import express from 'express'
import cors from 'cors'
import Anthropic from '@anthropic-ai/sdk'
import dotenv from 'dotenv'
import { fileURLToPath } from 'url'
import { dirname, join } from 'path'

dotenv.config()

const __filename = fileURLToPath(import.meta.url)
const __dirname = dirname(__filename)

const app = express()
const PORT = process.env.PORT || 3001

// Middleware
app.use(cors())
app.use(express.json({ limit: '50mb' }))

// Serve static files
app.use(express.static(join(__dirname, '../dist')))

// Initialize Anthropic client
const anthropic = new Anthropic({
  apiKey: process.env.ANTHROPIC_API_KEY
})

// Tool definitions for Claude
const tools = [
  {
    name: "get_sheet_info",
    description: "Get information about all available sheets in the workbook, including their names, dimensions, and a preview of the header row. Call this first to understand the spreadsheet structure.",
    input_schema: {
      type: "object",
      properties: {},
      required: []
    }
  },
  {
    name: "get_cell_range",
    description: "Read values from a range of cells. Use Excel-style notation like 'A1:D10'. Returns the values as a 2D array.",
    input_schema: {
      type: "object",
      properties: {
        sheet: { type: "string", description: "The sheet name to read from" },
        range: { type: "string", description: "The cell range in Excel notation (e.g., 'A1:D10')" }
      },
      required: ["sheet", "range"]
    }
  },
  {
    name: "set_cell_value",
    description: "Set a single cell's value. For formulas, use set_formula instead.",
    input_schema: {
      type: "object",
      properties: {
        sheet: { type: "string", description: "The sheet name" },
        cell: { type: "string", description: "The cell reference (e.g., 'A1')" },
        value: { type: "string", description: "The value to set" }
      },
      required: ["sheet", "cell", "value"]
    }
  },
  {
    name: "set_formula",
    description: "Set a formula in a cell. The formula should start with '=' and use Excel formula syntax.",
    input_schema: {
      type: "object",
      properties: {
        sheet: { type: "string", description: "The sheet name" },
        cell: { type: "string", description: "The cell reference (e.g., 'E2')" },
        formula: { type: "string", description: "The formula starting with '=' (e.g., '=A2*B2', '=SUM(A1:A10)')" }
      },
      required: ["sheet", "cell", "formula"]
    }
  },
  {
    name: "insert_column",
    description: "Insert a new column after a specified column. Optionally set a header value.",
    input_schema: {
      type: "object",
      properties: {
        sheet: { type: "string", description: "The sheet name" },
        afterColumn: { type: "string", description: "The column letter after which to insert" },
        header: { type: "string", description: "Optional header text for the new column" }
      },
      required: ["sheet", "afterColumn"]
    }
  },
  {
    name: "insert_row",
    description: "Insert a new row after a specified row number.",
    input_schema: {
      type: "object",
      properties: {
        sheet: { type: "string", description: "The sheet name" },
        afterRow: { type: "integer", description: "The row number after which to insert (1-indexed)" }
      },
      required: ["sheet", "afterRow"]
    }
  },
  {
    name: "apply_formula_to_range",
    description: "Apply a formula pattern to a range of cells. The formula will be automatically adjusted for each row.",
    input_schema: {
      type: "object",
      properties: {
        sheet: { type: "string", description: "The sheet name" },
        range: { type: "string", description: "The target range (e.g., 'E2:E100')" },
        formula: { type: "string", description: "The formula pattern for the first cell" }
      },
      required: ["sheet", "range", "formula"]
    }
  },
  {
    name: "delete_column",
    description: "Delete a column from the sheet.",
    input_schema: {
      type: "object",
      properties: {
        sheet: { type: "string", description: "The sheet name" },
        column: { type: "string", description: "The column letter to delete" }
      },
      required: ["sheet", "column"]
    }
  },
  {
    name: "delete_row",
    description: "Delete a row from the sheet.",
    input_schema: {
      type: "object",
      properties: {
        sheet: { type: "string", description: "The sheet name" },
        row: { type: "integer", description: "The row number to delete (1-indexed)" }
      },
      required: ["sheet", "row"]
    }
  }
]

// Chat endpoint
app.post('/api/chat', async (req, res) => {
  try {
    const { message, spreadsheetData, conversationHistory = [] } = req.body

    if (!process.env.ANTHROPIC_API_KEY) {
      return res.status(500).json({ 
        error: 'ANTHROPIC_API_KEY not configured',
        message: 'Please set the ANTHROPIC_API_KEY environment variable'
      })
    }

    // Build messages array
    const messages = [
      ...conversationHistory,
      { role: 'user', content: message }
    ]

    // Build comprehensive spreadsheet context
    const spreadsheetContext = buildSpreadsheetContext(spreadsheetData)

    // System prompt - Argentinian Spanish, friendly for Tomi
    const systemPrompt = `Sos un asistente experto en Excel que ayuda a Tomi a editar sus planillas. Hablás en español rioplatense (Argentina), de manera amigable, relajada y natural. Usás "vos" en vez de "tú", y expresiones como "dale", "buenísimo", "genial", "tranqui", etc.

PERSONALIDAD:
- Sos paciente y explicás las cosas de forma simple, porque Tomi no maneja bien las fórmulas
- Cuando creás una fórmula, explicá brevemente qué hace en palabras simples
- Celebrá los logros ("¡Listo! Quedó joya")
- Si algo puede ser confuso, aclaralo con ejemplos
- Sé proactivo: si ves algo que se podría mejorar, sugerilo

CAPACIDADES:
Tenés acceso completo al archivo Excel de Tomi. Podés:
- Ver TODOS los datos en tiempo real
- Crear y modificar fórmulas (SUM, AVERAGE, VLOOKUP, IF, etc.)
- Agregar/eliminar columnas y filas
- Aplicar fórmulas a rangos enteros
- Analizar datos y dar sugerencias

REGLAS IMPORTANTES:
1. SIEMPRE usá get_sheet_info primero para entender la estructura
2. Usá get_cell_range para leer datos antes de hacer cambios
3. Explicá en español simple qué va a hacer cada fórmula
4. Para fórmulas en múltiples filas, usá apply_formula_to_range
5. Sé preciso con las referencias de celdas

CONTEXTO ACTUAL DE LA PLANILLA:
${spreadsheetContext}`

    // Initial Claude call
    let response = await anthropic.messages.create({
      model: 'claude-sonnet-4-20250514',
      max_tokens: 4096,
      system: systemPrompt,
      tools,
      messages
    })

    // Collect tool calls and results
    const toolCalls = []
    const pendingChanges = []

    // Process tool use in a loop
    while (response.stop_reason === 'tool_use') {
      const toolUseBlocks = response.content.filter(block => block.type === 'tool_use')
      const toolResults = []

      for (const toolUse of toolUseBlocks) {
        const { id, name, input } = toolUse
        
        // Execute tool and get result
        const result = executeToolOnServer(name, input, spreadsheetData)
        
        toolCalls.push({
          tool: name,
          input,
          result
        })

        // Collect changes to send back to client
        if (result.change) {
          pendingChanges.push(result.change)
        }

        toolResults.push({
          type: 'tool_result',
          tool_use_id: id,
          content: JSON.stringify(result.output)
        })
      }

      // Continue conversation with tool results
      messages.push({ role: 'assistant', content: response.content })
      messages.push({ role: 'user', content: toolResults })

      response = await anthropic.messages.create({
        model: 'claude-sonnet-4-20250514',
        max_tokens: 4096,
        system: systemPrompt,
        tools,
        messages
      })
    }

    // Extract text response
    const textContent = response.content.find(block => block.type === 'text')
    const assistantMessage = textContent?.text || 'Done!'

    res.json({
      message: assistantMessage,
      toolCalls,
      changes: pendingChanges,
      conversationHistory: [
        ...conversationHistory,
        { role: 'user', content: message },
        { role: 'assistant', content: assistantMessage }
      ]
    })

  } catch (error) {
    console.error('Chat error:', error)
    res.status(500).json({ 
      error: error.message,
      details: error.response?.data || null
    })
  }
})

// Execute tool on server
function executeToolOnServer(toolName, input, spreadsheetData) {
  const { sheets, activeSheet } = spreadsheetData || { sheets: [], activeSheet: 0 }
  const currentSheet = sheets[activeSheet] || { data: [], formulas: {} }

  switch (toolName) {
    case 'get_sheet_info':
      return {
        output: {
          sheets: sheets.map((s, idx) => ({
            name: s.name,
            isActive: idx === activeSheet,
            rows: s.data?.length || 0,
            columns: s.data?.[0]?.length || 0,
            headers: s.data?.[0]?.slice(0, 20) || []
          }))
        }
      }

    case 'get_cell_range': {
      const sheet = sheets.find(s => s.name === input.sheet) || currentSheet
      const range = parseRange(input.range)
      if (!range) return { output: { error: 'Invalid range format' } }

      const data = []
      for (let r = range.startRow; r <= Math.min(range.endRow, (sheet.data?.length || 0) - 1); r++) {
        const row = []
        for (let c = range.startCol; c <= range.endCol; c++) {
          row.push(sheet.data?.[r]?.[c] ?? '')
        }
        data.push(row)
      }

      return { output: { range: input.range, data } }
    }

    case 'set_cell_value':
      return {
        output: { success: true, message: `Set ${input.cell} to "${input.value}"` },
        change: { type: 'setCellValue', sheet: input.sheet, cell: input.cell, value: input.value }
      }

    case 'set_formula':
      return {
        output: { success: true, message: `Set formula in ${input.cell}: ${input.formula}` },
        change: { type: 'setFormula', sheet: input.sheet, cell: input.cell, formula: input.formula }
      }

    case 'insert_column':
      return {
        output: { success: true, message: `Inserted column after ${input.afterColumn}` },
        change: { type: 'insertColumn', sheet: input.sheet, afterColumn: input.afterColumn, header: input.header }
      }

    case 'insert_row':
      return {
        output: { success: true, message: `Inserted row after ${input.afterRow}` },
        change: { type: 'insertRow', sheet: input.sheet, afterRow: input.afterRow }
      }

    case 'apply_formula_to_range':
      return {
        output: { success: true, message: `Applied formula to ${input.range}` },
        change: { type: 'applyFormulaToRange', sheet: input.sheet, range: input.range, formula: input.formula }
      }

    case 'delete_column':
      return {
        output: { success: true, message: `Deleted column ${input.column}` },
        change: { type: 'deleteColumn', sheet: input.sheet, cell: input.column + '1' }
      }

    case 'delete_row':
      return {
        output: { success: true, message: `Deleted row ${input.row}` },
        change: { type: 'deleteRow', sheet: input.sheet, row: input.row - 1 }
      }

    default:
      return { output: { error: `Unknown tool: ${toolName}` } }
  }
}

// Helper: Parse range like "A1:D10"
function parseRange(range) {
  const match = range.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/)
  if (!match) return null

  return {
    startCol: columnToIndex(match[1]),
    startRow: parseInt(match[2], 10) - 1,
    endCol: columnToIndex(match[3]),
    endRow: parseInt(match[4], 10) - 1
  }
}

function columnToIndex(col) {
  let index = 0
  for (let i = 0; i < col.length; i++) {
    index = index * 26 + (col.charCodeAt(i) - 64)
  }
  return index - 1
}

// Build comprehensive spreadsheet context for Claude
function buildSpreadsheetContext(spreadsheetData) {
  if (!spreadsheetData || !spreadsheetData.sheets) {
    return 'No hay planilla cargada todavía.'
  }

  const { sheets, activeSheet } = spreadsheetData
  const currentSheet = sheets[activeSheet]
  
  let context = ''
  
  // Sheet overview
  context += `📊 HOJAS DISPONIBLES: ${sheets.map((s, i) => i === activeSheet ? `[${s.name}] (activa)` : s.name).join(', ')}\n\n`
  
  // Current sheet details
  context += `📋 HOJA ACTIVA: "${currentSheet.name}"\n`
  context += `   Filas con datos: ${countDataRows(currentSheet.data)}\n`
  context += `   Columnas: ${currentSheet.data?.[0]?.length || 0}\n\n`
  
  // Headers with column letters
  if (currentSheet.data && currentSheet.data[0]) {
    context += `📝 ENCABEZADOS (Fila 1):\n`
    currentSheet.data[0].forEach((header, idx) => {
      if (header !== '' && header !== null && header !== undefined) {
        const colLetter = indexToColumnLetter(idx)
        context += `   ${colLetter}: "${header}"\n`
      }
    })
    context += '\n'
  }
  
  // Full data preview (up to 50 rows for context)
  const maxRows = Math.min(50, currentSheet.data?.length || 0)
  if (maxRows > 1) {
    context += `📊 DATOS (primeras ${maxRows} filas):\n`
    for (let r = 0; r < maxRows; r++) {
      const row = currentSheet.data[r]
      if (!row) continue
      
      const nonEmptyCells = []
      row.forEach((cell, idx) => {
        if (cell !== '' && cell !== null && cell !== undefined) {
          const colLetter = indexToColumnLetter(idx)
          nonEmptyCells.push(`${colLetter}${r+1}=${cell}`)
        }
      })
      
      if (nonEmptyCells.length > 0) {
        context += `   Fila ${r + 1}: ${nonEmptyCells.join(', ')}\n`
      }
    }
    
    const totalRows = countDataRows(currentSheet.data)
    if (totalRows > maxRows) {
      context += `   ... y ${totalRows - maxRows} filas más\n`
    }
  }
  
  // Existing formulas
  if (currentSheet.formulas && Object.keys(currentSheet.formulas).length > 0) {
    context += `\n🔢 FÓRMULAS EXISTENTES:\n`
    Object.entries(currentSheet.formulas).forEach(([cell, formula]) => {
      context += `   ${cell}: ${formula}\n`
    })
  }
  
  return context
}

// Count rows that have at least one non-empty cell
function countDataRows(data) {
  if (!data) return 0
  return data.filter(row => 
    row && row.some(cell => cell !== '' && cell !== null && cell !== undefined)
  ).length
}

// Convert column index to letter (0=A, 1=B, etc.)
function indexToColumnLetter(index) {
  let letter = ''
  index++
  while (index > 0) {
    const remainder = (index - 1) % 26
    letter = String.fromCharCode(65 + remainder) + letter
    index = Math.floor((index - 1) / 26)
  }
  return letter
}

// Health check
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() })
})

// Catch-all for SPA routing
app.get('*', (req, res) => {
  res.sendFile(join(__dirname, '../dist/index.html'))
})

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`)
  console.log(`Environment: ${process.env.NODE_ENV || 'development'}`)
})
