import * as XLSX from 'xlsx'
import JSZip from 'jszip'

/**
 * Parse an Excel file and convert it to our internal format
 * @param {File} file - The Excel file to parse
 * @returns {Promise<Object>} Parsed workbook data
 */
export async function parseExcelFile(file) {
  const arrayBuffer = await file.arrayBuffer()
  const data = new Uint8Array(arrayBuffer)
  
  // Store the original array buffer for later export
  const originalArrayBuffer = data.slice()
  
  // Parse with SheetJS
  const workbook = XLSX.read(data, { 
    type: 'array',
    cellFormula: true,
    cellStyles: true,
    cellNF: true,
    cellDates: true,
    bookVBA: true,
    bookDeps: true
  })
  
  // Extract data validations from raw XML
  const validations = await extractDataValidations(arrayBuffer, workbook.SheetNames)
  
  // Get named ranges
  const namedRanges = {}
  if (workbook.Workbook?.Names) {
    workbook.Workbook.Names.forEach(n => {
      namedRanges[n.Name] = n.Ref
    })
  }
  
  // Resolve dropdown lists from named ranges
  const dropdownLists = resolveDropdownLists(workbook, namedRanges)
  
  const sheets = workbook.SheetNames.map((sheetName, idx) => {
    const worksheet = workbook.Sheets[sheetName]
    const { data: sheetData, formulas } = convertSheetToData(worksheet)
    
    // Get dropdowns for this sheet
    const sheetValidations = validations[idx] || []
    const dropdowns = sheetValidations.map(v => ({
      range: v.sqref,
      listName: v.formula,
      options: dropdownLists[v.formula] || []
    })).filter(d => d.options.length > 0)
    
    return {
      name: sheetName,
      data: sheetData,
      formulas: formulas,
      dropdowns: dropdowns,
      merges: worksheet['!merges'] || [],
      colWidths: getColumnWidths(worksheet),
      rowHeights: getRowHeights(worksheet)
    }
  })
  
  return {
    sheets,
    namedRanges,
    dropdownLists,
    originalWorkbook: workbook,
    originalArrayBuffer: originalArrayBuffer,
    fileName: file.name
  }
}

/**
 * Extract data validations from the raw xlsx file
 */
async function extractDataValidations(arrayBuffer, sheetNames) {
  try {
    const zip = await JSZip.loadAsync(arrayBuffer)
    const validations = {}
    
    for (let i = 0; i < sheetNames.length; i++) {
      const sheetFile = `xl/worksheets/sheet${i + 1}.xml`
      const file = zip.file(sheetFile)
      if (!file) continue
      
      const xml = await file.async('string')
      const matches = xml.match(/<dataValidation[^]*?<\/dataValidation>/g)
      
      if (matches) {
        validations[i] = matches.map(v => {
          const sqref = v.match(/sqref="([^"]+)"/)
          const formula = v.match(/<formula1>([^<]+)<\/formula1>/)
          return {
            sqref: sqref ? sqref[1] : null,
            formula: formula ? formula[1] : null
          }
        }).filter(v => v.sqref && v.formula)
      }
    }
    
    return validations
  } catch (e) {
    console.error('Error extracting data validations:', e)
    return {}
  }
}

/**
 * Resolve dropdown lists from named ranges
 */
function resolveDropdownLists(workbook, namedRanges) {
  const lists = {}
  
  Object.entries(namedRanges).forEach(([name, ref]) => {
    // Parse OFFSET formulas like: OFFSET(Listas!$A$2,0,0,COUNTA(Listas!$A:$A)-1,1)
    const offsetMatch = ref.match(/OFFSET\(([^!]+)!\$([A-Z]+)\$(\d+)/)
    if (offsetMatch) {
      const sheetName = offsetMatch[1]
      const col = offsetMatch[2]
      const startRow = parseInt(offsetMatch[3])
      
      const sheet = workbook.Sheets[sheetName]
      if (!sheet) return
      
      // Get values from that column
      const values = []
      const colIndex = columnLetterToIndex(col)
      
      // Read up to 1000 rows
      for (let row = startRow; row < startRow + 1000; row++) {
        const cellRef = col + row
        const cell = sheet[cellRef]
        if (cell && cell.v !== undefined && cell.v !== '') {
          values.push(String(cell.v))
        } else if (values.length > 0) {
          // Stop at first empty cell after data
          break
        }
      }
      
      lists[name] = values
    }
  })
  
  return lists
}

/**
 * Convert a worksheet to a 2D array with formulas tracked separately
 */
function convertSheetToData(worksheet) {
  const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1')
  const data = []
  const formulas = {}
  
  // Ensure we have at least 100 rows and 26 columns for editing room
  const maxRow = Math.max(range.e.r + 1, 100)
  const maxCol = Math.max(range.e.c + 1, 26)
  
  for (let r = 0; r <= maxRow; r++) {
    const row = []
    for (let c = 0; c <= maxCol; c++) {
      const cellAddress = XLSX.utils.encode_cell({ r, c })
      const cell = worksheet[cellAddress]
      
      if (cell) {
        // Store formula if exists
        if (cell.f) {
          formulas[cellAddress] = '=' + cell.f
        }
        
        // Get display value
        if (cell.f) {
          row.push(cell.v !== undefined ? cell.v : '')
        } else if (cell.t === 'd') {
          row.push(cell.v)
        } else if (cell.v !== undefined) {
          row.push(cell.v)
        } else {
          row.push('')
        }
      } else {
        row.push('')
      }
    }
    data.push(row)
  }
  
  return { data, formulas }
}

/**
 * Get column widths from worksheet
 */
function getColumnWidths(worksheet) {
  const cols = worksheet['!cols'] || []
  return cols.map(col => col?.wpx || col?.width || 100)
}

/**
 * Get row heights from worksheet
 */
function getRowHeights(worksheet) {
  const rows = worksheet['!rows'] || []
  return rows.map(row => row?.hpx || row?.hpt || 23)
}

/**
 * Convert column letter to index (A=0, B=1, etc.)
 */
export function columnLetterToIndex(letter) {
  let index = 0
  for (let i = 0; i < letter.length; i++) {
    index = index * 26 + (letter.charCodeAt(i) - 64)
  }
  return index - 1
}

/**
 * Convert column index to letter (0=A, 1=B, etc.)
 */
export function indexToColumnLetter(index) {
  let letter = ''
  index++
  while (index > 0) {
    const remainder = (index - 1) % 26
    letter = String.fromCharCode(65 + remainder) + letter
    index = Math.floor((index - 1) / 26)
  }
  return letter
}

/**
 * Parse a cell reference like "A1" into { row, col }
 */
export function parseCellReference(ref) {
  const match = ref.match(/^([A-Z]+)(\d+)$/)
  if (!match) return null
  
  const col = columnLetterToIndex(match[1])
  const row = parseInt(match[2], 10) - 1
  
  return { row, col }
}

/**
 * Parse a range reference like "A1:B10" into { start, end }
 */
export function parseRangeReference(range) {
  const parts = range.split(':')
  if (parts.length === 1) {
    const cell = parseCellReference(parts[0])
    return { start: cell, end: cell }
  }
  
  return {
    start: parseCellReference(parts[0]),
    end: parseCellReference(parts[1])
  }
}
