import * as XLSX from 'xlsx'

/**
 * Parse an Excel file and convert it to our internal format
 * @param {File} file - The Excel file to parse
 * @returns {Promise<Object>} Parsed workbook data
 */
export async function parseExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result)
        const workbook = XLSX.read(data, { 
          type: 'array',
          cellFormula: true,
          cellStyles: true,
          cellNF: true,
          cellDates: true
        })
        
        const sheets = workbook.SheetNames.map(sheetName => {
          const worksheet = workbook.Sheets[sheetName]
          const { data: sheetData, formulas } = convertSheetToData(worksheet)
          
          return {
            name: sheetName,
            data: sheetData,
            formulas: formulas,
            merges: worksheet['!merges'] || [],
            colWidths: getColumnWidths(worksheet),
            rowHeights: getRowHeights(worksheet)
          }
        })
        
        resolve({
          sheets,
          originalWorkbook: workbook
        })
      } catch (error) {
        reject(error)
      }
    }
    
    reader.onerror = reject
    reader.readAsArrayBuffer(file)
  })
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
          // For formula cells, show the calculated value
          row.push(cell.v !== undefined ? cell.v : '')
        } else if (cell.t === 'd') {
          // Date
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
