import * as XLSX from 'xlsx'

/**
 * Export workbook by modifying the ORIGINAL file, preserving all features
 * @param {Object} workbook - Our workbook object (must include originalArrayBuffer)
 * @param {Object} currentSheetData - Current data from the active spreadsheet
 * @param {number} activeSheetIndex - Index of the currently active sheet
 * @param {string} fileName - Original filename
 */
export function downloadExcel(workbook, spreadsheetRef, fileName = 'export.xlsx') {
  if (!workbook.originalArrayBuffer) {
    console.error('No original file data available')
    alert('Error: No se puede exportar sin el archivo original')
    return
  }

  // Re-read the original workbook
  const originalWb = XLSX.read(workbook.originalArrayBuffer, {
    type: 'array',
    cellFormula: true,
    cellStyles: true,
    cellNF: true,
    cellDates: true,
    bookVBA: true,  // Preserve macros if any
    bookDeps: true, // Preserve dependencies
  })

  // Get current data from the spreadsheet ref for all sheets
  workbook.sheets.forEach((sheet, sheetIndex) => {
    const wsName = sheet.name
    const ws = originalWb.Sheets[wsName]
    if (!ws) return

    // Get the current data for this sheet
    const currentData = sheet.data
    const currentFormulas = sheet.formulas || {}

    // Update only the cells that have data
    if (currentData) {
      for (let r = 0; r < currentData.length; r++) {
        for (let c = 0; c < (currentData[r]?.length || 0); c++) {
          const cellRef = XLSX.utils.encode_cell({ r, c })
          const value = currentData[r][c]
          const formula = currentFormulas[cellRef]

          // Skip empty cells that were already empty
          if ((value === '' || value === null || value === undefined) && !ws[cellRef]) {
            continue
          }

          // Create or update the cell
          if (formula && formula.startsWith('=')) {
            // It's a formula - update the formula
            if (!ws[cellRef]) ws[cellRef] = {}
            ws[cellRef].f = formula.substring(1) // Remove leading =
            // Let Excel recalculate the value
            delete ws[cellRef].v
          } else if (value !== '' && value !== null && value !== undefined) {
            // It's a value
            if (!ws[cellRef]) ws[cellRef] = {}
            
            // Determine type
            if (typeof value === 'number') {
              ws[cellRef].t = 'n'
              ws[cellRef].v = value
            } else if (typeof value === 'boolean') {
              ws[cellRef].t = 'b'
              ws[cellRef].v = value
            } else if (value instanceof Date) {
              ws[cellRef].t = 'd'
              ws[cellRef].v = value
            } else {
              ws[cellRef].t = 's'
              ws[cellRef].v = String(value)
            }
            
            // If there was a formula before but now it's a value, remove formula
            if (ws[cellRef].f && !formula) {
              delete ws[cellRef].f
            }
          }
        }
      }
    }

    // Update the range if needed
    const range = XLSX.utils.decode_range(ws['!ref'] || 'A1')
    const newMaxRow = currentData?.length || 0
    const newMaxCol = currentData?.[0]?.length || 0
    
    if (newMaxRow - 1 > range.e.r) range.e.r = newMaxRow - 1
    if (newMaxCol - 1 > range.e.c) range.e.c = newMaxCol - 1
    
    ws['!ref'] = XLSX.utils.encode_range(range)
  })

  // Write the modified workbook
  const wbout = XLSX.write(originalWb, { 
    bookType: 'xlsx', 
    type: 'array',
    cellStyles: true,
    bookVBA: true
  })

  // Create blob and download
  const blob = new Blob([wbout], { 
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
  })
  
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  
  // Generate filename
  const baseName = fileName.replace(/\.[^/.]+$/, '')
  a.download = `${baseName}_edited.xlsx`
  
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  URL.revokeObjectURL(url)
}

/**
 * Get export data - sync current spreadsheet state to workbook
 */
export function getExportData(spreadsheetRef, workbook, activeSheet) {
  if (!spreadsheetRef?.current || !workbook) return workbook
  
  const currentData = spreadsheetRef.current.getData()
  const currentFormulas = spreadsheetRef.current.getFormulas()
  
  // Update the active sheet's data
  const updatedWorkbook = { ...workbook }
  updatedWorkbook.sheets = [...workbook.sheets]
  updatedWorkbook.sheets[activeSheet] = {
    ...workbook.sheets[activeSheet],
    data: currentData,
    formulas: currentFormulas
  }
  
  return updatedWorkbook
}
