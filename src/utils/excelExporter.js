import * as XLSX from 'xlsx'
import { indexToColumnLetter } from './excelParser'

/**
 * Export workbook data to an Excel file
 * @param {Object} workbook - Our internal workbook format
 * @param {string} fileName - Original filename for naming
 * @returns {Blob} Excel file blob
 */
export function exportToExcel(workbook, fileName = 'export.xlsx') {
  const wb = XLSX.utils.book_new()
  
  workbook.sheets.forEach(sheet => {
    const ws = createWorksheet(sheet)
    XLSX.utils.book_append_sheet(wb, ws, sheet.name)
  })
  
  // Generate blob
  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' })
  return new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
}

/**
 * Create a worksheet from our internal sheet format
 */
function createWorksheet(sheet) {
  const { data, formulas } = sheet
  
  // Create worksheet from data
  const ws = XLSX.utils.aoa_to_sheet(data)
  
  // Apply formulas
  if (formulas) {
    Object.entries(formulas).forEach(([cellRef, formula]) => {
      if (formula && formula.startsWith('=')) {
        const cell = ws[cellRef] || {}
        cell.f = formula.substring(1) // Remove leading =
        ws[cellRef] = cell
      }
    })
  }
  
  // Apply column widths if available
  if (sheet.colWidths && sheet.colWidths.length > 0) {
    ws['!cols'] = sheet.colWidths.map(w => ({ wpx: w || 100 }))
  }
  
  // Apply row heights if available
  if (sheet.rowHeights && sheet.rowHeights.length > 0) {
    ws['!rows'] = sheet.rowHeights.map(h => ({ hpx: h || 23 }))
  }
  
  // Apply merges if available
  if (sheet.merges && sheet.merges.length > 0) {
    ws['!merges'] = sheet.merges
  }
  
  return ws
}

/**
 * Download the workbook as an Excel file
 */
export function downloadExcel(workbook, fileName = 'export.xlsx') {
  const blob = exportToExcel(workbook, fileName)
  
  // Create download link
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = fileName.replace(/\.[^/.]+$/, '') + '_edited.xlsx'
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  URL.revokeObjectURL(url)
}

/**
 * Get all data from spreadsheet including formulas
 */
export function getExportData(spreadsheetRef, workbook, activeSheet) {
  if (!spreadsheetRef.current || !workbook) return workbook
  
  const currentData = spreadsheetRef.current.getData()
  const currentFormulas = spreadsheetRef.current.getFormulas()
  
  const updatedWorkbook = { ...workbook }
  updatedWorkbook.sheets = [...workbook.sheets]
  updatedWorkbook.sheets[activeSheet] = {
    ...workbook.sheets[activeSheet],
    data: currentData,
    formulas: currentFormulas
  }
  
  return updatedWorkbook
}
