import { useEffect, useRef, forwardRef, useImperativeHandle, useMemo, useCallback } from 'react'
import { HotTable } from '@handsontable/react'
import { HyperFormula } from 'hyperformula'
import { registerAllModules } from 'handsontable/registry'
import 'handsontable/dist/handsontable.full.min.css'
import { indexToColumnLetter, parseCellReference, parseRangeReference } from '../utils/excelParser'

// Register all Handsontable modules
registerAllModules()

const Spreadsheet = forwardRef(function Spreadsheet({ sheet }, ref) {
  const hotRef = useRef(null)
  const formulasRef = useRef({})
  const isApplyingChanges = useRef(false)

  // Initialize formulas from sheet data on mount
  useEffect(() => {
    if (sheet?.formulas) {
      formulasRef.current = { ...sheet.formulas }
    }
  }, [sheet?.name])

  // Create HyperFormula instance - only once
  const hyperformulaInstance = useMemo(() => {
    return HyperFormula.buildEmpty({
      licenseKey: 'gpl-v3'
    })
  }, [])

  // Adjust formula row references (e.g., =A2-B2 becomes =A3-B3)
  const adjustFormulaForRow = useCallback((formula, offset) => {
    return formula.replace(/([A-Z]+)(\d+)/g, (match, col, row) => {
      const newRow = parseInt(row, 10) + offset - 1
      return col + newRow
    })
  }, [])

  // Expose methods to parent
  useImperativeHandle(ref, () => ({
    getData: () => {
      const hot = hotRef.current?.hotInstance
      if (!hot) return sheet?.data || []
      return hot.getData()
    },
    
    getFormulas: () => {
      return { ...formulasRef.current }
    },
    
    applyChanges: (changes) => {
      const hot = hotRef.current?.hotInstance
      if (!hot) return
      
      // Set flag to prevent infinite loop
      isApplyingChanges.current = true
      
      try {
        changes.forEach(change => {
          const { type, cell, value, formula, range, afterColumn, header } = change
          
          switch (type) {
            case 'setCellValue': {
              const cellRef = parseCellReference(cell)
              if (cellRef) {
                hot.setDataAtCell(cellRef.row, cellRef.col, value, 'programmatic')
              }
              break
            }
            
            case 'setFormula': {
              const cellRef = parseCellReference(cell)
              if (cellRef) {
                formulasRef.current[cell] = formula
                hot.setDataAtCell(cellRef.row, cellRef.col, formula, 'programmatic')
              }
              break
            }
            
            case 'insertColumn': {
              const colIndex = afterColumn ? 
                afterColumn.charCodeAt(0) - 64 :
                hot.countCols()
              hot.alter('insert_col_start', colIndex)
              if (header) {
                hot.setDataAtCell(0, colIndex, header, 'programmatic')
              }
              break
            }
            
            case 'insertRow': {
              const rowIndex = change.afterRow || hot.countRows()
              hot.alter('insert_row_below', rowIndex)
              break
            }
            
            case 'applyFormulaToRange': {
              const rangeRef = parseRangeReference(range)
              if (rangeRef) {
                const { start, end } = rangeRef
                const cellChanges = []
                for (let r = start.row; r <= end.row; r++) {
                  const adjustedFormula = adjustFormulaForRow(formula, r - start.row + 1)
                  const cellRefStr = indexToColumnLetter(start.col) + (r + 1)
                  formulasRef.current[cellRefStr] = adjustedFormula
                  cellChanges.push([r, start.col, adjustedFormula])
                }
                hot.setDataAtCell(cellChanges, 'programmatic')
              }
              break
            }
            
            case 'deleteColumn': {
              const cellRef = parseCellReference(cell || 'A1')
              if (cellRef) {
                hot.alter('remove_col', cellRef.col)
              }
              break
            }
            
            case 'deleteRow': {
              const rowNum = change.row || 0
              hot.alter('remove_row', rowNum)
              break
            }
          }
        })
        
        hot.render()
      } finally {
        // Reset flag after a short delay to allow React to settle
        setTimeout(() => {
          isApplyingChanges.current = false
        }, 100)
      }
    },
    
    getCellValue: (cell) => {
      const hot = hotRef.current?.hotInstance
      if (!hot) return null
      const cellRef = parseCellReference(cell)
      if (!cellRef) return null
      return hot.getDataAtCell(cellRef.row, cellRef.col)
    },
    
    getCellRange: (range) => {
      const hot = hotRef.current?.hotInstance
      if (!hot) return []
      const rangeRef = parseRangeReference(range)
      if (!rangeRef) return []
      
      const { start, end } = rangeRef
      const data = []
      for (let r = start.row; r <= end.row; r++) {
        const row = []
        for (let c = start.col; c <= end.col; c++) {
          row.push(hot.getDataAtCell(r, c))
        }
        data.push(row)
      }
      return data
    }
  }), [sheet, adjustFormulaForRow])

  // Handle cell changes - only for user edits
  const handleAfterChange = useCallback((changes, source) => {
    // Skip if loading data or applying programmatic changes
    if (source === 'loadData' || source === 'programmatic' || isApplyingChanges.current) {
      return
    }
    
    const hot = hotRef.current?.hotInstance
    if (!hot || !changes) return
    
    // Update formulas ref for any formula changes
    changes.forEach(([row, col, oldValue, newValue]) => {
      const cellRef = indexToColumnLetter(col) + (row + 1)
      if (typeof newValue === 'string' && newValue.startsWith('=')) {
        formulasRef.current[cellRef] = newValue
      } else if (formulasRef.current[cellRef]) {
        delete formulasRef.current[cellRef]
      }
    })
  }, [])

  // Generate column headers (A, B, C, ... AA, AB, etc.)
  const colHeaders = useMemo(() => {
    const numCols = sheet?.data?.[0]?.length || 26
    return Array.from({ length: numCols }, (_, i) => indexToColumnLetter(i))
  }, [sheet?.data?.[0]?.length])

  if (!sheet) return null

  return (
    <div className="h-full w-full overflow-hidden bg-surface">
      <HotTable
        ref={hotRef}
        data={sheet.data}
        formulas={{
          engine: hyperformulaInstance
        }}
        rowHeaders={true}
        colHeaders={colHeaders}
        width="100%"
        height="100%"
        licenseKey="non-commercial-and-evaluation"
        stretchH="all"
        autoWrapRow={true}
        autoWrapCol={true}
        manualColumnResize={true}
        manualRowResize={true}
        contextMenu={true}
        dropdownMenu={true}
        filters={true}
        multiColumnSorting={true}
        undo={true}
        afterChange={handleAfterChange}
        className="htDark"
      />
    </div>
  )
})

export default Spreadsheet
