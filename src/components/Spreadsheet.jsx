import { useEffect, useRef, forwardRef, useImperativeHandle, useMemo } from 'react'
import { HotTable } from '@handsontable/react'
import { HyperFormula } from 'hyperformula'
import { registerAllModules } from 'handsontable/registry'
import 'handsontable/dist/handsontable.full.min.css'
import { indexToColumnLetter, parseCellReference, parseRangeReference } from '../utils/excelParser'

// Register all Handsontable modules
registerAllModules()

const Spreadsheet = forwardRef(function Spreadsheet({ sheet, onDataChange }, ref) {
  const hotRef = useRef(null)
  const formulasRef = useRef({})

  // Initialize formulas from sheet data
  useEffect(() => {
    if (sheet?.formulas) {
      formulasRef.current = { ...sheet.formulas }
    }
  }, [sheet?.name])

  // Create HyperFormula instance
  const hyperformulaInstance = useMemo(() => {
    return HyperFormula.buildEmpty({
      licenseKey: 'gpl-v3'
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
      
      changes.forEach(change => {
        const { type, sheet: sheetName, cell, value, formula, range, afterColumn, header } = change
        
        switch (type) {
          case 'setCellValue': {
            const ref = parseCellReference(cell)
            if (ref) {
              hot.setDataAtCell(ref.row, ref.col, value)
            }
            break
          }
          
          case 'setFormula': {
            const ref = parseCellReference(cell)
            if (ref) {
              // Store formula
              formulasRef.current[cell] = formula
              // Set the formula in the cell - Handsontable with formulas plugin will calculate it
              hot.setDataAtCell(ref.row, ref.col, formula)
            }
            break
          }
          
          case 'insertColumn': {
            const colIndex = afterColumn ? 
              afterColumn.charCodeAt(0) - 64 : // A=1, B=2, etc
              hot.countCols()
            hot.alter('insert_col_start', colIndex)
            if (header) {
              hot.setDataAtCell(0, colIndex, header)
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
              for (let r = start.row; r <= end.row; r++) {
                // Adjust formula row references
                const adjustedFormula = adjustFormulaForRow(formula, r - start.row + 1)
                const cellRef = indexToColumnLetter(start.col) + (r + 1)
                formulasRef.current[cellRef] = adjustedFormula
                hot.setDataAtCell(r, start.col, adjustedFormula)
              }
            }
            break
          }
          
          case 'deleteColumn': {
            const ref = parseCellReference(cell || 'A1')
            if (ref) {
              hot.alter('remove_col', ref.col)
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
    },
    
    getCellValue: (cell) => {
      const hot = hotRef.current?.hotInstance
      if (!hot) return null
      const ref = parseCellReference(cell)
      if (!ref) return null
      return hot.getDataAtCell(ref.row, ref.col)
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
  }), [sheet])

  // Adjust formula row references (e.g., =A2-B2 becomes =A3-B3)
  const adjustFormulaForRow = (formula, offset) => {
    return formula.replace(/([A-Z]+)(\d+)/g, (match, col, row) => {
      const newRow = parseInt(row, 10) + offset - 1
      return col + newRow
    })
  }

  // Handle cell changes
  const handleAfterChange = (changes, source) => {
    if (source === 'loadData') return
    
    const hot = hotRef.current?.hotInstance
    if (!hot) return
    
    // Update formulas ref for any formula changes
    if (changes) {
      changes.forEach(([row, col, oldValue, newValue]) => {
        const cellRef = indexToColumnLetter(col) + (row + 1)
        if (typeof newValue === 'string' && newValue.startsWith('=')) {
          formulasRef.current[cellRef] = newValue
        } else if (formulasRef.current[cellRef]) {
          delete formulasRef.current[cellRef]
        }
      })
    }
    
    if (onDataChange) {
      onDataChange(hot.getData(), formulasRef.current)
    }
  }

  // Generate column headers (A, B, C, ... AA, AB, etc.)
  const colHeaders = useMemo(() => {
    const numCols = sheet?.data?.[0]?.length || 26
    return Array.from({ length: numCols }, (_, i) => indexToColumnLetter(i))
  }, [sheet?.data])

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
        comments={true}
        customBorders={true}
        dropdownMenu={true}
        filters={true}
        multiColumnSorting={true}
        undo={true}
        afterChange={handleAfterChange}
        className="htDark"
        cell={Object.entries(formulasRef.current).map(([cellRef, formula]) => {
          const ref = parseCellReference(cellRef)
          if (!ref) return null
          return {
            row: ref.row,
            col: ref.col,
            className: 'formula-cell'
          }
        }).filter(Boolean)}
      />
    </div>
  )
})

export default Spreadsheet
