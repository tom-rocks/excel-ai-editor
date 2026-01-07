// Tool definitions for Claude AI
// These match the server-side tool definitions

export const CLAUDE_TOOLS = [
  {
    name: "get_sheet_info",
    description: "Get information about the spreadsheet including sheet names, dimensions, and column headers"
  },
  {
    name: "get_cell_range", 
    description: "Read values from a range of cells",
    parameters: {
      range: "Cell range in Excel notation (e.g., 'A1:D10')"
    }
  },
  {
    name: "set_cell_value",
    description: "Set a single cell's value",
    parameters: {
      cell: "Cell reference (e.g., 'A1')",
      value: "The value to set"
    }
  },
  {
    name: "set_formula",
    description: "Set a formula in a cell",
    parameters: {
      cell: "Cell reference (e.g., 'E2')",
      formula: "Excel formula starting with '='"
    }
  },
  {
    name: "apply_formula_to_range",
    description: "Apply a formula to multiple cells",
    parameters: {
      range: "Cell range (e.g., 'E2:E100')",
      formula: "Formula for first cell - row refs auto-increment"
    }
  },
  {
    name: "insert_column",
    description: "Insert a new column",
    parameters: {
      afterColumn: "Insert after this column letter",
      header: "Header text for the column"
    }
  },
  {
    name: "insert_row",
    description: "Insert a new row",
    parameters: {
      afterRow: "Insert after this row number"
    }
  }
]

// Convert tool name and input to spreadsheet change object
export function toolToChange(toolName, input) {
  switch (toolName) {
    case 'set_cell_value':
      return { type: 'setCellValue', ...input }
    case 'set_formula':
      return { type: 'setFormula', ...input }
    case 'apply_formula_to_range':
      return { type: 'applyFormulaToRange', ...input }
    case 'insert_column':
      return { type: 'insertColumn', ...input }
    case 'insert_row':
      return { type: 'insertRow', ...input }
    default:
      return null
  }
}
