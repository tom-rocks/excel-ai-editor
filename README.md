# Excel AI Editor

Upload Excel files, edit them with Claude AI, and download working .xlsx files.

## Features

- **Full Excel Support**: Import/export .xlsx files with formulas intact
- **Live Spreadsheet**: Real-time formula calculation with 400+ Excel functions
- **AI-Powered Editing**: Ask Claude to create complex formulas, add columns, manipulate data
- **Multiple Sheets**: Navigate between sheets in your workbook
- **Download**: Export your changes as a working Excel file

## Tech Stack

- **Frontend**: React + Vite + Tailwind CSS
- **Spreadsheet**: Handsontable + HyperFormula
- **Excel I/O**: SheetJS (xlsx)
- **AI**: Claude API with tool use
- **Backend**: Express.js

## Setup

1. Clone the repository
2. Install dependencies:
   ```bash
   npm install
   ```
3. Set your Anthropic API key as an environment variable:
   ```bash
   export ANTHROPIC_API_KEY=your_key_here
   ```
4. Run development server:
   ```bash
   npm run dev
   ```

## Deployment

This app is configured for Railway deployment. Set the `ANTHROPIC_API_KEY` environment variable in your Railway project settings.

## Usage

1. Drag and drop an Excel file onto the upload area
2. View and edit your spreadsheet directly
3. Use the chat panel to ask Claude for help:
   - "Add a column that calculates profit margin"
   - "Create a SUM formula for column D"
   - "Add a row with averages at the bottom"
4. Click "Download .xlsx" to export your changes

## License

MIT
