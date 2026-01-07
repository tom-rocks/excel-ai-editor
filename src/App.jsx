import { useState, useRef, useCallback } from 'react'
import FileUpload from './components/FileUpload'
import Spreadsheet from './components/Spreadsheet'
import SheetTabs from './components/SheetTabs'
import ChatPanel from './components/ChatPanel'
import DownloadButton from './components/DownloadButton'
import { parseExcelFile } from './utils/excelParser'
import { FileSpreadsheet, Sparkles } from 'lucide-react'

function App() {
  const [workbook, setWorkbook] = useState(null)
  const [activeSheet, setActiveSheet] = useState(0)
  const [isLoading, setIsLoading] = useState(false)
  const [fileName, setFileName] = useState('')
  const spreadsheetRef = useRef(null)

  const handleFileUpload = useCallback(async (file) => {
    setIsLoading(true)
    try {
      const parsed = await parseExcelFile(file)
      setWorkbook(parsed)
      setActiveSheet(0)
      setFileName(file.name)
    } catch (error) {
      console.error('Error parsing file:', error)
      alert('Error parsing Excel file. Please try another file.')
    } finally {
      setIsLoading(false)
    }
  }, [])

  const handleSheetChange = useCallback((index) => {
    setActiveSheet(index)
  }, [])

  const getSpreadsheetData = useCallback(() => {
    if (!workbook) return null
    return {
      sheets: workbook.sheets.map((sheet, idx) => ({
        name: sheet.name,
        data: idx === activeSheet && spreadsheetRef.current 
          ? spreadsheetRef.current.getData() 
          : sheet.data,
        formulas: idx === activeSheet && spreadsheetRef.current
          ? spreadsheetRef.current.getFormulas()
          : sheet.formulas
      })),
      activeSheet
    }
  }, [workbook, activeSheet])

  const applyChanges = useCallback((changes) => {
    if (spreadsheetRef.current) {
      spreadsheetRef.current.applyChanges(changes)
    }
  }, [])


  return (
    <div className="h-screen flex flex-col bg-midnight">
      {/* Header */}
      <header className="flex items-center justify-between px-6 py-4 border-b border-surface-light">
        <div className="flex items-center gap-3">
          <div className="w-10 h-10 rounded-xl bg-gradient-to-br from-accent to-success flex items-center justify-center">
            <FileSpreadsheet className="w-5 h-5 text-midnight" />
          </div>
          <div>
            <h1 className="text-xl font-display font-semibold text-white">Excel para Tomi</h1>
            {fileName && (
              <p className="text-sm text-gray-500 font-mono">{fileName}</p>
            )}
          </div>
        </div>
        
        <div className="flex items-center gap-4">
          {workbook && <DownloadButton workbook={workbook} spreadsheetRef={spreadsheetRef} fileName={fileName} />}
        </div>
      </header>

      {/* Main content */}
      <div className="flex-1 flex overflow-hidden">
        {!workbook ? (
          <FileUpload onFileUpload={handleFileUpload} isLoading={isLoading} />
        ) : (
          <>
            {/* Spreadsheet area */}
            <div className="flex-1 flex flex-col min-w-0">
              <SheetTabs 
                sheets={workbook.sheets} 
                activeSheet={activeSheet} 
                onSheetChange={handleSheetChange} 
              />
              <div className="flex-1 overflow-hidden">
                <Spreadsheet 
                  ref={spreadsheetRef}
                  sheet={workbook.sheets[activeSheet]}
                />
              </div>
            </div>

            {/* Chat panel */}
            <ChatPanel 
              getSpreadsheetData={getSpreadsheetData}
              applyChanges={applyChanges}
              activeSheet={activeSheet}
              sheetName={workbook.sheets[activeSheet]?.name}
            />
          </>
        )}
      </div>
    </div>
  )
}

export default App
