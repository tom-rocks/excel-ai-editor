import { Download, Loader2 } from 'lucide-react'
import { useState } from 'react'
import { downloadExcel, getExportData } from '../utils/excelExporter'

export default function DownloadButton({ workbook, spreadsheetRef, fileName }) {
  const [isExporting, setIsExporting] = useState(false)

  const handleDownload = async () => {
    setIsExporting(true)
    try {
      // Get latest data from spreadsheet
      const exportWorkbook = getExportData(spreadsheetRef, workbook, 0)
      downloadExcel(exportWorkbook, fileName)
    } catch (error) {
      console.error('Export error:', error)
      alert('Error exporting file. Please try again.')
    } finally {
      setIsExporting(false)
    }
  }

  return (
    <button
      onClick={handleDownload}
      disabled={isExporting}
      className="flex items-center gap-2 px-5 py-2.5 bg-gradient-to-r from-success/20 to-success/10 
        text-success border border-success/30 rounded-xl font-medium text-sm
        hover:from-success/30 hover:to-success/20 transition-all disabled:opacity-50"
    >
      {isExporting ? (
        <Loader2 className="w-4 h-4 animate-spin" />
      ) : (
        <Download className="w-4 h-4" />
      )}
      <span>{isExporting ? 'Exportando...' : 'Descargar .xlsx'}</span>
    </button>
  )
}
