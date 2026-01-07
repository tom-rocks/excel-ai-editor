import { useState, useCallback } from 'react'
import { Upload, FileSpreadsheet, Loader2 } from 'lucide-react'

export default function FileUpload({ onFileUpload, isLoading }) {
  const [isDragging, setIsDragging] = useState(false)

  const handleDrag = useCallback((e) => {
    e.preventDefault()
    e.stopPropagation()
  }, [])

  const handleDragIn = useCallback((e) => {
    e.preventDefault()
    e.stopPropagation()
    if (e.dataTransfer.items && e.dataTransfer.items.length > 0) {
      setIsDragging(true)
    }
  }, [])

  const handleDragOut = useCallback((e) => {
    e.preventDefault()
    e.stopPropagation()
    setIsDragging(false)
  }, [])

  const handleDrop = useCallback((e) => {
    e.preventDefault()
    e.stopPropagation()
    setIsDragging(false)

    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const file = e.dataTransfer.files[0]
      if (isValidExcelFile(file)) {
        onFileUpload(file)
      } else {
        alert('Please upload a valid Excel file (.xlsx, .xls)')
      }
    }
  }, [onFileUpload])

  const handleFileSelect = useCallback((e) => {
    if (e.target.files && e.target.files.length > 0) {
      const file = e.target.files[0]
      if (isValidExcelFile(file)) {
        onFileUpload(file)
      } else {
        alert('Please upload a valid Excel file (.xlsx, .xls)')
      }
    }
  }, [onFileUpload])

  const isValidExcelFile = (file) => {
    const validTypes = [
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/vnd.ms-excel',
      'application/excel',
      'application/x-excel',
      'application/x-msexcel'
    ]
    const validExtensions = ['.xlsx', '.xls']
    
    return validTypes.includes(file.type) || 
           validExtensions.some(ext => file.name.toLowerCase().endsWith(ext))
  }

  return (
    <div className="flex-1 flex items-center justify-center p-8">
      <div
        className={`drop-zone w-full max-w-2xl aspect-[4/3] rounded-2xl border-2 border-dashed 
          ${isDragging ? 'active border-accent' : 'border-surface-light'} 
          flex flex-col items-center justify-center gap-6 transition-all duration-300
          ${isLoading ? 'pointer-events-none opacity-50' : 'cursor-pointer hover:border-accent/50'}`}
        onDragEnter={handleDragIn}
        onDragLeave={handleDragOut}
        onDragOver={handleDrag}
        onDrop={handleDrop}
        onClick={() => !isLoading && document.getElementById('file-input').click()}
      >
        <input
          id="file-input"
          type="file"
          accept=".xlsx,.xls"
          onChange={handleFileSelect}
          className="hidden"
        />
        
        <div className={`w-24 h-24 rounded-2xl bg-surface flex items-center justify-center
          ${isDragging ? 'pulse-glow' : ''}`}>
          {isLoading ? (
            <Loader2 className="w-12 h-12 text-accent animate-spin" />
          ) : isDragging ? (
            <FileSpreadsheet className="w-12 h-12 text-accent" />
          ) : (
            <Upload className="w-12 h-12 text-gray-500" />
          )}
        </div>
        
        <div className="text-center">
          <h2 className="text-2xl font-display font-semibold text-white mb-2">
            {isLoading ? 'Processing...' : isDragging ? 'Drop it!' : 'Upload Excel File'}
          </h2>
          <p className="text-gray-500 max-w-md">
            {isLoading 
              ? 'Parsing your spreadsheet and preparing the editor...'
              : 'Drag and drop your .xlsx or .xls file here, or click to browse'
            }
          </p>
        </div>

        {!isLoading && (
          <div className="flex items-center gap-2 text-sm text-gray-600">
            <FileSpreadsheet className="w-4 h-4" />
            <span>Supports Excel 2007+ formats</span>
          </div>
        )}
      </div>
    </div>
  )
}
