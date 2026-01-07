import { Table2 } from 'lucide-react'

export default function SheetTabs({ sheets, activeSheet, onSheetChange }) {
  return (
    <div className="flex items-center gap-1 px-4 py-2 bg-surface border-b border-surface-light overflow-x-auto">
      {sheets.map((sheet, index) => (
        <button
          key={index}
          onClick={() => onSheetChange(index)}
          className={`flex items-center gap-2 px-4 py-2 rounded-lg text-sm font-medium transition-all
            ${index === activeSheet 
              ? 'bg-accent/10 text-accent border border-accent/30' 
              : 'text-gray-400 hover:text-white hover:bg-surface-light'
            }`}
        >
          <Table2 className="w-4 h-4" />
          <span className="whitespace-nowrap">{sheet.name}</span>
        </button>
      ))}
    </div>
  )
}
