import { useState, useRef, useEffect } from 'react'
import { Send, Loader2, Sparkles, User, Bot, Wrench, ChevronDown, ChevronRight } from 'lucide-react'

export default function ChatPanel({ getSpreadsheetData, applyChanges, activeSheet, sheetName }) {
  const [messages, setMessages] = useState([])
  const [input, setInput] = useState('')
  const [isLoading, setIsLoading] = useState(false)
  const [conversationHistory, setConversationHistory] = useState([])
  const messagesEndRef = useRef(null)
  const inputRef = useRef(null)

  const scrollToBottom = () => {
    messagesEndRef.current?.scrollIntoView({ behavior: 'smooth' })
  }

  useEffect(() => {
    scrollToBottom()
  }, [messages])

  const handleSubmit = async (e) => {
    e.preventDefault()
    if (!input.trim() || isLoading) return

    const userMessage = input.trim()
    setInput('')
    
    // Add user message to UI
    setMessages(prev => [...prev, { 
      role: 'user', 
      content: userMessage,
      timestamp: new Date()
    }])

    setIsLoading(true)

    try {
      // Get current spreadsheet data
      const spreadsheetData = getSpreadsheetData()

      const response = await fetch('/api/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          message: userMessage,
          spreadsheetData,
          conversationHistory
        })
      })

      if (!response.ok) {
        const errorData = await response.json()
        throw new Error(errorData.error || 'Failed to get response')
      }

      const data = await response.json()

      // Apply changes to spreadsheet
      if (data.changes && data.changes.length > 0) {
        applyChanges(data.changes)
      }

      // Update conversation history
      setConversationHistory(data.conversationHistory || [])

      // Add assistant message to UI
      setMessages(prev => [...prev, {
        role: 'assistant',
        content: data.message,
        toolCalls: data.toolCalls,
        timestamp: new Date()
      }])

    } catch (error) {
      console.error('Chat error:', error)
      setMessages(prev => [...prev, {
        role: 'error',
        content: error.message || 'An error occurred. Please try again.',
        timestamp: new Date()
      }])
    } finally {
      setIsLoading(false)
    }
  }

  const handleKeyDown = (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault()
      handleSubmit(e)
    }
  }

  return (
    <div className="w-96 border-l border-surface-light flex flex-col bg-surface">
      {/* Header */}
      <div className="px-4 py-3 border-b border-surface-light">
        <div className="flex items-center gap-2">
          <div className="w-8 h-8 rounded-lg bg-gradient-to-br from-accent to-success flex items-center justify-center">
            <Sparkles className="w-4 h-4 text-midnight" />
          </div>
          <div>
            <h2 className="font-semibold text-white text-sm">Asistente de Tomi</h2>
            <p className="text-xs text-gray-500">Editando: {sheetName}</p>
          </div>
        </div>
      </div>

      {/* Messages */}
      <div className="flex-1 overflow-y-auto p-4 space-y-4">
        {messages.length === 0 && (
          <div className="text-center py-8">
            <div className="w-16 h-16 mx-auto mb-4 rounded-2xl bg-gradient-to-br from-accent/20 to-success/20 flex items-center justify-center">
              <Sparkles className="w-8 h-8 text-accent" />
            </div>
            <h3 className="text-white font-medium mb-2">¡Hola Tomi! 👋</h3>
            <p className="text-gray-400 text-sm max-w-xs mx-auto">
              Soy tu asistente para Excel. Contame qué necesitás y te ayudo con las fórmulas.
            </p>
            <div className="mt-4 space-y-2">
              <SuggestionButton 
                onClick={() => setInput("Agregame una columna que calcule el total")} 
                text="📊 Agregar columna de totales"
              />
              <SuggestionButton 
                onClick={() => setInput("Sumame todos los valores de una columna")} 
                text="➕ Sumar una columna"
              />
              <SuggestionButton 
                onClick={() => setInput("Explicame qué datos tengo en la planilla")} 
                text="👀 Ver qué hay en mi archivo"
              />
              <SuggestionButton 
                onClick={() => setInput("Calculame el porcentaje de cada fila respecto al total")} 
                text="📈 Calcular porcentajes"
              />
            </div>
          </div>
        )}

        {messages.map((msg, idx) => (
          <Message key={idx} message={msg} />
        ))}

        {isLoading && (
          <div className="flex items-center gap-3 text-gray-400">
            <div className="w-8 h-8 rounded-lg bg-success/10 flex items-center justify-center">
              <Loader2 className="w-4 h-4 text-success animate-spin" />
            </div>
            <span className="text-sm">Pensando...</span>
          </div>
        )}

        <div ref={messagesEndRef} />
      </div>

      {/* Input */}
      <form onSubmit={handleSubmit} className="p-4 border-t border-surface-light">
        <div className="relative">
          <textarea
            ref={inputRef}
            value={input}
            onChange={(e) => setInput(e.target.value)}
            onKeyDown={handleKeyDown}
            placeholder="Contame qué necesitás..."
            rows={2}
            className="w-full px-4 py-3 pr-12 bg-surface-light rounded-xl text-white placeholder-gray-500 
              resize-none focus:outline-none focus:ring-2 focus:ring-accent/50 text-sm"
            disabled={isLoading}
          />
          <button
            type="submit"
            disabled={!input.trim() || isLoading}
            className="absolute right-2 bottom-2 p-2 rounded-lg bg-accent/10 text-accent 
              hover:bg-accent/20 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
          >
            {isLoading ? (
              <Loader2 className="w-5 h-5 animate-spin" />
            ) : (
              <Send className="w-5 h-5" />
            )}
          </button>
        </div>
        <p className="text-xs text-gray-600 mt-2 text-center">
          Enter para enviar · Shift+Enter para nueva línea
        </p>
      </form>
    </div>
  )
}

function Message({ message }) {
  const [showTools, setShowTools] = useState(false)
  
  if (message.role === 'user') {
    return (
      <div className="message-user rounded-xl p-4">
        <div className="flex items-start gap-3">
          <div className="w-8 h-8 rounded-lg bg-accent/10 flex items-center justify-center flex-shrink-0">
            <User className="w-4 h-4 text-accent" />
          </div>
          <div className="flex-1 min-w-0">
            <p className="text-white text-sm whitespace-pre-wrap">{message.content}</p>
          </div>
        </div>
      </div>
    )
  }

  if (message.role === 'error') {
    return (
      <div className="rounded-xl p-4 bg-red-500/10 border border-red-500/30">
        <p className="text-red-400 text-sm">{message.content}</p>
      </div>
    )
  }

  // Assistant message
  return (
    <div className="message-assistant rounded-xl p-4">
      <div className="flex items-start gap-3">
        <div className="w-8 h-8 rounded-lg bg-success/10 flex items-center justify-center flex-shrink-0">
          <Bot className="w-4 h-4 text-success" />
        </div>
        <div className="flex-1 min-w-0">
          <p className="text-white text-sm whitespace-pre-wrap">{message.content}</p>
          
          {message.toolCalls && message.toolCalls.length > 0 && (
            <button
              onClick={() => setShowTools(!showTools)}
              className="flex items-center gap-1 mt-3 text-xs text-gray-500 hover:text-gray-300 transition-colors"
            >
              {showTools ? <ChevronDown className="w-3 h-3" /> : <ChevronRight className="w-3 h-3" />}
              <Wrench className="w-3 h-3" />
              <span>{message.toolCalls.length} {message.toolCalls.length > 1 ? 'acciones' : 'acción'}</span>
            </button>
          )}
          
          {showTools && message.toolCalls && (
            <div className="mt-2 space-y-2">
              {message.toolCalls.map((tc, idx) => (
                <div key={idx} className="text-xs bg-surface-light rounded-lg p-2">
                  <div className="font-mono text-accent">{tc.tool}</div>
                  <div className="text-gray-500 mt-1 truncate">
                    {JSON.stringify(tc.input)}
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>
    </div>
  )
}

function SuggestionButton({ onClick, text }) {
  return (
    <button
      onClick={onClick}
      className="block w-full px-4 py-2 text-sm text-gray-400 bg-surface-light rounded-lg 
        hover:bg-accent/10 hover:text-accent transition-colors text-left"
    >
      {text}
    </button>
  )
}
