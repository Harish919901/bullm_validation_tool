import { useState } from 'react'

const navItems = [
  { key: 'dashboard', label: 'Dashboard', icon: '⌂' },
  { key: 'results', label: 'Results', icon: '✓' },
  { key: 'rules', label: 'Rules Info', icon: 'ℹ' },
]

function Sidebar({ currentView, onViewChange, onExportCSV, onExportExcel, hasResults, sessionId }) {
  return (
    <aside className="w-64 bg-sidebar-bg flex flex-col">
      {/* Logo Section */}
      <div className="flex justify-center py-6">
        <img src="/Logo.png" alt="Logo" className="h-14 w-auto" />
      </div>

      {/* Navigation Section */}
      <div className="px-4">
        <p className="text-text-muted text-xs uppercase tracking-wider mb-3 px-3">
          Navigation
        </p>
        <nav className="space-y-1">
          {navItems.map((item) => (
            <button
              key={item.key}
              onClick={() => onViewChange(item.key)}
              className={`w-full flex items-center px-4 py-3 rounded-lg text-sm font-medium transition-all-200 ${
                currentView === item.key
                  ? 'bg-sidebar-active text-white'
                  : 'text-gray-300 hover:bg-sidebar-hover hover:text-white'
              }`}
            >
              <span className="mr-3 text-lg">{item.icon}</span>
              {item.label}
            </button>
          ))}
        </nav>
      </div>

      {/* Export Section */}
      <div className="px-4 mt-10">
        <p className="text-text-muted text-xs uppercase tracking-wider mb-3 px-3">
          Export
        </p>
        <div className="space-y-1">
          <button
            onClick={onExportCSV}
            disabled={!hasResults}
            className={`w-full flex items-center px-4 py-3 rounded-lg text-sm font-medium transition-all-200 ${
              hasResults
                ? 'text-gray-300 hover:bg-sidebar-hover hover:text-white'
                : 'text-gray-500 cursor-not-allowed'
            }`}
          >
            <span className="mr-3 text-lg">⇩</span>
            Export CSV
          </button>
          <button
            onClick={onExportExcel}
            disabled={!hasResults}
            className={`w-full flex items-center px-4 py-3 rounded-lg text-sm font-medium transition-all-200 ${
              hasResults
                ? 'text-gray-300 hover:bg-sidebar-hover hover:text-white'
                : 'text-gray-500 cursor-not-allowed'
            }`}
          >
            <span className="mr-3 text-lg">⇩</span>
            Export Excel
          </button>
        </div>
      </div>

      {/* Spacer */}
      <div className="flex-1" />

      {/* Version */}
      <div className="px-4 py-6 text-center">
        <p className="text-text-muted text-xs">
          v1.0.0 | React + FastAPI
        </p>
      </div>
    </aside>
  )
}

export default Sidebar
