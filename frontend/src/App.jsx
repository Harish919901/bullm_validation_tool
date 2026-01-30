import { useState } from 'react'
import Sidebar from './components/Sidebar'
import Header from './components/Header'
import Dashboard from './components/Dashboard'
import Results from './components/Results'
import RulesInfo from './components/RulesInfo'
import { useValidation } from './hooks/useValidation'

function App() {
  const [currentView, setCurrentView] = useState('dashboard')
  const [validatorType, setValidatorType] = useState('QUOTE_WIN')
  const [filterStatus, setFilterStatus] = useState('all')

  const {
    results,
    loading,
    error,
    fileName,
    sessionId,
    uploadAndValidate,
    exportCSV,
    exportExcel,
    saveValidatedFile,
    clearResults
  } = useValidation()

  const handleValidatorTypeChange = (type) => {
    setValidatorType(type)
    clearResults()
  }

  const handleFileUpload = async (file) => {
    await uploadAndValidate(file, validatorType)
  }

  const renderContent = () => {
    switch (currentView) {
      case 'dashboard':
        return (
          <Dashboard
            results={results}
            validatorType={validatorType}
            onViewAllResults={() => setCurrentView('results')}
          />
        )
      case 'results':
        return (
          <Results
            results={results}
            filterStatus={filterStatus}
            onFilterChange={setFilterStatus}
          />
        )
      case 'rules':
        return <RulesInfo validatorType={validatorType} />
      default:
        return null
    }
  }

  return (
    <div className="flex h-screen bg-main-bg">
      {/* Sidebar */}
      <Sidebar
        currentView={currentView}
        onViewChange={setCurrentView}
        onExportCSV={exportCSV}
        onExportExcel={exportExcel}
        hasResults={results.length > 0}
        sessionId={sessionId}
      />

      {/* Main Content */}
      <div className="flex-1 flex flex-col overflow-hidden">
        {/* Header */}
        <Header
          validatorType={validatorType}
          onValidatorTypeChange={handleValidatorTypeChange}
          onFileUpload={handleFileUpload}
          onSaveFile={saveValidatedFile}
          loading={loading}
          fileName={fileName}
          results={results}
          sessionId={sessionId}
        />

        {/* Content Area */}
        <main className="flex-1 overflow-auto p-8">
          {error && (
            <div className="mb-4 p-4 bg-error-bg text-error rounded-lg">
              {error}
            </div>
          )}
          {renderContent()}
        </main>
      </div>
    </div>
  )
}

export default App
