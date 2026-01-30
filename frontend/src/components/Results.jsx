import ResultCard from './ResultCard'

function Results({ results, filterStatus, onFilterChange }) {
  // Filter results
  const filteredResults = results.filter((r) => {
    if (filterStatus === 'all') return true
    if (filterStatus === 'pass') return r.status === 'PASS'
    if (filterStatus === 'fail') return r.status === 'FAIL'
    return true
  })

  // Group by rule
  const groupedResults = filteredResults.reduce((acc, result) => {
    if (!acc[result.rule_name]) {
      acc[result.rule_name] = []
    }
    acc[result.rule_name].push(result)
    return acc
  }, {})

  const filters = [
    { key: 'all', label: 'All' },
    { key: 'pass', label: 'Passed' },
    { key: 'fail', label: 'Failed' },
  ]

  return (
    <div className="space-y-5">
      {/* Filter Bar */}
      <div className="bg-card-bg rounded-xl p-4 shadow-sm">
        <div className="flex items-center gap-4">
          <span className="text-sm font-medium text-text-primary">
            Filter Results:
          </span>
          {filters.map((filter) => (
            <button
              key={filter.key}
              onClick={() => onFilterChange(filter.key)}
              className={`px-4 py-2 text-sm font-medium rounded-lg transition-all-200 ${
                filterStatus === filter.key
                  ? 'bg-primary text-white'
                  : 'bg-transparent text-text-secondary hover:bg-success-bg'
              }`}
            >
              {filter.label}
            </button>
          ))}
        </div>
      </div>

      {/* Results */}
      {filteredResults.length === 0 ? (
        <div className="bg-card-bg rounded-xl p-8 shadow-sm text-center">
          <p className="text-text-secondary">
            {results.length === 0
              ? 'No results to display. Run validation first.'
              : 'No results match the selected filter.'}
          </p>
        </div>
      ) : (
        Object.entries(groupedResults).map(([ruleName, ruleResults]) => {
          const passed = ruleResults.filter((r) => r.status === 'PASS').length
          const failed = ruleResults.length - passed

          return (
            <div key={ruleName} className="bg-card-bg rounded-xl shadow-sm overflow-hidden">
              {/* Rule Header */}
              <div className="px-5 py-4 border-b border-gray-100 flex items-center justify-between">
                <h3 className="text-sm font-bold text-text-primary">{ruleName}</h3>
                <span className="text-xs text-text-muted">
                  {passed} passed, {failed} failed
                </span>
              </div>

              {/* Rule Results */}
              <div className="p-3 space-y-3">
                {ruleResults.map((result, index) => (
                  <ResultCard key={index} result={result} />
                ))}
              </div>
            </div>
          )
        })
      )}
    </div>
  )
}

export default Results
