function ResultCard({ result }) {
  const isPass = result.status === 'PASS'

  const truncate = (str, maxLen = 80) => {
    if (!str) return ''
    return str.length > maxLen ? str.substring(0, maxLen) + '...' : str
  }

  return (
    <div className="bg-card-bg rounded-lg p-4 shadow-sm">
      <div className="flex">
        {/* Left Indicator */}
        <div
          className={`w-1 rounded-full mr-4 ${isPass ? 'bg-success' : 'bg-error'}`}
        />

        {/* Content */}
        <div className="flex-1">
          {/* Header Row */}
          <div className="flex items-center justify-between mb-3">
            <h4 className="text-sm font-bold text-text-primary">
              {result.rule_name}
            </h4>
            <span
              className={`px-3 py-1 text-xs font-bold rounded ${
                isPass
                  ? 'bg-success-bg text-success'
                  : 'bg-error-bg text-error'
              }`}
            >
              {result.status}
            </span>
          </div>

          {/* Details */}
          <div className="space-y-1 text-xs">
            <div className="flex">
              <span className="w-16 text-text-muted">Sheet:</span>
              <span className="text-text-secondary">{result.sheet_name}</span>
            </div>
            <div className="flex">
              <span className="w-16 text-text-muted">Expected:</span>
              <span className="text-text-secondary flex-1">
                {truncate(result.expected)}
              </span>
            </div>
            <div className="flex">
              <span className="w-16 text-text-muted">Actual:</span>
              <span className="text-text-secondary flex-1">
                {truncate(result.actual)}
              </span>
            </div>
            {result.location && (
              <div className="flex">
                <span className="w-16 text-text-muted">Location:</span>
                <span className="text-text-secondary">{result.location}</span>
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  )
}

export default ResultCard
