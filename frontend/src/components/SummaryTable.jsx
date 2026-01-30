function SummaryTable({ title, results }) {
  const truncate = (str, maxLen = 40) => {
    if (!str) return ''
    return str.length > maxLen ? str.substring(0, maxLen) + '...' : str
  }

  // Extract rule number from rule name
  const getRuleNum = (ruleName) => {
    const match = ruleName.match(/Rule\s*(\d+)/)
    return match ? match[1] : ruleName.substring(0, 10)
  }

  return (
    <div className="bg-card-bg rounded-xl p-5 shadow-sm">
      <h3 className="text-sm font-bold text-text-primary mb-4">{title}</h3>

      {/* Table */}
      <div className="overflow-x-auto">
        <table className="w-full">
          <thead>
            <tr className="bg-gray-100 rounded-lg">
              <th className="text-left text-xs font-bold text-text-primary px-4 py-3 rounded-l-lg">
                Rule
              </th>
              <th className="text-left text-xs font-bold text-text-primary px-4 py-3">
                Sheet
              </th>
              <th className="text-left text-xs font-bold text-text-primary px-4 py-3">
                Status
              </th>
              <th className="text-left text-xs font-bold text-text-primary px-4 py-3">
                Expected
              </th>
              <th className="text-left text-xs font-bold text-text-primary px-4 py-3 rounded-r-lg">
                Location
              </th>
            </tr>
          </thead>
          <tbody>
            {results.slice(0, 10).map((result, index) => (
              <tr key={index} className="border-b border-gray-100 last:border-b-0">
                <td className="text-xs text-text-secondary px-4 py-3">
                  {getRuleNum(result.rule_name)}
                </td>
                <td className="text-xs text-text-secondary px-4 py-3">
                  {result.sheet_name}
                </td>
                <td className="px-4 py-3">
                  <span
                    className={`text-xs font-semibold ${
                      result.status === 'PASS' ? 'text-success' : 'text-error'
                    }`}
                  >
                    {result.status}
                  </span>
                </td>
                <td className="text-xs text-text-secondary px-4 py-3">
                  {truncate(result.expected, 35)}
                </td>
                <td className="text-xs text-text-secondary px-4 py-3">
                  {result.location || 'N/A'}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {results.length === 0 && (
        <div className="text-center py-8 text-text-muted text-sm">
          No results to display
        </div>
      )}
    </div>
  )
}

export default SummaryTable
