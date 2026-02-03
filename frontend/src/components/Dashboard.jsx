import StatCard from './StatCard'
import DonutChart from './DonutChart'
import SummaryTable from './SummaryTable'

function Dashboard({ results, validatorType, onViewAllResults }) {
  // Calculate stats
  const total = results.length
  const passed = results.filter((r) => r.status === 'PASS').length
  const failed = total - passed
  const defaultRules = validatorType === 'QUOTE_WIN' ? 4 : 19
  const rulesCount = results.length > 0
    ? new Set(results.map((r) => r.rule_name)).size
    : defaultRules
  const sheetsCount = new Set(results.map((r) => r.sheet_name)).size
  const passRate = total > 0 ? ((passed / total) * 100).toFixed(1) : 0

  // Determine pass rate color
  const getPassRateColor = () => {
    if (passRate >= 70) return 'bg-success'
    if (passRate >= 50) return 'bg-accent-yellow'
    return 'bg-error'
  }

  // Chart data
  const passFailData = [
    { name: 'Passed', value: passed },
    { name: 'Failed', value: failed },
  ]
  const passFailColors = ['#4CAF50', '#EF5350']

  // Rules distribution
  const ruleData = results.reduce((acc, result) => {
    const match = result.rule_name.match(/Rule\s*(\d+)/)
    const shortName = match ? `Rule ${match[1]}` : result.rule_name.substring(0, 10)
    acc[shortName] = (acc[shortName] || 0) + 1
    return acc
  }, {})

  const ruleChartData = Object.entries(ruleData)
    .sort((a, b) => {
      const numA = parseInt(a[0].replace('Rule ', '')) || 999
      const numB = parseInt(b[0].replace('Rule ', '')) || 999
      return numA - numB
    })
    .slice(0, 8)
    .map(([name, value]) => ({ name, value }))

  const ruleColors = [
    '#4A7C59', '#8B956D', '#7C5295', '#F5D547',
    '#E8B4B8', '#6B8E23', '#20B2AA', '#CD853F'
  ]

  return (
    <div className="space-y-6">
      {/* Stats Row */}
      <div className="grid grid-cols-6 gap-4">
        <StatCard title="Total Checks" value={total} color="bg-primary" />
        <StatCard title="Passed" value={passed} color="bg-success" />
        <StatCard title="Failed" value={failed} color="bg-error" />
        <StatCard title="Rules Checked" value={rulesCount} color="bg-accent-olive" />
        <StatCard title="Sheets Checked" value={sheetsCount} color="bg-accent-purple" />
        <StatCard
          title="Pass Rate"
          value={`${passRate}%`}
          color={getPassRateColor()}
          showProgress
          progress={parseFloat(passRate)}
        />
      </div>

      {/* Charts Row */}
      {results.length > 0 && (
        <div className="grid grid-cols-2 gap-6">
          <DonutChart
            title="Validation Results Distribution"
            data={passFailData}
            colors={passFailColors}
          />
          <DonutChart
            title="Results by Rule"
            data={ruleChartData}
            colors={ruleColors}
          />
        </div>
      )}

      {/* Summary Table or Empty State */}
      {results.length > 0 ? (
        <>
          <SummaryTable title="Validation Summary" results={results} />

          {/* View All Button */}
          <div className="flex justify-end">
            <button
              onClick={onViewAllResults}
              className="px-5 py-2 bg-primary hover:bg-primary-hover text-white text-sm font-semibold rounded-lg transition-all-200"
            >
              View All Results →
            </button>
          </div>
        </>
      ) : (
        <div className="bg-card-bg rounded-xl p-12 shadow-sm text-center">
          <div className="w-20 h-20 mx-auto mb-4 bg-success-bg rounded-full flex items-center justify-center">
            <span className="text-4xl text-primary">?</span>
          </div>
          <h3 className="text-lg font-bold text-text-primary mb-2">
            No Validation Results Yet
          </h3>
          <p className="text-sm text-text-secondary">
            Select an Excel file and click 'Select File' to get started
          </p>
        </div>
      )}
    </div>
  )
}

export default Dashboard
