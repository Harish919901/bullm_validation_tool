function StatCard({ title, value, color = 'bg-primary', showProgress = false, progress = 0 }) {
  return (
    <div className="bg-card-bg rounded-xl p-5 shadow-sm">
      {/* Title */}
      <p className="text-xs text-text-secondary mb-2">{title}</p>

      {/* Value with indicator */}
      <div className="flex items-center">
        <div className={`w-1 h-9 ${color} rounded-full mr-3`} />
        <span className="text-3xl font-bold text-text-primary">{value}</span>
      </div>

      {/* Progress bar (optional) */}
      {showProgress && (
        <div className="mt-3">
          <div className="h-2 bg-gray-200 rounded-full overflow-hidden">
            <div
              className={`h-full ${color} rounded-full transition-all duration-500`}
              style={{ width: `${Math.min(100, Math.max(0, progress))}%` }}
            />
          </div>
        </div>
      )}
    </div>
  )
}

export default StatCard
