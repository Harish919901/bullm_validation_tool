import { PieChart, Pie, Cell, ResponsiveContainer, Legend, Tooltip } from 'recharts'

function DonutChart({ title, data, colors }) {
  const total = data.reduce((sum, item) => sum + item.value, 0)

  // Custom label for center
  const renderCustomizedLabel = ({ cx, cy }) => {
    return (
      <text
        x={cx}
        y={cy}
        fill="#1A1A2E"
        textAnchor="middle"
        dominantBaseline="central"
        className="text-2xl font-bold"
      >
        {total}
      </text>
    )
  }

  // Custom tooltip
  const CustomTooltip = ({ active, payload }) => {
    if (active && payload && payload.length) {
      const item = payload[0]
      const percentage = ((item.value / total) * 100).toFixed(1)
      return (
        <div className="bg-white px-3 py-2 rounded-lg shadow-lg border border-gray-100">
          <p className="text-sm font-medium text-text-primary">
            {item.name}: {item.value} ({percentage}%)
          </p>
        </div>
      )
    }
    return null
  }

  // Custom legend
  const renderLegend = (props) => {
    const { payload } = props
    return (
      <div className="flex flex-col gap-2 ml-4">
        {payload.map((entry, index) => {
          const percentage = ((entry.payload.value / total) * 100).toFixed(1)
          return (
            <div key={`legend-${index}`} className="flex items-center gap-2">
              <div
                className="w-3 h-3 rounded-sm"
                style={{ backgroundColor: entry.color }}
              />
              <span className="text-xs text-text-secondary">
                {entry.value} ({percentage}%)
              </span>
            </div>
          )
        })}
      </div>
    )
  }

  if (total === 0) {
    return (
      <div className="bg-card-bg rounded-xl p-5 shadow-sm">
        <h3 className="text-sm font-bold text-text-primary mb-4">{title}</h3>
        <div className="flex items-center justify-center h-48 text-text-muted">
          No Data
        </div>
      </div>
    )
  }

  return (
    <div className="bg-card-bg rounded-xl p-5 shadow-sm">
      <h3 className="text-sm font-bold text-text-primary mb-4">{title}</h3>
      <div className="h-64">
        <ResponsiveContainer width="100%" height="100%">
          <PieChart>
            <Pie
              data={data}
              cx="40%"
              cy="50%"
              innerRadius={50}
              outerRadius={80}
              paddingAngle={2}
              dataKey="value"
              labelLine={false}
              label={renderCustomizedLabel}
            >
              {data.map((entry, index) => (
                <Cell
                  key={`cell-${index}`}
                  fill={colors[index % colors.length]}
                  stroke={colors[index % colors.length]}
                />
              ))}
            </Pie>
            <Tooltip content={<CustomTooltip />} />
            <Legend
              layout="vertical"
              align="right"
              verticalAlign="middle"
              content={renderLegend}
            />
          </PieChart>
        </ResponsiveContainer>
      </div>
    </div>
  )
}

export default DonutChart
