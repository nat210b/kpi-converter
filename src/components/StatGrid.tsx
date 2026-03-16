interface Stat { value: number; label: string }
interface Props { stats: Stat[] }

export default function StatGrid({ stats }: Props) {
  return (
    <div className="stat-grid">
      {stats.map((s) => (
        <div key={s.label} className="stat-card">
          <div className="stat-value">{s.value.toLocaleString()}</div>
          <div className="stat-label">{s.label}</div>
        </div>
      ))}
    </div>
  )
}
