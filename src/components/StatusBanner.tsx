import type { AppStatus } from '../types'

interface Props { status: AppStatus }

const ICONS: Record<string, string> = {
  loading: '⏳', success: '✅', error: '❌', idle: '',
}

export default function StatusBanner({ status }: Props) {
  if (status.type === 'idle') return null
  return (
    <div className={`status-banner status-${status.type}`}>
      <span>{ICONS[status.type]}</span>
      <span>{status.message}</span>
    </div>
  )
}
