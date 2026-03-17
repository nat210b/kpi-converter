import type { Page } from '../App'

function navHref(page: Page): string {
  if (page === 'qr') return '#/qr'
  if (page === 'checkpoint') return '#/checkpoint'
  return '#/kpi'
}

export default function TopNav({ page }: { page: Page }) {
  return (
    <header className="top-nav">
      <div className="nav-inner">
        <a className="brand" href={navHref('kpi')} aria-label="N4T Toolz Home">
          <span className="brand-name">N4T Toolz</span>
          <span className="brand-tag">tools that stay simple</span>
        </a>

        <nav className="nav-menu" aria-label="Primary">
          <a className={`nav-link ${page === 'checkpoint' ? 'is-active' : ''}`} href={navHref('checkpoint')}>
            หยอดผลงาน
          </a>
          <a className={`nav-link ${page === 'qr' ? 'is-active' : ''}`} href={navHref('qr')}>
            Make QR Code
          </a>
        </nav>
      </div>
    </header>
  )
}
