type Page = 'kpi' | 'qr'

function navHref(page: Page): string {
  return page === 'qr' ? '#/qr' : '#/kpi'
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
          <a
            className={`nav-link ${page === 'qr' ? 'is-active' : ''}`}
            href={navHref('qr')}
          >
            Make QR Code
          </a>
        </nav>
      </div>
    </header>
  )
}

