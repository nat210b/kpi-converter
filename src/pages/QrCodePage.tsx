import { useEffect, useMemo, useState } from 'react'
import QRCode from 'qrcode'

type ExpiryMode = 'permanent' | 'expires'
type QrSize = 224 | 256 | 320
type LinkMode = 'this_app' | 'custom'

function getAppBaseUrl(): string {
  // Keep origin + pathname only. Hash/query are controlled by us.
  return `${window.location.origin}${window.location.pathname}`
}

function getDefaultExpiresAtLocal(): string {
  const oneWeekMs = 7 * 24 * 60 * 60 * 1000
  const d = new Date(Date.now() + oneWeekMs)

  const pad2 = (n: number) => String(n).padStart(2, '0')
  const yyyy = d.getFullYear()
  const mm = pad2(d.getMonth() + 1)
  const dd = pad2(d.getDate())
  const hh = pad2(d.getHours())
  const min = pad2(d.getMinutes())
  return `${yyyy}-${mm}-${dd}T${hh}:${min}`
}

function buildKpiLink(expiryMode: ExpiryMode, expiresAtLocal: string): string {
  const base = getAppBaseUrl()
  if (expiryMode === 'permanent') return `${base}#/kpi`

  const expMs = new Date(expiresAtLocal).getTime()
  // If invalid date, fall back to permanent to avoid generating broken links.
  if (!Number.isFinite(expMs)) return `${base}#/kpi`

  return `${base}#/kpi?exp=${encodeURIComponent(String(expMs))}`
}

function normalizeCustomLink(raw: string): { ok: true; value: string } | { ok: false; error: string } {
  const trimmed = (raw || '').trim()
  if (!trimmed) return { ok: false, error: 'Link is required' }

  // Allow users to paste full URLs without forcing them to type scheme.
  const maybeUrl = trimmed.match(/^[a-zA-Z][a-zA-Z0-9+.-]*:\/\//)
    ? trimmed
    : `https://${trimmed}`

  try {
    const u = new URL(maybeUrl)
    if (u.protocol !== 'http:' && u.protocol !== 'https:') {
      return { ok: false, error: 'Only http/https links are allowed' }
    }
    return { ok: true, value: u.toString() }
  } catch {
    return { ok: false, error: 'Invalid link' }
  }
}

function clampQrDarkColor(hex: string): { color: string; warning: string } {
  const m = (hex || '').trim().match(/^#([0-9a-fA-F]{6})$/)
  if (!m) return { color: '#111111', warning: 'Invalid color; using default' }

  const rgb = m[1]
  const r = parseInt(rgb.slice(0, 2), 16)
  const g = parseInt(rgb.slice(2, 4), 16)
  const b = parseInt(rgb.slice(4, 6), 16)

  // Simple luminance check to avoid too-light QR modules.
  const luminance = (0.2126 * r + 0.7152 * g + 0.0722 * b) / 255
  if (luminance > 0.78) {
    return { color: '#111111', warning: 'Color was too light; using a darker color for scan reliability' }
  }

  return { color: `#${rgb.toLowerCase()}`, warning: '' }
}

export default function QrCodePage() {
  const [linkMode, setLinkMode] = useState<LinkMode>('this_app')
  const [customLink, setCustomLink] = useState('')

  const [expiryMode, setExpiryMode] = useState<ExpiryMode>('permanent')
  const [expiresAtLocal, setExpiresAtLocal] = useState(getDefaultExpiresAtLocal())

  const [size, setSize] = useState<QrSize>(256)
  const [darkColor, setDarkColor] = useState('#111111')

  const [dataUrl, setDataUrl] = useState<string>('')
  const [copyStatus, setCopyStatus] = useState<'idle' | 'copied' | 'failed'>('idle')
  const [qrWarning, setQrWarning] = useState<string>('')

  const thisAppLink = useMemo(
    () => buildKpiLink(expiryMode, expiresAtLocal),
    [expiryMode, expiresAtLocal]
  )

  // Initialize the custom link field once with a safe default.
  useEffect(() => {
    setCustomLink((prev) => prev || thisAppLink)
  }, [thisAppLink])

  const linkError = useMemo(() => {
    if (linkMode === 'this_app') return ''
    const norm = normalizeCustomLink(customLink)
    return norm.ok ? '' : norm.error
  }, [linkMode, customLink])

  const link = useMemo(() => {
    if (linkMode === 'this_app') return thisAppLink
    const norm = normalizeCustomLink(customLink)
    return norm.ok ? norm.value : ''
  }, [linkMode, thisAppLink, customLink])

  const palette = useMemo(() => {
    const { color, warning } = clampQrDarkColor(darkColor)
    setQrWarning(warning)
    return { dark: color, light: '#ffffff' }
  }, [darkColor])

  useEffect(() => {
    let cancelled = false
    async function run() {
      try {
        if (!link) {
          if (!cancelled) setDataUrl('')
          return
        }
        const url = await QRCode.toDataURL(link, {
          errorCorrectionLevel: 'M',
          width: size,
          margin: 2,
          color: palette,
        })
        if (!cancelled) setDataUrl(url)
      } catch {
        if (!cancelled) setDataUrl('')
      }
    }
    run()
    return () => {
      cancelled = true
    }
  }, [link, size, palette])

  async function handleCopy() {
    try {
      await navigator.clipboard.writeText(link)
      setCopyStatus('copied')
      window.setTimeout(() => setCopyStatus('idle'), 1200)
    } catch {
      setCopyStatus('failed')
      window.setTimeout(() => setCopyStatus('idle'), 1500)
    }
  }

  return (
    <div className="container">
      <div className="header">
        <h1 className="title">Make QR Code</h1>
        <p className="subtitle">
          Generates a QR you can scan on any device. Expires is enforced by this app when opened.
        </p>
      </div>

      <div className="card qr-card">
        <div className="qr-layout">
          <div className="qr-preview">
            <div className="qr-frame">
              {dataUrl ? (
                <img
                  className="qr-img"
                  src={dataUrl}
                  alt="QR code"
                  style={{ width: size, height: size, maxWidth: '100%', maxHeight: '100%' }}
                />
              ) : (
                <div className="qr-fallback">{linkError ? linkError : 'QR generation failed'}</div>
              )}
            </div>

            <div className="qr-link">
              <div className="qr-link-label">QR link</div>
              <div className="qr-link-row">
                <input className="qr-link-input" readOnly value={link || ''} />
                <button className="btn-secondary" onClick={handleCopy} disabled={!link}>
                  {copyStatus === 'copied' ? 'Copied' : copyStatus === 'failed' ? 'Copy failed' : 'Copy'}
                </button>
              </div>
            </div>
          </div>

          <div className="qr-controls">
            <div className="control">
              <div className="control-label">Link</div>
              <select
                className="control-select"
                value={linkMode}
                onChange={(e) => setLinkMode(e.target.value as LinkMode)}
              >
                <option value="this_app">This app (KPI tool)</option>
                <option value="custom">Custom</option>
              </select>
            </div>

            {linkMode === 'custom' && (
              <div className="control">
                <div className="control-label">Custom link (http/https)</div>
                <input
                  className="control-input"
                  value={customLink}
                  onChange={(e) => setCustomLink(e.target.value)}
                  placeholder="https://example.com"
                />
                {linkError && <div className="control-help is-error">{linkError}</div>}
              </div>
            )}

            <div className="control">
              <div className="control-label">Type</div>
              <select
                className="control-select"
                value={expiryMode}
                onChange={(e) => setExpiryMode(e.target.value as ExpiryMode)}
                disabled={linkMode === 'custom'}
              >
                <option value="permanent">Permanent</option>
                <option value="expires">Expires</option>
              </select>
              {linkMode === 'custom' && (
                <div className="control-help">
                  Expiration is only available for links that point back to this app.
                </div>
              )}
            </div>

            {expiryMode === 'expires' && linkMode === 'this_app' && (
              <div className="control">
                <div className="control-label">Expires at</div>
                <input
                  className="control-input"
                  type="datetime-local"
                  value={expiresAtLocal}
                  onChange={(e) => setExpiresAtLocal(e.target.value)}
                />
              </div>
            )}

            <div className="control">
              <div className="control-label">Size</div>
              <select
                className="control-select"
                value={String(size)}
                onChange={(e) => setSize(Number(e.target.value) as QrSize)}
              >
                <option value="224">Small</option>
                <option value="256">Medium</option>
                <option value="320">Large</option>
              </select>
            </div>

            <div className="control">
              <div className="control-label">QR color</div>
              <div className="control-row">
                <input
                  className="control-color"
                  type="color"
                  value={darkColor}
                  onChange={(e) => setDarkColor(e.target.value)}
                  aria-label="QR dark color"
                />
                <input
                  className="control-input"
                  value={darkColor}
                  onChange={(e) => setDarkColor(e.target.value)}
                  placeholder="#111111"
                />
              </div>
              {qrWarning && <div className="control-help">{qrWarning}</div>}
            </div>

            <div className="qr-note">
              Minimal customization on purpose: extreme styling often makes QR codes harder to scan.
            </div>
          </div>
        </div>
      </div>
    </div>
  )
}
