import { useCallback, useEffect, useState } from 'react'
import * as XLSX from 'xlsx'

import DropZone from './components/DropZone'
import StatGrid from './components/StatGrid'
import StatusBanner from './components/StatusBanner'
import TopNav from './components/TopNav'
import QrCodePage from './pages/QrCodePage'

import { parseKpiWorkbook, scanKpiSections } from './utils/kpiParser'
import { buildCsv, downloadCsv } from './utils/csvBuilder'

import type { AppStatus, ConvertResult, KpiSectionInfo } from './types'

type Page = 'kpi' | 'qr'

/* ---------------- ROUTER ---------------- */

function parseHash(hash: string): { page: Page; params: URLSearchParams } {
  const raw = (hash || '#/kpi').replace(/^#/, '')
  const [path, query = ''] = raw.split('?')

  const page: Page = path.startsWith('/qr') ? 'qr' : 'kpi'

  return {
    page,
    params: new URLSearchParams(query),
  }
}

function useHashRoute() {
  const [route, setRoute] = useState(() => parseHash(window.location.hash))

  useEffect(() => {
    const handler = () => setRoute(parseHash(window.location.hash))
    window.addEventListener('hashchange', handler)

    return () => window.removeEventListener('hashchange', handler)
  }, [])

  return route
}

/* ---------------- APP ---------------- */

export default function App() {
  const { page, params } = useHashRoute()

  const [fileName, setFileName] = useState('')
  const [fileSize, setFileSize] = useState('')

  const [status, setStatus] = useState<AppStatus>({
    type: 'idle',
    message: '',
  })

  const [result, setResult] = useState<ConvertResult | null>(null)
  const [csvContent, setCsvContent] = useState<string | null>(null)

  const [startDate, setStartDate] = useState('1/1/2025')
  const [endDate, setEndDate] = useState('12/31/2025')

  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null)
  const [loading, setLoading] = useState(false)

  const [kpiSections, setKpiSections] = useState<KpiSectionInfo[]>([])
  const [selectedSections, setSelectedSections] = useState<Record<string, boolean>>({})

  /* ---------------- FILE UPLOAD ---------------- */

  const handleFile = useCallback((file: File) => {
    if (!file.name.match(/\.xlsx?$/i)) {
      setStatus({
        type: 'error',
        message: 'กรุณาเลือกไฟล์ .xlsx หรือ .xls เท่านั้น',
      })
      return
    }

    setFileName(file.name)
    setFileSize((file.size / 1024).toFixed(1) + ' KB')

    setStatus({
      type: 'loading',
      message: 'กำลังโหลดไฟล์...',
    })

    setResult(null)
    setCsvContent(null)
    setWorkbook(null)
    setKpiSections([])
    setSelectedSections({})

    const reader = new FileReader()

    reader.onload = (e) => {
      try {
        const data = e.target?.result

        const wb = XLSX.read(data, { type: 'array' })

        const hasKpiSheet = wb.SheetNames.some((s) => s.includes('KPI'))

        if (!hasKpiSheet) {
          setStatus({
            type: 'error',
            message: 'ไม่พบ sheet ที่มีชื่อ "KPI" ในไฟล์นี้',
          })
          return
        }

        setWorkbook(wb)

        const sections = scanKpiSections(wb)
        setKpiSections(sections)
        setSelectedSections(Object.fromEntries(sections.map((s) => [s.name, true])))

        setStatus({
          type: 'success',
          message: `โหลดสำเร็จ — พบ ${wb.SheetNames.length} sheet กด "แปลงไฟล์" เพื่อดำเนินการต่อ`,
        })
      } catch (err) {
        setStatus({
          type: 'error',
          message: 'อ่านไฟล์ไม่ได้: ' + (err as Error).message,
        })
      }
    }

    reader.readAsArrayBuffer(file)
  }, [])

  /* ---------------- CONVERT KPI ---------------- */

  const handleConvert = useCallback(() => {
    if (!workbook) return

    setLoading(true)

    setStatus({
      type: 'loading',
      message: 'กำลังแปลงข้อมูล...',
    })

    try {
      const selected = new Set(
        Object.entries(selectedSections)
          .filter(([, on]) => on)
          .map(([name]) => name)
      )

      if (kpiSections.length > 0 && selected.size === 0) {
        setStatus({
          type: 'error',
          message: 'กรุณาเลือกประเภท KPI อย่างน้อย 1 ประเภท',
        })

        setLoading(false)
        return
      }

      const parsed = parseKpiWorkbook(
        workbook,
        kpiSections.length ? { includeSections: selected } : undefined
      )

      if (parsed.rowCount === 0) {
        setStatus({
          type: 'error',
          message: 'ไม่พบข้อมูล KPI ในไฟล์ กรุณาตรวจสอบ sheet',
        })

        setLoading(false)
        return
      }

      const csv = buildCsv(parsed.rows, startDate, endDate)

      setResult(parsed)
      setCsvContent(csv)

      setStatus({
        type: 'success',
        message: `แปลงสำเร็จ — ${parsed.rowCount} แถว จาก ${parsed.empCount} พนักงาน`,
      })
    } catch (err) {
      setStatus({
        type: 'error',
        message: 'เกิดข้อผิดพลาด: ' + (err as Error).message,
      })
    }

    setLoading(false)
  }, [workbook, startDate, endDate, selectedSections, kpiSections])

  /* ---------------- DOWNLOAD ---------------- */

  const handleDownload = useCallback(() => {
    if (!csvContent) return
    downloadCsv(csvContent, 'filled_sheet.csv')
  }, [csvContent])

  /* ---------------- QR EXPIRATION ---------------- */

  const exp = params.get('exp')

  const expMs = exp ? Number(exp) : NaN

  const isExpired = Number.isFinite(expMs) && Date.now() > expMs

  const selectedCount = Object.values(selectedSections).filter(Boolean).length

  const canConvert =
    Boolean(workbook) &&
    !loading &&
    (kpiSections.length === 0 || selectedCount > 0)

  function setAllSections(on: boolean) {
    setSelectedSections((prev) => {
      const next: Record<string, boolean> = { ...prev }
      for (const s of kpiSections) next[s.name] = on
      return next
    })
  }

  /* ---------------- UI ---------------- */

  return (
    <div className="app-shell">

      <TopNav page={page} />

      <main className="page">

        {page === 'qr' ? (
          <QrCodePage />

        ) : isExpired ? (

          <div className="container">

            <div className="header">
              <h1 className="title">QR code หมดอายุแล้ว</h1>
              <p className="subtitle">
                ลิงก์นี้มีเวลาหมดอายุ กรุณาขอ QR ใหม่ หรือสร้างใหม่ได้ด้านล่าง
              </p>
            </div>

            <div className="card">
              <div className="actions">

                <a className="btn-primary" href="#/qr">
                  สร้าง QR ใหม่
                </a>

                <a className="btn-secondary" href="#/kpi">
                  ไปหน้า KPI Tool
                </a>

              </div>
            </div>

          </div>

        ) : (

          <div className="container">

            <div className="header">
              <h1 className="title">KPI → filled_sheet.csv</h1>
              <p className="subtitle">
                อัพโหลดไฟล์ KPI Excel แล้วดาวน์โหลด filled_sheet.csv พร้อมใช้งานได้เลย
              </p>
            </div>

            <div className="card">

              <DropZone
                onFile={handleFile}
                disabled={loading}
                fileName={fileName}
                fileSize={fileSize}
              />

              <StatusBanner status={status} />

              {workbook && kpiSections.length > 0 && (
                <div className="section-filter">
                  <div className="filter-head">
                    <div className="filter-title">เลือกประเภท KPI</div>
                    <div className="filter-actions">
                      <button
                        type="button"
                        className="filter-link"
                        onClick={() => setAllSections(true)}
                      >
                        เลือกทั้งหมด
                      </button>
                      <button
                        type="button"
                        className="filter-link"
                        onClick={() => setAllSections(false)}
                      >
                        ไม่เลือกทั้งหมด
                      </button>
                    </div>
                  </div>

                  <div className="filter-sub">
                    พบ {kpiSections.length} ประเภท (Turnover/Time/Total ถูกตัดออกอัตโนมัติ)
                  </div>

                  <div className="filter-list" role="group" aria-label="KPI types">
                    {kpiSections.map((s) => (
                      <label key={s.name} className="filter-item">
                        <input
                          type="checkbox"
                          checked={selectedSections[s.name] ?? true}
                          onChange={(e) =>
                            setSelectedSections((prev) => ({
                              ...prev,
                              [s.name]: e.target.checked,
                            }))
                          }
                        />
                        <span className="filter-name">{s.name}</span>
                        <span className="filter-count">{s.kpiCount}</span>
                      </label>
                    ))}
                  </div>
                </div>
              )}

              {result && (
                <StatGrid
                  stats={[
                    { value: result.empCount, label: 'พนักงาน' },
                    { value: result.kpiCount, label: 'KPI' },
                    { value: result.rowCount, label: 'แถวทั้งหมด' },
                  ]}
                />
              )}

              <hr className="divider" />

              <div className="settings">

                <div className="setting-row">
                  <div>
                    <div className="setting-label">วันที่เริ่มต้น</div>
                    <div className="setting-hint">Start Date (mm/dd/yyyy)</div>
                  </div>

                  <input
                    type="text"
                    className="date-input"
                    value={startDate}
                    onChange={(e) => setStartDate(e.target.value)}
                    placeholder="1/1/2025"
                  />
                </div>

                <div className="setting-row">
                  <div>
                    <div className="setting-label">วันที่สิ้นสุด</div>
                    <div className="setting-hint">End Date (mm/dd/yyyy)</div>
                  </div>

                  <input
                    type="text"
                    className="date-input"
                    value={endDate}
                    onChange={(e) => setEndDate(e.target.value)}
                    placeholder="12/31/2025"
                  />
                </div>

              </div>

              <div className="actions">

                <button
                  className="btn-primary"
                  onClick={handleConvert}
                  disabled={!canConvert}
                >
                  {loading ? '⏳ กำลังแปลง...' : '⚙ แปลงไฟล์'}
                </button>

                {csvContent && (
                  <button
                    className="btn-success"
                    onClick={handleDownload}
                  >
                    ⬇ ดาวน์โหลด filled_sheet.csv
                  </button>
                )}

              </div>

            </div>

            <p className="footer">
              Turnover & Time KPIs ถูกตัดออกอัตโนมัติ · Levels จาก column M–Q ต่อ KPI
            </p>

          </div>

        )}

      </main>

    </div>
  )
}
