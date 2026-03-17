import { useCallback, useEffect, useState } from 'react'
import * as XLSX from 'xlsx'

import DropZone from './components/DropZone'
import StatGrid from './components/StatGrid'
import StatusBanner from './components/StatusBanner'
import TopNav from './components/TopNav'
import QrCodePage from './pages/QrCodePage'
import CheckpointPage from './pages/CheckpointPage'

import { autoDetectKpiSheet, parseKpiWorkbook, scanKpiSections, scanMissingLevels } from './utils/kpiParser'
import { buildCsv, downloadCsv } from './utils/csvBuilder'

import type { AppStatus, ConvertResult, KpiSectionInfo, MissingLevelKpi } from './types'

export type Page = 'kpi' | 'checkpoint' | 'qr'

/* ---------------- ROUTER ---------------- */

function parseHash(hash: string): { page: Page; params: URLSearchParams } {
  const raw = (hash || '#/kpi').replace(/^#/, '')
  const [path, query = ''] = raw.split('?')
  let page: Page = 'kpi'
  if (path.startsWith('/qr')) page = 'qr'
  else if (path.startsWith('/checkpoint')) page = 'checkpoint'
  return { page, params: new URLSearchParams(query) }
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

  const [fileName, setFileName]     = useState('')
  const [fileSize, setFileSize]     = useState('')
  const [status, setStatus]         = useState<AppStatus>({ type: 'idle', message: '' })
  const [result, setResult]         = useState<ConvertResult | null>(null)
  const [csvContent, setCsvContent] = useState<string | null>(null)
  const [startDate, setStartDate]   = useState('1/1/2025')
  const [endDate, setEndDate]       = useState('12/31/2025')
  const [workbook, setWorkbook]     = useState<XLSX.WorkBook | null>(null)
  const [loading, setLoading]       = useState(false)

  const [sheetNames, setSheetNames]         = useState<string[]>([])
  const [selectedSheet, setSelectedSheet]   = useState<string>('')
  const [kpiSections, setKpiSections]       = useState<KpiSectionInfo[]>([])
  const [selectedSections, setSelectedSections] = useState<Record<string, boolean>>({})

  // Missing level values
  const [missingLevels, setMissingLevels]   = useState<MissingLevelKpi[]>([])
  // User-edited level values: key = measureCode || kpiName
  const [levelEdits, setLevelEdits]         = useState<Record<string, { lvl1: string; lvl2: string; lvl3: string; lvl4: string; lvl5: string }>>({})

  /* ---------------- FILE UPLOAD ---------------- */

  const handleFile = useCallback((file: File) => {
    if (!file.name.match(/\.xlsx?$/i)) {
      setStatus({ type: 'error', message: 'กรุณาเลือกไฟล์ .xlsx หรือ .xls เท่านั้น' })
      return
    }
    setFileName(file.name)
    setFileSize((file.size / 1024).toFixed(1) + ' KB')
    setStatus({ type: 'loading', message: 'กำลังโหลดไฟล์...' })
    setResult(null); setCsvContent(null); setWorkbook(null)
    setSheetNames([]); setSelectedSheet(''); setKpiSections([]); setSelectedSections({})
    setMissingLevels([]); setLevelEdits({})

    const reader = new FileReader()
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target?.result, { type: 'array' })
        if (!wb.SheetNames.length) { setStatus({ type: 'error', message: 'ไฟล์นี้ไม่มี sheet เลย' }); return }
        const detected = autoDetectKpiSheet(wb)
        setWorkbook(wb); setSheetNames(wb.SheetNames); setSelectedSheet(detected)
        const sections = scanKpiSections(wb, detected)
        setKpiSections(sections)
        setSelectedSections(Object.fromEntries(sections.map((s) => [s.name, true])))
        setStatus({ type: 'success', message: `โหลดสำเร็จ — พบ ${wb.SheetNames.length} sheet · ใช้ "${detected}" กด "แปลงไฟล์" เพื่อดำเนินการต่อ` })
      } catch (err) {
        setStatus({ type: 'error', message: 'อ่านไฟล์ไม่ได้: ' + (err as Error).message })
      }
    }
    reader.readAsArrayBuffer(file)
  }, [])

  /* ---------------- SHEET CHANGE ---------------- */

  const handleSheetChange = useCallback((name: string) => {
    if (!workbook) return
    setSelectedSheet(name); setResult(null); setCsvContent(null)
    setMissingLevels([]); setLevelEdits({})
    try {
      const sections = scanKpiSections(workbook, name)
      setKpiSections(sections)
      setSelectedSections(Object.fromEntries(sections.map((s) => [s.name, true])))
      setStatus({ type: 'success', message: `เลือก sheet "${name}" แล้ว กด "แปลงไฟล์" เพื่อดำเนินการต่อ` })
    } catch (err) {
      setKpiSections([]); setSelectedSections({})
      setStatus({ type: 'error', message: (err as Error).message })
    }
  }, [workbook])

  /* ---------------- CONVERT KPI ---------------- */

  /** Step 1: scan for missing levels. If any found, show the table. Otherwise go straight to step 2. */
  const handleConvert = useCallback(() => {
    if (!workbook || !selectedSheet) return

    const selected = new Set(Object.entries(selectedSections).filter(([, on]) => on).map(([n]) => n))
    if (kpiSections.length > 0 && selected.size === 0) {
      setStatus({ type: 'error', message: 'กรุณาเลือกประเภท KPI อย่างน้อย 1 ประเภท' })
      return
    }

    const missing = scanMissingLevels(
      workbook,
      selectedSheet,
      kpiSections.length ? { includeSections: selected } : undefined
    )

    if (missing.length > 0) {
      // Pre-populate edits with whatever is already in the sheet
      const initial: typeof levelEdits = {}
      for (const m of missing) {
        const key = m.measureCode || m.kpiName
        initial[key] = { lvl1: m.lvl1, lvl2: m.lvl2, lvl3: m.lvl3, lvl4: m.lvl4, lvl5: m.lvl5 }
      }
      setLevelEdits(initial)
      setMissingLevels(missing)
      setStatus({ type: 'loading', message: `พบ ${missing.length} KPI ที่ยังไม่มีค่า Level — กรุณากรอกให้ครบแล้วกด "ยืนยันและแปลงไฟล์"` })
      return
    }

    runConvert(selected, new Map())
  }, [workbook, selectedSheet, selectedSections, kpiSections])

  /** Step 2: actually parse + build CSV, with optional level overrides from the table. */
  const runConvert = useCallback((
    selected: Set<string>,
    overrides: Map<string, { lvl1: string; lvl2: string; lvl3: string; lvl4: string; lvl5: string }>
  ) => {
    if (!workbook || !selectedSheet) return
    setLoading(true)
    setStatus({ type: 'loading', message: 'กำลังแปลงข้อมูล...' })
    try {
      const parsed = parseKpiWorkbook(
        workbook,
        selectedSheet,
        { includeSections: kpiSections.length ? selected : undefined, levelOverrides: overrides }
      )
      if (parsed.rowCount === 0) {
        setStatus({ type: 'error', message: 'ไม่พบข้อมูล KPI ในไฟล์ กรุณาตรวจสอบ sheet' })
        setLoading(false); return
      }
      const csv = buildCsv(parsed.rows, startDate, endDate)
      setResult(parsed); setCsvContent(csv)
      setMissingLevels([])  // clear table on success
      setStatus({ type: 'success', message: `แปลงสำเร็จ — ${parsed.rowCount} แถว จาก ${parsed.empCount} พนักงาน` })
    } catch (err) {
      setStatus({ type: 'error', message: 'เกิดข้อผิดพลาด: ' + (err as Error).message })
    }
    setLoading(false)
  }, [workbook, selectedSheet, startDate, endDate, kpiSections])

  /** Called when user clicks "ยืนยันและแปลงไฟล์" after filling the level table. */
  const handleConfirmConvert = useCallback(() => {
    const selected = new Set(Object.entries(selectedSections).filter(([, on]) => on).map(([n]) => n))
    const overrides = new Map(
      Object.entries(levelEdits).map(([k, v]) => [k, v])
    )
    runConvert(selected, overrides)
  }, [selectedSections, levelEdits, runConvert])

  /* ---------------- DOWNLOAD ---------------- */

  const handleDownload = useCallback(() => {
    if (!csvContent) return
    downloadCsv(csvContent, 'filled_sheet.csv')
  }, [csvContent])

  /* ---------------- HELPERS ---------------- */

  const exp = params.get('exp')
  const expMs = exp ? Number(exp) : NaN
  const isExpired = Number.isFinite(expMs) && Date.now() > expMs
  const selectedCount = Object.values(selectedSections).filter(Boolean).length
  const canConvert = Boolean(workbook) && Boolean(selectedSheet) && !loading && (kpiSections.length === 0 || selectedCount > 0)

  function setAllSections(on: boolean) {
    setSelectedSections((prev) => { const next = { ...prev }; for (const s of kpiSections) next[s.name] = on; return next })
  }

  /* ---------------- UI ---------------- */

  return (
    <div className="app-shell">
      <TopNav page={page} />
      <main className="page">

        {page === 'qr' ? (
          <QrCodePage />

        ) : page === 'checkpoint' ? (
          <CheckpointPage />

        ) : isExpired ? (
          <div className="container">
            <div className="header">
              <h1 className="title">QR code หมดอายุแล้ว</h1>
              <p className="subtitle">ลิงก์นี้มีเวลาหมดอายุ กรุณาขอ QR ใหม่ หรือสร้างใหม่ได้ด้านล่าง</p>
            </div>
            <div className="card">
              <div className="actions">
                <a className="btn-primary" href="#/qr">สร้าง QR ใหม่</a>
                <a className="btn-secondary" href="#/kpi">ไปหน้า KPI Tool</a>
              </div>
            </div>
          </div>

        ) : (
          <div className="container">
            <div className="header">
              <h1 className="title">KPI → filled_sheet.csv</h1>
              <p className="subtitle">อัพโหลดไฟล์ KPI Excel แล้วดาวน์โหลด filled_sheet.csv พร้อมใช้งานได้เลย</p>
            </div>
            <div className="card">

              <DropZone onFile={handleFile} disabled={loading} fileName={fileName} fileSize={fileSize} />
              <StatusBanner status={status} />

              {sheetNames.length > 0 && (
                <div className="sheet-selector">
                  <div className="sheet-label">
                    เลือก Sheet ที่เป็น KPI
                    <span className="sheet-count">{sheetNames.length} sheet</span>
                  </div>
                  <div className="sheet-list">
                    {sheetNames.map((name) => (
                      <button key={name} type="button"
                        className={`sheet-tab ${selectedSheet === name ? 'is-active' : ''}`}
                        onClick={() => handleSheetChange(name)} disabled={loading} title={name}>
                        {name}
                      </button>
                    ))}
                  </div>
                </div>
              )}

              {workbook && kpiSections.length > 0 && (
                <div className="section-filter">
                  <div className="filter-head">
                    <div className="filter-title">เลือกประเภท KPI</div>
                    <div className="filter-actions">
                      <button type="button" className="filter-link" onClick={() => setAllSections(true)}>เลือกทั้งหมด</button>
                      <button type="button" className="filter-link" onClick={() => setAllSections(false)}>ไม่เลือกทั้งหมด</button>
                    </div>
                  </div>
                  <div className="filter-sub">พบ {kpiSections.length} ประเภท (Turnover/Time/Total ถูกตัดออกอัตโนมัติ)</div>
                  <div className="filter-list" role="group">
                    {kpiSections.map((s) => (
                      <label key={s.name} className="filter-item">
                        <input type="checkbox" checked={selectedSections[s.name] ?? true}
                          onChange={(e) => setSelectedSections((prev) => ({ ...prev, [s.name]: e.target.checked }))} />
                        <span className="filter-name">{s.name}</span>
                        <span className="filter-count">{s.kpiCount}</span>
                      </label>
                    ))}
                  </div>
                </div>
              )}

              {/* ── Missing level values table ── */}
              {missingLevels.length > 0 && (
                <div className="missing-levels">
                  <div className="missing-levels-head">
                    <div>
                      <div className="missing-levels-title">
                        ⚠️ พบ {missingLevels.length} KPI ที่ยังไม่มีค่า Level
                      </div>
                      <div className="missing-levels-sub">
                        กรอกเฉพาะ Level ที่ต้องการ — รายการไหนไม่กรอกก็ปล่อยว่างได้ แล้วกด "ยืนยันและแปลงไฟล์"
                        &nbsp;·&nbsp; Level 5 = ดีที่สุด (Exceed) · Level 1 = ต่ำที่สุด (Below)
                      </div>
                    </div>
                  </div>

                  <div className="ml-table-wrap">
                    <table className="ml-table">
                      <thead>
                        <tr>
                          <th className="ml-th ml-th-name">ชื่อ KPI</th>
                          <th className="ml-th ml-th-code">MeasureCode</th>
                          <th className="ml-th ml-th-section">Section</th>
                          <th className="ml-th ml-th-lvl">★★★★★ L5</th>
                          <th className="ml-th ml-th-lvl">★★★★ L4</th>
                          <th className="ml-th ml-th-lvl">★★★ L3</th>
                          <th className="ml-th ml-th-lvl">★★ L2</th>
                          <th className="ml-th ml-th-lvl">★ L1</th>
                        </tr>
                      </thead>
                      <tbody>
                        {missingLevels.map((m) => {
                          const key = m.measureCode || m.kpiName
                          const edit = levelEdits[key] ?? { lvl1: m.lvl1, lvl2: m.lvl2, lvl3: m.lvl3, lvl4: m.lvl4, lvl5: m.lvl5 }
                          const setEdit = (field: keyof typeof edit, val: string) =>
                            setLevelEdits((prev) => ({
                              ...prev,
                              [key]: { ...(prev[key] ?? edit), [field]: val },
                            }))
                          return (
                            <tr key={key} className="ml-tr">
                              <td className="ml-td ml-td-name" title={m.kpiName}>{m.kpiName}</td>
                              <td className="ml-td ml-td-code">{m.measureCode}</td>
                              <td className="ml-td ml-td-section">{m.section}</td>
                              {(['lvl5', 'lvl4', 'lvl3', 'lvl2', 'lvl1'] as const).map((lk) => (
                                <td key={lk} className={`ml-td ml-td-input ${!edit[lk] ? 'ml-td-empty' : ''}`}>
                                  <input
                                    className="ml-input"
                                    type="text"
                                    placeholder="-"
                                    value={edit[lk]}
                                    onChange={(e) => setEdit(lk, e.target.value)}
                                  />
                                </td>
                              ))}
                            </tr>
                          )
                        })}
                      </tbody>
                    </table>
                  </div>

                  <div className="actions" style={{ marginTop: '1rem' }}>
                    <button
                      className="btn-primary"
                      onClick={handleConfirmConvert}
                      disabled={loading}
                    >
                      {loading ? '⏳ กำลังแปลง...' : '✓ ยืนยันและแปลงไฟล์'}
                    </button>
                    <button
                      className="btn-secondary"
                      onClick={() => { setMissingLevels([]); setStatus({ type: 'idle', message: '' }) }}
                      disabled={loading}
                    >
                      × ยกเลิก
                    </button>
                  </div>
                </div>
              )}

              {result && (
                <StatGrid stats={[
                  { value: result.empCount, label: 'พนักงาน' },
                  { value: result.kpiCount, label: 'KPI' },
                  { value: result.rowCount, label: 'แถวทั้งหมด' },
                ]} />
              )}

              <hr className="divider" />

              <div className="settings">
                <div className="setting-row">
                  <div><div className="setting-label">วันที่เริ่มต้น</div><div className="setting-hint">Start Date (mm/dd/yyyy)</div></div>
                  <input type="text" className="date-input" value={startDate} onChange={(e) => setStartDate(e.target.value)} placeholder="1/1/2025" />
                </div>
                <div className="setting-row">
                  <div><div className="setting-label">วันที่สิ้นสุด</div><div className="setting-hint">End Date (mm/dd/yyyy)</div></div>
                  <input type="text" className="date-input" value={endDate} onChange={(e) => setEndDate(e.target.value)} placeholder="12/31/2025" />
                </div>
              </div>

              <div className="actions">
                <button className="btn-primary" onClick={handleConvert} disabled={!canConvert}>
                  {loading ? '⏳ กำลังแปลง...' : '⚙ แปลงไฟล์'}
                </button>
                {csvContent && (
                  <button className="btn-success" onClick={handleDownload}>⬇ ดาวน์โหลด filled_sheet.csv</button>
                )}
              </div>
            </div>
            <p className="footer">Turnover & Time KPIs ถูกตัดออกอัตโนมัติ · Levels จาก column M–Q ต่อ KPI</p>
          </div>
        )}

      </main>
    </div>
  )
}
