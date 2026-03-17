import { useCallback, useRef, useState } from 'react'
import * as XLSX from 'xlsx'
import { fillCheckpoints } from '../utils/checkpointFiller'
import type { AppStatus, CheckpointFillStats } from '../types'

interface SrcState {
  workbook: XLSX.WorkBook
  fileName: string
  sheetNames: string[]
  selectedSheet: string
}

export default function CheckpointPage() {
  const [src, setSrc]       = useState<SrcState | null>(null)
  const [csvFiles, setCsvFiles] = useState<File[]>([])
  const [status, setStatus] = useState<AppStatus>({ type: 'idle', message: '' })
  const [loading, setLoading] = useState(false)
  const [results, setResults] = useState<
    { name: string; stats: CheckpointFillStats; blob: Blob; done: boolean }[]
  >([])

  const srcInputRef = useRef<HTMLInputElement>(null)
  const csvInputRef = useRef<HTMLInputElement>(null)

  // ── Load source Excel ──────────────────────────────────────────
  function handleSrcFile(file: File) {
    if (!file.name.match(/\.xlsx?$/i)) {
      setStatus({ type: 'error', message: 'กรุณาเลือกไฟล์ .xlsx หรือ .xls เท่านั้น' })
      return
    }
    setStatus({ type: 'loading', message: 'กำลังโหลด Excel...' })
    setSrc(null)
    setResults([])

    const reader = new FileReader()
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target?.result, { type: 'array' })
        const detected =
          wb.SheetNames.find((s) => s.toLowerCase().includes('pmmeasurecheckpoint')) ??
          wb.SheetNames.find((s) => s.toLowerCase().includes('checkpoint')) ??
          wb.SheetNames[0] ??
          ''
        setSrc({
          workbook: wb,
          fileName: file.name,
          sheetNames: wb.SheetNames,
          selectedSheet: detected,
        })
        setStatus({
          type: 'success',
          message: `โหลด Excel สำเร็จ — พบ ${wb.SheetNames.length} sheet · ใช้ "${detected}"`,
        })
      } catch (err) {
        setStatus({ type: 'error', message: 'อ่านไฟล์ไม่ได้: ' + (err as Error).message })
      }
    }
    reader.readAsArrayBuffer(file)
  }

  // ── Load output CSV files ─────────────────────────────────────
  function handleCsvFiles(files: FileList | null) {
    if (!files) return
    const valid = Array.from(files).filter((f) => f.name.match(/\.csv$/i))
    if (!valid.length) {
      setStatus({ type: 'error', message: 'กรุณาเลือกไฟล์ .csv' })
      return
    }
    setCsvFiles(valid)
    setResults([])
    setStatus({ type: 'success', message: `เลือก ${valid.length} ไฟล์ CSV แล้ว` })
  }

  // ── Run fill ──────────────────────────────────────────────────
  const handleFill = useCallback(async () => {
    if (!src || !csvFiles.length) return
    setLoading(true)
    setResults([])
    setStatus({ type: 'loading', message: 'กำลังหยอดผลงาน...' })

    const out: typeof results = []

    for (const file of csvFiles) {
      try {
        const res = await fillCheckpoints(file, src.workbook, src.selectedSheet)
        out.push({ name: res.fileName, stats: res.stats, blob: res.blob, done: true })
      } catch (err) {
        out.push({
          name: file.name,
          stats: { filled: 0, mismatchPlaced: 0, notFound: 0, total: 0 },
          blob: new Blob(),
          done: false,
        })
        setStatus({ type: 'error', message: `${file.name}: ${(err as Error).message}` })
      }
    }

    setResults(out)
    setLoading(false)

    const allOk = out.every((r) => r.done)
    if (allOk) {
      const totalFilled = out.reduce((s, r) => s + r.stats.filled + r.stats.mismatchPlaced, 0)
      setStatus({ type: 'success', message: `เสร็จแล้ว — รวมหยอด ${totalFilled} แถว จาก ${out.length} ไฟล์` })
    }
  }, [src, csvFiles])

  function downloadResult(r: { name: string; blob: Blob }) {
    const url = URL.createObjectURL(r.blob)
    const a = document.createElement('a')
    a.href = url
    a.download = r.name
    a.click()
    URL.revokeObjectURL(url)
  }

  const canFill = Boolean(src) && csvFiles.length > 0 && !loading

  return (
    <div className="container">
      <div className="header">
        <h1 className="title">หยอดผลงาน Checkpoint</h1>
        <p className="subtitle">
          อัพโหลด Excel (PMMeasureCheckpoint) + ไฟล์ CSV หยอดผลงาน แล้วดาวน์โหลด CSV ที่กรอก Checkpoint เรียบร้อย
        </p>
      </div>

      <div className="card">

        {/* ── Step 1: Source Excel ── */}
        <div className="cp-step">
          <div className="cp-step-label">
            <span className="cp-badge">1</span>
            เลือกไฟล์ Excel ที่มี sheet PMMeasureCheckpoint
          </div>

          <div
            className={`drop-zone ${src ? 'cp-loaded' : ''}`}
            onClick={() => srcInputRef.current?.click()}
          >
            <input
              ref={srcInputRef}
              type="file"
              accept=".xlsx,.xls"
              style={{ display: 'none' }}
              onChange={(e) => e.target.files?.[0] && handleSrcFile(e.target.files[0])}
            />
            {src ? (
              <div className="file-info">
                <span className="file-icon">📊</span>
                <div>
                  <div className="file-name">{src.fileName}</div>
                  <div className="file-meta">{src.sheetNames.length} sheets</div>
                </div>
                <span className="change-hint">คลิกเพื่อเปลี่ยน</span>
              </div>
            ) : (
              <>
                <div className="drop-icon">📊</div>
                <div className="drop-label">คลิกเพื่อเลือก Excel</div>
                <div className="drop-sub">.xlsx / .xls ที่มี sheet PMMeasureCheckpoint</div>
              </>
            )}
          </div>

          {/* Sheet selector */}
          {src && src.sheetNames.length > 0 && (
            <div className="sheet-selector" style={{ marginTop: '0.75rem' }}>
              <div className="sheet-label">
                เลือก Sheet ที่เป็น PMMeasureCheckpoint
                <span className="sheet-count">{src.sheetNames.length} sheet</span>
              </div>
              <div className="sheet-list">
                {src.sheetNames.map((name) => (
                  <button
                    key={name}
                    type="button"
                    className={`sheet-tab ${src.selectedSheet === name ? 'is-active' : ''}`}
                    onClick={() => setSrc((prev) => prev ? { ...prev, selectedSheet: name } : prev)}
                    disabled={loading}
                    title={name}
                  >
                    {name}
                  </button>
                ))}
              </div>
            </div>
          )}
        </div>

        <hr className="divider" />

        {/* ── Step 2: CSV files ── */}
        <div className="cp-step">
          <div className="cp-step-label">
            <span className="cp-badge">2</span>
            เลือกไฟล์ CSV หยอดผลงาน (เลือกหลายไฟล์ได้)
          </div>

          <div
            className={`drop-zone ${csvFiles.length ? 'cp-loaded' : ''}`}
            onClick={() => csvInputRef.current?.click()}
          >
            <input
              ref={csvInputRef}
              type="file"
              accept=".csv"
              multiple
              style={{ display: 'none' }}
              onChange={(e) => handleCsvFiles(e.target.files)}
            />
            {csvFiles.length > 0 ? (
              <div className="file-info">
                <span className="file-icon">📄</span>
                <div>
                  <div className="file-name">
                    {csvFiles.length === 1
                      ? csvFiles[0].name
                      : `${csvFiles.length} ไฟล์: ${csvFiles.map((f) => f.name).join(', ')}`}
                  </div>
                  <div className="file-meta">
                    รวม {(csvFiles.reduce((s, f) => s + f.size, 0) / 1024).toFixed(1)} KB
                  </div>
                </div>
                <span className="change-hint">คลิกเพื่อเปลี่ยน</span>
              </div>
            ) : (
              <>
                <div className="drop-icon">📄</div>
                <div className="drop-label">คลิกเพื่อเลือก CSV</div>
                <div className="drop-sub">ไฟล์ หยอดผลงาน .csv (เลือกหลายไฟล์ได้)</div>
              </>
            )}
          </div>
        </div>

        {/* ── Status ── */}
        {status.type !== 'idle' && (
          <div className={`status-banner status-${status.type}`} style={{ marginTop: '1rem' }}>
            <span>
              {status.type === 'loading' ? '⏳' : status.type === 'success' ? '✅' : '❌'}
            </span>
            <span>{status.message}</span>
          </div>
        )}

        <hr className="divider" />

        {/* ── Action ── */}
        <div className="actions">
          <button className="btn-primary" onClick={handleFill} disabled={!canFill}>
            {loading ? '⏳ กำลังหยอด...' : '⚙ หยอดผลงาน'}
          </button>
        </div>

        {/* ── Results ── */}
        {results.length > 0 && (
          <div className="cp-results">
            {results.map((r, i) => (
              <div key={i} className={`cp-result-row ${r.done ? '' : 'cp-result-error'}`}>
                <div className="cp-result-info">
                  <div className="cp-result-name">{r.name}</div>
                  {r.done && (
                    <div className="cp-result-stats">
                      ✅ {r.stats.filled} แถว (ตรง) · {r.stats.mismatchPlaced} แถว (แก้ freq) · {r.stats.notFound} ไม่พบ
                    </div>
                  )}
                  {!r.done && <div className="cp-result-stats">❌ เกิดข้อผิดพลาด</div>}
                </div>
                {r.done && r.blob.size > 0 && (
                  <button className="btn-success" onClick={() => downloadResult(r)}>
                    ⬇ ดาวน์โหลด
                  </button>
                )}
              </div>
            ))}
          </div>
        )}

      </div>

      <p className="footer">
        Checkpoint001–012 (col N–Y) · ความถี่ไม่ตรง → ใส่ค่าสุดท้ายใน chk ที่ถูกต้อง
      </p>
    </div>
  )
}
