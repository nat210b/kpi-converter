import { useState, useCallback } from 'react'
import * as XLSX from 'xlsx'
import DropZone from './components/DropZone'
import StatusBanner from './components/StatusBanner'
import StatGrid from './components/StatGrid'
import { parseKpiWorkbook } from './utils/kpiParser'
import { buildCsv, downloadCsv } from './utils/csvBuilder'
import type { AppStatus, ConvertResult } from './types'

export default function App() {
  const [fileName, setFileName] = useState('')
  const [fileSize, setFileSize] = useState('')
  const [status, setStatus] = useState<AppStatus>({ type: 'idle', message: '' })
  const [result, setResult] = useState<ConvertResult | null>(null)
  const [csvContent, setCsvContent] = useState<string | null>(null)
  const [startDate, setStartDate] = useState('1/1/2025')
  const [endDate, setEndDate] = useState('12/31/2025')
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null)
  const [loading, setLoading] = useState(false)

  const handleFile = useCallback((file: File) => {
    if (!file.name.match(/\.xlsx?$/i)) {
      setStatus({ type: 'error', message: 'กรุณาเลือกไฟล์ .xlsx หรือ .xls เท่านั้น' })
      return
    }
    setFileName(file.name)
    setFileSize((file.size / 1024).toFixed(1) + ' KB')
    setStatus({ type: 'loading', message: 'กำลังโหลดไฟล์...' })
    setResult(null)
    setCsvContent(null)
    setWorkbook(null)

    const reader = new FileReader()
    reader.onload = (e) => {
      try {
        const data = e.target?.result
        const wb = XLSX.read(data, { type: 'array' })
        const hasKpiSheet = wb.SheetNames.some((s) => s.includes('KPI'))
        if (!hasKpiSheet) {
          setStatus({ type: 'error', message: 'ไม่พบ sheet ที่มีชื่อ "KPI" ในไฟล์นี้' })
          return
        }
        setWorkbook(wb)
        setStatus({
          type: 'success',
          message: `โหลดสำเร็จ — พบ ${wb.SheetNames.length} sheet กด "แปลงไฟล์" เพื่อดำเนินการต่อ`,
        })
      } catch (err) {
        setStatus({ type: 'error', message: 'อ่านไฟล์ไม่ได้: ' + (err as Error).message })
      }
    }
    reader.readAsArrayBuffer(file)
  }, [])

  const handleConvert = useCallback(() => {
    if (!workbook) return
    setLoading(true)
    setStatus({ type: 'loading', message: 'กำลังแปลงข้อมูล...' })

    setTimeout(() => {
      try {
        const parsed = parseKpiWorkbook(workbook)
        if (parsed.rowCount === 0) {
          setStatus({ type: 'error', message: 'ไม่พบข้อมูล KPI ในไฟล์ กรุณาตรวจสอบ sheet' })
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
        setStatus({ type: 'error', message: 'เกิดข้อผิดพลาด: ' + (err as Error).message })
      }
      setLoading(false)
    }, 50)
  }, [workbook, startDate, endDate])

  const handleDownload = useCallback(() => {
    if (!csvContent) return
    downloadCsv(csvContent, 'filled_sheet.csv')
  }, [csvContent])

  return (
    <div className="page">
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
              disabled={!workbook || loading}
            >
              {loading ? '⏳ กำลังแปลง...' : '⚙ แปลงไฟล์'}
            </button>

            {csvContent && (
              <button className="btn-success" onClick={handleDownload}>
                ⬇ ดาวน์โหลด filled_sheet.csv
              </button>
            )}
          </div>
        </div>

        <p className="footer">
          Turnover &amp; Time KPIs ถูกตัดออกอัตโนมัติ · Levels จาก column M–Q ต่อ KPI
        </p>
      </div>
    </div>
  )
}