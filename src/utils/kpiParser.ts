import * as XLSX from 'xlsx'
import type { ConvertResult, KpiRow, KpiSectionInfo } from '../types'

const EXCLUDED_SECTIONS = ['Turnover', 'Time', 'Total']
const FALLBACK_SECTION = 'Other'

function cellVal(sheet: XLSX.WorkSheet, r: number, c: number): string {
  const addr = XLSX.utils.encode_cell({ r, c })
  const cell = sheet[addr]
  if (!cell || cell.v === undefined || cell.v === null) return ''
  return String(cell.v).trim()
}

function normalizeSectionName(raw: string): string {
  const t = (raw || '').trim()
  return t ? t : FALLBACK_SECTION
}

function isExcludedSection(sectionName: string): boolean {
  return EXCLUDED_SECTIONS.some((x) => sectionName.includes(x))
}

/**
 * If the cell value is a plain number (e.g. 0.9, 1.2), convert to percentage string (90%, 120%).
 * Values that already contain %, <, >, =, Thai text, or other symbols are returned as-is.
 */
function formatLevel(raw: string): string {
  if (!raw) return ''
  const trimmed = raw.trim()
  // Already has a % sign or non-numeric symbols â€” leave it alone
  if (/[%<>=a-zA-Zà¸-à¹™]/.test(trimmed)) return trimmed
  // Try to parse as a pure number
  const num = Number(trimmed)
  if (!isNaN(num) && trimmed !== '') {
    // Multiply by 100, remove trailing .00 noise
    const pct = Math.round(num * 10000) / 100 // e.g. 0.9 â†’ 90, 1.2 â†’ 120
    const formatted = pct % 1 === 0 ? `${pct}%` : `${pct}%`
    return formatted
  }
  return trimmed
}

function mapFreq(f: string): string {
  if (!f) return 'Semi-Annually'
  if (f.includes('à¸—à¸¸à¸à¹€à¸”à¸·à¸­à¸™') || f.includes('Monthly')) return 'Monthly'
  if (f.includes('à¸›à¸¥à¸²à¸¢à¸›à¸µ') && !f.includes('à¸„à¸£à¸¶à¹ˆà¸‡')) return 'Annually'
  return 'Semi-Annually'
}

function findKpiSheet(workbook: XLSX.WorkBook): XLSX.WorkSheet {
  // Accept any sheet name containing "KPI"
  const sheetName = workbook.SheetNames.find((s) => s.includes('KPI'))
  if (!sheetName) {
    throw new Error('à¹„à¸¡à¹ˆà¸žà¸š sheet à¸—à¸µà¹ˆà¸¡à¸µà¸Šà¸·à¹ˆà¸­ "KPI" à¹ƒà¸™à¹„à¸Ÿà¸¥à¹Œà¸™à¸µà¹‰')
  }
  return workbook.Sheets[sheetName]
}

export function scanKpiSections(workbook: XLSX.WorkBook): KpiSectionInfo[] {
  const sheet = findKpiSheet(workbook)
  const range = XLSX.utils.decode_range(sheet['!ref'] ?? 'A1:A1')
  const maxRow = range.e.r

  let currentSection = ''
  const counts = new Map<string, number>()

  for (let r = 12; r <= Math.min(maxRow, 60); r++) {
    const sec = cellVal(sheet, r, 0)
    if (sec) currentSection = sec

    const sectionName = normalizeSectionName(currentSection)
    if (isExcludedSection(sectionName)) continue

    const kpiName = cellVal(sheet, r, 4)
    if (!kpiName || kpiName === 'nan') continue

    counts.set(sectionName, (counts.get(sectionName) ?? 0) + 1)
  }

  return Array.from(counts.entries())
    .map(([name, kpiCount]) => ({ name, kpiCount }))
    .sort((a, b) => b.kpiCount - a.kpiCount || a.name.localeCompare(b.name))
}

export function parseKpiWorkbook(
  workbook: XLSX.WorkBook,
  options?: { includeSections?: Set<string> }
): ConvertResult {
  const sheet = findKpiSheet(workbook)

  const range = XLSX.utils.decode_range(sheet['!ref'] ?? 'A1:A1')
  const maxRow = range.e.r
  const maxCol = range.e.c

  // Row 5 (0-indexed) = emp codes, Row 8 = positions
  const empCodes: Record<number, string> = {}
  const empPositions: Record<number, string> = {}

  for (let c = 17; c <= maxCol; c++) {
    const code = cellVal(sheet, 5, c)
    if (code && code !== 'nan') {
      empCodes[c] = code.split('.')[0]
      empPositions[c] = cellVal(sheet, 8, c)
    }
  }

  const rows: KpiRow[] = []
  let currentSection = ''
  const kpiNames = new Set<string>()

  for (let r = 12; r <= Math.min(maxRow, 60); r++) {
    const sec = cellVal(sheet, r, 0)
    if (sec) currentSection = sec

    const sectionName = normalizeSectionName(currentSection)

    if (isExcludedSection(sectionName)) continue
    if (options?.includeSections && !options.includeSections.has(sectionName)) continue

    const kpiName = cellVal(sheet, r, 4)
    if (!kpiName || kpiName === 'nan') continue

    const measureCode = cellVal(sheet, r, 5)
    const unit = cellVal(sheet, r, 7) || 'à¸£à¹‰à¸­à¸¢à¸¥à¸°'
    const freq = mapFreq(cellVal(sheet, r, 10))

    const perspective = sectionName.includes('Department')
      ? 'Department KPIs'
      : 'Individual KPIs'

    // M=col12=Level5, N=col13=Level4, O=col14=Level3, P=col15=Level2, Q=col16=Level1
    const getRawLevel = (col: number): string => {
      const addr = XLSX.utils.encode_cell({ r, c: col })
      const cell = sheet[addr]
      if (!cell || cell.v === undefined || cell.v === null) return ''
      if (cell.t === 'n') return formatLevel(String(cell.v))
      return formatLevel(String(cell.v).trim())
    }

    const lvl5 = getRawLevel(12)
    const lvl4 = getRawLevel(13)
    const lvl3 = getRawLevel(14)
    const lvl2 = getRawLevel(15)
    const lvl1 = getRawLevel(16)

    let addedAny = false

    for (const [colStr, empCode] of Object.entries(empCodes)) {
      const col = parseInt(colStr)
      const wCell = sheet[XLSX.utils.encode_cell({ r, c: col })]
      if (!wCell || wCell.v === undefined || wCell.v === null) continue
      const w = parseFloat(String(wCell.v))
      if (isNaN(w) || w <= 0) continue

      rows.push({
        empCode,
        position: empPositions[col] ?? '',
        measureCode,
        kpiName,
        section: sectionName,
        weight: w,
        unit,
        perspective,
        freq,
        lvl1,
        lvl2,
        lvl3,
        lvl4,
        lvl5,
      })
      addedAny = true
    }

    if (addedAny) kpiNames.add(kpiName)
  }

  rows.sort((a, b) => a.empCode.localeCompare(b.empCode))

  const empSet = new Set(rows.map((r) => r.empCode))

  return {
    rows,
    empCount: empSet.size,
    kpiCount: kpiNames.size,
    rowCount: rows.length,
  }
}

