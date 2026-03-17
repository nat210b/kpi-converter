import * as XLSX from 'xlsx'
import type { ConvertResult, KpiRow, KpiSectionInfo, MissingLevelKpi } from '../types'

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

function formatLevel(raw: string): string {
  if (!raw) return ''
  const trimmed = raw.trim()
  if (/[%<>=a-zA-Zก-๙]/.test(trimmed)) return trimmed
  const num = Number(trimmed)
  if (!isNaN(num) && trimmed !== '') {
    const pct = Math.round(num * 10000) / 100
    return `${pct}%`
  }
  return trimmed
}

function mapFreq(f: string): string {
  if (!f) return 'Semi-Annually'
  if (f.includes('ทุกเดือน') || f.includes('Monthly')) return 'Monthly'
  if (f.includes('ปลายปี') && !f.includes('ครึ่ง')) return 'Annually'
  return 'Semi-Annually'
}

/** Returns the sheet by exact name. Throws a user-friendly error if not found. */
function getSheetByName(workbook: XLSX.WorkBook, sheetName: string): XLSX.WorkSheet {
  const sheet = workbook.Sheets[sheetName]
  if (!sheet) throw new Error(`ไม่พบ sheet "${sheetName}" ในไฟล์นี้`)
  return sheet
}

/**
 * Auto-detect the best KPI sheet from the workbook.
 * Returns the name of the first sheet whose name contains "KPI",
 * falling back to the first sheet in the workbook.
 */
export function autoDetectKpiSheet(workbook: XLSX.WorkBook): string {
  const match = workbook.SheetNames.find((s) => s.includes('KPI'))
  return match ?? workbook.SheetNames[0] ?? ''
}

export function scanKpiSections(
  workbook: XLSX.WorkBook,
  sheetName: string
): KpiSectionInfo[] {
  const sheet = getSheetByName(workbook, sheetName)
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

/**
 * Scan the KPI sheet and return all unique KPI rows that are missing
 * at least one of the five level values (lvl1–lvl5).
 * Excluded sections (Turnover / Time / Total) are skipped.
 */
export function scanMissingLevels(
  workbook: XLSX.WorkBook,
  sheetName: string,
  options?: { includeSections?: Set<string> }
): MissingLevelKpi[] {
  const sheet = getSheetByName(workbook, sheetName)
  const range = XLSX.utils.decode_range(sheet['!ref'] ?? 'A1:A1')
  const maxRow = range.e.r

  let currentSection = ''
  const seen = new Set<string>()
  const result: MissingLevelKpi[] = []

  for (let r = 12; r <= Math.min(maxRow, 60); r++) {
    const sec = cellVal(sheet, r, 0)
    if (sec) currentSection = sec

    const sectionName = normalizeSectionName(currentSection)
    if (isExcludedSection(sectionName)) continue
    if (options?.includeSections && !options.includeSections.has(sectionName)) continue

    const kpiName = cellVal(sheet, r, 4)
    if (!kpiName || kpiName === 'nan') continue

    const measureCode = cellVal(sheet, r, 5)
    const key = measureCode || kpiName
    if (seen.has(key)) continue
    seen.add(key)

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

    // Only include if at least one level is missing
    if (!lvl1 || !lvl2 || !lvl3 || !lvl4 || !lvl5) {
      result.push({ rowIndex: r, measureCode, kpiName, section: sectionName, lvl1, lvl2, lvl3, lvl4, lvl5 })
    }
  }

  return result
}

export function parseKpiWorkbook(
  workbook: XLSX.WorkBook,
  sheetName: string,
  options?: { includeSections?: Set<string>; levelOverrides?: Map<string, { lvl1: string; lvl2: string; lvl3: string; lvl4: string; lvl5: string }> }
): ConvertResult {
  const sheet = getSheetByName(workbook, sheetName)

  const range = XLSX.utils.decode_range(sheet['!ref'] ?? 'A1:A1')
  const maxRow = range.e.r
  const maxCol = range.e.c

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
    const unit = cellVal(sheet, r, 7) || 'ร้อยละ'
    const freq = mapFreq(cellVal(sheet, r, 10))

    const perspective = sectionName.includes('Department')
      ? 'Department KPIs'
      : 'Individual KPIs'

    const getRawLevel = (col: number): string => {
      const addr = XLSX.utils.encode_cell({ r, c: col })
      const cell = sheet[addr]
      if (!cell || cell.v === undefined || cell.v === null) return ''
      if (cell.t === 'n') return formatLevel(String(cell.v))
      return formatLevel(String(cell.v).trim())
    }

    let lvl5 = getRawLevel(12)
    let lvl4 = getRawLevel(13)
    let lvl3 = getRawLevel(14)
    let lvl2 = getRawLevel(15)
    let lvl1 = getRawLevel(16)

    // Apply user-supplied level overrides (keyed by measureCode || kpiName)
    // Empty string is intentional — user left it blank on purpose, so we respect that too
    const overrideKey = measureCode || kpiName
    const override = options?.levelOverrides?.get(overrideKey)
    if (override) {
      lvl1 = override.lvl1
      lvl2 = override.lvl2
      lvl3 = override.lvl3
      lvl4 = override.lvl4
      lvl5 = override.lvl5
    }

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
