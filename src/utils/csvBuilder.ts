import type { KpiRow } from '../types'

// Exact header from the original filled_sheet.csv — do not change
export const CSV_HEADER =
  'EmpCode*,PositionName,ObjectiveName,MeasureCode,Measure Name*,Measure Desc1,Measure Desc2,Measure Desc3,Measure Desc4,Measure Desc5,Measure Weight (%)*,Unit Of Measure,Measure Perspective Code,Measure Type Code*,Start Date (mm/dd/yyyy),End Date (mm/dd/yyyy),Rollup Type*,Checkpoint Frequency*,CheckpointPeriod,Baseline,Indicator Type*,Measure Value Type *,Level 1*,Level 2*,Level 3,Level 4,Level 5,Level 6,Level 7,Level 8,Level 9,Level 10,LinkSupMeasureCode,\\n'

function escapeCell(v: string | number | null | undefined): string {
  const s = String(v === null || v === undefined ? '' : v)
  if (s.includes(',') || s.includes('"') || s.includes('\n') || s.includes('\r')) {
    return '"' + s.replace(/"/g, '""') + '"'
  }
  return s
}

export function buildCsv(
  rows: KpiRow[],
  startDate: string,
  endDate: string
): string {
  const lines: string[] = [CSV_HEADER]

  for (const row of rows) {
    const cells = [
      row.empCode,
      row.position,
      '',                              // ObjectiveName
      row.measureCode,
      row.kpiName,
      '', '', '', '', '',              // Measure Desc 1-5
      row.weight,                      // decimal e.g. 0.15
      row.unit,
      '',                              // Measure Perspective Code
      row.perspective,                 // Measure Type Code*
      startDate,
      endDate,
      'Latest',                        // Rollup Type*
      row.freq,                        // Checkpoint Frequency*
      '', '',                          // CheckpointPeriod, Baseline
      'Standard Indicator 5 Level',   // Indicator Type*
      'Numeric',                       // Measure Value Type*
      row.lvl1,
      row.lvl2,
      row.lvl3,
      row.lvl4,
      row.lvl5,
      '', '', '', '', '',              // Level 6-10
      '',                              // LinkSupMeasureCode
    ]
    lines.push(cells.map(escapeCell).join(','))
  }

  // BOM for Excel Thai encoding + CRLF line endings
  return '\uFEFF' + lines.join('\r\n')
}

export function downloadCsv(content: string, filename = 'filled_sheet.csv'): void {
  const blob = new Blob([content], { type: 'text/csv;charset=utf-8' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  document.body.appendChild(a)
  a.click()
  document.body.removeChild(a)
  URL.revokeObjectURL(url)
}
