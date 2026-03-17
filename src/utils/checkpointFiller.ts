import * as XLSX from 'xlsx'
import type { CheckpointFillResult } from '../types'

const N_CHKPTS = 12
const SRC_CHKPT_COLS = Array.from({ length: N_CHKPTS }, (_, i) =>
  `Checkpoint${String(i + 1).padStart(3, '0')}`
)

// ── Frequency normalisation ────────────────────────────────────────
function normFreq(f: string): 'monthly' | 'semi-annually' | 'annually' | 'unknown' {
  const s = (f || '').trim().toLowerCase()
  if (s.includes('annual') && !s.includes('semi')) return 'annually'
  if (s.includes('semi')) return 'semi-annually'
  if (s.includes('month')) return 'monthly'
  return 'unknown'
}

/**
 * Returns 0-based target checkpoint index for mismatch cases.
 * Returns null if same frequency (copy directly).
 *
 * Rules:
 *   out=Monthly              → last value → chk12 (idx 11)
 *   out=Annually             → last value → chk1  (idx 0)
 *   out=Semi-Annually        → last value → chk2  (idx 1)
 */
function targetIndex(
  outFreq: ReturnType<typeof normFreq>,
  srcFreq: ReturnType<typeof normFreq>
): number | null {
  if (outFreq === srcFreq) return null
  if (outFreq === 'monthly') return 11
  if (outFreq === 'annually') return 0
  if (outFreq === 'semi-annually') return 1
  return null
}

// ── Parse source lookup from PMMeasureCheckpoint sheet ──────────────
interface SrcEntry {
  freq: ReturnType<typeof normFreq>
  vals: (number | null)[]
}

function buildSourceLookup(
  workbook: XLSX.WorkBook,
  sheetName: string
): Map<string, SrcEntry> {
  const sheet = workbook.Sheets[sheetName]
  if (!sheet) throw new Error(`ไม่พบ sheet "${sheetName}"`)

  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, { defval: null })
  const lookup = new Map<string, SrcEntry>()

  for (const row of rows) {
    const empRaw = row['Empcode']
    const codeRaw = row['MeasureCode-New'] ?? row['MeasureCode']
    if (empRaw === null || empRaw === undefined) continue

    const emp = String(Math.round(Number(empRaw)))
    const code = String(codeRaw ?? '').trim()
    if (!emp || !code) continue

    const freq = normFreq(String(row['CheckpointFrequency'] ?? ''))
    const vals: (number | null)[] = SRC_CHKPT_COLS.map((col) => {
      const v = row[col]
      return v !== null && v !== undefined && !Number.isNaN(Number(v)) ? Number(v) : null
    })

    lookup.set(`${emp}|${code}`, { freq, vals })
  }

  return lookup
}

// ── Parse the หยอดผลงาน CSV from an ArrayBuffer ────────────────────
function parseCsv(buffer: ArrayBuffer): { headers: string[]; rows: string[][] } {
  const text = new TextDecoder('utf-8').decode(buffer).replace(/^\uFEFF/, '')
  const lines = text.split(/\r?\n/)
  const headers = lines[0].split(',')
  const rows = lines
    .slice(1)
    .filter((l) => l.trim())
    .map((l) => l.split(','))
  return { headers, rows }
}

function buildCsvText(headers: string[], rows: string[][]): string {
  const escape = (v: string) => {
    if (v.includes(',') || v.includes('"') || v.includes('\n')) {
      return '"' + v.replace(/"/g, '""') + '"'
    }
    return v
  }
  const lines = [headers.map(escape).join(',')]
  for (const row of rows) lines.push(row.map(escape).join(','))
  return '\uFEFF' + lines.join('\r\n')
}

// ── Main fill function ─────────────────────────────────────────────
export function fillCheckpoints(
  outputFile: File,
  sourceWorkbook: XLSX.WorkBook,
  sourceSheetName: string
): Promise<CheckpointFillResult> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()

    reader.onerror = () => reject(new Error('อ่านไฟล์ CSV ไม่ได้'))

    reader.onload = (e) => {
      try {
        const buffer = e.target?.result as ArrayBuffer
        const { headers, rows } = parseCsv(buffer)

        // Find column indices
        const empIdx  = headers.findIndex((h) => h.trim() === 'Empcode')
        const codeIdx = headers.findIndex((h) => h.trim() === 'MeasureCode')
        const freqIdx = headers.findIndex((h) => h.trim() === 'CheckpointFrequency')

        // Find Checkpoint001..012 in output header
        const chkOutIndices: number[] = []
        for (let i = 1; i <= N_CHKPTS; i++) {
          const col = `Checkpoint${String(i).padStart(3, '0')}`
          const idx = headers.findIndex((h) => h.trim() === col)
          chkOutIndices.push(idx)
        }

        if (empIdx < 0 || codeIdx < 0 || freqIdx < 0) {
          throw new Error('ไฟล์ CSV ขาด column: Empcode / MeasureCode / CheckpointFrequency')
        }
        if (chkOutIndices[0] < 0) {
          throw new Error('ไม่พบ column Checkpoint001 ในไฟล์ CSV')
        }

        const lookup = buildSourceLookup(sourceWorkbook, sourceSheetName)

        let filled = 0
        let mismatchPlaced = 0
        let notFound = 0

        for (const row of rows) {
          const empRaw = row[empIdx]?.trim() ?? ''
          const code   = row[codeIdx]?.trim() ?? ''
          const outFreq = normFreq(row[freqIdx]?.trim() ?? '')

          const emp = empRaw.includes('.')
            ? String(Math.round(Number(empRaw)))
            : empRaw

          const entry = lookup.get(`${emp}|${code}`)
          if (!entry) {
            notFound++
            continue
          }

          const tidx = targetIndex(outFreq, entry.freq)

          if (tidx === null) {
            // Same frequency — copy values to matching positions
            for (let ci = 0; ci < N_CHKPTS; ci++) {
              const v = entry.vals[ci]
              const colIdx = chkOutIndices[ci]
              if (colIdx >= 0 && v !== null) {
                // Expand row if needed
                while (row.length <= colIdx) row.push('')
                row[colIdx] = String(v)
              }
            }
            filled++
          } else {
            // Frequency mismatch — find last non-null source value
            let lastVal: number | null = null
            for (const v of entry.vals) {
              if (v !== null) lastVal = v
            }
            if (lastVal !== null) {
              // Clear all 12 checkpoint columns first
              for (let ci = 0; ci < N_CHKPTS; ci++) {
                const colIdx = chkOutIndices[ci]
                if (colIdx >= 0) {
                  while (row.length <= colIdx) row.push('')
                  row[colIdx] = ''
                }
              }
              // Place at target position
              const destIdx = chkOutIndices[tidx]
              if (destIdx >= 0) {
                row[destIdx] = String(lastVal)
              }
              mismatchPlaced++
            }
          }
        }

        const csvText = buildCsvText(headers, rows)
        const blob = new Blob([csvText], { type: 'text/csv;charset=utf-8' })

        resolve({
          blob,
          fileName: outputFile.name.replace(/\.csv$/i, '_filled.csv'),
          stats: {
            filled,
            mismatchPlaced,
            notFound,
            total: rows.length,
          },
        })
      } catch (err) {
        reject(err)
      }
    }

    reader.readAsArrayBuffer(outputFile)
  })
}
