export interface KpiRow {
  empCode: string
  position: string
  measureCode: string
  kpiName: string
  section: string
  weight: number
  unit: string
  perspective: string
  freq: string
  lvl1: string
  lvl2: string
  lvl3: string
  lvl4: string
  lvl5: string
}

export interface ConvertResult {
  rows: KpiRow[]
  empCount: number
  kpiCount: number
  rowCount: number
}

export interface KpiSectionInfo {
  name: string
  kpiCount: number
}

export interface CheckpointFillStats {
  filled: number
  mismatchPlaced: number
  notFound: number
  total: number
}

export interface CheckpointFillResult {
  blob: Blob
  fileName: string
  stats: CheckpointFillStats
}

/** A unique KPI definition that is missing one or more level values */
export interface MissingLevelKpi {
  /** Row index in the sheet (0-based) — used to write overrides back */
  rowIndex: number
  measureCode: string
  kpiName: string
  section: string
  lvl1: string
  lvl2: string
  lvl3: string
  lvl4: string
  lvl5: string
}

export type StatusType = 'idle' | 'loading' | 'success' | 'error'

export interface AppStatus {
  type: StatusType
  message: string
}
