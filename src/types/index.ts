export interface KpiRow {
  empCode: string
  position: string
  measureCode: string
  kpiName: string
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

export type StatusType = 'idle' | 'loading' | 'success' | 'error'

export interface AppStatus {
  type: StatusType
  message: string
}
