export interface SysIsolationRange {
  /** Inclusive/exclusive bounds for max cable section, mm² */
  minExclusive: number
  maxInclusive: number
  voltage: number
  /** Random base range from sys sheet, MΩ */
  momMin: number
  momMax: number
}

export interface ShieldLine {
  /** Excel row number (1-based), for debugging */
  excelRow: number
  /** True for section/shield title rows inside a sheet */
  isTitle: boolean
  title?: string
  col1: string
  col2: string
  col3: string
  phase: string
  cableMark: string
  cores: string
  section: string
  /** яч14 — «Изм. изол.» */
  isolationHint: string
}

export interface ShieldSheet {
  name: string
  lines: ShieldLine[]
}

export interface ParsedWorkbook {
  shields: ShieldSheet[]
  sysRanges: SysIsolationRange[]
}

export type ResistanceCell = string

export interface ProtocolDataRow {
  kind: 'title' | 'data'
  /** Shield / panel title */
  title?: string
  /** сквозной № */
  index?: number
  name?: string
  cable?: string
  voltage?: string
  allowed?: string
  /** Word cols 6–15 */
  resistances?: ResistanceCell[]
  compliance?: 'Соотв.' | 'Не соотв.' | ''
}

export interface ProtocolResult {
  rows: ProtocolDataRow[]
  nonCompliantIndexes: number[]
  conclusion: string
}
