export interface SysIsolationRange {
  /** Inclusive/exclusive bounds for max cable section, mm² */
  minExclusive: number
  maxInclusive: number
  voltage: number
  /** Random base range from sys sheet, MΩ */
  momMin: number
  momMax: number
}

/** sys!H19:J21 — характеристики АВ (буквы кратности) */
export interface SysTripCurve {
  /** Characters that match яч12, e.g. "СсCc" */
  letters: string
  /** Display letter, e.g. "C" / "С" */
  displayLetter: string
  rangeMin: number
  rangeMax: number
}

/** sys!H24:I28 — коэффициент K по номинальному току */
export interface SysLoopKRange {
  minExclusive: number
  maxInclusive: number
  k: number
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
  /** яч8 — марка АВ */
  breakerMark: string
  /** яч9 — ном. ток */
  nominalCurrent: string
  /** яч10 — ОВВ */
  ovv: string
  /** яч11 — МД */
  md: string
  /** яч12 — кратность (В/С/8/10..) */
  multiplicity: string
  /** яч14 — «Изм. изол.» */
  isolationHint: string
  /** яч17 — «Изм. Петля» */
  loopHint: string
  /** яч18 — «Время петли» */
  loopTime: string
}

export interface ShieldSheet {
  name: string
  lines: ShieldLine[]
}

export interface ParsedWorkbook {
  shields: ShieldSheet[]
  sysRanges: SysIsolationRange[]
  tripCurves: SysTripCurve[]
  loopKRanges: SysLoopKRange[]
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
  /** Word cols 6–15 for isolation */
  resistances?: ResistanceCell[]
  compliance?: 'Соотв.' | 'Не соотв.' | '' | '-'
}

export interface ProtocolResult {
  rows: ProtocolDataRow[]
  nonCompliantIndexes: number[]
  conclusion: string
}

/** Row for phase-zero loop protocol (петля.docx), 15 columns */
export interface LoopDataRow {
  kind: 'title' | 'data'
  title?: string
  index?: number
  name?: string
  /** яч3 — марка АВ */
  breakerMark?: string
  /** яч4 — тип расцепителя */
  tripType?: string
  /** яч5 — ном. ток */
  nominalCurrent?: string
  /** яч6 — базовый диапазон */
  baseRange?: string
  /** Word cols 7–12: R_A,R_B,R_C,I_A,I_B,I_C */
  measurements?: string[]
  timeLimit?: string
  timeCurve?: string
  compliance?: string
}

export interface LoopProtocolResult {
  rows: LoopDataRow[]
  nonCompliantIndexes: number[]
  conclusion: string
}
