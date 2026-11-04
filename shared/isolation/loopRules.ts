import {
  normalizeNumberString,
  parseNumber
} from './numbers'
import type {
  LoopDataRow,
  LoopProtocolResult,
  ParsedWorkbook,
  ShieldLine,
  SysLoopKRange,
  SysTripCurve
} from './types'

/** Word measurement columns 7..12 (1-based as in TZ). */
const ALL_MEAS_COLS = [7, 8, 9, 10, 11, 12] as const

function normalizePhase(raw: string): string {
  return raw.replace(/[\s,;]+/g, '').trim()
}

function activeColumnsForPhase(phaseRaw: string): number[] {
  const p = normalizePhase(phaseRaw)
  const map: Record<string, number[]> = {
    '1': [7, 10],
    '2': [8, 11],
    '3': [9, 12],
    '12': [7, 8, 10, 11],
    '21': [7, 8, 10, 11],
    '23': [8, 9, 11, 12],
    '32': [8, 9, 11, 12],
    '13': [7, 9, 10, 12],
    '31': [7, 9, 10, 12],
    '123': [7, 8, 9, 10, 11, 12],
    '132': [7, 8, 9, 10, 11, 12],
    '213': [7, 8, 9, 10, 11, 12],
    '231': [7, 8, 9, 10, 11, 12],
    '321': [7, 8, 9, 10, 11, 12],
    '312': [7, 8, 9, 10, 11, 12]
  }
  return map[p] ? [...map[p]] : [...ALL_MEAS_COLS]
}

function hasCable(line: ShieldLine): boolean {
  return Boolean(line.cableMark.trim())
}

function buildLineName(line: ShieldLine): string {
  const left = line.col2 || line.col1 || ''
  let right = line.col3
  if (!right) {
    right = hasCable(line) ? 'Потребитель' : 'Резерв'
  }
  if (left && right) return `${left} - ${right}`
  return left || right
}

function ceilIfNeeded(n: number): number {
  return Number.isInteger(n) ? n : Math.ceil(n)
}

function formatHundredths(value: number): string {
  return (Math.round(value * 100) / 100).toFixed(2)
}

function formatIntOrPlain(value: number, forceInt: boolean): string {
  if (!forceInt) {
    // keep reasonable precision for "твёрдые" значения
    if (Number.isInteger(value)) return String(value)
    return String(Math.round(value * 100) / 100)
  }
  return String(Math.round(value))
}

function toCyrillicTripLetter(letter: string): string {
  const map: Record<string, string> = { C: 'С', D: 'Д', B: 'В', c: 'С', d: 'Д', b: 'В' }
  return map[letter] || letter
}

function findTripCurve(curves: SysTripCurve[], multiplicity: string): SysTripCurve | null {
  const raw = multiplicity.trim()
  if (!raw) return null
  // letter match: any char of яч12 appears in letters bag
  for (const curve of curves) {
    for (const ch of raw) {
      if (curve.letters.includes(ch)) return curve
    }
  }
  return null
}

function findLoopK(ranges: SysLoopKRange[], nominal: number): number {
  for (const r of ranges) {
    if (nominal > r.minExclusive && nominal <= r.maxInclusive) return r.k
  }
  return ranges[ranges.length - 1]?.k ?? 1.5
}

interface BaseTripInfo {
  tripType: string
  baseMin: number
  baseMax: number
  baseRangeText: string
  /** false, если нет яч9/яч12 для расчёта — вместо нулей ставим прочерки */
  validBase: boolean
}

function tripTypeFromParts(ovv: string, md: string, suffix?: string): string {
  if (ovv && md && suffix !== undefined && suffix !== '') {
    return `${ovv}; ${md}-${suffix}`
  }
  if (ovv && md) return `${ovv}; ${md}`
  const prefix = [ovv, md].filter(Boolean).join('; ')
  if (prefix && suffix) return `${prefix}-${suffix}`
  return prefix || suffix || '-'
}

function buildBaseTripInfo(line: ShieldLine, curves: SysTripCurve[]): BaseTripInfo {
  const ovv = line.ovv.trim()
  const md = line.md.trim()
  const mult = line.multiplicity.trim()
  const nominal = parseNumber(line.nominalCurrent)
  const hasNominal = nominal !== null
  const curve = findTripCurve(curves, mult)
  const multNum = parseNumber(mult)

  if (curve) {
    const letter = toCyrillicTripLetter(curve.displayLetter)
    const tripType = tripTypeFromParts(ovv, md, letter)
    if (!hasNominal) {
      return { tripType, baseMin: 0, baseMax: 0, baseRangeText: '-', validBase: false }
    }
    const baseMin = curve.rangeMin * nominal
    const baseMax = curve.rangeMax * nominal
    return {
      tripType,
      baseMin,
      baseMax,
      baseRangeText: `${formatPlain(baseMin)} - ${formatPlain(baseMax)}`,
      validBase: true
    }
  }

  if (multNum !== null && multNum < 80) {
    const tripType = tripTypeFromParts(ovv, md, String(multNum))
    if (!hasNominal) {
      return { tripType, baseMin: 0, baseMax: 0, baseRangeText: '-', validBase: false }
    }
    const base = multNum * nominal
    const baseMin = 0.8 * base
    const baseMax = 1.2 * base
    return {
      tripType,
      baseMin,
      baseMax,
      baseRangeText: `${formatPlain(baseMin)} - ${formatPlain(baseMax)}`,
      validBase: true
    }
  }

  if (multNum !== null && multNum >= 80) {
    // база = яч12, яч9 не нужен
    const tripType = tripTypeFromParts(ovv, md)
    const base = multNum
    const baseMin = 0.8 * base
    const baseMax = 1.2 * base
    return {
      tripType,
      baseMin,
      baseMax,
      baseRangeText: `${formatPlain(baseMin)} - ${formatPlain(baseMax)}`,
      validBase: true
    }
  }

  // нет яч12 (и не буква) — диапазон не считаем
  return {
    tripType: tripTypeFromParts(ovv, md),
    baseMin: 0,
    baseMax: 0,
    baseRangeText: '-',
    validBase: false
  }
}

function formatPlain(n: number): string {
  if (Number.isInteger(n)) return String(n)
  const rounded = Math.round(n * 100) / 100
  return String(rounded)
}

type HintMode =
  | { type: 'range', min: number, max: number, forceInt: boolean }
  | { type: 'fixed', values: number[] }

function parseLoopHint(
  hintRaw: string,
  baseMax: number,
  nominalCurrent: number,
  kRanges: SysLoopKRange[]
): HintMode {
  const hint = hintRaw.trim()
  if (!hint) {
    // без валидной базы не подставляем нули
    if (!(baseMax > 0)) {
      return { type: 'fixed', values: [] }
    }
    const x = baseMax
    const k = findLoopK(kRanges, nominalCurrent)
    const y = x * k
    return {
      type: 'range',
      min: ceilIfNeeded(y * 0.8),
      max: ceilIfNeeded(x * 1.2),
      forceInt: true
    }
  }

  if (hint.includes('/')) {
    const values = hint
      .split('/')
      .map(p => parseNumber(p))
      .filter((n): n is number => n !== null)
    return { type: 'fixed', values }
  }

  const rangeMatch = normalizeNumberString(hint).match(
    /^(\d+(?:\.\d+)?)\s*(?:[-–—]|\.\.)\s*(\d+(?:\.\d+)?)$/
  )
  if (rangeMatch) {
    return {
      type: 'range',
      min: ceilIfNeeded(Number(rangeMatch[1])),
      max: ceilIfNeeded(Number(rangeMatch[2])),
      forceInt: true
    }
  }

  const n = parseNumber(hint)
  if (n !== null) {
    return {
      type: 'range',
      min: ceilIfNeeded(n * 0.85),
      max: ceilIfNeeded(n * 1.15),
      forceInt: true
    }
  }

  const x = baseMax
  const k = findLoopK(kRanges, nominalCurrent)
  const y = x * k
  return {
    type: 'range',
    min: ceilIfNeeded(y * 0.8),
    max: ceilIfNeeded(x * 1.2),
    forceInt: true
  }
}

function randomIntInRange(min: number, max: number): number {
  const a = Math.ceil(Math.min(min, max))
  const b = Math.floor(Math.max(min, max))
  if (b <= a) return a
  return a + Math.floor(Math.random() * (b - a + 1))
}

function generateMeasurements(
  line: ShieldLine,
  baseMax: number,
  kRanges: SysLoopKRange[]
): { cells: string[], currents: number[], compliance: string } {
  // indices 0..5 ↔ Word cols 7..12
  const cells = Array.from({ length: 6 }, () => '-')
  const active = activeColumnsForPhase(line.phase)
  const currentCols = active.filter(c => c >= 10 && c <= 12)
  const resistanceCols = active.filter(c => c >= 7 && c <= 9)

  const nominal = parseNumber(line.nominalCurrent) ?? 0
  const hint = parseLoopHint(line.loopHint, baseMax, nominal, kRanges)
  const currents: number[] = []
  const forceInt = hint.type === 'range' ? hint.forceInt : false

  // Fill currents (cols 10-12) first
  if (hint.type === 'fixed') {
    currentCols.forEach((col, i) => {
      const v = hint.values[i]
      if (v === undefined) return
      cells[col - 7] = formatIntOrPlain(v, false)
      currents.push(v)
    })
  } else {
    for (const col of currentCols) {
      const v = randomIntInRange(hint.min, hint.max)
      cells[col - 7] = String(v)
      currents.push(v)
    }
  }

  // Resistances R = 230 / I (matches sample; TZ text has I/230 swapped)
  for (const col of resistanceCols) {
    const currentCol = col + 3 // 7→10, 8→11, 9→12
    const raw = cells[currentCol - 7]
    const iVal = parseNumber(raw)
    if (iVal === null || iVal === 0) {
      cells[col - 7] = '-'
      continue
    }
    cells[col - 7] = formatHundredths(230 / iVal)
  }

  const threshold = baseMax * 1.1
  const compliance = currents.some(v => v < threshold) ? 'Не соотв.' : 'Соотв.'
  return { cells, currents, compliance }
}

function dashOr(value: string | undefined | null): string {
  const v = (value ?? '').trim()
  return v ? v : '-'
}

function processLine(
  line: ShieldLine,
  parsed: ParsedWorkbook,
  nextIndex: () => number
): LoopDataRow | null {
  if (line.isTitle) {
    return { kind: 'title', title: line.title || buildLineName(line) }
  }

  const base = buildBaseTripInfo(line, parsed.tripCurves)
  const breakerMark = dashOr(line.breakerMark)
  const nominalCurrent = dashOr(line.nominalCurrent)
  const tripType = dashOr(base.tripType)
  const baseRange = base.validBase ? base.baseRangeText : '-'

  // Нет кабеля (яч5) — измерения прочерки; пустые яч8/9/12 тоже не превращаем в нули
  if (!hasCable(line)) {
    return {
      kind: 'data',
      name: buildLineName(line),
      breakerMark,
      tripType,
      nominalCurrent,
      baseRange,
      measurements: Array.from({ length: 6 }, () => '-'),
      timeLimit: '-',
      timeCurve: '-',
      compliance: '-'
    }
  }

  const hasLoopHint = Boolean(line.loopHint.trim())
  const timeRaw = line.loopTime.trim()

  // Без базы (нет яч9/яч12) и без яч17 — нечего считать; без № и с прочерками
  if (!base.validBase && !hasLoopHint) {
    return {
      kind: 'data',
      name: buildLineName(line),
      breakerMark,
      tripType,
      nominalCurrent,
      baseRange: '-',
      measurements: Array.from({ length: 6 }, () => '-'),
      timeLimit: '-',
      timeCurve: '-',
      compliance: '-'
    }
  }

  const { cells, compliance } = generateMeasurements(
    line,
    base.validBase ? base.baseMax : 0,
    parsed.loopKRanges
  )

  // Если в 7–12 все прочерки — без №; в 13, 14 и 15 тоже прочерки
  const measurementsDashed = cells.every(v => !v || v === '-')
  let timeLimit = timeRaw || '0.4'
  let timeCurve = `<${timeRaw || '0.4'}`
  let finalCompliance: string = base.validBase ? compliance : '-'

  if (measurementsDashed || !base.validBase || !line.breakerMark.trim()) {
    timeLimit = '-'
    timeCurve = '-'
    finalCompliance = '-'
  }

  return {
    kind: 'data',
    index: measurementsDashed ? undefined : nextIndex(),
    name: buildLineName(line),
    breakerMark,
    tripType,
    nominalCurrent,
    baseRange,
    measurements: cells,
    timeLimit,
    timeCurve,
    compliance: finalCompliance
  }
}

export function buildLoopProtocol(parsed: ParsedWorkbook): LoopProtocolResult {
  const rows: LoopDataRow[] = []
  const nonCompliantIndexes: number[] = []
  let counter = 0
  const nextIndex = () => {
    counter += 1
    return counter
  }

  for (const shield of parsed.shields) {
    rows.push({ kind: 'title', title: shield.name })
    for (const line of shield.lines) {
      const row = processLine(line, parsed, nextIndex)
      if (!row) continue
      rows.push(row)
      if (row.kind === 'data' && row.compliance === 'Не соотв.' && row.index != null) {
        nonCompliantIndexes.push(row.index)
      }
    }
  }

  const ok
    = 'Время срабатывания аппаратов защиты соответствует требованиям ПУЭ п.1.7.79; ГОСТ Р 50571.4.41-2022 п. 411.3.2, IEC 60898-1-2020, измеренные величины ожидаемого тока КЗ соответствует требованиям НТД'
  const conclusion = nonCompliantIndexes.length
    ? `${ok}, за исключением п/п: ${nonCompliantIndexes.join(', ')}.`
    : `${ok}.`

  return { rows, nonCompliantIndexes, conclusion }
}
