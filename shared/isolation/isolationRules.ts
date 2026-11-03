import {
  formatResistance,
  maxSection,
  normalizeNumberString,
  parseNumber,
  randomInRange
} from './numbers'
import type {
  ParsedWorkbook,
  ProtocolDataRow,
  ProtocolResult,
  ResistanceCell,
  ShieldLine,
  SysIsolationRange
} from './types'

/** Word columns 6..15 (1-based indexes from the rules). */
const ALL_RES_COLS = [6, 7, 8, 9, 10, 11, 12, 13, 14, 15] as const

function normalizePhase(raw: string): string {
  return raw.replace(/[\s,;]+/g, '').trim()
}

/** Active measurement columns for a given phase code. */
export function activeColumnsForPhase(phaseRaw: string): number[] {
  const p = normalizePhase(phaseRaw)
  const map: Record<string, number[]> = {
    '1': [9, 12, 15],
    '2': [10, 13, 15],
    '3': [11, 14, 15],
    '12': [6, 9, 10, 12, 13, 15],
    '21': [6, 9, 10, 12, 13, 15],
    '23': [7, 10, 11, 13, 14, 15],
    '32': [7, 10, 11, 13, 14, 15],
    '13': [8, 9, 11, 12, 14, 15],
    '31': [8, 9, 11, 12, 14, 15],
    '123': [6, 7, 8, 9, 10, 11, 12, 13, 14, 15],
    '132': [6, 7, 8, 9, 10, 11, 12, 13, 14, 15],
    '213': [6, 7, 8, 9, 10, 11, 12, 13, 14, 15],
    '231': [6, 7, 8, 9, 10, 11, 12, 13, 14, 15],
    '321': [6, 7, 8, 9, 10, 11, 12, 13, 14, 15],
    '312': [6, 7, 8, 9, 10, 11, 12, 13, 14, 15]
  }
  return map[p] ? [...map[p]] : [...ALL_RES_COLS]
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

function buildCableLabel(line: ShieldLine): string {
  if (!hasCable(line)) return ''
  const mark = line.cableMark.trim()
  const cores = line.cores.trim()
  const section = normalizeNumberString(line.section).replace(/\./g, '.') // keep dots
  // Prefer original section text but with comma→dot for decimals in numbers only
  const sectionOut = String(line.section).trim().replace(/(\d),(\d)/g, '$1.$2')
  if (cores && sectionOut) return `${mark} ${cores}x${sectionOut}`
  if (cores) return `${mark} ${cores}x${section}`
  return mark
}

function findSysRange(sysRanges: SysIsolationRange[], sectionRaw: string): SysIsolationRange {
  const max = maxSection(sectionRaw)
  if (max === null) return sysRanges[sysRanges.length - 1]!

  for (const r of sysRanges) {
    if (max > r.minExclusive && max <= r.maxInclusive) return r
  }
  // fallback: last open-ended or closest
  return sysRanges[sysRanges.length - 1]!
}

type HintMode =
  | { type: 'range', min: number, max: number }
  | { type: 'fixed', values: number[] }

function parseIsolationHint(hintRaw: string, sys: SysIsolationRange): HintMode {
  const hint = hintRaw.trim()
  if (!hint) {
    const x = randomInRange(sys.momMin, sys.momMax)
    return { type: 'range', min: x * 0.75, max: x * 1.25 }
  }

  // slash-separated fixed values
  if (hint.includes('/')) {
    const values = hint
      .split('/')
      .map(p => parseNumber(p))
      .filter((n): n is number => n !== null)
    return { type: 'fixed', values }
  }

  // explicit range: a-b or a..b
  const rangeMatch = normalizeNumberString(hint).match(
    /^(\d+(?:\.\d+)?)\s*(?:[-–—]|\.\.)\s*(\d+(?:\.\d+)?)$/
  )
  if (rangeMatch) {
    return { type: 'range', min: Number(rangeMatch[1]), max: Number(rangeMatch[2]) }
  }

  // single number
  const n = parseNumber(hint)
  if (n !== null) {
    let min = n * 0.75
    let max = n * 1.25
    if (n < 0.5) {
      max = Math.min(max, 0.49)
    } else {
      // n >= 0.5
      min = Math.max(min, 0.5)
    }
    return { type: 'range', min, max }
  }

  const x = randomInRange(sys.momMin, sys.momMax)
  return { type: 'range', min: x * 0.75, max: x * 1.25 }
}

function shouldSkipNPe(line: ShieldLine): boolean {
  const cores = parseNumber(line.cores)
  const sectionHasPlus = String(line.section).includes('+')
  return (cores === 2 || cores === 4) && !sectionHasPlus
}

function generateResistances(line: ShieldLine, sys: SysIsolationRange): {
  cells: ResistanceCell[]
  compliance: 'Соотв.' | 'Не соотв.'
} {
  // cells indexed as Word cols 6..15 → array length 10
  const cells: ResistanceCell[] = Array.from({ length: 10 }, () => '-')
  let active = activeColumnsForPhase(line.phase)

  if (shouldSkipNPe(line)) {
    active = active.filter(c => c !== 15)
  }

  const hint = parseIsolationHint(line.isolationHint, sys)
  const generated: number[] = []

  if (hint.type === 'fixed') {
    active.forEach((col, i) => {
      const v = hint.values[i]
      if (v === undefined) {
        cells[col - 6] = '-'
        return
      }
      cells[col - 6] = formatResistance(v)
      generated.push(v)
    })
  } else {
    for (const col of active) {
      const v = randomInRange(hint.min, hint.max)
      cells[col - 6] = formatResistance(v)
      generated.push(v)
    }
  }

  const compliance = generated.some(v => v < 0.5) ? 'Не соотв.' : 'Соотв.'
  return { cells, compliance }
}

function processLine(
  line: ShieldLine,
  sysRanges: SysIsolationRange[],
  nextIndex: () => number
): ProtocolDataRow | null {
  if (line.isTitle) {
    return { kind: 'title', title: line.title || buildLineName(line) }
  }

  if (!hasCable(line)) {
    // Lines without cable are omitted from the protocol table
    return null
  }

  const sys = findSysRange(sysRanges, line.section)
  const { cells, compliance } = generateResistances(line, sys)

  return {
    kind: 'data',
    index: nextIndex(),
    name: buildLineName(line),
    cable: buildCableLabel(line),
    voltage: String(sys.voltage),
    allowed: '0.5',
    resistances: cells,
    compliance
  }
}

export function buildProtocol(parsed: ParsedWorkbook): ProtocolResult {
  const rows: ProtocolDataRow[] = []
  const nonCompliantIndexes: number[] = []
  let counter = 0
  const nextIndex = () => {
    counter += 1
    return counter
  }

  for (const shield of parsed.shields) {
    rows.push({ kind: 'title', title: shield.name })

    for (const line of shield.lines) {
      const row = processLine(line, parsed.sysRanges, nextIndex)
      if (!row) continue
      rows.push(row)
      if (row.kind === 'data' && row.compliance === 'Не соотв.' && row.index != null) {
        nonCompliantIndexes.push(row.index)
      }
    }
  }

  let conclusion: string
  if (!nonCompliantIndexes.length) {
    conclusion
      = 'Сопротивление изоляции проводов и кабелей соответствует требованиям НТД.'
  } else {
    conclusion
      = `Сопротивление изоляции проводов и кабелей соответствует требованиям НТД, за исключением п/п: ${nonCompliantIndexes.join(', ')}.`
  }

  return { rows, nonCompliantIndexes, conclusion }
}
