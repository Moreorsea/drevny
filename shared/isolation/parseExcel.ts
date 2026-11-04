import ExcelJS from 'exceljs'
import type {
  ParsedWorkbook,
  ShieldLine,
  ShieldSheet,
  SysIsolationRange,
  SysLoopKRange,
  SysTripCurve
} from './types'
import { normalizeNumberString, parseNumber } from './numbers'

const EXCLUDED_SHEETS = new Set(['sys', 'Осн. инфо', 'ОБРАЗЕЦ', 'осн. инфо', 'образец'])

function fail(message: string): never {
  throw new Error(message)
}

function cellText(value: ExcelJS.CellValue): string {
  if (value === null || value === undefined) return ''
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') {
    return String(value).trim()
  }
  if (typeof value === 'object') {
    if ('result' in value && value.result !== undefined) return cellText(value.result as ExcelJS.CellValue)
    if ('text' in value && typeof value.text === 'string') return value.text.trim()
    if ('richText' in value && Array.isArray(value.richText)) {
      return value.richText.map(p => p.text).join('').trim()
    }
  }
  return String(value).trim()
}

function isExcludedSheet(name: string): boolean {
  return EXCLUDED_SHEETS.has(name.trim()) || EXCLUDED_SHEETS.has(name.trim().toLowerCase())
}

function parseSectionBound(
  raw: string,
  nextIsOpenEnded: boolean
): { minExclusive: number, maxInclusive: number } {
  const s = normalizeNumberString(raw)
  if (s.startsWith('..')) {
    return { minExclusive: Number.NEGATIVE_INFINITY, maxInclusive: parseNumber(s.slice(2)) ?? 0 }
  }
  if (s.endsWith('..')) {
    const edge = parseNumber(s.slice(0, -2)) ?? 0
    return { minExclusive: edge - 1e-9, maxInclusive: Number.POSITIVE_INFINITY }
  }
  const parts = s.split('..')
  if (parts.length === 2) {
    const lo = parseNumber(parts[0]) ?? 0
    const hi = parseNumber(parts[1]) ?? 0
    return {
      minExclusive: lo,
      maxInclusive: nextIsOpenEnded ? hi - 1e-9 : hi
    }
  }
  return { minExclusive: Number.NEGATIVE_INFINITY, maxInclusive: Number.POSITIVE_INFINITY }
}

function parseMomRange(raw: string): { momMin: number, momMax: number } {
  const s = normalizeNumberString(raw)
  const m = s.match(/^(\d+(?:\.\d+)?)\s*[-–—]\s*(\d+(?:\.\d+)?)$/)
  if (m) {
    return { momMin: Number(m[1]), momMax: Number(m[2]) }
  }
  const n = parseNumber(s)
  if (n !== null) return { momMin: n, momMax: n }
  return { momMin: 100, momMax: 100 }
}

function parseDotRange(raw: string): { min: number, max: number } | null {
  const s = normalizeNumberString(raw)
  const m = s.match(/^(\d+(?:\.\d+)?)\.\.(\d+(?:\.\d+)?)$/)
  if (m) return { min: Number(m[1]), max: Number(m[2]) }
  if (s.startsWith('..')) {
    const max = parseNumber(s.slice(2))
    return max === null ? null : { min: Number.NEGATIVE_INFINITY, max }
  }
  if (s.endsWith('..')) {
    const min = parseNumber(s.slice(0, -2))
    return min === null ? null : { min, max: Number.POSITIVE_INFINITY }
  }
  return null
}

export function parseSysRanges(sheet: ExcelJS.Worksheet): SysIsolationRange[] {
  const rows: { sectionRaw: string, voltage: number, momRaw: string }[] = []
  for (let r = 13; r <= 17; r++) {
    const sectionRaw = cellText(sheet.getCell(r, 8).value)
    const voltage = parseNumber(sheet.getCell(r, 9).value) ?? 1000
    const momRaw = cellText(sheet.getCell(r, 10).value)
    if (!sectionRaw) continue
    rows.push({ sectionRaw, voltage, momRaw })
  }

  return rows.map((row, index) => {
    const nextRaw = rows[index + 1]?.sectionRaw || ''
    const nextIsOpenEnded = normalizeNumberString(nextRaw).endsWith('..')
    const bounds = parseSectionBound(row.sectionRaw, nextIsOpenEnded)
    const mom = parseMomRange(row.momRaw)
    return {
      minExclusive: bounds.minExclusive,
      maxInclusive: bounds.maxInclusive,
      voltage: row.voltage,
      momMin: mom.momMin,
      momMax: mom.momMax
    }
  })
}

export function parseTripCurves(sheet: ExcelJS.Worksheet): SysTripCurve[] {
  const curves: SysTripCurve[] = []
  for (let r = 19; r <= 21; r++) {
    const letters = cellText(sheet.getCell(r, 8).value)
    const displayLetter = cellText(sheet.getCell(r, 9).value)
    const rangeRaw = cellText(sheet.getCell(r, 10).value)
    if (!letters || !displayLetter) continue
    const range = parseDotRange(rangeRaw)
    if (!range) continue
    curves.push({
      letters,
      displayLetter,
      rangeMin: range.min,
      rangeMax: range.max
    })
  }
  return curves
}

export function parseLoopKRanges(sheet: ExcelJS.Worksheet): SysLoopKRange[] {
  const rows: { raw: string, k: number }[] = []
  for (let r = 24; r <= 28; r++) {
    const raw = cellText(sheet.getCell(r, 8).value)
    const k = parseNumber(sheet.getCell(r, 9).value)
    if (!raw || k === null) continue
    rows.push({ raw, k })
  }
  return rows.map((row, index) => {
    const nextRaw = rows[index + 1]?.raw || ''
    const nextIsOpenEnded = normalizeNumberString(nextRaw).endsWith('..')
    const bounds = parseSectionBound(row.raw, nextIsOpenEnded)
    return {
      minExclusive: bounds.minExclusive,
      maxInclusive: bounds.maxInclusive,
      k: row.k
    }
  })
}

function resolveShieldName(sheet: ExcelJS.Worksheet, fallback: string): string {
  for (let c = 1; c <= 30; c++) {
    const t = cellText(sheet.getCell(1, c).value)
    if (/название\s*щита/i.test(t)) {
      const next = cellText(sheet.getCell(1, c + 1).value)
      if (next) return next
    }
  }
  const b1 = cellText(sheet.getCell(1, 2).value)
  if (b1 && !/инд\.?\s*обозн/i.test(b1) && !/наименование/i.test(b1)) {
    return b1
  }
  return fallback
}

function isHeaderRow(values: string[]): boolean {
  const joined = values.join(' ').toLowerCase()
  return joined.includes('марка кабеля') || joined.includes('адресация') || joined.includes('qs/qf')
}

function isTitleRow(col1: string, col2: string, col3: string, cable: string, phase: string, cores: string): boolean {
  if (cable || phase || cores) return false
  if (col1) return false
  return Boolean(col2 || col3)
}

function parseShieldSheet(sheet: ExcelJS.Worksheet): ShieldSheet {
  const name = resolveShieldName(sheet, sheet.name)
  const lines: ShieldLine[] = []

  sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    const col1 = cellText(row.getCell(1).value)
    const col2 = cellText(row.getCell(2).value)
    const col3 = cellText(row.getCell(3).value)
    const phase = cellText(row.getCell(4).value)
    const cableMark = cellText(row.getCell(5).value)
    const cores = cellText(row.getCell(6).value)
    const section = cellText(row.getCell(7).value)
    const breakerMark = cellText(row.getCell(8).value)
    const nominalCurrent = cellText(row.getCell(9).value)
    const ovv = cellText(row.getCell(10).value)
    const md = cellText(row.getCell(11).value)
    const multiplicity = cellText(row.getCell(12).value)
    const isolationHint = cellText(row.getCell(14).value)
    const loopHint = cellText(row.getCell(17).value)
    const loopTime = cellText(row.getCell(18).value)

    const probe = [col1, col2, col3, phase, cableMark, cores, section]
    if (isHeaderRow(probe)) return
    if (probe.every(v => !v) && !isolationHint && !loopHint && !breakerMark) return

    const extra = {
      breakerMark,
      nominalCurrent,
      ovv,
      md,
      multiplicity,
      isolationHint,
      loopHint,
      loopTime
    }

    if (isTitleRow(col1, col2, col3, cableMark, phase, cores)) {
      lines.push({
        excelRow: rowNumber,
        isTitle: true,
        title: col2 || col3,
        col1,
        col2,
        col3,
        phase,
        cableMark,
        cores,
        section,
        ...extra
      })
      return
    }

    if (!col1 && !col2 && !cableMark && !breakerMark) return

    lines.push({
      excelRow: rowNumber,
      isTitle: false,
      col1,
      col2,
      col3,
      phase,
      cableMark,
      cores,
      section,
      ...extra
    })
  })

  return { name, lines }
}

export async function parseWorkbook(data: ArrayBuffer | Uint8Array): Promise<ParsedWorkbook> {
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(data as never)

  const sys = wb.getWorksheet('sys')
  if (!sys) fail('В файле нет листа "sys"')

  const sysRanges = parseSysRanges(sys)
  if (!sysRanges.length) {
    fail('Не удалось прочитать диапазоны изоляции с листа sys (G12:J17)')
  }

  const tripCurves = parseTripCurves(sys)
  const loopKRanges = parseLoopKRanges(sys)

  let shields: ShieldSheet[] = []
  wb.eachSheet((sheet) => {
    if (isExcludedSheet(sheet.name)) return
    shields.push(parseShieldSheet(sheet))
  })

  if (!shields.length) {
    const sample = wb.getWorksheet('ОБРАЗЕЦ')
    if (sample) {
      const parsed = parseShieldSheet(sample)
      parsed.name = parsed.name === 'ОБРАЗЕЦ' ? 'ОБРАЗЕЦ' : parsed.name
      shields = [parsed]
    }
  }

  if (!shields.length) {
    fail('Нет листов со щитами (ожидаются листы кроме sys, «Осн. инфо», «ОБРАЗЕЦ»)')
  }

  return { shields, sysRanges, tripCurves, loopKRanges }
}
