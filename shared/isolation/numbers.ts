/** Normalize decimal comma → dot and trim. */
export function normalizeNumberString(raw: unknown): string {
  if (raw === null || raw === undefined) return ''
  return String(raw).trim().replace(/\s+/g, '').replace(/,/g, '.')
}

export function parseNumber(raw: unknown): number | null {
  const s = normalizeNumberString(raw)
  if (!s) return null
  const n = Number(s)
  return Number.isFinite(n) ? n : null
}

/** Extract all numbers from section-like strings: "2,5+1,5", "16", "2.5". */
export function extractNumbers(raw: unknown): number[] {
  const s = normalizeNumberString(raw)
  if (!s) return []
  const matches = s.match(/\d+(?:\.\d+)?/g) || []
  return matches.map(Number).filter(n => Number.isFinite(n))
}

export function maxSection(raw: unknown): number | null {
  const nums = extractNumbers(raw)
  if (!nums.length) return null
  return Math.max(...nums)
}

/**
 * Format resistance for Word:
 * < 1 → hundredths; ≤ 99.9 → tenths; > 99.9 → integer.
 * Always uses `.` as decimal separator.
 */
export function formatResistance(value: number): string {
  if (!Number.isFinite(value)) return ''
  if (value < 1) {
    return (Math.round(value * 100) / 100).toFixed(2)
  }
  if (value <= 99.9) {
    const rounded = Math.round(value * 10) / 10
    return Number.isInteger(rounded) ? String(rounded) : rounded.toFixed(1)
  }
  return String(Math.round(value))
}

/** Random float in [min, max]. */
export function randomInRange(min: number, max: number): number {
  let a = min
  let b = max
  if (b < a) {
    const t = a
    a = b
    b = t
  }
  if (a === b) return a
  return a + Math.random() * (b - a)
}
