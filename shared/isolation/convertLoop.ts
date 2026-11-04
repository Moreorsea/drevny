import { generateLoopDocxBlob } from './generateLoopDocx'
import { buildLoopProtocol } from './loopRules'
import { parseWorkbook } from './parseExcel'
import type { LoopProtocolResult } from './types'

export interface ConvertLoopResult {
  blob: Blob
  protocol: LoopProtocolResult
  fileName: string
}

export async function convertXlsxToLoopDocx(file: File | ArrayBuffer): Promise<ConvertLoopResult> {
  const buffer = file instanceof ArrayBuffer ? file : await file.arrayBuffer()
  const parsed = await parseWorkbook(buffer)
  if (!parsed.tripCurves.length) {
    throw new Error('На листе sys не найдены характеристики АВ (H19:J21)')
  }
  if (!parsed.loopKRanges.length) {
    throw new Error('На листе sys не найдены коэффициенты петли (H24:I28)')
  }
  const protocol = buildLoopProtocol(parsed)
  const blob = await generateLoopDocxBlob(protocol)
  const fileName = `petlya_${new Date().toISOString().slice(0, 10)}.docx`
  return { blob, protocol, fileName }
}
