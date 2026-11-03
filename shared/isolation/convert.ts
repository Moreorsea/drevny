import { buildProtocol } from './isolationRules'
import { generateIsolationDocxBlob } from './generateDocx'
import { parseWorkbook } from './parseExcel'
import type { ProtocolResult } from './types'

export interface ConvertResult {
  blob: Blob
  protocol: ProtocolResult
  fileName: string
}

export async function convertXlsxToIsolationDocx(file: File | ArrayBuffer): Promise<ConvertResult> {
  const buffer = file instanceof ArrayBuffer ? file : await file.arrayBuffer()
  const parsed = await parseWorkbook(buffer)
  const protocol = buildProtocol(parsed)
  const blob = await generateIsolationDocxBlob(protocol)
  const fileName = `izol_${new Date().toISOString().slice(0, 10)}.docx`
  return { blob, protocol, fileName }
}
