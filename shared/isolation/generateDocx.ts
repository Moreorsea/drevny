import {
  AlignmentType,
  BorderStyle,
  Document,
  Packer,
  PageOrientation,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  VerticalAlign,
  WidthType,
  convertMillimetersToTwip
} from 'docx'
import type { ProtocolDataRow, ProtocolResult } from './types'

const BORDER = { style: BorderStyle.SINGLE, size: 4, color: '000000' }
const BORDERS = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER }

/** 9pt in docx half-points */
const FONT_SIZE = 18
const FONT = 'Times New Roman'

/** A4 landscape: 297 × 210 mm */
const PAGE_WIDTH = convertMillimetersToTwip(297)
const PAGE_HEIGHT = convertMillimetersToTwip(210)
const TABLE_WIDTH = convertMillimetersToTwip(277)

const COL_WIDTHS = [
  convertMillimetersToTwip(10),
  convertMillimetersToTwip(42),
  convertMillimetersToTwip(36),
  convertMillimetersToTwip(16),
  convertMillimetersToTwip(18),
  convertMillimetersToTwip(14),
  convertMillimetersToTwip(14),
  convertMillimetersToTwip(14),
  convertMillimetersToTwip(16),
  convertMillimetersToTwip(16),
  convertMillimetersToTwip(16),
  convertMillimetersToTwip(14),
  convertMillimetersToTwip(14),
  convertMillimetersToTwip(14),
  convertMillimetersToTwip(14),
  convertMillimetersToTwip(19)
]

function buildDocument(result: ProtocolResult) {
  const bodyRows: TableRow[] = [
    ...headerRows(),
    ...result.rows.map(r => (r.kind === 'title' ? titleRow(r.title || '') : dataRow(r)))
  ]

  const mainTable = new Table({
    width: { size: TABLE_WIDTH, type: WidthType.DXA },
    columnWidths: COL_WIDTHS,
    rows: bodyRows
  })

  const conclusion = new Paragraph({
    spacing: { before: 240 },
    children: [
      new TextRun({
        text: `Заключение: ${result.conclusion}`,
        font: FONT,
        size: FONT_SIZE
      })
    ]
  })

  return new Document({
    styles: {
      default: {
        document: {
          run: {
            font: FONT,
            size: FONT_SIZE
          }
        }
      }
    },
    sections: [
      {
        properties: {
          page: {
            // A4 landscape. `docx` stores 210×297 + w:orient="landscape" (OOXML).
            size: {
              orientation: PageOrientation.LANDSCAPE,
              width: PAGE_WIDTH,
              height: PAGE_HEIGHT
            },
            margin: {
              top: convertMillimetersToTwip(10),
              bottom: convertMillimetersToTwip(10),
              left: convertMillimetersToTwip(10),
              right: convertMillimetersToTwip(10)
            }
          }
        },
        children: [mainTable, conclusion]
      }
    ]
  })
}

function para(content: string, opts?: { bold?: boolean, center?: boolean }) {
  return new Paragraph({
    alignment: opts?.center === false ? AlignmentType.LEFT : AlignmentType.CENTER,
    children: [
      new TextRun({
        text: content,
        bold: opts?.bold,
        size: FONT_SIZE,
        font: FONT
      })
    ]
  })
}

function cell(
  content: string | Paragraph[],
  opts?: {
    bold?: boolean
    center?: boolean
    columnSpan?: number
    rowSpan?: number
    width?: number
  }
) {
  const paragraphs = Array.isArray(content)
    ? content
    : [para(content, { bold: opts?.bold, center: opts?.center })]

  return new TableCell({
    borders: BORDERS,
    columnSpan: opts?.columnSpan,
    rowSpan: opts?.rowSpan,
    width: { size: opts?.width ?? 800, type: WidthType.DXA },
    verticalAlign: VerticalAlign.CENTER,
    children: paragraphs
  })
}

function lines(items: string[], bold = true): Paragraph[] {
  return items.map(line => para(line, { bold }))
}

function headerRows(): TableRow[] {
  const row1 = new TableRow({
    tableHeader: true,
    children: [
      cell(lines(['№', 'п/п']), { rowSpan: 2, width: COL_WIDTHS[0] }),
      cell(
        lines([
          'Наименование линий,',
          'электрических машин по проекту,',
          'рабочее напряжение.'
        ]),
        { rowSpan: 2, width: COL_WIDTHS[1] }
      ),
      cell(
        lines([
          'Марка провода, кабеля, кол-во жил,',
          'сечение провода, кабеля, мм²'
        ]),
        { rowSpan: 2, width: COL_WIDTHS[2] }
      ),
      cell(lines(['Напряжение', 'мегаомметра, В']), { rowSpan: 2, width: COL_WIDTHS[3] }),
      cell(
        lines(['Допустимое сопротивление изоляции,', 'не менее МОм']),
        { rowSpan: 2, width: COL_WIDTHS[4] }
      ),
      cell('Сопротивление изоляции, МОм', {
        columnSpan: 10,
        width: COL_WIDTHS.slice(5, 15).reduce((a, b) => a + b, 0),
        bold: true
      }),
      cell('Соотв. нормам', { rowSpan: 2, width: COL_WIDTHS[15], bold: true })
    ]
  })

  const sub = ['А-В', 'B-C', 'C-A', 'А-N (PEN)', 'B-N (PEN)', 'C-N (PEN)', 'А-PE', 'B-PE', 'C-PE', 'N-PE']
  const row2 = new TableRow({
    tableHeader: true,
    children: sub.map((label, i) => cell(label, { width: COL_WIDTHS[5 + i]!, bold: true }))
  })

  return [row1, row2]
}

function titleRow(title: string): TableRow {
  return new TableRow({
    children: [
      cell('', { width: COL_WIDTHS[0] }),
      cell(title, { width: COL_WIDTHS[1], bold: true }),
      ...COL_WIDTHS.slice(2).map(w => cell('', { width: w }))
    ]
  })
}

function dataRow(row: ProtocolDataRow): TableRow {
  const resistances = row.resistances ?? Array.from({ length: 10 }, () => '')
  return new TableRow({
    children: [
      cell(row.index != null ? String(row.index) : '', { width: COL_WIDTHS[0] }),
      cell(row.name ?? '', { width: COL_WIDTHS[1], center: false }),
      cell(row.cable ?? '', { width: COL_WIDTHS[2] }),
      cell(row.voltage ?? '', { width: COL_WIDTHS[3] }),
      cell(row.allowed ?? '', { width: COL_WIDTHS[4] }),
      ...resistances.map((v, i) => cell(v, { width: COL_WIDTHS[5 + i]! })),
      cell(row.compliance ?? '', { width: COL_WIDTHS[15] })
    ]
  })
}

/** Browser-friendly blob for downloads. */
export async function generateIsolationDocxBlob(result: ProtocolResult): Promise<Blob> {
  return Packer.toBlob(buildDocument(result))
}

/** Node/API-friendly bytes. */
export async function generateIsolationDocx(result: ProtocolResult): Promise<Uint8Array> {
  return Packer.toBuffer(buildDocument(result))
}
