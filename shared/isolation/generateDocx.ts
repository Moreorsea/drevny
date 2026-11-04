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

/** Matches изол.docx: tblBorders sz="8" (1pt) */
const BORDER = { style: BorderStyle.SINGLE, size: 8, color: '000000' }
const BORDERS = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER }

/** 9pt in docx half-points — as in изол.docx */
const FONT_SIZE = 18
const FONT = 'Times New Roman'

/**
 * A4 paper in portrait measures (docx swaps them when orientation is LANDSCAPE).
 * Resulting OOXML: w≈16838 h≈11906 orient="landscape" — same as изол.docx.
 */
const PAGE_WIDTH = convertMillimetersToTwip(210)
const PAGE_HEIGHT = convertMillimetersToTwip(297)

/**
 * Page margins from изол.docx pgMar:
 * top/right/bottom = 720 twips (12.7 mm), left = 1440 twips (25.4 mm)
 */
const MARGIN_TOP = 720
const MARGIN_RIGHT = 720
const MARGIN_BOTTOM = 720
const MARGIN_LEFT = 1440

/**
 * Exact column widths from изол.docx w:tblGrid (sum = 14715).
 */
const COL_WIDTHS = [
  511, // №
  3032, // наименование
  1697, // марка кабеля
  870, // напряжение
  1140, // допустимое
  646, 646, 646, 646, 646, 646, 646, 646, 646, 646, // сопротивления
  1005 // соотв.
]

const TABLE_WIDTH = COL_WIDTHS.reduce((a, b) => a + b, 0)

const ZERO_CELL_MARGINS = {
  top: 0,
  bottom: 0,
  left: 0,
  right: 0
}

function run(text: string, opts?: { bold?: boolean }) {
  return new TextRun({
    text,
    bold: opts?.bold,
    size: FONT_SIZE,
    font: FONT
  })
}

function para(content: string, opts?: { bold?: boolean, center?: boolean }) {
  return new Paragraph({
    alignment: opts?.center === false ? AlignmentType.LEFT : AlignmentType.CENTER,
    spacing: { before: 0, after: 0, line: 240 },
    children: [run(content, { bold: opts?.bold })]
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
    margins: ZERO_CELL_MARGINS,
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
      cell(lines(['Соотв.', 'нормам']), { rowSpan: 2, width: COL_WIDTHS[15] })
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
      cell(title, { width: COL_WIDTHS[1], bold: true, center: false }),
      ...COL_WIDTHS.slice(2).map(w => cell('', { width: w }))
    ]
  })
}

function dataRow(row: ProtocolDataRow): TableRow {
  const resistances = row.resistances ?? Array.from({ length: 10 }, () => '-')
  const dash = (value: string | undefined) => (value == null || value === '' ? '-' : value)
  return new TableRow({
    children: [
      // №: только номер или пусто, без прочерка
      cell(row.index != null ? String(row.index) : '', { width: COL_WIDTHS[0] }),
      cell(row.name ?? '', { width: COL_WIDTHS[1], center: false }),
      cell(dash(row.cable), { width: COL_WIDTHS[2] }),
      cell(dash(row.voltage), { width: COL_WIDTHS[3] }),
      cell(dash(row.allowed), { width: COL_WIDTHS[4] }),
      ...resistances.map((v, i) => cell(dash(v), { width: COL_WIDTHS[5 + i]! })),
      cell(dash(row.compliance), { width: COL_WIDTHS[15] })
    ]
  })
}

function buildDocument(result: ProtocolResult) {
  const bodyRows: TableRow[] = [
    ...headerRows(),
    ...result.rows.map(r => (r.kind === 'title' ? titleRow(r.title || '') : dataRow(r)))
  ]

  const mainTable = new Table({
    width: { size: TABLE_WIDTH, type: WidthType.DXA },
    columnWidths: COL_WIDTHS,
    alignment: AlignmentType.CENTER,
    rows: bodyRows
  })

  const conclusion = new Paragraph({
    spacing: { before: 200, after: 0, line: 240 },
    children: [
      run('Заключение: ', { bold: true }),
      run(result.conclusion)
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
            size: {
              orientation: PageOrientation.LANDSCAPE,
              width: PAGE_WIDTH,
              height: PAGE_HEIGHT
            },
            margin: {
              top: MARGIN_TOP,
              right: MARGIN_RIGHT,
              bottom: MARGIN_BOTTOM,
              left: MARGIN_LEFT,
              header: 0,
              footer: 305
            }
          }
        },
        children: [mainTable, conclusion]
      }
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
