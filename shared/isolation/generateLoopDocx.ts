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
import type { LoopDataRow, LoopProtocolResult } from './types'

const BORDER = { style: BorderStyle.SINGLE, size: 8, color: '000000' }
const BORDERS = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER }

const FONT_SIZE = 18
const FONT = 'Times New Roman'

/** A4 portrait measures — docx swaps on LANDSCAPE → w≈16834 h≈11909 */
const PAGE_WIDTH = convertMillimetersToTwip(210)
const PAGE_HEIGHT = convertMillimetersToTwip(297)

/** Margins from петля.docx pgMar */
const MARGIN_TOP = 1417
const MARGIN_RIGHT = 720
const MARGIN_BOTTOM = 720
const MARGIN_LEFT = 691

/** Exact widths from петля.docx tblGrid (sum = 14700) */
const COL_WIDTHS = [
  525, 2865, 1095, 1155, 705, 1590,
  720, 720, 720, 720, 720, 720,
  630, 885, 930
]
const TABLE_WIDTH = COL_WIDTHS.reduce((a, b) => a + b, 0)

const ZERO_CELL_MARGINS = { top: 0, bottom: 0, left: 0, right: 0 }

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
        lines(['Проверяемый участок цепи, место установки аппарата защиты']),
        { rowSpan: 2, width: COL_WIDTHS[1] }
      ),
      cell('Аппарат защиты от сверхтока', {
        columnSpan: 4,
        width: COL_WIDTHS[2]! + COL_WIDTHS[3]! + COL_WIDTHS[4]! + COL_WIDTHS[5]!,
        bold: true
      }),
      cell(
        lines(['Измеренное значение сопротивления цепи  «фаза – нуль», Ом']),
        {
          columnSpan: 3,
          width: COL_WIDTHS[6]! + COL_WIDTHS[7]! + COL_WIDTHS[8]!
        }
      ),
      cell(
        lines(['Измеренное значение тока однофазного замыкания, А']),
        {
          columnSpan: 3,
          width: COL_WIDTHS[9]! + COL_WIDTHS[10]! + COL_WIDTHS[11]!
        }
      ),
      cell(lines(['Время сраб. аппарата', 'защиты, с']), {
        columnSpan: 2,
        width: COL_WIDTHS[12]! + COL_WIDTHS[13]!
      }),
      cell(lines(['Соотв.', 'нормам']), { rowSpan: 2, width: COL_WIDTHS[14] })
    ]
  })

  const row2 = new TableRow({
    tableHeader: true,
    children: [
      cell(lines(['Типовое обозначение']), { width: COL_WIDTHS[2] }),
      cell(lines(['Тип расцепителя']), { width: COL_WIDTHS[3] }),
      cell(lines(['Ном.', 'ток, А']), { width: COL_WIDTHS[4] }),
      cell(lines(['Диапазон тока срабатывания расцепителя КЗ, А']), { width: COL_WIDTHS[5] }),
      cell('A', { width: COL_WIDTHS[6], bold: true }),
      cell('B', { width: COL_WIDTHS[7], bold: true }),
      cell('C', { width: COL_WIDTHS[8], bold: true }),
      cell('A', { width: COL_WIDTHS[9], bold: true }),
      cell('B', { width: COL_WIDTHS[10], bold: true }),
      cell('C', { width: COL_WIDTHS[11], bold: true }),
      cell('Доп.', { width: COL_WIDTHS[12], bold: true }),
      cell(lines(['По время токовой хар-ке']), { width: COL_WIDTHS[13] })
    ]
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

function dataRow(row: LoopDataRow): TableRow {
  const meas = row.measurements ?? Array.from({ length: 6 }, () => '-')
  const dash = (value: string | undefined) => (value == null || value === '' ? '-' : value)
  return new TableRow({
    children: [
      cell(row.index != null ? String(row.index) : '-', { width: COL_WIDTHS[0] }),
      cell(row.name ?? '', { width: COL_WIDTHS[1], center: false }),
      cell(dash(row.breakerMark), { width: COL_WIDTHS[2] }),
      cell(dash(row.tripType), { width: COL_WIDTHS[3] }),
      cell(dash(row.nominalCurrent), { width: COL_WIDTHS[4] }),
      cell(dash(row.baseRange), { width: COL_WIDTHS[5] }),
      ...meas.map((v, i) => cell(dash(v), { width: COL_WIDTHS[6 + i]! })),
      cell(dash(row.timeLimit), { width: COL_WIDTHS[12] }),
      cell(dash(row.timeCurve), { width: COL_WIDTHS[13] }),
      cell(dash(row.compliance), { width: COL_WIDTHS[14] })
    ]
  })
}

function buildDocument(result: LoopProtocolResult) {
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
          run: { font: FONT, size: FONT_SIZE }
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
              header: 720,
              footer: 720
            }
          }
        },
        children: [mainTable, conclusion]
      }
    ]
  })
}

export async function generateLoopDocxBlob(result: LoopProtocolResult): Promise<Blob> {
  return Packer.toBlob(buildDocument(result))
}

export async function generateLoopDocx(result: LoopProtocolResult): Promise<Uint8Array> {
  return Packer.toBuffer(buildDocument(result))
}
