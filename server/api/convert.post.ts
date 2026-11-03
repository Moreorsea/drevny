import { buildProtocol, generateIsolationDocx, parseWorkbook } from '#isolation'

export default defineEventHandler(async (event) => {
  const form = await readMultipartFormData(event)
  if (!form?.length) {
    throw createError({ statusCode: 400, statusMessage: 'Ожидается файл шаблона (.xlsx)' })
  }

  const file = form.find(p => p.name === 'file' && p.data?.length)
  if (!file) {
    throw createError({ statusCode: 400, statusMessage: 'Поле file с .xlsx не найдено' })
  }

  const filename = file.filename || 'template.xlsx'
  if (!/\.xlsx$/i.test(filename)) {
    throw createError({ statusCode: 400, statusMessage: 'Нужен файл формата .xlsx' })
  }

  try {
    const parsed = await parseWorkbook(file.data)
    const protocol = buildProtocol(parsed)
    const docx = await generateIsolationDocx(protocol)

    const outName = `izol_${new Date().toISOString().slice(0, 10)}.docx`
    setHeader(event, 'Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    setHeader(event, 'Content-Disposition', `attachment; filename="${outName}"`)
    setHeader(event, 'X-Protocol-Rows', String(protocol.rows.filter(r => r.kind === 'data').length))
    setHeader(event, 'X-Non-Compliant', protocol.nonCompliantIndexes.join(','))

    return docx
  } catch (err) {
    throw createError({
      statusCode: 400,
      statusMessage: err instanceof Error ? err.message : 'Ошибка конвертации'
    })
  }
})
