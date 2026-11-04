import fs from 'node:fs'
import { parseWorkbook } from '../shared/isolation/parseExcel'
import { buildProtocol } from '../shared/isolation/isolationRules'
import { generateIsolationDocx } from '../shared/isolation/generateDocx'

const buf = fs.readFileSync(
  `${process.env.USERPROFILE}/Downloads/Telegram Desktop/для теста.xlsx`
)
const parsed = await parseWorkbook(buf)
const protocol = buildProtocol(parsed)

for (const r of protocol.rows) {
  if (r.kind === 'title') {
    console.log('TITLE', r.title)
    continue
  }
  console.log(JSON.stringify({
    i: r.index ?? null,
    name: r.name,
    cable: r.cable,
    voltage: r.voltage,
    allowed: r.allowed,
    R: r.resistances,
    c: r.compliance
  }))
}

fs.writeFileSync('tmp_izol_smoke.docx', await generateIsolationDocx(protocol))
console.log('wrote docx, data rows', protocol.rows.filter(r => r.kind === 'data').length)
