import fs from 'node:fs'
import { parseWorkbook } from '../shared/isolation/parseExcel'
import { buildLoopProtocol } from '../shared/isolation/loopRules'
import { generateLoopDocx } from '../shared/isolation/generateLoopDocx'

const path
  = process.argv[2]
    || `${process.env.USERPROFILE}/Downloads/Telegram Desktop/шаблон (2).xlsx`

const buf = fs.readFileSync(path)
const parsed = await parseWorkbook(buf)
console.log('tripCurves', parsed.tripCurves)
console.log('loopK', parsed.loopKRanges)
console.log('shields', parsed.shields.map(s => ({ name: s.name, lines: s.lines.length })))

const protocol = buildLoopProtocol(parsed)
for (const r of protocol.rows.slice(0, 12)) {
  console.log(JSON.stringify(r))
}
fs.writeFileSync('tmp_loop_smoke.docx', await generateLoopDocx(protocol))
console.log('wrote tmp_loop_smoke.docx', 'nonCompliant', protocol.nonCompliantIndexes)
