import fs from 'node:fs'
import path from 'node:path'

const root = path.resolve(import.meta.dirname, '..')
const template = process.argv[2]
  || path.join(process.env.USERPROFILE || '', 'Downloads', 'Telegram Desktop', 'шаблон.xlsx')

if (!fs.existsSync(template)) {
  console.error('Template not found:', template)
  process.exit(1)
}

const form = new FormData()
form.append('file', new Blob([fs.readFileSync(template)]), 'shablon.xlsx')

const base = process.env.SMOKE_URL || 'http://[::1]:3001'
const res = await fetch(`${base}/api/convert`, {
  method: 'POST',
  body: form
})

if (!res.ok) {
  console.error('HTTP', res.status, await res.text())
  process.exit(1)
}

const out = path.join(root, 'tmp_izol_smoke.docx')
fs.writeFileSync(out, Buffer.from(await res.arrayBuffer()))
console.log('OK →', out)
console.log('rows:', res.headers.get('x-protocol-rows'))
console.log('nonCompliant:', res.headers.get('x-non-compliant') || '(none)')
