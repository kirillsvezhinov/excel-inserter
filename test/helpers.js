import { readFileSync } from 'node:fs'
import { join } from 'node:path'

export const fixture = (name) => readFileSync(join(process.cwd(), 'test/fixtures', `${name}.html`), 'utf8')
