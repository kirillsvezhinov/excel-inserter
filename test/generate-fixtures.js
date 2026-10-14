// Writes every fixture from test/fixtures.js into test/fixtures/. Run: npm run fixtures
import { writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { fixtures } from './fixtures.js'

const dir = join(process.cwd(), 'test/fixtures')
for (const [name, build] of Object.entries(fixtures)) {
  const bytes = build()
  writeFileSync(join(dir, name), bytes)
  console.log(`${name.padEnd(24)} ${String(bytes.length).padStart(8)} bytes`)
}
