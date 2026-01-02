import { readdir, readFile, writeFile } from 'fs/promises'
import { join } from 'path'

const { log } = console


const DIST_DIR = './dist'


// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════
// TRANSFORMATIONS
// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════

const TRANSFORMATIONS = [
  stripExports,
  stripImports,
  addBlankLinesBeforeFunctions,
]

function stripExports(content) {
  return content.replace(/^export\s+(function|const|let|var|class)/gm, '$1')
}

function stripImports(content) {
  return content.replace(/^import\s+.*?;?\s*$/gm, '')
}

function addBlankLinesBeforeFunctions(content) {
  return content.replace(/([^\n])\n(function\s+\w+)/g, '$1\n\n$2')
}


// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════
// RUN
// ══════════════════════════════════════════════════════════════════════════════════════════════════════════════════

(async () => {
  const files = await readdir(DIST_DIR)
  const jsFiles = files.filter(f => f.endsWith('.js'))

  for (const file of jsFiles) {
    const filePath = join(DIST_DIR, file)
    const originalContent = await readFile(filePath, 'utf-8')
    const content = TRANSFORMATIONS.reduce((accContent, transform) => transform(accContent), originalContent)

    await writeFile(filePath, content, 'utf-8')
  }

  log(`✓ Processed ${jsFiles.length} file(s)`)
})().catch(console.error)
