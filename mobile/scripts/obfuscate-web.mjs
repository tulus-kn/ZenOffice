// Minifies every JS file under mobile/www/ (the built web app) with esbuild,
// so the code shipped inside the APK is not readable as plain text.
// Runs after `node mobile/scripts/build-web.mjs` and before `npx cap sync android`.
// Skips files that are already minified (*.min.js) and anything under node_modules.
// esbuild comes from the repo root node_modules (transitive dep of vite);
// if it is missing, falls back to `npx --yes esbuild` (downloads it).
//
// This stops casual snooping and casual tampering only. It is not encryption:
// anyone skilled enough can still beautify the code and decode the strings.

import { execFileSync } from 'node:child_process'
import { existsSync, readdirSync, renameSync, statSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const mobileDir = resolve(here, '..')
const repoRoot = resolve(mobileDir, '..')
const wwwDir = join(mobileDir, 'www')

function collectJs(dir, out = []) {
  for (const entry of readdirSync(dir)) {
    const p = join(dir, entry)
    if (statSync(p).isDirectory()) collectJs(p, out)
    else if (p.endsWith('.js') && !p.endsWith('.min.js')) out.push(p)
  }
  return out
}

function esbuildCmd() {
  const local = join(repoRoot, 'node_modules', '.bin', process.platform === 'win32' ? 'esbuild.exe' : 'esbuild')
  if (existsSync(local)) return [local]
  return ['npx', '--yes', 'esbuild']
}

if (!existsSync(wwwDir)) {
  console.error(`obfuscate: www dir not found: ${wwwDir} (run build-web.mjs first)`)
  process.exit(1)
}

const files = collectJs(wwwDir)
if (files.length === 0) {
  console.log('obfuscate: no JS files found in', wwwDir)
  process.exit(0)
}

const bin = esbuildCmd()
for (const f of files) {
  const tmp = `${f}.obf-tmp`
  execFileSync(bin[0], [...bin.slice(1), f, '--minify', `--outfile=${tmp}`], { stdio: 'inherit' })
  renameSync(tmp, f)
  console.log('minified', f.replace(`${wwwDir}/`, ''))
}
console.log(`obfuscate: minified ${files.length} JS file(s)`)
