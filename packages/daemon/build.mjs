// Bundles the daemon into one CommonJS file so it runs on any Node (including
// the app's bundled runtime under ELECTRON_RUN_AS_NODE, with no node_modules).
// It pulls in the genoffice CLI's MCP tool surface from source, so the bundle
// carries the same engines the CLI ships.
import { build } from 'esbuild'
import { readFile } from 'node:fs/promises'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath, pathToFileURL } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))

// Vite's `?raw` text imports (the pptx op guides are markdown files); the CLI's
// bundle carries these too, so the daemon needs the same loader.
const rawText = {
  name: 'raw-text',
  setup(b) {
    b.onResolve({ filter: /\?raw$/ }, (args) => ({
      path: resolve(args.resolveDir, args.path.replace(/\?raw$/, '')),
      namespace: 'raw-text',
    }))
    b.onLoad({ filter: /.*/, namespace: 'raw-text' }, async (args) => ({
      contents: await readFile(args.path, 'utf8'),
      loader: 'text',
    }))
  },
}

/** The bundled CLI's own electron-require guard, reused so the daemon cannot ship the same trap. */
async function findElectronRequires(bundleText) {
  const modules = []
  let current = '(unknown module)'
  for (const line of bundleText.split('\n')) {
    const header = /^\s*\/\/ ((?:\.\.\/|src\/|node_modules\/)\S+)$/.exec(line)
    if (header) current = header[1]
    else if (/\brequire\((["'])electron\1\)/.test(line) && !modules.includes(current)) {
      modules.push(current)
    }
  }
  return modules
}

const outfile = join(here, 'dist/office-daemon.cjs')

export async function bundle() {
  const { version } = JSON.parse(await readFile(join(here, 'package.json'), 'utf8'))
  await build({
    entryPoints: [join(here, 'src/main.ts')],
    outfile,
    bundle: true,
    platform: 'node',
    target: 'node22',
    format: 'cjs',
    logLevel: 'info',
    // esbuild leaves import.meta empty in cjs output; the CLI's deps use it for
    // asset lookups, so give them the bundle's own location.
    define: {
      'import.meta.url': '__daemonImportMetaUrl',
      __GENOFFICE_DAEMON_VERSION__: JSON.stringify(version),
      // The daemon bundles the CLI's MCP tool surface, which reaches the CLI
      // entrypoint (../cli/src/cli.ts). In one CJS bundle that file's
      // `require.main === module` guard is true, so the CLI's own argv handling
      // would run on import and spew its help onto the MCP stdout stream.
      // Neutralize the guard: the daemon entry is the only main here.
      'require.main': 'undefined',
    },
    banner: {
      js: "const __daemonImportMetaUrl = require('node:url').pathToFileURL(__filename).href;",
    },
    // jsdom / mermaid are only reached by renderer-side code the CLI carries; the
    // daemon never runs them, but they must stay external so they are not inlined.
    external: ['electron', 'jsdom', 'mermaid'],
    jsx: 'automatic',
    loader: {
      '.css': 'empty',
      '.ttf': 'empty',
      '.woff': 'empty',
      '.woff2': 'empty',
      '.svg': 'empty',
      '.png': 'empty',
    },
    plugins: [rawText],
  })
  const offenders = await findElectronRequires(await readFile(outfile, 'utf8'))
  if (offenders.length > 0) {
    throw new Error(
      `daemon bundle requires "electron" (unavailable under ELECTRON_RUN_AS_NODE) via:\n` +
        offenders.map((m) => `  ${m}`).join('\n'),
    )
  }
}

// imported by tests, which only want the helpers above
if (process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url) await bundle()
