import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { CLI_PREFIX, GUI_PREFIX } from '../src/names'
import { createDaemonServer, type DaemonServer } from '../src/server'

/**
 * The daemon's full surface is exercised over an in-memory MCP transport: a
 * real `Client` speaks to a real `McpServer`, both families are registered, and
 * only the app itself is absent (the GUI family simply reports it unavailable).
 */

function firstText(result: unknown): string {
  return (result as { content: Array<{ type: 'text'; text: string }> }).content[0].text
}

describe('daemon server', () => {
  let daemon: DaemonServer
  let client: Client
  let userData: string

  beforeEach(async () => {
    userData = mkdtempSync(join(tmpdir(), 'genoffice-daemon-test-'))
    daemon = await createDaemonServer({
      cwd: process.cwd(),
      // an empty user-data dir -> no control.json -> no app is "running"
      env: { GENOFFICE_USER_DATA: userData, GENOFFICE_APP_BIN: '/nonexistent/genoffice' },
      log: () => {},
      // the CLI family is opt-in; the tests below exercise it, so turn it on
      cli: true,
      compactSchemas: true,
    })
    const [a, b] = InMemoryTransport.createLinkedPair()
    await daemon.server.connect(a)
    client = new Client({ name: 'daemon-test', version: '0' })
    await client.connect(b)
  })

  afterEach(async () => {
    await client.close()
    daemon.dispose()
    rmSync(userData, { recursive: true, force: true })
  })

  it('lists both tool families, keeping the bare upstream names private', async () => {
    const { tools } = await client.listTools()
    const names = tools.map((t) => t.name)
    expect(names).toContain('office_cli_info')
    expect(names).toContain('office_gui_opened_documents')
    expect(names).toContain('office_gui_read_pdf')
    expect(names).toContain('office_gui_open')

    // every tool carries a family prefix; no bare CLI/GUI name leaks through
    for (const name of names) {
      expect(name.startsWith(CLI_PREFIX) || name.startsWith(GUI_PREFIX)).toBe(true)
    }
    expect(names).not.toContain('info')
    expect(names).not.toContain('open_documents')
    expect(names).not.toContain('open_in_genoffice')

    // the cloud family is left out: no office_* tool sends work off the machine
    for (const cloud of ['capabilities', 'search', 'image', 'media']) {
      expect(names).not.toContain(`office_cli_${cloud}`)
    }
  })

  it('leaves the whole CLI family out unless asked for, keeping the GUI family', async () => {
    const guiOnly = await createDaemonServer({
      cwd: process.cwd(),
      env: { GENOFFICE_USER_DATA: userData, GENOFFICE_APP_BIN: '/nonexistent/genoffice' },
      log: () => {},
    })
    const [a, b] = InMemoryTransport.createLinkedPair()
    await guiOnly.server.connect(a)
    const c = new Client({ name: 'gui-only', version: '0' })
    await c.connect(b)
    try {
      const names = (await c.listTools()).tools.map((t) => t.name)
      expect(names.some((n) => n.startsWith(CLI_PREFIX))).toBe(false)
      expect(names).toContain('office_gui_opened_documents')
      expect(names).toContain('office_gui_open')
    } finally {
      await c.close()
      guiOnly.dispose()
    }
  })

  it('turns an unavailable app into an isError result, not a thrown call', async () => {
    const result = await client.callTool({
      name: 'office_gui_opened_documents',
      arguments: { action: 'list' },
    })
    expect(result.isError).toBe(true)
    expect(firstText(result).length).toBeGreaterThan(0)
  })
})
