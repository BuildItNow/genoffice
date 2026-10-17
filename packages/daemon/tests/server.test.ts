import { mkdtempSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { Client } from '@modelcontextprotocol/sdk/client/index.js'
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'
import { CLI_PREFIX, OFFICE_PREFIX } from '../src/names'
import { createDaemonServer, type DaemonServer } from '../src/server'

/**
 * The daemon's full surface is exercised over an in-memory MCP transport: a
 * real `Client` speaks to a real `McpServer`, the office_* family is always
 * registered, and the app itself is absent (so live-window calls report it
 * unavailable).
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

  it('lists every tool family, keeping the bare upstream names private', async () => {
    const { tools } = await client.listTools()
    const names = tools.map((t) => t.name)
    expect(names).toContain('office_cli_info')
    // the daemon's own live-window family
    expect(names).toContain('office_list')
    expect(names).toContain('office_open')
    expect(names).toContain('office_sheet_read')
    expect(names).toContain('office_sheet_ops_help')
    expect(names).toContain('office_pdf_info')
    expect(names).toContain('office_pdf_read')
    expect(names).toContain('office_pdf_search')

    // every tool carries a family prefix; no bare CLI/office name leaks through
    for (const name of names) {
      expect(name.startsWith(CLI_PREFIX) || name.startsWith(OFFICE_PREFIX)).toBe(true)
    }
    expect(names).not.toContain('info')
    expect(names).not.toContain('list')
    expect(names).not.toContain('sheet_read')
    expect(names).not.toContain('read_pdf')
    expect(names).not.toContain('search_pdf')

    // the cloud family is left out: no office_* tool sends work off the machine
    for (const cloud of ['capabilities', 'search', 'image', 'media']) {
      expect(names).not.toContain(`office_cli_${cloud}`)
    }
  })

  it('leaves the whole CLI family out unless asked for', async () => {
    const officeOnly = await createDaemonServer({
      cwd: process.cwd(),
      env: { GENOFFICE_USER_DATA: userData, GENOFFICE_APP_BIN: '/nonexistent/genoffice' },
      log: () => {},
    })
    const [a, b] = InMemoryTransport.createLinkedPair()
    await officeOnly.server.connect(a)
    const c = new Client({ name: 'office-only', version: '0' })
    await c.connect(b)
    try {
      const names = (await c.listTools()).tools.map((t) => t.name)
      expect(names.some((n) => n.startsWith(CLI_PREFIX))).toBe(false)
      // the office_* family, by contrast, is always served
      expect(names).toContain('office_list')
      expect(names).toContain('office_sheet_read')
    } finally {
      await c.close()
      officeOnly.dispose()
    }
  })

  it('reports the office_* family unavailable when no app answers', async () => {
    const result = await client.callTool({ name: 'office_list', arguments: {} })
    expect(result.isError).toBe(true)
    expect(firstText(result)).toContain('office_list')
  })
})
