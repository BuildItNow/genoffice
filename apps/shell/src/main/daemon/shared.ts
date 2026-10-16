import type { OpenDocumentTab } from '../../shared/tabs-api'
import type { SheetsControl } from '../mcp/tools/sheets-tools'
import type { OpenDocumentsControl } from '../mcp/tools/open-documents-tools'
import type { TabManager } from '../tab-manager'

/**
 * What the shell injects for the daemon's `office_*` tools, shared by every
 * per-category tool module (document-tools / sheets-tools / pdf-tools), plus the
 * two helpers those modules need in common: the visible tab list, and a trimmed
 * non-empty string. Kept apart from the dispatcher so no tool module imports
 * daemon-control and the imports stay acyclic.
 */

export interface DaemonControlDeps {
  /** the shell's tab strip; null only while no window is open (macOS, last window closed) */
  tabManager: TabManager | null
  /** the sheets renderer channel (createSheetsControl in index.ts) */
  sheets: SheetsControl
  /**
   * the `open_documents` control (createOpenDocumentsControl in index.ts): save
   * and close reuse its per-family writer/cleanup, so they behave exactly like
   * the built-in ones, across every editor family (not just sheets).
   */
  documents: OpenDocumentsControl
  /**
   * the shell's revealShellWindow: bring the shell window up, rebuilding it when
   * the user closed it (macOS keeps the process alive). Used by the tools that
   * create a tab, which need a window to put it in.
   */
  revealShellWindow: () => void
}

/** the tabs the user can see, or none while no window is open */
export async function listTabs(tabs: TabManager | null): Promise<OpenDocumentTab[]> {
  return tabs ? tabs.openDocuments() : []
}

/** a trimmed, non-empty string, else undefined */
export function asString(value: unknown): string | undefined {
  return typeof value === 'string' && value.trim() !== '' ? value.trim() : undefined
}
