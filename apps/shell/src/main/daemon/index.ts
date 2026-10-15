/**
 * The shell's daemon channel: the app-side handler for the control socket's
 * `office` command (the daemon's own `office_*` tool family). Importers outside
 * this folder should reach it through this barrel.
 */
export { createDaemonControl } from './daemon-control'
export type { OfficeControl, DaemonControlDeps } from './daemon-control'
