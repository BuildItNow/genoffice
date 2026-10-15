/**
 * The daemon is a naming/encapsulation layer over the tool sets that otherwise
 * share a namespace (`create_docx` exists on both the CLI and the app).
 * Prefixing each family keeps them side by side and tells the caller, at a
 * glance, where a tool will run.
 */

/** Headless tools served by the bundled genoffice CLI: no app, no window. */
export const CLI_PREFIX = 'office_cli_'

/**
 * The daemon's own family (`office_create`, `office_sheet_read`, …): it drives a
 * live window over the app's control socket, so the app's built-in MCP server —
 * and its tool list — stays untouched.
 */
export const OFFICE_PREFIX = 'office_'
