/**
 * The daemon is a naming/encapsulation layer over two upstream tool sets that
 * otherwise share a namespace (`create_docx` exists on both the CLI and the
 * app). Prefixing each family keeps them side by side and tells the caller, at
 * a glance, where a tool will run.
 */

/** Headless tools served by the bundled genoffice CLI: no app, no window. */
export const CLI_PREFIX = 'office_cli_'

/** Tools that drive a live GenOffice window; proxied to the app, launched on demand. */
export const GUI_PREFIX = 'office_gui_'
