// Read-unit conversion is shared with the app's MCP reads (`@genoffice/pptx-ops`),
// so geometry reads the same whether it came from the CLI or an open tab.
export { READ_UNITS, isReadUnit, fromEmu, emuPer, type ReadUnit } from '@genoffice/pptx-ops'
