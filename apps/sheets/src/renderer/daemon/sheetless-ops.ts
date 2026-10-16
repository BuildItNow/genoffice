import { z } from 'zod'
import { workbookOperationSchema } from '@genoffice/xlsx-gateway/domain/workbook-dsl'

/**
 * The workbook-DSL ops that name no worksheet — the workbook-level ops
 * (add_sheet, add_defined_name, …). Their schema carries no `sheetId`, so a
 * default sheet must not be seeded into them; every other op is sheet-scoped and
 * takes one (`seedSheetId` in `./commands`).
 *
 * Filtered out of `workbookOperationSchema` rather than listed by hand, so it
 * cannot drift from the ops: a new workbook-level op joins it the moment it is
 * added to the schema without a `sheetId`.
 */
export const SHEETLESS_OPS: ReadonlySet<string> = new Set(
  workbookOperationSchema.options
    .filter((option) => !('sheetId' in option.shape))
    .map((option) => (option.shape.op as z.ZodLiteral<string>).value),
)
