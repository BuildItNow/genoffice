/**
 * The daemon's own `office_*` renderer dispatcher, gathered under one entry.
 * Consume it through this barrel (`./daemon`) rather than reaching into
 * `commands.ts`, so the surface the renderer bridge depends on stays one import.
 */
export {
  addressInBounds,
  describeOpError,
  dispatchDaemonCommand,
  isDaemonCommand,
} from './commands'
export type {
  DaemonCommandName,
  DaemonCommandResult,
  RangeReadResult,
  SheetEditPayload,
  SheetEditResult,
  SheetInfoPayload,
  SheetInfoResult,
  SheetReadFormatsPayload,
  SheetReadFormatsResult,
  SheetReadPayload,
  SheetReadResult,
  SheetSearchMatch,
  SheetSearchPayload,
  SheetSearchResult,
} from './types'
