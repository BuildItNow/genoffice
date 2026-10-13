// The build replaces __GENOFFICE_DAEMON_VERSION__ with the package version; a
// plain `tsc`/vitest run leaves it undefined (typeof on an undeclared identifier
// is safe), so the fallback keeps both paths working.
declare const __GENOFFICE_DAEMON_VERSION__: string | undefined

export const DAEMON_VERSION =
  typeof __GENOFFICE_DAEMON_VERSION__ === 'string' ? __GENOFFICE_DAEMON_VERSION__ : '0.1.0'
