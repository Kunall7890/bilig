const repoRelativeRootPattern = /(?:\/[^\s:'")]+)+\/(?=(?:apps|docs|e2e|packages|scripts)\/)/gu

export function compactRepoLocalPaths(value: string, rootDir: string): string {
  return value.replaceAll(rootDir, '<repo>').replace(repoRelativeRootPattern, '<repo>/')
}
