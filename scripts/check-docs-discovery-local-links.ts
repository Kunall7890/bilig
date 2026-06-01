import { existsSync, readFileSync } from 'node:fs'
import { dirname, join, posix } from 'node:path'
import { docsSiteSources } from './check-docs-discovery-site-sources.ts'

function stripAnchorAndQuery(href: string): string {
  return href.split(/[?#]/u)[0] ?? ''
}

function isGeneratedHtmlLink(href: string): boolean {
  return stripAnchorAndQuery(href).endsWith('.html')
}

function resolveDocsHtmlLink(href: string, currentPage: string): string | undefined {
  if (!isGeneratedHtmlLink(href) || href.startsWith('#')) {
    return undefined
  }

  if (/^[a-z][a-z0-9+.-]*:/iu.test(href) && !/^https?:\/\//iu.test(href)) {
    return undefined
  }

  if (/^https?:\/\//iu.test(href)) {
    const url = new URL(href)
    if (url.hostname !== 'proompteng.github.io' || !url.pathname.startsWith('/bilig/')) {
      return undefined
    }
    return stripAnchorAndQuery(url.pathname.replace(/^\/bilig\//u, ''))
  }

  const path = stripAnchorAndQuery(href)
  if (path.startsWith('/bilig/')) {
    return path.replace(/^\/bilig\//u, '')
  }

  if (path.startsWith('/')) {
    return undefined
  }

  const pageDir = dirname(currentPage)
  const baseDir = pageDir === '.' ? '' : pageDir
  return posix.normalize(posix.join(baseDir, path.replace(/^\.\//u, '')))
}

function extractLocalHtmlLinks(source: string, currentPage: string): string[] {
  const links = new Set<string>()
  const push = (href: string | undefined): void => {
    if (href === undefined) {
      return
    }
    const resolved = resolveDocsHtmlLink(href.trim(), currentPage)
    if (resolved !== undefined && resolved.length > 0) {
      links.add(resolved)
    }
  }

  for (const match of source.matchAll(/\bhref=["']([^"']+)["']/gu)) {
    push(match[1])
  }
  for (const match of source.matchAll(/!?\[[^\]]*\]\(([^)\s]+)(?:\s+"[^"]*")?\)/gu)) {
    push(match[1])
  }
  for (const match of source.matchAll(/^\[[^\]]+\]:\s+(\S+)/gmu)) {
    push(match[1])
  }

  return [...links].toSorted()
}

function markdownWouldPublishHtml(sourcePath: string, page: string): boolean {
  if (!page.endsWith('.html') || !sourcePath.endsWith('.md')) {
    return true
  }
  return readFileSync(sourcePath, 'utf8').startsWith('---\n')
}

export function requireDocsLocalHtmlLinksHaveSources(docsRoot: string): void {
  const sourceFilesByPage = new Map(docsSiteSources.map(([urlPath, sourceFile]) => [urlPath, sourceFile]))
  const failures: string[] = []

  for (const [page, sourceFile] of docsSiteSources) {
    const sourcePath = join(docsRoot, sourceFile)
    if (!existsSync(sourcePath)) {
      failures.push(`${page}: source docs/${sourceFile} is missing`)
      continue
    }
    if (!markdownWouldPublishHtml(sourcePath, page)) {
      failures.push(`${page}: docs/${sourceFile} is missing YAML front matter`)
      continue
    }

    const sourceText = readFileSync(sourcePath, 'utf8')
    for (const linkedPage of extractLocalHtmlLinks(sourceText, page)) {
      const linkedSource = sourceFilesByPage.get(linkedPage)
      if (linkedSource === undefined) {
        failures.push(`${page} -> ${linkedPage}: missing from docsSiteSources`)
        continue
      }

      const linkedSourcePath = join(docsRoot, linkedSource)
      if (!existsSync(linkedSourcePath)) {
        failures.push(`${page} -> ${linkedPage}: missing source docs/${linkedSource}`)
        continue
      }
      if (!markdownWouldPublishHtml(linkedSourcePath, linkedPage)) {
        failures.push(`${page} -> ${linkedPage}: docs/${linkedSource} is missing YAML front matter`)
      }
    }
  }

  if (failures.length > 0) {
    throw new Error(`published docs contain local .html links that cannot be emitted:\n${failures.join('\n')}`)
  }
}
