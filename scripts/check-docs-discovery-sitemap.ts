import { extractSitemapUrls } from './check-docs-discovery-core.ts'

export interface SitemapPublishedSourcesVerification {
  actualSitemapUrls: string[]
  sourceFilesToVerify: string[]
}

export function requireSitemapPublishedSources({
  expectedSitemapUrls,
  sitemap,
  siteRoot,
  sourceFilesByUrl,
}: {
  expectedSitemapUrls: readonly string[]
  sitemap: string
  siteRoot: string
  sourceFilesByUrl: ReadonlyMap<string, string>
}): SitemapPublishedSourcesVerification {
  const actualSitemapUrls = extractSitemapUrls(sitemap)

  requireNoDuplicateSitemapUrls(actualSitemapUrls)

  if (actualSitemapUrls.length !== expectedSitemapUrls.length) {
    throw new Error(`sitemap has ${String(actualSitemapUrls.length)} urls, expected ${String(expectedSitemapUrls.length)}`)
  }

  const sourceFilesToVerify: string[] = []
  for (const expectedUrl of expectedSitemapUrls) {
    if (!actualSitemapUrls.includes(expectedUrl)) {
      throw new Error(`sitemap is missing ${expectedUrl}`)
    }

    const sourceFile = sourceFilesByUrl.get(expectedUrl)
    if (sourceFile === undefined) {
      throw new Error(`no source file mapping for ${expectedUrl}`)
    }
    sourceFilesToVerify.push(sourceFile)
  }

  for (const url of actualSitemapUrls) {
    if (!url.startsWith(siteRoot)) {
      throw new Error(`sitemap url is outside ${siteRoot}: ${url}`)
    }
  }

  return { actualSitemapUrls, sourceFilesToVerify }
}

function requireNoDuplicateSitemapUrls(actualSitemapUrls: readonly string[]): void {
  const seenUrls = new Set<string>()
  for (const url of actualSitemapUrls) {
    if (seenUrls.has(url)) {
      throw new Error(`sitemap contains duplicate url: ${url}`)
    }
    seenUrls.add(url)
  }
}
