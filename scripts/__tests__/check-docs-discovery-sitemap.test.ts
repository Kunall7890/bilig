import { describe, expect, it } from 'vitest'

import { requireSitemapPublishedSources } from '../check-docs-discovery-sitemap.ts'

const siteRoot = 'https://proompteng.github.io/bilig/'
const expectedSitemapUrls = [`${siteRoot}index.html`, `${siteRoot}llms.txt`]
const sourceFilesByUrl = new Map([
  [`${siteRoot}index.html`, 'index.md'],
  [`${siteRoot}llms.txt`, 'llms.txt'],
])

function sitemapFor(urls: readonly string[]): string {
  return [
    '<?xml version="1.0" encoding="UTF-8"?>',
    '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">',
    ...urls.map((url) => `  <url><loc>${url}</loc></url>`),
    '</urlset>',
  ].join('\n')
}

describe('docs discovery sitemap verification', () => {
  it('returns the source files that must be checked for published front matter', () => {
    expect(
      requireSitemapPublishedSources({
        expectedSitemapUrls,
        sitemap: sitemapFor(expectedSitemapUrls),
        siteRoot,
        sourceFilesByUrl,
      }),
    ).toEqual({
      actualSitemapUrls: expectedSitemapUrls,
      sourceFilesToVerify: ['index.md', 'llms.txt'],
    })
  })

  it('rejects duplicate sitemap URLs before trusting the source mapping', () => {
    expect(() =>
      requireSitemapPublishedSources({
        expectedSitemapUrls,
        sitemap: sitemapFor([expectedSitemapUrls[0] ?? '', expectedSitemapUrls[0] ?? '']),
        siteRoot,
        sourceFilesByUrl,
      }),
    ).toThrow(`sitemap contains duplicate url: ${expectedSitemapUrls[0]}`)
  })

  it('rejects source mappings that do not cover every expected sitemap URL', () => {
    expect(() =>
      requireSitemapPublishedSources({
        expectedSitemapUrls,
        sitemap: sitemapFor(expectedSitemapUrls),
        siteRoot,
        sourceFilesByUrl: new Map([[expectedSitemapUrls[0] ?? '', 'index.md']]),
      }),
    ).toThrow(`no source file mapping for ${expectedSitemapUrls[1]}`)
  })

  it('rejects sitemap URLs outside the published site root', () => {
    const externalUrl = 'https://example.com/bilig/index.html'

    expect(() =>
      requireSitemapPublishedSources({
        expectedSitemapUrls: [externalUrl],
        sitemap: sitemapFor([externalUrl]),
        siteRoot,
        sourceFilesByUrl: new Map([[externalUrl, 'index.md']]),
      }),
    ).toThrow(`sitemap url is outside ${siteRoot}: ${externalUrl}`)
  })
})
