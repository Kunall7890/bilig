import { communityLaunchPackRequiredLinks, llmsExternalSurfaceLinks } from './check-docs-discovery-growth-links.ts'

const productHuntLaunchKitRequiredText = [
  'title: Product Hunt launch kit for bilig',
  'Workbook formulas for TypeScript services and agents.',
  'start from an empty Node project, install @bilig/headless, run eval.ts',
  '46/46 comparable mean-latency rows are faster',
  'product-hunt-thumbnail.png',
  'product-hunt-gallery-01-workbook-api.png',
  'product-hunt-gallery-02-agent-readback.png',
  'product-hunt-gallery-03-node-service.png',
  'product-hunt-demo.webm',
  'try-bilig-headless-in-node.html',
  'what-workpaper-benchmark-proves.html',
  'where-bilig-is-not-excel-compatible-yet.html',
  'mcp-client-setup.html',
] as const

export const productHuntLaunchAssetFiles = [
  'product-hunt-thumbnail.png',
  'product-hunt-gallery-01-workbook-api.png',
  'product-hunt-gallery-02-agent-readback.png',
  'product-hunt-gallery-03-node-service.png',
  'product-hunt-demo.webm',
] as const

export function requireProductHuntLaunchKitDiscovery(
  productHuntLaunchKit: string,
  requireIncludes: (haystack: string, needle: string, context: string) => void,
): void {
  for (const required of productHuntLaunchKitRequiredText) {
    requireIncludes(productHuntLaunchKit, required, 'docs/product-hunt-launch-kit.md')
  }
}

export function requireGrowthSurfaceDiscovery(
  communityLaunchPack: string,
  llms: string,
  productHuntLaunchKit: string,
  requireIncludes: (haystack: string, needle: string, context: string) => void,
): void {
  for (const required of communityLaunchPackRequiredLinks) {
    requireIncludes(communityLaunchPack, required, 'docs/community-launch-pack.md')
  }
  for (const required of llmsExternalSurfaceLinks) {
    requireIncludes(llms, required, 'docs/llms.txt')
  }
  for (const required of [
    'https://proompteng.github.io/bilig/product-hunt-launch-kit.html',
    'https://github.com/proompteng/bilig/blob/main/docs/product-hunt-launch-kit.md',
  ] as const) {
    requireIncludes(llms, required, 'docs/llms.txt')
  }
  requireProductHuntLaunchKitDiscovery(productHuntLaunchKit, requireIncludes)
}
