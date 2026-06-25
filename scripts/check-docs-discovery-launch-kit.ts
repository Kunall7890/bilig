import { communityLaunchPackRequiredLinks } from './check-docs-discovery-growth-links.ts'
import { getBenchmarkDiscoveryEvidence } from './check-docs-discovery-benchmark-evidence.ts'

const productHuntLaunchKitRequiredText = [
  'title: Bilig product surface notes',
  'WorkPaper formulas for TypeScript services.',
  'without canned launch copy',
  'Do not use canned comments',
  'https://www.npmjs.com/package/@bilig/workpaper',
  'eval-workpaper-service.html',
  'eval-agent-mcp.html',
  'product-hunt-thumbnail.png',
  'product-hunt-gallery-01-workbook-api.png',
  'product-hunt-gallery-02-agent-readback.png',
  'product-hunt-gallery-03-node-service.png',
  'product-hunt-demo.webm',
  'what-workpaper-benchmark-proves.html',
  'where-bilig-is-not-excel-compatible-yet.html',
  'mcp-client-setup.html',
  'Fit Check',
  'https://www.producthunt.com/launch/preparing-for-launch',
  'https://www.producthunt.com/launch/',
  'personal maker account',
  'midnight PST',
  'Do not ask for\n  upvotes.',
  '240x240',
  '1270x760',
  'YouTube link',
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
  const benchmarkEvidence = getBenchmarkDiscoveryEvidence()
  const requiredText = [
    ...productHuntLaunchKitRequiredText,
    'The public package is `@bilig/workpaper`.',
    'returns `verified: true` after edit, recalculation, save, and\n  restore.',
    'the narrow comparable-row claim',
    benchmarkEvidence.p95HoldoutWorkload,
  ] as const

  for (const required of requiredText) {
    requireIncludes(productHuntLaunchKit, required, 'internal/growth/product-hunt-launch-kit.md')
  }
}

export function requireGrowthSurfaceDiscovery(
  communityLaunchPack: string,
  _headlessPackageVersion: string,
  _llms: string,
  productHuntLaunchKit: string,
  requireIncludes: (haystack: string, needle: string, context: string) => void,
): void {
  for (const required of communityLaunchPackRequiredLinks()) {
    requireIncludes(communityLaunchPack, required, 'internal/growth/community-launch-pack.md')
  }
  requireProductHuntLaunchKitDiscovery(productHuntLaunchKit, requireIncludes)
}
