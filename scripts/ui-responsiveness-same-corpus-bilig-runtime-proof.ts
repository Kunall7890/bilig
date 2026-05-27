import type { Page } from '@playwright/test'

import { hasBiligProductionRuntimeProof } from './ui-responsiveness-same-corpus-guardrails.ts'
import type { SameCorpusBiligRuntimeProof, SameCorpusBiligRuntimeProofSample } from './gen-ui-responsiveness-live-browser-scorecard.ts'

export function buildBiligRuntimeProof(source: string, samples: readonly SameCorpusBiligRuntimeProofSample[]): SameCorpusBiligRuntimeProof {
  const presentSamples = samples.filter((sample) => sample.present)
  const firstPresent = presentSamples[0]
  const actualBuildKind =
    presentSamples.length === 0
      ? 'unknown'
      : presentSamples.every((sample) => sample.buildKind === 'production')
        ? 'production'
        : (firstPresent?.buildKind ?? 'unknown')
  const prod = presentSamples.length > 0 && presentSamples.every((sample) => sample.prod)
  const dev = presentSamples.some((sample) => sample.dev)
  const remoteSyncEnabled = firstPresent?.remoteSyncEnabled ?? null
  const entryRoute = firstPresent?.entryRoute ?? null
  const verified =
    samples.length > 0 && samples.every((sample) => sample.present && sample.buildKind === 'production' && sample.prod && !sample.dev)
  return {
    product: 'bilig',
    source,
    verificationMethod: 'window.__biligRuntimeBuild',
    requiredBuildKind: 'production',
    actualBuildKind,
    mode: firstPresent?.mode ?? 'unknown',
    dev,
    prod,
    remoteSyncEnabled,
    entryRoute,
    sampleCount: samples.length,
    verified,
    samples: samples.map((sample) => ({ ...sample })),
  }
}

export function validateBiligRuntimeProof(proof: SameCorpusBiligRuntimeProof, source: string, caseId: string): void {
  if (proof.product !== 'bilig' || proof.source !== source || proof.verificationMethod !== 'window.__biligRuntimeBuild') {
    throw new Error(`UI responsiveness same-corpus Bilig runtime proof identity mismatch for ${caseId}`)
  }
  if (proof.requiredBuildKind !== 'production') {
    throw new Error(`UI responsiveness same-corpus Bilig runtime proof has stale required build kind for ${caseId}`)
  }
  if (!['development', 'production', 'unknown'].includes(proof.actualBuildKind)) {
    throw new Error(`UI responsiveness same-corpus Bilig runtime proof has invalid build kind for ${caseId}`)
  }
  if (proof.sampleCount !== proof.samples.length || proof.sampleCount <= 0) {
    throw new Error(`UI responsiveness same-corpus Bilig runtime proof sample count is stale for ${caseId}`)
  }
  for (const sample of proof.samples) {
    if (!Number.isInteger(sample.sampleIndex) || sample.sampleIndex < 0) {
      throw new Error(`UI responsiveness same-corpus Bilig runtime proof sample index is invalid for ${caseId}`)
    }
    if (!['development', 'production', 'unknown'].includes(sample.buildKind)) {
      throw new Error(`UI responsiveness same-corpus Bilig runtime proof sample build kind is invalid for ${caseId}`)
    }
  }
  const verified = hasBiligProductionRuntimeProof({ product: 'bilig', biligRuntimeProof: proof })
  if (proof.verified !== verified) {
    throw new Error(`UI responsiveness same-corpus Bilig runtime proof verified flag is stale for ${caseId}`)
  }
}

export async function readBiligRuntimeProofSample(page: Page, sampleIndex: number): Promise<SameCorpusBiligRuntimeProofSample> {
  const runtimeBuild = await page.evaluate(() => {
    const value = (window as Window & { __biligRuntimeBuild?: unknown }).__biligRuntimeBuild
    if (!value || typeof value !== 'object') {
      return null
    }
    const app = Reflect.get(value, 'app')
    const buildKind = Reflect.get(value, 'buildKind')
    const mode = Reflect.get(value, 'mode')
    const dev = Reflect.get(value, 'dev')
    const prod = Reflect.get(value, 'prod')
    const remoteSyncEnabled = Reflect.get(value, 'remoteSyncEnabled')
    const entryRoute = Reflect.get(value, 'entryRoute')
    return {
      app: typeof app === 'string' ? app : null,
      buildKind: buildKind === 'production' || buildKind === 'development' ? buildKind : 'unknown',
      mode: typeof mode === 'string' ? mode : 'unknown',
      dev: dev === true,
      prod: prod === true,
      remoteSyncEnabled: typeof remoteSyncEnabled === 'boolean' ? remoteSyncEnabled : null,
      entryRoute: typeof entryRoute === 'string' ? entryRoute : null,
    }
  })
  if (!runtimeBuild) {
    return {
      sampleIndex,
      present: false,
      app: null,
      buildKind: 'unknown',
      mode: 'unknown',
      dev: false,
      prod: false,
      remoteSyncEnabled: null,
      entryRoute: null,
    }
  }
  return {
    sampleIndex,
    present: runtimeBuild.app === 'bilig-web',
    app: runtimeBuild.app,
    buildKind: runtimeBuild.buildKind,
    mode: runtimeBuild.mode,
    dev: runtimeBuild.dev,
    prod: runtimeBuild.prod,
    remoteSyncEnabled: runtimeBuild.remoteSyncEnabled,
    entryRoute: runtimeBuild.entryRoute,
  }
}
