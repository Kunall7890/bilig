#!/usr/bin/env bun

import { existsSync, readdirSync, readFileSync, statSync } from 'node:fs'
import { basename, join, relative } from 'node:path'
import { pathToFileURL } from 'node:url'

import { outputPath, rootDir } from './bilig-dominance-scorecard-input.ts'

export interface PublicClaimViolation {
  readonly path: string
  readonly line: number
  readonly column: number
  readonly match: string
  readonly text: string
}

export interface PublicClaimCheckReport {
  readonly blanketTenXClaimAllowed: boolean
  readonly scannedFiles: readonly string[]
  readonly violations: readonly PublicClaimViolation[]
}

interface ScorecardClaimPolicy {
  readonly blanketTenXClaimAllowed: boolean
}

interface ScorecardLike {
  readonly claimPolicy: ScorecardClaimPolicy
}

const INTERNAL_DOC_NAME_PATTERNS: readonly RegExp[] = [
  /(?:^|-)plan(?:-|$)/i,
  /(?:^|-)program(?:-|$)/i,
  /(?:^|-)oracle(?:-|$)/i,
  /(?:^|-)design(?:-|$)/i,
  /(?:^|-)phase(?:-|$)/i,
  /(?:^|-)implementation(?:-|$)/i,
  /(?:^|-)remediation(?:-|$)/i,
]

const PUBLIC_DOC_EXTENSIONS = new Set(['.html', '.md'])

const BROAD_GOOGLE_SHEETS_TEN_X_PATTERNS: readonly RegExp[] = [
  /\b10\s*x\s+(?:faster|better|superior)\s+than\s+(?:(?:Excel|Microsoft Excel)(?:\s+Web)?\s+(?:and|or)\s+)?Google Sheets\b/gi,
  /\b10\s*x\s+(?:faster|better|superior)\s+than\s+(?:Sheets|Google's Sheets)\b/gi,
  /\b(?:beat|beats|beating|outperform|outperforms|outperforming)\s+Google Sheets\b[^\n.?!;:]{0,120}\b10\s*x\b/gi,
  /\b10\s*x\b[^\n.?!;:]{0,120}\b(?:beat|beats|beating|outperform|outperforms|outperforming)\s+Google Sheets\b/gi,
  /\b10\s*x\s+Google Sheets\b/gi,
]

export function collectPublicClaimFiles(repoRoot = rootDir): string[] {
  const files = new Set<string>()
  addIfFile(files, repoRoot, 'README.md')
  addIfFile(files, repoRoot, join('packages', 'headless', 'README.md'))
  collectDocs(files, repoRoot, join(repoRoot, 'docs'))
  return [...files].toSorted()
}

export function findBroadGoogleSheetsTenXClaims(source: string, repoPath: string): PublicClaimViolation[] {
  const violations: PublicClaimViolation[] = []
  const lines = source.split(/\r?\n/)
  for (const [lineIndex, line] of lines.entries()) {
    for (const pattern of BROAD_GOOGLE_SHEETS_TEN_X_PATTERNS) {
      pattern.lastIndex = 0
      for (let match = pattern.exec(line); match !== null; match = pattern.exec(line)) {
        violations.push({
          path: repoPath,
          line: lineIndex + 1,
          column: match.index + 1,
          match: match[0],
          text: line.trim(),
        })
      }
    }
  }
  return violations
}

export function buildPublicClaimCheckReport(
  input: {
    readonly repoRoot?: string | undefined
    readonly scorecard?: ScorecardLike | undefined
    readonly files?: readonly string[] | undefined
  } = {},
): PublicClaimCheckReport {
  const repoRoot = input.repoRoot ?? rootDir
  const scorecard = input.scorecard ?? loadScorecard(outputPath)
  const scannedFiles = input.files ?? collectPublicClaimFiles(repoRoot)
  const violations = scorecard.claimPolicy.blanketTenXClaimAllowed
    ? []
    : scannedFiles.flatMap((repoPath) => findBroadGoogleSheetsTenXClaims(readFileSync(join(repoRoot, repoPath), 'utf8'), repoPath))
  return {
    blanketTenXClaimAllowed: scorecard.claimPolicy.blanketTenXClaimAllowed,
    scannedFiles,
    violations,
  }
}

function main(): void {
  const report = buildPublicClaimCheckReport()
  if (report.violations.length > 0) {
    const formattedViolations = report.violations
      .map((violation) => `${violation.path}:${String(violation.line)}:${String(violation.column)} ${violation.match}`)
      .join('\n')
    throw new Error(
      [
        'Public claim check failed: broad Google Sheets 10x wording is forbidden until dominance scorecard claimPolicy.blanketTenXClaimAllowed is true.',
        formattedViolations,
      ].join('\n'),
    )
  }

  console.log(
    JSON.stringify(
      {
        scannedFileCount: report.scannedFiles.length,
        blanketTenXClaimAllowed: report.blanketTenXClaimAllowed,
        violationCount: report.violations.length,
      },
      null,
      2,
    ),
  )
}

function addIfFile(files: Set<string>, repoRoot: string, repoPath: string): void {
  if (existsSync(join(repoRoot, repoPath))) {
    files.add(repoPath)
  }
}

function collectDocs(files: Set<string>, repoRoot: string, dir: string): void {
  if (!existsSync(dir)) {
    return
  }
  for (const entry of readdirSync(dir)) {
    const path = join(dir, entry)
    const stat = statSync(path)
    if (stat.isDirectory()) {
      collectDocs(files, repoRoot, path)
      continue
    }
    if (!stat.isFile() || !isPublicDocFile(path)) {
      continue
    }
    const repoPath = relative(repoRoot, path)
    if (!isInternalPlanningDoc(repoPath)) {
      files.add(repoPath)
    }
  }
}

function isPublicDocFile(path: string): boolean {
  return [...PUBLIC_DOC_EXTENSIONS].some((extension) => path.endsWith(extension))
}

function isInternalPlanningDoc(repoPath: string): boolean {
  const name = basename(repoPath)
  return INTERNAL_DOC_NAME_PATTERNS.some((pattern) => pattern.test(name))
}

function loadScorecard(path: string): ScorecardLike {
  const parsed: unknown = JSON.parse(readFileSync(path, 'utf8'))
  if (!isRecord(parsed)) {
    throw new Error(`Dominance scorecard at ${path} must be an object`)
  }
  const claimPolicy = parsed['claimPolicy']
  if (!isRecord(claimPolicy) || typeof claimPolicy['blanketTenXClaimAllowed'] !== 'boolean') {
    throw new Error(`Dominance scorecard at ${path} is missing claimPolicy.blanketTenXClaimAllowed`)
  }
  return {
    claimPolicy: {
      blanketTenXClaimAllowed: claimPolicy['blanketTenXClaimAllowed'],
    },
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value)
}

if (process.argv[1] && pathToFileURL(process.argv[1]).href === import.meta.url) {
  main()
}
