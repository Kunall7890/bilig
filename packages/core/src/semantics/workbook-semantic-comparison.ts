import type { WorkbookSnapshot } from '@bilig/protocol'
import { projectWorkbookSemanticSnapshot, type WorkbookSemanticSnapshot } from './workbook-semantic-projection.js'

export interface WorkbookSemanticSnapshotDiff {
  readonly path: string
  readonly left: unknown
  readonly right: unknown
}

export function stableStringifyWorkbookSemanticSnapshot(snapshot: WorkbookSemanticSnapshot): string {
  return JSON.stringify(snapshot)
}

export function workbookSemanticSnapshotsEqual(left: WorkbookSnapshot, right: WorkbookSnapshot): boolean {
  return (
    stableStringifyWorkbookSemanticSnapshot(projectWorkbookSemanticSnapshot(left)) ===
    stableStringifyWorkbookSemanticSnapshot(projectWorkbookSemanticSnapshot(right))
  )
}

export function diffWorkbookSemanticSnapshots(left: WorkbookSnapshot, right: WorkbookSnapshot): WorkbookSemanticSnapshotDiff[] {
  return diffSemanticValues(projectWorkbookSemanticSnapshot(left), projectWorkbookSemanticSnapshot(right))
}

function diffSemanticValues(left: unknown, right: unknown, path: readonly string[] = []): WorkbookSemanticSnapshotDiff[] {
  if (Object.is(left, right)) {
    return []
  }
  if (Array.isArray(left) || Array.isArray(right)) {
    if (!Array.isArray(left) || !Array.isArray(right)) {
      return [semanticDiff(path, left, right)]
    }
    const length = Math.max(left.length, right.length)
    const diffs: WorkbookSemanticSnapshotDiff[] = []
    for (let index = 0; index < length; index += 1) {
      diffs.push(...diffSemanticValues(left[index], right[index], [...path, String(index)]))
    }
    return diffs
  }
  if (isRecord(left) || isRecord(right)) {
    if (!isRecord(left) || !isRecord(right)) {
      return [semanticDiff(path, left, right)]
    }
    const keys = [...new Set([...Object.keys(left), ...Object.keys(right)])].toSorted()
    return keys.flatMap((key) => diffSemanticValues(left[key], right[key], [...path, key]))
  }
  return [semanticDiff(path, left, right)]
}

function semanticDiff(path: readonly string[], left: unknown, right: unknown): WorkbookSemanticSnapshotDiff {
  return {
    path: path.length === 0 ? '$' : `$.${path.join('.')}`,
    left,
    right,
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value)
}
