import type { LiteralInput, WorkbookSnapshot } from '@bilig/protocol'

const importedXlsxSourceBytes = Symbol.for('bilig.importedXlsxSourceBytes')
const importedXlsxSourceCellPatches = Symbol.for('bilig.importedXlsxSourceCellPatches')

export type ImportedXlsxSourceReader = {
  readonly byteLength: number
  readBytes(): Uint8Array
}

export type ImportedXlsxSourceReference = Uint8Array | ImportedXlsxSourceReader

export interface ImportedXlsxSourceCellPatch {
  readonly kind: 'literal'
  readonly sheetName: string
  readonly address: string
  readonly value: LiteralInput
}

type SnapshotWithImportedXlsxSource = WorkbookSnapshot & {
  readonly [importedXlsxSourceBytes]?: ImportedXlsxSourceReference
  readonly [importedXlsxSourceCellPatches]?: readonly ImportedXlsxSourceCellPatch[]
}

export function readImportedXlsxSource(snapshot: WorkbookSnapshot): ImportedXlsxSourceReference | undefined {
  return (snapshot as SnapshotWithImportedXlsxSource)[importedXlsxSourceBytes]
}

export function readImportedXlsxSourceCellPatches(snapshot: WorkbookSnapshot): readonly ImportedXlsxSourceCellPatch[] {
  return (snapshot as SnapshotWithImportedXlsxSource)[importedXlsxSourceCellPatches] ?? []
}

export function attachImportedXlsxSourceMetadata(
  snapshot: WorkbookSnapshot,
  source: ImportedXlsxSourceReference | undefined,
  patches: readonly ImportedXlsxSourceCellPatch[] = [],
): WorkbookSnapshot {
  if (source !== undefined) {
    Object.defineProperty(snapshot, importedXlsxSourceBytes, {
      configurable: true,
      enumerable: false,
      value: source,
    })
  }
  if (patches.length > 0) {
    Object.defineProperty(snapshot, importedXlsxSourceCellPatches, {
      configurable: true,
      enumerable: false,
      value: patches,
    })
  }
  return snapshot
}

export function canRecordImportedXlsxLiteralPatch(
  source: ImportedXlsxSourceReference | undefined,
  content: unknown,
): content is LiteralInput {
  return source !== undefined && scalarContentCanPatchImportedXlsxSource(content)
}

export function setImportedXlsxLiteralPatch(
  patches: Map<string, ImportedXlsxSourceCellPatch>,
  sheetName: string,
  address: string,
  value: LiteralInput,
): void {
  patches.set(`${sheetName}!${address}`, {
    kind: 'literal',
    sheetName,
    address,
    value,
  })
}

export function scalarContentCanPatchImportedXlsxSource(content: unknown): content is LiteralInput {
  if (content === null || typeof content === 'number' || typeof content === 'boolean') {
    return true
  }
  return typeof content === 'string' && !content.startsWith('=')
}
