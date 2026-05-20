import type { WorkbookSnapshot } from '@bilig/protocol'

const importedXlsxSourceBytes = Symbol.for('bilig.importedXlsxSourceBytes')

type SnapshotWithImportedXlsxSource = WorkbookSnapshot & {
  readonly [importedXlsxSourceBytes]?: Uint8Array
}

export function attachImportedXlsxSourceBytes(snapshot: WorkbookSnapshot, bytes: Uint8Array): WorkbookSnapshot {
  Object.defineProperty(snapshot, importedXlsxSourceBytes, {
    configurable: true,
    enumerable: false,
    value: bytes,
  })
  return snapshot
}

export function readImportedXlsxSourceBytes(snapshot: WorkbookSnapshot): Uint8Array | undefined {
  return (snapshot as SnapshotWithImportedXlsxSource)[importedXlsxSourceBytes]
}
