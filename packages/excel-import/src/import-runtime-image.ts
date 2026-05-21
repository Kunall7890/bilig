const runtimeImageSymbol = Symbol.for('bilig.runtimeImage')

export function attachImportedRuntimeImage<T extends object>(carrier: T, runtimeImage: unknown): T {
  Object.defineProperty(carrier, runtimeImageSymbol, {
    value: runtimeImage,
    configurable: true,
    enumerable: false,
    writable: true,
  })
  return carrier
}
