export function createRequire(): (path: string) => never {
  return (path: string) => {
    throw new Error(`Node require is unavailable in the browser bundle: ${path}`)
  }
}
