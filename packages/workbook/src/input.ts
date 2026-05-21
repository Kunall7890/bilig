export type WorkbookActionInput =
  | null
  | string
  | number
  | boolean
  | readonly WorkbookActionInput[]
  | { readonly [key: string]: WorkbookActionInput }

export class WorkbookActionInputError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'WorkbookActionInputError'
  }
}

function typeName(value: object): string {
  return value.constructor?.name ?? 'object'
}

function childPath(path: string, key: string): string {
  return /^[A-Za-z_$][A-Za-z0-9_$]*$/.test(key) ? `${path}.${key}` : `${path}[${JSON.stringify(key)}]`
}

function normalizeInput(value: unknown, path: string, seen: WeakSet<object>): WorkbookActionInput {
  if (value === null || typeof value === 'string' || typeof value === 'boolean') {
    return value
  }

  if (typeof value === 'number') {
    if (!Number.isFinite(value)) {
      throw new WorkbookActionInputError(`Action input at ${path} must be a finite number`)
    }
    return value
  }

  if (value === undefined) {
    throw new WorkbookActionInputError(`Action input at ${path} must not be undefined`)
  }

  if (typeof value === 'bigint' || typeof value === 'function' || typeof value === 'symbol') {
    throw new WorkbookActionInputError(`Action input at ${path} must be JSON-safe, not ${typeof value}`)
  }

  if (typeof value !== 'object') {
    throw new WorkbookActionInputError(`Action input at ${path} must be JSON-safe`)
  }

  if (seen.has(value)) {
    throw new WorkbookActionInputError(`Action input at ${path} must not contain cycles`)
  }
  seen.add(value)

  try {
    if (Array.isArray(value)) {
      const output: WorkbookActionInput[] = []
      for (let index = 0; index < value.length; index += 1) {
        if (!Object.hasOwn(value, index)) {
          throw new WorkbookActionInputError(`Action input at ${path}[${index}] must not be a sparse array hole`)
        }
        output.push(normalizeInput(value[index], `${path}[${index}]`, seen))
      }
      return output
    }

    const prototype = Object.getPrototypeOf(value)
    if (prototype !== Object.prototype && prototype !== null) {
      throw new WorkbookActionInputError(`Action input at ${path} must be a plain JSON object, not ${typeName(value)}`)
    }

    if (Object.getOwnPropertySymbols(value).length > 0) {
      throw new WorkbookActionInputError(`Action input at ${path} must not contain symbol keys`)
    }

    const output: { [key: string]: WorkbookActionInput } = {}
    Object.entries(value)
      .toSorted(([left], [right]) => left.localeCompare(right))
      .forEach(([key, entry]) => {
        output[key] = normalizeInput(entry, childPath(path, key), seen)
      })
    return output
  } finally {
    seen.delete(value)
  }
}

export function normalizeWorkbookActionInput(input: unknown): WorkbookActionInput {
  return normalizeInput(input, 'input', new WeakSet())
}

export function normalizeOptionalWorkbookActionInput(input: WorkbookActionInput | undefined): WorkbookActionInput | undefined {
  return input === undefined ? undefined : normalizeWorkbookActionInput(input)
}

export function hasOwnActionInput(value: object): boolean {
  return Object.prototype.hasOwnProperty.call(value, 'input')
}

export function getOwnActionInput(value: object): unknown {
  return Object.getOwnPropertyDescriptor(value, 'input')?.value
}
