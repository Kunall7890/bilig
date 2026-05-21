const STRUCTURED_REFERENCE_ESCAPED_CHARACTERS = new Set(['[', ']', '#', "'", '@'])

export function isStructuredReferenceEscapedCharacter(character: string | undefined): boolean {
  return character !== undefined && STRUCTURED_REFERENCE_ESCAPED_CHARACTERS.has(character)
}

export function scanStructuredReferenceBracket(
  source: string,
  startIndex: number,
): { readonly content: string; readonly endIndex: number } | undefined {
  if (source[startIndex] !== '[') {
    return undefined
  }

  let depth = 0
  for (let index = startIndex; index < source.length; index += 1) {
    const character = source[index]
    if (character === "'" && isStructuredReferenceEscapedCharacter(source[index + 1])) {
      index += 1
      continue
    }
    if (character === '[') {
      depth += 1
      continue
    }
    if (character === ']') {
      depth -= 1
      if (depth === 0) {
        return {
          content: source.slice(startIndex + 1, index),
          endIndex: index + 1,
        }
      }
      if (depth < 0) {
        return undefined
      }
    }
  }
  return undefined
}

export function parseStructuredReferenceColumnSpecifier(source: string): string | undefined {
  const trimmed = source.trim()
  if (trimmed.length === 0) {
    return undefined
  }

  const columnText = unwrapSingleStructuredReferenceColumnItem(trimmed)
  return columnText === undefined ? undefined : unescapeStructuredReferenceColumnSpecifier(columnText)
}

export function formatStructuredReferenceColumnSpecifier(columnName: string): string {
  let output = ''
  for (const character of columnName) {
    output += isStructuredReferenceEscapedCharacter(character) ? `'${character}` : character
  }
  return output
}

function unwrapSingleStructuredReferenceColumnItem(source: string): string | undefined {
  if (!source.startsWith('[')) {
    return source
  }
  if (!source.endsWith(']')) {
    return undefined
  }
  const scanned = scanStructuredReferenceBracket(source, 0)
  if (scanned?.endIndex !== source.length) {
    return undefined
  }
  return scanned.content.trim()
}

function unescapeStructuredReferenceColumnSpecifier(source: string): string {
  let output = ''
  for (let index = 0; index < source.length; index += 1) {
    const character = source[index]!
    const next = source[index + 1]
    if (character === "'" && isStructuredReferenceEscapedCharacter(next)) {
      output += next
      index += 1
      continue
    }
    output += character
  }
  return output
}
