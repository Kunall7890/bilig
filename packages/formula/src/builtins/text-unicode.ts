export function excelTextLength(text: string): number {
  return Array.from(text).length
}

export function excelTextIndexToCodeUnitIndex(text: string, textIndex: number): number {
  if (textIndex <= 0) {
    return 0
  }

  let codeUnitIndex = 0
  let currentTextIndex = 0
  for (const char of text) {
    if (currentTextIndex >= textIndex) {
      return codeUnitIndex
    }
    codeUnitIndex += char.length
    currentTextIndex += 1
  }
  return text.length
}

export function excelTextPositionFromCodeUnitIndex(text: string, codeUnitIndex: number): number {
  let currentCodeUnitIndex = 0
  let textIndex = 0
  for (const char of text) {
    if (currentCodeUnitIndex >= codeUnitIndex) {
      return textIndex + 1
    }
    currentCodeUnitIndex += char.length
    textIndex += 1
  }
  return textIndex + 1
}

export function excelTextSlice(text: string, startIndex: number, count: number): string {
  if (count <= 0) {
    return ''
  }

  const startCodeUnitIndex = excelTextIndexToCodeUnitIndex(text, startIndex)
  const endCodeUnitIndex = excelTextIndexToCodeUnitIndex(text, startIndex + count)
  return text.slice(startCodeUnitIndex, endCodeUnitIndex)
}

export function excelTextReplace(text: string, start: number, count: number, replacement: string): string {
  const startIndex = start - 1
  if (startIndex >= excelTextLength(text)) {
    return text + replacement
  }

  const startCodeUnitIndex = excelTextIndexToCodeUnitIndex(text, startIndex)
  const endCodeUnitIndex = excelTextIndexToCodeUnitIndex(text, startIndex + count)
  return text.slice(0, startCodeUnitIndex) + replacement + text.slice(endCodeUnitIndex)
}
