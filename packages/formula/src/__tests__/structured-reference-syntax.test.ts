import { describe, expect, it } from 'vitest'
import {
  formatStructuredReferenceColumnSpecifier,
  isStructuredReferenceEscapedCharacter,
  parseStructuredReferenceColumnSpecifier,
  scanStructuredReferenceBracket,
} from '../structured-reference-syntax.js'

describe('structured reference syntax helpers', () => {
  it('scans bracketed references with escaped bracket characters', () => {
    expect(scanStructuredReferenceBracket('Sales[[Amount]]', 5)).toEqual({
      content: '[Amount]',
      endIndex: 'Sales[[Amount]]'.length,
    })
    expect(scanStructuredReferenceBracket("Sales[Col']umn]", 5)).toEqual({
      content: "Col']umn",
      endIndex: "Sales[Col']umn]".length,
    })
    expect(scanStructuredReferenceBracket('Sales[Amount', 5)).toBeUndefined()
    expect(scanStructuredReferenceBracket('Sales[Amount]', 0)).toBeUndefined()
  })

  it('formats and parses escaped column specifiers', () => {
    const columnName = "Q1 [Total] # '@"

    expect(formatStructuredReferenceColumnSpecifier(columnName)).toBe("Q1 '[Total'] '# '''@")
    expect(parseStructuredReferenceColumnSpecifier(`[${formatStructuredReferenceColumnSpecifier(columnName)}]`)).toBe(columnName)
    expect(parseStructuredReferenceColumnSpecifier('Amount')).toBe('Amount')
    expect(parseStructuredReferenceColumnSpecifier('[]')).toBe('')
    expect(parseStructuredReferenceColumnSpecifier('   ')).toBeUndefined()
    expect(parseStructuredReferenceColumnSpecifier('[Amount] trailing')).toBeUndefined()
    expect(isStructuredReferenceEscapedCharacter(undefined)).toBe(false)
    expect(isStructuredReferenceEscapedCharacter(']')).toBe(true)
  })
})
