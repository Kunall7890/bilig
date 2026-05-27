import { ErrorCode, ValueTag, type CellValue } from '@bilig/protocol'
import type { TextFixtureGroup } from './text-fixtures.js'

export const SUBSTITUTE_TEXT_FIXTURES: TextFixtureGroup = {
  builtin: 'SUBSTITUTE',
  cases: [
    {
      name: 'replaces all matching text',
      args: [text('banana'), text('an'), text('oo')],
      expected: text('booooa'),
    },
    {
      name: 'replaces the requested instance',
      args: [text('banana'), text('an'), text('oo'), number(2)],
      expected: text('banooa'),
    },
    {
      name: 'empty old text returns original text',
      args: [text('abc'), text(''), text('z')],
      expected: text('abc'),
      note: 'TrueCalc and xlsx-calc agree that empty old_text leaves the text unchanged.',
    },
    {
      name: 'empty old text returns original text with instance',
      args: [text('abc'), text(''), text('z'), number(1)],
      expected: text('abc'),
      note: 'TrueCalc and xlsx-calc agree that empty old_text leaves the text unchanged.',
    },
    {
      name: 'rejects zero instance number',
      args: [text('abc'), text('a'), text('z'), number(0)],
      expected: textValueError(),
    },
  ],
}

function number(value: number): CellValue {
  return { tag: ValueTag.Number, value }
}

function text(value: string): CellValue {
  return { tag: ValueTag.String, value, stringId: 0 }
}

function textValueError(): CellValue {
  return { tag: ValueTag.Error, code: ErrorCode.Value }
}
