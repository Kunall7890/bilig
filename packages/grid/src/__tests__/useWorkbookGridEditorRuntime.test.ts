// @vitest-environment jsdom
import { describe, expect, it } from 'vitest'
import { shouldPreserveWorkbookTextControlFocus } from '../useWorkbookGridEditorRuntime.js'

describe('useWorkbookGridEditorRuntime focus policy', () => {
  it('does not steal focus from workbook text controls during editor handoff', () => {
    const formulaInput = document.createElement('textarea')
    formulaInput.dataset['workbookTextControl'] = 'true'
    const nameBox = document.createElement('input')
    nameBox.dataset['workbookTextControl'] = 'true'
    const cellEditor = document.createElement('textarea')
    cellEditor.dataset['workbookTextControl'] = 'true'
    const toolbarInput = document.createElement('input')
    toolbarInput.dataset['workbookTextControl'] = 'false'
    const legacyTestIdOnlyInput = document.createElement('input')
    legacyTestIdOnlyInput.dataset['testid'] = 'formula-input'

    expect(shouldPreserveWorkbookTextControlFocus(formulaInput)).toBe(true)
    expect(shouldPreserveWorkbookTextControlFocus(nameBox)).toBe(true)
    expect(shouldPreserveWorkbookTextControlFocus(cellEditor)).toBe(true)
    expect(shouldPreserveWorkbookTextControlFocus(toolbarInput)).toBe(false)
    expect(shouldPreserveWorkbookTextControlFocus(legacyTestIdOnlyInput)).toBe(false)
    expect(shouldPreserveWorkbookTextControlFocus(formulaInput, { force: true })).toBe(false)
  })
})
