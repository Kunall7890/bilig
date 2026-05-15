import { useEffect } from 'react'
import { isTextEntryTarget } from './worker-workbook-app-model.js'

type SheetShortcutDirection = -1 | 1

function resolveSheetShortcutDirection(event: KeyboardEvent): SheetShortcutDirection | null {
  if (!event.altKey || event.ctrlKey || event.metaKey || event.shiftKey) {
    return null
  }
  if (event.key === 'ArrowDown' || event.key === 'ArrowRight') {
    return 1
  }
  if (event.key === 'ArrowUp' || event.key === 'ArrowLeft') {
    return -1
  }
  return null
}

function shouldRouteSheetShortcutFromActiveElement(activeElement: Element | null): boolean {
  if (isTextEntryTarget(activeElement)) {
    return false
  }

  const workbookScope = document.querySelector('[data-workbook-keyboard-scope="true"]')
  if (!activeElement || activeElement === document.body || activeElement === document.documentElement) {
    return true
  }

  return workbookScope?.contains(activeElement) ?? false
}

export function resolveAdjacentSheetName(
  sheetNames: readonly string[],
  currentSheetName: string,
  direction: SheetShortcutDirection,
): string | null {
  const currentIndex = sheetNames.indexOf(currentSheetName)
  if (currentIndex < 0) {
    return null
  }
  return sheetNames[currentIndex + direction] ?? null
}

export function useWorkbookSheetKeyboardShortcuts(input: {
  readonly address: string
  readonly enabled: boolean
  readonly onSelectSheet: (sheetName: string, address: string) => void
  readonly sheetName: string
  readonly sheetNames: readonly string[]
}) {
  const { address, enabled, onSelectSheet, sheetName, sheetNames } = input

  useEffect(() => {
    if (!enabled || sheetNames.length < 2) {
      return
    }

    const handleWindowKeyDown = (event: KeyboardEvent) => {
      if (event.defaultPrevented || !shouldRouteSheetShortcutFromActiveElement(document.activeElement)) {
        return
      }

      const direction = resolveSheetShortcutDirection(event)
      if (direction === null) {
        return
      }

      const nextSheetName = resolveAdjacentSheetName(sheetNames, sheetName, direction)
      if (!nextSheetName) {
        return
      }

      event.preventDefault()
      event.stopPropagation()
      onSelectSheet(nextSheetName, address)
    }

    window.addEventListener('keydown', handleWindowKeyDown, true)
    return () => {
      window.removeEventListener('keydown', handleWindowKeyDown, true)
    }
  }, [address, enabled, onSelectSheet, sheetName, sheetNames])
}
