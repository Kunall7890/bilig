// @vitest-environment jsdom
import { act } from 'react'
import { createRoot } from 'react-dom/client'
import { afterEach, describe, expect, it, vi } from 'vitest'
import { WorkbookAppErrorBoundary } from '../WorkbookAppErrorBoundary.js'

function ThrowingWorkbook(): never {
  throw new Error('route transition hook crash')
}

function HealthyWorkbook() {
  return <div data-testid="healthy-workbook">Grid ready</div>
}

describe('WorkbookAppErrorBoundary', () => {
  afterEach(() => {
    document.body.innerHTML = ''
    vi.restoreAllMocks()
  })

  it('renders a recoverable workbook shell error instead of a blank viewport', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    const onReload = vi.fn()
    const onError = vi.fn()
    vi.spyOn(console, 'error').mockImplementation(() => undefined)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <WorkbookAppErrorBoundary onError={onError} onReload={onReload} resetKey="doc-1">
          <ThrowingWorkbook />
        </WorkbookAppErrorBoundary>,
      )
    })

    expect(host.querySelector("[data-testid='workbook-app-error-state']")?.textContent).toContain('Workbook render failed')
    expect(host.querySelector("[data-testid='workbook-app-error-message']")?.textContent).toContain('route transition hook crash')
    expect(onError).toHaveBeenCalledTimes(1)

    const reloadButton = host.querySelector("[data-testid='workbook-app-error-reload']")
    expect(reloadButton).toBeInstanceOf(HTMLButtonElement)
    if (!(reloadButton instanceof HTMLButtonElement)) {
      throw new Error('Expected workbook error reload button')
    }

    await act(async () => {
      reloadButton.click()
    })

    expect(onReload).toHaveBeenCalledTimes(1)

    await act(async () => {
      root.unmount()
    })
  })

  it('resets the failure state when the workbook runtime key changes', async () => {
    ;(globalThis as typeof globalThis & { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true
    vi.spyOn(console, 'error').mockImplementation(() => undefined)

    const host = document.createElement('div')
    document.body.appendChild(host)
    const root = createRoot(host)

    await act(async () => {
      root.render(
        <WorkbookAppErrorBoundary resetKey="doc-1">
          <ThrowingWorkbook />
        </WorkbookAppErrorBoundary>,
      )
    })

    expect(host.querySelector("[data-testid='workbook-app-error-state']")).not.toBeNull()

    await act(async () => {
      root.render(
        <WorkbookAppErrorBoundary resetKey="doc-2">
          <HealthyWorkbook />
        </WorkbookAppErrorBoundary>,
      )
    })

    expect(host.querySelector("[data-testid='workbook-app-error-state']")).toBeNull()
    expect(host.querySelector("[data-testid='healthy-workbook']")?.textContent).toBe('Grid ready')

    await act(async () => {
      root.unmount()
    })
  })
})
