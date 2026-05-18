// @vitest-environment jsdom
import { createRoot as createReactRoot } from 'react-dom/client'
import { describe, expect, it, vi } from 'vitest'
import type { Root } from 'react-dom/client'
import { getOrCreateReactRoot, type ReactRootHotContext } from '../react-root-lifecycle.js'

function createHotContext(): ReactRootHotContext & { callbacks: ((data: ReactRootHotContext['data']) => void)[] } {
  const callbacks: ((data: ReactRootHotContext['data']) => void)[] = []
  return {
    callbacks,
    data: {},
    dispose(callback) {
      callbacks.push(callback)
    },
  }
}

function createRootStub(): Root {
  return createReactRoot(document.createElement('div'))
}

function createContainerStub(id: string): Element {
  const container = document.createElement('div')
  container.dataset.testId = id
  return container
}

describe('react root lifecycle', () => {
  it('reuses the existing Vite hot root for the same container', () => {
    const container = createContainerStub('root')
    const hot = createHotContext()
    const createdRoot = createRootStub()
    const createRoot = vi.fn(() => createdRoot)

    const first = getOrCreateReactRoot({ container, createRoot, hot })
    const second = getOrCreateReactRoot({ container, createRoot, hot })

    expect(first).toBe(createdRoot)
    expect(second).toBe(createdRoot)
    expect(createRoot).toHaveBeenCalledTimes(1)
  })

  it('carries the active root through hot dispose data instead of unmounting it', () => {
    const container = createContainerStub('root')
    const hot = createHotContext()
    const root = createRootStub()
    const unmount = vi.spyOn(root, 'unmount')

    getOrCreateReactRoot({ container, createRoot: () => root, hot })
    hot.callbacks.forEach((callback) => callback(hot.data))

    const createRootAfterDispose = vi.fn(() => createRootStub())
    const reused = getOrCreateReactRoot({ container, createRoot: createRootAfterDispose, hot })

    expect(reused).toBe(root)
    expect(unmount).not.toHaveBeenCalled()
    expect(createRootAfterDispose).not.toHaveBeenCalled()
  })

  it('unmounts a stale hot root if the DOM container changes', () => {
    const firstContainer = createContainerStub('first')
    const secondContainer = createContainerStub('second')
    const hot = createHotContext()
    const firstRoot = createRootStub()
    const secondRoot = createRootStub()
    const firstUnmount = vi.spyOn(firstRoot, 'unmount')
    const createRoot = vi.fn().mockReturnValueOnce(firstRoot).mockReturnValueOnce(secondRoot)

    getOrCreateReactRoot({ container: firstContainer, createRoot, hot })
    const resolved = getOrCreateReactRoot({ container: secondContainer, createRoot, hot })

    expect(resolved).toBe(secondRoot)
    expect(firstUnmount).toHaveBeenCalledTimes(1)
    expect(createRoot).toHaveBeenCalledTimes(2)
  })
})
