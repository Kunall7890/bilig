import type { Root } from 'react-dom/client'

const HOT_ROOT_KEY = '__biligReactRoot'
const HOT_ROOT_CONTAINER_KEY = '__biligReactRootContainer'

interface ReactRootHotData {
  [HOT_ROOT_KEY]?: Root | undefined
  [HOT_ROOT_CONTAINER_KEY]?: Element | undefined
}

export interface ReactRootHotContext {
  readonly data: ReactRootHotData
  dispose(callback: (data: ReactRootHotData) => void): void
}

export function getOrCreateReactRoot(input: {
  readonly container: Element
  readonly createRoot: (container: Element) => Root
  readonly hot?: ReactRootHotContext | null | undefined
}): Root {
  const hot = input.hot
  const hotData = hot?.data
  const existingRoot = hotData?.[HOT_ROOT_KEY]
  const existingContainer = hotData?.[HOT_ROOT_CONTAINER_KEY]
  if (existingRoot && existingContainer === input.container) {
    return existingRoot
  }

  if (existingRoot && existingContainer && existingContainer !== input.container) {
    existingRoot.unmount()
  }

  const root = input.createRoot(input.container)
  if (!hot) {
    return root
  }

  const writableHotData = hot.data
  writableHotData[HOT_ROOT_KEY] = root
  writableHotData[HOT_ROOT_CONTAINER_KEY] = input.container
  hot.dispose((data) => {
    data[HOT_ROOT_KEY] = root
    data[HOT_ROOT_CONTAINER_KEY] = input.container
  })
  return root
}
