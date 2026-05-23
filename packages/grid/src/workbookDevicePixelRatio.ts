export function getWorkbookDevicePixelRatio(): number {
  return typeof window === 'undefined' ? 1 : Math.max(1, window.devicePixelRatio || 1)
}

export function subscribeWorkbookDevicePixelRatioChange(onStoreChange: () => void): () => void {
  if (typeof window === 'undefined') {
    return () => {}
  }

  let disposed = false
  let resolutionQuery: MediaQueryList | null = null
  const handleChange = () => {
    if (disposed) {
      return
    }
    onStoreChange()
    resetResolutionQuery()
  }
  const resetResolutionQuery = () => {
    if (disposed) {
      return
    }
    removeResolutionQueryListener(resolutionQuery, handleChange)
    resolutionQuery =
      typeof window.matchMedia === 'function' ? window.matchMedia(`(resolution: ${getWorkbookDevicePixelRatio()}dppx)`) : null
    addResolutionQueryListener(resolutionQuery, handleChange)
  }

  resetResolutionQuery()
  window.addEventListener('resize', handleChange)
  window.visualViewport?.addEventListener('resize', handleChange)
  return () => {
    disposed = true
    removeResolutionQueryListener(resolutionQuery, handleChange)
    resolutionQuery = null
    window.removeEventListener('resize', handleChange)
    window.visualViewport?.removeEventListener('resize', handleChange)
  }
}

function addResolutionQueryListener(query: MediaQueryList | null, listener: () => void): void {
  if (!query) {
    return
  }
  if (typeof query.addEventListener === 'function') {
    query.addEventListener('change', listener)
    return
  }
  const legacyQuery = query as MediaQueryList & { addListener?: (listener: () => void) => void }
  legacyQuery.addListener?.(listener)
}

function removeResolutionQueryListener(query: MediaQueryList | null, listener: () => void): void {
  if (!query) {
    return
  }
  if (typeof query.removeEventListener === 'function') {
    query.removeEventListener('change', listener)
    return
  }
  const legacyQuery = query as MediaQueryList & { removeListener?: (listener: () => void) => void }
  legacyQuery.removeListener?.(listener)
}
