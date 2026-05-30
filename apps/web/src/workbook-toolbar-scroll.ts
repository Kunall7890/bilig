import { useEffect, useRef, useState } from 'react'

interface ToolbarScrollCueState {
  readonly hasOverflow: boolean
  readonly isAtStart: boolean
  readonly isAtEnd: boolean
}

export function useToolbarScrollCue() {
  const scrollContainerRef = useRef<HTMLDivElement | null>(null)
  const [state, setState] = useState<ToolbarScrollCueState>({
    hasOverflow: false,
    isAtStart: true,
    isAtEnd: true,
  })

  useEffect(() => {
    const scrollContainer = scrollContainerRef.current
    if (!scrollContainer) {
      return
    }

    let animationFrameId: number | null = null

    const updateScrollCue = () => {
      const maxScrollLeft = Math.max(0, scrollContainer.scrollWidth - scrollContainer.clientWidth)
      const nextState: ToolbarScrollCueState = {
        hasOverflow: maxScrollLeft > 1,
        isAtStart: scrollContainer.scrollLeft <= 1,
        isAtEnd: scrollContainer.scrollLeft >= maxScrollLeft - 1,
      }

      setState((currentState) =>
        currentState.hasOverflow === nextState.hasOverflow &&
        currentState.isAtStart === nextState.isAtStart &&
        currentState.isAtEnd === nextState.isAtEnd
          ? currentState
          : nextState,
      )
    }

    const scheduleScrollCueUpdate = () => {
      updateScrollCue()

      if (typeof window.requestAnimationFrame !== 'function' || animationFrameId !== null) {
        return
      }

      animationFrameId = window.requestAnimationFrame(() => {
        animationFrameId = null
        updateScrollCue()
      })
    }

    scheduleScrollCueUpdate()
    scrollContainer.addEventListener('scroll', scheduleScrollCueUpdate, { passive: true })
    window.addEventListener('resize', scheduleScrollCueUpdate)

    let resizeObserver: ResizeObserver | null = null
    if (typeof ResizeObserver !== 'undefined') {
      resizeObserver = new ResizeObserver(scheduleScrollCueUpdate)
      resizeObserver.observe(scrollContainer)
      for (const child of scrollContainer.children) {
        resizeObserver.observe(child)
      }
    }

    let mutationObserver: MutationObserver | null = null
    if (typeof MutationObserver !== 'undefined') {
      mutationObserver = new MutationObserver(() => {
        if (resizeObserver) {
          for (const child of scrollContainer.children) {
            resizeObserver.observe(child)
          }
        }
        scheduleScrollCueUpdate()
      })
      mutationObserver.observe(scrollContainer, {
        attributes: true,
        childList: true,
        subtree: true,
      })
    }

    return () => {
      if (animationFrameId !== null && typeof window.cancelAnimationFrame === 'function') {
        window.cancelAnimationFrame(animationFrameId)
      }
      scrollContainer.removeEventListener('scroll', scheduleScrollCueUpdate)
      window.removeEventListener('resize', scheduleScrollCueUpdate)
      resizeObserver?.disconnect()
      mutationObserver?.disconnect()
    }
  }, [])

  return {
    scrollContainerRef,
    showBackwardCue: state.hasOverflow && !state.isAtStart,
    showForwardCue: state.hasOverflow && !state.isAtEnd,
  } as const
}
