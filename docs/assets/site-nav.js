const localHashForLink = (link) => {
  const href = link.getAttribute('href')
  if (!href || href === '#') {
    return null
  }

  if (href.startsWith('#')) {
    return href
  }

  let url
  try {
    url = new URL(href, window.location.href)
  } catch {
    return null
  }

  if (
    url.origin !== window.location.origin ||
    normalizePathname(url.pathname) !== normalizePathname(window.location.pathname) ||
    url.hash.length === 0
  ) {
    return null
  }

  return url.hash
}

const normalizePathname = (pathname) => {
  const withoutIndex = pathname.replace(/\/index\.html$/u, '/')
  return withoutIndex.length > 1 ? withoutIndex.replace(/\/$/u, '') : withoutIndex
}

const findHashTarget = (hash) => {
  const id = hash.startsWith('#') ? hash.slice(1) : hash
  if (id.length === 0) {
    return null
  }

  const byId = document.getElementById(id)
  if (byId !== null) {
    return byId
  }

  try {
    return document.querySelector(hash)
  } catch {
    return null
  }
}

const syncTopbarHeight = () => {
  const topbar = document.querySelector('.topbar')
  if (!(topbar instanceof HTMLElement)) {
    return
  }

  const height = Math.ceil(topbar.getBoundingClientRect().height)
  if (height > 0) {
    document.documentElement.style.setProperty('--topbar-height', `${height}px`)
  }
}

const readScrollMarginTop = (target) => {
  syncTopbarHeight()
  const margin = Number.parseFloat(window.getComputedStyle(target).scrollMarginTop)
  return Number.isFinite(margin) ? margin : 0
}

const scrollToTarget = (target) => {
  const top = Math.max(0, window.scrollY + target.getBoundingClientRect().top - readScrollMarginTop(target))
  window.scrollTo(0, top)
}

const scrollToHash = (hash) => {
  const target = findHashTarget(hash)
  if (target instanceof HTMLElement) {
    scrollToTarget(target)
  }
}

const correctCurrentHashScroll = () => {
  if (window.location.hash.length > 0) {
    window.requestAnimationFrame(() => scrollToHash(window.location.hash))
  }
}

;(() => {
  syncTopbarHeight()

  const topbar = document.querySelector('.topbar')
  if (topbar instanceof HTMLElement && 'ResizeObserver' in window) {
    const resizeObserver = new ResizeObserver(() => {
      syncTopbarHeight()
      correctCurrentHashScroll()
    })
    resizeObserver.observe(topbar)
  }

  window.addEventListener('resize', syncTopbarHeight)
  window.addEventListener('orientationchange', () => {
    window.requestAnimationFrame(() => {
      syncTopbarHeight()
      correctCurrentHashScroll()
    })
  })

  document.addEventListener('click', (event) => {
    const target = event.target
    if (!(target instanceof Element)) {
      return
    }

    const link = target.closest('a[href]')
    if (!(link instanceof HTMLAnchorElement)) {
      return
    }

    const hash = localHashForLink(link)
    if (hash === null) {
      return
    }

    const scrollTarget = findHashTarget(hash)
    if (!(scrollTarget instanceof HTMLElement)) {
      return
    }

    event.preventDefault()
    if (window.location.hash !== hash) {
      window.history.pushState(null, '', hash)
    }
    scrollToTarget(scrollTarget)
  })

  window.addEventListener('hashchange', () => {
    correctCurrentHashScroll()
  })

  window.addEventListener('load', correctCurrentHashScroll)
  window.addEventListener('pageshow', correctCurrentHashScroll)

  if (document.fonts !== undefined) {
    void (async () => {
      try {
        await document.fonts.ready
        correctCurrentHashScroll()
      } catch {
        // Font readiness can reject in unsupported or interrupted page states.
      }
    })()
  }

  correctCurrentHashScroll()
})()
