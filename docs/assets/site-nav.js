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

  if (url.origin !== window.location.origin || url.pathname !== window.location.pathname || url.hash.length === 0) {
    return null
  }

  return url.hash
}

const findHashTarget = (hash) => {
  try {
    return document.querySelector(hash)
  } catch {
    return null
  }
}

const scrollToTarget = (target) => {
  target.scrollIntoView({ block: 'start', behavior: 'auto' })
}

;(() => {
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
})()
