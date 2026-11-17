// window.artifact: page capabilities, served to every artifact page. Calls go to the viewer that
// frames the page (the app origin), which checks what the artifact declared and talks to the API.
;(() => {
  const APP_ORIGIN = __APP_ORIGIN__
  const TIMEOUT_MS = 20000
  const framed = window.parent !== window
  const pending = new Map()
  const listeners = { db: new Set(), uploads: new Set() }
  let seq = 0
  let hello

  function call(method, params, transfer) {
    if (!framed) {
      return Promise.reject(new Error('Page capabilities work only inside the Artifacts viewer.'))
    }
    return new Promise((resolve, reject) => {
      const id = ++seq
      const timer = setTimeout(() => {
        pending.delete(id)
        reject(new Error(`${method} timed out: is this page open in the Artifacts viewer?`))
      }, TIMEOUT_MS)
      pending.set(id, { resolve, reject, timer })
      window.parent.postMessage({ __artifact: 1, id, method, params }, APP_ORIGIN, transfer || [])
    })
  }

  window.addEventListener('message', (event) => {
    if (event.source !== window.parent || event.origin !== APP_ORIGIN) return
    const message = event.data
    if (!message || message.__artifact !== 1) return
    if (message.event?.type === 'restore') {
      restoreScroll(message.event.scrollY)
      return
    }
    // Deck controls (slide decks only); the page is the only one listening, so a plain DOM event will do.
    if (typeof message.event?.type === 'string' && message.event.type.startsWith('deck-')) {
      // A picture shown large holds the deck still; anything else from the viewer closes it first.
      if (lightbox) {
        if (message.event.type === 'deck-key') return
        lightbox.close()
      }
      dispatchEvent(new CustomEvent('artifact-deck', { detail: message.event }))
      return
    }
    if (message.event) {
      for (const listener of listeners[message.event.type] || []) {
        try { listener(message.event) } catch (error) { console.error(error) }
      }
      return
    }
    const entry = pending.get(message.id)
    if (!entry) return
    pending.delete(message.id)
    clearTimeout(entry.timer)
    if (message.error) {
      entry.reject(Object.assign(new Error(message.error.message), { code: message.error.code }))
    } else {
      entry.resolve(message.result)
    }
  })

  // Where the reader is, so the viewer can keep them there across versions and in its address bar.
  let reportTimer = null
  function reportPlace() {
    reportTimer = null
    window.parent.postMessage({ __artifact: 1, report: { type: 'place', path: location.pathname, hash: location.hash.slice(0, 400), scrollY: Math.round(scrollY) } }, APP_ORIGIN)
  }
  // Content drawn by scripts may not be tall enough yet, so keep trying briefly unless the reader scrolls first.
  function restoreScroll(target) {
    if (typeof target !== 'number' || !Number.isFinite(target)) return
    let tries = 0
    const stop = () => { tries = Infinity }
    addEventListener('wheel', stop, { once: true, passive: true })
    addEventListener('touchstart', stop, { once: true, passive: true })
    const attempt = () => {
      if (tries++ > 40) return
      scrollTo(0, target)
      if (Math.abs(scrollY - target) > 2) setTimeout(attempt, 50)
    }
    if (document.readyState === 'complete') attempt()
    else addEventListener('load', attempt, { once: true })
  }
  // What breaks while the page runs (script errors, missing files, sandbox blocks) goes to the viewer's error log.
  const errorQueue = []
  let errorTimer = null
  let errorsReported = 0
  function reportError(kind, message, source) {
    if (!framed || errorsReported >= 20) return
    errorsReported++
    errorQueue.push({ kind, message: String(message).slice(0, 500), source: source ? String(source).slice(0, 500) : null })
    errorTimer ??= setTimeout(() => {
      errorTimer = null
      window.parent.postMessage({ __artifact: 1, report: { type: 'errors', path: location.pathname, errors: errorQueue.splice(0) } }, APP_ORIGIN)
    }, 300)
  }
  const where = (file, line, column) => (file ? `${file}:${line ?? 0}:${column ?? 0}` : null)
  addEventListener('error', (event) => {
    const target = event.target
    if (target instanceof Element) {
      reportError('resource', `<${target.localName}> failed to load`, target.getAttribute('src') ?? target.getAttribute('href'))
    } else {
      reportError('error', event.message || 'Script error', where(event.filename, event.lineno, event.colno))
    }
  }, true)
  addEventListener('unhandledrejection', (event) => {
    const reason = event.reason
    reportError('rejection', reason instanceof Error ? `${reason.name}: ${reason.message}` : String(reason), reason?.stack?.split('\n')[1]?.trim())
  })
  addEventListener('securitypolicyviolation', (event) => {
    reportError('csp', `${event.effectiveDirective} blocked ${event.blockedURI || 'inline code'}`, where(event.sourceFile, event.lineNumber, event.columnNumber))
  })

  if (framed) {
    addEventListener('DOMContentLoaded', reportPlace)
    addEventListener('hashchange', reportPlace)
    addEventListener('scroll', () => { reportTimer ??= setTimeout(reportPlace, 250) }, { passive: true })
  }

  // Pictures on the pages drawn here (Markdown, slides, Word, Excel, PowerPoint) open large on a click.
  // Hand-written pages are left alone: their pictures may already do something when clicked.
  const ZOOMABLE = '.markdown-body img, .slide-inner img, .pp-pic img, .dx-img, .xl-pic, .xl-cellpic'
  const LIGHTBOX_CSS = `
.artifact-lightbox { box-sizing: border-box; width: 100vw; height: 100vh; max-width: none; max-height: none; margin: 0; padding: 0; border: 0; background: transparent; overflow: auto; }
.artifact-lightbox[open] { display: grid; place-items: center; }
.artifact-lightbox::backdrop { background: rgba(0, 0, 0, 0.82); }
.artifact-lightbox img { display: block; max-width: calc(100vw - 32px); max-height: calc(100vh - 32px); width: auto; height: auto; object-fit: contain; background: #fff; cursor: zoom-in; box-shadow: 0 8px 40px rgba(0, 0, 0, 0.5); }
.artifact-lightbox.actual[open] { place-items: start; }
.artifact-lightbox.actual img { max-width: none; max-height: none; margin: 16px; cursor: zoom-out; }
.artifact-lightbox button { position: fixed; top: 10px; right: 12px; width: 36px; height: 36px; border: 0; border-radius: 50%; background: rgba(0, 0, 0, 0.55); color: #fff; font: 22px/36px system-ui, sans-serif; cursor: pointer; }
${ZOOMABLE} { cursor: zoom-in; }
a[href] :is(${ZOOMABLE}), button :is(${ZOOMABLE}), label :is(${ZOOMABLE}), body.overview .slide-inner img, body.overview .pp-pic img { cursor: inherit; }`
  let lightbox = null

  function openLightbox(source) {
    const dialog = document.createElement('dialog')
    dialog.className = 'artifact-lightbox'
    dialog.setAttribute('aria-label', source.alt || '圖片')
    const image = document.createElement('img')
    image.src = source.currentSrc || source.src
    image.alt = source.alt
    const close = document.createElement('button')
    close.type = 'button'
    close.textContent = '×'
    close.setAttribute('aria-label', '關閉')
    dialog.append(image, close)
    // Fitted first; a click shows a picture that was shrunk at its own size, scrollable.
    image.addEventListener('click', () => {
      if (dialog.classList.contains('actual') || image.naturalWidth > image.width + 1 || image.naturalHeight > image.height + 1) {
        dialog.classList.toggle('actual')
      } else {
        dialog.close()
      }
    })
    close.addEventListener('click', () => dialog.close())
    // A click on the dark area around the picture lands on the dialog itself.
    dialog.addEventListener('click', (event) => { if (event.target === dialog) dialog.close() })
    dialog.addEventListener('close', () => {
      dialog.remove()
      if (lightbox === dialog) lightbox = null
    })
    document.body.append(dialog)
    lightbox = dialog
    dialog.showModal()
  }

  addEventListener('DOMContentLoaded', () => {
    const style = document.createElement('style')
    style.textContent = LIGHTBOX_CSS
    document.head.append(style)
  })
  addEventListener('click', (event) => {
    if (event.defaultPrevented || event.button !== 0 || lightbox) return
    const image = event.target instanceof Element ? event.target.closest('img') : null
    if (!image || !image.matches(ZOOMABLE) || !image.complete || image.naturalWidth === 0) return
    if (image.closest('a[href], button, label, body.overview')) return
    event.preventDefault()
    openLightbox(image)
  })
  // Keys stay with the picture while it is open (the deck turns pages on window keydown); Escape still closes it.
  addEventListener('keydown', (event) => {
    if (lightbox) event.stopImmediatePropagation()
  }, true)

  const on = (type) => (listener) => {
    listeners[type].add(listener)
    return () => listeners[type].delete(listener)
  }

  window.artifact = Object.freeze({
    /** Resolves to { id, version, latestVersion, title, capabilities, viewer: { name } }. */
    info: () => (hello ??= call('hello')),
    db: Object.freeze({
      get: (collection, id) => call('db.get', { collection, id }),
      set: (collection, id, data) => call('db.set', { collection, id, data }),
      update: (collection, id, data) => call('db.update', { collection, id, data }),
      delete: (collection, id) => call('db.delete', { collection, id }),
      list: (collection, options = {}) => call('db.list', { collection, ...options }),
      query: (collection, where, options = {}) => call('db.query', { collection, where, ...options }),
      batch: (ops) => call('db.batch', { ops }),
      /** listener({ collection, id, deleted }) runs for every write by any viewer, including this one; { resync: true } (no collection) means events may have been missed, so re-read. */
      subscribe: on('db')
    }),
    uploads: Object.freeze({
      put: async (file, name) => {
        const bytes = await file.arrayBuffer()
        return call('uploads.put', { name: name || file.name || 'file', bytes }, [bytes])
      },
      list: () => call('uploads.list', {}),
      subscribe: on('uploads')
    })
  })
})()
