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
    if (message.event?.type === 'print') {
      lightbox?.close()
      print()
      return
    }
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

  // Pictures, Mermaid diagrams and wide Markdown tables open large on a click, on every page. A picture is
  // left alone when the page already gives it a job (inside a link or button, an onclick or role=button,
  // a pointer cursor on its parent), when it is icon-sized, or when it carries data-no-zoom.
  const SKIP = 'a[href], button, label, summary, select, [onclick], [role="button"], [data-no-zoom], body.overview'
  const LIGHTBOX_CSS = `
.artifact-lightbox { box-sizing: border-box; width: 100vw; height: 100vh; max-width: none; max-height: none; margin: 0; padding: 0; border: 0; background: transparent; overflow: auto; overscroll-behavior: contain; touch-action: pan-x pan-y; }
.artifact-lightbox[open] { display: grid; place-items: center; }
.artifact-lightbox::backdrop { background: rgba(0, 0, 0, 0.92); }
.artifact-lightbox.zoomed[open], .artifact-lightbox.sheet[open] { place-items: start; }
.artifact-lightbox .lb-item { display: block; box-sizing: border-box; margin: 16px; max-width: none; max-height: none; height: auto; cursor: zoom-in; box-shadow: 0 8px 40px rgba(0, 0, 0, 0.5); user-select: none; -webkit-user-drag: none; }
.artifact-lightbox img.lb-item { background: #fff; }
.artifact-lightbox.zoomed .lb-item { cursor: zoom-out; }
.artifact-lightbox .lb-sheet { box-sizing: border-box; margin: 16px; padding: 16px; max-width: none; background: var(--bg, #fff); color: var(--text, #1f2328); border-radius: 8px; cursor: auto; user-select: text; }
.artifact-lightbox .lb-sheet table { display: table; overflow: visible; }
.artifact-lightbox .lb-sheet svg { display: block; width: 100%; height: auto; max-width: none; }
.artifact-lightbox .lb-close { position: fixed; top: 10px; right: 12px; z-index: 1; width: 44px; height: 44px; border: 0; border-radius: 50%; background: rgba(0, 0, 0, 0.6); color: #fff; font: 26px/44px system-ui, sans-serif; cursor: pointer; }
:where(img:not([data-no-zoom])), :where(.markdown-body .diagram svg) { cursor: zoom-in; }
:where(a[href], button, label, [onclick], [role="button"]) img, body.overview :is(img, .diagram svg) { cursor: inherit; }
.table-zoom { display: inline-flex; align-items: center; gap: 4px; margin: 0 0 6px; padding: 6px 10px; min-height: 32px; border: 1px solid var(--border, #d1d9e0); border-radius: 6px; background: var(--surface, #f6f8fa); color: var(--text, #1f2328); font: 13px system-ui, sans-serif; cursor: pointer; }
.table-zoom[hidden] { display: none; }
@media (max-width: 600px) { .artifact-lightbox .lb-item, .artifact-lightbox .lb-sheet { margin: 8px; } }
@media print { .table-zoom { display: none !important; } }`
  let lightbox = null

  /**
   * Shows `item` over the page. A zoomable item (picture, diagram) starts fitted to the frame and zooms
   * with a click, double-tap, pinch or ctrl/⌘-wheel (a trackpad pinch); the dialog scrolls to pan.
   */
  function openLightbox(item, { ratio, natural, label, zoomable = true, sheet = false, shown = 0, at = null }) {
    const dialog = document.createElement('dialog')
    dialog.className = `artifact-lightbox${sheet ? ' sheet' : ''}`
    dialog.setAttribute('aria-label', label || '放大檢視')
    item.classList.add('lb-item')
    const close = document.createElement('button')
    close.type = 'button'
    close.className = 'lb-close'
    close.textContent = '×'
    close.setAttribute('aria-label', '關閉')
    dialog.append(item, close)

    let base = 0
    let scale = 1
    const maxScale = () => Math.max(4, (natural || 0) / base)
    function setScale(next, clientX = innerWidth / 2, clientY = innerHeight / 2) {
      if (!zoomable || !base) return
      const before = item.getBoundingClientRect()
      const fx = before.width ? (clientX - before.left) / before.width : 0.5
      const fy = before.height ? (clientY - before.top) / before.height : 0.5
      scale = Math.min(maxScale(), Math.max(1, next))
      item.style.width = `${Math.round(base * scale)}px`
      dialog.classList.toggle('zoomed', scale > 1.01)
      // Keep the point under the finger or cursor where it was.
      const after = item.getBoundingClientRect()
      dialog.scrollLeft += after.left + fx * after.width - clientX
      dialog.scrollTop += after.top + fy * after.height - clientY
    }
    function fit() {
      if (!zoomable) return
      const gutter = innerWidth <= 600 ? 16 : 32
      const room = { w: innerWidth - gutter, h: innerHeight - gutter }
      base = Math.max(1, Math.min(room.w, room.h * ratio, natural || Infinity))
      item.style.width = `${Math.round(base * scale)}px`
    }

    // A click zooms a fitted item to its own size (or 2x when it already fits), and back.
    let lastTap = 0
    item.addEventListener('click', (event) => {
      if (!zoomable) return
      event.stopPropagation()
      if (Date.now() - lastTap < 350) return
      lastTap = Date.now()
      if (scale > 1.01) setScale(1, event.clientX, event.clientY)
      else setScale(natural && natural > base * 1.05 ? natural / base : 2, event.clientX, event.clientY)
    })
    // Two fingers pinch; one finger scrolls the dialog natively (touch-action: pan-x pan-y).
    const pointers = new Map()
    let pinch = null
    dialog.addEventListener('pointerdown', (event) => {
      if (event.pointerType !== 'touch') return
      pointers.set(event.pointerId, { x: event.clientX, y: event.clientY })
      if (pointers.size === 2) {
        const [a, b] = [...pointers.values()]
        pinch = { distance: Math.hypot(a.x - b.x, a.y - b.y) || 1, scale }
      }
    })
    dialog.addEventListener('pointermove', (event) => {
      if (!pointers.has(event.pointerId)) return
      pointers.set(event.pointerId, { x: event.clientX, y: event.clientY })
      if (!pinch || pointers.size !== 2) return
      const [a, b] = [...pointers.values()]
      setScale(pinch.scale * (Math.hypot(a.x - b.x, a.y - b.y) / pinch.distance), (a.x + b.x) / 2, (a.y + b.y) / 2)
    })
    const lift = (event) => {
      pointers.delete(event.pointerId)
      if (pointers.size < 2) pinch = null
    }
    dialog.addEventListener('pointerup', lift)
    dialog.addEventListener('pointercancel', lift)
    dialog.addEventListener('wheel', (event) => {
      if (!event.ctrlKey && !event.metaKey) return
      event.preventDefault()
      setScale(scale * Math.exp(-event.deltaY * 0.01), event.clientX, event.clientY)
    }, { passive: false })

    close.addEventListener('click', () => dialog.close())
    // A click on the dark area around the item lands on the dialog itself.
    dialog.addEventListener('click', (event) => { if (event.target === dialog) dialog.close() })
    const onResize = () => fit()
    addEventListener('resize', onResize)
    dialog.addEventListener('close', () => {
      removeEventListener('resize', onResize)
      dialog.remove()
      if (lightbox === dialog) lightbox = null
    })
    document.body.append(dialog)
    lightbox = dialog
    fit()
    dialog.showModal()
    // When the fitted size is about what the page already showed (a full-width picture on a phone),
    // open zoomed in, around the point that was tapped, so the first tap visibly enlarges it.
    if (zoomable && shown && base < shown * 1.25) {
      const more = natural && natural > base * 1.05 ? natural / base : 2
      setScale(Math.min(2.5, more), at?.x, at?.y)
    }
  }

  function zoomPicture(source, at) {
    const image = document.createElement('img')
    image.src = source.currentSrc || source.src
    image.alt = source.alt
    image.draggable = false
    const width = source.naturalWidth || source.width
    const height = source.naturalHeight || source.height
    openLightbox(image, { ratio: width / (height || 1), natural: width, label: source.alt, shown: source.getBoundingClientRect().width, at })
  }

  function zoomDiagram(figure) {
    const svg = figure.querySelector('svg')
    if (!svg) return
    const box = svg.viewBox?.baseVal
    const width = box?.width || svg.getBoundingClientRect().width
    const height = box?.height || svg.getBoundingClientRect().height
    // Inside a .markdown-body wrapper the diagram keeps the colours markdown.css gives it.
    const sheet = document.createElement('div')
    sheet.className = 'markdown-body lb-sheet'
    const copy = document.createElement('figure')
    copy.className = 'diagram'
    copy.style.margin = '0'
    copy.append(svg.cloneNode(true))
    sheet.append(copy)
    openLightbox(sheet, { ratio: width / (height || 1), natural: width * 3, label: '圖表' })
  }

  function zoomTable(table) {
    const sheet = document.createElement('div')
    sheet.className = 'markdown-body lb-sheet'
    sheet.append(table.cloneNode(true))
    openLightbox(sheet, { zoomable: false, sheet: true, label: '表格' })
  }

  // Wide Markdown tables get a button; the table itself stays selectable and scrollable in place.
  function markWideTables() {
    for (const table of document.querySelectorAll('.markdown-body table')) {
      if (table.closest('dialog')) continue
      let button = table.previousElementSibling
      if (!button?.classList.contains('table-zoom')) {
        button = document.createElement('button')
        button.type = 'button'
        button.className = 'table-zoom'
        button.textContent = '⤢ 放大表格'
        button.addEventListener('click', () => zoomTable(table))
        table.before(button)
      }
      button.hidden = table.scrollWidth <= table.clientWidth + 1
    }
  }

  let lightboxStyle = null
  addEventListener('DOMContentLoaded', () => {
    const style = document.createElement('style')
    style.textContent = LIGHTBOX_CSS
    document.head.append(style)
    lightboxStyle = style
    markWideTables()
  })
  // The stylesheet that makes a table scroll may land after DOMContentLoaded, so measure again once it has.
  addEventListener('load', markWideTables)
  let tableTimer
  addEventListener('resize', () => {
    clearTimeout(tableTimer)
    tableTimer = setTimeout(markWideTables, 150)
  })
  // A page that gives a picture (or its parent) a pointer or zoom cursor of its own runs its own click there,
  // often its own lightbox. Our zoom-in cursor would hide that, so read the cursors without our stylesheet.
  function pageHandlesClick(image) {
    if (lightboxStyle) lightboxStyle.disabled = true
    try {
      return [image, image.parentElement].some((element) => element && /^(pointer|zoom-in)$/.test(getComputedStyle(element).cursor))
    } finally {
      if (lightboxStyle) lightboxStyle.disabled = false
    }
  }

  addEventListener('click', (event) => {
    if (event.defaultPrevented || event.button !== 0 || lightbox) return
    const target = event.target instanceof Element ? event.target : null
    if (!target || target.closest(SKIP)) return
    const figure = target.closest('figure.diagram')
    if (figure?.querySelector('svg')) {
      event.preventDefault()
      return zoomDiagram(figure)
    }
    const image = target.closest('img')
    if (!image || !image.complete || image.naturalWidth === 0) return
    const rect = image.getBoundingClientRect()
    if (Math.max(rect.width, rect.height) < 64) return
    if (pageHandlesClick(image)) return
    event.preventDefault()
    zoomPicture(image, { x: event.clientX, y: event.clientY })
  })
  // Keys stay with the open item (the deck turns pages on window keydown); Escape still closes it.
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
