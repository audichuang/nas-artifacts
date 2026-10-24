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
    window.parent.postMessage({ __artifact: 1, report: { type: 'place', path: location.pathname, scrollY: Math.round(scrollY) } }, APP_ORIGIN)
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
  if (framed) {
    addEventListener('DOMContentLoaded', reportPlace)
    addEventListener('hashchange', reportPlace)
    addEventListener('scroll', () => { reportTimer ??= setTimeout(reportPlace, 250) }, { passive: true })
  }

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
