// Slide deck runtime, inlined into every *.slides.md page. The slide on show lives in the hash (#3),
// which the bridge reports to the viewer so a new version opens on the same slide.
;(() => {
  const stage = document.getElementById('deck')
  const counter = document.getElementById('deck-counter')
  const notesPanel = document.getElementById('deck-notes')
  const button = (name) => document.querySelector(`[data-deck="${name}"]`)
  let slides = []
  let index = 0
  // The editor's live preview swaps the slides in; drafts are not reported as page errors.
  let previewing = false
  const reported = new Set()

  const fromHash = () => {
    const number = Number(/^#(\d{1,4})$/.exec(location.hash)?.[1])
    return Number.isInteger(number) && number >= 1 ? number - 1 : 0
  }

  function setHash(number) {
    if (location.hash === `#${number}`) return
    // replaceState keeps slide changes out of the browser's back button; the event tells the bridge.
    try { history.replaceState(null, '', `#${number}`) } catch { return }
    dispatchEvent(new HashChangeEvent('hashchange'))
  }

  function go(next) {
    if (slides.length === 0) return
    index = Math.max(0, Math.min(slides.length - 1, next))
    slides.forEach((slide, at) => slide.classList.toggle('active', at === index))
    counter.textContent = `${index + 1} / ${slides.length}`
    button('prev').disabled = index === 0
    button('next').disabled = index === slides.length - 1
    const notes = slides[index].querySelector('.slide-notes')
    notesPanel.innerHTML = notes ? notes.innerHTML : ''
    setHash(index + 1)
    if (document.body.classList.contains('overview')) slides[index].scrollIntoView({ block: 'nearest' })
  }

  // Shrinks a slide's text until it fits, like a presentation app's autofit; past 55% it reports the overflow.
  function fit(slide, at) {
    const inner = slide.querySelector('.slide-inner')
    if (!inner) return
    let scale = 1
    inner.style.setProperty('--fit', '1')
    while (inner.scrollHeight > inner.clientHeight + 1 && scale > 0.56) {
      scale = Math.round((scale - 0.05) * 100) / 100
      inner.style.setProperty('--fit', String(scale))
    }
    const key = `${at}:${inner.textContent.length}`
    if (inner.scrollHeight > inner.clientHeight + 1 && !previewing && !reported.has(key)) {
      reported.add(key)
      setTimeout(() => { throw new Error(`第 ${at + 1} 頁內容太多:字縮到 55% 仍超出版面,請拆成兩頁或刪減`) })
    }
  }

  function fitAll() {
    // Thumbnails share the full-size ratio, so measure in the normal view only.
    if (document.body.classList.contains('overview')) return
    slides.forEach(fit)
  }

  function load(at) {
    slides = [...stage.querySelectorAll('.slide')]
    for (const image of stage.querySelectorAll('img')) {
      if (!image.complete) image.addEventListener('load', fitAll, { once: true })
    }
    fitAll()
    go(at)
  }

  function setOverview(on) {
    document.body.classList.toggle('overview', on)
    button('overview').setAttribute('aria-pressed', String(on))
    if (on) slides[index]?.scrollIntoView({ block: 'nearest' })
    else fitAll()
  }

  function setNotes(on) {
    document.body.classList.toggle('show-notes', on)
    notesPanel.hidden = !on
    button('notes').setAttribute('aria-pressed', String(on))
    requestAnimationFrame(fitAll)
  }

  function fullscreen() {
    if (document.fullscreenElement) document.exitFullscreen?.().catch(() => {})
    else document.documentElement.requestFullscreen?.().catch(() => {})
  }

  const actions = {
    prev: () => go(index - 1),
    next: () => go(index + 1),
    overview: () => setOverview(!document.body.classList.contains('overview')),
    notes: () => setNotes(notesPanel.hidden),
    print: () => {
      setOverview(false)
      window.print()
    },
    fullscreen
  }
  document.querySelector('.deck-bar').addEventListener('click', (event) => {
    const name = event.target.closest('[data-deck]')?.dataset.deck
    if (name) actions[name]()
  })

  stage.addEventListener('click', (event) => {
    if (!document.body.classList.contains('overview')) return
    const slide = event.target.closest('.slide')
    if (!slide) return
    go(slides.indexOf(slide))
    setOverview(false)
  })

  const KEYS = {
    ArrowRight: 'next', ArrowDown: 'next', PageDown: 'next', ' ': 'next', Enter: 'next',
    ArrowLeft: 'prev', ArrowUp: 'prev', PageUp: 'prev', Backspace: 'prev',
    o: 'overview', O: 'overview', n: 'notes', N: 'notes', f: 'fullscreen', F: 'fullscreen'
  }
  function onKey(key, event) {
    if (key === 'Home') return go(0)
    if (key === 'End') return go(slides.length - 1)
    if (key === 'Escape' && document.body.classList.contains('overview')) return setOverview(false)
    const name = KEYS[key]
    if (!name) return false
    if (event && (name === 'next' || name === 'prev') && event.target.closest?.('button')) {
      if (key === ' ' || key === 'Enter') return false
    }
    actions[name]()
  }
  addEventListener('keydown', (event) => {
    if (event.metaKey || event.ctrlKey || event.altKey || event.isComposing) return
    if (event.target.closest?.('input, textarea, select, [contenteditable]')) return
    if (onKey(event.key, event) !== false) event.preventDefault()
  })

  // Swipe on touch screens.
  let touch = null
  stage.addEventListener('pointerdown', (event) => {
    if (event.pointerType !== 'mouse') touch = { x: event.clientX, y: event.clientY }
  })
  stage.addEventListener('pointerup', (event) => {
    if (!touch || document.body.classList.contains('overview')) return (touch = null)
    const dx = event.clientX - touch.x
    const dy = event.clientY - touch.y
    touch = null
    if (Math.abs(dx) > 40 && Math.abs(dx) > Math.abs(dy) * 1.5) go(index + (dx < 0 ? 1 : -1))
  })

  // In full screen the controls fade out until the pointer moves.
  let idleTimer
  addEventListener('pointermove', () => {
    document.body.classList.remove('idle')
    clearTimeout(idleTimer)
    if (document.fullscreenElement) idleTimer = setTimeout(() => document.body.classList.add('idle'), 2500)
  })
  document.addEventListener('fullscreenchange', () => {
    document.body.classList.remove('idle')
    requestAnimationFrame(fitAll)
  })

  addEventListener('hashchange', () => {
    const at = fromHash()
    if (at !== index && at < slides.length) go(at)
  })
  let resizeTimer
  addEventListener('resize', () => {
    clearTimeout(resizeTimer)
    resizeTimer = setTimeout(fitAll, 120)
  })

  // From the viewer, through the bridge: keys pressed outside the frame, the editor's preview, print.
  addEventListener('artifact-deck', (event) => {
    const message = event.detail
    if (message.type === 'deck-key' && typeof message.key === 'string') onKey(message.key)
    else if (message.type === 'deck-print') actions.print()
    else if (message.type === 'deck-preview' && typeof message.html === 'string') {
      previewing = true
      stage.innerHTML = message.html
      load(Number.isInteger(message.index) ? message.index : index)
    }
  })

  load(fromHash())
})()
