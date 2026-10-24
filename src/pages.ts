import type { ArtifactDetail } from './artifact-store.ts'
import type { Config } from './config.ts'
import { escapeHtml } from './content-render.ts'

function inlineJson(value: unknown): string {
  return JSON.stringify(value).replace(/</g, '\\u003c')
}

function shell(title: string, body: string, script: string, data: unknown): string {
  return `<!doctype html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml(title)}</title>
<link rel="icon" href="/static/favicon.svg">
<link rel="stylesheet" href="/static/app.css">
</head>
<body>
${body}
<script id="page-data" type="application/json">${inlineJson(data)}</script>
<script type="module" src="/static/${script}"></script>
</body>
</html>`
}

export function galleryPage(config: Config): string {
  return shell(
    'Artifacts',
    `<header class="topbar">
  <a class="brand" href="/"><span class="brand-mark" aria-hidden="true"></span>Artifacts</a>
  <label class="search">
    <span class="visually-hidden">搜尋</span>
    <input id="search" type="search" placeholder="搜尋標題或說明" autocomplete="off">
  </label>
</header>
<main class="gallery">
  <p id="status" class="muted" role="status"></p>
  <ul id="list" class="cards"></ul>
  <button id="more" class="button" hidden>載入更多</button>
</main>`,
    'gallery.js',
    { appOrigin: config.appOrigin }
  )
}

export function viewerPage(config: Config, detail: ArtifactDetail, version: number, path?: string): string {
  const row = detail.versions.find((candidate) => candidate.version === version)
  return shell(
    row?.title ?? detail.title,
    `<header class="topbar viewer-bar">
  <a class="icon-button" href="/" title="所有 Artifacts" aria-label="所有 Artifacts">←</a>
  <div class="viewer-title">
    <span class="artifact-icon" aria-hidden="true">${escapeHtml(detail.icon ?? '')}</span>
    <h1 id="title">${escapeHtml(row?.title ?? detail.title)}</h1>
  </div>
  <label class="version-picker">
    <span class="visually-hidden">版本</span>
    <select id="version"></select>
  </label>
  <div class="bar-actions">
    <button id="pin" class="icon-button" type="button" aria-pressed="false" title="釘選">☆</button>
    <button id="copy" class="icon-button" type="button" title="複製連結">⧉</button>
    <button id="fullscreen" class="icon-button" type="button" title="全螢幕" aria-label="全螢幕">⛶</button>
    <button id="errors" class="button danger" type="button" title="這頁執行時發生的錯誤" hidden>⚠ <span id="error-count"></span></button>
    <button id="toggle-comments" class="button" type="button" aria-expanded="false">留言 <span id="comment-count"></span></button>
    <button id="menu-delete" class="icon-button danger" type="button" title="刪除">🗑</button>
  </div>
</header>
<div class="viewer">
  <iframe id="frame" title="Artifact"
    sandbox="allow-scripts allow-forms allow-modals allow-downloads"
    allow="clipboard-write; fullscreen"></iframe>
  <aside id="comments" class="comments" hidden>
    <div class="comments-head">
      <h2>留言</h2>
      <div class="comments-head-actions">
        <label class="muted small"><input id="show-resolved" type="checkbox"> 顯示已解決</label>
        <button id="close-comments" class="icon-button" type="button" aria-label="關閉留言" title="關閉留言">×</button>
      </div>
    </div>
    <ol id="threads" class="threads"></ol>
    <form id="comment-form" class="comment-form">
      <textarea id="comment-body" rows="3" placeholder="留言…(⌘/Ctrl+Enter 送出)" required></textarea>
      <div class="comment-form-row">
        <span id="reply-target" class="muted small"></span>
        <button class="button primary" type="submit">送出</button>
      </div>
    </form>
  </aside>
</div>
<dialog id="error-dialog" class="error-dialog">
  <div class="comments-head">
    <h2>頁面錯誤</h2>
    <button id="close-errors" class="icon-button" type="button" aria-label="關閉" title="關閉">×</button>
  </div>
  <p class="muted small error-hint">這個版本被開啟時發生的錯誤。Agent 可用 <code>artifact errors ${escapeHtml(detail.id)}</code> 讀取。</p>
  <ol id="error-list" class="threads"></ol>
</dialog>
<div id="toast" class="toast" role="status" hidden></div>`,
    'viewer.js',
    {
      appOrigin: config.appOrigin,
      contentOrigin: config.contentOrigin,
      artifact: detail,
      version,
      path
    }
  )
}
