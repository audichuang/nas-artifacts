# nas-artifacts

自架、只在 tailnet 內使用的 Artifacts 服務:把 agent 產生的 HTML/Markdown 頁面發佈成固定網址,在自己的任何裝置上直接打開,不用登入任何帳號,內容也不離開自己的 NAS。

參考了 Claude 的 Artifacts(版本、檔案與資產、頁面互動能力、留言)與 [open-artifacts](https://github.com/coda0HQ/open-artifacts)(token 發佈、沙箱隔離)。

## 能做什麼

- **發佈與更新**:單檔或整個資料夾(多檔、圖片、JS/CSS);同一個本機路徑再發佈就是同網址的新版本,`?v=N` 看舊版,開著的頁面即時切到新版。
- **圖庫**:首頁列出全部 artifact,可搜尋、釘選。
- **留言**:每個 artifact 都有留言串(回覆、解決),agent 用 `artifact comments <id>` 讀回饋。
- **頁面互動能力**(發佈時用 `--cap` 宣告):
  - `db`:每個 artifact 自己的共享資料庫(set/update/get/list/query/batch),所有看的人即時同步。
  - `uploads`:讓看的人上傳檔案。
- **Agent skill + CLI**:`skill/nas-artifacts/`,零相依的 `scripts/artifact.mjs`。

## 安全模型

| 誰 | 能做什麼 |
|---|---|
| tailnet 外 | 什麼都連不到(DNS 指向 Tailscale 100.x 位址) |
| tailnet 內的瀏覽器 | 看所有頁面、留言、釘選、使用頁面宣告過的能力 |
| 持有發佈 token | 發佈、更新、刪除、直接讀寫任何 artifact 的 db |

頁面檔案由**另一個 origin** 提供,並帶 `Content-Security-Policy: sandbox allow-scripts …`(沒有 `allow-same-origin`、沒有 popups):頁面是 opaque origin,碰不到圖庫/API 的 origin;`connect-src 'none'` 擋掉 fetch/XHR/WebSocket;資源來源只限它自己的 `/a/<id>/` 路徑,讀不到別的 artifact。檢視頁再送 `frame-src <content>/a/<id>/`,所以框在檢視頁裡的頁面也無法把自己導向外站或別的 artifact。頁面能力只能經檢視頁的 `postMessage` 橋接,由檢視頁依宣告的能力代為呼叫 API。允許的外部腳本 CDN 與 Claude Artifacts 相同(cdnjs、jsdelivr、unpkg、tailwind、jquery)。

**已知缺口**:直接打開 content 網址(檢視頁的 ↗,或貼上 `artifacts-content…` 連結)時頁面是頂層文件,沒有 `frame-src` 約束,頁面仍可把自己導向外站、把它自己的內容帶在網址裡。這時沒有頁面能力(拿不到 db),但頁面本身的內容可能外流。修正方向:content 端把 `Sec-Fetch-Dest: document` 的頁面請求轉回檢視頁。

取捨:沙箱沒有 popups,所以頁面裡的外部連結無法開新分頁。

## 部署

一個容器、兩個 port:`8790` 是 app(圖庫、檢視頁、API),`8791` 是 content。前面用反向代理給兩個不同的 hostname。

| 環境變數 | 說明 |
|---|---|
| `APP_ORIGIN` / `CONTENT_ORIGIN` | 兩個對外 origin,必須不同 |
| `ARTIFACTS_TOKEN_FILE` | 發佈 token 檔(一行一個,可放多個);或 `ARTIFACTS_TOKENS` 逗號分隔 |
| `DATA_DIR` | SQLite 與 blob 存放處,預設 `/data` |

Image:`ghcr.io/audichuang/nas-artifacts`(CI 在 main 每次 push 時建置)。實際部署設定在 nas-ops repo 的 `hosts/ttnas/artifacts/`。

## 開發

```bash
pnpm install
pnpm dev         # app http://localhost:8790, content http://127.0.0.1:8791, token dev-token
pnpm tc && pnpm test
```

瀏覽器 e2e 測試需要 Chromium:設 `CHROMIUM_PATH`,或 `pnpm exec playwright-core install chromium-headless-shell`。

## 給 agent 使用

把 `skill/nas-artifacts` 連結到各 agent 的 skills 目錄,再設定 token:

```bash
ln -s "$PWD/skill/nas-artifacts" ~/.claude/skills/nas-artifacts
op read "op://Agent/nas-artifacts publish token/credential" | node skill/nas-artifacts/scripts/artifact.mjs config --token-stdin
```
