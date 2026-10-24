# nas-artifacts

自架、只在 tailnet 內使用的 Artifacts 服務:把 agent 產生的 HTML/Markdown 頁面發佈成固定網址,在自己的任何裝置上直接打開,不用登入任何帳號,內容也不離開自己的 NAS。

參考了 Claude 的 Artifacts(版本、檔案與資產、頁面互動能力、留言)與 [open-artifacts](https://github.com/coda0HQ/open-artifacts)(token 發佈、沙箱隔離)。

## 能做什麼

- **發佈與更新**:單檔或整個資料夾(多檔、圖片、JS/CSS);同一個本機路徑再發佈就是同網址的新版本,`?v=N` 看舊版,開著的頁面即時切到新版,並停在原本看的檔案與捲動位置;網址列跟著頁面內的換頁走,複製連結就是正在看的那頁。
- **版本差異**:檢視頁的 ± 與新版提示裡的「看變更」列出這版改了哪些檔、逐行增刪;`artifact diff <id>` 給 agent 看。`artifact prune <id> --keep N` 清掉舊版本並釋放只有它們用到的檔案。
- **Markdown**:程式碼上色、標題錨點、三節以上自動目錄;` ```mermaid ` 在伺服器端畫成 SVG(不在頁面跑腳本,跟著深色模式變色)。
- **簡報**:`*.slides.md`(每頁用一行 `---` 分開,`???` 之後是講稿)在伺服器端畫成 16:9 投影片:方向鍵或滑動翻頁、總覽、講稿、全螢幕、版面依內容自動挑(封面、章節、`|||` 分欄),字太多先自動縮小、再塞不下就記進頁面錯誤。檢視頁的「編輯」在旁邊改原始碼、即時預覽,⌘S 存成新版本;「PPTX」下載成 PowerPoint(文字可編輯),投影片列的「PDF」列印成一頁一張。`artifact export <id>` 給 agent 下載 .pptx。
- **圖庫**:首頁列出全部 artifact,可釘選;搜尋比對標題、說明、目前版本的內文(中文兩個字也找得到,附命中的那段文字)與發佈來源的機器和資料夾。
- **留言**:每個 artifact 都有留言串(回覆、解決)。Agent 用 `artifact comments <id> --open` 讀未解決的回饋,`--wait` 等新留言,`artifact comment reply … --resolve` 回覆並結案。
- **頁面錯誤**:頁面執行時的腳本錯誤、載入失敗的檔案與被沙箱擋下的請求會記在該版本底下,檢視頁顯示 ⚠ 數量,agent 用 `artifact errors <id>` 讀。發佈前 CLI 也會先靜態檢查一次(`artifact check`)。
- **頁面互動能力**(發佈時用 `--cap` 宣告):
  - `db`:每個 artifact 自己的共享資料庫(set/update/get/list/query/batch),所有看的人即時同步。
  - `uploads`:讓看的人上傳檔案。
- **Agent skill + CLI**:`skill/nas-artifacts/`,零相依的 `scripts/artifact.mjs`。

## 安全模型

| 誰 | 能做什麼 |
|---|---|
| tailnet 外 | 什麼都連不到(DNS 指向 Tailscale 100.x 位址) |
| tailnet 內的瀏覽器 | 看所有頁面、留言、釘選、回報頁面錯誤(每版最多 100 筆)、使用頁面宣告過的能力、編輯沒宣告任何能力的簡報(存成新版本,只換簡報原始碼那一個檔) |
| 持有發佈 token | 發佈、更新、刪除、直接讀寫任何 artifact 的 db |

頁面檔案由**另一個 origin** 提供,並帶 `Content-Security-Policy: sandbox allow-scripts …`(沒有 `allow-same-origin`、沒有 popups):頁面是 opaque origin,碰不到圖庫/API 的 origin;`connect-src 'none'` 擋掉 fetch/XHR/WebSocket;資源來源只限它自己的 `/a/<id>/` 路徑,讀不到別的 artifact。檢視頁再送 `frame-src <content>/a/<id>/`,所以框在檢視頁裡的頁面也無法把自己導向外站或別的 artifact。頁面能力只能經檢視頁的 `postMessage` 橋接,由檢視頁依宣告的能力代為呼叫 API。允許的外部腳本 CDN 與 Claude Artifacts 相同(cdnjs、jsdelivr、unpkg、tailwind、jquery)。

直接打開 content 網址(貼上 `artifacts-content…` 連結)時,HTML/Markdown 頁面的頂層導覽(`Sec-Fetch-Dest: document`)會被 302 送回檢視頁(帶 `v` 與 `path`),所以頁面永遠在檢視頁的 `frame-src` 約束下執行,沒有頂層文件可以把自己的內容帶去外站。檢視頁的框(`Sec-Fetch-Dest: iframe`)、不帶該標頭的用戶端(curl、CLI)與非頁面檔案(css、js、圖片)照常提供。

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
