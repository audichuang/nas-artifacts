# nas-artifacts

自架、只在 tailnet 內使用的 Artifacts 服務:把 agent 產生的 HTML/Markdown 頁面、簡報與 PowerPoint/Word/Excel 檔發佈成固定網址,在自己的任何裝置上直接打開,不用登入任何帳號,內容也不離開自己的 NAS。

參考了 Claude 的 Artifacts(版本、檔案與資產、頁面互動能力、留言)與 [open-artifacts](https://github.com/coda0HQ/open-artifacts)(token 發佈、沙箱隔離)。

## 能做什麼

- **發佈與更新**:單檔或整個資料夾(多檔、圖片、JS/CSS);同一個本機路徑再發佈就是同網址的新版本,`?v=N` 看舊版,開著的頁面即時切到新版,並停在原本看的檔案與捲動位置;網址列跟著頁面內的換頁走,複製連結就是正在看的那頁。
- **版本差異**:檢視頁的 ± 與新版提示裡的「看變更」列出這版改了哪些檔、逐行增刪;`artifact diff <id>` 給 agent 看。`artifact prune <id> --keep N` 清掉舊版本並釋放只有它們用到的檔案。
- **Markdown**:程式碼上色、標題錨點、三節以上自動目錄;` ```mermaid ` 在伺服器端畫成 SVG(不在頁面跑腳本,跟著深色模式變色)。
- **簡報**:`*.slides.md`(每頁用一行 `---` 分開,`???` 之後是講稿)在伺服器端畫成 16:9 投影片:方向鍵或滑動翻頁、總覽、講稿、全螢幕、版面依內容自動挑(封面、章節、`|||` 分欄),字太多先自動縮小、再塞不下就記進頁面錯誤。檢視頁的「編輯」在旁邊改原始碼、即時預覽,⌘S 存成新版本;「PPTX」下載成 PowerPoint(文字可編輯),投影片列的「PDF」列印成一頁一張。`artifact export <id>` 給 agent 下載 .pptx。
- **PowerPoint/Word/Excel**:直接發佈 `.pptx`/`.docx`/`.xlsx`(含巨集、放映、範本格式),伺服器端依檔案自己的母片、樣式與主題畫成頁面:`.pptx` 是投影片(同樣有總覽、講稿、PDF),`.docx` 是一整頁文件,`.xlsx` 每個可見的工作表一個分頁;文字可選取、可搜尋,巨集不會執行,下載鈕給回原檔。舊版 `.ppt`/`.doc`/`.xls` 與加密檔只提供下載。Office 檔上限 64 MB。
- **PowerPoint 物件編輯**:選配自架 ONLYOFFICE Docs。最新、未宣告互動能力的 `.pptx` 顯示「編輯」，可直接改文字、拖曳／縮放物件，使用 Office 工具列的「儲存」或 ⌘/Ctrl+S。「返回檢視」會先由 Office 確認未儲存的修改，再結束編輯並等待回存版本；若關閉自動儲存，請先使用原生儲存，或在 Office 提示中選擇繼續編輯。引擎的保存回呼會建立新版本；保存遇到其他人先發佈時不覆寫，提供「下載編輯結果」供合併。編輯引擎獨立於 artifact 沙箱，不會把頁面升權。未設定時維持原來的唯讀預覽。
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
| 持有發佈 token | 發佈、更新、刪除、清理舊版本(prune)、刪除資產、直接讀寫任何 artifact 的 db |

頁面檔案由**另一個 origin** 提供,並帶 `Content-Security-Policy: sandbox allow-scripts …`(沒有 `allow-same-origin`、沒有 popups):頁面是 opaque origin,碰不到圖庫/API 的 origin;`connect-src 'none'` 擋掉 fetch/XHR/WebSocket;資源來源只限它自己的 `/a/<id>/` 路徑,讀不到別的 artifact。檢視頁再送 `frame-src <content>/a/<id>/`,所以框在檢視頁裡的頁面也無法把自己導向外站或別的 artifact。頁面能力只能經檢視頁的 `postMessage` 橋接,由檢視頁依宣告的能力代為呼叫 API。允許的外部腳本 CDN 與 Claude Artifacts 相同(cdnjs、jsdelivr、unpkg、tailwind、jquery)。

直接打開 content 網址(貼上 `artifacts-content…` 連結)時,HTML/Markdown 頁面的頂層導覽(`Sec-Fetch-Dest: document`)會被 302 送回檢視頁(帶 `v` 與 `path`),所以頁面永遠在檢視頁的 `frame-src` 約束下執行,沒有頂層文件可以把自己的內容帶去外站。檢視頁的框(`Sec-Fetch-Dest: iframe`)、不帶該標頭的用戶端(curl、CLI)與非頁面檔案(css、js、圖片)照常提供。

取捨:沙箱沒有 popups,所以頁面裡的外部連結無法開新分頁。

## 部署

一個容器、兩個 port:`8790` 是 app(圖庫、檢視頁、API),`8791` 是 content。前面用反向代理給兩個不同的 hostname。

| 環境變數 | 說明 |
|---|---|
| `APP_ORIGIN` / `CONTENT_ORIGIN` | 兩個對外 origin,必須不同 |
| `APP_PORT` / `CONTENT_PORT` / `LISTEN_HOST` | 預設 `8790` / `8791` / `0.0.0.0` |
| `ARTIFACTS_TOKEN_FILE` | 發佈 token 檔(一行一個,可放多個,`#` 開頭的行略過);或 `ARTIFACTS_TOKENS` 逗號分隔。`NODE_ENV=production`(image 預設)時每個 token 至少 16 字元,否則啟動失敗 |
| `DATA_DIR` | SQLite 與 blob 存放處,預設 `/data` |
| `ONLYOFFICE_URL` | 選配，瀏覽器能連到的自架 Document Server origin，必須與 app、content 不同 |
| `ONLYOFFICE_INTERNAL_URL` | 伺服器下載結果的內部 origin，預設同 `ONLYOFFICE_URL` |
| `ONLYOFFICE_CALLBACK_ORIGIN` | Document Server 能連到的 app origin，預設 `APP_ORIGIN`；NAS 本機連不到 tailnet 名稱時使用區網 app 位址 |
| `ONLYOFFICE_JWT_SECRET_FILE` | 與 Document Server 相同的 JWT 秘密檔，至少 32 字元；設定 `ONLYOFFICE_URL` 時必填 |

ONLYOFFICE Docs 以獨立服務執行；部署設定仍放 nas-ops。開啟 JWT、使用相同秘密，並讓它能連到上述 callback origin；內部位址可能需要 Document Server 的 `ALLOW_PRIVATE_IP_ADDRESS=true`。不要使用公開試用站處理私人文件。瀏覽器與 app 只信任設定的 Office origin，保存下載不跟隨 redirect，回呼只採用簽章涵蓋的內容。

`DATA_DIR/office-edits/` 保存編輯來源與尚待提交的結果，服務重啟後仍可取回。衝突副本不自動刪除；已完成且保存成功（或無修改）的副本在一天後可清理，每次開啟新編輯時執行，上限 200 次工作階段。這個目錄也應納入資料備份。成功保存仍使用原本的版本／blob 提交流程；舊版 CLI 不需更新。

Image:`ghcr.io/audichuang/nas-artifacts`(CI 在 main 每次 push 時建置,標籤 `sha-<完整 sha>`;部署時釘 digest)。健康檢查:`/healthz`。實際部署設定在 nas-ops repo 的 `hosts/ttnas/artifacts/`。

## 開發

伺服器需要 Node 24+(直接執行 `.ts`);CLI 只需要 Node 20+。

```bash
pnpm install
pnpm dev         # app http://localhost:8790, content http://127.0.0.1:8791, token dev-token
pnpm tc && pnpm test
```

瀏覽器 e2e 測試需要 Chromium:設 `CHROMIUM_PATH`,或 `pnpm exec playwright-core install chromium-headless-shell`。

## 給 agent 使用

用 aghub 把 skill 裝到每台機器的各個 agent,再設定 token:

```bash
aghub-cli -g source sync audichuang/nas-artifacts --update --yes
op read "op://Agent/nas-artifacts publish token/credential" | node <skill 資料夾>/scripts/artifact.mjs config --token-stdin
```
