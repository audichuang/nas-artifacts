# 公開網址、擁有者登入與遠端 MCP（設計稿）

狀態：擁有者確認三階段（2026-11-18）。階段 1 已實作（`src/public-server.ts`）；階段 2、3 尚未開始。

## 目標

1. 每個 artifact 有「公開／私人」設定。公開的拿到網址就能看（唯讀、不列入任何公開清單、`noindex`）；私人的只有擁有者看得到。
2. 網址改走公開網域（Cloudflare Tunnel），不裝 Tailscale 的裝置也能開。
3. 擁有者在任何裝置登入後，可以看私人頁、留言、釘選、用頁面互動能力。
4. Agent 在任何機器都能發佈：先讓現有 CLI 帶 token 走公開網域，之後提供遠端 MCP（OAuth 登入）。

不做：多使用者帳號、公開首頁或作品列表、外人留言、第三方登入服務。

## 分三階段上線

| 階段 | 內容 | 外部可見的變化 | 前提 |
|---|---|---|---|
| 1 | 可見性欄位、公開唯讀的 host、CLI `--public` | 標成公開的頁面可從外網看 | 無登入，私人頁在外網一律 404 |
| 2 | 擁有者 passkey 登入、私人頁的內容 token | 擁有者在外網登入後可看私人頁、留言等 | 階段 1 |
| 3 | 遠端 MCP（Streamable HTTP + OAuth 2.1） | Claude、Codex 等直接連線發佈 | 階段 2 的登入 |

tailnet 上原本的兩個 port 在三個階段都**完全不變**：既有網址、舊版 CLI、tailnet 內免登入檢視照常可用，回滾不受影響。

## Origin 與 host

沙箱的前提是 app 與 content 分屬不同 origin，公開面也維持這個拆法：

| 用途 | tailnet（不變） | 公開（新增） |
|---|---|---|
| app：圖庫、檢視頁、API | `artifacts.ttnas…`（app port） | `<公開 app host>`，新 listener |
| content：artifact 檔案 | `artifacts-content.ttnas…`（content port） | `<公開 content host>`，新 listener |

- 公開面是**另外兩個 listener**，各自只掛白名單路由，其餘一律 404。不在現有路由逐一加判斷，漏一條就外洩。
- 公開 host 定為 `artifacts.audichuang.app`（app）與 `artifacts-content.audichuang.app`（content）。兩者同一個可註冊網域、不同 origin；session cookie 不設 `Domain`，只回 app host，content 收不到；沙箱的 opaque origin 讓頁面也讀寫不到任何 cookie。
- Cloudflare Tunnel `ttnas`（已在跑）新增兩條 ingress 與對應 DNS，用 `cf` CLI 設定。

## 可見性

- `artifacts.visibility TEXT NOT NULL DEFAULT 'private'`；既有 artifact 全部是私人，部署後外網看不到任何東西，直到擁有者把某頁設為公開。
- 設定方式：CLI `publish --public`／`--private`、`artifact visibility <id> public|private`（需 token，走 tailnet）；tailnet 檢視頁的 ⋯ 選單也能切換。登入狀態與公開開關不放在手機上會自動收起的工具列。
- **宣告了互動能力（`--cap db,uploads`）的 artifact 不能設為公開**：那些頁面依賴寫入，公開面又不開放任何寫入。正式站目前 38 個 artifact 中 0 個有宣告。
- CLI 發佈後印的網址：公開頁印公開網址；階段 2 上線後全部印公開網址（私人頁開啟時會要求登入）。
- 公開的只有**目前版本**：報告常是先寫內部草稿、修過才分享，舊版本只留在 tailnet。發新版後公開網址自動顯示新版，舊版的 content 網址回 404。

## 公開面的路由白名單

匿名訪客（未登入）：

| 路由 | 條件 |
|---|---|
| `GET /a/:id` | artifact 為公開；回一個**另寫的小型公開檢視頁**（標題、版本選單、iframe、下載原檔），不重用 tailnet 的 `viewer.js` |
| `GET /static/*`、`/healthz` | 一律 |
| `GET /api/artifacts/:id/files/*` | 公開；下載原檔 |
| content：`/a/:id/v/:n/*`、`/a/:id/assets/*`、`/_/bridge.js`、`/_/markdown.css` | 公開 |

不回給匿名訪客的欄位：發佈機器名與本機資料夾（`source_host`、`source_path`）、留言數、錯誤數、互動能力清單、版本標籤（`--label` 可能寫內部資訊，只顯示版本號與日期）。

檢視頁的標題、版本清單直接畫在 HTML 裡，所以公開面不需要 JSON API 與事件串流；新版本要重新整理才看得到（階段 1 的取捨）。

公開面**完全沒有寫入路徑**：留言、釘選、頁面錯誤回報、`db`、`uploads`、`PUT …/slides`、Office 編輯、發佈、更新、刪除、prune 都不掛。發佈仍走 tailnet（所有發佈的機器都在 tailnet 上）。`export.pptx` 也不掛：它是匿名可觸發的伺服器端運算。

不在白名單：`GET /`（圖庫）、`GET /api/artifacts`（列表、搜尋）、comments、errors、diff、assets 列表、export、所有 token 路由。

### 快取：Cloudflare 邊緣不能留住內容

版本檔案現在送 `Cache-Control: public, max-age=31536000, immutable`（`src/content-server.ts` `serveFile`、Office 圖片路由）。經過 Cloudflare 代理時，邊緣會依副檔名快取圖片等靜態檔；頁面改回私人後，邊緣與瀏覽器仍會回舊的公開副本，階段 2 的過期 token 網址也一樣。

- 公開 listener 的所有回應一律 `Cache-Control: private, no-store`（或 `private, max-age=…` 只給瀏覽器），並用 `cf` 對兩個公開 host 設 cache bypass 規則，兩道都做。
- tailnet listener 維持現狀。
- 測試：公開頁的圖片取得一次 → 改回私人 → 同網址立刻 404，回應標頭沒有 `public`。

### 每組 listener 各自的 origin

目前只認一組 origin 的地方，公開面都要用公開的那組：bridge.js 內嵌的 `__APP_ORIGIN__`（postMessage 的對象）、content CSP 的 `frame-ancestors`、檢視頁的 `frame-src`、頂層開啟時的轉址、API 回傳的 `url`／`contentUrl`。做法是 content server 與 app server 依 listener 建兩個實例，各自帶自己的 `{ appOrigin, contentOrigin }`，而不是在同一個實例裡看 Host 標頭分支。

擁有者登入後（階段 2），公開面開放和 tailnet 相同的能力；發佈、更新、刪除、prune 仍只認 publish token（CLI）或 MCP 的 OAuth token，瀏覽器 session 不能發佈。

## 擁有者登入（階段 2）

**建議用 passkey（WebAuthn）**：Face ID／Touch ID／Windows Hello，手機與電腦都能用，沒有密碼可以被猜，也不依賴第三方。比較：

| 方式 | 優點 | 缺點 |
|---|---|---|
| Passkey（建議） | 防釣魚、無密碼、不靠外部服務 | 要實作註冊與驗證（用成熟套件）；換裝置要先在舊裝置或用 token 加新 passkey |
| 密碼 | 最簡單 | 要自己顧強度、速率限制、外洩風險 |
| Cloudflare Access | 程式最少 | 依賴第三方；content host 的子資源驗證、MCP OAuth 仍要自己處理 |

- **第一個 passkey 的註冊**：CLI `artifact login-link`（需 publish token）印出 10 分鐘有效的一次性連結，打開後註冊 passkey。之後新增裝置也走這條，或從已登入的裝置新增。
- **第一把 passkey 建議在手機上註冊**：之後公司電腦可以用手機掃碼跨裝置登入（不用在公司電腦上存 passkey）。
- **Session**：cookie 名用 `__Host-` 前綴（強制 `Secure`、`Path=/`、不得設 `Domain`，`audichuang.app` 的其他子網域無法覆寫它），`HttpOnly; SameSite=Lax`，有效期 30 天，可登出、撤銷全部。
- **寫入請求**：所有靠 cookie 的寫入都要 `x-artifact-client: viewer` 標頭，強制 CORS preflight，同站的其他子網域也無法代送。
- **速率限制**：優先用 Cloudflare 的速率規則（`cf` 設定，程式最少）。若要在程式裡做，只能信 `CF-Connecting-IP`，而且前提是公開 listener 只有 cloudflared 連得到：artifacts 容器加入 cloudflared 所在的 docker network，公開 port 不發佈到 host。

### 私人頁的內容怎麼驗證（需先做實驗）

私人頁的檔案在 content host，iframe 是 opaque origin；session cookie 在 app host，content host 拿不到。
在 sandbox iframe 裡，連 content host 自己的 cookie 都不一定會隨子資源送出（瀏覽器把 opaque origin 的請求當跨站），而 Safari 會擋第三方 cookie。

做法：檢視頁（已登入）向 app 要一個**短效內容 token**：HMAC(artifact id, 到期時間)，例如 12 小時有效。iframe 改載 `<content>/t/<token>/a/<id>/v/<n>/…`，token 放在路徑裡，頁面內的相對連結自然帶著它。

- content listener 驗 token 後才回檔案；CSP 的自有檔案範圍改成 `<content>/t/<token>/a/<id>/`，一個 artifact 仍讀不到另一個的檔案。
- 公開頁不需要 token，路徑不變。
- token 外流的影響：拿到的人在到期前能看那一份私人頁。頁面本身被 `connect-src 'none'` 擋住，無法主動外送。
- bridge 的錯誤回報會帶檔案網址：存進 `page_errors` 前把 `/t/<token>` 剝掉，token 不落地。
- 頁面開著超過有效期，延遲載入的圖（xlsx 用 `loading="lazy"`）會失敗：到期前由檢視頁換新 token、重載 frame 並還原閱讀位置（沿用現有的換版還原機制）。
- **實作前先寫 real-Chromium 實驗（spike）**，確認 token 路徑在相對連結、`<base>`、CSS `url()`、Office 圖片路由、延遲載入、錯誤回報下都能用，並驗證事件串流（SSE）經過 Cloudflare Tunnel 不會被閒置逾時切斷，再定案。

## 遠端 MCP（階段 3）

先查各目標 client 的遠端 MCP 支援哪種驗證，再決定要不要做 OAuth：若要用的 coding agent（Claude Code、Codex 等）都能對遠端 MCP 帶固定的 `Authorization` 標頭，MCP 只是現有 API 外面一層薄殼，沿用 publish token 即可；OAuth 只有從 claude.ai 網頁或手機 App 連線時才需要。

- app 新增 `POST /mcp`（MCP Streamable HTTP）。工具：`publish`（檔案內容以 base64 傳）、`update`（帶 `baseVersion`，衝突回 409 與變更檔清單，和 CLI 同一套契約）、`list`、`read`、`comments`、`reply`、`set_visibility`、`delete`。
- 授權：OAuth 2.1 + PKCE，這個 app 自己當 authorization server，支援動態註冊（Claude、Codex 連 MCP 時需要）。授權頁要求 passkey 登入，擁有者按「允許」後發 access token（1 小時）與 refresh token；可在設定頁撤銷。
- CLI 與 publish token 保留，作為不支援 MCP 的 agent 與救援用途。

## 安全面的改變

- 一上公開網域，發佈、刪除只剩 publish token 一道牆（tailnet 不再是外圈）。token 維持 32 bytes 以上隨機值，只存摘要。
- 公開面要加：每個 IP 的請求速率上限（登入、token 驗證失敗、MCP）；request body 上限沿用現有值；所有回應帶 `X-Robots-Tag: noindex`。
- AGENTS.md 第一段（「只給一個人在 tailnet 上看」）、memory「個人使用不做防禦」改寫：公開面的輸入上限、速率限制屬信任邊界，必須做。
- 每條公開面路由、可見性切換、匿名寫入被拒、私人頁 404、內容 token 範圍，都要有先紅後綠的 real-Chromium e2e。

## 資料遷移與回滾

- 新欄位 `visibility` 可預設、舊 image 不讀它；新表 `passkeys`、`sessions`、`oauth_clients`、`oauth_tokens` 只增不改。回滾到舊 image：公開 listener 消失、外網全部斷，tailnet 照常。
- Cloudflare Tunnel `ttnas` 也服務 learning 與 ads.parallux.uk。改 ingress 前先用 `cf` 讀出現有設定並存檔，改完確認那兩個服務仍可連，再加新的 host。
- nas-ops：compose 讓 artifacts 加入 cloudflared 的 docker network（公開 port 不發佈到 host）、`PUBLIC_APP_ORIGIN`、`PUBLIC_CONTENT_ORIGIN`、內容 token 的簽章金鑰（0400 檔案，與 publish token 同樣管理）；Cloudflare Tunnel 加兩條 ingress。

## 待決定

1. ~~公開網址用哪兩個 host~~：已定 `artifacts.audichuang.app`、`artifacts-content.audichuang.app`。
2. 階段 1 先上，還是三階段做完再一起上。
3. 外網的私人頁是否也要能留言（擁有者登入後；建議可以，和 tailnet 相同）。
4. 是否需要從 claude.ai 網頁或手機發佈：需要才做 OAuth，否則階段 3 只是帶 token 的 MCP。
