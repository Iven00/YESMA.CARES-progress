# CG01 進度網頁（GitHub Pages）

| 項目 | 說明 |
|---|---|
| 頁面 | [`index.html`](./index.html) 進度｜[`exceptions.html`](./exceptions.html) 異常入口 |
| 預期 URL | `https://iven00.github.io/YESMA.CARES/progress/`（開 Pages 後） |
| 異常權威 | Google Sheet `PO_Actions`＋Apps Script Web App（程式 PM／SM 寫入） |
| 票據權威 | Sheet `Tickets`（打通後）；repo Markdown＝快照（**不整批 Issues**） |

## 啟用（一次）

Repo → **Settings → Pages** → Source：`Deploy from a branch` → Branch：`main` → Folder：`/docs` → Save。

## 更新進度

1. 改 `2工作分解過程文件/tickets/進度板.md`／Backlog／事件日誌  
2. 同步改本目錄 `index.html` 數字與焦點  
3. commit + push `main`  
4. 瀏覽器重新整理上述 URL  

單張票若要在 GitHub 討論：手動開一則 Issue 並連回 `TK-xxxx`，不要整批匯入。
