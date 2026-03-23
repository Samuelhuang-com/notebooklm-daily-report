# 投資日報自動化 – 設定指南

每天早上 9:00 自動執行分析，結果送到 Email、LINE、Google Drive。
NB 完全不需要開著。

---

## 整體流程

```
GitHub Actions (雲端) 每天 09:00
    ↓
YouTube API 搜尋熱門影片
    ↓
NotebookLM API 分析（Summary + Report）
    ↓
yfinance 取得投資報價
    ↓
生成完整 HTML 報告
    ↓
同時送出：
  📧 Email（附件 HTML）
  📱 LINE Notify（摘要文字 + Drive 連結）
  ☁️  Google Drive（完整 HTML）
```

---

## Step 1：GitHub 倉庫設定

### 1-1. 建立倉庫

```bash
# 在你的 GitHub 建立新倉庫（private 即可）
# 名稱建議：notebooklm-daily-report
```

### 1-2. 上傳以下檔案到倉庫根目錄

```
your-repo/
├── headless_runner.py          ← 核心執行腳本
├── requirements_headless.txt   ← 相依套件
└── .github/
    └── workflows/
        └── daily_analysis.yml  ← GitHub Actions 排程
```

---

## Step 2：設定 GitHub Secrets

進入倉庫 → Settings → Secrets and variables → Actions → New repository secret

### 必填 Secrets

| Secret 名稱           | 說明                             | 取得方式             |
| --------------------- | -------------------------------- | -------------------- |
| `YOUTUBE_API_KEY`   | YouTube Data API Key             | Google Cloud Console |
| `NOTEBOOK_ID`       | NotebookLM Notebook ID           | NotebookLM 網址列    |
| `STORAGE_STATE_B64` | NotebookLM 登入 Cookie（base64） | 見 Step 3            |
| `ANALYSIS_COMMAND`  | 分析指令                         | 見下方範例           |

### 選填 Secrets

| Secret 名稱            | 說明                                                     |
| ---------------------- | -------------------------------------------------------- |
| `EMAIL_TO`           | 收件人，多人用逗號分隔，例如 `a@gmail.com,b@gmail.com` |
| `EMAIL_FROM`         | 寄件 Gmail，例如 `yourname@gmail.com`                  |
| `EMAIL_APP_PASSWORD` | Gmail App 密碼（非一般密碼，見 Step 4）                  |
| `LINE_NOTIFY_TOKEN`  | LINE Notify Token（見 Step 5）                           |
| `GDRIVE_SA_JSON_B64` | Google Drive 服務帳號 JSON（base64）（見 Step 6）        |
| `GDRIVE_FOLDER_ID`   | Google Drive 資料夾 ID                                   |
| `EXTRA_TICKERS`      | 額外報價代號，例如 `2330,AAPL,GC=F,^GSPC`              |

### ANALYSIS_COMMAND 範例

```
台股 美股 ETF 投資 財經｜請整理今天最熱門的投資影片，重點萃取：提到的股票/ETF、目標價、博主核心觀點、共識與分歧。輸出：先5-8點重點摘要，再影片表格，最後股票彙整表。
```

---

## Step 3：取得 STORAGE_STATE_B64（最重要）

NotebookLM 需要你的 Google 帳號登入狀態。

### 方法：用現有 GUI 程式取得

**3-1.** 在你的 NB 打開 GUI 程式，點擊「🔑 登入 NotebookLM」，完成 Google 登入

**3-2.** 登入成功後，`~/.notebooklm/storage_state.json` 會自動建立

**3-3.** 在 PowerShell 執行以下指令，把內容轉成 base64：

```powershell
# Windows PowerShell
$bytes = [System.IO.File]::ReadAllBytes("$env:USERPROFILE\.notebooklm\storage_state.json")
$b64 = [Convert]::ToBase64String($bytes)
$b64 | Set-Clipboard
Write-Host "已複製到剪貼簿，長度：$($b64.Length) 字元"
```

**3-4.** 把複製的 base64 字串貼到 GitHub Secret `STORAGE_STATE_B64`

> ⚠️  **注意** ：storage_state.json 包含登入 Cookie，有效期約 2–4 週。
> 過期後需要重新在 NB 上用 GUI 登入，再重複 3-3 步驟更新 Secret。

---

## Step 4：Gmail App Password 設定

（不用 App Password 就不設定 EMAIL_ 相關 Secrets）

**4-1.** 登入 Google 帳號 → 管理帳戶 → 安全性

**4-2.** 開啟「兩步驟驗證」（若尚未開啟）

**4-3.** 搜尋「App 密碼」→ 建立 → 選擇「郵件」→ 複製 16 碼密碼

**4-4.** 把 16 碼密碼（不含空格）存入 `EMAIL_APP_PASSWORD`

---

## Step 5：LINE Notify Token

**5-1.** 前往 https://notify-bot.line.me/zh_TW/

**5-2.** 登入 → 個人頁面 → 發行存取權杖

**5-3.** 服務名稱填「投資日報」，選擇要接收的聊天室

**5-4.** 複製 Token → 存入 `LINE_NOTIFY_TOKEN`

---

## Step 6：Google Drive 服務帳號設定

**6-1.** 前往 Google Cloud Console → IAM 與管理 → 服務帳戶

**6-2.** 建立服務帳戶（名稱：`notebooklm-drive`）

**6-3.** 建立金鑰 → JSON → 下載 JSON 檔

**6-4.** 在 PowerShell 轉成 base64：

```powershell
$bytes = [System.IO.File]::ReadAllBytes("C:\path\to\service-account.json")
$b64 = [Convert]::ToBase64String($bytes)
$b64 | Set-Clipboard
```

**6-5.** 存入 `GDRIVE_SA_JSON_B64`

**6-6.** 在 Google Drive 建立資料夾「投資日報」→ 右鍵分享 → 把服務帳號的 Email 加為「編輯者」

**6-7.** 複製資料夾 URL 中的 ID（網址最後一段）→ 存入 `GDRIVE_FOLDER_ID`

```
https://drive.google.com/drive/folders/1aBcDeFgHiJkLmNoPqRsTuVwXyZ
                                        ↑ 這一段就是 FOLDER_ID
```

---

## Step 7：調整執行時間

修改 `.github/workflows/daily_analysis.yml` 中的 cron 設定：

```yaml
schedule:
  - cron: "0 1 * * 1-5"
#          │ │ │ │ └─ 週一到週五
#          │ │ │ └─── 每月（* = 每天）
#          │ │ └───── 每日（* = 每天）
#          │ └─────── UTC 01:00 = 台北 09:00
#          └───────── 分鐘 0
```

常用時間對照：

| 台北時間       | UTC Cron                           |
| -------------- | ---------------------------------- |
| 每天 07:00     | `0 23 * * *`（前一天 UTC 23:00） |
| 每天 08:00     | `0 0 * * *`                      |
| 每天 09:00     | `0 1 * * *`                      |
| 每天 18:00     | `0 10 * * *`                     |
| 週一到五 09:00 | `0 1 * * 1-5`                    |

---

## Step 8：手動測試

**8-1.** 進入 GitHub 倉庫 → Actions 分頁

**8-2.** 左側點選「投資日報自動分析」

**8-3.** 右上角「Run workflow」→ 可輸入自訂指令 → 執行

**8-4.** 點進執行紀錄查看 Log，確認每個步驟都通過

**8-5.** 執行完成後在 Artifacts 下載 HTML 報告確認內容

---

## 常見問題

### Q: storage_state 過期了怎麼辦？

A: 在 NB 上重新用 GUI 程式點「登入 NotebookLM」，完成後重複 Step 3 的 3-3 步驟，更新 GitHub Secret。

### Q: YouTube API 配額超過怎麼辦？

A: 每天只執行一次約消耗 400–600 單位，10,000 單位的限額完全夠用。若用多個指令可能需要申請提高配額或用多個 API Key。

### Q: GitHub Actions 每月免費額度夠嗎？

A: 公開倉庫：無限。私有倉庫：2,000 分鐘/月。每次分析約 10–15 分鐘，一個月 22 個交易日 = 220–330 分鐘，完全夠用。

### Q: 如何更換分析主題？

A: 直接修改 `ANALYSIS_COMMAND` Secret，或在 GitHub Actions 手動執行時輸入新指令。
