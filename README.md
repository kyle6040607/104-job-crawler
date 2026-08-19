# 104 Job Crawler

一個以 Python 撰寫的 104 人力銀行職缺搜尋工具。可依關鍵字、工作經歷、學歷與職缺內容篩選資料，並將結果整理成終端機摘要與 Excel 報表。

摘要直接從 104 的職缺原文與欄位整理，不使用外部 AI 服務，也不需要 API 金鑰。

## 功能

- 搜尋 104 職缺並逐頁取得結果
- 支援精確搜尋或 104 預設的模糊搜尋
- 依職稱、職缺內容、排除詞、工作經歷及學歷篩選
- 擷取職缺條件、工作內容、福利與公司簡介
- 在終端機顯示重點摘要
- 匯出含超連結與自動篩選功能的 Excel 報表
- 自動控制請求間隔，並在暫時失敗時重試

## 環境需求

- Python 3.13 以上
- [uv](https://docs.astral.sh/uv/)（建議）

## 安裝

```powershell
git clone https://github.com/kyle6040607/104-job-crawler.git
cd 104-job-crawler
uv sync
```

若不使用 uv，也可以建立虛擬環境後安裝相依套件：

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install openpyxl requests
```

## 使用方式

### 互動模式

直接執行程式，依畫面提示輸入搜尋條件：

```powershell
uv run python main.py
```

### 命令列模式

搜尋 Python 職缺，內容需包含「遠端」，並限制工作經歷、學歷與筆數：

```powershell
uv run python main.py --keyword python --content 遠端 --exp 2,3 --edu 4 --limit 20
```

只保留職稱中含有關鍵字的職缺：

```powershell
uv run python main.py --keyword 資料工程師 --title --limit 20
```

排除結構、土木及工地相關結果：

```powershell
uv run python main.py --keyword 工程師 --exclude 結構,土木,工地 --limit 20
```

多個內容關鍵字預設必須全部符合；加入 `--any` 後改為任一符合：

```powershell
uv run python main.py --keyword 工程師 --content Python,SQL --any --limit 20
```

## 參數

| 參數 | 說明 |
| --- | --- |
| `--keyword` | 職缺搜尋關鍵字 |
| `--content` | 內容關鍵字，以逗號分隔；預設需全部符合 |
| `--any` | 內容關鍵字改為任一符合 |
| `--exclude` | 要排除的關鍵字，以逗號分隔 |
| `--title` | 只保留職稱包含搜尋關鍵字的職缺 |
| `--loose` | 使用 104 預設的模糊比對 |
| `--exp` | 工作經歷選單編號，可用逗號複選 |
| `--edu` | 學歷選單編號，可用逗號複選 |
| `--limit` | 最多輸出的職缺筆數 |

工作經歷編號：`0` 不拘、`1` 1 年以下、`2` 1～3 年、`3` 3～5 年、`4` 5～10 年、`5` 10 年以上。

學歷編號：`0` 不拘、`1` 高中以下、`2` 高中、`3` 專科、`4` 大學、`5` 碩士、`6` 博士。

完整說明也可由程式查詢：

```powershell
uv run python main.py --help
```

## 輸出

程式會在終端機顯示每筆職缺的重點資訊，並在目前目錄產生：

```text
104_<搜尋關鍵字>_<日期時間>.xlsx
```

Excel 包含職稱、公司、產業、地點、薪資、經歷、學歷、技能、摘要、福利、職缺及公司網址、公司簡介與職缺全文等欄位。產生的 Excel 已由 `.gitignore` 排除，不會被意外提交。

## 注意事項

- 本工具僅供個人查詢與資料整理使用。
- 請遵守 104 人力銀行的服務條款、網站規範與相關法律。
- 網站 API 或欄位格式若調整，程式可能需要同步更新。
- 程式預設在請求之間暫停，請勿移除限制或進行高頻率抓取。

## 授權

目前尚未指定開源授權。未經作者許可，請勿將程式用於商業用途或大量資料蒐集。
