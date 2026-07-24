# InsightFlow — AI 輿情分析系統

輸入品牌、排除條件與觀測期間，**一鍵**從 KEYPO 大數據關鍵引擎抓取資料，自動產出
**輿情報告簡報（PPT）**、**詳細分析報告（Word）**、**視覺圖表**與**市場議題分析規劃（Word）**。

把過去「手寫關鍵字 → 排除雜訊/抽獎 → 上 KEYPO 設期間 → 逐項看數據 → 下載表格 → 貼進簡報」
的半天手工流程，收斂成一次執行、數分鐘完成。

---

## 🚀 快速開始

```bash
# 1) 安裝相依
pip install -r requirements.txt

# 2) 前置：KEYPO CLI 置於 ../KEYPO_API/keypo_cli.py；啟動預設 AI 後端 Ollama
ollama pull qwen2.5:7b        # 或改用雲端：於 .env 設 AI_BACKEND=claude/openai + 金鑰

# 3) 執行，依提示輸入品牌、期間等
py main.py
```

完成後於 `output/{品牌}/` 取得**簡報 PPT、詳細分析報告 Word、圖表、市場議題分析規劃 Word**。
詳細設定與其他工具見下方章節。

---

## 產出成果

執行一次 `main.py`，於 `output/{品牌}/` 產出：

| 檔案 | 內容 |
|---|---|
| `{品牌}_..._Report.pptx` | 26 頁 KEYPO 風格簡報：聲量趨勢（事件標註）、來源分布、情緒好感度、關鍵字對應事件、熱門文章、KOL 四維、競品分析、觀測彙整 |
| `{品牌}_..._詳細分析報告.docx` | §1–§12 完整報告（真表格）：執行摘要、聲量趨勢、情緒、關鍵字、熱門文章、KOL、核心議題、風險、AI 洞察、建議、研究方法、競品 |
| `charts/*.png` | 聲量趨勢、情緒甜甜圈、來源分布、關鍵字長條與文字雲、競品比較 |
| `市場議題分析規劃_{品牌}.docx` | 前瞻規劃：AI 依品牌自動擬定「議題 × 分析維度」矩陣與逐議題分析腳本（供分析師核定） |

---

## 系統需求

- **Python 3.10+**（開發環境為 3.14）
- **KEYPO CLI**：`../KEYPO_API/keypo_cli.py`（與本專案同層的另一個目錄，**不含在本 repo**）
- **AI 後端（擇一）**：
  - 本機 [Ollama](https://ollama.com/)（預設、免費、離線）— 需先 `ollama pull qwen2.5:7b` 並在 `localhost:11434` 執行
  - 或雲端 API：Anthropic Claude、OpenAI（需 API 金鑰）
- **作業系統**：Windows（字型寫死 `Microsoft JhengHei`／`msjh.ttc`，並以 `os.startfile` 自動開啟 PPT）

---

## 安裝

```bash
pip install -r requirements.txt
```

相依套件：`httpx`、`matplotlib`、`wordcloud`、`python-pptx`、`python-docx`、
`anthropic`、`python-dotenv`、`opencc-python-reimplemented`（AI 輸出簡→繁保險）。
使用 OpenAI 後端時另需 `pip install openai`。

---

## 設定（`.env`）

於專案根目錄建立 `.env`：

```dotenv
# AI 後端：ollama（預設）／ claude ／ openai
AI_BACKEND=ollama

# 本機 Ollama（免費、離線）
OLLAMA_MODEL=qwen2.5:7b
# OLLAMA_URL=http://localhost:11434/api/generate

# Claude（Anthropic）
# ANTHROPIC_API_KEY=sk-ant-...
# CLAUDE_MODEL=claude-opus-4-8

# OpenAI
# OPENAI_API_KEY=sk-...
# OPENAI_MODEL=gpt-4o-mini

# 簡→繁轉換（預設開啟，設 off 可關閉）
# OPENCC=on
```

> `.env` 已列入 `.gitignore`，金鑰只會留在本機。切換 AI 後端只需改 `AI_BACKEND`，程式不用動。

---

## 使用方式

### 1. 產出完整報告（主流程）

```bash
py main.py
```

依提示輸入：品牌名稱、查詢字串（可用 KEYPO 布林語法）、觀測期間、是否排除抽獎、
自訂雜訊排除、競品（選填，最多 3 個，可各自設排除字串）、產業／背景（選填，
供市場議題規劃 AI 推斷）。完成後自動產出上述四類成果。

### 2. 只產市場議題分析規劃

```bash
py framework_generator.py
```

輸入品牌與產業，AI 依品牌自動擬定議題矩陣與分析腳本，輸出至
`output/{品牌}/市場議題分析規劃_{品牌}.docx`。

### 3. 產出專案實作報告簡報

```bash
py project_showcase.py
```

產出介紹本系統的簡報至 `output/專案實作報告/`。

---

## 系統架構

```
輸入（品牌／排除／期間）
   ↓  keypo_fetcher   呼叫 KEYPO CLI，抓 5 端點（sentidist / hotkw / hotrank / freqdist / sprdtrnd）
   ↓  keypo_loader    解析為情緒／關鍵字／熱門文章／KOL；分類、來源分布、關鍵字對應事件
   ↓  period_history  上期（MoM）／去年同期（YoY）成長率
   ↓  chart_generator matplotlib／wordcloud 產圖
   ↓  ai_backend      可切換後端（ollama／claude／openai）＋簡→繁；driven by local_ai_analyzer
   ↓  local_ai_analyzer  六節洞察、聲量高峰事件敘事、表格摘要
   ↓  輸出：ppt_generator（簡報）／detailed_report_docx（詳細報告 Word）／framework_generator（市場規劃 Word）
```

---

## 專案結構

| 模組 | 職責 |
|---|---|
| `main.py` | 主流程：互動輸入 → 抓取 → 分析 → 產出全部成果 |
| `keypo_fetcher.py` | 呼叫 KEYPO CLI 抓取（含抽獎排除布林語法） |
| `keypo_loader.py` | 解析 KEYPO JSON；文章分類、KOL 聚合、來源分布、關鍵字對應事件 |
| `period_history.py` | 品牌期間快照，計算 MoM／YoY |
| `chart_generator.py` | 各式圖表 |
| `ai_backend.py` | 可切換 AI 後端 + OpenCC 簡→繁 |
| `local_ai_analyzer.py` | AI 敘事（六節分析、趨勢事件敘事、表格摘要） |
| `report_generator.py` | 精簡摘要（內部用） |
| `detailed_report_generator.py` | 詳細報告文字版（並提供 helper 給 Word 版） |
| `detailed_report_docx.py` | 詳細分析報告 **Word 版**（真表格，正式交付） |
| `ppt_generator.py` | 輿情報告簡報（KEYPO 風格） |
| `framework_generator.py` | 市場議題分析規劃 Word（AI 依品牌擬定議題腳本） |
| `project_showcase.py` | 專案實作報告簡報 |
| `ai_analyzer.py` | 早期 Claude API 路徑（目前 `main.py` 未使用） |

---

## AI 後端說明

| 後端 | 品質 | 成本 | 備註 |
|---|---|---|---|
| `ollama`（預設） | 中（本地 7B 模型） | 免費、離線 | 資料留本機；偶有小幻覺，簡體字已由 OpenCC 自動轉繁 |
| `claude` | 高 | 低（每份約數美分） | 需 `ANTHROPIC_API_KEY`；用原生 Anthropic SDK |
| `openai` | 高 | 低 | 需 `OPENAI_API_KEY` 與 `pip install openai` |

所有 AI 輸出統一經過 OpenCC（`s2twp`）簡→繁保險，確保繁體報告不夾簡體字。

---

## 注意事項與已知限制

- **相依外部工具**：需同層目錄有 KEYPO CLI（`../KEYPO_API/keypo_cli.py`）；本機需可執行對應查詢。
- **Windows 限定**：字型與自動開檔為 Windows 行為。
- **本地模型天花板**：`ollama`（qwen2.5:7b）內容深度有限、偶有幻覺；正式交付建議切換雲端後端。
- **來源分布為參考值**：目前以熱門文章篇數推估（非 KEYPO 全站聲量端點）。
- **市場議題規劃 vs 輿情報告**：前者是「接下來要分析哪些市場議題」的**前瞻規劃**，後者是「本期實際數據」的**監測報告**，用途不同。

---

## 未來方向

擴充 KEYPO 端點（datadist／hotchnl／opleader／hotkwV5）· 多議題 × 多品牌矩陣分析 ·
QoQ 季比較 · 市場規劃改為真實資料驅動 · 開發圖形化 UI 介面。
