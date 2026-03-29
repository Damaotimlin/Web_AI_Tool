# Web AI Tool

A local-first research and presentation tool that crawls news sites, filters articles with AI, and generates dual-language Keynote/PDF presentations — powered by any local LLM with an OpenAI-compatible API.

## What It Does

1. **Research** — Enter a topic, and the tool crawls multiple news/tech sites, uses AI to filter and score article relevance, then deep-reads the best matches
2. **Translate & Summarize** — Translates article content into your target language with a detailed summary
3. **Generate Slides** — Produces a dual-language presentation (e.g. zh-TW + English) with themed slides, source citations, and cover images
4. **Export** — Outputs PDF (via LibreOffice) and PPTX, auto-opens when done

## Stack

| Component | Technology |
|-----------|-----------|
| AI Backend | Any OpenAI-compatible local LLM server (LM Studio, Ollama, vLLM, LocalAI, Jan, llama.cpp, Text Gen WebUI) |
| Web Scraping | Requests + BeautifulSoup, Playwright (paywall sites), Jina Reader (fallback) |
| Frontend | [Gradio](https://gradio.app) on `localhost:7860` |
| Slide Export | python-pptx (PPTX) + LibreOffice headless (PDF) |
| Prompts | Centralized in `prompts.py` for easy editing |

## Setup

```bash
# 1. Install Python dependencies
pip install -r requirements.txt

# 2. Install Playwright browser (for paywall site support)
python -m playwright install chromium

# 3. Install LibreOffice (for PDF export)
brew install --cask libreoffice

# 4. Start your local AI provider (any one of these):
#    - LM Studio → enable Local Server on port 1234
#    - Ollama → runs on port 11434 by default
#    - vLLM → runs on port 8000 by default
#    - LocalAI / llama.cpp → port 8080
#    - Jan → port 1337
#    Select your provider in the app's "AI Provider" dropdown

# 5. Run the app
python app.py
```

Open http://localhost:7860 in your browser.

## Project Structure

```
Web_AI_Tool/
├── app.py                 # Main application (UI + pipelines)
├── prompts.py             # All AI prompt templates
├── requirements.txt       # Python dependencies
├── outputs/               # Generated PPTX/PDF files
│   └── .img_cache/        # Downloaded article images
├── .saved_categories.json # Site categories (auto-created)
├── .saved_prefs.json      # User preferences (auto-created)
└── README.md
```

## Features

### Research Tab

- **Topic-based research** — Enter a topic like "Impact of AI regulation on tech companies"
- **Multi-site crawling** — Crawls 10+ sites simultaneously (finance, tech, or custom)
- **AI keyword extraction** — Generates 15-25 search keywords from your topic
- **AI title filtering** — Scores article relevance in batches of 40
- **Deep reading** — AI reads full article content and scores top candidates
- **Topic grouping** — Groups related articles by sub-theme
- **Auto-generate** — Optionally auto-generates a keynote from the top result

### Site Categories

- **Built-in categories**: Finance (MarketWatch, WSJ, Barron's, Economist, Reuters, Investopedia) and Tech (TechCrunch, Ars Technica, The Verge, Wired, VentureBeat, and more)
- **Custom categories** — Add, rename, delete categories via the Manage Categories panel
- **Per-category memory** — Each category remembers its own site selections
- **"All categories" mode** — Combine all sites for broad research

### Crawl Depth

| Mode | Behavior | Pages/Site | Time |
|------|----------|-----------|------|
| Standard | Crawl homepage only | ~80 links | ~2-5 min total |
| Deep | Follow section/category pages matching keywords | ~150 links | ~5-10 min total |

Deep crawl identifies section pages (e.g. `/technology/`, `/markets/`) and prioritizes those matching your search keywords.

### Keynote Tab

- **Direct URL mode** — Paste any article URL to generate a presentation
- **Dual-language** — Primary + secondary language on each slide (e.g. zh-TW + English)
- **3 themes** — Dark, Light, Blue
- **Cover images** — Auto-extracted from article OG/meta images
- **Source citations** — Clickable hyperlinks on each slide
- **Batch generation** — Slides generated in batches of 4 for reliability

### Supported Languages

zh-TW, English, Japanese, Korean, Spanish, French, German

### Export Formats

- **PDF** — Auto-generated via LibreOffice headless
- **PPTX** — Always generated (python-pptx)
- **Auto-download** — PDF auto-downloads in browser; PPTX available as backup

### Paywall Support

For sites you have subscriptions to (WSJ, Barron's, Economist, etc.):

1. The tool auto-loads cookies from your Chrome browser
2. If direct fetch returns a paywall teaser, it falls through to a Playwright browser with your cookies
3. Cookie extraction is **domain-sandboxed** — only the target site's cookies are ever used

You can verify cookie isolation by running:
```bash
python test_cookie_sandbox.py
```

## Supported AI Providers

The tool works with any local LLM server that exposes an OpenAI-compatible `/v1/chat/completions` API. Select your provider from the dropdown — the endpoint URL and port are auto-filled:

| Provider | Default Port | Notes |
|----------|-------------|-------|
| LM Studio | 1234 | Enable "Local Server" in settings |
| Ollama | 11434 | `ollama serve` — models via `ollama pull` |
| vLLM | 8000 | `vllm serve <model>` |
| LocalAI | 8080 | Docker or binary install |
| Jan | 1337 | Enable "Local API Server" in settings |
| llama.cpp | 8080 | `llama-server -m <model>` |
| Text Gen WebUI | 5000 | oobabooga with API extension enabled |
| Custom | any | Enter any URL manually |

The selected provider and URL are saved between sessions.

## Model Routing

The tool supports a **2-stage pipeline** when two models are loaded:

| Stage | Model Pattern | Purpose |
|-------|--------------|---------|
| Extraction | `*35b-a3b*` | Keyword extraction, title filtering (fast) |
| Refinement | `*claude*` or `*reasoning*` | Translation, slide generation (quality) |

If only one model is loaded, it's used for everything. Select "auto" in the model dropdown to enable routing.

## Fetch Pipeline

Content fetching uses a 4-method fallback chain:

1. **Direct fetch** with session cookies + full browser headers
2. **Direct fetch** with auto-loaded Chrome domain cookies
3. **Playwright browser** with domain-scoped cookies (for paywalled content)
4. **Jina Reader** as last resort

Sites with RSS feeds (WSJ, Barron's, Economist, TechCrunch, etc.) fall back to RSS when direct crawling is blocked.

## Number Formatting

Prompts enforce consistent number formatting:

- **English**: `$15B`, `$1.25M`, `3,500`
- **Chinese**: `150億美元`, `1,250萬`
- Always digits, never spelled out

## Configuration

All preferences are auto-saved between sessions:

- AI provider and endpoint URL
- Language selections
- Slide count
- Theme
- Site categories and selections
- Crawl depth
- Auto-keynote toggle
- Last used category

## Requirements

- **macOS** (for Keynote integration and browser cookie access)
- **Python 3.10+**
- **Local LLM server** — any OpenAI-compatible provider (LM Studio, Ollama, vLLM, LocalAI, Jan, llama.cpp, or Text Gen WebUI)
- **LibreOffice** (for PDF export) — `brew install --cask libreoffice`
- **Chrome** (optional, for paywall cookie extraction)

---

# Web AI Tool — 繁體中文說明

一款本機優先的研究與簡報工具——爬取新聞與科技網站、以 AI 篩選與評分文章、自動生成雙語 Keynote/PDF 簡報。支援任何相容 OpenAI API 的本機大型語言模型。

## 功能簡介

1. **智慧研究** — 輸入主題後，工具同時爬取多個新聞/科技網站，AI 自動提取關鍵字、篩選標題相關性、深度閱讀最佳文章
2. **翻譯摘要** — 將文章內容翻譯為目標語言，產出 500-800 字的詳盡摘要
3. **自動生成簡報** — 產出雙語簡報（如 zh-TW + English），含主題佈景、來源引用、封面圖片
4. **匯出格式** — 自動產出 PDF（透過 LibreOffice）及 PPTX，完成後自動開啟

## 支援的 AI 服務

工具支援任何提供 `/v1/chat/completions` API 的本機 LLM 伺服器，在介面的「AI Provider」下拉選單中切換：

| 服務 | 預設埠號 | 說明 |
|------|---------|------|
| LM Studio | 1234 | 設定中啟用「Local Server」 |
| Ollama | 11434 | `ollama serve`，透過 `ollama pull` 下載模型 |
| vLLM | 8000 | `vllm serve <模型名>` |
| LocalAI | 8080 | Docker 或直接安裝 |
| Jan | 1337 | 設定中啟用「Local API Server」 |
| llama.cpp | 8080 | `llama-server -m <模型檔>` |
| Text Gen WebUI | 5000 | oobabooga，需啟用 API 擴充 |
| 自訂 | 任意 | 手動輸入任何 URL |

選擇的服務與 URL 會自動記憶，下次啟動自動套用。

## 安裝步驟

```bash
# 1. 安裝 Python 套件
pip install -r requirements.txt

# 2. 安裝 Playwright 瀏覽器（用於付費牆網站）
python -m playwright install chromium

# 3. 安裝 LibreOffice（用於 PDF 匯出）
brew install --cask libreoffice

# 4. 啟動本機 AI 服務（任選一個）
#    例：LM Studio → 啟用 Local Server（port 1234）
#    例：Ollama → ollama serve（port 11434）

# 5. 啟動工具
python app.py
```

開啟瀏覽器前往 http://localhost:7860

## 主要功能

### 研究標籤頁
- **主題研究** — 輸入如「AI 監管對科技公司的影響」等主題
- **多站爬取** — 同時爬取 10+ 個網站（財經、科技或自訂）
- **AI 關鍵字提取** — 從主題產生 15-25 個搜尋關鍵字
- **AI 標題篩選** — 批次評分文章相關性（每批 40 篇）
- **深度閱讀** — AI 閱讀完整文章內容並評分
- **主題分群** — 自動將相關文章依子主題分組

### 網站分類管理
- **內建分類**：財經（MarketWatch、WSJ、Barron's、Economist 等）和科技（TechCrunch、The Verge、Wired 等）
- **自訂分類** — 透過「Manage Categories」面板新增、重新命名、刪除分類
- **分類記憶** — 每個分類各自記憶網站選擇

### 爬取深度

| 模式 | 行為 | 每站連結數 | 時間 |
|------|------|-----------|------|
| Standard | 僅爬首頁 | ~80 連結 | ~2-5 分鐘 |
| Deep | 追蹤符合關鍵字的分類頁面 | ~150 連結 | ~5-10 分鐘 |

### 簡報標籤頁
- **直接 URL 模式** — 貼上任何文章網址即可生成簡報
- **雙語簡報** — 每張投影片同時顯示主要語言與次要語言
- **3 種佈景** — 深色、淺色、藍色
- **封面圖片** — 自動從文章擷取 OG/meta 圖片
- **來源引用** — 每張投影片含可點擊的來源超連結

### 付費牆支援

對於已訂閱的網站（WSJ、Barron's、Economist 等）：
1. 工具自動從 Chrome 載入該網站的 Cookie
2. 若取得的內容為付費牆預覽，自動切換至 Playwright 瀏覽器使用您的訂閱 Cookie
3. Cookie 擷取採**網域隔離**——僅使用目標網站的 Cookie，絕不存取其他網站

### 支援語言

繁體中文、English、日本語、한국어、Español、Français、Deutsch

## 系統需求

- **macOS**（Keynote 整合及瀏覽器 Cookie 存取）
- **Python 3.10+**
- **本機 LLM 伺服器**（LM Studio、Ollama、vLLM、LocalAI、Jan、llama.cpp 或 Text Gen WebUI 任一）
- **LibreOffice**（PDF 匯出）— `brew install --cask libreoffice`
- **Chrome**（選用，用於付費牆 Cookie 擷取）
