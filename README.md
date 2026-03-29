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
