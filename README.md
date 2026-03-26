# Web → Keynote Generator

Fetches any web page, translates it with a local LLM, and exports a `.pptx` that opens in Keynote.

## Stack
- **AI**: LM Studio (MLX) → `localhost:1234`
- **Web Fetch**: Jina Reader (bot-bypass)
- **Frontend**: Gradio → `localhost:7860`
- **Export**: python-pptx → Keynote

## Setup

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Start LM Studio → enable Local Server on port 1234
#    Load: qwen2.5-72b-instruct (MLX 4bit)

# 3. Run
python src/app.py
```

Open browser at http://localhost:7860

## Project Structure

```
web-to-keynote/
├── src/
│   └── app.py          # Main app (Gradio UI + pipeline)
├── outputs/            # Generated .pptx files
├── .vscode/
│   ├── launch.json     # F5 to run
│   └── tasks.json      # Cmd+Shift+B to run
├── requirements.txt
└── README.md
```

## Usage

1. Paste any URL
2. Select language (zh-TW / English / Japanese / Korean)
3. Set number of slides (3–12)
4. Pick a theme (Dark / Light / Blue)
5. Click **Generate Keynote**
6. `.pptx` auto-opens in Keynote

## Notes
- LM Studio must be running before launching the app
- Outputs saved to `./outputs/` with timestamps
- Change `MODEL` in `app.py` to match your loaded model name in LM Studio
