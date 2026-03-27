import gradio as gr
import requests
import http.cookiejar
import browser_cookie3
import json
import subprocess
import os
import re
from urllib.parse import urljoin, urlparse
from bs4 import BeautifulSoup
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from datetime import datetime

LM_STUDIO_BASE = "http://localhost:1234"
LM_STUDIO_URL = f"{LM_STUDIO_BASE}/v1/chat/completions"
PREFERRED_MODELS = [
    "qwen3.5-35b-a3b",
    "qwen3.5-27b-claude-4.6-opus-reasoning-distilled",
]
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "outputs")
os.makedirs(OUTPUT_DIR, exist_ok=True)
SAVED_SITES_FILE = os.path.join(os.path.dirname(__file__), ".saved_sites.json")
SITES_HISTORY_FILE = os.path.join(os.path.dirname(__file__), ".sites_history.json")
SAVED_PREFS_FILE = os.path.join(os.path.dirname(__file__), ".saved_prefs.json")

DEFAULT_PREFS = {
    "max_articles": 30,
    "slides": 10,
    "auto_keynote": False,
    "lang1": "zh-TW",
    "lang2": "English",
    "theme": "Dark",
}


def load_saved_sites() -> list[str]:
    """Load saved site URLs from disk."""
    try:
        with open(SAVED_SITES_FILE, "r") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return []


def load_sites_history() -> list[str]:
    """Load all ever-used site URLs (history for dropdown choices)."""
    try:
        with open(SITES_HISTORY_FILE, "r") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return []


def save_sites_history(sites: list[str]):
    """Add sites to history (never removes, only appends new ones)."""
    history = load_sites_history()
    existing = set(history)
    for s in sites:
        s = s.strip()
        if s and s not in existing:
            history.append(s)
            existing.add(s)
    with open(SITES_HISTORY_FILE, "w") as f:
        json.dump(history, f)


def load_saved_prefs() -> dict:
    """Load saved user preferences from disk."""
    try:
        with open(SAVED_PREFS_FILE, "r") as f:
            saved = json.load(f)
            return {**DEFAULT_PREFS, **saved}
    except (FileNotFoundError, json.JSONDecodeError):
        return DEFAULT_PREFS.copy()


def save_prefs(**kwargs):
    """Save user preferences to disk (merges with existing)."""
    prefs = load_saved_prefs()
    prefs.update(kwargs)
    with open(SAVED_PREFS_FILE, "w") as f:
        json.dump(prefs, f)


def save_sites(sites: list[str]):
    """Save active site URLs to disk, and add to history."""
    seen = set()
    unique = []
    for s in sites:
        s = s.strip()
        if s and s not in seen:
            seen.add(s)
            unique.append(s)
    with open(SAVED_SITES_FILE, "w") as f:
        json.dump(unique, f)
    # Also add to permanent history
    save_sites_history(unique)


def get_available_models() -> list[str]:
    """Fetch available model IDs from LM Studio."""
    try:
        res = requests.get(f"{LM_STUDIO_BASE}/v1/models", timeout=5)
        res.raise_for_status()
        return [m["id"] for m in res.json().get("data", [])]
    except Exception:
        return []


def pick_default_model(available: list[str]) -> str:
    """Return the first preferred model that is available, or the first available model."""
    for pref in PREFERRED_MODELS:
        for model_id in available:
            if pref.lower() in model_id.lower():
                return model_id
    return available[0] if available else ""


BROWSER_UA = "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"


def build_session(cookies_text: str = "") -> requests.Session:
    """Build a requests.Session with optional cookies.

    Accepts either:
      - Raw cookie header string: "name1=val1; name2=val2"
      - Path to a Netscape/Mozilla cookies.txt file
    """
    session = requests.Session()
    session.headers.update({"User-Agent": BROWSER_UA})

    cookies_text = (cookies_text or "").strip()
    if not cookies_text:
        return session

    # If it looks like a file path, try loading as cookies.txt
    if os.path.isfile(cookies_text):
        jar = http.cookiejar.MozillaCookieJar(cookies_text)
        jar.load(ignore_discard=True, ignore_expires=True)
        session.cookies.update(jar)
        return session

    # Otherwise parse as "name=value; name2=value2" header string
    for pair in cookies_text.split(";"):
        pair = pair.strip()
        if "=" in pair:
            name, value = pair.split("=", 1)
            session.cookies.set(name.strip(), value.strip())

    return session


BROWSER_LOADERS = {
    "Chrome":  browser_cookie3.chrome,
    "Firefox": browser_cookie3.firefox,
    "Safari":  browser_cookie3.safari,
    "Edge":    browser_cookie3.edge,
}


def load_browser_cookies(browser_name: str, domain: str) -> str:
    """Extract cookies from an installed browser for a given domain.
    Returns a cookie header string (name=val; name2=val2).
    """
    loader = BROWSER_LOADERS.get(browser_name)
    if not loader:
        return f"Error: Unknown browser '{browser_name}'"
    try:
        jar = loader(domain_name=domain)
        pairs = [f"{c.name}={c.value}" for c in jar if domain in (c.domain or "")]
        if not pairs:
            return f"Error: No cookies found for {domain} in {browser_name}. Are you logged in?"
        return "; ".join(pairs)
    except PermissionError:
        return f"Error: Permission denied reading {browser_name} cookies. Close {browser_name} and retry, or grant Full Disk Access in System Settings > Privacy."
    except Exception as e:
        return f"Error: {e}"


def html_to_text(html: str) -> str:
    """Extract clean article text from HTML."""
    soup = BeautifulSoup(html, "lxml")
    for tag in soup.find_all(["nav", "footer", "script", "style", "aside", "iframe", "header"]):
        tag.decompose()
    article = soup.find("article") or soup.find("main") or soup.find("body")
    return article.get_text(separator="\n", strip=True) if article else soup.get_text(separator="\n", strip=True)


def extract_images_from_html(html: str, base_url: str = "", max_images: int = 10) -> list[str]:
    """Extract article images from HTML — OG image, meta images, and large article images."""
    soup = BeautifulSoup(html, "lxml")
    images = []
    seen = set()

    # 1. Open Graph image (highest quality, most relevant)
    og = soup.find("meta", property="og:image")
    if og and og.get("content"):
        images.append(og["content"])
        seen.add(og["content"])

    # 2. Twitter card image
    tw = soup.find("meta", attrs={"name": "twitter:image"})
    if tw and tw.get("content") and tw["content"] not in seen:
        images.append(tw["content"])
        seen.add(tw["content"])

    # 3. Large images from article body
    article = soup.find("article") or soup.find("main") or soup.find("body")
    if article:
        for img in article.find_all("img", src=True):
            src = img["src"]
            if src.startswith("data:"):
                continue
            src = urljoin(base_url, src) if base_url else src
            # Skip tiny icons/tracking pixels
            w = img.get("width", "")
            h = img.get("height", "")
            if w and str(w).isdigit() and int(w) < 100:
                continue
            if h and str(h).isdigit() and int(h) < 100:
                continue
            if src not in seen:
                images.append(src)
                seen.add(src)
            if len(images) >= max_images:
                break

    return images


IMG_CACHE_DIR = os.path.join(os.path.dirname(__file__), "outputs", ".img_cache")
os.makedirs(IMG_CACHE_DIR, exist_ok=True)


def download_image(url: str) -> str | None:
    """Download an image and return the local path, or None on failure."""
    try:
        res = requests.get(url, headers={"User-Agent": BROWSER_UA}, timeout=15, stream=True)
        res.raise_for_status()
        content_type = res.headers.get("Content-Type", "")
        if "image" not in content_type and not url.lower().endswith((".jpg", ".jpeg", ".png", ".webp", ".gif")):
            return None
        # Determine extension
        ext = ".jpg"
        if "png" in content_type or url.lower().endswith(".png"):
            ext = ".png"
        elif "webp" in content_type or url.lower().endswith(".webp"):
            ext = ".webp"
        import hashlib
        fname = hashlib.md5(url.encode()).hexdigest() + ext
        path = os.path.join(IMG_CACHE_DIR, fname)
        if not os.path.exists(path):
            with open(path, "wb") as f:
                for chunk in res.iter_content(8192):
                    f.write(chunk)
        return path
    except Exception:
        return None


def normalize_url(url: str) -> str:
    """Ensure URL has a scheme."""
    url = url.strip()
    if url and not url.startswith(("http://", "https://")):
        url = f"https://{url}"
    return url


def resolve_google_news_url(url: str) -> str:
    """Resolve Google News redirect URLs to actual article URLs."""
    if "news.google.com/rss/articles/" not in url:
        return url
    try:
        res = requests.head(url, allow_redirects=True, timeout=10, headers={"User-Agent": BROWSER_UA})
        if res.url and "news.google.com" not in res.url:
            return res.url
        # Sometimes HEAD doesn't resolve, try GET
        res = requests.get(url, allow_redirects=True, timeout=10, headers={"User-Agent": BROWSER_UA})
        return res.url if res.url else url
    except Exception:
        return url


def fetch_url(url: str, session: requests.Session | None = None) -> str:
    """Fetch page content. Tries: 1) direct with session, 2) direct without, 3) Jina Reader."""
    url = normalize_url(url)
    url = resolve_google_news_url(url)

    # Try 1: direct fetch with session cookies
    if session and session.cookies:
        try:
            res = session.get(url, timeout=30)
            res.raise_for_status()
            text = html_to_text(res.text)
            if len(text) > 200:
                return text
        except Exception:
            pass

    # Try 2: direct fetch without cookies (build fresh session for this domain)
    try:
        s = build_session_for_url(url)
        res = s.get(url, timeout=30)
        res.raise_for_status()
        text = html_to_text(res.text)
        if len(text) > 200:
            return text
    except Exception:
        pass

    # Try 3: Jina Reader as last resort
    try:
        res = requests.get(f"https://r.jina.ai/{url}", timeout=30)
        res.raise_for_status()
        return res.text
    except Exception:
        pass

    raise ValueError(f"All fetch methods failed for {url}")


def fetch_article_images(url: str, session: requests.Session | None = None, max_images: int = 3) -> list[str]:
    """Fetch and download article images. Returns list of local file paths."""
    url = normalize_url(url)
    url = resolve_google_news_url(url)
    try:
        s = session or build_session_for_url(url)
        res = s.get(url, timeout=15)
        res.raise_for_status()
        image_urls = extract_images_from_html(res.text, url, max_images=max_images)
        local_paths = []
        for img_url in image_urls:
            path = download_image(img_url)
            if path:
                local_paths.append(path)
            if len(local_paths) >= max_images:
                break
        return local_paths
    except Exception:
        return []


def chat(prompt: str, model: str = "") -> str:
    """Send prompt to LM Studio and return response text."""
    payload = {
        "model": model,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.3,
    }
    try:
        res = requests.post(LM_STUDIO_URL, json=payload, timeout=300)
        res.raise_for_status()
        return res.json()["choices"][0]["message"]["content"]
    except Exception as e:
        raise ValueError(f"LM Studio error: {e}")


def translate_content(content: str, language: str, model: str = "", user_prompt: str = "") -> str:
    """Translate and summarize content."""
    user_instruction = ""
    if user_prompt:
        user_instruction = f"""
The user's original research prompt was: "{user_prompt}"
Consider this context when summarizing — focus on aspects most relevant to the user's interest.
"""
    prompt = f"""
Translate and provide a comprehensive summary of the following content into {language}.
The summary should be detailed and thorough (500-800 words). Cover all key points, data, quotes, and analysis.
Do NOT be brief — include specific numbers, names, dates, and details from the article(s).
{user_instruction}
Content:
{content[:12000]}
"""
    return chat(prompt, model)


def _slide_json_template(lang1: str, lang2: str = "") -> str:
    """Build the JSON example template based on selected languages."""
    if lang2:
        return f"""{{
  "title_primary": "Title in {lang1}",
  "title_secondary": "Title in {lang2}",
  "subtitle_primary": "Subtitle in {lang1}",
  "subtitle_secondary": "Subtitle in {lang2}",
  "slides": [
    {{
      "heading_primary": "Heading in {lang1}",
      "heading_secondary": "Heading in {lang2}",
      "bullets_primary": ["Detailed point with specifics in {lang1}", "Another detailed point with data/numbers in {lang1}", "Third point with analysis in {lang1}", "Fourth point with context in {lang1}", "Fifth point with implications in {lang1}"],
      "bullets_secondary": ["Same point in {lang2}", "Same point in {lang2}", "Same point in {lang2}", "Same point in {lang2}", "Same point in {lang2}"]
    }}
  ]
}}"""
    else:
        return f"""{{
  "title_primary": "Title in {lang1}",
  "subtitle_primary": "Subtitle in {lang1}",
  "slides": [
    {{
      "heading_primary": "Heading in {lang1}",
      "bullets_primary": ["Detailed point with specifics in {lang1}", "Another detailed point with data/numbers in {lang1}", "Third point with analysis in {lang1}", "Fourth point with context in {lang1}", "Fifth point with implications in {lang1}"]
    }}
  ]
}}"""


def generate_slide_structure(content: str, language: str, num_slides: int, model: str = "", source_url: str = "", source_title: str = "", user_prompt: str = "", lang2: str = "") -> dict:
    """Generate structured slide JSON from content — supports single or dual-language."""
    user_instruction = ""
    if user_prompt:
        user_instruction = f"""
The user's original research prompt was: "{user_prompt}"
Use this to guide the presentation angle, emphasis, and what aspects to highlight.
If the prompt contains specific requests about how to present or summarize, follow those instructions.
"""
    if lang2:
        lang_instruction = f"Each slide MUST have BOTH {language} AND {lang2} versions of the heading and bullets."
    else:
        lang_instruction = f"All content must be in {language}."

    prompt = f"""
Create a keynote presentation structure with exactly {num_slides} content slides.
{lang_instruction}
{user_instruction}
IMPORTANT rules for content quality:
- Each slide MUST have 4-6 detailed bullet points
- Each bullet should be a full sentence (15-30 words), not just a short phrase
- Include specific data: numbers, percentages, dollar amounts, dates, company names
- Cover different angles: facts, analysis, market impact, expert quotes, future outlook
- Spread the content across all {num_slides} slides — do NOT front-load everything into the first few slides
- Each slide should have a distinct sub-topic or angle
- If multiple articles are provided (separated by ---), SYNTHESIZE the information from ALL sources, do NOT just repeat one article
- Combine complementary details from different sources to build a comprehensive picture

Based on this content:

{content[:15000]}

Respond ONLY with valid JSON in this exact format, no markdown, no extra text:
{_slide_json_template(language, lang2)}
"""
    raw = chat(prompt, model)
    raw = raw.strip().replace("```json", "").replace("```", "").strip()
    data = json.loads(raw)
    # Attach source info for slides to reference
    if source_url:
        data["source_url"] = source_url
        data["source_title"] = source_title or source_url
    return data


def add_hyperlink(paragraph, url, text, font_size, color, slide_part=None):
    """Add a clickable hyperlink run to a paragraph."""
    from pptx.oxml.ns import qn
    run = paragraph.add_run()
    run.text = text
    run.font.size = font_size
    run.font.color.rgb = color
    run.font.underline = True

    if not slide_part:
        # Walk up to find the slide part
        obj = paragraph._parent
        while obj is not None:
            if hasattr(obj, 'part'):
                slide_part = obj.part
                break
            obj = getattr(obj, '_parent', None) or getattr(obj, 'parent', None)

    if not slide_part:
        return  # Can't create hyperlink without slide part

    # Register the relationship
    rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    rel = slide_part.relate_to(url, rel_type, is_external=True)
    # rel can be a string rId or an object with .rId
    rId = rel if isinstance(rel, str) else rel.rId

    # Create the hyperlink XML on the run
    r_elem = run._r
    hlinkClick = r_elem.makeelement(qn("a:hlinkClick"), {})
    hlinkClick.set(qn("r:id"), rId)
    r_elem.append(hlinkClick)


def build_pptx(data: dict, theme: str, issues: list[str] | None = None, cover_image: str | None = None) -> str:
    """Build .pptx file from slide data — dual-language with source links."""
    themes = {
        "Dark":  {"bg": RGBColor(0x1A, 0x1A, 0x2E), "title": RGBColor(0xE9, 0x4F, 0x37), "text": RGBColor(0xFF, 0xFF, 0xFF), "sub": RGBColor(0xAA, 0xAA, 0xAA), "link": RGBColor(0x4F, 0xC3, 0xF7)},
        "Light": {"bg": RGBColor(0xFA, 0xFA, 0xFA), "title": RGBColor(0x22, 0x22, 0x22),  "text": RGBColor(0x33, 0x33, 0x33), "sub": RGBColor(0x77, 0x77, 0x77), "link": RGBColor(0x1A, 0x73, 0xE8)},
        "Blue":  {"bg": RGBColor(0x0F, 0x3D, 0x66), "title": RGBColor(0x4F, 0xC3, 0xF7), "text": RGBColor(0xFF, 0xFF, 0xFF), "sub": RGBColor(0xAA, 0xCC, 0xDD), "link": RGBColor(0x8B, 0xD9, 0xFF)},
    }
    colors = themes.get(theme, themes["Dark"])

    prs = Presentation()
    prs.slide_width  = Inches(13.33)
    prs.slide_height = Inches(7.5)

    source_url = data.get("source_url", "")
    all_sources = data.get("all_sources", [source_url] if source_url else [])

    def set_bg(slide, color):
        bg = slide.background
        fill = bg.fill
        fill.solid()
        fill.fore_color.rgb = color

    has_secondary = bool(data.get("title_secondary"))

    # --- Title Slide ---
    title_slide = prs.slides.add_slide(prs.slide_layouts[6])
    set_bg(title_slide, colors["bg"])

    # Primary title
    txBox = title_slide.shapes.add_textbox(Inches(1.5), Inches(1.8), Inches(10), Inches(1.2))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = data.get("title_primary", data.get("title_zh", data.get("title", "")))
    p.font.size = Pt(44)
    p.font.bold = True
    p.font.color.rgb = colors["title"]

    # Secondary title
    if has_secondary:
        enBox = title_slide.shapes.add_textbox(Inches(1.5), Inches(3.2), Inches(10), Inches(1))
        entf = enBox.text_frame
        enp = entf.paragraphs[0]
        enp.text = data.get("title_secondary", data.get("title_en", ""))
        enp.font.size = Pt(28)
        enp.font.color.rgb = colors["sub"]

    # Subtitle
    sub_y = 4.5 if has_secondary else 3.2
    subBox = title_slide.shapes.add_textbox(Inches(1.5), Inches(sub_y), Inches(10), Inches(1))
    stf = subBox.text_frame
    sp = stf.paragraphs[0]
    sp.text = data.get("subtitle_primary", data.get("subtitle_zh", data.get("subtitle", "")))
    sp.font.size = Pt(20)
    sp.font.color.rgb = colors["sub"]
    sec_sub = data.get("subtitle_secondary", data.get("subtitle_en", ""))
    if sec_sub:
        sp2 = stf.add_paragraph()
        sp2.text = sec_sub
        sp2.font.size = Pt(18)
        sp2.font.color.rgb = colors["sub"]

    # Cover image on title slide (lower right, below text)
    if cover_image and os.path.exists(cover_image):
        try:
            from PIL import Image
            with Image.open(cover_image) as img:
                img_w, img_h = img.size
            aspect = img_w / img_h if img_h else 1
            max_w, max_h = 3.5, 2.5
            if aspect > max_w / max_h:
                w = min(max_w, img_w / 96)
                h = w / aspect
            else:
                h = min(max_h, img_h / 96)
                w = h * aspect
            # Position: bottom-right corner with padding
            left = Inches(13.33 - w - 0.6)
            top = Inches(7.5 - h - 0.6)
            title_slide.shapes.add_picture(cover_image, left, top, Inches(w), Inches(h))
        except ImportError:
            title_slide.shapes.add_picture(cover_image, Inches(9.0), Inches(4.5), Inches(3.5), Inches(2.2))
        except Exception:
            pass

    # Source links on title slide
    if all_sources:
        src_y = 5.2 if len(all_sources) > 1 else 5.8
        srcBox = title_slide.shapes.add_textbox(Inches(1.5), Inches(src_y), Inches(8), Inches(1.8))
        srctf = srcBox.text_frame
        srctf.word_wrap = True
        for si, src_url in enumerate(all_sources):
            src_domain = urlparse(normalize_url(src_url)).netloc
            sp = srctf.paragraphs[0] if si == 0 else srctf.add_paragraph()
            add_hyperlink(sp, src_url, f"📎 {src_domain}", Pt(11), colors["link"], slide_part=title_slide.part)
            sp.space_after = Pt(2)

    # --- Summary Slide ---
    summary_text = data.get("summary", "")
    if summary_text:
        # Split into lines first, then paginate by line count
        all_lines = [l.strip() for l in summary_text.split("\n") if l.strip()]
        max_lines_per_slide = 14
        summary_pages = []
        for i in range(0, len(all_lines), max_lines_per_slide):
            summary_pages.append(all_lines[i:i + max_lines_per_slide])

        # If only one page but too many chars, split by char count
        if len(summary_pages) == 1 and len(summary_text) > 700:
            summary_pages = []
            remaining = summary_text
            while remaining:
                if len(remaining) <= 700:
                    summary_pages.append(remaining.split("\n"))
                    break
                cut = remaining.rfind("。", 0, 700)
                if cut < 0:
                    cut = remaining.rfind(". ", 0, 700)
                if cut < 0:
                    cut = remaining.rfind("\n", 0, 700)
                if cut < 0:
                    cut = 700
                summary_pages.append([l.strip() for l in remaining[:cut + 1].split("\n") if l.strip()])
                remaining = remaining[cut + 1:].strip()

        for pg_idx, pg_lines in enumerate(summary_pages):
            sum_slide = prs.slides.add_slide(prs.slide_layouts[6])
            set_bg(sum_slide, colors["bg"])

            stitle = "Executive Summary"
            if len(summary_pages) > 1:
                stitle += f" ({pg_idx + 1}/{len(summary_pages)})"
            shBox = sum_slide.shapes.add_textbox(Inches(0.8), Inches(0.3), Inches(11.5), Inches(0.7))
            shtf = shBox.text_frame
            shp = shtf.paragraphs[0]
            shp.text = stitle
            shp.font.size = Pt(28)
            shp.font.bold = True
            shp.font.color.rgb = colors["title"]

            sbBox = sum_slide.shapes.add_textbox(Inches(0.8), Inches(1.1), Inches(11.5), Inches(6.0))
            sbtf = sbBox.text_frame
            sbtf.word_wrap = True
            first = True
            for line in pg_lines:
                if not line:
                    continue
                sp = sbtf.paragraphs[0] if first else sbtf.add_paragraph()
                first = False
                sp.text = line
                sp.font.size = Pt(14)
                sp.font.color.rgb = colors["text"]
                sp.space_after = Pt(3)

    # --- Content Slides ---
    for slide_idx, slide_data in enumerate(data["slides"]):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        set_bg(slide, colors["bg"])

        # Primary heading
        hBox = slide.shapes.add_textbox(Inches(0.8), Inches(0.3), Inches(11.5), Inches(0.7))
        htf = hBox.text_frame
        hp = htf.paragraphs[0]
        hp.text = slide_data.get("heading_primary", slide_data.get("heading_zh", slide_data.get("heading", "")))
        hp.font.size = Pt(32)
        hp.font.bold = True
        hp.font.color.rgb = colors["title"]

        # Secondary heading
        if has_secondary:
            heBox = slide.shapes.add_textbox(Inches(0.8), Inches(1.0), Inches(11.5), Inches(0.5))
            hetf = heBox.text_frame
            hep = hetf.paragraphs[0]
            hep.text = slide_data.get("heading_secondary", slide_data.get("heading_en", ""))
            hep.font.size = Pt(20)
            hep.font.color.rgb = colors["sub"]

        # Bullets
        pri_bullets = slide_data.get("bullets_primary", slide_data.get("bullets_zh", slide_data.get("bullets", [])))
        sec_bullets = slide_data.get("bullets_secondary", slide_data.get("bullets_en", []))

        # Primary bullets (left column if dual, full width if single)
        bullet_top = Inches(1.7) if has_secondary else Inches(1.5)
        bullet_w = Inches(5.5) if has_secondary else Inches(11)
        bBox = slide.shapes.add_textbox(Inches(0.8), bullet_top, bullet_w, Inches(4.5))
        btf = bBox.text_frame
        btf.word_wrap = True
        for i, bullet in enumerate(pri_bullets):
            bp = btf.paragraphs[0] if i == 0 else btf.add_paragraph()
            bp.text = f"• {bullet}"
            bp.font.size = Pt(18)
            bp.font.color.rgb = colors["text"]
            bp.space_after = Pt(6)

        # Secondary bullets (right column, only if dual-language)
        if has_secondary and sec_bullets:
            eBox = slide.shapes.add_textbox(Inches(6.8), bullet_top, Inches(5.5), Inches(4.5))
            etf = eBox.text_frame
            etf.word_wrap = True
            for i, bullet in enumerate(sec_bullets):
                ep = etf.paragraphs[0] if i == 0 else etf.add_paragraph()
                ep.text = f"• {bullet}"
                ep.font.size = Pt(16)
                ep.font.color.rgb = colors["sub"]
                ep.space_after = Pt(6)

        # Source link at bottom — rotate through all sources
        if all_sources:
            src = all_sources[slide_idx % len(all_sources)]
            src_domain = urlparse(normalize_url(src)).netloc
            linkBox = slide.shapes.add_textbox(Inches(0.8), Inches(6.6), Inches(11), Inches(0.4))
            ltf = linkBox.text_frame
            lp = ltf.paragraphs[0]
            add_hyperlink(lp, src, f"📎 {src_domain}", Pt(10), colors["link"], slide_part=slide.part)

    # --- Issues/Warnings Slides (paginated) ---
    if issues:
        issues_per_page = 12
        for page_start in range(0, len(issues), issues_per_page):
            page_issues = issues[page_start:page_start + issues_per_page]
            page_num = page_start // issues_per_page + 1
            total_pages = (len(issues) + issues_per_page - 1) // issues_per_page

            issue_slide = prs.slides.add_slide(prs.slide_layouts[6])
            set_bg(issue_slide, colors["bg"])

            title_text = "⚠️ Issues & Warnings"
            if total_pages > 1:
                title_text += f" ({page_num}/{total_pages})"

            ihBox = issue_slide.shapes.add_textbox(Inches(0.8), Inches(0.3), Inches(11.5), Inches(0.7))
            ihtf = ihBox.text_frame
            ihp = ihtf.paragraphs[0]
            ihp.text = title_text
            ihp.font.size = Pt(28)
            ihp.font.bold = True
            ihp.font.color.rgb = RGBColor(0xE9, 0x4F, 0x37)

            ibBox = issue_slide.shapes.add_textbox(Inches(0.8), Inches(1.2), Inches(11.5), Inches(5.5))
            ibtf = ibBox.text_frame
            ibtf.word_wrap = True
            for i, issue in enumerate(page_issues):
                ip = ibtf.paragraphs[0] if i == 0 else ibtf.add_paragraph()
                ip.text = f"• {issue}"
                ip.font.size = Pt(13)
                ip.font.color.rgb = RGBColor(0xFF, 0xAA, 0x00)
                ip.space_after = Pt(3)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    path = os.path.join(OUTPUT_DIR, f"keynote_{timestamp}.pptx")
    prs.save(path)
    return path


def extract_keywords(prompt: str, model: str = "") -> list[str]:
    """Use AI to break down a prompt into broad search keywords."""
    ai_prompt = f"""
Analyze the following topic and extract 15-25 search keywords for finding related news articles.

Rules:
- Each keyword should be 1-2 words MAX
- Cast a WIDE net — include:
  * Direct terms (e.g. "Iran", "oil", "tariff")
  * Related sectors/industries (e.g. "energy", "defense", "shipping")
  * Key people/companies/indices (e.g. "S&P 500", "OPEC", "Boeing")
  * Broader themes that connect (e.g. "inflation", "sanctions", "geopolitical")
  * Synonyms and alternative phrasings (e.g. "crude" + "petroleum", "conflict" + "war")
- The goal is to catch ALL potentially related articles, not just exact matches

If the topic contains specific instructions about what to focus on, use those to guide which keywords to prioritize.

Return ONLY a JSON array of strings, no markdown, no extra text.

Topic: {prompt}
"""
    raw = chat(ai_prompt, model)
    raw = raw.strip().replace("```json", "").replace("```", "").strip()
    return json.loads(raw)


RSS_FEEDS = {
    "marketwatch.com": [
        "https://www.marketwatch.com/rss/topstories",
        "https://www.marketwatch.com/rss/marketpulse",
    ],
    "wsj.com": [
        "https://feeds.a.dj.com/rss/RSSWorldNews.xml",
        "https://feeds.a.dj.com/rss/WSJcomUSBusiness.xml",
        "https://feeds.a.dj.com/rss/RSSMarketsMain.xml",
    ],
}

# Sites that need Google News RSS as proxy (strong bot detection, no RSS)
# Note: Google News URLs can't be resolved to original URLs, so articles
# from these sites are title-only (no deep-scan)
GOOGLE_NEWS_FALLBACK = {"reuters.com", "investopedia.com"}


def crawl_rss_feeds(feed_urls: list[str], max_links: int = 80) -> list[dict]:
    """Fetch article links from RSS feeds."""
    links = []
    seen = set()
    for feed_url in feed_urls:
        try:
            res = requests.get(feed_url, headers={"User-Agent": BROWSER_UA}, timeout=15)
            res.raise_for_status()
            soup = BeautifulSoup(res.content, "lxml-xml")
            for item in soup.find_all("item"):
                title_tag = item.find("title")
                link_tag = item.find("link")
                if not title_tag or not link_tag:
                    continue
                title = title_tag.get_text(strip=True)
                url = link_tag.get_text(strip=True)
                if url in seen or not title or len(title) < 10:
                    continue
                seen.add(url)
                links.append({"url": url, "title": title})
                if len(links) >= max_links:
                    return links
        except Exception:
            continue
    return links


def crawl_google_news(domain: str, max_links: int = 80) -> list[dict]:
    """Use Google News RSS as proxy for sites that block direct crawl."""
    feed_url = f"https://news.google.com/rss/search?q=site:{domain}&hl=en-US&gl=US&ceid=US:en"
    try:
        res = requests.get(feed_url, headers={"User-Agent": BROWSER_UA}, timeout=15)
        res.raise_for_status()
    except Exception:
        return []

    soup = BeautifulSoup(res.content, "lxml-xml")
    links = []
    seen = set()
    for item in soup.find_all("item"):
        title_tag = item.find("title")
        link_tag = item.find("link")
        if not title_tag or not link_tag:
            continue
        title = title_tag.get_text(strip=True)
        # Clean title — Google appends " - Source Name"
        title = re.sub(r'\s*-\s*[^-]+$', '', title).strip()
        gn_url = link_tag.get_text(strip=True)
        # Resolve Google redirect to actual article URL
        url = resolve_google_news_url(gn_url)
        if url in seen or not title or len(title) < 10:
            continue
        seen.add(url)
        links.append({"url": url, "title": title})
        if len(links) >= max_links:
            break
    return links


# Search engines — when user adds these as a "site", we search instead of crawl
# Note: Google blocks scraping; DuckDuckGo and Bing work reliably
SEARCH_ENGINES = {
    "google.com": "https://www.google.com/search?q={query}&num={num}",
    "bing.com": "https://www.bing.com/search?q={query}&count={num}",
    "duckduckgo.com": "https://html.duckduckgo.com/html/?q={query}",
}

def search_engine_crawl(engine_domain: str, query: str, max_links: int = 80) -> list[dict]:
    """Use a search engine to find articles related to a query."""
    links = []
    seen = set()

    # Build search URL
    template = None
    for domain, url_template in SEARCH_ENGINES.items():
        if domain in engine_domain:
            template = url_template
            break
    if not template:
        return []

    search_url = template.format(query=requests.utils.quote(query), num=min(max_links, 30))

    try:
        headers = {
            "User-Agent": BROWSER_UA,
            "Accept": "text/html,application/xhtml+xml",
            "Accept-Language": "en-US,en;q=0.9",
        }
        res = requests.get(search_url, headers=headers, timeout=15)
        res.raise_for_status()
    except Exception:
        return []

    soup = BeautifulSoup(res.text, "lxml")

    def _extract_real_url(href: str) -> str:
        """Extract real URL from search engine redirect wrappers."""
        # DuckDuckGo wraps: //duckduckgo.com/l/?uddg=https%3A%2F%2F...
        if "uddg=" in href:
            from urllib.parse import unquote, parse_qs
            parsed_q = parse_qs(urlparse(href).query)
            if "uddg" in parsed_q:
                return unquote(parsed_q["uddg"][0])
        # Google wraps: /url?q=https://...
        if "/url?" in href and "q=" in href:
            from urllib.parse import unquote, parse_qs
            parsed_q = parse_qs(urlparse(href).query)
            if "q" in parsed_q:
                return unquote(parsed_q["q"][0])
        return href

    # Google results
    if "google.com" in engine_domain:
        for div in soup.find_all("div", class_="g"):
            a = div.find("a", href=True)
            h3 = div.find("h3")
            if not a or not h3:
                continue
            url = _extract_real_url(a["href"])
            title = h3.get_text(strip=True)
            if not url.startswith("http") or "google.com" in url:
                continue
            if url not in seen and title and len(title) > 5:
                seen.add(url)
                links.append({"url": url, "title": title})
        # Fallback: if no div.g results, try broader search
        if not links:
            for a in soup.find_all("a", href=True):
                url = _extract_real_url(a["href"])
                h3 = a.find("h3")
                if not h3:
                    continue
                title = h3.get_text(strip=True)
                if not url.startswith("http") or "google.com" in url:
                    continue
                if url not in seen and title and len(title) > 5:
                    seen.add(url)
                    links.append({"url": url, "title": title})

    # Bing results
    elif "bing.com" in engine_domain:
        for li in soup.find_all("li", class_="b_algo"):
            a = li.find("a", href=True)
            if not a:
                continue
            url = _extract_real_url(a["href"])
            title = a.get_text(strip=True)
            if not url.startswith("http") or "bing.com" in url:
                continue
            if url not in seen and title and len(title) > 5:
                seen.add(url)
                links.append({"url": url, "title": title})

    # DuckDuckGo results
    elif "duckduckgo.com" in engine_domain:
        for a in soup.find_all("a", class_="result__a"):
            raw_href = a.get("href", "")
            url = _extract_real_url(raw_href)
            title = a.get_text(strip=True)
            if not url.startswith("http") or "duckduckgo.com" in url:
                continue
            if url not in seen and title and len(title) > 5:
                seen.add(url)
                links.append({"url": url, "title": title})

    return links[:max_links]


def is_search_engine(domain: str) -> bool:
    """Check if a domain is a search engine."""
    return any(se in domain for se in SEARCH_ENGINES)


def crawl_site_links(base_url: str, session: requests.Session | None = None, max_links: int = 80) -> list[dict]:
    """Crawl a website. Falls back to RSS/Google News if direct crawl fails."""
    base_url = normalize_url(base_url)
    parsed = urlparse(base_url)
    domain = parsed.netloc.removeprefix("www.")

    try:
        s = session or requests.Session()
        s.headers.setdefault("User-Agent", BROWSER_UA)
        # If session has no cookies, try auto-loading for this domain
        if not s.cookies:
            fresh = build_session_for_url(base_url)
            if fresh.cookies:
                s.cookies.update(fresh.cookies)
        res = s.get(base_url, timeout=15)
        res.raise_for_status()
    except Exception as e:
        # Fallback 1: RSS feeds
        for rss_domain, feeds in RSS_FEEDS.items():
            if rss_domain in domain:
                rss_links = crawl_rss_feeds(feeds, max_links)
                if rss_links:
                    return rss_links
        # Fallback 2: Google News RSS proxy
        for gn_domain in GOOGLE_NEWS_FALLBACK:
            if gn_domain in domain:
                gn_links = crawl_google_news(gn_domain, max_links)
                if gn_links:
                    return gn_links
        raise ValueError(f"Failed to fetch {base_url}: {e}")

    soup = BeautifulSoup(res.text, "lxml")
    parsed_base = urlparse(base_url)
    links = []
    seen = set()

    for a in soup.find_all("a", href=True):
        href = a["href"].strip()
        full_url = urljoin(base_url, href)
        parsed = urlparse(full_url)

        if parsed.netloc != parsed_base.netloc:
            continue
        # Deduplicate by path (ignore query params like ?mod=...)
        canon = f"{parsed.scheme}://{parsed.netloc}{parsed.path}"
        if canon in seen:
            continue
        path_parts = [pt for pt in parsed.path.split("/") if pt]
        if len(path_parts) < 2:
            continue
        skip_paths = {"login", "signup", "subscribe", "account", "search", "video", "podcasts", "#"}
        if any(sk in parsed.path.lower() for sk in skip_paths):
            continue

        title = a.get_text(strip=True)
        if not title or len(title) < 10:
            continue

        seen.add(canon)
        links.append({"url": canon, "title": title})
        if len(links) >= max_links:
            break

    return links


def score_text(text: str, keywords: list[str]) -> float:
    """Score text against keywords (0-1). Matches both full phrases and individual words."""
    if not keywords:
        return 0
    text_lower = text.lower()
    score = 0
    for kw in keywords:
        kw_lower = kw.lower()
        if kw_lower in text_lower:
            # Full keyword match — full point
            score += 1
        else:
            # Partial: check individual words within the keyword
            words = kw_lower.split()
            if len(words) > 1:
                word_hits = sum(1 for w in words if w in text_lower)
                score += word_hits / len(words)
    return score / len(keywords)


def score_article(title: str, keywords: list[str]) -> float:
    """Score an article title against keywords."""
    return score_text(title, keywords)


def ai_filter_titles(titles: list[dict], prompt: str, model: str = "", keywords: list[str] = None, batch_size: int = 60) -> list[dict]:
    """Pre-filter by keywords, then use AI to judge relevance in batches."""
    # Step 1: fast keyword pre-filter to reduce volume
    if keywords and len(titles) > batch_size:
        scored = []
        for t in titles:
            s = score_text(t["title"], keywords)
            if s > 0:
                scored.append((s, t))
        scored.sort(key=lambda x: x[0], reverse=True)
        titles = [t for _, t in scored[:batch_size * 2]]  # keep top candidates

    # Step 2: batch AI filtering
    all_results = []
    for batch_start in range(0, len(titles), batch_size):
        batch = titles[batch_start:batch_start + batch_size]
        batch_results = _ai_filter_batch(batch, prompt, model)
        all_results.extend(batch_results)

    return all_results


def _ai_filter_batch(titles: list[dict], prompt: str, model: str = "") -> list[dict]:
    """AI judges a single batch of titles."""
    title_list = "\n".join(f"{i}. {t['title']}" for i, t in enumerate(titles))
    ai_prompt = f"""
You are filtering news article titles for relevance to a research topic.

Topic: {prompt}

Article titles:
{title_list}

For each title, decide if it is relevant to the topic.
Return ONLY a JSON array of objects for the RELEVANT titles, no markdown, no extra text.
Each object: {{"index": <number>, "relevance": <0.0-1.0>, "reason": "<brief reason>"}}

Example: [{{"index": 3, "relevance": 0.9, "reason": "Directly about oil prices"}}]

If none are relevant, return an empty array: []
"""
    raw = chat(ai_prompt, model)
    raw = raw.strip().replace("```json", "").replace("```", "").strip()
    try:
        judgments = json.loads(raw)
    except json.JSONDecodeError:
        return []

    results = []
    for j in judgments:
        idx = j.get("index", -1)
        if 0 <= idx < len(titles):
            results.append({
                **titles[idx],
                "score": j.get("relevance", 0.5),
                "reason": j.get("reason", ""),
            })
    return results


def ai_score_article(url: str, prompt: str, session: requests.Session | None = None, model: str = "") -> dict:
    """Fetch article content and use AI to judge relevance to user's topic."""
    try:
        content = fetch_url(url, session)[:4000]
    except Exception:
        return {"score": 0, "reason": "Failed to fetch"}

    ai_prompt = f"""
You are evaluating whether a news article is relevant to a research topic.

Research topic: {prompt}

Article content (excerpt):
{content}

Rate the relevance from 0.0 to 1.0 and explain briefly why.
Return ONLY valid JSON, no markdown:
{{"relevance": <0.0-1.0>, "reason": "<1 sentence>"}}
"""
    try:
        raw = chat(ai_prompt, model)
        raw = raw.strip().replace("```json", "").replace("```", "").strip()
        result = json.loads(raw)
        return {"score": result.get("relevance", 0), "reason": result.get("reason", "")}
    except Exception:
        return {"score": 0, "reason": "AI scoring failed"}


def build_session_for_url(url: str, cookies_text: str = "") -> requests.Session:
    """Build a session with cookies specific to the URL's domain."""
    session = requests.Session()
    session.headers.update({"User-Agent": BROWSER_UA})
    # Increase header limit for sites like Yahoo that send excessive Set-Cookie headers
    import http.client
    http.client._MAXHEADERS = 1000

    # If manual cookies provided, use those
    if cookies_text and cookies_text.strip():
        for pair in cookies_text.split(";"):
            pair = pair.strip()
            if "=" in pair:
                name, value = pair.split("=", 1)
                session.cookies.set(name.strip(), value.strip())
        return session

    # Auto-load cookies for this specific domain
    parsed = urlparse(url if "://" in url else f"https://{url}")
    domain = parsed.netloc.removeprefix("www.")
    if domain:
        try:
            result = load_browser_cookies("Chrome", domain)
            if not result.startswith("Error:"):
                for pair in result.split(";"):
                    pair = pair.strip()
                    if "=" in pair:
                        name, value = pair.split("=", 1)
                        session.cookies.set(name.strip(), value.strip())
        except Exception:
            pass  # Cookie loading failed silently, proceed without
    return session


_research_issues = []  # Shared between research and pipeline
_research_results = []  # Store latest results for topic grouping

def run_research(prompt, site_urls, max_articles, model, cookies_text):
    """Research pipeline: keywords → crawl multiple sites → AI filters titles → deep scan."""
    global _research_issues
    _research_issues = []
    logs = []
    no_update = gr.update()

    # Normalize site_urls to a list
    if isinstance(site_urls, str):
        url_list = [u.strip() for u in site_urls.split(";") if u.strip()]
    elif isinstance(site_urls, list):
        url_list = [u.strip() for u in site_urls if u and u.strip()]
    else:
        url_list = []

    if not url_list:
        yield "❌ No site URLs provided", "", "", no_update
        return

    # Save sites for future use
    save_sites(url_list)

    try:
        logs.append("🔍 Analyzing prompt with AI...")
        yield "\n".join(logs), "", "", no_update

        keywords = extract_keywords(prompt, model)
        kw_display = ", ".join(keywords)
        logs.append(f"🔑 Keywords: {kw_display}")
        yield "\n".join(logs), kw_display, "", no_update

        # Crawl all sites — each with its own session/cookies
        all_links = []
        search_query = " ".join(keywords[:8])  # Use top keywords as search query
        for site_url in url_list:
            site_domain = urlparse(normalize_url(site_url)).netloc.removeprefix("www.")

            # Detect search engines — search instead of crawl
            if is_search_engine(site_domain):
                logs.append(f"🔎 Searching {site_domain} for: {search_query[:50]}...")
                yield "\n".join(logs), kw_display, "", no_update
                try:
                    site_links = search_engine_crawl(site_domain, search_query, max_links=30)
                    logs.append(f"  📄 Found {len(site_links)} results from {site_domain}")
                    all_links.extend(site_links)
                except Exception as e:
                    logs.append(f"  ⚠️ Search failed: {e}")
                    _research_issues.append(f"Search {site_domain}: {e}")
            else:
                session = build_session_for_url(site_url, cookies_text)
                if session.cookies:
                    logs.append(f"🔐 Loaded {len(session.cookies)} cookies for {urlparse(normalize_url(site_url)).netloc}")
                logs.append(f"🕷️ Crawling {site_url}...")
                yield "\n".join(logs), kw_display, "", no_update

                try:
                    site_links = crawl_site_links(site_url, session=session, max_links=80)
                    logs.append(f"  📄 Found {len(site_links)} links from {urlparse(normalize_url(site_url)).netloc}")
                    all_links.extend(site_links)
                except Exception as e:
                    logs.append(f"  ⚠️ Failed: {e}")
                    _research_issues.append(f"Crawl {urlparse(normalize_url(site_url)).netloc}: {e}")

            yield "\n".join(logs), kw_display, "", no_update

        links = all_links
        logs.append(f"📄 Total: {len(links)} article links from {len(url_list)} site(s)")
        yield "\n".join(logs), kw_display, "", no_update

        # Phase 1: AI judges which titles are relevant
        logs.append("🧠 AI filtering article titles...")
        yield "\n".join(logs), kw_display, "", no_update

        candidates = ai_filter_titles(links, prompt, model, keywords=keywords)
        candidates.sort(key=lambda x: x["score"], reverse=True)

        if not candidates:
            logs.append("📊 AI found no relevant titles")
            yield "\n".join(logs), kw_display, "No matching articles found. Try a different topic or site.", gr.update(choices=[], value="")
            return

        logs.append(f"📊 AI selected {len(candidates)} relevant articles")
        for c in candidates:
            logs.append(f"  → [{int(c['score']*100)}%] {c['title'][:60]} — {c['reason']}")
        yield "\n".join(logs), kw_display, "", no_update

        # Phase 2: AI reads top candidates and scores relevance
        top = candidates[:max_articles]
        logs.append(f"📖 AI deep-reading top {len(top)} of {len(candidates)} articles...")
        yield "\n".join(logs), kw_display, "", no_update

        results = []
        for i, article in enumerate(top):
            logs.append(f"  📖 [{i+1}/{len(top)}] {article['title'][:50]}...")
            yield "\n".join(logs), kw_display, "", no_update

            art_session = build_session_for_url(article["url"], cookies_text)
            ai_result = ai_score_article(article["url"], prompt, art_session, model)
            combined = (article["score"] * 0.4) + (ai_result["score"] * 0.6)
            reason = ai_result.get("reason") or article.get("reason", "")
            results.append({**article, "score": combined, "reason": reason})

        results.sort(key=lambda x: x["score"], reverse=True)
        global _research_results
        _research_results = results

        # Phase 3: AI cross-checks articles to group related topics
        if len(results) > 1:
            logs.append("🔗 AI cross-checking articles for related topics...")
            yield "\n".join(logs), kw_display, "", no_update

            try:
                titles_for_grouping = "\n".join(f"{i}. {r['title']}" for i, r in enumerate(results))
                group_prompt = f"""
You have a list of news articles found for the topic: "{prompt}"

Articles:
{titles_for_grouping}

Group these articles by sub-topic/theme. Articles covering the same event or angle should be grouped together.
Return ONLY valid JSON, no markdown:
[
  {{"topic": "Brief topic name", "articles": [0, 2, 5], "summary": "What these articles share"}}
]
"""
                raw = chat(group_prompt, model)
                raw = raw.strip().replace("```json", "").replace("```", "").strip()
                topic_groups = json.loads(raw)
                logs.append(f"📊 Found {len(topic_groups)} topic groups:")
                for tg in topic_groups:
                    article_count = len(tg.get("articles", []))
                    logs.append(f"  • {tg['topic']} ({article_count} articles) — {tg.get('summary', '')[:60]}")
                # Tag each result with its topic group
                for tg in topic_groups:
                    for idx in tg.get("articles", []):
                        if 0 <= idx < len(results):
                            results[idx]["topic_group"] = tg["topic"]
            except Exception as e:
                logs.append(f"  ⚠️ Grouping failed: {e}")

            yield "\n".join(logs), kw_display, "", no_update

        # Format output + build dropdown choices
        output_lines = []
        url_choices = []
        current_topic = None
        for i, r in enumerate(results, 1):
            pct = int(r["score"] * 100)
            reason = f" — {r['reason']}" if r.get("reason") else ""
            topic = r.get("topic_group", "")
            if topic and topic != current_topic:
                output_lines.append(f"\n── {topic} ──")
                current_topic = topic
            output_lines.append(f"{i}. [{pct}% match] {r['title']}{reason}\n   {r['url']}")
            url_choices.append(f"{r['title'][:60]}  |  {r['url']}")

        result_text = "\n".join(output_lines)
        logs.append(f"\n✅ Found {len(results)} relevant articles")
        top_choice = url_choices[0] if url_choices else ""
        yield "\n".join(logs), kw_display, result_text, gr.update(choices=url_choices, value=top_choice)

    except Exception as e:
        logs.append(f"❌ Error: {e}")
        yield "\n".join(logs), "", "", no_update


def run_pipeline(url, lang1, lang2, num_slides, theme, open_keynote, model, cookies_text, user_prompt="", extra_urls=None):
    """Generate keynote from one or more article URLs. If extra_urls provided, combines content."""
    logs = []
    issues = list(_research_issues)  # Carry over any crawl issues
    lang2 = lang2 if lang2 and lang2 != "None" else ""

    # Combine primary URL with any extra URLs from same topic group
    all_urls = [url] + (extra_urls or [])
    # Deduplicate and skip Google News redirect URLs (can't be fetched)
    seen = set()
    unique_urls = []
    for u in all_urls:
        if u not in seen and "news.google.com/rss/articles/" not in u:
            seen.add(u)
            unique_urls.append(u)

    try:
        lang_label = f"{lang1} + {lang2}" if lang2 else lang1
        if user_prompt:
            logs.append(f"📝 User prompt: {user_prompt[:80]}...")
        logs.append(f"🌍 Languages: {lang_label}")
        if len(unique_urls) > 1:
            logs.append(f"📰 Combining {len(unique_urls)} related articles")

        # Fetch all article content
        all_content = []
        fetched_sources = []  # Track successfully fetched article URLs
        for i, article_url in enumerate(unique_urls):
            logs.append(f"📥 Fetching [{i+1}/{len(unique_urls)}] {urlparse(normalize_url(article_url)).path[:50]}...")
            yield "\n".join(logs), None, ""

            session = build_session_for_url(article_url, cookies_text)
            try:
                text = fetch_url(article_url, session)
                logs.append(f"  ✅ {len(text)} characters")
                all_content.append(text)
                fetched_sources.append(article_url)
            except Exception as e:
                issues.append(f"Fetch {article_url}: {e}")
                logs.append(f"  ⚠️ {e}")
                try:
                    text = requests.get(f"https://r.jina.ai/{article_url}", timeout=30).text
                    logs.append(f"  ✅ Fallback: {len(text)} characters")
                    all_content.append(text)
                    fetched_sources.append(article_url)
                except Exception as e2:
                    issues.append(f"Fallback {article_url}: {e2}")
                    logs.append(f"  ⚠️ Fallback also failed")
            yield "\n".join(logs), None, ""

        content = "\n\n---\n\n".join(all_content) if all_content else f"[No content fetched from {url}]"

        # Fetch 1 representative image from the primary article
        logs.append("🖼️ Looking for article image...")
        yield "\n".join(logs), None, ""
        article_image = None
        for article_url in unique_urls:
            try:
                session = build_session_for_url(article_url, cookies_text)
                imgs = fetch_article_images(article_url, session, max_images=1)
                if imgs:
                    article_image = imgs[0]
                    logs.append(f"  ✅ Found image from {urlparse(normalize_url(article_url)).netloc}")
                    break
            except Exception:
                continue
        if not article_image:
            logs.append("  ℹ️ No suitable image found")
        yield "\n".join(logs), None, ""

        logs.append("🌐 Translating & summarizing...")
        yield "\n".join(logs), None, ""

        try:
            summary = translate_content(content, lang1, model, user_prompt=user_prompt)
            logs.append("✅ Translation done")
        except Exception as e:
            issues.append(f"Translation error: {e}")
            summary = content[:500]
            logs.append(f"⚠️ Translation failed, using raw content excerpt")
        yield "\n".join(logs), None, summary

        logs.append(f"🧠 Generating {num_slides} slides...")
        yield "\n".join(logs), None, summary

        try:
            data = generate_slide_structure(content, lang1, num_slides, model, source_url=url, source_title=url, user_prompt=user_prompt, lang2=lang2)
            # Attach all fetched sources for slide-level attribution
            data["all_sources"] = fetched_sources
            logs.append(f"✅ Structure ready: {len(data['slides'])} slides")
        except Exception as e:
            issues.append(f"Slide generation error: {e}")
            logs.append(f"⚠️ Slide generation failed: {e}")
            yield "\n".join(logs), None, summary
            return
        yield "\n".join(logs), None, summary

        if issues:
            logs.append(f"⚠️ {len(issues)} issue(s) will be noted in the final slide")

        logs.append("📊 Building PPTX...")
        yield "\n".join(logs), None, summary

        data["summary"] = summary
        pptx_path = build_pptx(data, theme, issues=issues, cover_image=article_image)
        logs.append(f"✅ Saved: {os.path.basename(pptx_path)}")
        yield "\n".join(logs), None, summary

        # Export to PDF via Keynote
        logs.append("📄 Exporting to PDF...")
        yield "\n".join(logs), None, summary

        pdf_path = pptx_path.replace(".pptx", ".pdf")
        try:
            export_script = f'''
            tell application "Keynote"
                set theDoc to open POSIX file "{pptx_path}"
                delay 2
                export theDoc to POSIX file "{pdf_path}" as PDF
                close theDoc saving no
            end tell
            '''
            subprocess.run(["osascript", "-e", export_script], timeout=30, check=True)
            logs.append(f"✅ PDF: {os.path.basename(pdf_path)}")
            path = pdf_path
        except Exception as e:
            issues.append(f"PDF export failed: {e}")
            logs.append(f"⚠️ PDF export failed, using PPTX: {e}")
            path = pptx_path

        if open_keynote:
            subprocess.run(["open", path])
            logs.append(f"🎬 Opened {os.path.basename(path)}")

        yield "\n".join(logs), path, summary

    except Exception as e:
        logs.append(f"❌ Error: {e}")
        yield "\n".join(logs), None, ""


# --- UI ---
css = """
/* === Chrome Extension popup — compact dark === */
*, *::before, *::after { box-sizing: border-box; }

body {
    font-family: 'Segoe UI', 'SF Pro Text', -apple-system, Roboto, sans-serif;
    background: #1a1a1a;
}

/* ── Container ── */
.gradio-container {
    max-width: 900px !important;
    margin: 0 auto !important;
    padding: 0 12px 12px !important;
    background: #2b2b2b;
    border-left: 1px solid #3a3a3a;
    border-right: 1px solid #3a3a3a;
    box-shadow: 0 0 40px rgba(0,0,0,0.5);
}

/* ── Kill Gradio's default spacing ── */
.gradio-container > .flex,
.gradio-container .form,
.gradio-container .contain,
.gradio-container .panel {
    gap: 5px !important;
    padding: 0 !important;
    margin: 0 !important;
}
.gradio-container .panel,
.gradio-container .gr-box {
    background: none !important;
    border: none !important;
    box-shadow: none !important;
}
.gradio-container .row { gap: 8px !important; align-items: end !important; }

/* ── Header bar ── */
.ext-header {
    background: linear-gradient(135deg, #1a73e8, #1558b0);
    padding: 12px 16px !important;
    margin: 0 -12px 10px !important;
    border-bottom: 1px solid #1558b0;
}
.ext-header h3, .ext-header p { margin: 0 !important; color: #fff !important; }
.ext-header h3 { font-size: 14px !important; font-weight: 600 !important; }
.ext-header p { font-size: 10.5px !important; opacity: 0.8; margin-top: 1px !important; }

/* ── Labels — plain text, no background ── */
.gradio-container label,
.gradio-container .label-wrap,
.gradio-container span[data-testid="block-info"],
.gradio-container span.text-gray-500,
.gradio-container .gradio-label,
.gradio-container label span {
    font-size: 10px !important;
    font-weight: 600 !important;
    color: #888 !important;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    padding: 0 !important;
    margin: 0 0 1px !important;
    background: none !important;
    background-color: transparent !important;
    border: none !important;
    box-shadow: none !important;
    line-height: 1.4 !important;
}
/* Checkbox/radio labels — normal case, readable size */
.gradio-container label:has(input[type="checkbox"]) span,
.gradio-container label:has(input[type="radio"]) span {
    font-size: 12px !important;
    text-transform: none !important;
    letter-spacing: 0 !important;
    color: #ccc !important;
}

/* ── Inputs ── */
.gradio-container input[type="text"],
.gradio-container textarea,
.gradio-container select,
.gradio-container .wrap input,
.gradio-container .secondary-wrap {
    background: #333 !important;
    border: 1px solid #444 !important;
    color: #e0e0e0 !important;
    border-radius: 5px !important;
    font-size: 12.5px !important;
    padding: 7px 9px !important;
    transition: border-color 0.15s;
}
.gradio-container input[type="text"]:focus,
.gradio-container textarea:focus {
    border-color: #1a73e8 !important;
    outline: none !important;
    box-shadow: 0 0 0 2px rgba(26,115,232,0.2) !important;
}

/* ── Dropdown overrides ── */
.gradio-container .wrap.svelte-aqlk7e,
.gradio-container .wrap[data-testid] {
    background: #333 !important;
    border: 1px solid #444 !important;
    border-radius: 5px !important;
    min-height: 32px !important;
    padding: 0 8px !important;
}

/* ── Tabs ── */
.ext-tabs > .tab-nav {
    background: #252525 !important;
    border-bottom: 1px solid #3a3a3a !important;
    gap: 0 !important;
    padding: 0 !important;
    margin: 0 -12px 6px !important;
}
.ext-tabs > .tab-nav button {
    background: transparent !important;
    color: #777 !important;
    border: none !important;
    border-bottom: 2px solid transparent !important;
    font-size: 11.5px !important;
    font-weight: 600 !important;
    padding: 8px 14px !important;
    margin: 0 !important;
    transition: all 0.15s;
}
.ext-tabs > .tab-nav button.selected {
    color: #e0e0e0 !important;
    border-bottom-color: #1a73e8 !important;
    background: rgba(26,115,232,0.08) !important;
}
.ext-tabs > .tab-nav button:hover:not(.selected) {
    color: #bbb !important;
    background: rgba(255,255,255,0.04) !important;
}
.ext-tabs .tabitem { padding: 6px 0 0 !important; }

/* ── Buttons ── */
.ext-generate-btn {
    background: linear-gradient(135deg, #1a73e8, #1558b0) !important;
    color: #fff !important;
    border: none !important;
    border-radius: 6px !important;
    font-size: 12.5px !important;
    font-weight: 600 !important;
    padding: 9px 0 !important;
    margin-top: 4px !important;
    cursor: pointer;
    transition: opacity 0.15s, transform 0.1s;
}
.ext-generate-btn:hover { opacity: 0.9 !important; }
.ext-generate-btn:active { transform: scale(0.98) !important; }

.ext-refresh-btn {
    background: #333 !important;
    color: #aaa !important;
    border: 1px solid #444 !important;
    border-radius: 5px !important;
    font-size: 13px !important;
    padding: 2px 8px !important;
    min-width: 36px !important;
    max-height: 32px !important;
    cursor: pointer;
}
.ext-refresh-btn:hover { background: #3e3e3e !important; }

/* ── Slider compact ── */
.gradio-container input[type="range"] { accent-color: #1a73e8; }
.gradio-container input[type="number"] {
    background: #333 !important; border: 1px solid #444 !important;
    color: #ddd !important; border-radius: 4px !important;
    width: 44px !important; font-size: 11px !important;
    padding: 2px 4px !important;
}

/* ── Radio / checkbox — fully restore ── */
.gradio-container input[type="radio"],
.gradio-container input[type="checkbox"] {
    accent-color: #1a73e8;
    width: 16px !important;
    height: 16px !important;
    min-width: 16px !important;
    margin: 0 !important;
    padding: 0 !important;
    cursor: pointer !important;
    pointer-events: auto !important;
    position: static !important;
    opacity: 1 !important;
    flex-shrink: 0 !important;
    -webkit-appearance: auto !important;
    appearance: auto !important;
}
/* Every ancestor of a checkbox/radio must allow clicks and be visible */
.gradio-container *:has(input[type="checkbox"]),
.gradio-container *:has(input[type="radio"]) {
    pointer-events: auto !important;
    overflow: visible !important;
    max-height: none !important;
    height: auto !important;
    min-height: 0 !important;
}
.gradio-container label:has(input[type="checkbox"]),
.gradio-container label:has(input[type="radio"]) {
    display: flex !important;
    align-items: center !important;
    gap: 6px !important;
    padding: 4px 8px !important;
    cursor: pointer !important;
}

/* ── Section divider ── */
.ext-section {
    border-top: 1px solid #333;
    margin: 6px 0 4px !important;
    padding: 0 !important;
}

/* ── Progress log ── */
@keyframes border-pulse {
    0%   { box-shadow: 0 0 0 0 rgba(139,195,74,0), inset 0 0 0 0 rgba(139,195,74,0); }
    50%  { box-shadow: 0 0 10px 2px rgba(139,195,74,0.35), inset 0 0 4px 0 rgba(139,195,74,0.1); }
    100% { box-shadow: 0 0 0 0 rgba(139,195,74,0), inset 0 0 0 0 rgba(139,195,74,0); }
}
.ext-progress {
    position: relative !important;
    border-radius: 5px !important;
}
.ext-progress textarea {
    background: #1a1a1a !important;
    color: #8bc34a !important;
    font-family: 'SF Mono', 'Cascadia Code', 'Fira Code', monospace !important;
    font-size: 10.5px !important;
    line-height: 1.5 !important;
    border: 1.5px solid #2a2a2a !important;
    border-radius: 5px !important;
    padding: 6px 8px !important;
}
/* Flash only the outer glow, not the text */
.ext-progress.ext-active {
    animation: border-pulse 1.5s ease-in-out infinite !important;
    border-radius: 5px !important;
}

/* ── Summary / results ── */
.ext-summary textarea {
    background: #1a1a1a !important;
    color: #ccc !important;
    font-size: 11.5px !important;
    line-height: 1.5 !important;
    border: 1px solid #2a2a2a !important;
    border-radius: 5px !important;
    padding: 6px 8px !important;
}

/* ── File download ── */
.gradio-container .file-preview,
.gradio-container .upload-button {
    background: #333 !important;
    border: 1px dashed #444 !important;
    border-radius: 5px !important;
    color: #aaa !important;
    font-size: 11px !important;
}

/* ── Scrollbar ── */
::-webkit-scrollbar { width: 5px; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-thumb { background: #444; border-radius: 3px; }
::-webkit-scrollbar-thumb:hover { background: #666; }

/* ── Accordion (cookies) ── */
.ext-accordion {
    border: 1px solid #3a3a3a !important;
    border-radius: 6px !important;
    background: #252525 !important;
    margin: 4px 0 !important;
    padding: 0 !important;
}
.ext-accordion > .label-wrap {
    padding: 6px 10px !important;
    background: none !important;
    color: #999 !important;
    font-size: 11px !important;
}
.ext-accordion > .label-wrap span { color: #999 !important; font-size: 11px !important; }
.ext-accordion .prose { padding: 0 !important; }
.ext-accordion > div:not(.label-wrap) {
    padding: 4px 10px 8px !important;
    gap: 4px !important;
}
.ext-accordion > div:not(.label-wrap) > div { gap: 4px !important; padding: 0 !important; margin: 0 !important; }
.ext-accordion > div:not(.label-wrap) > div > div { gap: 2px !important; padding: 0 !important; margin: 0 !important; }
.ext-accordion .row { gap: 6px !important; }
.ext-accordion button { margin: 0 !important; padding: 7px 0 !important; }
.ext-accordion label {
    background: none !important; border: none !important; box-shadow: none !important;
    padding: 0 !important; margin: 0 0 1px !important;
}

/* ── Site URL tags ── */
.ext-site-tags .token {
    background: #1a73e8 !important;
    color: #fff !important;
    border-radius: 12px !important;
    padding: 2px 10px !important;
    font-size: 11px !important;
    margin: 2px !important;
}
.ext-site-tags .token-remove {
    color: #fff !important;
    margin-left: 4px !important;
    cursor: pointer !important;
}
.ext-site-tags .secondary-wrap {
    min-height: 36px !important;
    flex-wrap: wrap !important;
    gap: 3px !important;
    padding: 4px 8px !important;
}
/* Hide the clear-all × and dropdown arrow */
.ext-site-tags .icon-wrap,
.ext-site-tags button[aria-label="Clear"],
.ext-site-tags .clear-btn,
.ext-site-tags > div > div > button:not(.token-remove),
.ext-site-tags svg.icon-clear {
    display: none !important;
}

/* ── Hide footer ── */
footer { display: none !important; }
"""

available_models = get_available_models()
default_model = pick_default_model(available_models)

progress_js = """
() => {
    // Watch progress textareas and toggle 'active' class on content change
    const observer = new MutationObserver(() => {
        document.querySelectorAll('.ext-progress textarea').forEach(ta => {
            const wrapper = ta.closest('.ext-progress');
            if (!wrapper) return;
            if (ta.value && ta.value.trim()) {
                wrapper.classList.add('ext-active');
                // Clear existing timer
                clearTimeout(wrapper._pulseTimer);
                // Stop pulsing after 3s of no changes
                wrapper._pulseTimer = setTimeout(() => {
                    wrapper.classList.remove('ext-active');
                }, 3000);
            }
        });
    });
    setTimeout(() => {
        observer.observe(document.body, { childList: true, subtree: true, characterData: true });
    }, 1000);
}
"""

with gr.Blocks(title="Web AI Tool", theme=gr.themes.Soft(), css=css, js=progress_js) as app:
    gr.HTML("""
        <div class="ext-header">
            <h3>🌐 Web AI Tool</h3>
            <p>Research · Translate · Export to Keynote</p>
        </div>
    """)

    # Shared controls at the top
    with gr.Row():
        model_input = gr.Dropdown(
            choices=available_models,
            value=default_model,
            label="Model",
            allow_custom_value=True,
            scale=4,
        )
        refresh_btn = gr.Button("↻", elem_classes=["ext-refresh-btn"], scale=1, min_width=40)

    saved_sites = load_saved_sites()
    sites_history = load_sites_history()
    # Merge: history has all ever-used, saved_sites are currently active
    all_choices = list(dict.fromkeys(sites_history + saved_sites))  # dedupe, preserve order
    prefs = load_saved_prefs()

    research_site = gr.Dropdown(
        label="Site URLs (type and press Enter to add)",
        choices=all_choices,
        value=saved_sites,
        multiselect=True,
        allow_custom_value=True,
        elem_classes=["ext-site-tags"],
    )

    with gr.Accordion("🔐 Session / Cookies", open=False, elem_classes=["ext-accordion"]):
        with gr.Row():
            cookie_browser = gr.Dropdown(
                choices=["Chrome", "Firefox", "Safari", "Edge"],
                value="Chrome",
                label="Browser",
                scale=3,
            )
            cookie_load_btn = gr.Button("Load Cookies", elem_classes=["ext-generate-btn"], scale=4)
        cookies_input = gr.Textbox(
            label="Cookies",
            placeholder="Auto-filled from Site URL, or paste manually (name=val; name2=val2)",
            lines=2,
        )
        cookie_status = gr.HTML('<p style="font-size:10px;color:#666;margin:2px 0 0;">Loads cookies for the domain in Site URL above.</p>')

    with gr.Tabs(elem_classes=["ext-tabs"]):

        # ── Research Tab ──
        with gr.TabItem("🔍 Research"):
            with gr.Row():
                with gr.Column(scale=1):
                    research_prompt = gr.Textbox(
                        label="Topic / Prompt",
                        placeholder="e.g. Impact of AI regulation on tech companies",
                        lines=3,
                    )
                    research_max = gr.Number(value=prefs["max_articles"], label="Articles to deep-scan", precision=0, minimum=1)
                    auto_keynote = gr.Checkbox(value=prefs["auto_keynote"], label="Auto-generate Keynote from top result")
                    research_btn = gr.Button("Search Articles", variant="primary", elem_classes=["ext-generate-btn"])
                with gr.Column(scale=1):
                    research_log = gr.Textbox(label="Progress", lines=5, interactive=False, elem_classes=["ext-progress"])
                    research_kw  = gr.Textbox(label="Extracted Keywords", lines=1, interactive=False)
                    research_out = gr.Textbox(label="Matching Articles", lines=10, interactive=False, elem_classes=["ext-summary"])

            gr.HTML('<div class="ext-section"></div>')
            article_picker = gr.Dropdown(label="Select article for Keynote", choices=[], interactive=True)
            with gr.Row():
                pick_lang1  = gr.Dropdown(["zh-TW", "English", "Japanese", "Korean", "Spanish", "French", "German"], value=prefs["lang1"], label="Primary Lang", scale=2)
                pick_lang2  = gr.Dropdown(["English", "zh-TW", "Japanese", "Korean", "Spanish", "French", "German", "None"], value=prefs["lang2"], label="Secondary Lang", scale=2)
                pick_slides = gr.Number(value=prefs["slides"], label="Slides", precision=0, minimum=1, scale=1)
                pick_theme  = gr.Radio(["Dark", "Light", "Blue"], value=prefs["theme"], label="Theme", scale=2)
            generate_from_research = gr.Button("Generate Keynote from Article", variant="primary", elem_classes=["ext-generate-btn"])
            pick_log     = gr.Textbox(label="Keynote Progress", lines=5, interactive=False, elem_classes=["ext-progress"])
            pick_file    = gr.File(label="Download")
            pick_summary = gr.Textbox(label="Summary", lines=4, interactive=False, elem_classes=["ext-summary"])

        # ── Keynote Tab ──
        with gr.TabItem("📊 Keynote"):
            with gr.Row():
                with gr.Column(scale=1):
                    url_input = gr.Textbox(label="URL", placeholder="https://example.com/article", lines=1)
                    with gr.Row():
                        lang1      = gr.Dropdown(["zh-TW", "English", "Japanese", "Korean", "Spanish", "French", "German"], value=prefs["lang1"], label="Primary Lang", scale=2)
                        lang2      = gr.Dropdown(["English", "zh-TW", "Japanese", "Korean", "Spanish", "French", "German", "None"], value=prefs["lang2"], label="Secondary Lang", scale=2)
                        num_slides = gr.Number(value=prefs["slides"], label="Slides", precision=0, minimum=1, scale=1)
                    with gr.Row():
                        theme        = gr.Radio(["Dark", "Light", "Blue"], value=prefs["theme"], label="Theme", scale=2)
                        open_keynote = gr.Checkbox(value=True, label="Auto-open file", scale=1)
                    btn = gr.Button("Generate Keynote", variant="primary", elem_classes=["ext-generate-btn"])
                with gr.Column(scale=1):
                    log_output     = gr.Textbox(label="Progress", lines=6, interactive=False, elem_classes=["ext-progress"])
                    file_output    = gr.File(label="Download")
                    summary_output = gr.Textbox(label="Summary", lines=5, interactive=False, elem_classes=["ext-summary"])

    # ── Event handlers ──
    def refresh_models():
        models = get_available_models()
        default = pick_default_model(models)
        return gr.update(choices=models, value=default)

    def on_load_cookies(browser, site_urls):
        # Support both list (multiselect) and string
        if isinstance(site_urls, list):
            url_list = [u.strip() for u in site_urls if u and u.strip()]
        else:
            url_list = [u.strip() for u in (site_urls or "").split(";") if u.strip()]
        if not url_list:
            return "", '<p style="font-size:10px;color:#e74c3c;">Add at least one Site URL first.</p>'
        all_cookies = []
        domains_loaded = []
        for raw_url in url_list:
            parsed = urlparse(raw_url if "://" in raw_url else f"https://{raw_url}")
            domain = parsed.netloc.removeprefix("www.")
            if not domain:
                continue
            result = load_browser_cookies(browser, domain)
            if not result.startswith("Error:"):
                all_cookies.append(result)
                domains_loaded.append(domain)
        if not all_cookies:
            return "", '<p style="font-size:10px;color:#e74c3c;">No cookies found for any domain.</p>'
        combined = "; ".join(all_cookies)
        count = combined.count("=")
        domains_str = ", ".join(domains_loaded)
        return combined, f'<p style="font-size:10px;color:#8bc34a;">Loaded {count} cookies for {domains_str} from {browser}</p>'

    def on_sites_change(sites):
        if isinstance(sites, list):
            save_sites(sites)

    refresh_btn.click(refresh_models, outputs=[model_input])
    research_site.change(on_sites_change, inputs=[research_site])
    cookie_load_btn.click(on_load_cookies, inputs=[cookie_browser, research_site], outputs=[cookies_input, cookie_status])

    def run_research_and_maybe_keynote(prompt, site_url, max_articles, model, cookies, auto_gen, l1, l2, slides, thm):
        """Run research, then optionally auto-generate keynote from top result."""
        # Save preferences
        save_prefs(max_articles=int(max_articles or 30), auto_keynote=bool(auto_gen),
                   slides=int(slides or 10), lang1=l1, lang2=l2, theme=thm)

        # 7 outputs: research_log, research_kw, research_out, article_picker, pick_log, pick_file, pick_summary
        no_kn = gr.update()
        last_res_log = ""
        last_kw = ""
        last_out = ""
        last_picker = no_kn
        top_selection = None

        for res_log, res_kw, res_out, picker_update in run_research(prompt, site_url, int(max_articles), model, cookies):
            last_res_log = res_log
            last_kw = res_kw
            last_out = res_out
            last_picker = picker_update
            # Capture the top selection from the final picker update
            if isinstance(picker_update, dict) and picker_update.get("value"):
                top_selection = picker_update["value"]
            yield res_log, res_kw, res_out, picker_update, no_kn, no_kn, no_kn

        if not auto_gen:
            return

        # Auto-generate from top article
        if not top_selection or "|" not in top_selection:
            last_res_log += "\n\n⚠️ Auto-generate: no article found to generate from"
            yield last_res_log, last_kw, last_out, last_picker, "⚠️ No article to generate from", no_kn, no_kn
            return

        last_pipe = top_selection.rfind("|")
        url = top_selection[last_pipe + 1:].strip()
        title = top_selection[:last_pipe].strip()
        if not url.startswith("http"):
            last_res_log += f"\n\n⚠️ Auto-generate: invalid URL extracted: {url}"
            yield last_res_log, last_kw, last_out, last_picker, f"⚠️ Invalid URL: {url}", no_kn, no_kn
            return

        num = int(slides or 10)

        # Find related articles in the same topic group (max 3, from different domains)
        extra_urls = []
        top_topic = None
        top_domain = urlparse(normalize_url(url)).netloc
        for r in _research_results:
            if r.get("url") == url:
                top_topic = r.get("topic_group")
                break
        if top_topic:
            seen_domains = {top_domain}
            for r in _research_results:
                if r.get("topic_group") == top_topic and r.get("url") != url:
                    r_domain = urlparse(normalize_url(r["url"])).netloc
                    if r_domain not in seen_domains:
                        extra_urls.append(r["url"])
                        seen_domains.add(r_domain)
                if len(extra_urls) >= 3:
                    break

        # Show transition in both logs
        combine_msg = f" + {len(extra_urls)} related" if extra_urls else ""
        last_res_log += f"\n\n🚀 Auto-generating Keynote from: {title[:60]}{combine_msg}..."
        yield last_res_log, last_kw, last_out, last_picker, f"🚀 Starting Keynote generation...\n📎 {title[:70]}{combine_msg}\n🔗 {url}", no_kn, no_kn

        for kn_log, kn_file, kn_summary in run_pipeline(url, l1, l2, num, thm, True, model, cookies, user_prompt=prompt, extra_urls=extra_urls):
            yield last_res_log, last_kw, last_out, last_picker, kn_log, kn_file, kn_summary

    research_btn.click(
        run_research_and_maybe_keynote,
        inputs=[research_prompt, research_site, research_max, model_input, cookies_input, auto_keynote, pick_lang1, pick_lang2, pick_slides, pick_theme],
        outputs=[research_log, research_kw, research_out, article_picker, pick_log, pick_file, pick_summary],
    )

    def run_from_picker(selection, l1, l2, slides, thm, mdl, cookies, prompt):
        """Extract URL from picker selection and run keynote pipeline."""
        try:
            if not selection or "|" not in selection:
                yield "❌ No article selected — run Research first", None, ""
                return
            last_pipe = selection.rfind("|")
            url = selection[last_pipe + 1:].strip()
            if not url.startswith("http"):
                yield f"❌ Invalid URL: {url}", None, ""
                return
            yield from run_pipeline(url, l1, l2, int(slides), thm, True, mdl, cookies, user_prompt=prompt)
        except Exception as e:
            yield f"❌ Error: {e}", None, ""

    generate_from_research.click(
        run_from_picker,
        inputs=[article_picker, pick_lang1, pick_lang2, pick_slides, pick_theme, model_input, cookies_input, research_prompt],
        outputs=[pick_log, pick_file, pick_summary],
    )

    def run_pipeline_and_save(url, l1, l2, slides, thm, open_kn, mdl, cookies):
        save_prefs(slides=int(slides or 10), lang1=l1, lang2=l2, theme=thm)
        yield from run_pipeline(url, l1, l2, int(slides), thm, open_kn, mdl, cookies)

    btn.click(
        run_pipeline_and_save,
        inputs=[url_input, lang1, lang2, num_slides, theme, open_keynote, model_input, cookies_input],
        outputs=[log_output, file_output, summary_output],
    )

if __name__ == "__main__":
    app.launch(server_port=7860, share=False)
