"""
AI prompt templates for Web AI Tool.

All prompts used for LLM calls are defined here for easy management.
Each function returns a formatted prompt string.
"""


def translate_and_summarize(content: str, language: str, user_prompt: str = "") -> str:
    user_instruction = ""
    if user_prompt:
        user_instruction = f"""
The user's original research prompt was: "{user_prompt}"
Consider this context when summarizing — focus on aspects most relevant to the user's interest.
"""
    return f"""
Translate and provide a comprehensive summary of the following content into {language}.
The summary should be detailed and thorough (500-800 words). Cover all key points, data, quotes, and analysis.
Do NOT be brief — include specific numbers, names, dates, and details from the article(s).
Number formatting rules (apply to ALL languages):
- Use abbreviations for large numbers: M for million, B for billion, T for trillion (e.g. "$15B" not "$15,000,000,000", "$1.25M" not "$1,250,000")
- For Chinese/zh-TW: follow Chinese units — use 萬 and 億 with digits (e.g. "150億美元" not "$15,000,000,000", "1,250萬" not "12,500,000")
- Numbers under 1 million: use thousand separators (e.g. "3,500" not "three thousand")
- Always use Arabic numerals and digits, never spell out numbers in any language
- Percentages: "7.2%" not "seven percent" or "百分之七"
{user_instruction}
Content:
{content[:12000]}
"""


def slide_json_template(lang1: str, lang2: str = "") -> str:
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


def generate_slides(content: str, language: str, num_slides: int, lang2: str = "", user_instruction: str = "", slide_context: str = "") -> str:
    """Build the prompt for generating slides."""
    if lang2:
        lang_instruction = f"Each slide MUST have BOTH {language} AND {lang2} versions of the heading and bullets."
    else:
        lang_instruction = f"All content must be in {language}."

    return f"""
Create exactly {num_slides} content slides for a keynote presentation.
{lang_instruction}
{user_instruction}
{slide_context}
IMPORTANT rules:
- Each slide MUST have 4-6 detailed bullet points
- Each bullet: full sentence (15-30 words), include specific data
- Cover different angles: facts, analysis, market impact, quotes, outlook
- Each slide should have a distinct sub-topic
- SYNTHESIZE information from ALL sources if multiple are provided
- Number formatting: use M/B/T for large numbers ("$15B" not "$15,000,000,000", "$1.25M" not "$1,250,000"). For Chinese: use 萬/億 with digits ("150億美元", "1,250萬"). Under 1M use thousand separators ("3,500"). Always use digits, never spell out. Percentages: "7.2%" not "seven percent"

Content:
{content}

Respond ONLY with valid JSON, no markdown:
{slide_json_template(language, lang2)}
"""


def extract_keywords(prompt: str) -> str:
    return f"""
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


def filter_titles(prompt: str, title_list: str) -> str:
    return f"""
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


def score_article(prompt: str, content: str) -> str:
    return f"""
You are evaluating whether a news article is relevant to a research topic.

Research topic: {prompt}

Article content (excerpt):
{content}

Rate the relevance from 0.0 to 1.0 and explain briefly why.
Return ONLY valid JSON, no markdown:
{{"relevance": <0.0-1.0>, "reason": "<1 sentence>"}}
"""


def group_articles(prompt: str, titles_for_grouping: str) -> str:
    return f"""
You have a list of news articles found for the topic: "{prompt}"

Articles:
{titles_for_grouping}

Group these articles by sub-topic/theme. Articles covering the same event or angle should be grouped together.
Return ONLY valid JSON, no markdown:
[
  {{"topic": "Brief topic name", "articles": [0, 2, 5], "summary": "What these articles share"}}
]
"""
