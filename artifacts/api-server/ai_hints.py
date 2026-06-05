"""
Gemini-powered AI formatting suggestions for WordForge.
"""

import os
from google import genai

_client = None


def _get_client():
    global _client
    if _client is None:
        api_key = os.environ.get("GEMINI_API_KEY")
        if not api_key:
            raise RuntimeError("GEMINI_API_KEY is not set")
        _client = genai.Client(api_key=api_key)
    return _client


TRIGGER_PROMPTS = {
    "header_focus": (
        "A student is formatting an academic Word document and just clicked on the header field. "
        "Give a single short tip (1-2 sentences) about using different headers for odd and even pages, "
        "which is standard for dissertations and reports."
    ),
    "h1_large": (
        "A student set Heading 1 to more than 20pt in their academic document formatter. "
        "Suggest a specific complementary Heading 2 size in pt that would look well-balanced. "
        "Keep it to one sentence with the actual number."
    ),
    "h1_h2_same": (
        "A student has set Heading 1 and Heading 2 to the same font size in their Word document. "
        "Warn them in one concise sentence that headings won't be visually distinguishable in the output."
    ),
    "dark_color": (
        "A student picked a dark (non-black) font tint in their document formatter. "
        "Ask in one friendly sentence if they'd like a matching shaded background on section headers."
    ),
    "double_spacing": (
        "A student selected double line spacing (2.0) in their academic document formatter. "
        "Note in one sentence that most Indian university submissions require 1.5 spacing "
        "and ask if this is intentional."
    ),
    "no_header_export": (
        "A student is about to download their formatted Word document but left the header field blank. "
        "Remind them in one sentence that many universities require department name or subject in the header."
    ),
    "footer_left_pagenum": (
        "A student has set page numbers to align left and has footer text. "
        "Suggest in one sentence whether 'Page X of Y' format might look more professional than just the page number."
    ),
    "ipynb_large": (
        "A student uploaded a Jupyter notebook with more than 10 cells to convert to Word. "
        "Suggest in one sentence that adding a Table of Contents page at the start might help navigation."
    ),
    "ipynb_images": (
        "A student's Jupyter notebook has matplotlib or plot image outputs. "
        "Ask in one sentence whether they'd like automatic figure captions ('Figure 1', 'Figure 2', etc.) "
        "added below each plot image in the Word document."
    ),
}


def get_suggestion(trigger: str, context: str = "") -> str:
    """Call Gemini to get a contextual formatting hint."""
    prompt = TRIGGER_PROMPTS.get(trigger)

    if not prompt:
        prompt = (
            f"A student is formatting an academic Word document. Context: {context}. "
            "Give one brief, helpful formatting tip (1-2 sentences)."
        )
    elif context:
        prompt += f" Additional context: {context}"

    client = _get_client()
    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt,
    )

    text = response.text or ""
    return text.strip()[:350]
