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
            raise RuntimeError("GEMINI_API_KEY environment variable is not set")
        _client = genai.Client(api_key=api_key)
    return _client


TRIGGER_PROMPTS = {
    "header_focus": (
        "The user just clicked on the header text field in a Word document formatter. "
        "Give a single short, helpful tip (1-2 sentences) about advanced header options "
        "like different first-page or odd/even page headers, common in academic reports."
    ),
    "h1_large": (
        "The user set Heading 1 font size to more than 20pt in a Word document formatter. "
        "Suggest a complementary Heading 2 size that would look balanced. "
        "Give a single short suggestion (1-2 sentences)."
    ),
    "dark_color": (
        "The user picked a dark font color in a Word document formatter. "
        "Ask in one friendly sentence if they'd like a matching dark section header background."
    ),
    "ipynb_large": (
        "The user uploaded a Jupyter notebook with more than 10 cells to convert to Word. "
        "Suggest in one sentence that they consider adding a Table of Contents page."
    ),
    "double_spacing": (
        "The user selected double line spacing (2.0) in a Word document formatter. "
        "Note in one sentence that most Indian university submissions require 1.5 spacing "
        "and ask if they want to confirm their choice."
    ),
}


def get_suggestion(trigger: str, context: str = "") -> str:
    """Call Gemini to get a contextual formatting hint."""
    prompt = TRIGGER_PROMPTS.get(trigger)

    if not prompt:
        # Generic fallback for unknown triggers
        prompt = (
            f"The user is formatting a Word document. Context: {context}. "
            "Give a single brief, helpful formatting tip (1-2 sentences)."
        )
    elif context:
        prompt += f" Additional context: {context}"

    client = _get_client()
    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt,
    )

    text = response.text or ""
    # Keep it concise — truncate at 300 chars if model is verbose
    return text.strip()[:300]
