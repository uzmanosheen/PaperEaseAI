# This is the file that talks to Gemini. It handles three input types: photos of handwriting, a docx file, or raw text.

import json
from google import genai
from google.genai import types
from docx import Document as DocxReader
import config

client = genai.Client(api_key=config.GEMINI_API_KEY)

EXTRACTION_PROMPT = """
You are an expert exam-paper formatter and OCR reader for school exam papers,
which may be handwritten or typed, in English or Urdu.

Read the provided exam paper content and return ONLY valid JSON (no markdown
fences, no explanation, no extra text before or after) matching exactly this
schema:

{
  "exam_title": "string",
  "subject": "string",
  "language": "english" or "urdu",
  "instructions": ["string", ...],
  "sections": [
    {
      "section_title": "string",
      "section_marks": "string",
      "questions": [
        {
          "number": "string",
          "text": "string",
          "marks": "string",
          "sub_parts": ["string", ...]
        }
      ]
    }
  ],
  "total_marks": "string"
}

Rules:
- Preserve the original wording as closely as possible. Only silently fix
  obvious spelling/handwriting-recognition mistakes.
- If the paper is in Urdu, keep ALL question text in Urdu script (do not
  translate), and set "language" to "urdu".
- If a value (e.g. marks) is missing in the source, use an empty string "".
- Group MCQs / short questions / long questions into separate "sections" if
  the original paper visually separates them; otherwise use one section.
- Do NOT invent questions that are not present in the source material.
- Numbers in "number" should match the original numbering (e.g. "1", "Q2", "3a").
- The "instructions" field is ONLY for genuine exam-taking instructions
  (e.g. "Attempt all questions", "Write your answers in blue ink",
  "Section A is compulsory"). Do NOT put student fill-in fields like
  Name, Roll No, Class, Date, or Time into "instructions" — those are
  handled separately. If there are no genuine instructions on the paper,
  return an empty list [].
  - For multiple-choice questions, the question stem (the part ending in ":" or
  "?") goes in "text", and EACH option (a, b, c, d...) goes as its own
  separate string in "sub_parts" — e.g. "a) 8", "b) 9", "c) 5" as three
  separate list items. Never merge options into the "text" field.

"""


def _safe_json_parse(text: str) -> dict:
    """Gemini sometimes wraps JSON in ```json fences even when asked not to.
    This strips that safely before parsing."""
    text = text.strip()
    if text.startswith("```"):
        text = text.strip("`")
        if text.lower().startswith("json"):
            text = text[4:]
    return json.loads(text.strip())


def extract_from_images(image_files):
    """image_files: list of (bytes, mime_type) tuples, one per page/photo."""
    parts = [
        types.Part.from_bytes(data=data, mime_type=mime)
        for data, mime in image_files
    ]
    parts.append(EXTRACTION_PROMPT)

    response = client.models.generate_content(
        model=config.MODEL_NAME,
        contents=parts,
        config=types.GenerateContentConfig(response_mime_type="application/json"),
    )
    return _safe_json_parse(response.text)


def extract_from_docx(file_path: str):
    doc = DocxReader(file_path)
    raw_text = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    return extract_from_text(raw_text)


def extract_from_text(raw_text: str):
    response = client.models.generate_content(
        model=config.MODEL_NAME,
        contents=[raw_text, EXTRACTION_PROMPT],
        config=types.GenerateContentConfig(response_mime_type="application/json"),
    )
    return _safe_json_parse(response.text)