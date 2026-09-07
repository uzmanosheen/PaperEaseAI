# PaperEaseAI 📝

**Smart Exam Formatting Assistant** — turn messy, handwritten or unformatted exam papers into polished, print-ready Word documents in under two minutes.

## The Problem

Teachers and coaching institutes spend **hours** on every exam paper — typing questions out, fixing numbering, drawing header tables, aligning MCQ options, adding marks columns. It's repetitive, error-prone work that takes time away from actual teaching. Existing "solutions" (typing it into Word yourself, or asking a chatbot) still leave the formatting job entirely in your hands.

## How It Works

1. **Upload** — photos of a handwritten draft (multi-page supported), an existing `.docx`, or pasted plain text
2. **AI Extraction** — Google Gemini reads the messy input and structures every question into strict JSON: sections, question numbers, marks, MCQ options
3. **Template Engine** — a professional exam layout is applied deterministically: masthead with school name & logo, student info grid, section headers, MCQ alignment, page numbers
4. **Download** — a print-ready `.docx`, named automatically from the subject and class (e.g. `formatted-english-6-exam.docx`)

AI handles the fuzzy part (reading messy input); deterministic code handles the precise part (pixel-perfect formatting). We never ask the LLM to produce formatting — that's why the output is clean every time.

## Features

- **Multi-modal input** — handwritten photos, typed `.docx`, or raw text: meet teachers where they are
- **Two built-in templates** + **custom school letterhead** upload (your own `.docx` with `{{ school_name }}`, `{{ logo }}` placeholders)
- **Smart MCQ rendering** — short option sets go on one evenly-spaced line, longer sets flow into two-per-line tab-aligned columns
- **Objective/Subjective labels** — detected from question structure, printed only when the type changes
- **Professional masthead** — logo + centered school name (full page width), student info grid (name, roll number, marks, invigilator sign, remarks)
- **Running header from page 2** — Name / Roll No. / Subject-Class + automatic `Page X of Y` numbering; page 1 keeps the full masthead
- **English + Urdu papers** — Urdu output uses Nastaliq script fonts and proper right-to-left paragraphs
- **Dynamic filenames** — slugged from the Subject and Class form fields

## Architecture

```
Input (photo / .docx / text)
        │
        ▼
Google Gemini (multimodal extraction → strict JSON)
        │
        ▼
Template Engine (python-docx, OXML-level formatting)
        │
        ▼
Print-ready .docx
```

## Tech Stack

| Layer | Tool |
|---|---|
| Web UI | Streamlit |
| AI extraction | Google Gemini API (`google-genai`) |
| Document rendering | python-docx |
| Custom templates | docxtpl (Jinja2) |
| Image handling | Pillow |

## Getting Started

Requires Python 3.9+.

```bash
git clone <repo-url>
cd PaperEaseAI
python -m venv venv
venv\Scripts\activate        # Windows  (macOS/Linux: source venv/bin/activate)
pip install -r requirements.txt
copy .env.example .env       # macOS/Linux: cp .env.example .env
# → edit .env and paste your Gemini API key
streamlit run app.py
```

Get a free API key at [Google AI Studio](https://aistudio.google.com/apikey). The app runs at `http://localhost:8501`.

## Project Structure

```
app.py               # Streamlit UI + the 4-step workflow
extractor.py         # Gemini extraction & JSON structuring
doc_builder.py       # DOCX template engine (masthead, MCQs, sections, headers)
custom_template.py   # Custom school-letterhead flow (docxtpl)
config.py            # API key & model configuration
```

## Future Roadmap

- CBSE / ICSE / university-specific template packs
- Batch processing of multiple papers at once
- PDF export
- AI-generated marking schemes
