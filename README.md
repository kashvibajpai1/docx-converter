# docx2llm

**Production-grade DOCX → HTML5 converter optimized for LLM ingestion.**

Zero dependencies. Pure Python stdlib. Deterministic output.

---

## Why docx2llm?

Standard converters (Pandoc, LibreOffice) produce HTML littered with Word noise:
`mso-*` classes, inline styles, empty spans, broken list nesting. That noise wastes
tokens and confuses LLMs.

`docx2llm` is a **semantic structural extractor** — not a visual renderer.
It produces the minimum valid HTML5 needed to faithfully represent the document's
structure, stripping everything else.

**Typical token reduction vs Pandoc: 30–60%.**

---

## Installation

```bash
pip install docx2llm
```

Or from source:
```bash
git clone https://github.com/you/docx2llm
cd docx2llm
pip install -e .
```

No runtime dependencies. Requires Python ≥ 3.9.

---

## Quick Start

```bash
# LLM mode (default) — minimal HTML, images as placeholders
docx2llm report.docx > report.html

# Preserve mode — base64 images, more structure
docx2llm report.docx --mode=preserve -o report.html

# Plain mode — strip images and footnotes entirely
docx2llm report.docx --mode=plain > report.html

# Compare token count with Pandoc
docx2llm report.docx --token-count --compare-pandoc
```

---

## Python API

```python
from docx2llm import convert

with open("report.docx", "rb") as f:
    docx_bytes = f.read()

html = convert(docx_bytes)                          # llm mode
html = convert(docx_bytes, mode="preserve")         # keep images
html = convert(docx_bytes, mode="plain")            # strip everything
html = convert(docx_bytes, track_changes="accept")  # accept tracked changes
```

---

## CLI Options

| Option | Default | Description |
|--------|---------|-------------|
| `--mode` | `llm` | `llm` \| `preserve` \| `plain` |
| `--track-changes` | `accept` | `accept` \| `reject` \| `preserve` |
| `--normalize` | off | Normalize fake headings |
| `-o FILE` | stdout | Write to file |
| `--token-count` | off | Print estimated token count |
| `--compare-pandoc` | off | Compare with Pandoc (requires pandoc in PATH) |

---

## Output Modes

### `--mode=llm` (default)
- Images replaced by `[IMAGE: alt text]` placeholders
- Footnotes appended as `<section id="footnotes">`
- Minimal, semantic HTML only

### `--mode=preserve`
- Images embedded as `<img src="data:image/...;base64,...">` 
- All structural information retained

### `--mode=plain`
- Images and footnotes stripped entirely
- Maximum token efficiency

---

## Structural Mapping

| Word element | HTML output |
|---|---|
| Heading 1–6 | `<h1>`–`<h6>` |
| Normal paragraph | `<p>` |
| Bullet list | `<ul><li>` (nested) |
| Numbered list | `<ol><li>` (nested) |
| Table with header row | `<table><thead><tbody>` |
| Merged cells | `colspan` / `rowspan` |
| Bold | `<strong>` |
| Italic | `<em>` |
| Underline | `<u>` |
| Superscript | `<sup>` |
| Subscript | `<sub>` |
| Hyperlink | `<a href="...">` |
| Page break | `<hr data-page-break>` |
| Footnote reference | `<sup><a href="#fn1">1</a></sup>` |
| RTL paragraph | `<p dir="rtl">` |

---

## Track Changes

```bash
docx2llm doc.docx --track-changes=accept   # accept inserts, discard deletes
docx2llm doc.docx --track-changes=reject   # discard inserts, keep deletes
docx2llm doc.docx --track-changes=preserve # emit <ins> and <del> markup
```

---

## Security

- No macro execution
- No remote URL fetching  
- Hyperlink sanitization (only `http://`, `https://`, `mailto:`, `ftp://`, anchors)
- ZIP extraction size limits (500 MB total, 100 MB per member)
- Invalid XML handled gracefully — never crashes

---

## Architecture

```
docx2llm/
├── __init__.py        Public API: convert()
├── cli.py             CLI entrypoint
├── converter.py       Orchestrator: opens ZIP, builds registries, runs parser
├── parser.py          Core document.xml streaming parser
├── styles.py          Styles registry (styleId → HTML tag)
├── numbering.py       Numbering registry (numId + ilvl → ul/ol)
├── relations.py       Relationship registry (rId → URL / media path)
├── notes.py           Footnote/endnote parser
└── builder.py         HTML assembly, token optimization, list stack
```

### Design principles

1. **Streaming-first**: Uses `xml.etree.ElementTree` iteratively, not a full DOM
2. **Separation of concerns**: Each XML part has its own registry parsed once upfront
3. **ListStack**: Tracks open `<ul>`/`<ol>` nesting precisely, handling level changes and type switches
4. **Zero noise**: No `class=`, `style=`, `div`, `span` in output
5. **Deterministic**: Same input always produces identical output

---

## Testing

```bash
python run_tests.py          # built-in runner (no dependencies)
python -m pytest tests/ -v   # if pytest is installed
```

Test coverage:
- Empty document
- Headings (h1–h6)
- Inline formatting (bold, italic, underline, sup, sub)
- Nested lists (5+ levels)
- Ordered lists with sub-levels
- Tables with `<thead>`, `colspan`
- Hyperlinks
- Footnotes
- RTL / Arabic text
- Track changes (accept/reject/preserve)
- Smart quote normalization
- Corrupt ZIP, corrupt XML, missing styles.xml, corrupt numbering.xml
- All output modes

---

## Limitations

- **rowspan**: Best-effort only (full rowspan count requires a pre-pass)
- **Text boxes / SmartArt / Charts**: Fallback to `[IMAGE]` placeholder
- **Embedded spreadsheets**: Not extracted (treated as images)
- **Complex fields** (`=SUM(...)`, TOC): Display text only

---

## License

MIT
