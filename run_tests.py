#!/usr/bin/env python3
"""Run all tests without pytest dependency."""
import sys
import io
import zipfile
import traceback

sys.path.insert(0, '/home/claude/docx2llm')

from docx2llm import convert, ConversionError
from tests.build_fixtures import (
    make_headings_doc, make_nested_lists_doc, make_ordered_lists_doc,
    make_table_doc, make_inline_formatting_doc, make_hyperlink_doc,
    make_empty_doc, make_unicode_rtl_doc, make_footnotes_doc,
    CONTENT_TYPES, RELS, WORD_RELS, STYLES, NUMBERING, W_NS, R_NS,
)

results = []


def check(name, condition, detail=''):
    if condition:
        results.append(('PASS', name))
        print(f'  \033[32mPASS\033[0m  {name}')
    else:
        results.append(('FAIL', name))
        print(f'  \033[31mFAIL\033[0m  {name}' + (f' — {detail}' if detail else ''))


def section(title):
    print(f'\n\033[1m--- {title} ---\033[0m')


# ---- EMPTY DOCUMENT ----
section('Empty document')
html = convert(make_empty_doc())
check('valid html structure', '<!DOCTYPE html>' in html and '</html>' in html)
check('no paragraph tags', '<p>' not in html)
check('no heading tags', '<h1>' not in html)

# ---- HEADINGS ----
section('Headings')
html = convert(make_headings_doc())
check('h1 present', '<h1>Chapter One</h1>' in html)
check('h2 present', '<h2>Section 1.1</h2>' in html)
check('normal → p', '<p>This is normal text.</p>' in html)
check('no mso-* noise', 'mso-' not in html)
check('no class attributes', 'class=' not in html)
check('no inline styles', 'style=' not in html)
check('no div tags', '<div' not in html)
check('no word namespaces', 'xmlns:w=' not in html)

# ---- INLINE FORMATTING ----
section('Inline formatting')
html = convert(make_inline_formatting_doc())
check('bold → strong', '<strong>Bold</strong>' in html)
check('italic → em', '<em>italic</em>' in html)
check('underline → u', '<u>underline</u>' in html)
check('superscript → sup', '<sup>sup</sup>' in html)
check('subscript → sub', '<sub>sub</sub>' in html)

# ---- NESTED LISTS (5 levels) ----
section('Nested lists (5+ levels)')
html = convert(make_nested_lists_doc())
check('ul present', '<ul>' in html)
check('item at ilvl 0', 'Item 1' in html)
check('item at ilvl 2', 'Item 1.1.1' in html)
check('item at ilvl 4', 'Item 2.1.1.1.1' in html)
check('multiple nesting levels (>=2 ul)', html.count('<ul>') >= 2, f'got {html.count("<ul>")}')
check('balanced ul tags', html.count('<ul>') == html.count('</ul>'),
      f'open={html.count("<ul>")} close={html.count("</ul>")}')

# ---- ORDERED LISTS ----
section('Ordered lists')
html = convert(make_ordered_lists_doc())
check('ol present', '<ol>' in html)
check('first item', 'First' in html)
check('nested sub-items', 'Sub A' in html)
check('balanced ol tags', html.count('<ol>') == html.count('</ol>'))

# ---- TABLES ----
section('Tables with merged cells')
html = convert(make_table_doc())
check('table tag', '<table>' in html)
check('thead detected', '<thead>' in html)
check('tbody present', '<tbody>' in html)
check('header cells are th', '<th>' in html)
check('colspan=2 preserved', 'colspan="2"' in html)
check('cell content', 'Alpha' in html and 'Name' in html)
check('& escaped as &amp;', '&amp;' in html)
check('balanced table tags', html.count('<table>') == html.count('</table>'))
check('balanced tr tags', html.count('<tr>') == html.count('</tr>'))

# ---- HYPERLINKS ----
section('Hyperlinks')
html = convert(make_hyperlink_doc())
check('anchor href', '<a href="https://example.com">' in html)
check('link text preserved', 'example.com' in html)
check('surrounding text', 'Visit' in html and 'for more.' in html)

# ---- FOOTNOTES ----
section('Footnotes')
html = convert(make_footnotes_doc())
check('footnote section', '<section id="footnotes">' in html)
check('footnote content', 'This is the first footnote.' in html)
check('bold in footnote', '<strong>bold</strong>' in html)
check('inline ref sup', '<sup><a href="#fn1">' in html)
html_plain = convert(make_footnotes_doc(), mode='plain')
check('plain mode: no footnotes', '<section id="footnotes">' not in html_plain)
html_llm = convert(make_footnotes_doc(), mode='llm')
check('llm mode: footnotes present', '<section id="footnotes">' in html_llm)

# ---- RTL / UNICODE ----
section('RTL and Unicode')
html = convert(make_unicode_rtl_doc())
check('Arabic text preserved', 'مرحبا' in html)
check('RTL dir attribute', 'dir="rtl"' in html)
check('LTR text no dir attr', 'Hello World' in html)
check('valid HTML', '<!DOCTYPE html>' in html and '</html>' in html)

# ---- CORRUPT INPUTS ----
section('Corrupt / malformed inputs')
try:
    convert(b'not a docx file')
    check('reject non-zip bytes', False)
except ConversionError:
    check('reject non-zip bytes', True)

try:
    convert(b'')
    check('reject empty bytes', False)
except ConversionError:
    check('reject empty bytes', True)

# Corrupt document.xml
buf = io.BytesIO()
with zipfile.ZipFile(buf, 'w') as zf:
    zf.writestr('[Content_Types].xml', '<Types/>')
    zf.writestr('_rels/.rels', '')
    zf.writestr('word/document.xml', '<<< NOT XML >>>')
try:
    html = convert(buf.getvalue())
    check('corrupt XML: no crash', '<!DOCTYPE html>' in html)
except Exception as e:
    check('corrupt XML: no crash', False, str(e))

# Missing styles.xml
buf2 = io.BytesIO()
doc_xml = '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p></w:body></w:document>'
with zipfile.ZipFile(buf2, 'w') as zf:
    zf.writestr('[Content_Types].xml', '<Types/>')
    zf.writestr('_rels/.rels', '')
    zf.writestr('word/document.xml', doc_xml)
html = convert(buf2.getvalue())
check('missing styles.xml: no crash', 'Hello' in html)

# Corrupt numbering.xml
docx = make_nested_lists_doc()
buf3 = io.BytesIO()
with zipfile.ZipFile(io.BytesIO(docx)) as src:
    with zipfile.ZipFile(buf3, 'w') as dst:
        for item in src.infolist():
            data = b'<<< CORRUPT >>>' if item.filename == 'word/numbering.xml' else src.read(item.filename)
            dst.writestr(item, data)
html = convert(buf3.getvalue())
check('corrupt numbering.xml: no crash', '<!DOCTYPE html>' in html)

# ---- TRACK CHANGES ----
section('Track changes')
body_tc = '''<w:p>
  <w:r><w:t xml:space="preserve">Keep this. </w:t></w:r>
  <w:ins w:id="1" w:author="Alice" w:date="2024-01-01T00:00:00Z">
    <w:r><w:t>Inserted.</w:t></w:r>
  </w:ins>
  <w:del w:id="2" w:author="Bob" w:date="2024-01-01T00:00:00Z">
    <w:r><w:delText>Deleted.</w:delText></w:r>
  </w:del>
</w:p>'''
doc_tc = f'<?xml version="1.0"?><w:document {W_NS} {R_NS}><w:body>{body_tc}</w:body></w:document>'
buf_tc = io.BytesIO()
with zipfile.ZipFile(buf_tc, 'w') as zf:
    zf.writestr('[Content_Types].xml', CONTENT_TYPES)
    zf.writestr('_rels/.rels', RELS)
    zf.writestr('word/_rels/document.xml.rels', WORD_RELS)
    zf.writestr('word/styles.xml', STYLES)
    zf.writestr('word/numbering.xml', NUMBERING)
    zf.writestr('word/document.xml', doc_tc)
tc_bytes = buf_tc.getvalue()

html_accept = convert(tc_bytes, track_changes='accept')
check('accept: inserted text present', 'Inserted.' in html_accept)
check('accept: deleted text absent', 'Deleted.' not in html_accept)

html_preserve = convert(tc_bytes, track_changes='preserve')
check('preserve: del tag present', '<del>' in html_preserve)
check('preserve: deleted text present', 'Deleted.' in html_preserve)

# ---- SMART QUOTES ----
section('Token optimization')
body_sq = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>\u201cHello\u201d and \u2018world\u2019</w:t></w:r></w:p>'
doc_sq = f'<?xml version="1.0"?><w:document {W_NS} {R_NS}><w:body>{body_sq}</w:body></w:document>'
buf_sq = io.BytesIO()
with zipfile.ZipFile(buf_sq, 'w') as zf:
    zf.writestr('[Content_Types].xml', CONTENT_TYPES)
    zf.writestr('_rels/.rels', RELS)
    zf.writestr('word/_rels/document.xml.rels', WORD_RELS)
    zf.writestr('word/styles.xml', STYLES)
    zf.writestr('word/numbering.xml', NUMBERING)
    zf.writestr('word/document.xml', doc_sq)
html = convert(buf_sq.getvalue())
check('smart double quotes normalized', '\u201c' not in html and '\u201d' not in html)
check('smart single quotes normalized', '\u2018' not in html and '\u2019' not in html)
check('ascii quotes in output', '&quot;Hello&quot;' in html or '"Hello"' in html)
check('no empty p tags', '<p></p>' not in html)
check('no span tags', '<span>' not in html)

# ---- MODES ----
section('Output modes')
html_llm = convert(make_headings_doc(), mode='llm')
check('llm: no base64', 'base64,' not in html_llm)
html_plain = convert(make_headings_doc(), mode='plain')
check('plain: valid html', '<!DOCTYPE html>' in html_plain)
html_preserve = convert(make_headings_doc(), mode='preserve')
check('preserve: valid html', '<!DOCTYPE html>' in html_preserve)

# ---- SUMMARY ----
total = len(results)
passed_n = sum(1 for r in results if r[0] == 'PASS')
failed_n = total - passed_n
print(f'\n{"="*50}')
print(f'\033[1m{passed_n}/{total} tests passed\033[0m', end='')
if failed_n:
    print(f'  (\033[31m{failed_n} failed\033[0m)')
else:
    print('  \033[32m✓ ALL PASS\033[0m')

sys.exit(0 if failed_n == 0 else 1)
