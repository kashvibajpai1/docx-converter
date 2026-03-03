"""
Build test fixture .docx files using python-docx.
Run once: python -m tests.build_fixtures
"""
from __future__ import annotations
import io
import os
import zipfile

# We build minimal valid .docx files from raw XML to avoid needing python-docx
# This ensures we test our own parser, not a third-party writer.

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>'''

RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    Target="word/document.xml"/>
</Relationships>'''

WORD_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    Target="https://example.com" TargetMode="External"/>
</Relationships>'''

STYLES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
</w:styles>'''

NUMBERING = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/></w:lvl>
    <w:lvl w:ilvl="1"><w:numFmt w:val="bullet"/></w:lvl>
    <w:lvl w:ilvl="2"><w:numFmt w:val="bullet"/></w:lvl>
    <w:lvl w:ilvl="3"><w:numFmt w:val="bullet"/></w:lvl>
    <w:lvl w:ilvl="4"><w:numFmt w:val="bullet"/></w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl>
    <w:lvl w:ilvl="1"><w:numFmt w:val="decimal"/></w:lvl>
    <w:lvl w:ilvl="2"><w:numFmt w:val="decimal"/></w:lvl>
    <w:lvl w:ilvl="3"><w:numFmt w:val="decimal"/></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>
  <w:num w:numId="2">
    <w:abstractNumId w:val="1"/>
  </w:num>
</w:numbering>'''

W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
R_NS = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'


def _p(text: str, style: str = 'Normal', bold: bool = False,
       italic: bool = False, num_id: str = '', ilvl: int = 0) -> str:
    rpr = ''
    if bold or italic:
        b = '<w:b/>' if bold else ''
        i = '<w:i/>' if italic else ''
        rpr = f'<w:rPr>{b}{i}</w:rPr>'
    numpr = ''
    if num_id:
        numpr = f'<w:numPr><w:ilvl w:val="{ilvl}"/><w:numId w:val="{num_id}"/></w:numPr>'
    ppr = f'<w:pPr><w:pStyle w:val="{style}"/>{numpr}</w:pPr>'
    run = f'<w:r>{rpr}<w:t xml:space="preserve">{text}</w:t></w:r>'
    return f'<w:p>{ppr}{run}</w:p>'


def _make_docx(document_xml: str, numbering: str = NUMBERING) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', CONTENT_TYPES)
        zf.writestr('_rels/.rels', RELS)
        zf.writestr('word/_rels/document.xml.rels', WORD_RELS)
        zf.writestr('word/styles.xml', STYLES)
        zf.writestr('word/numbering.xml', numbering)
        zf.writestr('word/document.xml', document_xml)
    return buf.getvalue()


def make_headings_doc() -> bytes:
    body = '\n'.join([
        _p('Chapter One', 'Heading1'),
        _p('Section 1.1', 'Heading2'),
        _p('This is normal text.'),
        _p('Another paragraph.'),
    ])
    doc = f'<?xml version="1.0" encoding="UTF-8"?><w:document {W_NS} {R_NS}><w:body>{body}</w:body></w:document>'
    return _make_docx(doc)


def make_nested_lists_doc() -> bytes:
    items = [
        _p('Item 1', num_id='1', ilvl=0),
        _p('Item 1.1', num_id='1', ilvl=1),
        _p('Item 1.1.1', num_id='1', ilvl=2),
        _p('Item 1.1.2', num_id='1', ilvl=2),
        _p('Item 1.2', num_id='1', ilvl=1),
        _p('Item 2', num_id='1', ilvl=0),
        _p('Item 2.1', num_id='1', ilvl=1),
        _p('Item 2.1.1', num_id='1', ilvl=2),
        _p('Item 2.1.1.1', num_id='1', ilvl=3),
        _p('Item 2.1.1.1.1', num_id='1', ilvl=4),
    ]
    body = '\n'.join(items)
    doc = f'<?xml version="1.0" encoding="UTF-8"?><w:document {W_NS} {R_NS}><w:body>{body}</w:body></w:document>'
    return _make_docx(doc)


def make_ordered_lists_doc() -> bytes:
    items = [
        _p('First', num_id='2', ilvl=0),
        _p('Second', num_id='2', ilvl=0),
        _p('Sub A', num_id='2', ilvl=1),
        _p('Sub B', num_id='2', ilvl=1),
        _p('Third', num_id='2', ilvl=0),
    ]
    body = '\n'.join(items)
    doc = f'<?xml version="1.0" encoding="UTF-8"?><w:document {W_NS} {R_NS}><w:body>{body}</w:body></w:document>'
    return _make_docx(doc)


def make_table_doc() -> bytes:
    table = '''<w:tbl>
      <w:tr>
        <w:trPr><w:tblHeader/></w:trPr>
        <w:tc><w:p><w:r><w:t>Name</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Value</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Notes</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Alpha</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>42</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>First item</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc>
          <w:tcPr><w:gridSpan w:val="2"/></w:tcPr>
          <w:p><w:r><w:t>Beta &amp; Gamma merged</w:t></w:r></w:p>
        </w:tc>
        <w:tc><w:p><w:r><w:t>Merged col</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>'''
    doc = f'<?xml version="1.0" encoding="UTF-8"?><w:document {W_NS} {R_NS}><w:body>{table}</w:body></w:document>'
    return _make_docx(doc)


def make_inline_formatting_doc() -> bytes:
    body = f'''<w:p>
      <w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r>
      <w:r><w:t xml:space="preserve"> and </w:t></w:r>
      <w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r>
      <w:r><w:t xml:space="preserve"> and </w:t></w:r>
      <w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>underline</w:t></w:r>
      <w:r><w:t xml:space="preserve"> and </w:t></w:r>
      <w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>sup</w:t></w:r>
      <w:r><w:t xml:space="preserve"> and </w:t></w:r>
      <w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t>sub</w:t></w:r>
    </w:p>'''
    doc = f'<?xml version="1.0" encoding="UTF-8"?><w:document {W_NS} {R_NS}><w:body>{body}</w:body></w:document>'
    return _make_docx(doc)


def make_hyperlink_doc() -> bytes:
    body = '''<w:p>
      <w:r><w:t xml:space="preserve">Visit </w:t></w:r>
      <w:hyperlink r:id="rId1">
        <w:r><w:t>example.com</w:t></w:r>
      </w:hyperlink>
      <w:r><w:t xml:space="preserve"> for more.</w:t></w:r>
    </w:p>'''
    doc = f'<?xml version="1.0" encoding="UTF-8"?><w:document {W_NS} {R_NS}><w:body>{body}</w:body></w:document>'
    return _make_docx(doc)


def make_empty_doc() -> bytes:
    doc = f'<?xml version="1.0" encoding="UTF-8"?><w:document {W_NS} {R_NS}><w:body></w:body></w:document>'
    return _make_docx(doc)


def make_unicode_rtl_doc() -> bytes:
    body = f'''<w:p>
        <w:pPr><w:bidi/></w:pPr>
        <w:r><w:t>مرحبا بالعالم</w:t></w:r>
      </w:p>
      <w:p>
        <w:r><w:t>Hello World</w:t></w:r>
      </w:p>'''
    doc = f'<?xml version="1.0" encoding="UTF-8"?><w:document {W_NS} {R_NS}><w:body>{body}</w:body></w:document>'
    return _make_docx(doc)


def make_footnotes_doc() -> bytes:
    footnotes_xml = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:id="1">
    <w:p><w:r><w:t>This is the first footnote.</w:t></w:r></w:p>
  </w:footnote>
  <w:footnote w:id="2">
    <w:p><w:r><w:t>Second footnote with </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r><w:r><w:t> text.</w:t></w:r></w:p>
  </w:footnote>
</w:footnotes>'''

    body = '''<w:p>
      <w:r><w:t xml:space="preserve">Text with footnote</w:t></w:r>
      <w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>
        <w:footnoteReference w:id="1"/>
      </w:r>
      <w:r><w:t xml:space="preserve"> and another</w:t></w:r>
      <w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>
        <w:footnoteReference w:id="2"/>
      </w:r>
      <w:r><w:t>.</w:t></w:r>
    </w:p>'''

    doc = f'<?xml version="1.0" encoding="UTF-8"?><w:document {W_NS} {R_NS}><w:body>{body}</w:body></w:document>'

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('[Content_Types].xml', CONTENT_TYPES)
        zf.writestr('_rels/.rels', RELS)
        zf.writestr('word/_rels/document.xml.rels', WORD_RELS)
        zf.writestr('word/styles.xml', STYLES)
        zf.writestr('word/numbering.xml', NUMBERING)
        zf.writestr('word/document.xml', doc)
        zf.writestr('word/footnotes.xml', footnotes_xml)
    return buf.getvalue()


# Save all fixtures
FIXTURES_DIR = os.path.join(os.path.dirname(__file__), '..', 'test_fixtures')


def build_all() -> None:
    os.makedirs(FIXTURES_DIR, exist_ok=True)
    fixtures = {
        'headings.docx': make_headings_doc(),
        'nested_lists.docx': make_nested_lists_doc(),
        'ordered_lists.docx': make_ordered_lists_doc(),
        'table.docx': make_table_doc(),
        'inline_formatting.docx': make_inline_formatting_doc(),
        'hyperlink.docx': make_hyperlink_doc(),
        'empty.docx': make_empty_doc(),
        'unicode_rtl.docx': make_unicode_rtl_doc(),
        'footnotes.docx': make_footnotes_doc(),
    }
    for name, data in fixtures.items():
        path = os.path.join(FIXTURES_DIR, name)
        with open(path, 'wb') as f:
            f.write(data)
        print(f'  wrote {path} ({len(data):,} bytes)')


if __name__ == '__main__':
    build_all()
