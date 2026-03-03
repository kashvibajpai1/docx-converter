"""Parse /word/footnotes.xml and /word/endnotes.xml."""
from __future__ import annotations
import xml.etree.ElementTree as ET
from html import escape
from typing import Dict

from .builder import normalize_text, collapse_inline

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'


def _w(tag: str) -> str:
    return f'{{{W}}}{tag}'


def _local(tag: str) -> str:
    if '}' in tag:
        return tag.split('}')[1]
    return tag


def parse_notes(xml_bytes: bytes, kind: str = 'fn') -> Dict[str, str]:
    """
    Return dict of note_id → HTML fragment (a <li> element).
    kind: 'fn' for footnotes, 'en' for endnotes.
    """
    result: Dict[str, str] = {}
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError:
        return result

    note_tag = 'footnote' if kind == 'fn' else 'endnote'
    prefix = 'fn' if kind == 'fn' else 'en'

    for note in root.findall(_w(note_tag)):
        nid = note.get(_w('id'), '')
        # Skip separator notes
        ntype = note.get(_w('type'), '')
        if ntype in ('separator', 'continuationSeparator', 'continuationNotice'):
            continue

        texts = []
        for p in note.findall('.//' + _w('p')):
            para_texts = []
            for r in p.findall('.//' + _w('r')):
                rpr = r.find(_w('rPr'))
                t_el = r.find(_w('t'))
                if t_el is not None:
                    text = escape(normalize_text(t_el.text or ''))
                    if rpr is not None:
                        b = rpr.find(_w('b'))
                        i = rpr.find(_w('i'))
                        if b is not None and b.get(_w('val'), '1') not in ('0', 'false'):
                            text = f'<strong>{text}</strong>'
                        if i is not None and i.get(_w('val'), '1') not in ('0', 'false'):
                            text = f'<em>{text}</em>'
                    para_texts.append(text)
            if para_texts:
                texts.append(''.join(para_texts))

        content = ' '.join(texts)
        if content:
            result[nid] = f'<li id="{prefix}{nid}">{content}</li>\n'

    return result
