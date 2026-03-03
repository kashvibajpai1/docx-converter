"""Build clean, token-optimized HTML5 from structured data."""
from __future__ import annotations
from html import escape
from typing import List, Optional
import re

# Smart quote normalization map
_SMART_QUOTES = str.maketrans({
    '\u2018': "'", '\u2019': "'",
    '\u201c': '"', '\u201d': '"',
    '\u2013': '-', '\u2014': '--',
    '\u00a0': ' ', '\u200b': '',
    '\ufeff': '',
})

_WHITESPACE = re.compile(r'[ \t]+')


def normalize_text(text: str) -> str:
    text = text.translate(_SMART_QUOTES)
    text = _WHITESPACE.sub(' ', text)
    return text


def build_html(body_parts: List[str], footnotes: List[str], mode: str) -> str:
    parts = ['<!DOCTYPE html>\n<html>\n<head><meta charset="utf-8"></head>\n<body>\n']
    parts.extend(body_parts)
    if footnotes and mode != 'plain':
        parts.append('\n<section id="footnotes">\n<ol>\n')
        parts.extend(footnotes)
        parts.append('</ol>\n</section>\n')
    parts.append('</body>\n</html>')
    return ''.join(parts)


def tag(name: str, content: str, attrs: Optional[dict] = None) -> str:
    if not content and content != '0':
        return ''
    attr_str = ''
    if attrs:
        attr_str = ''.join(f' {k}="{escape(str(v))}"' for k, v in attrs.items() if v is not None)
    return f'<{name}{attr_str}>{content}</{name}>'


def collapse_inline(html: str) -> str:
    """Remove redundant adjacent identical inline tags."""
    for t in ('strong', 'em', 'u', 'sub', 'sup'):
        html = re.sub(rf'</{t}><{t}>', '', html)
    return html


class ListStack:
    """
    Manage nested list open/close tags.
    State: list of (list_tag, num_id, ilvl)
    """

    def __init__(self) -> None:
        self._stack: List[tuple] = []  # (list_tag, num_id, ilvl)
        self._buf: List[str] = []

    def emit(self, list_tag: str, num_id: str, ilvl: int, item_html: str) -> str:
        out = []

        # Close deeper levels first
        while self._stack and self._stack[-1][2] > ilvl:
            out.append(f'</li>\n</{self._stack[-1][0]}>\n')
            self._stack.pop()

        # Same level but different list type or num_id → close and reopen
        if self._stack and self._stack[-1][2] == ilvl:
            top = self._stack[-1]
            if top[0] != list_tag or top[1] != num_id:
                out.append(f'</li>\n</{top[0]}>\n')
                self._stack.pop()
                out.append(f'<{list_tag}>\n<li>')
                self._stack.append((list_tag, num_id, ilvl))
            else:
                out.append(f'</li>\n<li>')
        elif not self._stack or self._stack[-1][2] < ilvl:
            # Need to open a new level
            out.append(f'<{list_tag}>\n<li>')
            self._stack.append((list_tag, num_id, ilvl))

        out.append(item_html)
        return ''.join(out)

    def close_all(self) -> str:
        out = []
        while self._stack:
            out.append(f'</li>\n</{self._stack[-1][0]}>\n')
            self._stack.pop()
        return ''.join(out)

    @property
    def is_open(self) -> bool:
        return bool(self._stack)
