"""Parse /word/styles.xml to build a style-id → HTML tag / style-name map."""
from __future__ import annotations
import xml.etree.ElementTree as ET
from typing import Dict, Optional, Tuple

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
}

# Official Word heading style names → HTML tag
_HEADING_MAP: Dict[str, str] = {
    'heading 1': 'h1', 'heading1': 'h1',
    'heading 2': 'h2', 'heading2': 'h2',
    'heading 3': 'h3', 'heading3': 'h3',
    'heading 4': 'h4', 'heading4': 'h4',
    'heading 5': 'h5', 'heading5': 'h5',
    'heading 6': 'h6', 'heading6': 'h6',
}


class StyleRegistry:
    """Maps styleId → (html_tag, style_name)."""

    def __init__(self) -> None:
        self._by_id: Dict[str, Tuple[str, str]] = {}   # styleId → (tag, name)
        self._by_name: Dict[str, str] = {}              # lower name → tag

    # ------------------------------------------------------------------
    def parse(self, xml_bytes: bytes) -> None:
        try:
            root = ET.fromstring(xml_bytes)
        except ET.ParseError:
            return
        for style in root.findall('w:style', NS):
            sid = style.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}styleId', '')
            stype = style.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}type', '')
            name_el = style.find('w:name', NS)
            if name_el is None:
                continue
            name = name_el.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '')
            tag = self._resolve_tag(name, stype)
            self._by_id[sid] = (tag, name)
            self._by_name[name.lower()] = tag

    # ------------------------------------------------------------------
    def _resolve_tag(self, name: str, stype: str) -> str:
        lower = name.lower()
        if lower in _HEADING_MAP:
            return _HEADING_MAP[lower]
        # Numbered heading variants like "Heading1Char"
        for k, v in _HEADING_MAP.items():
            if lower.startswith(k.replace(' ', '')):
                return v
        return 'p'

    # ------------------------------------------------------------------
    def get_tag(self, style_id: Optional[str]) -> str:
        if style_id and style_id in self._by_id:
            return self._by_id[style_id][0]
        return 'p'

    def get_name(self, style_id: Optional[str]) -> str:
        if style_id and style_id in self._by_id:
            return self._by_id[style_id][1]
        return ''

    def is_heading(self, style_id: Optional[str]) -> bool:
        return self.get_tag(style_id) in ('h1', 'h2', 'h3', 'h4', 'h5', 'h6')
