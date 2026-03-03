"""Parse /word/_rels/document.xml.rels for hyperlinks and image references."""
from __future__ import annotations
import xml.etree.ElementTree as ET
from typing import Dict
import re

REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'
HYPERLINK_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
IMAGE_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
FOOTNOTES_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes'
ENDNOTES_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes'

_SAFE_SCHEMES = re.compile(r'^(https?|mailto|ftp)://', re.IGNORECASE)


class RelationshipRegistry:
    def __init__(self) -> None:
        self._rels: Dict[str, Dict[str, str]] = {}  # rId → {type, target, targetMode}

    def parse(self, xml_bytes: bytes) -> None:
        try:
            root = ET.fromstring(xml_bytes)
        except ET.ParseError:
            return
        for rel in root:
            rid = rel.get('Id', '')
            self._rels[rid] = {
                'type': rel.get('Type', ''),
                'target': rel.get('Target', ''),
                'mode': rel.get('TargetMode', ''),
            }

    def get_hyperlink(self, rid: str) -> str:
        rel = self._rels.get(rid, {})
        if rel.get('type') != HYPERLINK_TYPE:
            return ''
        url = rel.get('target', '')
        return _sanitize_url(url)

    def get_image_path(self, rid: str) -> str:
        """Return the media path inside the zip, e.g. word/media/image1.png"""
        rel = self._rels.get(rid, {})
        if rel.get('type') != IMAGE_TYPE:
            return ''
        target = rel.get('target', '')
        # Targets are relative to word/ directory
        if not target.startswith('/'):
            target = 'word/' + target
        return target.lstrip('/')

    def get_image_alt(self, rid: str) -> str:
        return ''


def _sanitize_url(url: str) -> str:
    url = url.strip()
    if not url:
        return ''
    if _SAFE_SCHEMES.match(url):
        return url
    if url.startswith('#'):
        return url
    return ''
