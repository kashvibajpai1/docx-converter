"""
docx2llm.converter
==================
Top-level orchestrator: opens the .docx ZIP, loads all sub-documents,
runs the parser, and returns clean HTML5.
"""
from __future__ import annotations
import io
import zipfile
from typing import Dict, Optional

from .styles import StyleRegistry
from .numbering import NumberingRegistry
from .relations import RelationshipRegistry
from .notes import parse_notes
from .parser import DocumentParser
from .builder import build_html

# Safety limits
MAX_UNCOMPRESSED_BYTES = 512 * 1024 * 1024  # 500 MB
MAX_MEMBER_BYTES = 100 * 1024 * 1024         # 100 MB


class ConversionError(Exception):
    pass


def convert(
    docx_bytes: bytes,
    mode: str = 'llm',
    track_changes: str = 'accept',
    normalize: bool = False,
) -> str:
    """
    Convert a .docx file (as bytes) to HTML5.

    Parameters
    ----------
    docx_bytes : bytes
        Raw .docx file content.
    mode : str
        'llm' | 'preserve' | 'plain'
    track_changes : str
        'accept' | 'reject' | 'preserve'
    normalize : bool
        Normalize fake headings by style analysis.

    Returns
    -------
    str
        Valid HTML5 document.
    """
    # --- Open ZIP safely ---
    try:
        zf = zipfile.ZipFile(io.BytesIO(docx_bytes))
    except zipfile.BadZipFile as exc:
        raise ConversionError(f'Not a valid .docx file: {exc}') from exc

    names = set(zf.namelist())

    def _read(path: str, required: bool = False) -> Optional[bytes]:
        if path not in names:
            if required:
                raise ConversionError(f'Missing required part: {path}')
            return None
        info = zf.getinfo(path)
        if info.file_size > MAX_MEMBER_BYTES:
            raise ConversionError(f'Member too large: {path} ({info.file_size} bytes)')
        return zf.read(path)

    # --- Load sub-documents ---
    doc_xml = _read('word/document.xml', required=True)
    styles_xml = _read('word/styles.xml')
    numbering_xml = _read('word/numbering.xml')
    rels_xml = _read('word/_rels/document.xml.rels')
    footnotes_xml = _read('word/footnotes.xml')
    endnotes_xml = _read('word/endnotes.xml')

    # --- Build registries ---
    styles = StyleRegistry()
    if styles_xml:
        styles.parse(styles_xml)

    numbering = NumberingRegistry()
    if numbering_xml:
        numbering.parse(numbering_xml)

    rels = RelationshipRegistry()
    if rels_xml:
        rels.parse(rels_xml)

    footnote_map: Dict[str, str] = {}
    if footnotes_xml:
        footnote_map = parse_notes(footnotes_xml, 'fn')

    endnote_map: Dict[str, str] = {}
    if endnotes_xml:
        endnote_map = parse_notes(endnotes_xml, 'en')

    # --- Load images (preserve mode only) ---
    image_data: Dict[str, bytes] = {}
    if mode == 'preserve':
        for name in names:
            if name.startswith('word/media/'):
                info = zf.getinfo(name)
                if info.file_size < MAX_MEMBER_BYTES:
                    image_data[name] = zf.read(name)

    zf.close()

    # --- Parse document ---
    parser = DocumentParser(
        styles=styles,
        numbering=numbering,
        rels=rels,
        footnote_map=footnote_map,
        endnote_map=endnote_map,
        image_data=image_data,
        mode=mode,
        track_changes=track_changes,
        normalize=normalize,
    )

    body_parts = parser.parse(doc_xml)

    # Collect footnotes in order
    all_notes = list(footnote_map.values()) + list(endnote_map.values())

    return build_html(body_parts, all_notes, mode)
