"""
Core streaming parser for /word/document.xml.

Walks the XML tree and emits HTML fragments. Uses iterative
element processing to stay memory-efficient.
"""
from __future__ import annotations
import base64
import re
import xml.etree.ElementTree as ET
from html import escape
from typing import Dict, List, Optional, Tuple

from .styles import StyleRegistry
from .numbering import NumberingRegistry
from .relations import RelationshipRegistry
from .builder import normalize_text, tag, collapse_inline, ListStack

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
WPS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
V = 'urn:schemas-microsoft-com:vml'


def _w(tag: str) -> str:
    return f'{{{W}}}{tag}'


def _r(tag: str) -> str:
    return f'{{{R_NS}}}{tag}'


class DocumentParser:
    def __init__(
        self,
        styles: StyleRegistry,
        numbering: NumberingRegistry,
        rels: RelationshipRegistry,
        footnote_map: Dict[str, str],   # fnId → html fragment
        endnote_map: Dict[str, str],
        image_data: Dict[str, bytes],   # media path → bytes
        mode: str = 'llm',
        track_changes: str = 'accept',
        normalize: bool = False,
    ) -> None:
        self.styles = styles
        self.numbering = numbering
        self.rels = rels
        self.footnote_map = footnote_map
        self.endnote_map = endnote_map
        self.image_data = image_data
        self.mode = mode
        self.track_changes = track_changes
        self.normalize = normalize

        self._parts: List[str] = []
        self._list_stack = ListStack()
        self._fn_counter = 0
        self._used_fn_ids: List[str] = []

    # ------------------------------------------------------------------
    def parse(self, xml_bytes: bytes) -> List[str]:
        try:
            root = ET.fromstring(xml_bytes)
        except ET.ParseError as exc:
            return [f'<!-- XML parse error: {escape(str(exc))} -->']

        body = root.find(f'{{{W}}}body')
        if body is None:
            body = root

        for child in body:
            local = _local(child.tag)
            if local == 'p':
                self._process_paragraph(child)
            elif local == 'tbl':
                self._flush_list()
                self._process_table(child)
            elif local == 'sdt':
                self._process_sdt(child)
            elif local == 'sectPr':
                pass  # section properties – skip

        self._flush_list()
        return self._parts

    # ------------------------------------------------------------------
    # PARAGRAPH
    # ------------------------------------------------------------------
    def _process_paragraph(self, p_el: ET.Element) -> None:
        ppr = p_el.find(_w('pPr'))
        style_id = None
        num_id = None
        ilvl = 0
        is_rtl = False

        if ppr is not None:
            pstyle = ppr.find(_w('pStyle'))
            if pstyle is not None:
                style_id = pstyle.get(_w('val'))

            # List properties
            numpr = ppr.find(_w('numPr'))
            if numpr is not None:
                ilvl_el = numpr.find(_w('ilvl'))
                numid_el = numpr.find(_w('numId'))
                if numid_el is not None:
                    num_id = numid_el.get(_w('val'))
                if ilvl_el is not None:
                    try:
                        ilvl = int(ilvl_el.get(_w('val'), '0'))
                    except ValueError:
                        ilvl = 0

            # RTL
            bidi = ppr.find(_w('bidi'))
            if bidi is not None:
                val = bidi.get(_w('val'), '1')
                is_rtl = val not in ('0', 'false')

        html_tag = self.styles.get_tag(style_id)

        # Collect inline content
        inline_html = self._collect_inline(p_el)
        inline_html = collapse_inline(inline_html)

        if not inline_html.strip():
            # Empty paragraph — skip (don't emit empty tags)
            if self.numbering.is_list_item(num_id):
                # Empty list item — still flush if needed
                self._flush_list()
            return

        # Page break detection
        if self._is_page_break(p_el):
            self._flush_list()
            self._parts.append('<hr data-page-break>\n')
            return

        if self.numbering.is_list_item(num_id):
            list_tag = self.numbering.get_list_type(num_id, ilvl)
            fragment = self._list_stack.emit(list_tag, num_id, ilvl, inline_html)
            self._parts.append(fragment)
        else:
            self._flush_list()
            attrs = {}
            if is_rtl:
                attrs['dir'] = 'rtl'
            self._parts.append(tag(html_tag, inline_html, attrs or None) + '\n')

    # ------------------------------------------------------------------
    def _is_page_break(self, p_el: ET.Element) -> bool:
        for r in p_el.findall('.//' + _w('br')):
            btype = r.get(_w('type'), '')
            if btype == 'page':
                return True
        return False

    # ------------------------------------------------------------------
    def _flush_list(self) -> None:
        if self._list_stack.is_open:
            self._parts.append(self._list_stack.close_all())

    # ------------------------------------------------------------------
    # INLINE CONTENT
    # ------------------------------------------------------------------
    def _collect_inline(self, container: ET.Element) -> str:
        parts: List[str] = []
        for child in container:
            local = _local(child.tag)
            if local == 'r':
                parts.append(self._process_run(child))
            elif local == 'hyperlink':
                parts.append(self._process_hyperlink(child))
            elif local == 'ins':
                # Track changes: inserted text
                if self.track_changes in ('accept', 'preserve'):
                    for r in child.findall('.//' + _w('r')):
                        parts.append(self._process_run(r))
            elif local == 'del':
                # Track changes: deleted text
                if self.track_changes == 'preserve':
                    parts.append('<del>')
                    for dr in child.findall('.//' + _w('delText')):
                        parts.append(escape(normalize_text(dr.text or '')))
                    parts.append('</del>')
                # accept → skip deleted
            elif local == 'fldSimple':
                parts.append(self._process_fld_simple(child))
            elif local == 'bookmarkStart':
                bid = child.get(_w('id'), '')
                bname = child.get(_w('name'), '')
                if bname:
                    parts.append(f'<a id="{escape(bname)}"></a>')
            elif local == 'drawing':
                parts.append(self._process_drawing(child))
            elif local == 'pict':
                parts.append(self._process_pict(child))
            elif local in ('proofErr', 'bookmarkEnd', 'commentRangeStart',
                           'commentRangeEnd', 'rPrChange', 'pPrChange',
                           'sectPr', 'pPr'):
                pass
            elif local == 'sdt':
                parts.append(self._process_sdt_inline(child))
        return ''.join(parts)

    # ------------------------------------------------------------------
    def _process_run(self, r_el: ET.Element) -> str:
        # Check track-changes deletion
        rpr = r_el.find(_w('rPr'))
        if rpr is not None:
            del_el = rpr.find(_w('del'))
            if del_el is not None and self.track_changes == 'accept':
                return ''

        # Separate: raw HTML fragments (refs/images) vs plain text
        html_parts: List[str] = []
        text_buf: List[str] = []

        def flush_text() -> None:
            if text_buf:
                raw = ''.join(text_buf)
                text_buf.clear()
                if rpr is not None:
                    html_parts.append(self._apply_rpr(raw, rpr))
                else:
                    html_parts.append(escape(raw))

        for child in r_el:
            local = _local(child.tag)
            if local == 't':
                text_buf.append(normalize_text(child.text or ''))
            elif local == 'delText':
                if self.track_changes == 'preserve':
                    text_buf.append(normalize_text(child.text or ''))
            elif local == 'br':
                btype = child.get(_w('type'), '')
                if btype != 'page':
                    flush_text()
                    html_parts.append('<br>')
            elif local == 'tab':
                text_buf.append(' ')
            elif local == 'footnoteReference':
                flush_text()
                fn_id = child.get(_w('id'), '')
                html_parts.append(self._footnote_ref(fn_id, 'fn'))
            elif local == 'endnoteReference':
                flush_text()
                fn_id = child.get(_w('id'), '')
                html_parts.append(self._footnote_ref(fn_id, 'en'))
            elif local == 'drawing':
                flush_text()
                html_parts.append(self._process_drawing(child))

        flush_text()
        return ''.join(html_parts)

    def _apply_rpr(self, text: str, rpr: ET.Element) -> str:
        bold = rpr.find(_w('b'))
        italic = rpr.find(_w('i'))
        underline = rpr.find(_w('u'))
        strike = rpr.find(_w('strike'))
        sup = rpr.find(_w('vertAlign'))
        sub_el = None

        if sup is not None:
            val = sup.get(_w('val'), '')
            if val == 'subscript':
                sub_el = sup
                sup = None

        if underline is not None:
            uval = underline.get(_w('val'), 'single')
            if uval == 'none':
                underline = None

        def _is_on(el: Optional[ET.Element]) -> bool:
            if el is None:
                return False
            v = el.get(_w('val'), '1')
            return v not in ('0', 'false')

        escaped = escape(text)

        if _is_on(sup):
            escaped = f'<sup>{escaped}</sup>'
        elif sub_el is not None:
            escaped = f'<sub>{escaped}</sub>'
        if underline is not None:
            escaped = f'<u>{escaped}</u>'
        if _is_on(italic):
            escaped = f'<em>{escaped}</em>'
        if _is_on(bold):
            escaped = f'<strong>{escaped}</strong>'

        return escaped

    # ------------------------------------------------------------------
    def _process_hyperlink(self, hl_el: ET.Element) -> str:
        rid = hl_el.get(f'{{{R_NS}}}id', '')
        anchor = hl_el.get(_w('anchor'), '')
        url = ''
        if rid:
            url = self.rels.get_hyperlink(rid)
        elif anchor:
            url = f'#{anchor}'

        inner = ''
        for r in hl_el:
            local = _local(r.tag)
            if local == 'r':
                inner += self._process_run(r)
        if not inner:
            return ''
        if url:
            return f'<a href="{escape(url)}">{inner}</a>'
        return inner

    # ------------------------------------------------------------------
    def _process_fld_simple(self, fld: ET.Element) -> str:
        # Try to get display text from runs inside
        parts = []
        for r in fld.findall('.//' + _w('r')):
            parts.append(self._process_run(r))
        return ''.join(parts)

    # ------------------------------------------------------------------
    def _footnote_ref(self, fn_id: str, kind: str) -> str:
        prefix = 'fn' if kind == 'fn' else 'en'
        self._fn_counter += 1
        n = self._fn_counter
        return f'<sup><a href="#{prefix}{fn_id}">{n}</a></sup>'

    # ------------------------------------------------------------------
    # IMAGES
    # ------------------------------------------------------------------
    def _process_drawing(self, drawing_el: ET.Element) -> str:
        # Find blipFill → blip → embed r:id
        for blip in drawing_el.iter(f'{{{A}}}blip'):
            rid = blip.get(f'{{{R_NS}}}embed', '')
            if rid:
                return self._render_image(rid, self._get_drawing_alt(drawing_el))
        # SmartArt / charts fallback
        return '[IMAGE]'

    def _get_drawing_alt(self, drawing_el: ET.Element) -> str:
        for docPr in drawing_el.iter('{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr'):
            alt = docPr.get('descr') or docPr.get('name') or ''
            return alt
        return ''

    def _process_pict(self, pict_el: ET.Element) -> str:
        for imagedata in pict_el.iter(f'{{{V}}}imagedata'):
            rid = imagedata.get(f'{{{R_NS}}}id', '')
            if rid:
                return self._render_image(rid, '')
        return '[IMAGE]'

    def _render_image(self, rid: str, alt: str) -> str:
        img_path = self.rels.get_image_path(rid)
        if self.mode == 'llm':
            alt_text = alt or img_path.split('/')[-1] if img_path else 'image'
            return f'[IMAGE: {escape(alt_text)}]'
        elif self.mode == 'plain':
            return ''
        else:  # preserve
            if img_path and img_path in self.image_data:
                data = self.image_data[img_path]
                mime = _guess_mime(img_path)
                b64 = base64.b64encode(data).decode('ascii')
                alt_escaped = escape(alt or '')
                return f'<img src="data:{mime};base64,{b64}" alt="{alt_escaped}">'
            return f'[IMAGE: {escape(alt or "")}]'

    # ------------------------------------------------------------------
    # TABLES
    # ------------------------------------------------------------------
    def _process_table(self, tbl_el: ET.Element) -> None:
        rows = tbl_el.findall(_w('tr'))
        if not rows:
            return

        html_rows: List[Tuple[bool, str]] = []  # (is_header, row_html)

        for row_el in rows:
            is_header = self._is_header_row(row_el)
            cells = self._process_row(row_el)
            html_rows.append((is_header, cells))

        # Find header boundary
        header_end = 0
        for i, (is_h, _) in enumerate(html_rows):
            if is_h:
                header_end = i + 1
            else:
                break

        out = ['<table>\n']
        if header_end > 0:
            out.append('<thead>\n')
            for _, row_html in html_rows[:header_end]:
                out.append(row_html)
            out.append('</thead>\n')
            body_rows = html_rows[header_end:]
        else:
            body_rows = html_rows

        if body_rows:
            out.append('<tbody>\n')
            for _, row_html in body_rows:
                out.append(row_html)
            out.append('</tbody>\n')

        out.append('</table>\n')
        self._parts.append(''.join(out))

    def _is_header_row(self, row_el: ET.Element) -> bool:
        trpr = row_el.find(_w('trPr'))
        if trpr is not None:
            trhdr = trpr.find(_w('tblHeader'))
            if trhdr is not None:
                val = trhdr.get(_w('val'), '1')
                return val not in ('0', 'false')
        return False

    def _process_row(self, row_el: ET.Element) -> str:
        is_header = self._is_header_row(row_el)
        cell_tag = 'th' if is_header else 'td'
        cells = []
        for tc in row_el.findall(_w('tc')):
            cells.append(self._process_cell(tc, cell_tag))
        return '<tr>\n' + ''.join(cells) + '</tr>\n'

    def _process_cell(self, tc_el: ET.Element, cell_tag: str) -> str:
        attrs: Dict[str, str] = {}

        # Merge spans
        tcpr = tc_el.find(_w('tcPr'))
        if tcpr is not None:
            gridspan = tcpr.find(_w('gridSpan'))
            if gridspan is not None:
                span = gridspan.get(_w('val'), '1')
                if span and int(span) > 1:
                    attrs['colspan'] = span

            vmerge = tcpr.find(_w('vMerge'))
            if vmerge is not None:
                val = vmerge.get(_w('val'), '')
                if val == 'restart':
                    attrs['rowspan'] = '2'  # best-effort, real count needs pre-pass
                elif val == '':
                    # continuation — emit hidden td
                    return ''

        # Collect cell content
        content_parts = []
        for child in tc_el:
            local = _local(child.tag)
            if local == 'p':
                inline = self._collect_inline(child)
                inline = collapse_inline(inline)
                if inline.strip():
                    content_parts.append(inline)
            elif local == 'tbl':
                # Nested table
                saved = self._parts
                self._parts = []
                self._process_table(child)
                nested = ''.join(self._parts)
                self._parts = saved
                content_parts.append(nested)

        content = '<br>'.join(content_parts)
        attr_str = ''.join(f' {k}="{escape(v)}"' for k, v in attrs.items())
        return f'<{cell_tag}{attr_str}>{content}</{cell_tag}>\n'

    # ------------------------------------------------------------------
    # SDT (Structured Document Tags)
    # ------------------------------------------------------------------
    def _process_sdt(self, sdt_el: ET.Element) -> None:
        content = sdt_el.find(_w('sdtContent'))
        if content is None:
            return
        for child in content:
            local = _local(child.tag)
            if local == 'p':
                self._process_paragraph(child)
            elif local == 'tbl':
                self._flush_list()
                self._process_table(child)

    def _process_sdt_inline(self, sdt_el: ET.Element) -> str:
        content = sdt_el.find(_w('sdtContent'))
        if content is None:
            return ''
        parts = []
        for child in content:
            local = _local(child.tag)
            if local == 'r':
                parts.append(self._process_run(child))
            elif local == 'p':
                parts.append(self._collect_inline(child))
        return ''.join(parts)


# ------------------------------------------------------------------
def _local(tag: str) -> str:
    if '}' in tag:
        return tag.split('}')[1]
    return tag


def _guess_mime(path: str) -> str:
    ext = path.rsplit('.', 1)[-1].lower()
    return {
        'png': 'image/png',
        'jpg': 'image/jpeg',
        'jpeg': 'image/jpeg',
        'gif': 'image/gif',
        'webp': 'image/webp',
        'svg': 'image/svg+xml',
        'emf': 'image/emf',
        'wmf': 'image/wmf',
    }.get(ext, 'image/png')
