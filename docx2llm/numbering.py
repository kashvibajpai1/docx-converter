"""Parse /word/numbering.xml for list reconstruction."""
from __future__ import annotations
import xml.etree.ElementTree as ET
from typing import Dict, Optional, Tuple

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}


def _w(tag: str) -> str:
    return f'{{{W}}}{tag}'


class NumberingRegistry:
    """
    Maps (numId, ilvl) → list type ('ul' | 'ol').
    Also tracks abstract num definitions.
    """

    def __init__(self) -> None:
        # abstractNumId → {ilvl → fmt}
        self._abstract: Dict[str, Dict[int, str]] = {}
        # numId → abstractNumId
        self._num_to_abstract: Dict[str, str] = {}
        # override cache
        self._override: Dict[Tuple[str, int], str] = {}

    # ------------------------------------------------------------------
    def parse(self, xml_bytes: bytes) -> None:
        try:
            root = ET.fromstring(xml_bytes)
        except ET.ParseError:
            return

        # Parse abstractNum
        for an in root.findall('w:abstractNum', NS):
            aid = an.get(_w('abstractNumId'), '')
            levels: Dict[int, str] = {}
            for lvl in an.findall('w:lvl', NS):
                ilvl = int(lvl.get(_w('ilvl'), '0'))
                numfmt_el = lvl.find('w:numFmt', NS)
                fmt = 'bullet'
                if numfmt_el is not None:
                    fmt = numfmt_el.get(_w('val'), 'bullet')
                levels[ilvl] = fmt
            self._abstract[aid] = levels

        # Parse num → abstractNum
        for num in root.findall('w:num', NS):
            nid = num.get(_w('numId'), '')
            ref = num.find('w:abstractNumId', NS)
            if ref is not None:
                aid = ref.get(_w('val'), '')
                self._num_to_abstract[nid] = aid

            # Level overrides
            for lo in num.findall('w:lvlOverride', NS):
                ilvl = int(lo.get(_w('ilvl'), '0'))
                numfmt_el = lo.find('.//w:numFmt', NS)
                if numfmt_el is not None:
                    fmt = numfmt_el.get(_w('val'), 'bullet')
                    self._override[(nid, ilvl)] = fmt

    # ------------------------------------------------------------------
    def get_list_type(self, num_id: str, ilvl: int) -> str:
        """Return 'ul' or 'ol'."""
        if (num_id, ilvl) in self._override:
            fmt = self._override[(num_id, ilvl)]
            return _fmt_to_tag(fmt)

        aid = self._num_to_abstract.get(num_id, '')
        if aid and aid in self._abstract:
            fmt = self._abstract[aid].get(ilvl, 'bullet')
            return _fmt_to_tag(fmt)

        return 'ul'  # safe default

    def is_list_item(self, num_id: Optional[str]) -> bool:
        return bool(num_id) and num_id != '0'


def _fmt_to_tag(fmt: str) -> str:
    ordered = {
        'decimal', 'upperRoman', 'lowerRoman',
        'upperLetter', 'lowerLetter', 'ordinal',
        'cardinalText', 'ordinalText', 'hex',
        'chicago', 'ideographDigital', 'japaneseCounting',
        'aiueo', 'iroha', 'decimalFullWidth', 'decimalHalfWidth',
        'japaneseLegal', 'japaneseDigitalTenThousand',
        'decimalEnclosedCircle', 'decimalFullWidth2', 'aiueoFullWidth',
        'irohaFullWidth', 'decimalZero', 'ganada', 'chosung',
        'decimalEnclosedFullstop', 'decimalEnclosedParen',
        'decimalEnclosedCircleChinese', 'ideographEnclosedCircle',
        'ideographTraditional', 'ideographZodiac', 'ideographZodiacTraditional',
        'taiwaneseCounting', 'ideographLegalTraditional', 'taiwaneseCountingThousand',
        'taiwaneseDigital', 'chineseCounting', 'chineseLegalSimplified',
        'chineseCountingThousand', 'koreanDigital', 'koreanCounting',
        'koreanLegal', 'koreanDigital2', 'vietnameseCounting', 'russianLower',
        'russianUpper', 'none', 'numberInDash', 'hebrew1', 'hebrew2',
        'arabicAlpha', 'arabicAbjad', 'hindiVowels', 'hindiConsonants',
        'hindiNumbers', 'hindiCounting', 'thaiLetters', 'thaiNumbers',
        'thaiCounting',
    }
    if fmt in ordered and fmt != 'none':
        return 'ol'
    return 'ul'
