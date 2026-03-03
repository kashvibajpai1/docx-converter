#!/usr/bin/env python3
"""
docx2llm — DOCX → token-optimized HTML5 for LLM ingestion.

Usage
-----
    docx2llm input.docx [options] > output.html
    docx2llm input.docx --mode=preserve -o output.html
    docx2llm input.docx --mode=plain --track-changes=reject

Options
-------
    --mode=llm          Default. Images as placeholders, minimal HTML.
    --mode=preserve     Images as base64, more structure hints.
    --mode=plain        Strip images and footnotes entirely.

    --track-changes=accept   Include inserted, discard deleted. (default)
    --track-changes=reject   Discard inserted, include deleted.
    --track-changes=preserve Emit both with <ins>/<del> markup.

    --normalize         Normalize fake headings by style analysis.

    -o / --output FILE  Write to FILE instead of stdout.
    --token-count       Print token estimate to stderr.
    --compare-pandoc    Run Pandoc and compare token counts (requires pandoc).
    -h / --help         Show this help.
"""
from __future__ import annotations
import argparse
import sys
import os


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        prog='docx2llm',
        description='Convert .docx to token-optimized HTML5 for LLM ingestion.',
        add_help=True,
    )
    parser.add_argument('input', help='Path to input .docx file')
    parser.add_argument(
        '--mode',
        choices=['llm', 'preserve', 'plain'],
        default='llm',
        help='Output mode (default: llm)',
    )
    parser.add_argument(
        '--track-changes',
        dest='track_changes',
        choices=['accept', 'reject', 'preserve'],
        default='accept',
        help='How to handle tracked changes (default: accept)',
    )
    parser.add_argument(
        '--normalize',
        action='store_true',
        help='Normalize fake headings by style analysis',
    )
    parser.add_argument(
        '-o', '--output',
        default=None,
        help='Output file (default: stdout)',
    )
    parser.add_argument(
        '--token-count',
        action='store_true',
        help='Print estimated token count to stderr',
    )
    parser.add_argument(
        '--compare-pandoc',
        action='store_true',
        help='Compare token count with Pandoc output (requires pandoc in PATH)',
    )

    args = parser.parse_args(argv)

    # Read input
    try:
        with open(args.input, 'rb') as f:
            docx_bytes = f.read()
    except FileNotFoundError:
        print(f'Error: file not found: {args.input}', file=sys.stderr)
        return 1
    except OSError as exc:
        print(f'Error reading file: {exc}', file=sys.stderr)
        return 1

    # Convert
    try:
        from .converter import convert, ConversionError
        html = convert(
            docx_bytes,
            mode=args.mode,
            track_changes=args.track_changes,
            normalize=args.normalize,
        )
    except Exception as exc:  # noqa: BLE001
        print(f'Conversion error: {exc}', file=sys.stderr)
        return 1

    # Write output
    if args.output:
        try:
            with open(args.output, 'w', encoding='utf-8') as f:
                f.write(html)
        except OSError as exc:
            print(f'Error writing output: {exc}', file=sys.stderr)
            return 1
    else:
        sys.stdout.write(html)

    # Token count estimate (rough: chars / 4)
    if args.token_count or args.compare_pandoc:
        our_tokens = len(html) // 4
        print(f'[docx2llm] estimated tokens: {our_tokens:,}', file=sys.stderr)

    if args.compare_pandoc:
        _compare_pandoc(args.input, our_tokens)

    return 0


def _compare_pandoc(input_path: str, our_tokens: int) -> None:
    import subprocess
    try:
        result = subprocess.run(
            ['pandoc', '--from=docx', '--to=html5', input_path],
            capture_output=True, text=True, timeout=60,
        )
        pandoc_tokens = len(result.stdout) // 4
        reduction = (1 - our_tokens / max(pandoc_tokens, 1)) * 100
        print(f'[pandoc]   estimated tokens: {pandoc_tokens:,}', file=sys.stderr)
        print(f'[docx2llm] token reduction:  {reduction:.1f}%', file=sys.stderr)
    except FileNotFoundError:
        print('[compare-pandoc] pandoc not found in PATH', file=sys.stderr)
    except subprocess.TimeoutExpired:
        print('[compare-pandoc] pandoc timed out', file=sys.stderr)


if __name__ == '__main__':
    sys.exit(main())
