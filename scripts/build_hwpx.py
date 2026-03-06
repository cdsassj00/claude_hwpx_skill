#!/usr/bin/env python3
"""Build a new HWPX file from a template and a content specification.

Usage:
    python build_hwpx.py <template.hwpx> <content_spec.json> <output.hwpx>

Two modes:
  1. "replace" mode (default): Positional text replacement in template XML.
     Preserves ALL layout data (linesegarray, margins, run styles, etc.)
  2. "generate" mode: Build section0.xml from scratch (may lose layout hints)

Replace mode content_spec format:
{
  "mode": "replace",
  "texts": [
    "첫번째 텍스트",    // index 0: replaces 1st <hp:t> content
    "두번째 텍스트",    // index 1: replaces 2nd <hp:t> content
    null,               // null = keep original text
    "",                 // empty string = clear text (keep run structure)
    ...
  ]
}

Run `python build_hwpx.py --dump-texts <template.hwpx>` to see all text
positions and their current content for building the replacement list.

Generate mode content_spec format:
{
  "mode": "generate",
  "paragraphs": [ ... ]
}
"""

import argparse
import json
import sys
import zipfile
import re
from pathlib import Path

# ============================================================
# REPLACE MODE: Positional text replacement preserving all XML structure
# ============================================================

def _find_hp_t_tags(xml_str):
    """Find all <hp:t>...</hp:t> and <hp:t/> tags using regex.

    Returns list of (start, end, is_self_closing) tuples.
    Uses regex to avoid matching <hp:tbl>, <hp:tc>, <hp:tr> etc.
    """
    tags = []
    # Match <hp:t> (open) or <hp:t/> (self-closing)
    for m in re.finditer(r'<hp:t(/?)>', xml_str):
        if m.group(1) == '/':
            tags.append((m.start(), m.end(), True))
        else:
            # Find closing </hp:t>
            t_close = xml_str.find('</hp:t>', m.end())
            if t_close != -1:
                tags.append((m.start(), t_close + 7, False))
    return tags


def replace_texts_positional(xml_str, texts):
    """Replace <hp:t>...</hp:t> contents by position index.

    Finds all non-self-closing <hp:t> tags in order and replaces their
    text content with the corresponding entry from the texts list.
    - null entries keep the original text
    - Empty string entries clear the text but keep the run structure
    - Self-closing <hp:t/> tags are skipped (not counted)

    Everything else (XML structure, attributes, linesegarray, runs) is
    preserved exactly as-is.
    """
    tags = _find_hp_t_tags(xml_str)
    result = []
    prev_end = 0
    text_idx = 0

    for tag_start, tag_end, is_self_closing in tags:
        # Add everything before this tag
        result.append(xml_str[prev_end:tag_start])

        if is_self_closing:
            result.append('<hp:t/>')
        else:
            # Extract original text between <hp:t> and </hp:t>
            content_start = tag_start + 6  # len('<hp:t>')
            content_end = tag_end - 7      # len('</hp:t>')
            orig_text = xml_str[content_start:content_end]

            if text_idx < len(texts) and texts[text_idx] is not None:
                result.append(f'<hp:t>{escape_xml(texts[text_idx])}</hp:t>')
            else:
                result.append(f'<hp:t>{orig_text}</hp:t>')
            text_idx += 1

        prev_end = tag_end

    # Add remaining content after last tag
    result.append(xml_str[prev_end:])
    return ''.join(result)


def dump_texts(xml_str):
    """Extract all <hp:t> text positions for building replacement list."""
    tags = _find_hp_t_tags(xml_str)
    entries = []
    idx = 0
    for tag_start, tag_end, is_self_closing in tags:
        if is_self_closing:
            continue
        content_start = tag_start + 6
        content_end = tag_end - 7
        text = xml_str[content_start:content_end]
        text = (text.replace('&amp;', '&').replace('&lt;', '<')
                .replace('&gt;', '>').replace('&quot;', '"'))
        entries.append({'index': idx, 'text': text})
        idx += 1
    return entries


def find_closing_tag(xml, start, tag):
    """Find the closing tag position (end of </tag>) for a given opening tag position."""
    depth = 0
    i = start
    while i < len(xml):
        # Opening tag (not self-closing)
        open_match = re.match(rf'<{re.escape(tag)}[\s>]', xml[i:])
        if open_match:
            # Check if self-closing
            tag_end = xml.find('>', i)
            if tag_end != -1 and xml[tag_end - 1] == '/':
                # Self-closing, skip
                i = tag_end + 1
                continue
            depth += 1
            i += 1
            continue

        # Closing tag
        close_match = re.match(rf'</{re.escape(tag)}>', xml[i:])
        if close_match:
            depth -= 1
            if depth == 0:
                return i + len(f'</{tag}>')
            i += 1
            continue

        i += 1

    return -1


def escape_xml(text):
    """Escape text for XML content."""
    return (text
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;'))


# ============================================================
# GENERATE MODE: Build from scratch (legacy)
# ============================================================

HWPX_NS_DECL = (
    'xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" '
    'xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" '
    'xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph" '
    'xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" '
    'xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" '
    'xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" '
    'xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history" '
    'xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page" '
    'xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" '
    'xmlns:dc="http://purl.org/dc/elements/1.1/" '
    'xmlns:opf="http://www.idpf.org/2007/opf/" '
    'xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart" '
    'xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar" '
    'xmlns:epub="http://www.idpf.org/2007/ops" '
    'xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"'
)


def build_run_gen(run_spec):
    char_id = run_spec.get('charPrIDRef', 0)
    text = escape_xml(run_spec.get('text', ''))
    return f'<hp:run charPrIDRef="{char_id}"><hp:t>{text}</hp:t></hp:run>'


def build_paragraph_gen(para_spec):
    para_id = para_spec.get('id', '2147483648')
    para_pr = para_spec.get('paraPrIDRef', 0)
    style_id = para_spec.get('styleIDRef', 0)
    runs_xml = ''.join(build_run_gen(r) for r in para_spec.get('runs', []))
    return (f'<hp:p id="{para_id}" paraPrIDRef="{para_pr}" styleIDRef="{style_id}" '
            f'pageBreak="0" columnBreak="0" merged="0">{runs_xml}</hp:p>')


def build_cell_gen(cell_spec, col_idx, row_idx, cell_widths):
    bfid = cell_spec.get('borderFillIDRef', 10)
    vert_align = cell_spec.get('vertAlign', 'CENTER')
    col_span = cell_spec.get('colSpan', 1)
    row_span = cell_spec.get('rowSpan', 1)
    width = cell_spec.get('width', cell_widths[col_idx] if col_idx < len(cell_widths) else 10000)
    height = cell_spec.get('height', 3000)
    header = cell_spec.get('header', 0)
    cm = cell_spec.get('cellMargin', {})
    cm_l, cm_r = cm.get('left', 141), cm.get('right', 141)
    cm_t, cm_b = cm.get('top', 141), cm.get('bottom', 141)
    has_margin = 1 if cm else 0

    paras_xml = ''.join(build_paragraph_gen(p) for p in cell_spec.get('paragraphs', []))
    if not paras_xml:
        paras_xml = '<hp:p id="2147483648" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"/>'

    return (
        f'<hp:tc name="" header="{header}" hasMargin="{has_margin}" protect="0" editable="0" '
        f'dirty="0" borderFillIDRef="{bfid}">'
        f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" '
        f'vertAlign="{vert_align}" linkListIDRef="0" linkListNextIDRef="0" '
        f'textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">'
        f'{paras_xml}'
        f'</hp:subList>'
        f'<hp:cellAddr colAddr="{col_idx}" rowAddr="{row_idx}"/>'
        f'<hp:cellSpan colSpan="{col_span}" rowSpan="{row_span}"/>'
        f'<hp:cellSz width="{width}" height="{height}"/>'
        f'<hp:cellMargin left="{cm_l}" right="{cm_r}" top="{cm_t}" bottom="{cm_b}"/>'
        f'</hp:tc>')


def build_table_gen(table_spec):
    row_cnt = table_spec.get('rowCnt', 0)
    col_cnt = table_spec.get('colCnt', 0)
    bf_id = table_spec.get('borderFillIDRef', 4)
    cell_widths = table_spec.get('cellWidths', [])
    total_width = sum(cell_widths) if cell_widths else 48018
    if not cell_widths:
        w = total_width // col_cnt
        cell_widths = [w] * col_cnt
        cell_widths[-1] = total_width - w * (col_cnt - 1)

    rows_xml = ''
    for ri, row_spec in enumerate(table_spec.get('rows', [])):
        cells_xml = ''.join(build_cell_gen(c, ci, ri, cell_widths)
                            for ci, c in enumerate(row_spec.get('cells', [])))
        rows_xml += f'<hp:tr>{cells_xml}</hp:tr>'

    total_height = sum(
        max((c.get('height', 3000) for c in r.get('cells', [{'height': 3000}])), default=3000)
        for r in table_spec.get('rows', []))
    tbl_id = table_spec.get('id', '2065086322')
    z_order = table_spec.get('zOrder', 0)
    no_adjust = table_spec.get('noAdjust', 1)
    page_break = table_spec.get('pageBreak', 'NONE')
    out_m = table_spec.get('outMargin', {})
    in_m = table_spec.get('inMargin', {})

    return (
        f'<hp:tbl id="{tbl_id}" zOrder="{z_order}" numberingType="TABLE" '
        f'textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" '
        f'pageBreak="{page_break}" repeatHeader="1" '
        f'rowCnt="{row_cnt}" colCnt="{col_cnt}" cellSpacing="0" '
        f'borderFillIDRef="{bf_id}" noAdjust="{no_adjust}">'
        f'<hp:sz width="{total_width}" widthRelTo="ABSOLUTE" '
        f'height="{total_height}" heightRelTo="ABSOLUTE" protect="0"/>'
        f'<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" '
        f'holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="PARA" vertAlign="TOP" '
        f'horzAlign="LEFT" vertOffset="0" horzOffset="0"/>'
        f'<hp:outMargin left="{out_m.get("left",140)}" right="{out_m.get("right",140)}" '
        f'top="{out_m.get("top",140)}" bottom="{out_m.get("bottom",140)}"/>'
        f'<hp:inMargin left="{in_m.get("left",140)}" right="{in_m.get("right",140)}" '
        f'top="{in_m.get("top",140)}" bottom="{in_m.get("bottom",140)}"/>'
        f'{rows_xml}</hp:tbl>')


def build_section_gen(content_spec, sec_pr_xml):
    parts = [f'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hs:sec {HWPX_NS_DECL}>']
    for para_spec in content_spec.get('paragraphs', []):
        ptype = para_spec.get('type', 'paragraph')
        if ptype == 'table_host':
            pr = para_spec.get('paraPrIDRef', 0)
            cr = para_spec.get('charPrIDRef', 0)
            si = para_spec.get('styleIDRef', 0)
            parts.append(f'<hp:p id="2147483648" paraPrIDRef="{pr}" styleIDRef="{si}" '
                         f'pageBreak="0" columnBreak="0" merged="0">')
            for run in para_spec.get('preRuns', []):
                parts.append(build_run_gen(run))
            parts.append(f'<hp:run charPrIDRef="{cr}">')
            parts.append(build_table_gen(para_spec['table']))
            parts.append('<hp:t/></hp:run>')
            for run in para_spec.get('postRuns', []):
                parts.append(build_run_gen(run))
            parts.append('</hp:p>')
        elif ptype == 'secpr_host':
            pr = para_spec.get('paraPrIDRef', 0)
            cr = para_spec.get('charPrIDRef', 0)
            si = para_spec.get('styleIDRef', 0)
            parts.append(f'<hp:p id="0" paraPrIDRef="{pr}" styleIDRef="{si}" '
                         f'pageBreak="0" columnBreak="0" merged="0">')
            parts.append(f'<hp:run charPrIDRef="{cr}">{sec_pr_xml}')
            parts.append('<hp:ctrl><hp:colPr id="" type="NEWSPAPER" layout="LEFT" '
                         'colCount="1" sameSz="1" sameGap="0"/></hp:ctrl></hp:run>')
            for run in para_spec.get('runs', []):
                parts.append(build_run_gen(run))
            parts.append('</hp:p>')
        else:
            parts.append(build_paragraph_gen(para_spec))
    parts.append('</hs:sec>')
    return ''.join(parts)


def extract_secpr(zf):
    raw = zf.read('Contents/section0.xml').decode('utf-8')
    start = raw.find('<hp:secPr')
    if start == -1:
        return ''
    end = raw.find('</hp:secPr>', start)
    if end != -1:
        return raw[start:end + len('</hp:secPr>')]
    return ''


# ============================================================
# Main build logic
# ============================================================

def generate_preview_text(xml_str):
    """Generate preview text from XML content."""
    texts = re.findall(r'<hp:t>(.*?)</hp:t>', xml_str)
    lines = []
    for t in texts:
        t = t.replace('&lt;', '<').replace('&gt;', '>').replace('&amp;', '&').replace('&quot;', '"')
        if t.strip():
            lines.append(t.strip())
    return '\n'.join(lines[:50])


def build_hwpx(template_path, content_spec, output_path):
    """Build a new HWPX from template + content spec."""
    mode = content_spec.get('mode', 'replace')

    with zipfile.ZipFile(template_path, 'r') as zf_in:
        if mode == 'replace':
            orig_section = zf_in.read('Contents/section0.xml').decode('utf-8')
            texts = content_spec.get('texts', [])
            section_xml = replace_texts_positional(orig_section, texts)
        else:
            sec_pr_xml = extract_secpr(zf_in)
            section_xml = build_section_gen(content_spec, sec_pr_xml)

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf_out:
            for item in zf_in.infolist():
                if item.filename == 'Contents/section0.xml':
                    zf_out.writestr(item.filename, section_xml.encode('utf-8'))
                elif item.filename == 'Preview/PrvText.txt':
                    preview = generate_preview_text(section_xml)
                    zf_out.writestr(item.filename, preview.encode('utf-8'))
                else:
                    zf_out.writestr(item.filename, zf_in.read(item.filename))


def main():
    parser = argparse.ArgumentParser(description='Build HWPX from template + content spec')
    parser.add_argument('template', help='Path to template HWPX file')
    parser.add_argument('content', nargs='?', help='Path to content specification JSON file')
    parser.add_argument('output', nargs='?', help='Output HWPX file path')
    parser.add_argument('--dump-texts', action='store_true',
                        help='Dump all text positions from template (for building replacement list)')
    args = parser.parse_args()

    if not Path(args.template).exists():
        print(f"Error: Template not found: {args.template}", file=sys.stderr)
        sys.exit(1)

    if args.dump_texts:
        with zipfile.ZipFile(args.template, 'r') as zf:
            xml = zf.read('Contents/section0.xml').decode('utf-8')
        entries = dump_texts(xml)
        json.dump(entries, sys.stdout, ensure_ascii=False, indent=2)
        return

    if not args.content or not args.output:
        parser.error('content and output are required when not using --dump-texts')

    if not Path(args.content).exists():
        print(f"Error: Content spec not found: {args.content}", file=sys.stderr)
        sys.exit(1)

    content_spec = json.loads(Path(args.content).read_text(encoding='utf-8'))
    build_hwpx(args.template, content_spec, args.output)
    print(f"HWPX created: {args.output}", file=sys.stderr)


if __name__ == '__main__':
    main()
