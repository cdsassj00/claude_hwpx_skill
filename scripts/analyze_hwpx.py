#!/usr/bin/env python3
"""Analyze a template HWPX file and extract its style system as JSON.

Usage:
    python analyze_hwpx.py <template.hwpx> [--output style_map.json]

Extracts:
  - Page layout (paper size, margins, orientation)
  - Font definitions
  - Character properties (charPr): font size, bold, color, font refs
  - Paragraph properties (paraPr): alignment, line spacing, margins
  - Border/Fill styles (borderFill): cell borders, background colors
  - Named styles
  - Document structure: paragraphs, tables, and their style references
"""

import argparse
import json
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

# HWPX namespace map
NS = {
    'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
    'hp10': 'http://www.hancom.co.kr/hwpml/2016/paragraph',
    'hs': 'http://www.hancom.co.kr/hwpml/2011/section',
    'hc': 'http://www.hancom.co.kr/hwpml/2011/core',
    'hh': 'http://www.hancom.co.kr/hwpml/2011/head',
    'hm': 'http://www.hancom.co.kr/hwpml/2011/master-page',
    'ha': 'http://www.hancom.co.kr/hwpml/2011/app',
    'opf': 'http://www.idpf.org/2007/opf/',
}


def parse_xml_from_zip(zf, path):
    """Parse an XML file from inside a ZIP archive."""
    raw = zf.read(path)
    return ET.fromstring(raw)


def extract_page_layout(section_root):
    """Extract page layout from section's secPr."""
    info = {}
    sec_pr = section_root.find('.//hp:secPr', NS)
    if sec_pr is None:
        return info
    page_pr = sec_pr.find('hp:pagePr', NS)
    if page_pr is not None:
        info['landscape'] = page_pr.get('landscape', '')
        info['width'] = int(page_pr.get('width', '0'))
        info['height'] = int(page_pr.get('height', '0'))
        margin = page_pr.find('hp:margin', NS)
        if margin is not None:
            info['margin'] = {k: int(margin.get(k, '0')) for k in
                              ['left', 'right', 'top', 'bottom', 'header', 'footer', 'gutter']}
    return info


def extract_fonts(header_root):
    """Extract font definitions."""
    fonts = {}
    for fontface in header_root.findall('.//hh:fontface', NS):
        lang = fontface.get('lang', '')
        font_list = []
        for font in fontface.findall('hh:font', NS):
            font_list.append({
                'id': int(font.get('id', '0')),
                'face': font.get('face', ''),
                'type': font.get('type', ''),
            })
        fonts[lang] = font_list
    return fonts


def extract_char_properties(header_root):
    """Extract character property definitions."""
    props = []
    for cp in header_root.findall('.//hh:charProperties/hh:charPr', NS):
        entry = {
            'id': int(cp.get('id', '0')),
            'height': int(cp.get('height', '0')),  # font size in 1/100 pt
            'textColor': cp.get('textColor', '#000000'),
            'bold': cp.find('hh:bold', NS) is not None,
            'italic': cp.find('hh:italic', NS) is not None,
            'borderFillIDRef': int(cp.get('borderFillIDRef', '0')),
        }
        font_ref = cp.find('hh:fontRef', NS)
        if font_ref is not None:
            entry['fontRef'] = {
                'hangul': int(font_ref.get('hangul', '0')),
                'latin': int(font_ref.get('latin', '0')),
            }
        spacing = cp.find('hh:spacing', NS)
        if spacing is not None:
            entry['spacing_hangul'] = int(spacing.get('hangul', '0'))
        props.append(entry)
    return props


def extract_para_properties(header_root):
    """Extract paragraph property definitions."""
    props = []
    for pp in header_root.findall('.//hh:paraProperties/hh:paraPr', NS):
        entry = {
            'id': int(pp.get('id', '0')),
            'condense': int(pp.get('condense', '0')),
        }
        align = pp.find('hh:align', NS)
        if align is not None:
            entry['horizontal'] = align.get('horizontal', '')
            entry['vertical'] = align.get('vertical', '')

        # Try hp:switch > hp:default first, then direct children
        for source in [pp.find('.//hp:default', NS), pp]:
            margin = source.find('hh:margin', NS) if source is not None else None
            if margin is not None:
                entry['margin'] = {}
                for child in margin:
                    tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
                    entry['margin'][tag] = int(child.get('value', '0'))
                break

        for source in [pp.find('.//hp:default', NS), pp]:
            ls = source.find('hh:lineSpacing', NS) if source is not None else None
            if ls is not None:
                entry['lineSpacing'] = {
                    'type': ls.get('type', ''),
                    'value': int(ls.get('value', '0')),
                }
                break

        props.append(entry)
    return props


def extract_border_fills(header_root):
    """Extract border/fill style definitions."""
    fills = []
    for bf in header_root.findall('.//hh:borderFills/hh:borderFill', NS):
        entry = {'id': int(bf.get('id', '0')), 'borders': {}, 'fillColor': None}
        for side in ['leftBorder', 'rightBorder', 'topBorder', 'bottomBorder']:
            border_el = bf.find(f'hh:{side}', NS)
            if border_el is not None:
                entry['borders'][side] = {
                    'type': border_el.get('type', 'NONE'),
                    'width': border_el.get('width', ''),
                    'color': border_el.get('color', ''),
                }
        win_brush = bf.find('.//hc:winBrush', NS)
        if win_brush is not None:
            fc = win_brush.get('faceColor', 'none')
            if fc and fc != 'none':
                entry['fillColor'] = fc
        fills.append(entry)
    return fills


def extract_styles(header_root):
    """Extract named styles."""
    styles = []
    for st in header_root.findall('.//hh:styles/hh:style', NS):
        styles.append({
            'id': int(st.get('id', '0')),
            'type': st.get('type', ''),
            'name': st.get('name', ''),
            'engName': st.get('engName', ''),
            'paraPrIDRef': int(st.get('paraPrIDRef', '0')),
            'charPrIDRef': int(st.get('charPrIDRef', '0')),
            'nextStyleIDRef': int(st.get('nextStyleIDRef', '0')),
        })
    return styles


def collect_text(el):
    """Recursively collect text from hp:t elements."""
    parts = []
    for t in el.iter(f'{{{NS["hp"]}}}t'):
        if t.text:
            parts.append(t.text)
    return ''.join(parts)


def extract_table_structure(tbl_el):
    """Extract table structure with rows, cells, style refs, and text."""
    table = {
        'rowCnt': int(tbl_el.get('rowCnt', '0')),
        'colCnt': int(tbl_el.get('colCnt', '0')),
        'borderFillIDRef': int(tbl_el.get('borderFillIDRef', '0')),
        'rows': [],
    }
    for tr in tbl_el.findall('hp:tr', NS):
        row = []
        for tc in tr.findall('hp:tc', NS):
            cell = {
                'borderFillIDRef': int(tc.get('borderFillIDRef', '0')),
            }
            addr = tc.find('hp:cellAddr', NS)
            if addr is not None:
                cell['col'] = int(addr.get('colAddr', '0'))
                cell['row'] = int(addr.get('rowAddr', '0'))
            span = tc.find('hp:cellSpan', NS)
            if span is not None:
                cell['colSpan'] = int(span.get('colSpan', '1'))
                cell['rowSpan'] = int(span.get('rowSpan', '1'))
            sz = tc.find('hp:cellSz', NS)
            if sz is not None:
                cell['width'] = int(sz.get('width', '0'))
                cell['height'] = int(sz.get('height', '0'))

            # Extract paragraphs inside cell
            cell_paras = []
            sub = tc.find('hp:subList', NS)
            if sub is not None:
                cell['vertAlign'] = sub.get('vertAlign', '')
                for p in sub.findall('hp:p', NS):
                    cp = {
                        'paraPrIDRef': int(p.get('paraPrIDRef', '0')),
                        'styleIDRef': int(p.get('styleIDRef', '0')),
                        'runs': [],
                    }
                    for run in p.findall('hp:run', NS):
                        run_text = collect_text(run)
                        if run_text:
                            cp['runs'].append({
                                'charPrIDRef': int(run.get('charPrIDRef', '0')),
                                'text': run_text,
                            })
                    cell_paras.append(cp)
            cell['paragraphs'] = cell_paras
            row.append(cell)
        table['rows'].append(row)
    return table


def extract_document_structure(section_root):
    """Extract the full document structure: paragraphs and tables."""
    elements = []
    for p in section_root.findall('hp:p', NS):
        para = {
            'type': 'paragraph',
            'paraPrIDRef': int(p.get('paraPrIDRef', '0')),
            'styleIDRef': int(p.get('styleIDRef', '0')),
            'runs': [],
            'tables': [],
        }
        for run in p.findall('hp:run', NS):
            char_id = int(run.get('charPrIDRef', '0'))
            text = collect_text(run)

            # Check for embedded table
            tbl = run.find('hp:tbl', NS)
            if tbl is not None:
                para['tables'].append(extract_table_structure(tbl))

            if text:
                para['runs'].append({
                    'charPrIDRef': char_id,
                    'text': text,
                })
        elements.append(para)
    return elements


def analyze_hwpx(hwpx_path):
    """Main analysis: extract complete style map from HWPX."""
    result = {}
    with zipfile.ZipFile(hwpx_path, 'r') as zf:
        file_list = zf.namelist()
        result['files'] = file_list

        # Parse header.xml
        if 'Contents/header.xml' in file_list:
            header = parse_xml_from_zip(zf, 'Contents/header.xml')
            result['fonts'] = extract_fonts(header)
            result['charProperties'] = extract_char_properties(header)
            result['paraProperties'] = extract_para_properties(header)
            result['borderFills'] = extract_border_fills(header)
            result['styles'] = extract_styles(header)

        # Parse section0.xml
        if 'Contents/section0.xml' in file_list:
            section = parse_xml_from_zip(zf, 'Contents/section0.xml')
            result['pageLayout'] = extract_page_layout(section)
            result['documentStructure'] = extract_document_structure(section)

        # Content.hpf metadata
        if 'Contents/content.hpf' in file_list:
            hpf = parse_xml_from_zip(zf, 'Contents/content.hpf')
            meta = {}
            for m in hpf.findall('.//opf:meta', NS):
                name = m.get('name', '')
                if name and m.text:
                    meta[name] = m.text
            result['metadata'] = meta

    return result


def main():
    parser = argparse.ArgumentParser(description='Analyze HWPX template style system')
    parser.add_argument('hwpx', help='Path to template HWPX file')
    parser.add_argument('--output', '-o', help='Output JSON file path (default: stdout)')
    args = parser.parse_args()

    if not Path(args.hwpx).exists():
        print(f"Error: File not found: {args.hwpx}", file=sys.stderr)
        sys.exit(1)

    result = analyze_hwpx(args.hwpx)

    json_str = json.dumps(result, ensure_ascii=False, indent=2)
    if args.output:
        Path(args.output).write_text(json_str, encoding='utf-8')
        print(f"Style map saved to: {args.output}", file=sys.stderr)
    else:
        sys.stdout.buffer.write(json_str.encode('utf-8'))
        sys.stdout.buffer.write(b'\n')


if __name__ == '__main__':
    main()
