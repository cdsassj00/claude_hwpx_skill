---
name: hwpx-writer
description: Create and edit HWPX (Hancom/한글) documents by combining a template HWPX file's style and layout with new content from a markdown file. Use when the user provides a .hwpx template file and a .md content file (or describes content to write), and wants to generate a new .hwpx document that preserves the original template's formatting, fonts, table structure, bullet styles, and page layout while replacing the content. Also use when the user asks to create, modify, or generate Korean government-style documents in HWPX format.
---

# HWPX Writer

Generate HWPX documents that preserve a template's style while replacing content with new material.

## Workflow

### Step 1: Dump text positions from the template

Extract the ordered list of all replaceable text slots:

```bash
python scripts/build_hwpx.py --dump-texts <template.hwpx> > text_positions.json
```

Read the output JSON. Each entry has an `index` and `text` field showing the current text at that position. This is the map for building replacements.

### Step 2: Analyze the template (optional, for understanding styles)

If you need to understand the template's style system (fonts, colors, alignments):

```bash
python scripts/analyze_hwpx.py <template.hwpx> -o style_map.json
```

This extracts charProperties, paraProperties, borderFills, and documentStructure. Useful for understanding which text positions correspond to titles, headers, bullets, etc.

### Step 3: Read the markdown content

Parse the user's markdown file. Map markdown elements to template text positions:

| Markdown | Typical template position |
|----------|--------------------------|
| `# Title` | The title text inside the decorative 1x1 table |
| `> metadata` | The □ info line (training period, participants, location) |
| Table headers | Header row cells (교시, 강의시간, 교과목, 교육내용) |
| Table data rows | Data cells with time, subject, and bullet content |
| `- item` | Bullet text (starts with ․) |
| `  - sub` | Sub-item text (starts with -) |

### Step 4: Build content_spec.json

Create a JSON file with the `"replace"` mode and a `"texts"` list.

```json
{
  "mode": "replace",
  "texts": [
    "<붙임>",
    "새 문서 제목",
    "□",
    " 교육기간 : 2026. 5. 12.~13.(10H) ",
    "□",
    " 교육인원 : 30명 내외 ",
    null,
    null
  ]
}
```

Rules:
- Each entry replaces the `<hp:t>` tag at that position index (from Step 1)
- `null` = keep original text unchanged
- `""` (empty string) = clear the text but keep the XML run structure
- The list length can be shorter than total positions (remaining positions keep original text)
- **Critical**: The template's XML structure (linesegarray, margins, cell dimensions, run styles) is 100% preserved. Only text content inside `<hp:t>` tags changes.

For complex paragraphs with multiple runs (e.g., "보고서 요약: " + "글로벌 스마트팜" + ...), each run's text is a separate index. Match the original run count to preserve styling per-run.

### Step 5: Generate the HWPX

```bash
python scripts/build_hwpx.py <template.hwpx> content_spec.json <output.hwpx>
```

This copies the entire template, replacing only the text content in `Contents/section0.xml`. All styles (header.xml), images (BinData/), and metadata remain intact.

### Step 6: Verify

Confirm the output HWPX is valid:
```bash
python -c "import zipfile; zf=zipfile.ZipFile('output.hwpx'); print([f.filename for f in zf.infolist()]); zf.close()"
```

Inform the user the file is ready and can be opened in Hancom Office.

## Scripts

- **`scripts/analyze_hwpx.py`**: Extract complete style map from a template HWPX as JSON.
- **`scripts/build_hwpx.py`**: Build a new HWPX from template + content_spec.json. Two modes:
  - `--dump-texts <template.hwpx>`: List all text positions for building replacement list
  - `<template.hwpx> <content_spec.json> <output.hwpx>`: Build the output file

## References

- **[references/hwpx-structure.md](references/hwpx-structure.md)**: HWPX format specification, XML structure, style reference system, and common Korean document patterns.
