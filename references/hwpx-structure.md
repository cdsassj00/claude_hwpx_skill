# HWPX File Format Reference

## Overview

HWPX is Hancom Office's XML-based document format (successor to binary HWP). It is a ZIP archive containing XML files.

## File Structure

```
document.hwpx (ZIP)
├── mimetype                    # "application/hwp+zip"
├── version.xml                 # Hancom Office version info
├── settings.xml                # Document settings
├── META-INF/
│   ├── container.xml           # Root file references
│   ├── container.rdf           # RDF metadata
│   └── manifest.xml            # File manifest
├── Contents/
│   ├── content.hpf             # Document metadata (OPF format)
│   ├── header.xml              # ALL style definitions (fonts, charPr, paraPr, borderFill, styles)
│   └── section0.xml            # Body content (paragraphs, tables, images)
├── BinData/                    # Embedded images/binaries
└── Preview/
    ├── PrvText.txt             # Plain text preview
    └── PrvImage.png            # Thumbnail preview
```

## Key XML Namespaces

| Prefix | URI | Used for |
|--------|-----|----------|
| `hp:` | `.../2011/paragraph` | Paragraphs, runs, tables, cells |
| `hh:` | `.../2011/head` | Header definitions (fonts, charPr, paraPr, styles) |
| `hs:` | `.../2011/section` | Section root element |
| `hc:` | `.../2011/core` | Core elements (margins, fills) |

## Style Reference System

The style system uses numeric ID references. Content in section0.xml references style definitions in header.xml.

### Character Properties (`charPr`)

Defined in `header.xml` under `<hh:charProperties>`. Referenced by `charPrIDRef` attribute on `<hp:run>`.

Key attributes:
- `height`: Font size in 1/100 pt (e.g., 1200 = 12pt)
- `textColor`: Color in hex (e.g., "#000000")
- `bold`: Present as `<hh:bold/>` child element
- `fontRef`: Maps to font IDs per language (hangul, latin, etc.)
- `spacing`: Character spacing adjustment

### Paragraph Properties (`paraPr`)

Defined in `header.xml` under `<hh:paraProperties>`. Referenced by `paraPrIDRef` attribute on `<hp:p>`.

Key attributes:
- `align horizontal`: LEFT, CENTER, RIGHT, JUSTIFY
- `lineSpacing`: Type (PERCENT/FIXED) and value
- `margin`: left, right, intent (first line indent), prev (space before), next (space after)

### Border/Fill Styles (`borderFill`)

Defined in `header.xml` under `<hh:borderFills>`. Referenced by `borderFillIDRef` on tables and cells.

Key attributes:
- Border sides: leftBorder, rightBorder, topBorder, bottomBorder
- Each border has: type (NONE/SOLID), width, color
- Fill: `<hc:winBrush faceColor="..."/>` for background color

## Content Structure (section0.xml)

### Basic Paragraph

```xml
<hp:p id="2147483648" paraPrIDRef="25" styleIDRef="0"
      pageBreak="0" columnBreak="0" merged="0">
  <hp:run charPrIDRef="13">
    <hp:t>Text content here</hp:t>
  </hp:run>
</hp:p>
```

### Table Structure

Tables are embedded inside a paragraph's `<hp:run>`:

```xml
<hp:p paraPrIDRef="25" ...>
  <hp:run charPrIDRef="14">
    <hp:tbl rowCnt="3" colCnt="4" borderFillIDRef="4" ...>
      <hp:sz width="48018" ... />
      <hp:pos treatAsChar="1" ... />
      <hp:outMargin ... />
      <hp:inMargin ... />
      <hp:tr>
        <hp:tc borderFillIDRef="6">
          <hp:subList vertAlign="CENTER" ...>
            <hp:p paraPrIDRef="26" ...>
              <hp:run charPrIDRef="15"><hp:t>Cell text</hp:t></hp:run>
            </hp:p>
          </hp:subList>
          <hp:cellAddr colAddr="0" rowAddr="0"/>
          <hp:cellSpan colSpan="1" rowSpan="1"/>
          <hp:cellSz width="2718" height="4014"/>
          <hp:cellMargin left="141" right="141" top="141" bottom="141"/>
        </hp:tc>
      </hp:tr>
    </hp:tbl>
    <hp:t/>
  </hp:run>
</hp:p>
```

### Section Properties (secPr)

The first paragraph must contain `<hp:secPr>` inside its first `<hp:run>`, defining page layout:

```xml
<hp:secPr textDirection="HORIZONTAL" ...>
  <hp:pagePr landscape="WIDELY" width="59528" height="84186">
    <hp:margin left="5669" right="5669" top="3600" bottom="3600" .../>
  </hp:pagePr>
  ...
</hp:secPr>
```

Common paper sizes (HWPUNIT):
- A4 Portrait: width=59528, height=84186, landscape="NARROWLY"
- A4 Landscape: width=59528, height=84186, landscape="WIDELY"

## Critical Table Attributes

These attributes on `<hp:tbl>` significantly affect rendering:

| Attribute | Values | Impact |
|-----------|--------|--------|
| `noAdjust` | **1** (recommended) / 0 | 1=fixed column widths, 0=auto-adjust (can break layout) |
| `pageBreak` | **NONE** / CELL / TABLE | NONE=keep table together, CELL=allow break at cell boundary |

**Title/header tables** (1x1 decorative boxes): use `noAdjust="0"`, `pageBreak="CELL"`
**Content tables** (multi-row data): use `noAdjust="1"`, `pageBreak="NONE"`

### Cell Margins

Header cells typically use larger margins than body cells:
- Header cells: `<hp:cellMargin left="494" right="494" top="0" bottom="0"/>`
- Body cells: `<hp:cellMargin left="141" right="141" top="141" bottom="141"/>`

### Table Margins

Title tables typically have larger inner margins:
- Title: `<hp:inMargin left="566" right="566" top="1133" bottom="1133"/>`
- Content: `<hp:inMargin left="140" right="140" top="140" bottom="140"/>`

## Common Korean Document Patterns

### Bullet Characters

| Character | Unicode | Common use |
|-----------|---------|------------|
| □ | U+25A1 | Section headers / key info labels |
| ○ | U+25CB | Sub-items |
| ․ | U+2024 | List items (one dot leader) |
| - | U+002D | Sub-list items |
| ※ | U+203B | Notes/remarks |

### Typical Style Mapping for Government Documents

| Element | Typical charPr traits | Typical paraPr traits |
|---------|----------------------|----------------------|
| Document title | 16-20pt, bold | CENTER align |
| Section label (□) | 12-14pt, bold | JUSTIFY, no indent |
| Table header | 11-12pt, bold, colored bg | CENTER align |
| Table body | 10-12pt, normal | LEFT/JUSTIFY |
| Bullet items (․) | 10-12pt | LEFT, left margin indent |
| Sub-items (-) | 10-11pt | LEFT, deeper indent |

## content_spec.json Format (Replace Mode)

The `build_hwpx.py` script uses positional text replacement. First dump text positions:

```bash
python build_hwpx.py --dump-texts template.hwpx > text_positions.json
```

Output (example):
```json
[
  {"index": 0, "text": "<붙임>"},
  {"index": 1, "text": "「AI시대 교육시간표」"},
  {"index": 2, "text": "□"},
  {"index": 3, "text": " 교육기간 : 2026. 3. 31.(5H) "},
  ...
]
```

Then build content_spec.json with a `texts` list mapping to these indices:

```json
{
  "mode": "replace",
  "texts": [
    "<붙임>",
    "새 문서 제목",
    "□",
    " 교육기간 : 2026. 5. 12.~13.(10H) ",
    null,
    null
  ]
}
```

### Replace Mode Rules

- Each entry in `texts` replaces the `<hp:t>` content at that position index
- `null` = keep original text unchanged
- `""` = clear text but keep XML run structure
- List can be shorter than total positions (remaining keep original)
- ALL XML structure is preserved: linesegarray, margins, cell sizes, run attributes
- Multi-run paragraphs: each run's text is a separate index. Match run count to preserve per-run styling

### Generate Mode (Legacy)

For building section0.xml from scratch (loses linesegarray layout hints):

```json
{
  "mode": "generate",
  "paragraphs": [
    {"type": "secpr_host", "paraPrIDRef": 27, "charPrIDRef": 24},
    {"type": "table_host", "paraPrIDRef": 30, "charPrIDRef": 25, "table": {...}},
    {"paraPrIDRef": 25, "runs": [{"charPrIDRef": 13, "text": "□"}]}
  ]
}
```
