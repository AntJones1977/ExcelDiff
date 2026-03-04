# ExcelDiff - Architecture Document

## 1. System Overview

```
┌─────────────────────────────────────────────────────────────────┐
│                      ExcelDiff System                            │
│                                                                  │
│  ┌──────────┐  ┌──────────┐  ┌──────────┐  ┌──────────┐        │
│  │  CLI     │  │  Web UI  │  │  Python  │  │ CI / CD  │        │
│  │ (Click)  │  │ (Flask)  │  │  API     │  │ Pipeline │        │
│  └────┬─────┘  └────┬─────┘  └────┬─────┘  └────┬─────┘        │
│       │              │              │              │              │
│  ┌────┴──────────────┴──────────────┴──────────────┴─────┐      │
│  │                  Comparison Engine                      │      │
│  │  compare(left.xlsx, right.xlsx) → WorkbookDiff          │      │
│  └────────────────────┬────────────────────────────────────┘      │
│                       │                                           │
│  ┌────────────────────┴────────────────────────────────────┐      │
│  │                  Parsing Layer                           │      │
│  │  openpyxl → WorkbookData (frozen, immutable snapshots)  │      │
│  └─────────────────────────────────────────────────────────┘      │
│                       │                                           │
│  ┌────────────────────┴────────────────────────────────────┐      │
│  │                  Renderer Layer                          │      │
│  │  Terminal (Rich) │ HTML (Jinja2) │ Excel (openpyxl)     │      │
│  └─────────────────────────────────────────────────────────┘      │
│                                                                   │
│  ┌───────────────────────────────────────────────────────────┐    │
│  │              .xlsx Files (openpyxl)                         │    │
│  │  data_only=False → raw formulas, not cached values         │    │
│  └───────────────────────────────────────────────────────────┘    │
└───────────────────────────────────────────────────────────────────┘
```

## 2. Core Design Principles

| Principle | Implementation |
|-----------|---------------|
| **Formula-first** | Compares underlying formulas, not computed values. `openpyxl.load_workbook(data_only=False)` |
| **Immutable snapshots** | All parsed data stored in `frozen=True` dataclasses, completely decoupled from openpyxl |
| **Three-level cascade** | Workbook → Sheet → Cell comparison, each level independent |
| **Pluggable output** | Engine returns a `WorkbookDiff`; renderers are swappable (terminal, HTML, Excel) |
| **CI-friendly** | Exit code 0 = identical, 1 = differences found |

## 3. Project Structure

```
ExcelDiff/
├── pyproject.toml                          # Setuptools config, deps, entry point
├── ARCHITECTURE.md                         # This document
├── docs/
│   └── PROCESS_MAPS.md                    # Mermaid process diagrams
│
├── src/exceldiff/
│   ├── __init__.py                        # __version__ = "0.1.0"
│   ├── cli.py                             # Click CLI: diff, inspect, web
│   │
│   ├── parsing/                           # Layer 1: .xlsx → immutable data models
│   │   ├── models.py                      # CellKind, CellData, SheetData, WorkbookData
│   │   ├── cell_parser.py                 # Cell extraction + formula detection
│   │   ├── sheet_parser.py                # Sheet structure extraction
│   │   └── workbook_parser.py             # Top-level workbook loader
│   │
│   ├── comparison/                        # Layer 2: diff engine
│   │   ├── models.py                      # ChangeType, CellDiff, SheetDiff, WorkbookDiff
│   │   ├── cell_comparator.py             # Cell-level formula + style comparison
│   │   ├── sheet_comparator.py            # Sheet-level structural comparison
│   │   ├── workbook_comparator.py         # Workbook-level sheet set operations
│   │   └── engine.py                      # Top-level orchestrator
│   │
│   ├── renderers/                         # Layer 3: output formatters
│   │   ├── terminal.py                    # Rich colour-coded console output
│   │   ├── html.py                        # Jinja2 self-contained HTML report
│   │   ├── excel.py                       # Annotated .xlsx with comments + colour fills
│   │   └── templates/
│   │       ├── report.html                # HTML report template
│   │       └── style.css                  # Inline CSS for HTML report
│   │
│   └── web/                               # Layer 4: browser-based UI
│       ├── app.py                         # Flask application factory
│       ├── routes.py                      # Upload, compare, results, download
│       └── templates/
│           ├── base.html                  # Navbar + flash messages layout
│           ├── upload.html                # Two-file upload form
│           ├── results.html               # Summary cards + sheet table
│           └── sheet_detail.html          # Per-sheet drill-down
│
└── tests/
    ├── conftest.py                        # Programmatic .xlsx fixture builder
    ├── test_parsing/                      # 11 parser tests
    ├── test_comparison/                   # 19 engine integration tests
    ├── test_renderers/                    # 11 renderer tests (HTML + Excel)
    └── test_web/                          # 7 Flask route tests
```

## 4. Data Flow Architecture

```
.xlsx File (LEFT)                              .xlsx File (RIGHT)
       │                                              │
       ▼                                              ▼
  openpyxl.load_workbook                      openpyxl.load_workbook
  (data_only=False)                           (data_only=False)
       │                                              │
       ▼                                              ▼
  workbook_parser.py                           workbook_parser.py
  ├── sheet_parser.py                          ├── sheet_parser.py
  │   └── cell_parser.py                       │   └── cell_parser.py
  │       └── Formula detection                │       └── Formula detection
  │       └── Style normalisation              │       └── Style normalisation
       │                                              │
       ▼                                              ▼
  WorkbookData (frozen)                       WorkbookData (frozen)
  ├── sheet_names                              ├── sheet_names
  ├── sheets: {name → SheetData}               ├── sheets: {name → SheetData}
  │   ├── cells: {coord → CellData}            │   ├── cells: {coord → CellData}
  │   ├── merged_cells                         │   ├── merged_cells
  │   └── ...structural props                  │   └── ...structural props
  └── defined_names                            └── defined_names
       │                                              │
       └───────────────┬──────────────────────────────┘
                       │
                       ▼
              compare_workbooks()
              ├── Set operations on sheet names
              ├── compare_sheets() for each shared sheet
              │   ├── Structural property diffs
              │   └── compare_cells() for each coordinate
              │       ├── Formula comparison (case-insensitive)
              │       └── Style comparison (per-field)
              └── Defined name comparison
                       │
                       ▼
              WorkbookDiff (frozen)
              ├── sheet_diffs: [SheetDiff]
              │   ├── structural_diffs: [PropertyDiff]
              │   └── cell_diffs: [CellDiff]
              │       ├── formula_changed: bool
              │       └── properties: [PropertyDiff]
              ├── sheets_only_in_left / right
              └── defined_name_diffs
                       │
          ┌────────────┼────────────┬──────────────┐
          ▼            ▼            ▼              ▼
     Terminal       HTML         Excel          Web UI
     Renderer      Renderer     Renderer       (Flask)
     (Rich)        (Jinja2)     (openpyxl)     Delegates
                                               to renderers
```

## 5. Parsing Data Models

All models are `@dataclass(frozen=True)` — immutable once created, safe for comparison.

```
WorkbookData
├── file_path: str
├── sheet_names: tuple[str, ...]
├── sheets: dict[str, SheetData]
│   ├── name: str
│   ├── dimensions: str ("A1:D10")
│   ├── max_row / max_column: int
│   ├── merged_cells: tuple[str, ...]
│   ├── freeze_panes: str | None
│   ├── column_widths: dict[str, float]
│   ├── row_heights: dict[int, float]
│   ├── data_validations: tuple[DataValidationData, ...]
│   ├── conditional_formats: tuple[ConditionalFormatData, ...]
│   ├── auto_filter / print_area / sheet_protection
│   └── cells: dict[str, CellData]
│       ├── coordinate: str ("B2")
│       ├── kind: CellKind (FORMULA | VALUE | EMPTY | ARRAY_FORMULA)
│       ├── formula: str | None ("=SUM(A1:A10)")
│       ├── value: object
│       ├── data_type: str
│       ├── style: CellStyleData
│       │   ├── number_format, font_name, font_size
│       │   ├── font_bold, font_italic, font_colour
│       │   ├── fill_colour, fill_pattern
│       │   ├── border_left/right/top/bottom
│       │   └── alignment, wrap_text
│       ├── comment: str | None
│       └── hyperlink: str | None
└── defined_names: tuple[DefinedNameData, ...]
    ├── name: str
    ├── value: str
    └── scope: str | None
```

### Formula Detection Logic

```python
if isinstance(value, ArrayFormula):
    kind = CellKind.ARRAY_FORMULA
    formula = value.text
elif isinstance(value, str) and len(value) > 1 and value.startswith("="):
    kind = CellKind.FORMULA
    formula = value
elif value is None and (has_style or has_comment):
    kind = CellKind.EMPTY
else:  # value is None with nothing → skip cell entirely
    kind = CellKind.VALUE
```

## 6. Comparison Data Models

```
WorkbookDiff
├── left_path / right_path: str
├── sheets_only_in_left: tuple[str, ...]     ← removed sheets
├── sheets_only_in_right: tuple[str, ...]    ← added sheets
├── sheets_in_both: tuple[str, ...]          ← shared (compared)
├── defined_name_diffs: tuple[PropertyDiff, ...]
├── sheet_diffs: tuple[SheetDiff, ...]
│   ├── sheet_name: str
│   ├── change_type: ChangeType (ADDED | REMOVED | MODIFIED | UNCHANGED)
│   ├── structural_diffs: tuple[PropertyDiff, ...]
│   │   ├── property_name: str ("merged_cells", "freeze_panes", ...)
│   │   ├── left_value / right_value: Any
│   │   └── change_type: ChangeType
│   └── cell_diffs: tuple[CellDiff, ...]
│       ├── coordinate: str
│       ├── change_type: ChangeType
│       ├── formula_changed: bool
│       ├── style_changed: bool
│       └── properties: tuple[PropertyDiff, ...]
│
├── has_differences: bool              (computed property)
├── total_formula_diffs: int           (computed property)
├── total_cell_diffs: int              (computed property)
└── total_style_diffs: int             (computed property)
```

## 7. Three-Level Comparison Cascade

### Level 1: Workbook

| Check | Method |
|-------|--------|
| Sheet names | Set operations: `left_only`, `right_only`, `in_both` |
| Sheet order | Compare `sheet_names` tuples |
| Defined names | Compare name/value/scope for each named range |

### Level 2: Sheet (structural)

| Property | Comparison |
|----------|------------|
| Dimensions (used range) | String equality |
| Merged cell regions | Set operations on sorted tuples |
| Freeze panes | String equality |
| Column widths | Dict key union, value comparison |
| Row heights | Dict key union, value comparison |
| Data validations | Rule-by-rule comparison |
| Conditional formatting | Rule-by-rule comparison |
| Auto-filter, print area | String equality |
| Sheet protection | Boolean equality |

### Level 3: Cell (the core)

| Property | Comparison | Notes |
|----------|------------|-------|
| Formula | String equality after `.upper()` | Case-insensitive: `=sum()` == `=SUM()` |
| Cell kind | Enum equality | Detects formula ↔ value transitions |
| Data type | String equality | e.g. `"s"` (string) vs `"n"` (number) |
| Style | Per-field dataclass comparison | 15 individual style properties |
| Comment | String equality | Cell notes/comments |
| Hyperlink | String equality | URL links |

## 8. Renderer Architecture

### Terminal Renderer (Rich)

```
ExcelDiff Report
├── Summary Panel (blue border)
│   ├── Left / Right file names
│   ├── Sheets compared / added / removed
│   └── Formula diffs / Style diffs / Total cell diffs
├── Defined Name Changes Table (if any)
├── Sheet Summary Table
│   └── Sheet Name | Status | Formula | Style | Structural
└── Per-Sheet Detail (for modified sheets only)
    ├── Structural Changes Table (cyan border)
    ├── Formula Changes Table (yellow border)
    ├── Style Changes Table (verbose mode only, dim border)
    └── Other Cell Changes Table (blue border)
```

### HTML Renderer (Jinja2)

- Self-contained single-file output (CSS inlined via `<style>` tags)
- Summary panel with stat cards
- Per-sheet sections with pill navigation
- Formula changes shown in monospace font
- Colour scheme: green (added), red (removed), yellow (modified)

### Excel Renderer (openpyxl)

- Copies structure from left (baseline) workbook
- Inserts `_ExcelDiff_Summary` sheet at position 0
- Annotates changed cells with:
  - **Comments**: "ExcelDiff: formula changed from X to Y"
  - **Fill colours**: Yellow (formula), Orange (style), Green (added), Red (removed), Blue (structural)
- Added sheets shown as placeholder tabs: `(+) SheetName`

### Colour Scheme

```
┌──────────────────────────────────────────────────────────┐
│ Green (#27AE60)  - Added cells/sheets                     │
├──────────────────────────────────────────────────────────┤
│ Red (#C0392B)    - Removed cells/sheets                   │
├──────────────────────────────────────────────────────────┤
│ Yellow (#F39C12) - Formula changes                        │
├──────────────────────────────────────────────────────────┤
│ Orange (#E67E22) - Style/formatting changes               │
├──────────────────────────────────────────────────────────┤
│ Blue (#3498DB)   - Structural changes                     │
├──────────────────────────────────────────────────────────┤
│ Dim/Grey         - Unchanged (terminal only)              │
└──────────────────────────────────────────────────────────┘
```

## 9. Web UI Architecture

```
Flask Application (create_app factory)
├── Configuration
│   ├── MAX_CONTENT_LENGTH: 16 MB
│   ├── UPLOAD_FOLDER: tempdir
│   └── secret_key: local-dev
│
├── Blueprint: "main" (routes.py)
│   ├── GET  /                           → Upload form
│   ├── POST /compare                    → File upload + comparison
│   ├── GET  /results/<job_id>           → Results summary page
│   ├── GET  /results/<job_id>/sheet/<name> → Sheet detail drill-down
│   └── GET  /results/<job_id>/download/<fmt> → Download HTML/Excel
│
├── In-Memory Storage
│   ├── _results: dict[job_id → WorkbookDiff]
│   └── _file_paths: dict[job_id → (left_path, right_path)]
│
└── Templates (Jinja2)
    ├── base.html         → Navbar, flash messages, CSS framework
    ├── upload.html       → Two file inputs, option checkboxes
    ├── results.html      → Stat cards, sheet table, download buttons
    └── sheet_detail.html → Structural/formula/style/other tables
```

## 10. CLI Interface

```
exceldiff
├── diff LEFT RIGHT                      # Compare two workbooks
│   ├── --format / -f [terminal|html|excel]
│   ├── --output / -o PATH
│   ├── --formulas-only                  # Skip value/style comparison
│   ├── --ignore-style                   # Skip style comparison
│   ├── --sheets / -s NAME              # Filter to specific sheets
│   └── --verbose / -v                   # Show style diffs in terminal
│
├── inspect WORKBOOK                     # View single workbook structure
│   └── Shows: sheets, dimensions, cell counts, formula counts, merged regions
│
└── web                                  # Launch Flask web UI
    ├── --host (default: 127.0.0.1)
    ├── --port / -p (default: 5000)
    └── --debug
```

Exit codes: `0` = no differences, `1` = differences found.

## 11. Dependencies

| Package | Purpose | Version |
|---------|---------|---------|
| openpyxl | Parse and write .xlsx files | >= 3.1.0 |
| click | CLI framework | >= 8.0 |
| rich | Terminal colour output | >= 13.0 |
| jinja2 | HTML report templates | >= 3.1 (optional) |
| flask | Web UI server | >= 3.0 (optional) |
| pytest | Test framework | >= 7.0 (dev) |
| pytest-cov | Coverage reporting | >= 4.0 (dev) |
| ruff | Linter / formatter | >= 0.1.0 (dev) |

## 12. Testing Strategy

- **48 tests** across 4 test suites
- **Fixtures**: All test .xlsx files created programmatically in `conftest.py` (no binary files in git)
- **6 fixture workbooks**: base, modified_formulas, modified_structure, modified_formatting, identical, empty
- **Unit tests**: Each parser and comparator tested independently
- **Integration tests**: Full `compare()` pipeline asserting specific diffs
- **Renderer tests**: HTML string parsing, Excel re-read verification
- **Web tests**: Flask test client, file upload simulation, route assertions

| Suite | Tests | Coverage |
|-------|-------|----------|
| test_parsing | 11 | Workbook/sheet/cell parsers |
| test_comparison | 19 | Engine, all comparator levels |
| test_renderers | 11 | HTML + Excel output |
| test_web | 7 | Flask routes + downloads |

## 13. Key Technical Decisions

| Decision | Rationale |
|----------|-----------|
| `data_only=False` | Returns formula strings instead of cached computed values |
| `read_only=False` | Required for merged cells, styles, data validations |
| Frozen dataclasses | Ensures immutability; safe for comparison, hashing, and caching |
| Case-insensitive formulas | `=sum()` and `=SUM()` are functionally identical |
| Decouple from openpyxl | Parsed models have no openpyxl dependency; enables future format support |
| In-memory web results | Suitable for local/dev use; swap for Redis/DB in production |
| `(+)` sheet prefix | Excel forbids `[` `]` in sheet names; parentheses are safe |
