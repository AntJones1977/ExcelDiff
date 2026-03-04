# ExcelDiff - Process Maps

Visual process maps documenting how the application works, how users interact with each interface, and how data flows through the comparison pipeline. All diagrams use [Mermaid](https://mermaid.js.org/) syntax and render natively in GitHub and VSCode.

---

## 1. High-Level System Flow

The core pipeline from input files to output report, showing all entry points and output formats.

```mermaid
graph TB
    subgraph INPUTS["Input Sources"]
        direction LR
        CLI["CLI<br/>exceldiff diff"]
        WEB["Web UI<br/>Browser Upload"]
        API["Python API<br/>compare() function"]
    end

    subgraph FILES["Excel Files"]
        direction LR
        LEFT["LEFT.xlsx<br/>(Baseline)"]
        RIGHT["RIGHT.xlsx<br/>(Comparison)"]
    end

    INPUTS --> FILES

    FILES --> PARSE

    subgraph PARSE["Parsing Layer"]
        direction TB
        WBP["workbook_parser<br/>━━━━━━━━━━━━━━━<br/>openpyxl.load_workbook<br/>(data_only=False)"]
        WBP --> SP["sheet_parser<br/>━━━━━━━━━━━━━━━<br/>Merged cells, freeze panes<br/>Column widths, row heights<br/>Validations, conditions"]
        SP --> CP["cell_parser<br/>━━━━━━━━━━━━━━━<br/>Formula detection<br/>Style normalisation<br/>Comment extraction"]
    end

    PARSE --> SNAPSHOT["WorkbookData<br/>(Frozen Snapshots)<br/>━━━━━━━━━━━━━━━<br/>Left + Right"]

    SNAPSHOT --> ENGINE

    subgraph ENGINE["Comparison Engine"]
        direction TB
        WBC["workbook_comparator<br/>Sheet set operations<br/>Defined name diffs"]
        WBC --> SC["sheet_comparator<br/>Structural property diffs"]
        SC --> CC["cell_comparator<br/>Formula + style comparison<br/>(case-insensitive)"]
    end

    ENGINE --> RESULT["WorkbookDiff<br/>(Frozen Result Tree)<br/>━━━━━━━━━━━━━━━<br/>has_differences: bool<br/>total_formula_diffs: int"]

    RESULT --> RENDER

    subgraph RENDER["Output Renderers"]
        direction LR
        TERM["Terminal<br/>(Rich)"]
        HTML["HTML Report<br/>(Jinja2)"]
        EXCEL["Annotated Excel<br/>(openpyxl)"]
        WEBUI["Web Results<br/>(Flask Templates)"]
    end

    style INPUTS fill:#FFFFCC,stroke:#2C3E50
    style FILES fill:#D4E6F1,stroke:#2C3E50
    style PARSE fill:#f0f4f8,stroke:#34495E,stroke-width:2px
    style ENGINE fill:#f0f4f8,stroke:#E67E22,stroke-width:2px
    style RENDER fill:#d4edda,stroke:#27AE60,stroke-width:2px
    style SNAPSHOT fill:#D4E6F1,stroke:#2C3E50
    style RESULT fill:#27AE60,color:#fff
```

---

## 2. CLI User Workflow

The complete journey of a user running ExcelDiff from the command line.

```mermaid
flowchart TD
    START(["User Runs<br/>exceldiff command"]) --> CMD{"Which<br/>subcommand?"}

    CMD -->|"diff"| DIFF
    CMD -->|"inspect"| INSPECT
    CMD -->|"web"| WEB

    subgraph DIFF["exceldiff diff LEFT RIGHT"]
        direction TB
        ARGS["Parse Arguments<br/>━━━━━━━━━━━━━━━<br/>LEFT.xlsx, RIGHT.xlsx<br/>--format (terminal/html/excel)<br/>--output path<br/>--formulas-only<br/>--ignore-style<br/>--sheets filter<br/>--verbose"]

        ARGS --> VALIDATE{"Files<br/>exist?"}
        VALIDATE -->|No| ERR1["Error: Path does not exist<br/>Exit code 2"]
        VALIDATE -->|Yes| COMPARE["compare()<br/>Parse + Diff Engine"]

        COMPARE --> FORMAT{"Output<br/>format?"}

        FORMAT -->|terminal| RICH["Terminal Renderer<br/>━━━━━━━━━━━━━━━<br/>Summary panel<br/>Sheet summary table<br/>Formula changes<br/>Style changes (if -v)"]

        FORMAT -->|html| HTMLR["HTML Renderer<br/>━━━━━━━━━━━━━━━<br/>Self-contained .html<br/>Inline CSS<br/>'HTML report saved to: path'"]

        FORMAT -->|excel| EXCELR["Excel Renderer<br/>━━━━━━━━━━━━━━━<br/>Annotated .xlsx<br/>Summary sheet + comments<br/>'Annotated Excel saved to: path'"]

        RICH --> EXIT
        HTMLR --> EXIT
        EXCELR --> EXIT

        EXIT{"Differences<br/>found?"}
        EXIT -->|No| CODE0["Exit Code 0<br/>(CI: pass)"]
        EXIT -->|Yes| CODE1["Exit Code 1<br/>(CI: fail)"]
    end

    subgraph INSPECT["exceldiff inspect WORKBOOK"]
        direction TB
        IPARSE["Parse Single Workbook<br/>━━━━━━━━━━━━━━━<br/>parse_workbook(path)"]
        IPARSE --> IPANEL["Rich Panel<br/>File path<br/>Sheet count<br/>Defined names count"]
        IPANEL --> ITABLE["Rich Table<br/>━━━━━━━━━━━━━━━<br/>Sheet Name | Dimensions<br/>Cells | Formulas<br/>Merged Regions | Freeze"]
    end

    subgraph WEB["exceldiff web"]
        direction TB
        FLASK["create_app()<br/>Flask factory"]
        FLASK --> SERVE["app.run()<br/>━━━━━━━━━━━━━━━<br/>--host (127.0.0.1)<br/>--port (5000)<br/>--debug"]
        SERVE --> BROWSER["Open browser at<br/>http://127.0.0.1:5000"]
    end

    style START fill:#27AE60,color:#fff,stroke:#27AE60
    style DIFF fill:#fff,stroke:#2C3E50
    style INSPECT fill:#fff,stroke:#3498DB
    style WEB fill:#fff,stroke:#E67E22
    style ERR1 fill:#C0392B,color:#fff
    style CODE0 fill:#27AE60,color:#fff
    style CODE1 fill:#F39C12,color:#fff
```

---

## 3. Web UI User Interaction Flow

The end-to-end journey of using the Flask web interface to compare workbooks.

```mermaid
flowchart TD
    START(["User Opens Browser<br/>http://127.0.0.1:5000"]) --> UPLOAD

    subgraph UPLOAD["Upload Page (GET /)"]
        direction TB
        FORM["Upload Form<br/>━━━━━━━━━━━━━━━━━━━<br/>📁 Left File (Baseline)<br/>📁 Right File (Comparison)<br/>☐ Ignore style differences<br/>☐ Formulas only<br/>━━━━━━━━━━━━━━━━━━━<br/>🔄 Compare Button"]
    end

    FORM --> SUBMIT{{"User Clicks<br/>Compare"}}

    SUBMIT --> VALIDATE

    subgraph VALIDATE["Server Validation (POST /compare)"]
        direction TB
        V1{"Both files<br/>uploaded?"}
        V1 -->|No| FLASH1["Flash: 'Please select<br/>a left/right file'"]
        FLASH1 --> REDIR1["Redirect → /"]

        V1 -->|Yes| V2{"Both .xlsx<br/>extension?"}
        V2 -->|No| FLASH2["Flash: 'Only .xlsx<br/>files are supported'"]
        FLASH2 --> REDIR2["Redirect → /"]

        V2 -->|Yes| SAVE["Save to temp directory<br/>━━━━━━━━━━━━━━━━━━━<br/>upload_dir/job_id/<br/>├── left.xlsx<br/>└── right.xlsx"]
    end

    SAVE --> COMPARE

    subgraph COMPARE["Comparison Engine"]
        direction TB
        RUN["compare(left_path, right_path)<br/>━━━━━━━━━━━━━━━━━━━<br/>ignore_style: from checkbox<br/>formulas_only: from checkbox"]

        RUN --> STORE["Store in memory<br/>━━━━━━━━━━━━━━━━━━━<br/>_results[job_id] = WorkbookDiff<br/>_file_paths[job_id] = (left, right)"]
    end

    STORE --> REDIRECT["Redirect → /results/job_id"]

    REDIRECT --> RESULTS

    subgraph RESULTS["Results Page (GET /results/job_id)"]
        direction TB
        STATS["Summary Stats Cards<br/>━━━━━━━━━━━━━━━━━━━<br/>📊 Total Cell Diffs<br/>📐 Formula Diffs<br/>🎨 Style Diffs<br/>📋 Sheets Compared"]

        STATS --> SHEETTABLE["Sheet Table<br/>━━━━━━━━━━━━━━━━━━━<br/>Name | Status | Formulas | Styles<br/>Each row → View Details link"]

        SHEETTABLE --> DOWNLOADS["Download Buttons<br/>━━━━━━━━━━━━━━━━━━━<br/>📄 Download HTML Report<br/>📊 Download Annotated Excel"]
    end

    SHEETTABLE -->|"Click<br/>View Details"| DETAIL

    subgraph DETAIL["Sheet Detail Page (GET /results/job_id/sheet/name)"]
        direction TB
        D1["Structural Changes Table<br/>Property | Left | Right | Change"]
        D2["Formula Changes Table<br/>Cell | Left Formula | Right Formula"]
        D3["Style Changes Table<br/>Cell | Property | Left | Right"]
        D4["Other Cell Changes Table<br/>Cell | Property | Left | Right"]

        D1 --> D2 --> D3 --> D4
    end

    DOWNLOADS -->|"Click Download<br/>HTML"| DLHTML["GET /results/job_id/download/html<br/>━━━━━━━━━━━━━━━━━━━<br/>Generate HTML via renderer<br/>Send as file attachment"]

    DOWNLOADS -->|"Click Download<br/>Excel"| DLEXCEL["GET /results/job_id/download/excel<br/>━━━━━━━━━━━━━━━━━━━<br/>Generate .xlsx via renderer<br/>Send as file attachment"]

    style START fill:#27AE60,color:#fff,stroke:#27AE60
    style SUBMIT fill:#2C3E50,color:#fff
    style UPLOAD fill:#FFFFCC,stroke:#2C3E50,stroke-width:2px
    style VALIDATE fill:#f0f4f8,stroke:#34495E
    style COMPARE fill:#f0f4f8,stroke:#E67E22,stroke-width:2px
    style RESULTS fill:#D4E6F1,stroke:#2C3E50,stroke-width:2px
    style DETAIL fill:#fff,stroke:#3498DB
    style DLHTML fill:#d4edda,stroke:#27AE60
    style DLEXCEL fill:#d4edda,stroke:#27AE60
    style FLASH1 fill:#C0392B,color:#fff
    style FLASH2 fill:#C0392B,color:#fff
```

---

## 4. Parsing Pipeline

How a raw .xlsx file is transformed into an immutable `WorkbookData` snapshot, cell by cell.

```mermaid
flowchart LR
    subgraph INPUT["Raw .xlsx File"]
        direction TB
        XLSX["file.xlsx<br/>(ZIP archive containing<br/>XML worksheets)"]
    end

    INPUT --> LOAD["openpyxl.load_workbook<br/>━━━━━━━━━━━━━━━━━━━<br/>data_only=False<br/>  → raw formulas<br/>read_only=False<br/>  → full structure access"]

    LOAD --> WBP

    subgraph WBP["workbook_parser.parse_workbook()"]
        direction TB

        SHEETS["Iterate wb.sheetnames<br/>For each worksheet:"]

        SHEETS --> SP

        subgraph SP["sheet_parser.parse_sheet(ws)"]
            direction TB
            STRUCT["Extract Structure<br/>━━━━━━━━━━━━━━━━━━━<br/>• dimensions (used range)<br/>• merged_cells (sorted tuple)<br/>• freeze_panes<br/>• column_widths (non-default)<br/>• row_heights (non-default)<br/>• data_validations<br/>• conditional_formats<br/>• auto_filter, print_area<br/>• sheet_protection"]

            STRUCT --> CELLS["Iterate ws.iter_rows()<br/>For each cell:"]

            CELLS --> CP

            subgraph CP["cell_parser.parse_cell(cell)"]
                direction TB
                DETECT{"Cell value<br/>type?"}

                DETECT -->|"ArrayFormula<br/>object"| AF["CellKind.ARRAY_FORMULA<br/>formula = value.text"]
                DETECT -->|"String starting<br/>with '='"| F["CellKind.FORMULA<br/>formula = value"]
                DETECT -->|"None + has<br/>style/comment"| E["CellKind.EMPTY"]
                DETECT -->|"None + nothing<br/>else"| SKIP["Skip cell entirely<br/>(not stored)"]
                DETECT -->|"Any other<br/>value"| V["CellKind.VALUE"]

                AF --> STYLE
                F --> STYLE
                E --> STYLE
                V --> STYLE

                STYLE["Extract CellStyleData<br/>━━━━━━━━━━━━━━━━━━━<br/>number_format<br/>font (name, size, bold, italic, colour)<br/>fill (colour, pattern)<br/>borders (left, right, top, bottom)<br/>alignment (h, v, wrap)"]

                STYLE --> COMMENT["Extract comment<br/>+ hyperlink"]
            end
        end

        SP --> DN["Extract defined_names<br/>━━━━━━━━━━━━━━━━━━━<br/>for name in wb.defined_names:<br/>  dn = wb.defined_names[name]<br/>  → DefinedNameData"]
    end

    WBP --> OUTPUT["WorkbookData<br/>(frozen dataclass)<br/>━━━━━━━━━━━━━━━━━━━<br/>Immutable snapshot<br/>Decoupled from openpyxl<br/>Safe for comparison"]

    style INPUT fill:#D4E6F1,stroke:#2C3E50
    style WBP fill:#f0f4f8,stroke:#34495E,stroke-width:2px
    style SP fill:#fff,stroke:#3498DB
    style CP fill:#FFFFCC,stroke:#2C3E50
    style OUTPUT fill:#27AE60,color:#fff
    style SKIP fill:#BDC3C7,stroke:#7F8C8D
```

---

## 5. Comparison Engine Pipeline

How two `WorkbookData` snapshots are compared to produce a `WorkbookDiff` result tree.

```mermaid
flowchart TD
    subgraph INPUTS["Parsed Snapshots"]
        direction LR
        LEFT["WorkbookData<br/>(Left / Baseline)"]
        RIGHT["WorkbookData<br/>(Right / Comparison)"]
    end

    INPUTS --> OPTIONS

    subgraph OPTIONS["Comparison Options"]
        direction LR
        OPT1["ignore_style: bool<br/>Skip all style comparison"]
        OPT2["formulas_only: bool<br/>Only compare formulas"]
        OPT3["sheet_filter: tuple<br/>Compare specific sheets only"]
    end

    OPTIONS --> WBC

    subgraph WBC["workbook_comparator.compare_workbooks()"]
        direction TB

        SETOPS["Sheet Set Operations<br/>━━━━━━━━━━━━━━━━━━━<br/>left_names = set(left.sheet_names)<br/>right_names = set(right.sheet_names)<br/>━━━━━━━━━━━━━━━━━━━<br/>only_in_left = left - right<br/>only_in_right = right - left<br/>in_both = left & right"]

        SETOPS --> FILTER{"sheet_filter<br/>specified?"}
        FILTER -->|Yes| APPLY["in_both &= set(sheet_filter)"]
        FILTER -->|No| COMPARE

        APPLY --> COMPARE

        COMPARE["For each sheet in_both:"]

        COMPARE --> SC

        subgraph SC["sheet_comparator.compare_sheets()"]
            direction TB

            STRUCTURAL["Compare Structural Properties<br/>━━━━━━━━━━━━━━━━━━━━━━━━━<br/>dimensions, merged_cells (set ops),<br/>freeze_panes, column_widths (dict merge),<br/>row_heights, auto_filter, print_area,<br/>sheet_protection, data_validations,<br/>conditional_formats"]

            STRUCTURAL --> CC

            subgraph CC["cell_comparator.compare_cells()"]
                direction TB

                UNION["Union all coordinates<br/>━━━━━━━━━━━━━━━━━━━<br/>all_coords = left_cells.keys()<br/>             | right_cells.keys()"]

                UNION --> EACH["For each coordinate:"]

                EACH --> CLASS{"Cell in<br/>which set?"}

                CLASS -->|"Left only"| REM["CellDiff<br/>change_type = REMOVED"]
                CLASS -->|"Right only"| ADD["CellDiff<br/>change_type = ADDED"]
                CLASS -->|"Both"| CMP["Compare cell pair"]

                CMP --> FORMULA["Formula Comparison<br/>━━━━━━━━━━━━━━━━━━━<br/>left.formula.upper()<br/>  vs<br/>right.formula.upper()<br/>(case-insensitive)"]

                CMP --> KIND["Kind Comparison<br/>━━━━━━━━━━━━━━━━━━━<br/>FORMULA ↔ VALUE<br/>VALUE ↔ EMPTY<br/>etc."]

                CMP --> STYLECMP["Style Comparison<br/>(if not ignore_style)<br/>━━━━━━━━━━━━━━━━━━━<br/>Compare each of 15<br/>CellStyleData fields"]

                FORMULA --> CELLDIFF
                KIND --> CELLDIFF
                STYLECMP --> CELLDIFF

                CELLDIFF["CellDiff<br/>━━━━━━━━━━━━━━━<br/>formula_changed: bool<br/>style_changed: bool<br/>properties: [PropertyDiff]"]
            end

            CC --> SHEETDIFF["SheetDiff<br/>━━━━━━━━━━━━━━━<br/>structural_diffs<br/>cell_diffs<br/>change_type"]
        end

        SC --> DNCOMP["Compare Defined Names<br/>━━━━━━━━━━━━━━━━━━━<br/>Set operations on name keys<br/>Value comparison for shared names"]
    end

    WBC --> RESULT["WorkbookDiff<br/>━━━━━━━━━━━━━━━━━━━<br/>sheet_diffs<br/>defined_name_diffs<br/>sheets_only_in_left/right<br/>━━━━━━━━━━━━━━━━━━━<br/>has_differences: bool<br/>total_formula_diffs: int<br/>total_cell_diffs: int<br/>total_style_diffs: int"]

    style INPUTS fill:#D4E6F1,stroke:#2C3E50
    style OPTIONS fill:#FFFFCC,stroke:#2C3E50
    style WBC fill:#f0f4f8,stroke:#E67E22,stroke-width:2px
    style SC fill:#fff,stroke:#3498DB
    style CC fill:#fff,stroke:#F39C12
    style RESULT fill:#27AE60,color:#fff
    style REM fill:#C0392B,color:#fff
    style ADD fill:#27AE60,color:#fff
```

---

## 6. Output Rendering Pipeline

How a `WorkbookDiff` result is transformed into each output format.

```mermaid
flowchart TD
    DIFF["WorkbookDiff<br/>(Frozen Result Tree)"] --> FORMAT{"Output<br/>format?"}

    FORMAT -->|Terminal| TERM
    FORMAT -->|HTML| HTML
    FORMAT -->|Excel| EXCEL
    FORMAT -->|Web| WEBR

    subgraph TERM["Terminal Renderer (Rich)"]
        direction TB
        T1["Summary Panel<br/>━━━━━━━━━━━━━━━<br/>Blue border<br/>Left/Right filenames<br/>Stats: compared, added, removed"]
        T2["Defined Name Changes<br/>(if any)"]
        T3["Sheet Summary Table<br/>━━━━━━━━━━━━━━━<br/>Name | Status | Formulas<br/>Style | Structural"]
        T4["Per-Sheet Details<br/>━━━━━━━━━━━━━━━<br/>Structural Changes (cyan)<br/>Formula Changes (yellow)<br/>Style Changes (dim, -v only)<br/>Other Cell Changes (blue)"]

        T1 --> T2 --> T3 --> T4
        T4 --> STDOUT["stdout<br/>(colour-coded)"]
    end

    subgraph HTML["HTML Renderer (Jinja2)"]
        direction TB
        H1["Load report.html template"]
        H2["Load style.css → inline"]
        H3["Render with WorkbookDiff context<br/>━━━━━━━━━━━━━━━━━━━━━━━━━<br/>Summary panel + stat cards<br/>Sheet navigation pills<br/>Defined name changes<br/>Per-sheet sections:<br/>  structural / formula / style"]
        H4["Write self-contained .html"]

        H1 --> H2 --> H3 --> H4
        H4 --> HTMLFILE["report.html<br/>(single file,<br/>no external deps)"]
    end

    subgraph EXCEL["Excel Renderer (openpyxl)"]
        direction TB
        E1["Load left (baseline) .xlsx<br/>as starting template"]
        E2["Insert _ExcelDiff_Summary<br/>at sheet position 0<br/>━━━━━━━━━━━━━━━━━━━━━━━━━<br/>ExcelDiff Report header<br/>Left/Right file info<br/>Per-sheet stats table<br/>Colour legend"]
        E3["Annotate changed cells<br/>━━━━━━━━━━━━━━━━━━━━━━━━━<br/>Comments: 'Formula changed<br/>  from X to Y'<br/>Fill colours:<br/>  Yellow = formula<br/>  Orange = style<br/>  Green = added<br/>  Red = removed<br/>  Blue = structural"]
        E4["Add placeholder sheets<br/>for added sheets: '(+) Name'"]

        E1 --> E2 --> E3 --> E4
        E4 --> XLSXFILE["annotated.xlsx"]
    end

    subgraph WEBR["Web Renderer (Flask)"]
        direction TB
        W1["Store WorkbookDiff<br/>in _results dict"]
        W2["Render results.html template<br/>━━━━━━━━━━━━━━━━━━━━━━━━━<br/>Stat cards grid<br/>Sheet table with detail links<br/>Download buttons"]
        W3["On download request:<br/>delegate to HTML or<br/>Excel renderer"]

        W1 --> W2 --> W3
        W3 --> BROWSER["Browser renders<br/>results page"]
    end

    style DIFF fill:#27AE60,color:#fff
    style TERM fill:#fff,stroke:#2C3E50
    style HTML fill:#fff,stroke:#C0392B
    style EXCEL fill:#fff,stroke:#3498DB
    style WEBR fill:#fff,stroke:#E67E22
    style STDOUT fill:#f0f4f8,stroke:#34495E
    style HTMLFILE fill:#d4edda,stroke:#27AE60
    style XLSXFILE fill:#D4E6F1,stroke:#2C3E50
    style BROWSER fill:#FFFFCC,stroke:#2C3E50
```

---

## 7. Test Fixture Architecture

How the 6 test workbooks are created programmatically and used across test suites.

```mermaid
flowchart TD
    subgraph CONFTEST["conftest.py — Fixture Builder"]
        direction TB
        BUILDER["openpyxl.Workbook()<br/>Create .xlsx programmatically<br/>(no binary files in git)"]

        BUILDER --> BASE
        BUILDER --> MODF
        BUILDER --> MODS
        BUILDER --> MODFMT
        BUILDER --> IDENT
        BUILDER --> EMPTY

        subgraph BASE["base.xlsx (3 sheets)"]
            direction TB
            B1["Summary<br/>• A1: 'Revenue' (value)<br/>• B1: =SUM(Data!B2:B5)<br/>• B2: =AVERAGE(Data!B2:B5)<br/>• Merge: A1:A2<br/>• Freeze: A3"]
            B2["Data<br/>• 9 rows of data<br/>• Column B: values (100-900)<br/>• Column C: =B*1.2 formulas<br/>• Column D: markup formulas"]
            B3["Config<br/>• Single-cell settings<br/>• A1: 'VAT Rate'<br/>• B1: 0.20"]
        end

        subgraph MODF["modified_formulas.xlsx"]
            direction TB
            MF1["Summary<br/>• B1: =SUMIF(Data!A2:A5,'>0',...)<br/>  (was =SUM)"]
            MF2["Data<br/>• C2: =B2*1.25 (was *1.2)<br/>• D2: =LEN(A2) (new column)"]
        end

        subgraph MODS["modified_structure.xlsx"]
            direction TB
            MS1["Summary<br/>• Merge removed<br/>• Freeze changed A2→A3"]
            MS2["Settings (renamed from Config)"]
            MS3["Audit (new sheet)"]
            MS4["Column width changed"]
        end

        subgraph MODFMT["modified_formatting.xlsx"]
            direction TB
            MFT1["Data headers: bold + size 14"]
            MFT2["B2: yellow fill added"]
        end

        IDENT["identical.xlsx<br/>(exact copy of base)"]
        EMPTY["empty.xlsx<br/>(single empty sheet)"]
    end

    BASE --> TESTS
    MODF --> TESTS
    MODS --> TESTS
    MODFMT --> TESTS
    IDENT --> TESTS
    EMPTY --> TESTS

    subgraph TESTS["Test Suites (48 tests)"]
        direction LR
        TP["test_parsing/<br/>11 tests"]
        TC["test_comparison/<br/>19 tests"]
        TR["test_renderers/<br/>11 tests"]
        TW["test_web/<br/>7 tests"]
    end

    style CONFTEST fill:#f0f4f8,stroke:#34495E,stroke-width:2px
    style BASE fill:#D4E6F1,stroke:#2C3E50
    style MODF fill:#FFFFCC,stroke:#F39C12
    style MODS fill:#FFFFCC,stroke:#E67E22
    style MODFMT fill:#FFFFCC,stroke:#3498DB
    style IDENT fill:#d4edda,stroke:#27AE60
    style EMPTY fill:#f4f6f7,stroke:#BDC3C7
    style TESTS fill:#fff,stroke:#2C3E50,stroke-width:2px
```

---

## 8. CI/CD & Integration Pipeline

How ExcelDiff integrates into automated workflows and CI pipelines.

```mermaid
flowchart LR
    subgraph TRIGGERS["CI Triggers"]
        direction TB
        PUSH["git push"]
        PR["Pull Request"]
        MANUAL["Manual run"]
    end

    TRIGGERS --> CI

    subgraph CI["CI Pipeline"]
        direction TB
        INSTALL["pip install -e '.[dev,html,web]'"]
        INSTALL --> LINT["ruff check src/ tests/"]
        LINT --> TEST["pytest --cov=exceldiff"]
        TEST --> REPORT["Coverage Report"]
    end

    subgraph USAGE["Integration Usage"]
        direction TB

        subgraph CICHECK["CI Spreadsheet Audit"]
            direction TB
            CIRUN["exceldiff diff<br/>baseline.xlsx updated.xlsx"]
            CIRUN --> EXITCODE{"Exit code?"}
            EXITCODE -->|"0"| PASS["Pass<br/>(no changes)"]
            EXITCODE -->|"1"| FAIL["Fail<br/>(formulas changed!)"]
        end

        subgraph PRECOMMIT["Pre-Commit Hook"]
            direction TB
            HOOK["On .xlsx file change:<br/>exceldiff diff HEAD:file.xlsx file.xlsx"]
            HOOK --> BLOCK{"Unexpected<br/>formula changes?"}
            BLOCK -->|Yes| REJECT["Block commit"]
            BLOCK -->|No| ALLOW["Allow commit"]
        end

        subgraph REPORT_GEN["Batch Report Generation"]
            direction TB
            BATCH["exceldiff diff old.xlsx new.xlsx<br/>-f html -o report.html"]
            BATCH --> PUBLISH["Publish report<br/>to shared drive / wiki"]
        end
    end

    CI --> USAGE

    style TRIGGERS fill:#FFFFCC,stroke:#2C3E50
    style CI fill:#f0f4f8,stroke:#34495E,stroke-width:2px
    style CICHECK fill:#D4E6F1,stroke:#2C3E50
    style PRECOMMIT fill:#fff,stroke:#E67E22
    style REPORT_GEN fill:#d4edda,stroke:#27AE60
    style PASS fill:#27AE60,color:#fff
    style FAIL fill:#C0392B,color:#fff
    style REJECT fill:#C0392B,color:#fff
    style ALLOW fill:#27AE60,color:#fff
```

---

## Diagram Legend

| Symbol | Meaning |
|--------|---------|
| Rounded rectangle | Start/end point |
| Rectangle | Screen, process step, or component |
| Diamond | Decision point |
| Hexagon | User action |
| Cylinder | Database / persistent storage |
| Dashed arrow | Optional or conditional path |
| Solid arrow | Primary flow direction |

### Colour Coding

| Colour | Usage in Diagrams |
|--------|-------------------|
| Yellow (#FFFFCC) | User inputs / options / configuration |
| Light Blue (#D4E6F1) | Data / file storage / snapshots |
| Green (#27AE60) | Start points / success states / final output |
| Dark Blue (#2C3E50) | User action buttons |
| Orange (#E67E22) | Comparison engine / processing |
| Red (#C0392B) | Error states / removed items |
| Blue (#3498DB) | Detail views / structural elements |
| White | Screen content areas |
| Light grey (#f0f4f8) | Service / processing layers |
