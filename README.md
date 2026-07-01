# openexcel

A high-performance Python library for reading and writing `.xlsx` files, implemented as a C extension. Designed as a drop-in accelerator for workloads where [openpyxl](https://openpyxl.readthedocs.io/) is too slow.

## Why

openpyxl is pure Python. For large files (100k+ rows), the bottlenecks are ZIP extraction, XML parsing, and per-cell object allocation — all done in the Python interpreter. openexcel moves these to C:

- **Streaming SAX parsing** via libexpat — never loads the full XML into memory
- **Flat sorted cell array** — O(1) append during parse, O(1) sequential iteration
- **GIL released** during read and write — other threads run freely
- **Vendored miniz** — no external ZIP dependency

Benchmarked against openpyxl on mixed workloads (integers, floats, booleans, dates, strings):

| Rows | Write speedup | Read speedup |
|---:|---:|---:|
| 1,000 | 8.5× | 5.3× |
| 10,000 | 8.8× | 5.2× |
| 50,000 | 8.1× | 5.3× |
| 100,000 | 8.0× | 5.3× |
| 250,000 | 7.8× | 5.4× |

**~8× faster writes, ~5× faster reads** across all sizes tested.

### Feature benchmarks

Benchmarked on macOS arm64, Python 3.11. Each result is the median of 3 runs.

| Benchmark | openexcel | openpyxl | Speedup |
|---|---:|---:|---:|
| Cell access (10k reads) | 1.69 ms | 1.82 ms | 1.1× |
| Number format set (1k cells) | 220 µs | 429 µs | 2.0× |
| Number format roundtrip (1k cells) | 2.04 ms | 18.04 ms | **8.8×** |
| Formula write (1k cells) | 390 µs | 942 µs | 2.4× |
| Formula roundtrip (1k cells) | 2.88 ms | 23.09 ms | **8.0×** |
| Merged cells roundtrip (100 ranges) | 366 µs | 12.90 ms | **35.2×** |
| Column/row dimensions roundtrip (50+50) | 329 µs | 4.54 ms | **13.8×** |
| Named ranges roundtrip (20) | 321 µs | 3.36 ms | **10.5×** |
| Freeze panes roundtrip | 525 µs | 5.56 ms | **10.6×** |
| Hyperlinks roundtrip (50 cells) | 662 µs | 5.26 ms | **7.9×** |
| Data validation roundtrip (10 rules) | 390 µs | 3.44 ms | **8.8×** |
| Page setup roundtrip | 355 µs | 3.11 ms | **8.8×** |
| Sheet protection roundtrip | 347 µs | 3.42 ms | **9.9×** |

The pattern: **in-memory cell manipulation is ~2× faster; any operation involving save+load is 8–35× faster**. The C XML serializer and SAX parser are where the largest gains appear.

## Installation

```bash
pip install openexcel-c
```

Pre-built wheels are available for macOS (arm64, x86_64) and Linux (x86_64, aarch64) for Python 3.10–3.13.

### Building from source

You need CMake ≥ 3.20, a C11 compiler, and libexpat development headers.

```bash
# macOS
brew install expat cmake

# Ubuntu/Debian
sudo apt-get install libexpat1-dev cmake

pip install openexcel-c --no-binary openexcel-c
```

## Usage

### Reading

```python
import openexcel

wb = openexcel.load_workbook("data.xlsx")
ws = wb.active   # first sheet

# Iterate rows — returns tuples of Python values
for row in ws.iter_rows():
    print(row)   # e.g. (1, "hello", 3.14, True, datetime.date(2024, 6, 1))

# Slice a range
for row in ws.iter_rows(min_row=2, max_row=100, min_col=1, max_col=5):
    print(row)
```

### Cell objects

```python
# Single cell by A1 notation
cell = ws["A1"]
print(cell.value, cell.row, cell.column, cell.coordinate)  # 42, 1, 1, "A1"

# Single cell by row/column
cell = ws.cell(row=1, column=1)

# Range — returns tuple of tuples of Cell objects
cells = ws["A1:C3"]

# Set a value
cell.value = "hello"
cell.value = 3.14
cell.value = datetime.date(2024, 6, 1)
```

Cell values map to Python types:

| Excel type | Python type |
|---|---|
| Number | `float` |
| String | `str` |
| Boolean | `bool` |
| Date / datetime | `datetime.date` / `datetime.datetime` |
| Formula | `str` (e.g. `"=SUM(A1:A10)"`) |
| Empty | `None` |
| Error | `str` (e.g. `"#DIV/0!"`) |

### Writing

```python
import openexcel
import datetime

wb = openexcel.Workbook()
ws = wb.create_sheet("Sheet1")

ws.append(["Name", "Score", "Date"])
ws.append(["Alice", 98.5, datetime.date(2024, 6, 1)])
ws.append(["Bob",   72.0, datetime.date(2024, 6, 2)])

wb.save("output.xlsx")
```

### Number formats

```python
cell = ws["A1"]
cell.value = 0.1234
cell.number_format = "0.00%"   # displays as 12.34%

cell2 = ws["B1"]
cell2.value = 1234567.89
cell2.number_format = "#,##0.00"

# Read back the format string after loading
wb2 = openexcel.load_workbook("output.xlsx")
print(wb2[0]["A1"].number_format)  # "0.00%"
```

All built-in Excel format IDs (0–49) are recognized by format string. Custom format strings are assigned IDs starting at 164 and are round-tripped correctly.

### Formulas

```python
ws["A1"].value = 10
ws["A2"].value = 20
ws["A3"].value = "=SUM(A1:A2)"   # leading "=" marks a formula

# After save/load:
cell = wb2[0]["A3"]
print(cell.value)      # "=SUM(A1:A2)"
print(cell.data_type)  # "f"
```

Formulas are stored and round-tripped as strings. openexcel does not evaluate formulas — cached values from the original file are not preserved (matching openpyxl's default `data_only=False` behavior).

### Sheet formatting

```python
# Column and row dimensions
ws.set_column_width("A", 20)       # or set_column_width(1, 20)
ws.set_row_height(1, 30)

# Merged cells
ws.merge_cells("A1:C3")
ws.unmerge_cells("A1:C3")
print(ws.merged_cells)             # list of merged ranges

# Freeze panes
ws.freeze_panes = "B2"             # freeze row 1 and column A
ws.freeze_panes = None             # unfreeze

# Sheet view
ws.zoom_scale = 75
ws.show_gridlines = False
ws.tab_color = "FF0000"            # RRGGBB hex string

# Auto-filter
ws.auto_filter_ref = "A1:D1"
```

### Fonts, fills, borders, and alignment

```python
from openexcel import Font, PatternFill, Border, Side, Alignment

cell = ws["A1"]

# Font
cell.font = Font(name="Arial", size=14, bold=True, italic=False,
                 underline="single", color="FFFF0000")   # ARGB hex

# Fill
cell.fill = PatternFill(fill_type="solid", fgColor="FFFFFF00")   # ARGB hex

# Border
cell.border = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="medium"),
    bottom=Side(style="medium"),
)

# Alignment
cell.alignment = Alignment(horizontal="center", vertical="center",
                           wrap_text=True, indent=0)

# Read back
print(cell.font.bold)              # True
print(cell.fill.fill_type)         # "solid"
print(cell.border.left.style)      # "thin"
print(cell.alignment.horizontal)   # "center"
```

All style objects are immutable value types. Assigning a new style object replaces the cell's style while preserving any other styles already set (e.g., setting `.font` leaves `.fill` and `.border` unchanged). Style deduplication is handled automatically — identical styles share a single XF entry in styles.xml.

Available border styles: `"thin"`, `"medium"`, `"thick"`, `"dashed"`, `"dotted"`, `"double"`, `"hair"`, `"mediumDashed"`, and more (any valid OOXML border style string).

### Hyperlinks

```python
cell = ws["A1"]
cell.value = "Visit our site"
cell.hyperlink = "https://example.com"

# Read back after save/load
wb2 = openexcel.load_workbook("output.xlsx")
print(wb2[0]["A1"].hyperlink)   # "https://example.com"
print(wb2[0]["A1"].value)       # "Visit our site"

# Clear a hyperlink
cell.hyperlink = None
```

Hyperlinks are stored in the standard OOXML location (`xl/worksheets/_rels/sheetN.xml.rels`). Both external URLs and internal sheet anchors (e.g. `"#Sheet2!A1"`) are supported. Hyperlinks compose with all other cell properties — a cell can have a hyperlink, a value, a font, and a number format simultaneously.

### Data validation

```python
from openexcel import DataValidation

# Drop-down list
dv = DataValidation(
    type="list",
    formula1='"Yes,No,Maybe"',
    sqref="A1:A100",
    show_error_message=True,
    error_title="Invalid input",
    error_message="Please choose Yes, No, or Maybe.",
)
ws.add_data_validation(dv)

# Whole number in range
dv2 = DataValidation(
    type="whole",
    operator="between",
    formula1="1",
    formula2="10",
    sqref="B1:B50",
    allow_blank=True,
)
ws.add_data_validation(dv2)

# Read back after save/load
wb2 = openexcel.load_workbook("output.xlsx")
validations = wb2[0].data_validations   # list of DataValidation objects
print(validations[0].type)             # "list"
print(validations[0].sqref)            # "A1:A100"
```

`DataValidation` constructor parameters:

| Parameter | Type | Description |
|---|---|---|
| `type` | `str` | `"list"`, `"whole"`, `"decimal"`, `"date"`, `"time"`, `"textLength"`, `"custom"` |
| `operator` | `str` | `"between"`, `"notBetween"`, `"equal"`, `"notEqual"`, `"greaterThan"`, `"lessThan"`, `"greaterThanOrEqual"`, `"lessThanOrEqual"` |
| `formula1` | `str` | First value/formula |
| `formula2` | `str` | Second value/formula (for `between`/`notBetween`) |
| `sqref` | `str` | Cell range (e.g. `"A1:A100"`) |
| `allow_blank` | `bool` | Allow empty cells (default `False`) |
| `show_drop_down` | `bool` | Hide the dropdown arrow for list validations (default `False`) |
| `show_error_message` | `bool` | Show error alert (default `False`) |
| `error_title` | `str` | Error dialog title |
| `error_message` | `str` | Error dialog message |
| `error_style` | `str` | `"stop"`, `"warning"`, or `"information"` |
| `show_input_message` | `bool` | Show input prompt (default `False`) |
| `prompt_title` | `str` | Input prompt title |
| `prompt_message` | `str` | Input prompt message |

### Page setup and print options

```python
from openexcel import PageSetup, PageMargins, PrintOptions

# Page setup
ws.page_setup = PageSetup(
    orientation="landscape",  # "portrait" or "landscape"
    paper_size=9,             # 9 = A4, 1 = Letter
    scale=75,                 # zoom percentage (10–400)
)

# Fit-to-page mode
ws.page_setup = PageSetup(
    fit_to_page=True,
    fit_to_width=1,
    fit_to_height=0,          # 0 = unlimited pages tall
)

# Page margins (in inches)
ws.page_margins = PageMargins(
    left=0.5, right=0.5,
    top=1.0,  bottom=1.0,
    header=0.5, footer=0.5,
)

# Print options
ws.print_options = PrintOptions(
    grid_lines=True,          # print gridlines
    headings=True,            # print row/column headings
    horizontal_centered=True, # center content horizontally on page
    vertical_centered=True,   # center content vertically on page
)

# Read back
ws2 = openexcel.load_workbook("output.xlsx")[0]
print(ws2.page_setup.orientation)       # "landscape"
print(ws2.page_margins.left)            # 0.5
print(ws2.print_options.grid_lines)     # True
```

Common paper size codes: `1` = Letter (8.5×11"), `9` = A4, `5` = Legal, `13` = B5.

### Sheet protection

```python
from openexcel import SheetProtection

# Lock the sheet (prevent edits)
ws.protection = SheetProtection(
    sheet=True,           # enable protection
    format_cells=False,   # deny format-cells (True = allow)
    insert_rows=False,    # deny inserting rows
    delete_rows=False,    # deny deleting rows
)

# Password-protect (password hash must be pre-computed)
ws.protection = SheetProtection(
    sheet=True,
    password_hash="ABCD",         # legacy XOR hash
)

# Read back
ws2 = openexcel.load_workbook("output.xlsx")[0]
prot = ws2.protection
if prot is not None:
    print(prot.sheet)         # True
    print(prot.format_cells)  # False

# Remove protection
ws.protection = None
```

`SheetProtection` constructor parameters — all optional, default `False`/`None`:

| Parameter | Type | Description |
|---|---|---|
| `sheet` | `bool` | Enable sheet protection |
| `objects` | `bool` | Protect objects |
| `scenarios` | `bool` | Protect scenarios |
| `format_cells` | `bool` | **Allow** formatting cells (inverted: `True` = allowed) |
| `format_columns` | `bool` | Allow formatting columns |
| `format_rows` | `bool` | Allow formatting rows |
| `insert_columns` | `bool` | Allow inserting columns |
| `insert_rows` | `bool` | Allow inserting rows |
| `insert_hyperlinks` | `bool` | Allow inserting hyperlinks |
| `delete_columns` | `bool` | Allow deleting columns |
| `delete_rows` | `bool` | Allow deleting rows |
| `select_locked` | `bool` | Allow selecting locked cells |
| `sort` | `bool` | Allow sorting |
| `auto_filter` | `bool` | Allow using auto-filter |
| `pivot_tables` | `bool` | Allow using pivot tables |
| `select_unlocked` | `bool` | Allow selecting unlocked cells |
| `password_hash` | `str` | Legacy XOR password hash |
| `algorithm_name` | `str` | Hash algorithm (e.g. `"SHA-512"`) |
| `hash_value` | `str` | Base64-encoded hash (modern protection) |
| `salt_value` | `str` | Base64-encoded salt |
| `spin_count` | `int` | Spin count for hash iterations |

> **Note on `format_cells` and similar flags**: In OOXML, `formatCells="1"` means the action is *denied*. openexcel inverts this for readability — `format_cells=True` means the action is *allowed*.

### Named ranges

```python
wb.add_defined_name("MyRange", "Sheet1!$A$1:$C$10")
print(wb.defined_names)            # {"MyRange": "Sheet1!$A$1:$C$10"}
```

### Multiple sheets

```python
wb = openexcel.load_workbook("multi.xlsx")

# By index
ws = wb[0]

# By name
ws = wb["Summary"]

# Iterate all sheets
for ws in wb:
    print(ws.title, ws.max_row, ws.max_column)
```

### Context manager

```python
with openexcel.load_workbook("data.xlsx") as wb:
    ws = wb.active
    data = [row for row in ws.iter_rows()]
```

## API reference

### `openexcel.load_workbook(path: str) -> Workbook`

Open an existing `.xlsx` file. Parses the entire file on load; subsequent access is from memory.

### `openexcel.Workbook()`

Create a new empty workbook.

### `Workbook`

| Member | Description |
|---|---|
| `.active` | First sheet |
| `[0]` / `["Name"]` | Sheet by index or name |
| `.create_sheet(name)` | Add a new sheet |
| `.save(path)` | Write to `.xlsx`. GIL released during write. |
| `.add_defined_name(name, value)` | Add a named range |
| `.defined_names` | Dict of all named ranges |

### `Worksheet`

| Member | Description |
|---|---|
| `.title` | Sheet name |
| `.max_row` / `.max_column` | Dimensions |
| `.append(row)` | Append a row of values |
| `.iter_rows(min_row, max_row, min_col, max_col)` | Iterate rows as value tuples |
| `["A1"]` | Get `Cell` by A1 notation |
| `["A1:C3"]` | Get range as tuple-of-tuples of `Cell` |
| `.cell(row, column)` | Get `Cell` by 1-based row/col |
| `.merge_cells(range_str)` | Merge a cell range |
| `.unmerge_cells(range_str)` | Remove a merge |
| `.merged_cells` | List of merged ranges |
| `.set_column_width(col, width)` | Set column width in character units |
| `.set_row_height(row, height)` | Set row height in points |
| `.freeze_panes` | Get/set freeze pane cell (e.g. `"B2"`) or `None` |
| `.zoom_scale` | Get/set zoom percentage (e.g. `75`) |
| `.show_gridlines` | Get/set gridline visibility |
| `.tab_color` | Get/set tab color as RRGGBB hex string |
| `.auto_filter_ref` | Get/set auto-filter range |
| `.add_data_validation(dv)` | Add a `DataValidation` rule to the sheet |
| `.data_validations` | List of `DataValidation` objects on this sheet |
| `.page_setup` | Get/set `PageSetup` object (orientation, paper size, scale) |
| `.page_margins` | Get/set `PageMargins` object (left, right, top, bottom, header, footer) |
| `.print_options` | Get/set `PrintOptions` object (gridlines, headings, centering) |
| `.protection` | Get/set `SheetProtection` object, or `None` to clear |

### `Cell`

| Member | Description |
|---|---|
| `.value` | Get/set cell value |
| `.number_format` | Get/set format string (e.g. `"0.00%"`) |
| `.font` | Get/set `Font` object |
| `.fill` | Get/set `PatternFill` object |
| `.border` | Get/set `Border` object |
| `.alignment` | Get/set `Alignment` object |
| `.hyperlink` | Get/set hyperlink URL string, or `None` |
| `.row` | Row number (1-based) |
| `.column` | Column number (1-based) |
| `.coordinate` | A1-style coordinate string |
| `.column_letter` | Column letter string |
| `.data_type` | `"n"` number, `"s"` string, `"b"` bool, `"d"` date, `"f"` formula, `"e"` error |

## Differences from openpyxl

| Feature | openexcel | openpyxl |
|---|---|---|
| Read speed | **~5× faster** (C, SAX) | Baseline |
| Write speed | **~8× faster** (C, buffer) | Baseline |
| Cell objects | `Cell` with value, format, coordinate, data_type | Full `Cell` |
| Number formats | Read/write format strings | Full format API |
| Formulas | Read/write formula strings (no evaluation) | Read/write formula strings |
| Merged cells | Supported | Supported |
| Named ranges | Supported | Supported |
| Freeze panes / zoom | Supported | Supported |
| Column/row dimensions | Supported | Supported |
| Font / fill / border / alignment | **Supported** (`Font`, `PatternFill`, `Border`, `Side`, `Alignment`) | Full style API |
| Hyperlinks | **Supported** (`cell.hyperlink`) | Supported |
| Data validation | **Supported** (`DataValidation`, `ws.add_data_validation`) | Supported |
| Page setup / margins | **Supported** (`PageSetup`, `PageMargins`, `PrintOptions`) | Supported |
| Sheet protection | **Supported** (`SheetProtection`, `ws.protection`) | Supported |
| Conditional formatting | Not yet supported | Supported |
| Charts / images | Not yet supported | Supported |
| Comments | Not yet supported | Supported |

## License

MIT — see [LICENSE](LICENSE).

Vendored dependencies:
- [miniz](https://github.com/richgel999/miniz) — public domain
- [khash](https://github.com/attractivechaos/klib) — MIT
