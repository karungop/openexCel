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

### `Cell`

| Member | Description |
|---|---|
| `.value` | Get/set cell value |
| `.number_format` | Get/set format string (e.g. `"0.00%"`) |
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
| Font / fill / border | Not yet supported | Full style API |
| Conditional formatting | Not yet supported | Supported |
| Charts / images | Not yet supported | Supported |
| Data validation | Not yet supported | Supported |
| Comments | Not yet supported | Supported |

## License

MIT — see [LICENSE](LICENSE).

Vendored dependencies:
- [miniz](https://github.com/richgel999/miniz) — public domain
- [khash](https://github.com/attractivechaos/klib) — MIT
