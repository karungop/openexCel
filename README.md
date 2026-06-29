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

**~8× faster writes, ~5× faster reads** across all sizes tested. Speedup is consistent — there is no warm-up effect.

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

Cell values map to Python types:

| Excel type | Python type |
|---|---|
| Number | `float` |
| String | `str` |
| Boolean | `bool` |
| Date / datetime | `datetime.date` / `datetime.datetime` |
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
# file resources released on exit
```

## API reference

### `openexcel.load_workbook(path: str) -> Workbook`

Open an existing `.xlsx` file for reading. Parses the entire file on load; subsequent access is from memory.

### `openexcel.Workbook()`

Create a new empty workbook.

### `Workbook.create_sheet(name: str) -> Worksheet`

Add a new sheet and return it.

### `Workbook.save(path: str)`

Write the workbook to an `.xlsx` file. GIL is released during write.

### `Workbook.active -> Worksheet`

Returns the first sheet.

### `Worksheet.iter_rows(min_row=None, max_row=None, min_col=None, max_col=None)`

Yields one tuple per row. Indices are 1-based and inclusive, matching openpyxl's convention.

### `Worksheet.append(row: list | tuple)`

Append a row of values. Accepts `int`, `float`, `str`, `bool`, `None`, `datetime.date`, `datetime.datetime`.

### `Worksheet.max_row -> int`

### `Worksheet.max_column -> int`

### `Worksheet.title -> str`

## Differences from openpyxl

| Feature | openexcel | openpyxl |
|---|---|---|
| Read speed | **Fast** (C, SAX) | Slow (pure Python) |
| Write speed | **Fast** (C, arena buffer) | Slow (pure Python) |
| Cell objects | Not exposed (values only) | Full `Cell` with font/fill/border |
| Formulas | Read result value only | Read formula string |
| Styles / formatting | Date detection only | Full style API |
| Merged cells | Not supported | Supported |
| Charts / images | Not supported | Supported |

openexcel is optimized for **data workloads** — reading and writing large tables of values. If you need full formatting control or chart support, use openpyxl.

## License

MIT — see [LICENSE](LICENSE).

Vendored dependencies:
- [miniz](https://github.com/richgel999/miniz) — public domain
- [khash](https://github.com/attractivechaos/klib) — MIT
