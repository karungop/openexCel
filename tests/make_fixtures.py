"""Generate test fixture .xlsx files using openpyxl.
Run once: python tests/make_fixtures.py
"""
import datetime, pathlib, openpyxl

here = pathlib.Path(__file__).parent / "fixtures"
here.mkdir(exist_ok=True)

# simple.xlsx — integers, floats, strings
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Data"
ws.append([1, 2, 3])
ws.append([4, 5, 6])
ws.append([1.5, 2.5, 3.5])
wb.save(here / "simple.xlsx")
print("simple.xlsx")

# strings.xlsx — shared strings
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Strings"
for i in range(100):
    ws.append([f"hello {i}", f"world {i}", "constant"])
wb.save(here / "strings.xlsx")
print("strings.xlsx")

# dates.xlsx
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Dates"
ws.append([datetime.date(2024, 1, 15)])
ws.append([datetime.datetime(2024, 6, 1, 12, 30, 0)])
ws.append([datetime.date(1900, 3, 1)])  # after the phantom leap day
wb.save(here / "dates.xlsx")
print("dates.xlsx")

# mixed.xlsx — numbers, strings, bools, None
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Mixed"
ws.append([1, "hello", 3.14, True, False, None])
wb.save(here / "mixed.xlsx")
print("mixed.xlsx")

print("All fixtures created.")
