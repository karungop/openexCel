import openexcel
from conftest import fixture_path

def test_open_returns_workbook():
    wb = openexcel.load_workbook(fixture_path("simple.xlsx"))
    assert wb is not None

def test_active_sheet():
    wb = openexcel.load_workbook(fixture_path("simple.xlsx"))
    ws = wb.active
    assert ws is not None

def test_read_integers():
    wb = openexcel.load_workbook(fixture_path("simple.xlsx"))
    rows = list(wb.active)
    assert rows[0] == (1.0, 2.0, 3.0)
    assert rows[1] == (4.0, 5.0, 6.0)

def test_read_floats():
    wb = openexcel.load_workbook(fixture_path("simple.xlsx"))
    rows = list(wb.active)
    assert rows[2] == (1.5, 2.5, 3.5)

def test_max_row():
    wb = openexcel.load_workbook(fixture_path("simple.xlsx"))
    assert wb.active.max_row == 3

def test_max_column():
    wb = openexcel.load_workbook(fixture_path("simple.xlsx"))
    assert wb.active.max_column == 3

def test_sheet_title():
    wb = openexcel.load_workbook(fixture_path("simple.xlsx"))
    assert wb.active.title == "Data"

def test_getitem_by_name():
    wb = openexcel.load_workbook(fixture_path("simple.xlsx"))
    ws = wb["Data"]
    assert ws is not None

def test_iter_rows():
    wb = openexcel.load_workbook(fixture_path("simple.xlsx"))
    rows = list(wb.active.iter_rows(min_row=0, max_row=0))
    assert len(rows) == 1
    assert rows[0] == (1.0, 2.0, 3.0)
