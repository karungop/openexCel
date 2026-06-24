import openexcel
from conftest import fixture_path

def test_shared_strings():
    wb = openexcel.load_workbook(fixture_path("strings.xlsx"))
    rows = list(wb.active)
    assert rows[0][0] == "hello 0"
    assert rows[0][1] == "world 0"
    assert rows[0][2] == "constant"

def test_string_count():
    wb = openexcel.load_workbook(fixture_path("strings.xlsx"))
    rows = list(wb.active)
    assert len(rows) == 100

def test_last_row_strings():
    wb = openexcel.load_workbook(fixture_path("strings.xlsx"))
    rows = list(wb.active)
    assert rows[99][0] == "hello 99"
    assert rows[99][2] == "constant"
