import datetime, openexcel
from conftest import fixture_path

def test_create_workbook(tmp_path):
    wb = openexcel.Workbook()
    assert wb is not None

def test_create_sheet(tmp_path):
    wb = openexcel.Workbook()
    ws = wb.create_sheet("MySheet")
    assert ws is not None
    assert ws.title == "MySheet"

def test_save_and_reload(tmp_path):
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Data")
    ws.append([1, 2, 3])
    ws.append([4, 5, 6])
    out = str(tmp_path / "out.xlsx")
    wb.save(out)
    wb2 = openexcel.load_workbook(out)
    rows = list(wb2.active)
    assert rows[0] == (1.0, 2.0, 3.0)
    assert rows[1] == (4.0, 5.0, 6.0)

def test_write_strings(tmp_path):
    wb = openexcel.Workbook()
    ws = wb.create_sheet("S")
    ws.append(["hello", "world"])
    out = str(tmp_path / "s.xlsx")
    wb.save(out)
    wb2 = openexcel.load_workbook(out)
    rows = list(wb2.active)
    assert rows[0] == ("hello", "world")

def test_write_bool(tmp_path):
    wb = openexcel.Workbook()
    ws = wb.create_sheet("B")
    ws.append([True, False])
    out = str(tmp_path / "b.xlsx")
    wb.save(out)
    wb2 = openexcel.load_workbook(out)
    rows = list(wb2.active)
    assert rows[0] == (True, False)
