import datetime, openexcel
from conftest import fixture_path

def test_roundtrip_mixed(tmp_path):
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    ws.append([1, "hello", 3.14, True, datetime.date(2024, 1, 1)])
    out = str(tmp_path / "rt.xlsx")
    wb.save(out)

    wb2 = openexcel.load_workbook(out)
    row = list(wb2.active)[0]
    assert row[0] == 1.0
    assert row[1] == "hello"
    assert abs(row[2] - 3.14) < 1e-9
    assert row[3] is True
    assert isinstance(row[4], datetime.date)
    assert row[4] == datetime.date(2024, 1, 1)

def test_roundtrip_many_rows(tmp_path):
    N = 1000
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Big")
    for i in range(N):
        ws.append([i, f"row{i}", float(i) * 1.1])
    out = str(tmp_path / "big.xlsx")
    wb.save(out)

    wb2 = openexcel.load_workbook(out)
    rows = list(wb2.active)
    assert len(rows) == N
    assert rows[0] == (0.0, "row0", 0.0)
    assert rows[999][0] == 999.0
    assert rows[999][1] == "row999"

def test_context_manager(tmp_path):
    out = str(tmp_path / "ctx.xlsx")
    with openexcel.Workbook() as wb:
        ws = wb.create_sheet("S")
        ws.append([42])
        wb.save(out)
    wb2 = openexcel.load_workbook(out)
    assert list(wb2.active)[0][0] == 42.0

def test_read_existing_fixture_roundtrip(tmp_path):
    # Read openpyxl-written fixture, write with openexcel, read back
    wb = openexcel.load_workbook(fixture_path("simple.xlsx"))
    out = str(tmp_path / "rewrite.xlsx")
    wb.save(out)
    wb2 = openexcel.load_workbook(out)
    rows = list(wb2.active)
    assert rows[0] == (1.0, 2.0, 3.0)
