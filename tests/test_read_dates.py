import datetime, openexcel
from conftest import fixture_path

def test_read_date():
    wb = openexcel.load_workbook(fixture_path("dates.xlsx"))
    rows = list(wb.active)
    val = rows[0][0]
    assert isinstance(val, datetime.date)
    assert val == datetime.date(2024, 1, 15)

def test_read_datetime():
    wb = openexcel.load_workbook(fixture_path("dates.xlsx"))
    rows = list(wb.active)
    val = rows[1][0]
    assert isinstance(val, datetime.datetime)
    assert val.year == 2024
    assert val.month == 6
    assert val.day == 1

def test_date_after_1900_leap_bug():
    wb = openexcel.load_workbook(fixture_path("dates.xlsx"))
    rows = list(wb.active)
    val = rows[2][0]
    assert isinstance(val, datetime.date)
    assert val == datetime.date(1900, 3, 1)
