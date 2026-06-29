import pytest
import openexcel
from openexcel import _openexcel


def test_cell_subscript_basic():
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    ws.append([1, "hello", True])
    cell = ws['A1']
    assert cell.value == 1.0
    assert cell.row == 1
    assert cell.column == 1
    assert cell.coordinate == "A1"
    assert cell.column_letter == "A"


def test_cell_method():
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    ws.append([42, "world"])
    cell = ws.cell(row=1, column=2)
    assert cell.value == "world"
    assert cell.coordinate == "B1"


def test_cell_set_value(tmp_path):
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    ws.append([0])
    ws['A1'].value = 99
    path = str(tmp_path / "out.xlsx")
    wb.save(path)
    wb2 = openexcel.load_workbook(path)
    assert wb2[0]['A1'].value == 99.0


def test_cell_range_slice():
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    ws.append([1, 2, 3])
    ws.append([4, 5, 6])
    cells = ws['A1:C2']
    assert len(cells) == 2       # 2 rows
    assert len(cells[0]) == 3   # 3 cols
    assert cells[0][0].value == 1.0
    assert cells[1][2].value == 6.0


def test_cell_new_coordinate():
    # Setting value at a coordinate that doesn't exist yet
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    ws['B3'].value = "inserted"
    assert ws['B3'].value == "inserted"


def test_column_utilities():
    assert _openexcel.column_index_from_string("A") == 1
    assert _openexcel.column_index_from_string("Z") == 26
    assert _openexcel.column_index_from_string("AA") == 27
    assert _openexcel.get_column_letter(1) == "A"
    assert _openexcel.get_column_letter(26) == "Z"
    assert _openexcel.get_column_letter(27) == "AA"


def test_cell_repr():
    wb = openexcel.Workbook()
    ws = wb.create_sheet("MySheet")
    ws.append([1])
    cell = ws['A1']
    assert repr(cell) == "<Cell 'MySheet'.A1>"


def test_cell_data_type():
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    ws.append([1.5, "text", True])
    assert ws['A1'].data_type == 'n'
    assert ws['B1'].data_type == 's'
    assert ws['C1'].data_type == 'b'


def test_cell_tuple_subscript():
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    ws.append([10, 20, 30])
    cell = ws[(1, 3)]
    assert cell.value == 30.0
    assert cell.coordinate == "C1"


def test_cell_large_column():
    assert _openexcel.column_index_from_string("AZ") == 52
    assert _openexcel.get_column_letter(52) == "AZ"
    assert _openexcel.column_index_from_string("AAA") == 703
    assert _openexcel.get_column_letter(703) == "AAA"


def test_cell_set_none_existing():
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    ws.append([42])
    ws['A1'].value = None
    # After setting to None, value should be None (empty cell)
    assert ws['A1'].value is None


def test_cell_overwrite_string():
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    ws.append(["original"])
    ws['A1'].value = "updated"
    assert ws['A1'].value == "updated"
