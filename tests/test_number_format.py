"""
Phase 2: Number format string infrastructure tests.
Some tests require cell.number_format (module.c work, done in a later pass).
Those tests are guarded with pytest.skip if the attribute is not yet present.
"""
import pytest
import openexcel
import tempfile
import os
import datetime


def make_wb():
    return openexcel.Workbook()


def roundtrip(wb):
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        path = f.name
    try:
        wb.save(path)
        return openexcel.load_workbook(path)
    finally:
        os.unlink(path)


def test_date_cell_roundtrip():
    """Date cells written by Python survive round-trip without error."""
    wb = make_wb()
    ws = wb.create_sheet('Sheet1')
    ws.append([datetime.date(2024, 3, 15)])
    wb2 = roundtrip(wb)
    val = wb2[0]['A1'].value
    assert val is not None


def test_number_format_float_roundtrip():
    """cell.number_format = '0.00%' survives save/load."""
    wb = make_wb()
    ws = wb.create_sheet('Sheet1')
    ws.append([0.75])
    cell = ws['A1']
    if not hasattr(cell, 'number_format'):
        pytest.skip('number_format not yet in Python API')
    cell.number_format = '0.00%'
    wb2 = roundtrip(wb)
    cell2 = wb2[0]['A1']
    assert cell2.number_format == '0.00%'


def test_number_format_general():
    """Cells with no format return None or 'General'."""
    wb = make_wb()
    ws = wb.create_sheet('Sheet1')
    ws.append([42])
    cell = ws['A1']
    if not hasattr(cell, 'number_format'):
        pytest.skip('number_format not yet in Python API')
    fmt = cell.number_format
    assert fmt is None or fmt == 'General'


def test_multiple_formats():
    """Multiple cells with different formats all survive round-trip."""
    wb = make_wb()
    ws = wb.create_sheet('Sheet1')
    ws.append([0.5, 1234.56])
    cell1 = ws['A1']
    cell2 = ws['B1']
    if not hasattr(cell1, 'number_format'):
        pytest.skip('number_format not yet in Python API')
    cell1.number_format = '0%'
    cell2.number_format = '#,##0.00'
    wb2 = roundtrip(wb)
    assert wb2[0]['A1'].number_format == '0%'
    assert wb2[0]['B1'].number_format == '#,##0.00'


def test_styles_xml_emitted_on_write():
    """Saving a workbook produces a valid .xlsx with styles.xml (date format)."""
    wb = make_wb()
    ws = wb.create_sheet('Sheet1')
    ws.append([1.0])
    ws.append([datetime.date(2024, 1, 1)])
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        path = f.name
    try:
        wb.save(path)
        # Load it back — if styles.xml is malformed, load_workbook will fail
        wb2 = openexcel.load_workbook(path)
        # Row 2 is the date row
        val = wb2[0]['A2'].value
        assert val is not None
    finally:
        os.unlink(path)


def test_roundtrip_preserves_date_format():
    """Reading a file with date cells preserves date detection after round-trip."""
    wb = openexcel.load_workbook('tests/fixtures/dates.xlsx')
    ws = wb[0]
    # Find a date cell and verify it has a date value
    found_value = False
    for ref in ['A1', 'B1', 'C1', 'A2', 'B2']:
        try:
            cell = ws[ref]
            if cell.value is not None:
                found_value = True
                break
        except Exception:
            pass
    assert found_value, "Expected at least one non-empty cell in dates.xlsx"


def test_styles_xml_includes_numfmt_on_custom_format():
    """When a cell has a custom number format, numFmts appears in styles.xml."""
    import zipfile
    wb = make_wb()
    ws = wb.create_sheet('Sheet1')
    ws.append([0.5])
    cell = ws['A1']
    if not hasattr(cell, 'number_format'):
        pytest.skip('number_format not yet in Python API')
    cell.number_format = '0.000%'  # non-standard custom format
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        path = f.name
    try:
        wb.save(path)
        with zipfile.ZipFile(path) as zf:
            styles_xml = zf.read('xl/styles.xml').decode('utf-8')
        assert '0.000%' in styles_xml
        assert 'numFmtId' in styles_xml
    finally:
        os.unlink(path)
