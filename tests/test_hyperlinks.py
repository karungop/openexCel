import pytest
import openexcel
import tempfile
import os


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


def test_hyperlink_set_get():
    wb = make_wb()
    ws = wb.create_sheet('S')
    ws['A1'].value = 'Click me'
    ws['A1'].hyperlink = 'https://example.com'
    assert ws['A1'].hyperlink == 'https://example.com'


def test_hyperlink_roundtrip():
    wb = make_wb()
    ws = wb.create_sheet('S')
    ws['A1'].value = 'Click me'
    ws['A1'].hyperlink = 'https://example.com'
    wb2 = roundtrip(wb)
    assert wb2[0]['A1'].hyperlink == 'https://example.com'
    assert wb2[0]['A1'].value == 'Click me'


def test_hyperlink_multiple_cells():
    wb = make_wb()
    ws = wb.create_sheet('S')
    ws['A1'].value = 'Link 1'
    ws['A1'].hyperlink = 'https://first.com'
    ws['B2'].value = 'Link 2'
    ws['B2'].hyperlink = 'https://second.com'
    wb2 = roundtrip(wb)
    assert wb2[0]['A1'].hyperlink == 'https://first.com'
    assert wb2[0]['B2'].hyperlink == 'https://second.com'


def test_hyperlink_none_default():
    wb = make_wb()
    ws = wb.create_sheet('S')
    ws['A1'].value = 'plain'
    assert ws['A1'].hyperlink is None


def test_hyperlink_clear():
    wb = make_wb()
    ws = wb.create_sheet('S')
    ws['A1'].value = 'x'
    ws['A1'].hyperlink = 'https://example.com'
    ws['A1'].hyperlink = None
    assert ws['A1'].hyperlink is None


def test_hyperlink_preserves_value():
    wb = make_wb()
    ws = wb.create_sheet('S')
    ws['A1'].value = 42
    ws['A1'].hyperlink = 'https://example.com'
    wb2 = roundtrip(wb)
    assert wb2[0]['A1'].value == 42.0
    assert wb2[0]['A1'].hyperlink == 'https://example.com'


def test_hyperlink_preserves_font():
    """Hyperlink and font style can coexist."""
    from openexcel import Font
    wb = make_wb()
    ws = wb.create_sheet('S')
    ws['A1'].value = 'styled link'
    ws['A1'].hyperlink = 'https://example.com'
    ws['A1'].font = Font(bold=True)
    wb2 = roundtrip(wb)
    assert wb2[0]['A1'].hyperlink == 'https://example.com'
    assert wb2[0]['A1'].font.bold is True


def test_existing_tests_not_broken():
    """Sanity: basic write/read still works."""
    wb = make_wb()
    ws = wb.create_sheet('S')
    ws['A1'].value = 'hello'
    ws['B1'].value = 42
    wb2 = roundtrip(wb)
    assert wb2[0]['A1'].value == 'hello'
    assert wb2[0]['B1'].value == 42.0


def test_xlsx_contains_rels_file():
    """A saved file with hyperlinks has the _rels file."""
    import zipfile
    wb = make_wb()
    ws = wb.create_sheet('S')
    ws['A1'].value = 'link'
    ws['A1'].hyperlink = 'https://example.com'
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as f:
        path = f.name
    try:
        wb.save(path)
        with zipfile.ZipFile(path) as zf:
            names = zf.namelist()
        assert any('_rels/sheet' in n and n.endswith('.rels') for n in names), \
            f"No sheet rels file found, got: {names}"
    finally:
        os.unlink(path)
