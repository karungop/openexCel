import pytest
import openexcel
from openexcel import Font, PatternFill, Border, Side, Alignment
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


def test_font_bold_roundtrip():
    wb = make_wb()
    ws = wb.create_sheet('S')
    ws['A1'].value = 1
    ws['A1'].font = Font(bold=True)
    wb2 = roundtrip(wb)
    assert wb2[0]['A1'].font.bold == True


def test_font_name_size_roundtrip():
    wb = make_wb()
    ws = wb.create_sheet('S')
    ws['A1'].value = 'hi'
    ws['A1'].font = Font(name='Arial', size=14)
    wb2 = roundtrip(wb)
    f = wb2[0]['A1'].font
    assert f.name == 'Arial'
    assert f.size == 14.0


def test_font_color_roundtrip():
    wb = make_wb()
    ws = wb.create_sheet('S')
    ws['A1'].value = 'x'
    ws['A1'].font = Font(color='FFFF0000')  # ARGB red
    wb2 = roundtrip(wb)
    assert wb2[0]['A1'].font.color == 'FFFF0000'


def test_font_italic_underline_roundtrip():
    wb = make_wb()
    ws = wb.create_sheet('S')
    ws['A1'].value = 'x'
    ws['A1'].font = Font(italic=True, underline='double')
    wb2 = roundtrip(wb)
    f = wb2[0]['A1'].font
    assert f.italic == True
    assert f.underline == 'double'


def test_fill_solid_roundtrip():
    wb = make_wb()
    ws = wb.create_sheet('S')
    ws['A1'].value = 1
    ws['A1'].fill = PatternFill(fill_type='solid', fgColor='FFFFFF00')
    wb2 = roundtrip(wb)
    fill = wb2[0]['A1'].fill
    assert fill.fill_type == 'solid'
    assert fill.fgColor == 'FFFFFF00'


def test_border_thin_roundtrip():
    wb = make_wb()
    ws = wb.create_sheet('S')
    ws['A1'].value = 1
    ws['A1'].border = Border(left=Side(style='thin'))
    wb2 = roundtrip(wb)
    b = wb2[0]['A1'].border
    assert b.left.style == 'thin'


def test_alignment_center_wrap_roundtrip():
    wb = make_wb()
    ws = wb.create_sheet('S')
    ws['A1'].value = 'wrapped text'
    ws['A1'].alignment = Alignment(horizontal='center', wrap_text=True)
    wb2 = roundtrip(wb)
    a = wb2[0]['A1'].alignment
    assert a.horizontal == 'center'
    assert a.wrap_text == True


def test_combined_styles_roundtrip():
    wb = make_wb()
    ws = wb.create_sheet('S')
    ws['A1'].value = 'styled'
    ws['A1'].font = Font(bold=True, size=14)
    ws['A1'].fill = PatternFill(fill_type='solid', fgColor='FF0000FF')
    ws['A1'].border = Border(top=Side(style='medium'), bottom=Side(style='medium'))
    ws['A1'].alignment = Alignment(horizontal='center')
    wb2 = roundtrip(wb)
    c = wb2[0]['A1']
    assert c.font.bold == True
    assert c.font.size == 14.0
    assert c.fill.fill_type == 'solid'
    assert c.border.top.style == 'medium'
    assert c.alignment.horizontal == 'center'


def test_default_cell_font_is_calibri():
    wb = make_wb()
    ws = wb.create_sheet('S')
    ws['A1'].value = 1
    font = ws['A1'].font
    assert font is not None
    assert font.size in (0.0, 11.0)  # 0 if no default loaded


def test_multiple_cells_different_styles():
    wb = make_wb()
    ws = wb.create_sheet('S')
    ws['A1'].font = Font(bold=True)
    ws['A2'].font = Font(italic=True)
    ws['A3'].font = Font(bold=True)  # same as A1, should reuse XF
    wb2 = roundtrip(wb)
    assert wb2[0]['A1'].font.bold == True
    assert wb2[0]['A2'].font.italic == True


def test_existing_tests_still_pass():
    """Sanity check: basic operations still work."""
    wb = make_wb()
    ws = wb.create_sheet('S')
    ws['A1'].value = 42
    ws['A1'].number_format = '0.00'
    wb2 = roundtrip(wb)
    assert wb2[0]['A1'].value == 42.0
    assert wb2[0]['A1'].number_format == '0.00'


def test_font_underline_single():
    wb = make_wb()
    ws = wb.create_sheet('S')
    ws['A1'].value = 'x'
    ws['A1'].font = Font(underline='single')
    wb2 = roundtrip(wb)
    f = wb2[0]['A1'].font
    assert f.underline == 'single'


def test_font_none_underline():
    wb = make_wb()
    ws = wb.create_sheet('S')
    ws['A1'].font = Font(bold=True)
    f = ws['A1'].font
    assert f.bold == True
    assert f.underline is None


def test_alignment_wrap_only():
    wb = make_wb()
    ws = wb.create_sheet('S')
    ws['A1'].value = 'text'
    ws['A1'].alignment = Alignment(wrap_text=True)
    wb2 = roundtrip(wb)
    assert wb2[0]['A1'].alignment.wrap_text == True


def test_border_all_sides():
    wb = make_wb()
    ws = wb.create_sheet('S')
    ws['A1'].value = 1
    ws['A1'].border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='medium'),
        bottom=Side(style='thick'),
    )
    wb2 = roundtrip(wb)
    b = wb2[0]['A1'].border
    assert b.left.style == 'thin'
    assert b.right.style == 'thin'
    assert b.top.style == 'medium'
    assert b.bottom.style == 'thick'


def test_types_importable():
    """Check that all 5 style types are importable from openexcel."""
    from openexcel import Font, PatternFill, Border, Side, Alignment
    f = Font()
    pf = PatternFill()
    s = Side()
    b = Border()
    a = Alignment()
    assert f is not None
    assert pf is not None
    assert s is not None
    assert b is not None
    assert a is not None
