import openexcel
import os
import tempfile
import zipfile
import pytest


class TestSheetProtection:
    def test_basic_protection_roundtrip(self):
        wb = openexcel.Workbook()
        ws = wb.create_sheet("Sheet1")
        ws["A1"].value = "protected"
        ws.protection = openexcel.SheetProtection(sheet=True)
        tmp = tempfile.mktemp(suffix=".xlsx")
        wb.save(tmp)
        try:
            wb2 = openexcel.load_workbook(tmp)
            p = wb2[0].protection
            assert p is not None
            assert p.sheet == True
        finally:
            os.unlink(tmp)

    def test_no_protection_returns_none(self):
        wb = openexcel.Workbook()
        ws = wb.create_sheet("Sheet1")
        assert ws.protection is None

    def test_xml_has_sheet_protection_element(self):
        wb = openexcel.Workbook()
        ws = wb.create_sheet("Sheet1")
        ws.protection = openexcel.SheetProtection(sheet=True)
        tmp = tempfile.mktemp(suffix=".xlsx")
        wb.save(tmp)
        try:
            with zipfile.ZipFile(tmp) as zf:
                xml = zf.read("xl/worksheets/sheet1.xml").decode()
            assert "<sheetProtection" in xml
            assert 'sheet="1"' in xml
        finally:
            os.unlink(tmp)

    def test_no_protection_no_xml_element(self):
        wb = openexcel.Workbook()
        ws = wb.create_sheet("Sheet1")
        ws["A1"].value = "hello"
        tmp = tempfile.mktemp(suffix=".xlsx")
        wb.save(tmp)
        try:
            with zipfile.ZipFile(tmp) as zf:
                xml = zf.read("xl/worksheets/sheet1.xml").decode()
            assert "<sheetProtection" not in xml
        finally:
            os.unlink(tmp)

    def test_allow_format_cells_roundtrip(self):
        wb = openexcel.Workbook()
        ws = wb.create_sheet("Sheet1")
        ws.protection = openexcel.SheetProtection(sheet=True, format_cells=True)
        tmp = tempfile.mktemp(suffix=".xlsx")
        wb.save(tmp)
        try:
            wb2 = openexcel.load_workbook(tmp)
            p = wb2[0].protection
            assert p.format_cells == True
        finally:
            os.unlink(tmp)

    def test_allow_insert_rows_roundtrip(self):
        wb = openexcel.Workbook()
        ws = wb.create_sheet("Sheet1")
        ws.protection = openexcel.SheetProtection(sheet=True, insert_rows=True, delete_rows=True)
        tmp = tempfile.mktemp(suffix=".xlsx")
        wb.save(tmp)
        try:
            wb2 = openexcel.load_workbook(tmp)
            p = wb2[0].protection
            assert p.insert_rows == True
            assert p.delete_rows == True
        finally:
            os.unlink(tmp)

    def test_objects_and_scenarios_roundtrip(self):
        wb = openexcel.Workbook()
        ws = wb.create_sheet("Sheet1")
        ws.protection = openexcel.SheetProtection(sheet=True, objects=True, scenarios=True)
        tmp = tempfile.mktemp(suffix=".xlsx")
        wb.save(tmp)
        try:
            wb2 = openexcel.load_workbook(tmp)
            p = wb2[0].protection
            assert p.objects == True
            assert p.scenarios == True
        finally:
            os.unlink(tmp)

    def test_protection_element_order(self):
        # sheetProtection must appear BEFORE printOptions in XML
        wb = openexcel.Workbook()
        ws = wb.create_sheet("Sheet1")
        ws.protection = openexcel.SheetProtection(sheet=True)
        tmp = tempfile.mktemp(suffix=".xlsx")
        wb.save(tmp)
        try:
            with zipfile.ZipFile(tmp) as zf:
                xml = zf.read("xl/worksheets/sheet1.xml").decode()
            assert "<sheetProtection" in xml
            # If printOptions is present, sheetProtection must come before it
            if "<printOptions" in xml:
                assert xml.index("<sheetProtection") < xml.index("<printOptions")
        finally:
            os.unlink(tmp)

    def test_protection_default_sheet_true(self):
        p = openexcel.SheetProtection()
        assert p.sheet == True

    def test_protection_default_format_cells_false(self):
        p = openexcel.SheetProtection()
        assert p.format_cells == False

    def test_set_protection_none_clears_it(self):
        wb = openexcel.Workbook()
        ws = wb.create_sheet("Sheet1")
        ws.protection = openexcel.SheetProtection(sheet=True)
        ws.protection = None
        tmp = tempfile.mktemp(suffix=".xlsx")
        wb.save(tmp)
        try:
            with zipfile.ZipFile(tmp) as zf:
                xml = zf.read("xl/worksheets/sheet1.xml").decode()
            assert "<sheetProtection" not in xml
        finally:
            os.unlink(tmp)

    def test_xml_format_cells_attribute_value(self):
        # When format_cells=True (allowed), the XML should have formatCells="0"
        wb = openexcel.Workbook()
        ws = wb.create_sheet("Sheet1")
        ws.protection = openexcel.SheetProtection(sheet=True, format_cells=True)
        tmp = tempfile.mktemp(suffix=".xlsx")
        wb.save(tmp)
        try:
            with zipfile.ZipFile(tmp) as zf:
                xml = zf.read("xl/worksheets/sheet1.xml").decode()
            assert 'formatCells="0"' in xml
        finally:
            os.unlink(tmp)

    def test_xml_insert_rows_attribute_value(self):
        # When insert_rows=True (allowed), the XML should have insertRows="0"
        wb = openexcel.Workbook()
        ws = wb.create_sheet("Sheet1")
        ws.protection = openexcel.SheetProtection(sheet=True, insert_rows=True)
        tmp = tempfile.mktemp(suffix=".xlsx")
        wb.save(tmp)
        try:
            with zipfile.ZipFile(tmp) as zf:
                xml = zf.read("xl/worksheets/sheet1.xml").decode()
            assert 'insertRows="0"' in xml
        finally:
            os.unlink(tmp)
