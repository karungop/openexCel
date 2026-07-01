import openexcel
import os
import tempfile
import zipfile
import pytest


class TestPageSetup:
    def test_page_setup_orientation_landscape_roundtrip(self):
        """Landscape orientation survives save/load."""
        wb = openexcel.Workbook()
        ws = wb.create_sheet("Sheet1")
        ws.page_setup = openexcel.PageSetup(orientation="landscape")
        tmp = tempfile.mktemp(suffix=".xlsx")
        wb.save(tmp)
        try:
            wb2 = openexcel.load_workbook(tmp)
            assert wb2[0].page_setup.orientation == "landscape"
        finally:
            os.unlink(tmp)

    def test_page_setup_paper_size_roundtrip(self):
        """paper_size=9 (A4) survives save/load."""
        wb = openexcel.Workbook()
        ws = wb.create_sheet("Sheet1")
        ws.page_setup = openexcel.PageSetup(paper_size=9)
        tmp = tempfile.mktemp(suffix=".xlsx")
        wb.save(tmp)
        try:
            wb2 = openexcel.load_workbook(tmp)
            assert wb2[0].page_setup.paper_size == 9
        finally:
            os.unlink(tmp)

    def test_page_setup_scale_roundtrip(self):
        """scale=75 survives save/load."""
        wb = openexcel.Workbook()
        ws = wb.create_sheet("Sheet1")
        ws.page_setup = openexcel.PageSetup(scale=75)
        tmp = tempfile.mktemp(suffix=".xlsx")
        wb.save(tmp)
        try:
            wb2 = openexcel.load_workbook(tmp)
            assert wb2[0].page_setup.scale == 75
        finally:
            os.unlink(tmp)

    def test_page_setup_fit_to_page_roundtrip(self):
        """fit_to_page=True, fit_to_width=1, fit_to_height=0 survive save/load."""
        wb = openexcel.Workbook()
        ws = wb.create_sheet("Sheet1")
        ws.page_setup = openexcel.PageSetup(fit_to_page=1, fit_to_width=1, fit_to_height=0)
        tmp = tempfile.mktemp(suffix=".xlsx")
        wb.save(tmp)
        try:
            wb2 = openexcel.load_workbook(tmp)
            ps = wb2[0].page_setup
            assert ps.fit_to_page == True
            assert ps.fit_to_width == 1
        finally:
            os.unlink(tmp)

    def test_page_margins_roundtrip(self):
        """Custom page margins survive save/load."""
        wb = openexcel.Workbook()
        ws = wb.create_sheet("Sheet1")
        ws.page_margins = openexcel.PageMargins(
            left=0.5, right=0.5, top=1.0, bottom=1.0, header=0.5, footer=0.5
        )
        tmp = tempfile.mktemp(suffix=".xlsx")
        wb.save(tmp)
        try:
            wb2 = openexcel.load_workbook(tmp)
            m = wb2[0].page_margins
            assert abs(m.left - 0.5) < 0.001
            assert abs(m.right - 0.5) < 0.001
            assert abs(m.top - 1.0) < 0.001
            assert abs(m.bottom - 1.0) < 0.001
            assert abs(m.header - 0.5) < 0.001
            assert abs(m.footer - 0.5) < 0.001
        finally:
            os.unlink(tmp)

    def test_print_options_gridlines_roundtrip(self):
        """grid_lines=True survives save/load."""
        wb = openexcel.Workbook()
        ws = wb.create_sheet("Sheet1")
        ws.print_options = openexcel.PrintOptions(grid_lines=1)
        tmp = tempfile.mktemp(suffix=".xlsx")
        wb.save(tmp)
        try:
            wb2 = openexcel.load_workbook(tmp)
            assert wb2[0].print_options.grid_lines == True
        finally:
            os.unlink(tmp)

    def test_print_options_centered_roundtrip(self):
        """horizontal_centered and vertical_centered survive save/load."""
        wb = openexcel.Workbook()
        ws = wb.create_sheet("Sheet1")
        ws.print_options = openexcel.PrintOptions(horizontal_centered=1, vertical_centered=1)
        tmp = tempfile.mktemp(suffix=".xlsx")
        wb.save(tmp)
        try:
            wb2 = openexcel.load_workbook(tmp)
            po = wb2[0].print_options
            assert po.horizontal_centered == True
            assert po.vertical_centered == True
        finally:
            os.unlink(tmp)

    def test_xml_has_page_setup_element(self):
        """Verify <pageSetup> appears in the XML when page_setup is set."""
        wb = openexcel.Workbook()
        ws = wb.create_sheet("Sheet1")
        ws.page_setup = openexcel.PageSetup(orientation="portrait", paper_size=9)
        tmp = tempfile.mktemp(suffix=".xlsx")
        wb.save(tmp)
        try:
            with zipfile.ZipFile(tmp) as zf:
                xml = zf.read("xl/worksheets/sheet1.xml").decode()
            assert "<pageSetup" in xml
            assert 'orientation="portrait"' in xml
            assert 'paperSize="9"' in xml
        finally:
            os.unlink(tmp)

    def test_no_page_setup_no_element(self):
        """When not set, <pageSetup> should NOT appear in XML."""
        wb = openexcel.Workbook()
        ws = wb.create_sheet("Sheet1")
        ws["A1"].value = "hello"
        tmp = tempfile.mktemp(suffix=".xlsx")
        wb.save(tmp)
        try:
            with zipfile.ZipFile(tmp) as zf:
                xml = zf.read("xl/worksheets/sheet1.xml").decode()
            assert "<pageSetup" not in xml
        finally:
            os.unlink(tmp)

    def test_default_page_margins_properties(self):
        """Default PageMargins object has Excel standard values."""
        m = openexcel.PageMargins()
        assert abs(m.left - 0.7) < 0.001
        assert abs(m.right - 0.7) < 0.001
        assert abs(m.top - 0.75) < 0.001
        assert abs(m.bottom - 0.75) < 0.001
        assert abs(m.header - 0.3) < 0.001
        assert abs(m.footer - 0.3) < 0.001

    def test_page_setup_default_orientation_is_none(self):
        """Default PageSetup has orientation=None."""
        ps = openexcel.PageSetup()
        assert ps.orientation is None

    def test_page_setup_default_numerics_are_zero(self):
        """Default PageSetup has all numeric fields at 0."""
        ps = openexcel.PageSetup()
        assert ps.paper_size == 0
        assert ps.scale == 0
        assert ps.fit_to_width == 0
        assert ps.fit_to_height == 0
        assert ps.fit_to_page == False

    def test_page_setup_element_order(self):
        """pageMargins must appear before pageSetup in XML."""
        wb = openexcel.Workbook()
        ws = wb.create_sheet("Sheet1")
        ws.page_setup = openexcel.PageSetup(orientation="portrait", paper_size=9)
        ws.page_margins = openexcel.PageMargins(
            left=1.0, right=1.0, top=1.0, bottom=1.0, header=0.5, footer=0.5
        )
        tmp = tempfile.mktemp(suffix=".xlsx")
        wb.save(tmp)
        try:
            with zipfile.ZipFile(tmp) as zf:
                xml = zf.read("xl/worksheets/sheet1.xml").decode()
            assert "<pageMargins" in xml
            assert "<pageSetup" in xml
            assert xml.index("<pageMargins") < xml.index("<pageSetup")
        finally:
            os.unlink(tmp)

    def test_xml_has_print_options_element(self):
        """Verify <printOptions> appears in the XML when print_options is set."""
        wb = openexcel.Workbook()
        ws = wb.create_sheet("Sheet1")
        ws.print_options = openexcel.PrintOptions(grid_lines=1, headings=1)
        tmp = tempfile.mktemp(suffix=".xlsx")
        wb.save(tmp)
        try:
            with zipfile.ZipFile(tmp) as zf:
                xml = zf.read("xl/worksheets/sheet1.xml").decode()
            assert "<printOptions" in xml
            assert 'gridLines="1"' in xml
            assert 'headings="1"' in xml
        finally:
            os.unlink(tmp)

    def test_xml_element_order_print_options_before_page_margins(self):
        """printOptions must appear before pageMargins in XML."""
        wb = openexcel.Workbook()
        ws = wb.create_sheet("Sheet1")
        ws.print_options = openexcel.PrintOptions(horizontal_centered=1)
        ws.page_margins = openexcel.PageMargins(left=0.5, right=0.5, top=1.0, bottom=1.0,
                                                header=0.3, footer=0.3)
        ws.page_setup = openexcel.PageSetup(paper_size=1)
        tmp = tempfile.mktemp(suffix=".xlsx")
        wb.save(tmp)
        try:
            with zipfile.ZipFile(tmp) as zf:
                xml = zf.read("xl/worksheets/sheet1.xml").decode()
            assert "<printOptions" in xml
            assert "<pageMargins" in xml
            assert "<pageSetup" in xml
            assert xml.index("<printOptions") < xml.index("<pageMargins")
            assert xml.index("<pageMargins") < xml.index("<pageSetup")
        finally:
            os.unlink(tmp)

    def test_print_options_default_all_false(self):
        """Default PrintOptions has all fields False."""
        po = openexcel.PrintOptions()
        assert po.grid_lines == False
        assert po.headings == False
        assert po.horizontal_centered == False
        assert po.vertical_centered == False

    def test_page_setup_combined_roundtrip(self):
        """Multiple page_setup fields set simultaneously survive save/load."""
        wb = openexcel.Workbook()
        ws = wb.create_sheet("Sheet1")
        ws.page_setup = openexcel.PageSetup(
            orientation="landscape", paper_size=9, scale=50
        )
        tmp = tempfile.mktemp(suffix=".xlsx")
        wb.save(tmp)
        try:
            wb2 = openexcel.load_workbook(tmp)
            ps = wb2[0].page_setup
            assert ps.orientation == "landscape"
            assert ps.paper_size == 9
            assert ps.scale == 50
        finally:
            os.unlink(tmp)
