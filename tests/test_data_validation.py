import openexcel
import os
import tempfile
import zipfile
import pytest


def make_wb_with_dv(dv_kwargs, sqref="A1:A10"):
    """Helper: create workbook, add data validation, save to temp file."""
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    ws["A1"].value = "test"
    dv = openexcel.DataValidation(sqref=sqref, **dv_kwargs)
    ws.add_data_validation(dv)
    tmp = tempfile.mktemp(suffix=".xlsx")
    wb.save(tmp)
    return tmp


class TestDataValidation:
    def test_list_validation_roundtrip(self):
        """Basic list validation survives save/load."""
        tmp = make_wb_with_dv({"type": "list", "formula1": '"Yes,No,Maybe"'})
        try:
            wb2 = openexcel.load_workbook(tmp)
            dvs = wb2[0].data_validations
            assert len(dvs) == 1
            assert dvs[0].type == "list"
            assert dvs[0].formula1 == '"Yes,No,Maybe"'
            assert dvs[0].sqref == "A1:A10"
        finally:
            os.unlink(tmp)

    def test_whole_number_between(self):
        """Whole number between validation roundtrip."""
        tmp = make_wb_with_dv({"type": "whole", "operator": "between",
                                "formula1": "1", "formula2": "100"})
        try:
            wb2 = openexcel.load_workbook(tmp)
            dvs = wb2[0].data_validations
            assert dvs[0].type == "whole"
            assert dvs[0].formula1 == "1"
            assert dvs[0].formula2 == "100"
        finally:
            os.unlink(tmp)

    def test_allow_blank_default_true(self):
        """allow_blank defaults to True."""
        dv = openexcel.DataValidation(type="list", formula1='"A,B"', sqref="A1")
        assert dv.allow_blank == True

    def test_show_drop_down_default_false(self):
        """show_drop_down defaults to False (dropdown is visible)."""
        dv = openexcel.DataValidation(type="list", formula1='"A,B"', sqref="A1")
        assert dv.show_drop_down == False

    def test_error_message_roundtrip(self):
        """Error title, message, and style survive save/load."""
        tmp = make_wb_with_dv({
            "type": "list",
            "formula1": '"X,Y"',
            "error_title": "Bad input",
            "error_message": "Please select from the list",
            "error_style": "stop",
            "show_error_message": True,
        })
        try:
            wb2 = openexcel.load_workbook(tmp)
            dv = wb2[0].data_validations[0]
            assert dv.error_title == "Bad input"
            assert dv.error_message == "Please select from the list"
            assert dv.error_style == "stop"
            assert dv.show_error_message == True
        finally:
            os.unlink(tmp)

    def test_prompt_message_roundtrip(self):
        """Prompt title and message survive save/load."""
        tmp = make_wb_with_dv({
            "type": "whole",
            "formula1": "1",
            "prompt_title": "Enter value",
            "prompt_message": "Enter a whole number",
            "show_input_message": True,
        })
        try:
            wb2 = openexcel.load_workbook(tmp)
            dv = wb2[0].data_validations[0]
            assert dv.prompt_title == "Enter value"
            assert dv.prompt_message == "Enter a whole number"
            assert dv.show_input_message == True
        finally:
            os.unlink(tmp)

    def test_multiple_validations(self):
        """Multiple data validations on the same sheet."""
        wb = openexcel.Workbook()
        ws = wb.create_sheet("Sheet1")
        ws.add_data_validation(openexcel.DataValidation(type="list", formula1='"A,B"', sqref="A1:A5"))
        ws.add_data_validation(openexcel.DataValidation(type="whole", formula1="1", formula2="10", sqref="B1:B5"))
        tmp = tempfile.mktemp(suffix=".xlsx")
        wb.save(tmp)
        try:
            wb2 = openexcel.load_workbook(tmp)
            dvs = wb2[0].data_validations
            assert len(dvs) == 2
            assert dvs[0].type == "list"
            assert dvs[1].type == "whole"
        finally:
            os.unlink(tmp)

    def test_xml_has_data_validations_element(self):
        """Verify the XML actually contains <dataValidations> at the right place."""
        tmp = make_wb_with_dv({"type": "list", "formula1": '"Yes,No"'})
        try:
            with zipfile.ZipFile(tmp) as zf:
                xml = zf.read("xl/worksheets/sheet1.xml").decode()
            assert "<dataValidations" in xml
            assert "<dataValidation" in xml
            assert 'type="list"' in xml
            # dataValidations must appear BEFORE hyperlinks (or at end if no hyperlinks)
            if "<hyperlinks" in xml:
                assert xml.index("<dataValidations") < xml.index("<hyperlinks")
        finally:
            os.unlink(tmp)

    def test_no_validation_no_element(self):
        """Workbooks without data validation should have no <dataValidations> element."""
        wb = openexcel.Workbook()
        ws = wb.create_sheet("Sheet1")
        ws["A1"].value = "hello"
        tmp = tempfile.mktemp(suffix=".xlsx")
        wb.save(tmp)
        try:
            with zipfile.ZipFile(tmp) as zf:
                xml = zf.read("xl/worksheets/sheet1.xml").decode()
            assert "<dataValidations" not in xml
        finally:
            os.unlink(tmp)

    def test_data_validation_properties(self):
        """Verify all DataValidation constructor args are accessible as properties."""
        dv = openexcel.DataValidation(
            type="decimal",
            operator="greaterThan",
            formula1="0",
            sqref="C1:C10",
            allow_blank=False,
            error_title="Err",
            error_message="Must be positive",
            error_style="warning",
        )
        assert dv.type == "decimal"
        assert dv.operator == "greaterThan"
        assert dv.formula1 == "0"
        assert dv.sqref == "C1:C10"
        assert dv.allow_blank == False
        assert dv.error_title == "Err"
        assert dv.error_message == "Must be positive"
        assert dv.error_style == "warning"

    def test_data_validations_empty_list_default(self):
        """ws.data_validations returns empty list when no validations added."""
        wb = openexcel.Workbook()
        ws = wb.create_sheet("Sheet1")
        assert ws.data_validations == []

    def test_decimal_validation(self):
        """Decimal type validation roundtrip."""
        tmp = make_wb_with_dv({
            "type": "decimal",
            "operator": "greaterThanOrEqual",
            "formula1": "0.5",
        }, sqref="C1:C20")
        try:
            wb2 = openexcel.load_workbook(tmp)
            dvs = wb2[0].data_validations
            assert len(dvs) == 1
            assert dvs[0].type == "decimal"
            assert dvs[0].sqref == "C1:C20"
        finally:
            os.unlink(tmp)

    def test_none_fields_default(self):
        """Optional fields default to None when not set."""
        dv = openexcel.DataValidation(type="list", formula1='"A,B"', sqref="A1")
        assert dv.formula2 is None
        assert dv.error_title is None
        assert dv.error_message is None
        assert dv.error_style is None
        assert dv.prompt_title is None
        assert dv.prompt_message is None
        assert dv.operator is None

    def test_show_drop_down_true_hides_dropdown(self):
        """show_drop_down=True maps to showDropDown=1 in XML (which hides the dropdown)."""
        wb = openexcel.Workbook()
        ws = wb.create_sheet("Sheet1")
        dv = openexcel.DataValidation(type="list", formula1='"A,B"', sqref="A1", show_drop_down=True)
        assert dv.show_drop_down == True
        ws.add_data_validation(dv)
        tmp = tempfile.mktemp(suffix=".xlsx")
        wb.save(tmp)
        try:
            with zipfile.ZipFile(tmp) as zf:
                xml = zf.read("xl/worksheets/sheet1.xml").decode()
            assert 'showDropDown="1"' in xml
        finally:
            os.unlink(tmp)
