"""Tests for Phase 16: Conditional Formatting"""
import pytest
import tempfile
import os
import openexcel


def roundtrip(wb):
    """Save and reload a workbook."""
    tmp = tempfile.mktemp(suffix=".xlsx")
    wb.save(tmp)
    wb2 = openexcel.load_workbook(tmp)
    os.unlink(tmp)
    return wb2


def test_cfrule_basic_creation():
    rule = openexcel.ConditionalFormattingRule(
        type="cellIs", operator="greaterThan", formula="100",
    )
    assert rule.type == "cellIs"
    assert rule.operator == "greaterThan"
    assert rule.formula == "100"
    assert rule.formula2 is None
    assert rule.priority == 0
    assert rule.stop_if_true == 0


def test_cfrule_defaults():
    rule = openexcel.ConditionalFormattingRule(type="cellIs")
    assert rule.top10_top == 1
    assert rule.top10_rank == 10
    assert rule.above_average == 1
    assert rule.equal_average == 0
    assert rule.data_bar_show_value == 1


def test_add_cf_returns_none():
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    rule = openexcel.ConditionalFormattingRule(
        type="cellIs", operator="greaterThan", formula="0",
    )
    result = ws.add_conditional_formatting("A1:A10", rule)
    assert result is None


def test_conditional_formatting_property_empty():
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    assert ws.conditional_formatting == []


def test_conditional_formatting_one_rule():
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    rule = openexcel.ConditionalFormattingRule(
        type="cellIs", operator="greaterThan", formula="100",
    )
    ws.add_conditional_formatting("A1:A10", rule)
    cf = ws.conditional_formatting
    assert len(cf) == 1
    sqref, rules = cf[0]
    assert sqref == "A1:A10"
    assert len(rules) == 1
    assert rules[0].type == "cellIs"
    assert rules[0].operator == "greaterThan"
    assert rules[0].formula == "100"


def test_multiple_rules_same_sqref():
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    rule1 = openexcel.ConditionalFormattingRule(
        type="cellIs", operator="greaterThan", formula="100", priority=1,
    )
    rule2 = openexcel.ConditionalFormattingRule(
        type="cellIs", operator="lessThan", formula="0", priority=2,
    )
    ws.add_conditional_formatting("A1:A10", rule1)
    ws.add_conditional_formatting("A1:A10", rule2)
    cf = ws.conditional_formatting
    assert len(cf) == 1
    sqref, rules = cf[0]
    assert sqref == "A1:A10"
    assert len(rules) == 2


def test_multiple_sqrefs():
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    ws.add_conditional_formatting("A1:A10",
        openexcel.ConditionalFormattingRule(type="cellIs", operator="greaterThan", formula="0"))
    ws.add_conditional_formatting("B1:B10",
        openexcel.ConditionalFormattingRule(type="cellIs", operator="lessThan", formula="0"))
    assert len(ws.conditional_formatting) == 2


def test_priority_stored():
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    rule = openexcel.ConditionalFormattingRule(
        type="cellIs", operator="greaterThan", formula="50", priority=7,
    )
    ws.add_conditional_formatting("A1:A10", rule)
    sqref, rules = ws.conditional_formatting[0]
    assert rules[0].priority == 7


def test_formula2_stored():
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    rule = openexcel.ConditionalFormattingRule(
        type="cellIs", operator="between", formula="10", formula2="20",
    )
    ws.add_conditional_formatting("B1:B5", rule)
    sqref, rules = ws.conditional_formatting[0]
    assert rules[0].operator == "between"
    assert rules[0].formula == "10"
    assert rules[0].formula2 == "20"


def test_text_stored():
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    rule = openexcel.ConditionalFormattingRule(
        type="containsText", text="hello",
        formula='NOT(ISERROR(SEARCH("hello",A1)))',
    )
    ws.add_conditional_formatting("A1:A10", rule)
    sqref, rules = ws.conditional_formatting[0]
    assert rules[0].type == "containsText"
    assert rules[0].text == "hello"
    assert "SEARCH" in rules[0].formula


def test_expression_rule():
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    rule = openexcel.ConditionalFormattingRule(
        type="expression", formula="A1>AVERAGE(A:A)",
    )
    ws.add_conditional_formatting("A1:A20", rule)
    sqref, rules = ws.conditional_formatting[0]
    assert rules[0].type == "expression"
    assert "AVERAGE" in rules[0].formula


def test_color_scale_rule():
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    rule = openexcel.ConditionalFormattingRule(
        type="colorScale",
        color_scale=["FFFF0000", "FF00FF00"],
    )
    ws.add_conditional_formatting("A1:A10", rule)
    sqref, rules = ws.conditional_formatting[0]
    assert rules[0].type == "colorScale"


def test_data_bar_rule():
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    rule = openexcel.ConditionalFormattingRule(
        type="dataBar", data_bar_color="FF0070C0",
    )
    ws.add_conditional_formatting("A1:A10", rule)
    sqref, rules = ws.conditional_formatting[0]
    assert rules[0].type == "dataBar"


def test_top10_rule_properties():
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    rule = openexcel.ConditionalFormattingRule(
        type="top10", top10_rank=5, top10_top=True, top10_percent=False,
    )
    ws.add_conditional_formatting("A1:A100", rule)
    sqref, rules = ws.conditional_formatting[0]
    assert rules[0].type == "top10"
    assert rules[0].top10_rank == 5
    assert rules[0].top10_top == 1


def test_above_average_rule():
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    rule = openexcel.ConditionalFormattingRule(
        type="aboveAverage", above_average=True,
    )
    ws.add_conditional_formatting("A1:A10", rule)
    sqref, rules = ws.conditional_formatting[0]
    assert rules[0].type == "aboveAverage"
    assert rules[0].above_average == 1


def test_with_fill_styling():
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    rule = openexcel.ConditionalFormattingRule(
        type="cellIs", operator="greaterThan", formula="100",
        fill=openexcel.PatternFill(fill_type="solid", fgColor="FF90EE90"),
    )
    ws.add_conditional_formatting("A1:A10", rule)
    sqref, rules = ws.conditional_formatting[0]
    assert rules[0].type == "cellIs"


def test_save_does_not_crash():
    """Test that saving with conditional formatting doesn't crash."""
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    rule = openexcel.ConditionalFormattingRule(
        type="cellIs", operator="greaterThan", formula="100",
        fill=openexcel.PatternFill(fill_type="solid", fgColor="FFFF0000"),
    )
    ws.add_conditional_formatting("A1:A10", rule)
    tmp = tempfile.mktemp(suffix=".xlsx")
    wb.save(tmp)
    assert os.path.exists(tmp)
    os.unlink(tmp)


def test_save_color_scale_does_not_crash():
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    rule = openexcel.ConditionalFormattingRule(
        type="colorScale",
        color_scale=["FFFF0000", "FFFFFF00", "FF00FF00"],
    )
    ws.add_conditional_formatting("A1:A10", rule)
    tmp = tempfile.mktemp(suffix=".xlsx")
    wb.save(tmp)
    assert os.path.exists(tmp)
    os.unlink(tmp)


def test_save_data_bar_does_not_crash():
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    rule = openexcel.ConditionalFormattingRule(
        type="dataBar", data_bar_color="FF0070C0",
    )
    ws.add_conditional_formatting("A1:A10", rule)
    tmp = tempfile.mktemp(suffix=".xlsx")
    wb.save(tmp)
    assert os.path.exists(tmp)
    os.unlink(tmp)


def test_no_cf_empty_list():
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    assert ws.conditional_formatting == []
    tmp = tempfile.mktemp(suffix=".xlsx")
    wb.save(tmp)
    assert os.path.exists(tmp)
    os.unlink(tmp)


def test_duplicate_values_rule():
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    rule = openexcel.ConditionalFormattingRule(type="duplicateValues")
    ws.add_conditional_formatting("A1:A10", rule)
    sqref, rules = ws.conditional_formatting[0]
    assert rules[0].type == "duplicateValues"


def test_unique_values_rule():
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    rule = openexcel.ConditionalFormattingRule(type="uniqueValues")
    ws.add_conditional_formatting("A1:A10", rule)
    sqref, rules = ws.conditional_formatting[0]
    assert rules[0].type == "uniqueValues"


def test_cfrule_with_font():
    rule = openexcel.ConditionalFormattingRule(
        type="cellIs", operator="greaterThan", formula="0",
        font=openexcel.Font(bold=True, color="FFFF0000"),
    )
    assert rule.font is not None
    assert rule.font != None  # not None object


def test_cfrule_color_scale_list():
    rule = openexcel.ConditionalFormattingRule(
        type="colorScale",
        color_scale=["FFFF0000", "FF00FF00"],
    )
    assert rule.color_scale is not None
    assert len(rule.color_scale) == 2


def test_stop_if_true():
    wb = openexcel.Workbook()
    ws = wb.create_sheet("Sheet1")
    rule = openexcel.ConditionalFormattingRule(
        type="cellIs", operator="greaterThan", formula="100",
        stop_if_true=True,
    )
    ws.add_conditional_formatting("A1:A10", rule)
    sqref, rules = ws.conditional_formatting[0]
    assert rules[0].stop_if_true == 1
