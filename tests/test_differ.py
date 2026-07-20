"""Pure tests for VBA-aware component comparison semantics."""

import pytest

from xlvbatools.vba.differ import _compare, _normalize_vba_line


@pytest.mark.unit
def test_vba_diff_treats_identifier_and_keyword_case_as_equivalent():
    result = _compare(
        "modMain",
        ["PUBLIC Sub CalculateTotal()", "    fileCOUNT = 1", "End SUB"],
        ["Public Sub CalculateTotal()", "    FileCount = 1", "End Sub"],
        "modMain.bas",
    )

    assert result == {
        "name": "modMain",
        "status": "equivalent",
        "comparison": "vba",
        "equivalence": "vba_token_equivalent",
        "lines_added": 0,
        "lines_removed": 0,
    }


@pytest.mark.unit
@pytest.mark.parametrize(
    ("workbook", "source"),
    [
        ('Debug.Print "READY"', 'Debug.Print "ready"'),
        ("value = 1 ' Important", "value = 1 ' important"),
        ("Rem Important", "REM important"),
        ("10 Rem Important", "10 REM important"),
        ("value = 1: Rem Important", "VALUE = 1: REM important"),
    ],
)
def test_vba_diff_preserves_literal_and_comment_case(workbook, source):
    result = _compare("modMain", [workbook], [source], "modMain.bas")

    assert result["status"] == "modified"
    assert result["lines_added"] == 1
    assert result["lines_removed"] == 1


@pytest.mark.unit
def test_raw_text_comparison_reports_identifier_case_changes():
    result = _compare(
        "modMain", ["fileCOUNT = 1"], ["FileCount = 1"], "modMain.bas",
        comparison="text",
    )

    assert result["status"] == "modified"
    assert result["comparison"] == "text"


@pytest.mark.unit
def test_vba_diff_ignores_insignificant_spacing_between_code_tokens():
    result = _compare(
        "modMain", ["If ready Then total=total+1"],
        ["IF  Ready   THEN total = total + 1"], "modMain.bas",
    )

    assert result["status"] == "equivalent"
    assert result["equivalence"] == "vba_token_equivalent"


@pytest.mark.unit
@pytest.mark.parametrize(
    ("workbook", "source"),
    [("value = [A 1]", "value = [A1]"), ("value = #1 Jan 2026#", "value = #1Jan2026#")],
)
def test_vba_diff_preserves_bracket_and_date_literal_contents(workbook, source):
    assert _compare("modMain", [workbook], [source], "modMain.bas")["status"] == "modified"


@pytest.mark.unit
def test_vba_normalization_supports_labels_and_colon_statements():
    assert _normalize_vba_line(
        "100 IF Ready THEN: REM Preserve This"
    ) == (
        ("number", "100"),
        ("identifier", "if"),
        ("identifier", "ready"),
        ("identifier", "then"),
        ("punctuation", ":"),
        ("identifier", "rem"),
        ("comment", " Preserve This"),
    )


@pytest.mark.unit
def test_diff_rejects_unknown_comparison_mode():
    with pytest.raises(ValueError, match="comparison must be one of"):
        _compare("modMain", [], [], "modMain.bas", comparison="semantic-ish")
