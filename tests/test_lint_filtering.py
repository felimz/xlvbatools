"""Tests for stable lint selection and baseline behavior."""

import json

import pytest

from xlvbatools.analysis.filtering import (
    lint_issue_fingerprint,
    read_lint_baseline,
    select_lint_issues,
    write_lint_baseline,
)
from xlvbatools.analysis.issue import VBAIssue


def _issue(
    *, rule="IP001", severity="ERROR", module="modules/modMain.bas",
    line=3, message="Variable 'FileCount' is undeclared", procedure="Run",
):
    return VBAIssue(rule, severity, module, line, message, procedure)


@pytest.mark.unit
def test_fingerprint_ignores_line_separator_and_vba_case_changes():
    first = _issue()
    moved = _issue(
        module="MODULES\\MODMAIN.BAS", line=900,
        message="Variable   'filecount' is undeclared", procedure="run",
    )

    assert lint_issue_fingerprint(first) == lint_issue_fingerprint(moved)


@pytest.mark.unit
def test_selection_combines_severity_and_repeatable_rule_filters():
    issues = (
        _issue(rule="IP001", severity="ERROR"),
        _issue(rule="DV001", severity="ERROR"),
        _issue(rule="ST001", severity="STYLE"),
    )

    selection = select_lint_issues(
        issues, severities=["error"], rules=["dv001", "IP001"],
    )

    assert [issue.rule_id for issue in selection.issues] == ["IP001", "DV001"]
    assert selection.metadata()["suppressed_count"] == 1


@pytest.mark.unit
def test_new_only_baseline_uses_multiset_occurrences(tmp_path):
    known = _issue(line=3)
    duplicate = _issue(line=30)
    new = _issue(rule="DV001", message="Duplicate declaration")
    baseline = tmp_path / "lint-baseline.json"
    write_lint_baseline(baseline, [known])

    selection = select_lint_issues(
        [known, duplicate, new], baseline=baseline, new_only=True,
    )

    assert selection.issues == (duplicate, new)
    assert selection.known_count == 1
    assert selection.new_count == 2
    assert selection.baseline_count == 1


@pytest.mark.unit
def test_baseline_is_deterministic_versioned_and_atomic(tmp_path):
    baseline = tmp_path / "nested" / "lint-baseline.json"
    resolved = write_lint_baseline(
        baseline,
        [_issue(rule="ST001", severity="STYLE"), _issue()],
    )

    payload = json.loads(baseline.read_text(encoding="utf-8"))
    assert resolved == str(baseline.resolve())
    assert payload["schema_version"] == "1.0"
    assert len(read_lint_baseline(baseline)) == 2
    assert not list(baseline.parent.glob("*.tmp"))


@pytest.mark.unit
def test_new_only_requires_a_baseline():
    with pytest.raises(ValueError, match="requires a baseline"):
        select_lint_issues([_issue()], new_only=True)


@pytest.mark.unit
def test_unknown_rule_filter_fails_closed():
    with pytest.raises(ValueError, match="Unknown lint rule"):
        select_lint_issues([_issue()], rules=["IP00I"])


@pytest.mark.unit
def test_invalid_baseline_fails_closed(tmp_path):
    baseline = tmp_path / "lint-baseline.json"
    baseline.write_text(
        '{"schema_version":"9.0","findings":[]}', encoding="utf-8",
    )

    with pytest.raises(ValueError, match="Unsupported lint baseline schema"):
        select_lint_issues([_issue()], baseline=baseline, new_only=True)


@pytest.mark.unit
def test_tampered_baseline_fingerprint_fails_closed(tmp_path):
    baseline = tmp_path / "lint-baseline.json"
    write_lint_baseline(baseline, [_issue()])
    payload = json.loads(baseline.read_text(encoding="utf-8"))
    payload["findings"][0]["message"] = "A different finding"
    baseline.write_text(json.dumps(payload), encoding="utf-8")

    with pytest.raises(ValueError, match="fingerprint mismatch"):
        read_lint_baseline(baseline)
