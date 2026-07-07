"""
Tests for xlvbatools.analysis.rules -- VBA static analysis rules.
"""

import pytest
from xlvbatools.analysis.rules import (
    check_dim_after_exec,
    check_line_continuation,
    check_unbalanced_blocks,
    check_msgbox,
    check_implicit_variant,
    check_active_refs,
    check_option_explicit,
    run_all_rules,
)


@pytest.mark.unit
class TestDimAfterExec:
    """DS001: Dim statement after executable code."""

    def test_dim_at_top_passes(self):
        lines = [
            "Public Sub Test()\n",
            "    Dim x As Double\n",
            "    Dim y As Long\n",
            "    x = 42\n",
            "End Sub\n",
        ]
        issues = check_dim_after_exec("test.bas", lines)
        assert len(issues) == 0

    def test_dim_after_exec_fails(self):
        lines = [
            "Public Sub Test()\n",
            "    Dim x As Double\n",
            "    x = 42\n",
            "    Dim y As Long\n",
            "End Sub\n",
        ]
        issues = check_dim_after_exec("test.bas", lines)
        assert len(issues) == 1
        assert issues[0].rule_id == "DS001"
        assert issues[0].severity == "ERROR"
        assert issues[0].line_num == 4

    def test_dim_in_separate_procs_passes(self):
        lines = [
            "Public Sub A()\n",
            "    Dim x As Double\n",
            "    x = 1\n",
            "End Sub\n",
            "Public Sub B()\n",
            "    Dim y As Long\n",
            "    y = 2\n",
            "End Sub\n",
        ]
        issues = check_dim_after_exec("test.bas", lines)
        assert len(issues) == 0


@pytest.mark.unit
class TestLineContinuation:
    """LC001: Orphaned line continuation."""

    def test_correct_continuation_passes(self):
        lines = [
            "    x = a + b _\n",
            "        + c\n",
        ]
        issues = check_line_continuation("test.bas", lines)
        assert len(issues) == 0

    def test_orphaned_continuation_warns(self):
        lines = [
            "    x = SomeFunc_\n",
        ]
        issues = check_line_continuation("test.bas", lines)
        assert len(issues) == 1
        assert issues[0].rule_id == "LC001"
        assert issues[0].severity == "WARNING"


@pytest.mark.unit
class TestUnbalancedBlocks:
    """SB001: Unbalanced Sub/Function blocks."""

    def test_balanced_passes(self):
        lines = [
            "Public Sub Test()\n",
            "    Debug.Print 1\n",
            "End Sub\n",
        ]
        issues = check_unbalanced_blocks("test.bas", lines)
        assert len(issues) == 0

    def test_missing_end_sub_fails(self):
        lines = [
            "Public Sub Test()\n",
            "    Debug.Print 1\n",
        ]
        issues = check_unbalanced_blocks("test.bas", lines)
        assert len(issues) == 1
        assert "no matching End" in issues[0].message

    def test_extra_end_sub_fails(self):
        lines = [
            "End Sub\n",
        ]
        issues = check_unbalanced_blocks("test.bas", lines)
        assert len(issues) == 1
        assert "without matching start" in issues[0].message

    def test_nested_function_balanced(self):
        lines = [
            "Public Sub Outer()\n",
            "End Sub\n",
            "Public Function Inner() As Long\n",
            "    Inner = 42\n",
            "End Function\n",
        ]
        issues = check_unbalanced_blocks("test.bas", lines)
        assert len(issues) == 0


@pytest.mark.unit
class TestMsgBox:
    """PF001: MsgBox detection."""

    def test_msgbox_warns(self):
        lines = [
            "Public Sub Test()\n",
            '    MsgBox "Hello"\n',
            "End Sub\n",
        ]
        issues = check_msgbox("test.bas", lines)
        assert len(issues) == 1
        assert issues[0].rule_id == "PF001"

    def test_commented_msgbox_passes(self):
        lines = [
            "'    MsgBox \"Hello\"\n",
        ]
        issues = check_msgbox("test.bas", lines)
        assert len(issues) == 0


@pytest.mark.unit
class TestImplicitVariant:
    """PF002: Dim without explicit type."""

    def test_explicit_type_passes(self):
        lines = ["    Dim x As Double\n"]
        issues = check_implicit_variant("test.bas", lines)
        assert len(issues) == 0

    def test_implicit_variant_warns(self):
        lines = ["    Dim x\n"]
        issues = check_implicit_variant("test.bas", lines)
        assert len(issues) == 1
        assert issues[0].rule_id == "PF002"


@pytest.mark.unit
class TestActiveRefs:
    """PF003: ActiveSheet/ActiveCell usage."""

    def test_activesheet_warns(self):
        lines = ["    ActiveSheet.Range(\"A1\").Value = 1\n"]
        issues = check_active_refs("test.bas", lines)
        assert len(issues) == 1
        assert "ActiveSheet" in issues[0].message

    def test_explicit_ref_passes(self):
        lines = ["    ws.Range(\"A1\").Value = 1\n"]
        issues = check_active_refs("test.bas", lines)
        assert len(issues) == 0


@pytest.mark.unit
class TestOptionExplicit:
    """OE001: Missing Option Explicit."""

    def test_with_option_explicit_passes(self):
        lines = [
            "Option Explicit\n",
            "Dim x As Long\n",
        ]
        issues = check_option_explicit("test.bas", lines)
        assert len(issues) == 0

    def test_missing_option_explicit_warns(self):
        lines = [
            "Dim x As Long\n",
            "Public Sub Test()\n",
            "End Sub\n",
        ]
        issues = check_option_explicit("test.bas", lines)
        assert len(issues) == 1
        assert issues[0].rule_id == "OE001"


@pytest.mark.unit
class TestRunAllRules:
    """Test the rule orchestrator."""

    def test_all_rules_run(self):
        lines = [
            "Dim x\n",
            "Public Sub Test()\n",
            '    MsgBox "hi"\n',
            "    ActiveSheet.Range(\"A1\").Value = 1\n",
            "End Sub\n",
        ]
        issues = run_all_rules("test.bas", lines)
        rule_ids = {i.rule_id for i in issues}
        assert "OE001" in rule_ids  # Missing Option Explicit
        assert "PF001" in rule_ids  # MsgBox
        assert "PF002" in rule_ids  # Implicit Variant
        assert "PF003" in rule_ids  # ActiveSheet

    def test_disabled_rules_skipped(self):
        lines = [
            "Dim x\n",
            "Public Sub Test()\n",
            '    MsgBox "hi"\n',
            "End Sub\n",
        ]
        issues = run_all_rules("test.bas", lines, disabled_rules=["PF001", "PF002"])
        rule_ids = {i.rule_id for i in issues}
        assert "PF001" not in rule_ids
        assert "PF002" not in rule_ids
