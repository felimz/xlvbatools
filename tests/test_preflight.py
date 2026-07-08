"""
Tests for xlvbatools.analysis.rules -- VBA static analysis rules.
"""

import pytest
from xlvbatools.analysis.rules import (
    check_dim_after_exec,
    check_const_after_exec,
    check_line_continuation,
    check_unbalanced_blocks,
    check_msgbox,
    check_implicit_variant,
    check_active_refs,
    check_option_explicit,
    check_undeclared_variables,
    check_block_declarations,
    check_one_declaration_per_line,
    check_avoid_integer,
    check_explicit_param_passing,
    check_explicit_access_modifiers,
    check_line_length,
    check_no_call_keyword,
    check_hardcoded_secrets,
    check_absolute_paths,
    check_type_suffixes,
    check_fragile_selection,
    check_silent_error_suppression,
    check_error_handler,
    check_filedialog_guard,
    check_unused_local_variables,
    check_empty_procedures,
    check_consecutive_blank_lines,
    check_double_spacing,
    run_all_rules,
    _is_entry_point,
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
        assert issues[0].severity == "WARNING"
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
        assert "lacks a corresponding 'End Sub' or 'End Function'" in issues[0].message

    def test_extra_end_sub_fails(self):
        lines = [
            "End Sub\n",
        ]
        issues = check_unbalanced_blocks("test.bas", lines)
        assert len(issues) == 1
        assert "without a matching start statement" in issues[0].message

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
        assert issues[0].severity == "ERROR"


@pytest.mark.unit
class TestBlockDeclarations:
    """BK001: Local variable declaration inside control-flow blocks."""

    def test_block_declarations_warns(self):
        lines = [
            "Public Sub Test()\n",
            "    If True Then\n",
            "        Dim x As Long\n",
            "    End If\n",
            "End Sub\n",
        ]
        issues = check_block_declarations("test.bas", lines)
        assert len(issues) == 1
        assert issues[0].rule_id == "BK001"
        assert issues[0].severity == "WARNING"
        assert "inside a 'If' block" in issues[0].message

    def test_block_declarations_passes(self):
        lines = [
            "Public Sub Test()\n",
            "    Dim x As Long\n",
            "    If True Then\n",
            "        x = 5\n",
            "    End If\n",
            "End Sub\n",
        ]
        issues = check_block_declarations("test.bas", lines)
        assert len(issues) == 0


@pytest.mark.unit
class TestLinterStyleGuideRules:
    """Test cases for the newly introduced style guide rules."""

    def test_one_declaration_per_line(self):
        # Commas inside dim are forbidden
        lines = ["    Dim a, b As Long\n"]
        issues = check_one_declaration_per_line("test.bas", lines)
        assert len(issues) == 1
        assert issues[0].rule_id == "SD002"
        assert issues[0].severity == "STYLE"

        # Safe for single var or array parens
        lines = ["    Dim arr(1 to 5, 1 to 2) As Long\n"]
        issues = check_one_declaration_per_line("test.bas", lines)
        assert len(issues) == 0

    def test_avoid_integer(self):
        lines = ["    Dim count As Integer\n"]
        issues = check_avoid_integer("test.bas", lines)
        assert len(issues) == 1
        assert issues[0].rule_id == "PF004"
        assert issues[0].severity == "WARNING"

        lines = ["    Dim count As Long\n"]
        issues = check_avoid_integer("test.bas", lines)
        assert len(issues) == 0

    def test_explicit_param_passing(self):
        lines = ["Public Sub Run(x As Long, ByVal y As Double)\n"]
        issues = check_explicit_param_passing("test.bas", lines)
        assert len(issues) == 1
        assert issues[0].rule_id == "SD005"
        assert "Argument 'x' does not specify ByVal or ByRef" in issues[0].message

        # Test Optional parameters
        lines = ["Public Sub Run(Optional cnt As Long, ParamArray args() As Variant)\n"]
        issues = check_explicit_param_passing("test.bas", lines)
        assert len(issues) == 2
        assert "Argument 'cnt' does not specify" in issues[0].message
        assert "Argument 'args()' does not specify" in issues[1].message

        lines = ["Public Sub Run(ByRef x As Long, ByVal y As Double)\n"]
        issues = check_explicit_param_passing("test.bas", lines)
        assert len(issues) == 0

    def test_explicit_access_modifiers(self):
        # Procedure missing modifier
        lines = [
            "Sub Calculate()\n",
            "End Sub\n",
        ]
        issues = check_explicit_access_modifiers("test.bas", lines)
        assert len(issues) == 1
        assert issues[0].rule_id == "SD006"

        # Module-level variable missing modifier (bare Dim)
        lines = [
            "Dim moduleVar As Long\n",
            "Public Sub Test()\n",
            "End Sub\n",
        ]
        issues = check_explicit_access_modifiers("test.bas", lines)
        assert len(issues) == 1
        assert "Module-level variable declaration lacks" in issues[0].message

    def test_line_length(self):
        lines = ["    " + "x" * 120 + "\n"]
        issues = check_line_length("test.bas", lines)
        assert len(issues) == 1
        assert issues[0].rule_id == "SD010"

    def test_no_call_keyword(self):
        lines = ["    Call SaveFile(filepath)\n"]
        issues = check_no_call_keyword("test.bas", lines)
        assert len(issues) == 1
        assert issues[0].rule_id == "SD014"


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


@pytest.mark.unit
class TestAdvancedLinterRules:
    """Test cases for the newly introduced advanced security and dead code rules."""

    def test_hardcoded_secrets(self):
        lines = [
            "Dim pass As String\n",
            'pass = "secret123"\n',
            'apiKey = "AB125"\n',
            '\' Hardcoded password in comment should be ignored: pass = "secret123"\n',
            'x = 1 \' inline comment: pwd = "xyz"\n',
        ]
        issues = check_hardcoded_secrets("test.bas", lines)
        assert len(issues) == 1
        assert issues[0].rule_id == "SC001"
        assert issues[0].severity == "WARNING"

    def test_absolute_paths(self):
        lines = [
            'Dim path As String\n',
            'path = "C:\\Users\\felim\\Documents\\file.txt"\n',
            '\' Commented path: "C:\\Users\\felim\\Documents\\file.txt"\n',
            'y = 2 \' inline: "C:\\Users\\felim\\file.txt"\n',
        ]
        issues = check_absolute_paths("test.bas", lines)
        assert len(issues) == 1
        assert issues[0].rule_id == "SC002"
        assert issues[0].severity == "WARNING"

    def test_type_suffixes(self):
        lines = [
            "Dim count%\n",
            "Dim name$\n",
            "' Comment: Dim test%\n",
            "Dim x As String ' inline suffix Dim y%\n",
        ]
        issues = check_type_suffixes("test.bas", lines)
        assert len(issues) == 2
        assert issues[0].rule_id == "CL001"
        assert "Obsolete type declaration suffix" in issues[0].message

    def test_fragile_selection(self):
        lines = [
            "    Range(\"A1\").Select\n",
            "    Selection.Value = 5\n",
            "    ' Commented: Range(\"A1\").Select\n",
            "    x = 1 ' inline: Range(\"A1\").Select\n",
        ]
        issues = check_fragile_selection("test.bas", lines)
        assert len(issues) == 2
        assert all(i.rule_id == "CL002" for i in issues)

    def test_silent_error_suppression(self):
        # Unverified (no Err check or reset in 5 lines)
        lines1 = [
            "On Error Resume Next\n",
            "x = 1 / 0\n",
            "y = 2\n",
        ]
        assert len(check_silent_error_suppression("test.bas", lines1)) == 1

        # Verified by GoTo 0
        lines2 = [
            "On Error Resume Next\n",
            "x = 1 / 0\n",
            "On Error GoTo 0\n",
        ]
        assert len(check_silent_error_suppression("test.bas", lines2)) == 0

        # Verified by checking Err.Number
        lines3 = [
            "On Error Resume Next\n",
            "x = 1 / 0\n",
            "If Err.Number <> 0 Then\n",
        ]
        assert len(check_silent_error_suppression("test.bas", lines3)) == 0

    def test_unused_local_variables(self):
        lines = [
            "Public Sub Run()\n",
            "    Dim x As Long\n",
            "    Dim y As String\n",
            '    y = "test"\n',
            "End Sub\n",
        ]
        issues = check_unused_local_variables("test.bas", lines)
        assert len(issues) == 1
        assert issues[0].rule_id == "DC001"
        assert "Unused local variable 'x'" in issues[0].message

    def test_empty_procedures(self):
        lines = [
            "Public Sub EmptySub()\n",
            "End Sub\n",
            "Public Sub NonEmptySub()\n",
            "    Dim x As Long\n",
            "    x = 1\n",
            "End Sub\n",
        ]
        issues = check_empty_procedures("test.bas", lines)
        assert len(issues) == 1
        assert issues[0].rule_id == "DC002"
        assert "Empty procedure 'EmptySub'" in issues[0].message

    def test_dead_procedures(self, tmp_path):
        from xlvbatools.analysis.preflight import lint_files
        
        modules_dir = tmp_path / "modules"
        classes_dir = tmp_path / "classes"
        modules_dir.mkdir()
        classes_dir.mkdir()
        
        mod_content = (
            "Attribute VB_Name = \"modTest\"\n"
            "Option Explicit\n"
            "Public Sub MainEntryPoint()\n"
            "    Call HelperUsed\n"
            "End Sub\n"
            "Private Sub HelperUsed()\n"
            "    Dim x As Long\n"
            "    x = 1\n"
            "End Sub\n"
            "Private Sub UnusedHelper()\n"
            "    Dim y As Long\n"
            "    y = 2\n"
            "End Sub\n"
        )
        with open(modules_dir / "modTest.bas", "w", encoding="utf-8") as f:
            f.write(mod_content)
            
        issues = lint_files(str(tmp_path))
        rule_ids = {i.rule_id for i in issues}
        assert "DC003" in rule_ids
        
        dead_procs = [i.procedure for i in issues if i.rule_id == "DC003"]
        assert "UnusedHelper" in dead_procs
        assert "HelperUsed" not in dead_procs
        assert "MainEntryPoint" not in dead_procs

    def test_consecutive_blank_lines(self):
        lines = [
            "Option Explicit\n",
            "\n",
            "\n",  # Line 3 (second consecutive blank line)
            "Public Sub Test()\n",
            "End Sub\n",
        ]
        issues = check_consecutive_blank_lines("test.bas", lines)
        assert len(issues) == 1
        assert issues[0].rule_id == "SD015"
        assert issues[0].line_num == 3

    def test_double_spacing(self):
        lines = [
            "Public Sub Test()\n",
            "    x = 1\n",
            "\n",  # line 3
            "    y = 2\n",
            "\n",  # line 5
            "    z = 3\n",
            "\n",  # line 7
            "    End Sub\n",
        ]
        issues = check_double_spacing("test.bas", lines)
        assert len(issues) == 1
        assert issues[0].rule_id == "SD016"
        assert issues[0].line_num == 3


@pytest.mark.unit
class TestConstAfterExec:
    """CS001: Const after executable statement."""

    def test_const_at_top_passes(self):
        lines = [
            "Public Sub Test()\n",
            "    Const MAX_VAL As Long = 100\n",
            "    Dim x As Long\n",
            "    x = MAX_VAL\n",
            "End Sub\n",
        ]
        issues = check_const_after_exec("test.bas", lines)
        assert len(issues) == 0

    def test_const_after_exec_warns(self):
        lines = [
            "Public Sub Test()\n",
            "    Dim x As Long\n",
            "    x = 1\n",
            "    Const MAX_VAL As Long = 100\n",
            "End Sub\n",
        ]
        issues = check_const_after_exec("test.bas", lines)
        assert len(issues) == 1
        assert issues[0].rule_id == "CS001"


@pytest.mark.unit
class TestUndeclaredVariables:
    """UV001: Undeclared variable usage."""

    def test_all_declared_passes(self):
        lines = [
            "Option Explicit\n",
            "Public Sub Test()\n",
            "    Dim x As Long\n",
            "    x = 1\n",
            "End Sub\n",
        ]
        issues = check_undeclared_variables("test.bas", lines)
        assert len(issues) == 0

    def test_undeclared_variable_errors(self):
        lines = [
            "Option Explicit\n",
            "Public Sub Test()\n",
            "    Dim x As Long\n",
            "    x = 1\n",
            "    filePath = \"test\"\n",
            "End Sub\n",
        ]
        issues = check_undeclared_variables("test.bas", lines)
        assert len(issues) == 1
        assert issues[0].rule_id == "UV001"
        assert "filePath" in issues[0].message

    def test_no_option_explicit_skips(self):
        lines = [
            "Public Sub Test()\n",
            "    filePath = \"test\"\n",
            "End Sub\n",
        ]
        issues = check_undeclared_variables("test.bas", lines)
        assert len(issues) == 0

    def test_for_loop_variable_detected(self):
        lines = [
            "Option Explicit\n",
            "Public Sub Test()\n",
            "    For ii = 1 To 10\n",
            "    Next ii\n",
            "End Sub\n",
        ]
        issues = check_undeclared_variables("test.bas", lines)
        assert len(issues) == 1
        assert "ii" in issues[0].message

    def test_parameter_not_flagged(self):
        lines = [
            "Option Explicit\n",
            "Public Sub Test(ByVal x As Long)\n",
            "    x = x + 1\n",
            "End Sub\n",
        ]
        issues = check_undeclared_variables("test.bas", lines)
        assert len(issues) == 0

    def test_function_return_value_not_flagged(self):
        """Assigning to the function name is the VBA return idiom, not an undeclared var."""
        lines = [
            "Option Explicit\n",
            "Public Function SafeDbl(ByVal val As Variant) As Double\n",
            "    SafeDbl = CDbl(val)\n",
            "End Function\n",
        ]
        issues = check_undeclared_variables("test.bas", lines)
        assert len(issues) == 0

    def test_property_get_return_not_flagged(self):
        """Property Get also uses name = value for return."""
        lines = [
            "Option Explicit\n",
            "Public Property Get Name() As String\n",
            "    Name = m_name\n",
            "End Property\n",
        ]
        issues = check_undeclared_variables("test.bas", lines)
        # m_name might be flagged (module-level), but Name should not
        flagged_names = [i.message for i in issues]
        assert not any("'Name'" in m for m in flagged_names)


@pytest.mark.unit
class TestErrorHandler:
    """EH001: Missing error handler."""

    def test_public_with_handler_passes(self):
        lines = [
            "Public Sub Test()\n",
            "    On Error GoTo ErrHandler\n",
            "    Dim x As Long\n",
            "    x = 1\n",
            "    Exit Sub\n",
            "ErrHandler:\n",
            "    Err.Raise Err.Number\n",
            "End Sub\n",
        ]
        issues = check_error_handler("test.bas", lines)
        assert len(issues) == 0

    def test_public_without_handler_warns(self):
        lines = [
            "Public Sub Test()\n",
            "    Dim x As Long\n",
            "    x = 1\n",
            "End Sub\n",
        ]
        issues = check_error_handler("test.bas", lines)
        assert len(issues) == 1
        assert issues[0].rule_id == "EH001"

    def test_private_without_handler_passes(self):
        lines = [
            "Private Sub Helper()\n",
            "    Dim x As Long\n",
            "    x = 1\n",
            "End Sub\n",
        ]
        issues = check_error_handler("test.bas", lines)
        assert len(issues) == 0


@pytest.mark.unit
class TestFileDialogGuard:
    """FD001: FileDialog without UserControl guard."""

    def test_guarded_filedialog_passes(self):
        lines = [
            "Private Sub Test()\n",
            "    If Not Application.UserControl Then Exit Sub\n",
            "    Application.FileDialog(msoFileDialogOpen).Show\n",
            "End Sub\n",
        ]
        issues = check_filedialog_guard("test.bas", lines)
        assert len(issues) == 0

    def test_unguarded_filedialog_warns(self):
        lines = [
            "Private Sub Test()\n",
            "    Application.FileDialog(msoFileDialogOpen).Show\n",
            "End Sub\n",
        ]
        issues = check_filedialog_guard("test.bas", lines)
        assert len(issues) == 1
        assert issues[0].rule_id == "FD001"


@pytest.mark.unit
class TestImprovedOptionExplicit:
    """OE001 improvements: no Dim required, impact assessment."""

    def test_module_with_assignments_but_no_dim(self):
        """OE001 should fire even without Dim statements."""
        lines = [
            "Public Sub Test()\n",
            "    filePath = \"test\"\n",
            "End Sub\n",
        ]
        issues = check_option_explicit("test.bas", lines)
        assert len(issues) == 1
        assert issues[0].rule_id == "OE001"

    def test_impact_assessment_in_message(self):
        """OE001 should list undeclared variables in IMPACT note."""
        lines = [
            "Public Sub Test()\n",
            "    Dim x As Long\n",
            "    x = 1\n",
            "    filePath = \"test\"\n",
            "End Sub\n",
        ]
        issues = check_option_explicit("test.bas", lines)
        assert len(issues) == 1
        assert "IMPACT" in issues[0].message
        assert "filePath" in issues[0].message


@pytest.mark.unit
class TestImprovedImplicitVariant:
    """PF002 improvements: now catches Const without type."""

    def test_const_without_type_warns(self):
        lines = [
            "Public Sub Test()\n",
            "    Const ForReading = 1\n",
            "End Sub\n",
        ]
        issues = check_implicit_variant("test.bas", lines)
        assert len(issues) == 1
        assert issues[0].rule_id == "PF002"
        assert "ForReading" in issues[0].message

    def test_const_with_type_passes(self):
        lines = [
            "Public Sub Test()\n",
            "    Const ForReading As Long = 1\n",
            "End Sub\n",
        ]
        issues = check_implicit_variant("test.bas", lines)
        assert len(issues) == 0


@pytest.mark.unit
class TestEntryPointWhitelist:
    """DC003 entry-point detection."""

    def test_event_handlers_recognized(self):
        assert _is_entry_point("Worksheet_Change") is True
        assert _is_entry_point("Workbook_Open") is True
        assert _is_entry_point("CommandButton1_Click") is True
        assert _is_entry_point("Worksheet_SelectionChange") is True

    def test_regular_procs_not_entry_points(self):
        assert _is_entry_point("CalculateWeight") is False
        assert _is_entry_point("ParseLine") is False

    def test_named_entry_points_recognized(self):
        assert _is_entry_point("OnRetrieve") is True
        assert _is_entry_point("Main") is True
