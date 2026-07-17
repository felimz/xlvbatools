"""
Unit tests for the xlvba command line interface.
"""

import os
import json
import pytest
from unittest.mock import patch
from xlvbatools.cli.main import main
from xlvbatools.results import AttemptDiagnostic, Diagnostics, OperationResult


def _result(operation, data=None, *, success=True):
    return OperationResult(
        operation=operation,
        success=success,
        phase="complete",
        data=data,
        elapsed_seconds=0.1,
    )


@pytest.mark.unit
def test_cli_version(capsys):
    """Even the short version flag is machine-readable by default."""
    with pytest.raises(SystemExit) as exc_info:
        main(["--version"])
    assert exc_info.value.code == 0
    captured = capsys.readouterr()
    output = captured.out or captured.err
    assert json.loads(output)["version"]


@pytest.mark.unit
def test_cli_detailed_version_is_structured_json(capsys):
    main(["version"])

    payload = json.loads(capsys.readouterr().out)
    assert payload["operation"] == "version"
    assert payload["success"] is True
    assert payload["data"]["version"]
    assert os.path.basename(payload["data"]["python_executable"]).lower().startswith("python")
    assert payload["data"]["package_path"].endswith("xlvbatools")


@pytest.mark.unit
def test_human_and_table_presentations_are_explicit_opt_ins(capsys):
    main(["version", "--text"])
    assert capsys.readouterr().out.startswith("xlvba ")

    main(["version", "--table"])
    output = capsys.readouterr().out
    assert "field" in output
    assert "version" in output


@pytest.mark.unit
def test_unexpected_command_failure_is_a_json_envelope(tmp_path, capsys):
    fake_project = patch("xlvbatools.cli.commands._project").start().return_value
    fake_project.run.side_effect = RuntimeError("unexpected worker boundary failure")
    with patch("xlvbatools.config.loader.load_config") as mock_config, \
         patch("xlvbatools.logging.setup_logging"):
        mock_config.return_value.workbook = "book.xlsm"
        mock_config.return_value.log_dir = str(tmp_path)
        mock_config.return_value.log_name = "test"
        with pytest.raises(SystemExit) as exc_info:
            main(["run", "MyMacro"])

    assert exc_info.value.code == 1
    payload = json.loads(capsys.readouterr().out)
    assert payload["success"] is False
    assert payload["error"]["code"] == "unhandled_cli_error"
    assert "unexpected worker boundary failure" in payload["error"]["message"]
    patch.stopall()


@pytest.mark.unit
def test_cli_help(capsys):
    """Test running xlvba with no args displays help."""
    with pytest.raises(SystemExit) as exc_info:
        main([])
    assert exc_info.value.code == 0
    captured = capsys.readouterr()
    output = captured.out or captured.err
    assert "usage:" in output.lower()
    assert "xlvba help" in output
    assert "xlvba agents install" in output
    assert ".agents/" in output
    assert "OperationResult JSON envelope" in output


@pytest.mark.unit
def test_every_public_command_has_conventional_help(capsys):
    """All discovery-catalog commands expose useful argparse help."""
    from xlvbatools.cli.discovery import COMMAND_SPECS

    for spec in COMMAND_SPECS:
        with pytest.raises(SystemExit) as exc_info:
            main([spec.name, "--help"])
        assert exc_info.value.code == 0
        output = capsys.readouterr().out
        assert "usage:" in output.lower(), spec.name
        assert spec.summary in output, spec.name

    for command in ("create", "list", "info", "restore", "diff", "prune"):
        with pytest.raises(SystemExit) as exc_info:
            main(["snapshot", command, "--help"])
        assert exc_info.value.code == 0
        assert "usage:" in capsys.readouterr().out.lower()

    with pytest.raises(SystemExit) as exc_info:
        main(["agents", "install", "--help"])
    assert exc_info.value.code == 0
    install_help = capsys.readouterr().out
    assert "--destination" in install_help
    assert "--force" in install_help


@pytest.mark.unit
def test_machine_readable_help_catalog_supports_command_detail(capsys):
    main(["help"])
    payload = json.loads(capsys.readouterr().out)
    assert payload["operation"] == "help"
    assert payload["data"]["discovery_schema_version"] == "1.0"
    assert payload["data"]["output_contract"]["default"] == "json"
    assert payload["data"]["agent_templates"]["destination"] == ".agents/"
    assert any(item["name"] == "extract" for item in payload["data"]["commands"])

    main(["help", "extract"])
    detail = json.loads(capsys.readouterr().out)["data"]["command"]
    assert detail["name"] == "extract"
    assert detail["excel_required"] is True
    assert detail["examples"]
    options = {item["name"]: item for item in detail["options"]}
    assert options["workbook"]["flags"] == ["--workbook", "-w"]
    assert options["timeout"]["default"] == 120.0

    main(["help", "run"])
    run_detail = json.loads(capsys.readouterr().out)["data"]["command"]
    run_options = {item["name"]: item for item in run_detail["options"]}
    assert run_options["named_range"]["flags"] == ["--named-range"]
    assert run_options["save"]["flags"] == ["--save", "--no-save"]
    assert run_options["save"]["default"] is True
    assert run_options["visible"]["default"] is False

    main(["help", "agents"])
    agents = json.loads(capsys.readouterr().out)["data"]["command"]
    install = next(item for item in agents["subcommands"] if item["name"] == "install")
    assert any(option["name"] == "destination" for option in install["options"])


@pytest.mark.unit
def test_cli_search(temp_vba_source, sample_bas_file, tmp_path, capsys):
    """Test xlvba search subcommand."""
    # Test literal search
    with patch("xlvbatools.config.loader.load_config") as mock_config:
        mock_cfg = mock_config.return_value
        mock_cfg.vba_source = str(temp_vba_source)
        mock_cfg.log_dir = str(tmp_path)
        mock_cfg.log_name = "test_search"
        main(["search", "TestSub"])
    
    payload = json.loads(capsys.readouterr().out)
    assert payload["operation"] == "search"
    assert payload["metadata"]["match_count"] == 1
    assert payload["data"][0]["line"] == "Public Sub TestSub()"

    # Test regex search
    with patch("xlvbatools.config.loader.load_config") as mock_config:
        mock_cfg = mock_config.return_value
        mock_cfg.vba_source = str(temp_vba_source)
        mock_cfg.log_dir = str(tmp_path)
        mock_cfg.log_name = "test_search"
        main(["search", "--regex", r"Dim\s+\w+\s+As"])

    payload = json.loads(capsys.readouterr().out)
    assert payload["data"][0]["line"].strip() == "Dim x As Double"


@pytest.mark.unit
def test_cli_fmt(temp_vba_source, sample_bas_file, tmp_path, capsys):
    """Test xlvba fmt subcommand."""
    # Test formatting with dry-run
    with patch("xlvbatools.config.loader.load_config") as mock_config:
        mock_cfg = mock_config.return_value
        mock_cfg.vba_source = str(temp_vba_source)
        mock_cfg.log_dir = str(tmp_path)
        mock_cfg.log_name = "test_fmt"
        main(["fmt", "--dry-run"])
    
    payload = json.loads(capsys.readouterr().out)
    assert payload["operation"] == "format"
    assert payload["metadata"]["dry_run"] is True
    assert payload["metadata"]["file_count"] >= 1


@pytest.mark.unit
def test_cli_graph(temp_vba_source, sample_bas_file, tmp_path, capsys):
    """Test xlvba graph subcommand."""
    # Default graph output is structured JSON.
    with patch("xlvbatools.config.loader.load_config") as mock_config:
        mock_cfg = mock_config.return_value
        mock_cfg.vba_source = str(temp_vba_source)
        mock_cfg.log_dir = str(tmp_path)
        mock_cfg.log_name = "test_graph"
        main(["graph"])

    payload = json.loads(capsys.readouterr().out)
    assert payload["operation"] == "graph"
    assert payload["data"]["graph_format"] == "json"
    assert payload["data"]["graph"]["node_count"] >= 1

    with patch("xlvbatools.config.loader.load_config") as mock_config:
        mock_config.return_value.vba_source = str(temp_vba_source)
        mock_config.return_value.log_dir = str(tmp_path)
        mock_config.return_value.log_name = "test_graph"
        main(["graph", "--graph-format", "mermaid", "--text"])
    assert "graph TD" in capsys.readouterr().out


@pytest.mark.unit
def test_cli_init(tmp_path, monkeypatch, capsys):
    """Test xlvba init subcommand."""
    monkeypatch.chdir(tmp_path)
    
    # 1. First init
    main(["init", "--workbook", "my_test.xlsm"])
    payload = json.loads(capsys.readouterr().out)
    assert payload["success"] is True
    assert payload["data"]["config_path"].endswith("xlvbatools.toml")
    assert os.path.exists("xlvbatools.toml")
    
    # Check config content
    with open("xlvbatools.toml", "r") as f:
        content = f.read()
    assert 'workbook = "my_test.xlsm"' in content

    # 2. Re-init without force (should fail)
    with pytest.raises(SystemExit) as exc_info:
        main(["init", "--workbook", "new.xlsm"])
    assert exc_info.value.code == 1
    payload = json.loads(capsys.readouterr().out)
    assert payload["success"] is False
    assert "already exists" in payload["error"]["message"]

    # 3. Re-init with force (should succeed)
    main(["init", "--workbook", "new.xlsm", "--force"])
    assert json.loads(capsys.readouterr().out)["success"] is True
    with open("xlvbatools.toml", "r") as f:
        content = f.read()
    assert 'workbook = "new.xlsm"' in content

    # 4. Re-init with force and --agents (should install templates)
    main(["init", "--workbook", "new.xlsm", "--force", "--agents"])
    payload = json.loads(capsys.readouterr().out)
    assert payload["data"]["agents_status"] == "installed"
    assert os.path.exists(".agents")
    assert os.path.exists(".agents/AGENTS.md")
    assert os.path.exists(".agents/skills/xlvba-toolchain/SKILL.md")
    assert os.path.exists(".agents/workflows/get-started.md")
    assert os.path.exists(".agents/workflows/vba-debug.md")
    assert os.path.exists(".agents/workflows/vba-edit.md")
    assert os.path.exists(".agents/rules/vba-rules.md")
    assert os.path.exists(".agents/rules/python-rules.md")


@pytest.mark.unit
def test_global_agents_help_exits_successfully(capsys):
    """The global --agents flag remains distinct from init --agents."""
    with pytest.raises(SystemExit) as exc_info:
        main(["--agents"])
    assert exc_info.value.code == 0
    payload = json.loads(capsys.readouterr().out)
    assert payload["operation"] == "agents"
    assert payload["data"]["default_output"] == "json"


@pytest.mark.unit
def test_agents_install_is_incremental_and_force_is_scoped(tmp_path, monkeypatch, capsys):
    monkeypatch.chdir(tmp_path)

    main(["agents", "install"])
    first = json.loads(capsys.readouterr().out)
    assert first["operation"] == "agents_install"
    assert first["data"]["destination"] == str((tmp_path / ".agents").resolve())
    assert "AGENTS.md" in first["data"]["installed"]
    assert first["data"]["skipped"] == []

    agents_file = tmp_path / ".agents/AGENTS.md"
    agents_file.write_text("project customization", encoding="utf-8")
    custom_file = tmp_path / ".agents/project-only.md"
    custom_file.write_text("keep me", encoding="utf-8")

    main(["agents", "install"])
    second = json.loads(capsys.readouterr().out)
    assert "AGENTS.md" in second["data"]["skipped"]
    assert agents_file.read_text(encoding="utf-8") == "project customization"

    main(["agents", "install", "--force"])
    forced = json.loads(capsys.readouterr().out)
    assert "AGENTS.md" in forced["data"]["overwritten"]
    assert "xlvbatools v1 - Agent Guide" in agents_file.read_text(encoding="utf-8")
    assert custom_file.read_text(encoding="utf-8") == "keep me"


@pytest.mark.unit
def test_cli_run_forwards_timeout(tmp_path):
    project = patch("xlvbatools.cli.commands._project").start().return_value
    project.run.return_value = _result("run_macro", {"macro": "MyMacro"})
    with patch("xlvbatools.config.loader.load_config") as mock_config, \
         patch("xlvbatools.logging.setup_logging"):
        mock_config.return_value.workbook = "book.xlsm"
        mock_config.return_value.log_dir = str(tmp_path)
        mock_config.return_value.log_name = "test"

        with pytest.raises(SystemExit) as exc_info:
            main(["run", "MyMacro", "--timeout", "7.5"])

    assert exc_info.value.code == 0
    project.run.assert_called_once_with(
        "MyMacro",
        named_ranges=None,
        timeout=7.5,
        visible=False,
        save=True,
    )
    patch.stopall()


@pytest.mark.unit
def test_cli_run_emits_executor_attempt_diagnostics(tmp_path, capsys):
    project = patch("xlvbatools.cli.commands._project").start().return_value
    project.run.return_value = OperationResult(
        operation="run_macro",
        success=True,
        phase="complete",
        data={"macro": "MyMacro"},
        attempt_count=2,
        diagnostics=Diagnostics(attempts=(
            AttemptDiagnostic(
                attempt=1,
                phase="worker_start",
                error_code="worker_start_failed",
                retryable=True,
                retry_reason="worker_creation_failed",
            ),
            AttemptDiagnostic(attempt=2, phase="complete"),
        )),
    )
    with patch("xlvbatools.config.loader.load_config") as mock_config, \
         patch("xlvbatools.logging.setup_logging"):
        mock_config.return_value.workbook = "book.xlsm"
        mock_config.return_value.log_dir = str(tmp_path)
        mock_config.return_value.log_name = "test"

        with pytest.raises(SystemExit) as exc_info:
            main(["run", "MyMacro"])

    payload = json.loads(capsys.readouterr().out)
    assert exc_info.value.code == 0
    assert payload["attempt_count"] == 2
    assert payload["diagnostics"]["attempts"][0]["retry_reason"] == (
        "worker_creation_failed"
    )
    patch.stopall()


@pytest.mark.unit
def test_cli_run_forwards_typed_named_ranges_save_and_visibility(tmp_path):
    project = patch("xlvbatools.cli.commands._project").start().return_value
    project.run.return_value = _result("run_macro", {"macro": "MyMacro"})
    with patch("xlvbatools.config.loader.load_config") as mock_config, \
         patch("xlvbatools.logging.setup_logging"):
        mock_config.return_value.workbook = "book.xlsm"
        mock_config.return_value.log_dir = str(tmp_path)
        mock_config.return_value.log_name = "test"

        with pytest.raises(SystemExit) as exc_info:
            main([
                "run", "MyMacro",
                "--named-range", "Count=42",
                "--named-range", "Ratio=0.707",
                "--named-range", "Enabled=true",
                "--named-range", 'Mode="Design"',
                "--named-range", "Label=North Sea",
                "--named-range", "Clear=null",
                "--named-range", "Empty=",
                "--no-save", "--visible", "--timeout", "8",
            ])

    assert exc_info.value.code == 0
    project.run.assert_called_once_with(
        "MyMacro",
        named_ranges={
            "Count": 42,
            "Ratio": 0.707,
            "Enabled": True,
            "Mode": "Design",
            "Label": "North Sea",
            "Clear": None,
            "Empty": "",
        },
        timeout=8.0,
        visible=True,
        save=False,
    )
    patch.stopall()


@pytest.mark.unit
@pytest.mark.parametrize(
    "options",
    (
        ["--named-range", "MissingEquals"],
        ["--named-range", "=42"],
        ["--named-range", "Input=1", "--named-range", "input=2"],
    ),
)
def test_cli_run_rejects_invalid_named_range_inputs(options, tmp_path, capsys):
    project = patch("xlvbatools.cli.commands._project").start().return_value
    with patch("xlvbatools.config.loader.load_config") as mock_config, \
         patch("xlvbatools.logging.setup_logging"):
        mock_config.return_value.workbook = "book.xlsm"
        mock_config.return_value.log_dir = str(tmp_path)
        mock_config.return_value.log_name = "test"

        with pytest.raises(SystemExit) as exc_info:
            main(["run", "MyMacro", *options])

    assert exc_info.value.code == 2
    payload = json.loads(capsys.readouterr().out)
    assert payload["error"]["code"] == "invalid_arguments"
    project.run.assert_not_called()
    patch.stopall()


@pytest.mark.com
@pytest.mark.integration
def test_cli_run_named_range_and_no_save_reach_live_worker(
    runtime_error_workbook, tmp_path, capsys,
):
    import xml.etree.ElementTree as ET
    import zipfile

    from xlvbatools.config.schema import XlvbaConfig

    config = XlvbaConfig(
        workbook=runtime_error_workbook,
        vba_source=str(tmp_path / "vba_source"),
        log_dir=str(tmp_path / "logs"),
        log_name="cli_live_run",
    )
    with patch("xlvbatools.config.loader.load_config", return_value=config), \
         patch("xlvbatools.logging.setup_logging"):
        with pytest.raises(SystemExit) as exc_info:
            main([
                "run", "VerifyNamedRange",
                "--named-range", "TestInput=42",
                "--no-save",
                "--timeout", "90",
            ])

    assert exc_info.value.code == 0
    payload = json.loads(capsys.readouterr().out)
    assert payload["success"] is True, payload
    assert payload["data"]["macro"] == "VerifyNamedRange"
    assert payload["diagnostics"]["cleanup"]["still_running"] is False

    with zipfile.ZipFile(runtime_error_workbook) as workbook_zip:
        sheet_xml = ET.fromstring(workbook_zip.read("xl/worksheets/sheet1.xml"))
    namespace = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    saved_value = sheet_xml.find(".//x:c[@r='C1']/x:v", namespace)
    assert saved_value is not None
    assert saved_value.text == "0"


@pytest.mark.unit
def test_cli_dump_forwards_parser_defaults_and_prints_structured_json(tmp_path, capsys):
    """Dump owns its timeout/hidden defaults and emits machine-readable results."""
    from xlvbatools.results import InspectionOutput
    expected = _result(
        "inspect",
        InspectionOutput(None, {"Input": "screenshots/Input.png"}),
    )
    fake_project = patch("xlvbatools.cli.commands._project").start().return_value
    fake_project.inspect.return_value = expected
    with patch("xlvbatools.config.loader.load_config") as mock_config, \
         patch("xlvbatools.logging.setup_logging"):
        mock_config.return_value.workbook = "configured.xlsm"
        mock_config.return_value.log_dir = str(tmp_path)
        mock_config.return_value.log_name = "test_dump"
        main([
            "dump", "--workbook", "book.xlsm", "--sheets", "Input",
            "--screenshot", "--range", "B91:K99",
        ])

    assert json.loads(capsys.readouterr().out) == expected.to_dict()
    fake_project.inspect.assert_called_once_with(
        ["Input"], output_dir="screenshots",
        cell_range="B91:K99", include_data=False, include_screenshots=True,
        output_json=None, output_markdown=None, timeout=60.0,
        include_hidden_sheets=False,
    )
    patch.stopall()


@pytest.mark.unit
def test_cli_lint_supports_one_source_file(tmp_path, capsys):
    from xlvbatools.config.schema import LintConfig, XlvbaConfig

    source = tmp_path / "Broken.bas"
    source.write_text("Option Explicit\nx = 42\n", encoding="utf-8")

    with patch("xlvbatools.config.loader.load_config") as mock_config:
        mock_config.return_value = XlvbaConfig(
            workbook=str(tmp_path / "book.xlsm"),
            vba_source=str(tmp_path / "vba_source"),
            log_dir=str(tmp_path),
            log_name="test_lint",
            lint=LintConfig(disabled_rules=["DC003"]),
        )

        with pytest.raises(SystemExit) as exc_info:
            main(["lint", "--source", str(source)])

    assert exc_info.value.code == 1
    output = capsys.readouterr()
    payload = json.loads(output.out)
    assert payload["success"] is False
    assert any(issue["rule_id"] == "IP001" for issue in payload["data"])


@pytest.mark.unit
def test_cli_lint_missing_target_fails_without_pass(tmp_path, capsys):
    from xlvbatools.config.schema import XlvbaConfig

    with patch("xlvbatools.config.loader.load_config") as mock_config:
        mock_config.return_value = XlvbaConfig(
            workbook=str(tmp_path / "book.xlsm"),
            vba_source=str(tmp_path / "vba_source"),
            log_dir=str(tmp_path),
            log_name="test_lint",
        )

        with pytest.raises(SystemExit) as exc_info:
            main(["lint", "--source", str(tmp_path / "missing.bas")])

    assert exc_info.value.code == 2
    output = capsys.readouterr()
    payload = json.loads(output.out)
    assert payload["success"] is False
    assert payload["error"]["code"] == "invalid_lint_target"
    assert "expected a directory or VBA source file" in payload["error"]["message"]


@pytest.mark.unit
def test_cli_modify(tmp_path, capsys):
    """Test xlvba modify subcommand."""
    fake_project = patch("xlvbatools.cli.commands._project").start().return_value
    fake_project.modify.return_value = _result("modify", True)
    with patch("xlvbatools.config.loader.load_config") as mock_config, \
         patch("xlvbatools.logging.setup_logging"):
        mock_cfg = mock_config.return_value
        mock_cfg.workbook = "test.xlsm"
        mock_cfg.log_dir = str(tmp_path)
        mock_cfg.log_name = "test_modify"
        with pytest.raises(SystemExit) as exc_info:
            main(["modify", "--cell", "A1", "--value", "100", "--sheet", "Sheet2"])
        assert exc_info.value.code == 0
        
        fake_project.modify.assert_called_once_with(
            sheet="Sheet2",
            cell="A1",
            value=100,
            formula=None,
            name=None,
            refers_to=None,
            delete_name=False,
            timeout=120.0,
        )
    payload = json.loads(capsys.readouterr().out)
    assert payload["operation"] == "modify"
    assert payload["success"] is True
    patch.stopall()


@pytest.mark.unit
def test_cli_snapshot_commands(tmp_path, capsys):
    """Test xlvba snapshot create and list subcommands."""
    from xlvbatools import SnapshotRecord

    with patch("xlvbatools.config.loader.load_config") as mock_config, \
         patch("xlvbatools.SnapshotService") as mock_service_cls:
        
        mock_cfg = mock_config.return_value
        mock_cfg.workbook = "test.xlsm"
        mock_cfg.vba_source = "vba_source"
        mock_cfg.snapshots_dir = "snapshots"
        mock_cfg.log_dir = str(tmp_path)
        mock_cfg.log_name = "test_snapshot"
        mock_cfg.snapshots.rolling_limit = 10
        
        mock_service = mock_service_cls.return_value
        created = SnapshotRecord(
            snapshot_id="20260707T120000",
            timestamp="2026-07-07T12:00:00",
            description="checkpoint 1",
            workbook_file="rolling/20260707T120000.xlsm",
            workbook_hash="abc",
            workbook_size_bytes=10,
            vba_source_dir=None,
            vba_hash="def",
            milestone=True,
        )
        mock_service.create.return_value = created
        mock_service.list.return_value = (created,)
        
        # Test create
        main(["snapshot", "create", "--desc", "checkpoint 1", "--milestone"])
        mock_service.create.assert_called_once_with(
            description="checkpoint 1", milestone=True
        )
        payload = json.loads(capsys.readouterr().out)
        assert payload["operation"] == "snapshot_create"
        assert payload["data"]["snapshot_id"] == "20260707T120000"

        # Test list
        main(["snapshot", "list"])
        mock_service.list.assert_called_once()
        payload = json.loads(capsys.readouterr().out)
        assert payload["operation"] == "snapshot_list"
        assert payload["data"][0]["description"] == "checkpoint 1"
        assert payload["data"][0]["milestone"] is True


@pytest.mark.unit
def test_cli_extract_inject_diff(tmp_path, capsys):
    """Test xlvba extract, inject, and diff subcommands."""
    from xlvbatools import ExtractionOutput, InjectionChange, InjectionOutput

    fake_project = patch("xlvbatools.cli.commands._project").start().return_value
    fake_project.extract.return_value = _result(
        "extract",
        ExtractionOutput(
            workbook="test.xlsm",
            output_dir="vba_source",
            extracted_at="",
            components=(),
        ),
    )
    fake_project.inject.return_value = _result(
        "inject",
        InjectionOutput(
            changes=(InjectionChange(name="modTest", status="injected"),),
        ),
    )
    fake_project.diff.return_value = _result("diff", [])
    with patch("xlvbatools.config.loader.load_config") as mock_config, \
         patch("xlvbatools.logging.setup_logging"):
        
        mock_cfg = mock_config.return_value
        mock_cfg.workbook = "test.xlsm"
        mock_cfg.vba_source = "vba_source"
        mock_cfg.log_dir = str(tmp_path)
        mock_cfg.log_name = "test_extract"
        mock_cfg.backups.limit = 5
        
        # Test extract
        main(["extract"])
        assert json.loads(capsys.readouterr().out)["operation"] == "extract"
        
        # Test inject
        main(["inject"])
        assert json.loads(capsys.readouterr().out)["operation"] == "inject"

        # Test diff
        main(["diff"])
        assert json.loads(capsys.readouterr().out)["operation"] == "diff"
        fake_project.extract.assert_called_once_with(
            output="vba_source", component=None, timeout=120.0,
        )
        fake_project.inject.assert_called_once_with(
            source="vba_source", component=None, backup=True,
            dry_run=False, timeout=120.0,
        )
        fake_project.diff.assert_called_once_with(
            source="vba_source", component=None, timeout=120.0,
        )
    patch.stopall()


@pytest.mark.unit
def test_cli_resolves_config_paths_from_nested_directory(
    tmp_path, monkeypatch,
):
    project = tmp_path / "project"
    nested = project / "tools" / "nested"
    nested.mkdir(parents=True)
    (project / "xlvbatools.toml").write_text(
        "[xlvbatools]\n"
        'workbook = "workbook/book.xlsm"\n'
        'vba_source = "workbook/vba_source"\n'
        'log_dir = "logs"\n',
        encoding="utf-8",
    )
    monkeypatch.chdir(nested)

    with patch("xlvbatools.execution.IsolatedExecutor.execute") as execute:
        execute.return_value = _result("extract", {"components": []})
        main(["extract"])

    request = execute.call_args.args[0]
    assert request.operation.value == "extract"
    assert request.arguments == {
        "workbook_path": str((project / "workbook" / "book.xlsm").resolve()),
        "output_dir": str((project / "workbook" / "vba_source").resolve()),
        "component": None,
    }
    assert request.timeout == 120.0
    assert (project / "logs").is_dir()
