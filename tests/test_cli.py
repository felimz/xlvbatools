"""
Unit tests for the xlvba command line interface.
"""

import os
import sys
import json
import pytest
from unittest.mock import patch
from xlvbatools.cli.main import main
from xlvbatools.results import OperationResult


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
    """Test xlvba --version."""
    with pytest.raises(SystemExit) as exc_info:
        main(["--version"])
    assert exc_info.value.code == 0
    # version prints to stdout or stderr depending on python version
    captured = capsys.readouterr()
    output = captured.out or captured.err
    assert "xlvba" in output


@pytest.mark.unit
def test_cli_detailed_version_is_structured_json(capsys):
    main(["version", "--json"])

    payload = json.loads(capsys.readouterr().out)
    assert payload["version"]
    assert os.path.basename(payload["python_executable"]).lower().startswith("python")
    assert payload["package_path"].endswith("xlvbatools")


@pytest.mark.unit
def test_cli_help(capsys):
    """Test running xlvba with no args displays help."""
    with pytest.raises(SystemExit) as exc_info:
        main([])
    assert exc_info.value.code == 0
    captured = capsys.readouterr()
    assert "usage:" in captured.out.lower() or "usage:" in captured.err.lower()


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
    
    captured = capsys.readouterr()
    assert "modTest.bas:4: Public Sub TestSub()" in captured.out
    assert "1 match(es)" in captured.out

    # Test regex search
    with patch("xlvbatools.config.loader.load_config") as mock_config:
        mock_cfg = mock_config.return_value
        mock_cfg.vba_source = str(temp_vba_source)
        mock_cfg.log_dir = str(tmp_path)
        mock_cfg.log_name = "test_search"
        main(["search", "--regex", r"Dim\s+\w+\s+As"])

    captured = capsys.readouterr()
    assert "Dim x As Double" in captured.out


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
    
    captured = capsys.readouterr()
    assert "file(s) would be changed" in captured.out or "No changes" in captured.out


@pytest.mark.unit
def test_cli_graph(temp_vba_source, sample_bas_file, tmp_path, capsys):
    """Test xlvba graph subcommand."""
    # Test generating graph to stdout (mermaid format)
    with patch("xlvbatools.config.loader.load_config") as mock_config:
        mock_cfg = mock_config.return_value
        mock_cfg.vba_source = str(temp_vba_source)
        mock_cfg.log_dir = str(tmp_path)
        mock_cfg.log_name = "test_graph"
        main(["graph", "--format", "mermaid"])

    captured = capsys.readouterr()
    assert "graph TD" in captured.out
    assert "modTest.TestSub" in captured.out or "modTest" in captured.out


@pytest.mark.unit
def test_cli_init(tmp_path, monkeypatch, capsys):
    """Test xlvba init subcommand."""
    monkeypatch.chdir(tmp_path)
    
    # 1. First init
    main(["init", "--workbook", "my_test.xlsm"])
    captured = capsys.readouterr()
    assert "xlvbatools.toml" in captured.out
    assert os.path.exists("xlvbatools.toml")
    
    # Check config content
    with open("xlvbatools.toml", "r") as f:
        content = f.read()
    assert 'workbook = "my_test.xlsm"' in content

    # 2. Re-init without force (should fail)
    with pytest.raises(SystemExit) as exc_info:
        main(["init", "--workbook", "new.xlsm"])
    assert exc_info.value.code == 1
    captured = capsys.readouterr()
    assert "already exists" in captured.out

    # 3. Re-init with force (should succeed)
    main(["init", "--workbook", "new.xlsm", "--force"])
    captured = capsys.readouterr()
    assert "xlvbatools.toml" in captured.out
    with open("xlvbatools.toml", "r") as f:
        content = f.read()
    assert 'workbook = "new.xlsm"' in content

    # 4. Re-init with force and --agents (should install templates)
    main(["init", "--workbook", "new.xlsm", "--force", "--agents"])
    assert os.path.exists(".agents")
    assert os.path.exists(".agents/AGENTS.md")
    assert os.path.exists(".agents/skills/xlvba-toolchain/SKILL.md")
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
    assert "xlvbatools v1 agent integration" in capsys.readouterr().out


@pytest.mark.unit
def test_cli_run_forwards_timeout(tmp_path):
    project = patch("xlvbatools.cli.main._project").start().return_value
    project.run.return_value = _result("run_macro", {"macro": "MyMacro"})
    with patch("xlvbatools.config.loader.load_config") as mock_config, \
         patch("xlvbatools.logging.setup_logging"):
        mock_config.return_value.workbook = "book.xlsm"
        mock_config.return_value.log_dir = str(tmp_path)
        mock_config.return_value.log_name = "test"

        with pytest.raises(SystemExit) as exc_info:
            main(["run", "MyMacro", "--timeout", "7.5"])

    assert exc_info.value.code == 0
    project.run.assert_called_once_with("MyMacro", timeout=7.5)
    patch.stopall()


@pytest.mark.unit
def test_cli_dump_forwards_parser_defaults_and_prints_structured_json(tmp_path, capsys):
    """Dump owns its timeout/hidden defaults and emits machine-readable results."""
    from xlvbatools.results import InspectionOutput
    expected = _result(
        "inspect",
        InspectionOutput(None, {"Input": "screenshots/Input.png"}),
    )
    fake_project = patch("xlvbatools.cli.main._project").start().return_value
    fake_project.inspect.return_value = expected
    with patch("xlvbatools.config.loader.load_config") as mock_config, \
         patch("xlvbatools.logging.setup_logging"):
        mock_config.return_value.workbook = "configured.xlsm"
        mock_config.return_value.log_dir = str(tmp_path)
        mock_config.return_value.log_name = "test_dump"
        main([
            "dump", "--workbook", "book.xlsm", "--sheets", "Input",
            "--screenshot", "--range", "B91:K99", "--json",
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
    assert "IP001" in output.out
    assert "FAIL:" in output.out
    assert "PASS" not in output.out


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
    assert "expected a directory or VBA source file" in output.err
    assert "PASS" not in output.out + output.err


@pytest.mark.unit
def test_cli_modify(tmp_path, capsys):
    """Test xlvba modify subcommand."""
    fake_project = patch("xlvbatools.cli.main._project").start().return_value
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
    captured = capsys.readouterr()
    assert "OK" in captured.out
    patch.stopall()


@pytest.mark.unit
def test_cli_snapshot_commands(tmp_path, capsys):
    """Test xlvba snapshot create and list subcommands."""
    with patch("xlvbatools.config.loader.load_config") as mock_config, \
         patch("xlvbatools.snapshot.manager.SnapshotManager") as mock_mgr_cls:
        
        mock_cfg = mock_config.return_value
        mock_cfg.workbook = "test.xlsm"
        mock_cfg.vba_source = "vba_source"
        mock_cfg.snapshots_dir = "snapshots"
        mock_cfg.log_dir = str(tmp_path)
        mock_cfg.log_name = "test_snapshot"
        mock_cfg.snapshots.rolling_limit = 10
        
        mock_mgr = mock_mgr_cls.return_value
        mock_mgr.create.return_value = "20260707T120000000000"
        mock_mgr.list.return_value = [
            {"snapshot_id": "snap1", "description": "checkpoint 1", "milestone": True}
        ]
        
        # Test create
        main(["snapshot", "create", "--desc", "checkpoint 1", "--milestone"])
        mock_mgr.create.assert_called_once_with(description="checkpoint 1", milestone=True)
        captured = capsys.readouterr()
        assert "Snapshot created: 20260707T120000000000" in captured.out

        # Test list
        main(["snapshot", "list"])
        mock_mgr.list.assert_called_once()
        captured = capsys.readouterr()
        assert "snap1" in captured.out
        assert "checkpoint 1" in captured.out
        assert "[MILESTONE]" in captured.out


@pytest.mark.unit
def test_cli_extract_inject_diff(tmp_path, capsys):
    """Test xlvba extract, inject, and diff subcommands."""
    fake_project = patch("xlvbatools.cli.main._project").start().return_value
    fake_project.extract.return_value = _result(
        "extract", {"components": [{"name": "modTest"}]},
    )
    fake_project.inject.return_value = _result(
        "inject", [{"name": "modTest", "status": "injected"}],
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
        
        # Test inject
        main(["inject"])

        # Test diff
        main(["diff"])
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
