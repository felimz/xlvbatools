"""
Unit tests for the xlvba command line interface.
"""

import os
import sys
import json
import pytest
from unittest.mock import patch
from xlvbatools.cli.main import main


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
    assert "Agent Integration Guide" in capsys.readouterr().out


@pytest.mark.unit
def test_cli_run_forwards_timeout(tmp_path):
    with patch("xlvbatools.config.loader.load_config") as mock_config, \
         patch("xlvbatools.macro.runner.run_macro") as mock_run:
        mock_config.return_value.workbook = "book.xlsm"
        mock_config.return_value.log_dir = str(tmp_path)
        mock_config.return_value.log_name = "test"
        mock_run.return_value = {"success": True, "elapsed_seconds": 0.1}

        with pytest.raises(SystemExit) as exc_info:
            main(["run", "MyMacro", "--timeout", "7.5"])

    assert exc_info.value.code == 0
    mock_run.assert_called_once_with("book.xlsm", "MyMacro", timeout=7.5)


@pytest.mark.unit
def test_cli_modify(tmp_path, capsys):
    """Test xlvba modify subcommand."""
    with patch("xlvbatools.config.loader.load_config") as mock_config, \
         patch("xlvbatools.workbook.modifier.modify_cell") as mock_modify:
        mock_cfg = mock_config.return_value
        mock_cfg.workbook = "test.xlsm"
        mock_cfg.log_dir = str(tmp_path)
        mock_cfg.log_name = "test_modify"
        mock_modify.return_value = True
        
        with pytest.raises(SystemExit) as exc_info:
            main(["modify", "--cell", "A1", "--value", "100", "--sheet", "Sheet2"])
        assert exc_info.value.code == 0
        
        mock_modify.assert_called_once_with(
            "test.xlsm",
            sheet="Sheet2",
            cell="A1",
            value=100,
            formula=None,
            name=None,
            refers_to=None,
            delete_name=False
        )
    captured = capsys.readouterr()
    assert "OK" in captured.out


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
    with patch("xlvbatools.config.loader.load_config") as mock_config, \
         patch("xlvbatools.vba.extractor.extract_all") as mock_extract, \
         patch("xlvbatools.vba.injector.inject_all") as mock_inject, \
         patch("xlvbatools.vba.differ.diff_all") as mock_diff:
        
        mock_cfg = mock_config.return_value
        mock_cfg.workbook = "test.xlsm"
        mock_cfg.vba_source = "vba_source"
        mock_cfg.log_dir = str(tmp_path)
        mock_cfg.log_name = "test_extract"
        mock_cfg.backups.limit = 5
        
        mock_extract.return_value = {"components": [{"name": "modTest"}]}
        mock_inject.return_value = [{"name": "modTest", "status": "injected"}]
        mock_diff.return_value = []
        
        # Test extract
        main(["extract"])
        mock_extract.assert_called_once_with("test.xlsm", "vba_source")
        
        # Test inject
        main(["inject"])
        mock_inject.assert_called_once_with(
            "test.xlsm", "vba_source", backup=True, dry_run=False, backup_limit=5
        )

        # Test diff
        main(["diff"])
        mock_diff.assert_called_once_with("test.xlsm", "vba_source")
