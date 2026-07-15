"""
Tests for xlvbatools.config -- Configuration loading and validation.
"""

import pytest


@pytest.mark.unit
class TestConfigSchema:
    """Test XlvbaConfig dataclass."""

    def test_defaults(self):
        from xlvbatools.config.schema import XlvbaConfig
        cfg = XlvbaConfig()
        assert cfg.workbook == "workbook.xlsm"
        assert cfg.vba_source == "vba_source"
        assert cfg.snapshots.rolling_limit == 10
        assert cfg.backups.limit == 5
        assert cfg.lint.disabled_rules == []

    def test_validation_passes(self):
        from xlvbatools.config.schema import XlvbaConfig
        cfg = XlvbaConfig()
        errors = cfg.validate()
        assert errors == []

    def test_validation_fails_bad_ext(self):
        from xlvbatools.config.schema import XlvbaConfig
        cfg = XlvbaConfig(workbook="file.xlsx")
        errors = cfg.validate()
        assert len(errors) == 1
        assert ".xlsm" in errors[0]

    def test_validation_fails_bad_limit(self):
        from xlvbatools.config.schema import XlvbaConfig, SnapshotConfig
        cfg = XlvbaConfig(snapshots=SnapshotConfig(rolling_limit=0))
        errors = cfg.validate()
        assert len(errors) == 1


@pytest.mark.unit
class TestConfigLoader:
    """Test config file loading."""

    def test_load_default_when_no_file(self, tmp_path, monkeypatch):
        from xlvbatools.config.loader import load_config
        monkeypatch.chdir(tmp_path)
        cfg = load_config()
        assert cfg.workbook == "workbook.xlsm"

    def test_load_from_file(self, tmp_path, monkeypatch):
        from xlvbatools.config.loader import load_config
        toml_content = '''
[xlvbatools]
workbook = "my_project.xlsm"
vba_source = "src/vba"
log_dir = "output/logs"

[xlvbatools.snapshots]
rolling_limit = 5

[xlvbatools.lint]
disabled_rules = ["PF001", "PF003"]
'''
        (tmp_path / "xlvbatools.toml").write_text(toml_content, encoding="utf-8")
        monkeypatch.chdir(tmp_path)
        cfg = load_config()
        assert cfg.workbook == "my_project.xlsm"
        assert cfg.vba_source == "src/vba"
        assert cfg.snapshots.rolling_limit == 5
        assert "PF001" in cfg.lint.disabled_rules
        assert cfg.workbook_path == str((tmp_path / "my_project.xlsm").resolve())
        assert cfg.vba_source_path == str((tmp_path / "src" / "vba").resolve())
        assert cfg.log_dir_path == str((tmp_path / "output" / "logs").resolve())

    def test_find_walks_up(self, tmp_path, monkeypatch):
        from xlvbatools.config.loader import find_config
        # Put config in parent
        (tmp_path / "xlvbatools.toml").write_text("[xlvbatools]\n", encoding="utf-8")
        child = tmp_path / "subdir" / "deep"
        child.mkdir(parents=True)
        monkeypatch.chdir(child)
        result = find_config()
        assert result is not None
        assert result == str(tmp_path / "xlvbatools.toml")

    def test_find_returns_none(self, tmp_path, monkeypatch):
        from xlvbatools.config.loader import find_config
        monkeypatch.chdir(tmp_path)
        assert find_config() is None
