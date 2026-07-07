"""
Tests for xlvbatools.logging -- Centralized logging configuration.
"""

import os
import pytest


@pytest.mark.unit
class TestLogging:
    """Test logging setup."""

    def test_setup_logging_creates_file(self, tmp_path):
        from xlvbatools.logging import setup_logging
        log_file = setup_logging(
            verbose=False,
            tool_name="test",
            log_dir=str(tmp_path),
            log_name="test_log",
        )
        assert os.path.exists(log_file)
        assert "test_log.log" in log_file

    def test_setup_logging_verbose(self, tmp_path):
        import logging
        from xlvbatools.logging import setup_logging
        setup_logging(
            verbose=True,
            tool_name="test_verbose",
            log_dir=str(tmp_path),
        )
        root = logging.getLogger()
        # Console handler should be at DEBUG level when verbose
        console_handlers = [h for h in root.handlers
                           if isinstance(h, logging.StreamHandler)
                           and not isinstance(h, logging.FileHandler)]
        assert any(h.level == logging.DEBUG for h in console_handlers)

    def test_get_latest_log_returns_none_for_missing(self, tmp_path):
        from xlvbatools.logging import get_latest_log
        result = get_latest_log(str(tmp_path))
        assert result is None

    def test_get_latest_log_returns_path_after_setup(self, tmp_path):
        from xlvbatools.logging import setup_logging, get_latest_log
        setup_logging(log_dir=str(tmp_path))
        result = get_latest_log(str(tmp_path))
        assert result is not None
        assert os.path.exists(result)
