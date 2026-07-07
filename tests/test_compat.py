"""
Tests for xlvbatools._compat -- Cross-platform compatibility layer.
"""

import sys
import pytest


@pytest.mark.unit
class TestCompat:
    """Tests for platform detection and error handling."""

    def test_is_windows_flag(self):
        from xlvbatools._compat import IS_WINDOWS
        assert IS_WINDOWS == (sys.platform == "win32")

    def test_platform_error_type(self):
        from xlvbatools._compat import PlatformError
        assert issubclass(PlatformError, RuntimeError)

    def test_require_windows_on_windows(self):
        from xlvbatools._compat import IS_WINDOWS, require_windows
        if IS_WINDOWS:
            # Should not raise
            require_windows("test operation")
        else:
            from xlvbatools._compat import PlatformError
            with pytest.raises(PlatformError, match="requires Windows"):
                require_windows("test operation")

    def test_import_win32com_on_windows(self):
        from xlvbatools._compat import IS_WINDOWS, import_win32com
        if IS_WINDOWS:
            mod = import_win32com()
            assert hasattr(mod, "Dispatch")
        else:
            from xlvbatools._compat import PlatformError
            with pytest.raises(PlatformError):
                import_win32com()
