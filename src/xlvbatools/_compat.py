"""
Cross-Platform Compatibility Layer
====================================
Provides graceful degradation for non-Windows platforms where COM
automation is not available. VBA source file operations (lint, diff,
search, format) work on any platform; COM-dependent operations raise
PlatformError with a clear message.
"""

import sys

IS_WINDOWS = sys.platform == "win32"


class PlatformError(RuntimeError):
    """Raised when a COM-dependent operation is called on a non-Windows platform."""
    pass


def require_windows(operation: str = "This operation"):
    """Raise PlatformError if not on Windows."""
    if not IS_WINDOWS:
        raise PlatformError(
            f"{operation} requires Windows with Excel installed. "
            f"Current platform: {sys.platform}"
        )


def import_win32com():
    """Import and return win32com.client, or raise PlatformError."""
    require_windows("Excel COM automation")
    import win32com.client
    return win32com.client


def import_win32process():
    """Import and return win32process, or raise PlatformError."""
    require_windows("Excel process management")
    import win32process
    return win32process
