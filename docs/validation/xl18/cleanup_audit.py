"""Local XL-18 evidence collector; delegates the strict public assertion unchanged."""
import ctypes
from ctypes import wintypes
import json
import os
from pathlib import Path

from xlvbatools import OperationResult

_original = OperationResult.require_clean_shutdown
_probe_details = {}

def _running(pid, expected_image):
    kernel = ctypes.WinDLL('kernel32', use_last_error=True)
    kernel.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    kernel.OpenProcess.restype = wintypes.HANDLE
    kernel.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
    kernel.WaitForSingleObject.restype = wintypes.DWORD
    kernel.CloseHandle.argtypes = [wintypes.HANDLE]
    handle = kernel.OpenProcess(0x00101000, False, pid)
    if not handle:
        error = ctypes.get_last_error()
        if error == 87:
            return False
        raise ctypes.WinError(error)
    try:
        state = kernel.WaitForSingleObject(handle, 0)
        if state == 0:
            return False
        if state == 258:
            import time
            code = wintypes.DWORD()
            kernel.GetExitCodeProcess.argtypes = [wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD)]
            ok = kernel.GetExitCodeProcess(handle, ctypes.byref(code))
            buffer = ctypes.create_unicode_buffer(32768)
            size = wintypes.DWORD(len(buffer))
            kernel.QueryFullProcessImageNameW.argtypes = [wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD)]
            image_ok = kernel.QueryFullProcessImageNameW(handle, 0, buffer, ctypes.byref(size))
            created, ended, kernel_time, user_time = (wintypes.FILETIME() for _ in range(4))
            kernel.GetProcessTimes.argtypes = [wintypes.HANDLE] + [ctypes.POINTER(wintypes.FILETIME)] * 4
            times_ok = kernel.GetProcessTimes(handle, ctypes.byref(created), ctypes.byref(ended), ctypes.byref(kernel_time), ctypes.byref(user_time))
            time.sleep(0.05)
            _probe_details[str(pid)] = {
                'image': buffer.value if image_ok else None, 'expected_image': expected_image,
                'pid_reused_for_other_image': bool(image_ok and Path(buffer.value).name.casefold() != expected_image.casefold()),
                'exit_code': code.value if ok else None,
                'creation_filetime': (created.dwHighDateTime << 32) | created.dwLowDateTime if times_ok else None,
                'exit_filetime': (ended.dwHighDateTime << 32) | ended.dwLowDateTime if times_ok else None,
                'wait_after_50ms': kernel.WaitForSingleObject(handle, 0),
            }
            if image_ok and Path(buffer.value).name.casefold() != expected_image.casefold():
                return False  # This is not the operation-owned Excel/worker process.
            return True
        raise ctypes.WinError(ctypes.get_last_error())
    finally:
        kernel.CloseHandle(handle)

def _audited(self, *args, **kwargs):
    try:
        return _original(self, *args, **kwargs)
    finally:
        diagnostics = self.to_dict().get('diagnostics') or {}
        if diagnostics.get('attempts') and diagnostics.get('worker_pid'):
            pids = [diagnostics.get('excel_pid'), diagnostics.get('worker_pid')]
            _probe_details.clear()
            running = {str(pid): _running(pid, 'EXCEL.EXE' if pid == diagnostics.get('excel_pid') else 'python.exe') for pid in pids if pid}
            path = Path(os.environ['XL18_AUDIT_PATH'])
            with path.open('a', encoding='utf-8') as out:
                out.write(json.dumps({
                    'request_id': self.request_id, 'operation': self.operation,
                    'success': self.success, 'diagnostics': diagnostics,
                    'owned_pids_still_running': running, 'live_pid_details': dict(_probe_details),
                }, default=str) + '\n')
            assert not any(running.values()), running

def pytest_configure(config):
    OperationResult.require_clean_shutdown = _audited

def pytest_unconfigure(config):
    OperationResult.require_clean_shutdown = _original
