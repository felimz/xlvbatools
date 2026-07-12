"""Isolated worker entry point for read-only workbook inspection."""

import json
import os
import sys


def main() -> int:
    request_path, result_path, progress_path = sys.argv[1:4]
    with open(request_path, encoding="utf-8") as handle:
        request = json.load(handle)

    def report_excel_pid(pid: int) -> None:
        temp_path = progress_path + ".tmp"
        with open(temp_path, "w", encoding="utf-8") as handle:
            json.dump({"excel_pid": pid, "worker_pid": os.getpid()}, handle)
        os.replace(temp_path, progress_path)

    from xlvbatools.workbook.dumper import _inspect_workbook_in_process
    result = _inspect_workbook_in_process(
        **request, on_excel_started=report_excel_pid,
    )
    temp_path = result_path + ".tmp"
    with open(temp_path, "w", encoding="utf-8") as handle:
        json.dump(result, handle, default=str)
    os.replace(temp_path, result_path)
    return 0 if result.get("success") else 1


if __name__ == "__main__":
    raise SystemExit(main())
