"""Public static-analysis API."""

from xlvbatools.analysis.issue import VBAIssue
from xlvbatools.analysis.preflight import lint_files, lint_workbook, print_report

__all__ = ["VBAIssue", "lint_files", "lint_workbook", "print_report"]
