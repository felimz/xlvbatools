"""
VBA Issue Dataclass
====================
Structured representation of a VBA static analysis finding.
"""

from dataclasses import dataclass, asdict


@dataclass
class VBAIssue:
    """A single issue found during VBA static analysis."""

    rule_id: str        # e.g. "RK001", "CW002", "SB001"
    severity: str       # "ERROR" or "WARNING"
    module: str         # Module/file name where issue was found
    line_num: int       # 1-indexed line number (0 if not applicable)
    message: str        # Human-readable description

    def to_dict(self) -> dict:
        return asdict(self)

    def __str__(self) -> str:
        loc = f"{self.module}:{self.line_num}" if self.line_num else self.module
        return f"{self.severity} [{self.rule_id}] {loc}: {self.message}"
