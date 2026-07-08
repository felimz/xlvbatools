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
    procedure: str | None = None  # Containing procedure/method name

    def to_dict(self) -> dict:
        return asdict(self)

    def __str__(self) -> str:
        loc_parts = [self.module]
        if self.procedure:
            loc_parts.append(f"In procedure '{self.procedure}'")
        if self.line_num:
            loc_parts.append(f"line {self.line_num}")
        loc = " -> ".join(loc_parts)
        return f"{self.severity} [{self.rule_id}] {loc}: {self.message}"
