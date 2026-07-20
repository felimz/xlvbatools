"""Stable selection and baseline support for public lint operations."""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
import hashlib
import json
import os
from pathlib import Path
import re
import tempfile
from typing import Iterable, Sequence

from xlvbatools.analysis.issue import VBAIssue
from xlvbatools.analysis.rules import ALL_RULE_IDS


LINT_BASELINE_SCHEMA_VERSION = "1.0"
LINT_SEVERITIES = frozenset({"ERROR", "WARNING", "STYLE"})
_WHITESPACE_RE = re.compile(r"\s+")


@dataclass(frozen=True)
class LintSelection:
    """Selected findings plus machine-readable selection counts."""

    issues: tuple[VBAIssue, ...]
    raw_count: int
    selected_count: int
    suppressed_count: int
    baseline_count: int
    known_count: int
    new_count: int
    severities: tuple[str, ...]
    rules: tuple[str, ...]
    baseline: str | None
    new_only: bool

    def metadata(self) -> dict[str, object]:
        return {
            "raw_issue_count": self.raw_count,
            "issue_count": self.selected_count,
            "suppressed_count": self.suppressed_count,
            "baseline_count": self.baseline_count,
            "known_issue_count": self.known_count,
            "new_issue_count": self.new_count,
            "severity_filter": self.severities,
            "rule_filter": self.rules,
            "baseline": self.baseline,
            "new_only": self.new_only,
        }


def select_lint_issues(
    issues: Iterable[VBAIssue],
    *,
    severities: Sequence[str] | None = None,
    rules: Sequence[str] | None = None,
    baseline: str | os.PathLike[str] | None = None,
    new_only: bool = False,
) -> LintSelection:
    """Apply severity, rule, and optional baseline/new-only selection."""
    raw = tuple(issues)
    severity_filter = _normalize_severities(severities)
    rule_filter = _normalize_rules(rules)
    if new_only and baseline is None:
        raise ValueError("new_only requires a baseline path")

    baseline_path = str(Path(baseline).resolve()) if baseline is not None else None
    baseline_fingerprints = (
        read_lint_baseline(baseline_path) if baseline_path is not None else ()
    )
    remaining_known = Counter(baseline_fingerprints)

    selected: list[VBAIssue] = []
    known_count = 0
    new_count = 0
    for issue in raw:
        fingerprint = lint_issue_fingerprint(issue)
        is_known = remaining_known[fingerprint] > 0
        if is_known:
            remaining_known[fingerprint] -= 1
            known_count += 1
        else:
            new_count += 1

        if severity_filter and issue.severity.upper() not in severity_filter:
            continue
        if rule_filter and issue.rule_id.upper() not in rule_filter:
            continue
        if new_only and is_known:
            continue
        selected.append(issue)

    return LintSelection(
        issues=tuple(selected),
        raw_count=len(raw),
        selected_count=len(selected),
        suppressed_count=len(raw) - len(selected),
        baseline_count=len(baseline_fingerprints),
        known_count=known_count,
        new_count=new_count,
        severities=tuple(sorted(severity_filter)),
        rules=tuple(sorted(rule_filter)),
        baseline=baseline_path,
        new_only=new_only,
    )


def lint_issue_fingerprint(issue: VBAIssue) -> str:
    """Return a stable identity that survives line movement and VBA casing changes."""
    identity = {
        "rule_id": issue.rule_id.strip().upper(),
        "severity": issue.severity.strip().upper(),
        "module": _normalize_text(issue.module).replace("\\", "/").casefold(),
        "procedure": _normalize_text(issue.procedure or "").casefold(),
        "message": _normalize_text(issue.message).casefold(),
    }
    encoded = json.dumps(
        identity, sort_keys=True, ensure_ascii=False, separators=(",", ":"),
    ).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def write_lint_baseline(
    path: str | os.PathLike[str], issues: Iterable[VBAIssue],
) -> str:
    """Atomically write a deterministic, versioned lint baseline."""
    target = Path(path).resolve()
    target.parent.mkdir(parents=True, exist_ok=True)
    entries = [
        {
            "fingerprint": lint_issue_fingerprint(issue),
            "rule_id": issue.rule_id,
            "severity": issue.severity,
            "module": issue.module,
            "line_num": issue.line_num,
            "procedure": issue.procedure,
            "message": issue.message,
        }
        for issue in issues
    ]
    entries.sort(
        key=lambda item: (
            str(item["fingerprint"]),
            str(item["module"]).casefold(),
            str(item["procedure"] or "").casefold(),
        )
    )
    payload = {
        "schema_version": LINT_BASELINE_SCHEMA_VERSION,
        "findings": entries,
    }

    descriptor, temporary = tempfile.mkstemp(
        prefix=f".{target.name}.", suffix=".tmp", dir=str(target.parent),
    )
    try:
        with os.fdopen(descriptor, "w", encoding="utf-8", newline="\n") as stream:
            json.dump(payload, stream, indent=2, ensure_ascii=False)
            stream.write("\n")
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temporary, target)
    except Exception:
        try:
            os.unlink(temporary)
        except OSError:
            pass
        raise
    return str(target)


def read_lint_baseline(path: str | os.PathLike[str]) -> tuple[str, ...]:
    """Read and validate fingerprints from a lint baseline."""
    target = Path(path).resolve()
    try:
        payload = json.loads(target.read_text(encoding="utf-8"))
    except FileNotFoundError:
        raise FileNotFoundError(f"Lint baseline not found: {target}") from None
    except json.JSONDecodeError as error:
        raise ValueError(f"Invalid lint baseline JSON at {target}: {error}") from error

    if not isinstance(payload, dict):
        raise ValueError(f"Invalid lint baseline at {target}: root must be an object")
    version = payload.get("schema_version")
    if version != LINT_BASELINE_SCHEMA_VERSION:
        raise ValueError(
            f"Unsupported lint baseline schema {version!r}; "
            f"expected {LINT_BASELINE_SCHEMA_VERSION!r}"
        )
    findings = payload.get("findings")
    if not isinstance(findings, list):
        raise ValueError(f"Invalid lint baseline at {target}: findings must be a list")
    fingerprints: list[str] = []
    for index, item in enumerate(findings):
        if not isinstance(item, dict):
            raise ValueError(
                f"Invalid lint baseline at {target}: finding {index} must be an object"
            )
        fingerprint = item.get("fingerprint")
        if not isinstance(fingerprint, str) or not re.fullmatch(
            r"[0-9a-f]{64}", fingerprint,
        ):
            raise ValueError(
                f"Invalid lint baseline at {target}: finding {index} has no valid fingerprint"
            )
        required = ("rule_id", "severity", "module", "message")
        if any(not isinstance(item.get(name), str) for name in required):
            raise ValueError(
                f"Invalid lint baseline at {target}: finding {index} has invalid identity fields"
            )
        procedure = item.get("procedure")
        if procedure is not None and not isinstance(procedure, str):
            raise ValueError(
                f"Invalid lint baseline at {target}: finding {index} has invalid procedure"
            )
        expected = lint_issue_fingerprint(VBAIssue(
            rule_id=item["rule_id"],
            severity=item["severity"],
            module=item["module"],
            line_num=0,
            message=item["message"],
            procedure=procedure,
        ))
        if fingerprint != expected:
            raise ValueError(
                f"Invalid lint baseline at {target}: finding {index} fingerprint mismatch"
            )
        fingerprints.append(fingerprint)
    return tuple(fingerprints)


def _normalize_severities(values: Sequence[str] | None) -> frozenset[str]:
    normalized = frozenset(str(value).strip().upper() for value in (values or ()))
    invalid = normalized - LINT_SEVERITIES
    if invalid:
        choices = ", ".join(sorted(LINT_SEVERITIES))
        raise ValueError(
            f"Unknown lint severity {sorted(invalid)[0]!r}; choose from: {choices}"
        )
    return normalized


def _normalize_rules(values: Sequence[str] | None) -> frozenset[str]:
    normalized = frozenset(str(value).strip().upper() for value in (values or ()))
    if "" in normalized:
        raise ValueError("lint rule filters must be non-empty")
    invalid = normalized - ALL_RULE_IDS
    if invalid:
        raise ValueError(
            f"Unknown lint rule {sorted(invalid)[0]!r}; use 'xlvba help lint' "
            "to discover valid rule IDs"
        )
    return normalized


def _normalize_text(value: str) -> str:
    return _WHITESPACE_RE.sub(" ", value.strip())
