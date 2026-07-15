"""Contract tests for documentation and generated agent guidance."""

from pathlib import Path
import re

import pytest


ROOT = Path(__file__).resolve().parents[1]
TEMPLATE_FILES = (
    Path("AGENTS.md"),
    Path("rules/python-rules.md"),
    Path("rules/vba-rules.md"),
    Path("skills/xlvba-toolchain/SKILL.md"),
    Path("workflows/vba-debug.md"),
    Path("workflows/vba-edit.md"),
)


def _documentation_files() -> tuple[Path, ...]:
    top_level = (ROOT / "README.md", ROOT / "CONTRIBUTING.md")
    trees = (
        ROOT / "docs",
        ROOT / ".agents",
        ROOT / "templates/agents",
        ROOT / "src/xlvbatools/templates/agents",
    )
    return top_level + tuple(
        path
        for tree in trees
        for path in tree.rglob("*.md")
    )


@pytest.mark.unit
def test_agent_templates_are_identical_across_all_three_surfaces():
    for relative in TEMPLATE_FILES:
        active = (ROOT / ".agents" / relative).read_bytes()
        repository = (ROOT / "templates/agents" / relative).read_bytes()
        packaged = (
            ROOT / "src/xlvbatools/templates/agents" / relative
        ).read_bytes()
        assert active == repository == packaged, relative


@pytest.mark.unit
def test_documented_local_markdown_links_resolve():
    link_pattern = re.compile(r"\[[^\]]+\]\(([^)]+)\)")
    failures: list[str] = []
    for document in _documentation_files():
        for target in link_pattern.findall(document.read_text(encoding="utf-8")):
            if target.startswith(("http://", "https://", "mailto:", "#")):
                continue
            path_text = target.split("#", 1)[0]
            if path_text and not (document.parent / path_text).exists():
                failures.append(f"{document.relative_to(ROOT)} -> {target}")
    assert not failures, failures


@pytest.mark.unit
def test_application_guidance_uses_only_the_v1_public_boundary():
    combined = "\n".join(
        path.read_text(encoding="utf-8") for path in _documentation_files()
    )
    stale_patterns = (
        r"\bXlvbaProject\b",
        r"\bfrom_legacy\b",
        r"\bsnapshot_manager\b",
        r"from xlvbatools\.(?:core|analysis|macro|snapshot|vba|workbook)\b",
        r"\bUse ExcelSession\b",
        r"\bKill stale Excel\b",
    )
    for pattern in stale_patterns:
        assert re.search(pattern, combined) is None, pattern
    assert "xlvbatools.Project" in combined
    assert "require_clean_shutdown()" in combined


@pytest.mark.unit
def test_vba_agent_rule_explicitly_forbids_global_excel_termination():
    rules = (ROOT / ".agents/rules/vba-rules.md").read_text(encoding="utf-8")
    assert "Never run `taskkill /im EXCEL.EXE`" in rules
    assert "Only an operation-owned PID" in (
        ROOT / ".agents/rules/python-rules.md"
    ).read_text(encoding="utf-8")
