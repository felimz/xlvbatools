"""Contract tests for documentation and generated agent guidance."""

from pathlib import Path
import re

import pytest


ROOT = Path(__file__).resolve().parents[1]


def _documentation_files() -> tuple[Path, ...]:
    top_level = (ROOT / "README.md", ROOT / "CONTRIBUTING.md")
    trees = (
        ROOT / "docs",
        ROOT / ".agents",
        ROOT / "src/xlvbatools/templates/agents",
    )
    return top_level + tuple(
        path
        for tree in trees
        for path in tree.rglob("*.md")
    )


@pytest.mark.unit
def test_active_agent_guidance_matches_packaged_templates():
    active_root = ROOT / ".agents"
    packaged_root = ROOT / "src/xlvbatools/templates/agents"
    active_files = {
        path.relative_to(active_root)
        for path in active_root.rglob("*")
        if path.is_file() and "__pycache__" not in path.parts
    }
    packaged_files = {
        path.relative_to(packaged_root)
        for path in packaged_root.rglob("*")
        if path.is_file() and "__pycache__" not in path.parts
    }
    assert active_files == packaged_files
    for relative in active_files:
        active = (ROOT / ".agents" / relative).read_bytes()
        packaged = (packaged_root / relative).read_bytes()
        assert active == packaged, relative
    assert not (ROOT / "templates/agents").exists()


@pytest.mark.unit
def test_readme_indexes_every_document():
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    missing = [
        path.name
        for path in (ROOT / "docs").glob("*.md")
        if f"docs/{path.name}" not in readme
    ]
    assert not missing, missing


@pytest.mark.unit
def test_api_reference_names_every_public_export():
    import xlvbatools

    reference = (ROOT / "docs/api-reference.md").read_text(encoding="utf-8")
    missing = [name for name in xlvbatools.__all__ if f"`{name}`" not in reference]
    assert not missing, missing


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
def test_documentation_describes_machine_first_cli_output():
    combined = "\n".join(
        path.read_text(encoding="utf-8") for path in _documentation_files()
    )
    assert "machine-first" in combined
    assert "JSON envelope by default" in combined
    assert "--output-format text" in combined
    assert "--output-format table" in combined
    assert " --json" not in combined


@pytest.mark.unit
def test_get_started_covers_supported_invocation_contract():
    guide = (ROOT / "docs/get-started.md").read_text(encoding="utf-8")

    required_snippets = (
        r".\.venv\Scripts\xlvba.exe",
        "xlvba help",
        "--workbook",
        "--source",
        "--timeout",
        "--dry-run",
        "--include-hidden-sheets",
        "--named-range",
        "--no-save",
        "--visible",
        "--text",
        "--table",
        "ConvertFrom-Json",
        "from xlvbatools import Project",
        "Project.from_config()",
        "Project.open(",
        "require_success()",
        "require_clean_shutdown()",
    )
    missing = [snippet for snippet in required_snippets if snippet not in guide]
    assert not missing, missing

    workflow = (ROOT / ".agents/workflows/get-started.md").read_text(
        encoding="utf-8"
    )
    assert "from xlvbatools import Project" in workflow
    assert "--dry-run --timeout" in workflow
    assert "--named-range" in workflow
    assert "--no-save" in workflow
    assert "Default stdout is one JSON result envelope" in workflow
    assert "Not currently exposed by `xlvba run`" not in guide


@pytest.mark.unit
def test_agent_installation_and_discovery_are_documented():
    readme = (ROOT / "README.md").read_text(encoding="utf-8")
    guide = (ROOT / "docs/agent-integration.md").read_text(encoding="utf-8")
    agents = (ROOT / ".agents/AGENTS.md").read_text(encoding="utf-8")
    combined = "\n".join((readme, guide, agents))

    assert "xlvba help" in combined
    assert "xlvba COMMAND --help" in combined
    assert "xlvba agents install" in combined
    assert "xlvba init --agents" in combined
    assert ".agents/` (plural)" in combined
    assert "project-specific extra files" in combined
    assert "does not install the Python package" in combined
    assert "does not" in combined and "xlvbatools.toml" in combined


@pytest.mark.unit
def test_vba_agent_rule_explicitly_forbids_global_excel_termination():
    rules = (ROOT / ".agents/rules/vba-rules.md").read_text(encoding="utf-8")
    assert "Never run `taskkill /im EXCEL.EXE`" in rules
    assert "Only an operation-owned PID" in (
        ROOT / ".agents/rules/python-rules.md"
    ).read_text(encoding="utf-8")
