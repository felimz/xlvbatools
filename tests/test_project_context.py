"""Regression coverage for project-scoped VBA name resolution."""

from __future__ import annotations

from dataclasses import dataclass

import pytest

from xlvbatools.analysis.preflight import lint_files, lint_workbook
from xlvbatools.analysis.project_context import (
    CLASS_MODULE,
    STANDARD_MODULE,
    VBAModuleSource,
    build_project_index,
    lint_project_modules,
)
from xlvbatools.analysis.rules import ALL_RULES
from xlvbatools.vba.constants import TYPE_STD_MODULE


def _uv_only_disabled_rules() -> list[str]:
    return [rule_id for rule_id in ALL_RULES if rule_id != "UV001"] + [
        "DC003",
        "DP001",
    ]


def _module(
    name: str,
    code: str,
    *,
    kind: str = STANDARD_MODULE,
    directory: str = "modules",
    extension: str = ".bas",
) -> VBAModuleSource:
    return VBAModuleSource.create(
        name=name,
        rel_path=f"{directory}/{name}{extension}",
        module_kind=kind,
        lines=code.splitlines(),
    )


@pytest.mark.unit
@pytest.mark.parametrize("declaration", ["Public FileCount As Long", "Global FileCount As Long"])
def test_standard_module_project_symbol_resolves_across_modules(declaration):
    modules = [
        _module("modConstants", f"Option Explicit\n{declaration}\n"),
        _module(
            "modConsumer",
            "Option Explicit\n"
            "Public Sub UpdateFileCount()\n"
            "    FileCount = 1\n"
            "End Sub\n",
        ),
    ]

    issues, _ = lint_project_modules(modules, _uv_only_disabled_rules())

    assert [issue for issue in issues if issue.rule_id == "UV001"] == []


@pytest.mark.unit
def test_private_symbol_does_not_resolve_across_modules():
    modules = [
        _module("modConstants", "Option Explicit\nPrivate FileCount As Long\n"),
        _module(
            "modConsumer",
            "Option Explicit\n"
            "Public Sub UpdateFileCount()\n"
            "    FileCount = 1\n"
            "End Sub\n",
        ),
    ]

    issues, _ = lint_project_modules(modules, _uv_only_disabled_rules())

    undeclared = [issue for issue in issues if issue.rule_id == "UV001"]
    assert len(undeclared) == 1
    assert "FileCount" in undeclared[0].message
    assert undeclared[0].module == "modules/modConsumer.bas"


@pytest.mark.unit
def test_public_class_field_does_not_leak_into_project_namespace():
    modules = [
        _module(
            "FileState",
            "Option Explicit\nPublic FileCount As Long\n",
            kind=CLASS_MODULE,
            directory="classes",
            extension=".cls",
        ),
        _module(
            "modConsumer",
            "Option Explicit\n"
            "Public Sub UpdateFileCount()\n"
            "    FileCount = 1\n"
            "End Sub\n",
        ),
    ]

    issues, index = lint_project_modules(modules, _uv_only_disabled_rules())

    assert "filecount" not in index.project_public_names
    assert index.class_registry["filestate"] == {"filecount"}
    assert len([issue for issue in issues if issue.rule_id == "UV001"]) == 1


@pytest.mark.unit
def test_current_module_private_symbol_remains_visible():
    module = _module(
        "modConsumer",
        "Option Explicit\n"
        "Private FileCount As Long\n"
        "Public Sub UpdateFileCount()\n"
        "    FileCount = 1\n"
        "End Sub\n",
    )

    index = build_project_index([module])
    issues, _ = lint_project_modules([module], _uv_only_disabled_rules())

    assert "filecount" in index.visible_names_for(module.rel_path)
    assert [issue for issue in issues if issue.rule_id == "UV001"] == []


class _FakeCodeModule:
    def __init__(self, code: str):
        self._code = code
        self.CountOfLines = len(code.splitlines())

    def Lines(self, _start: int, _count: int) -> str:
        return self._code


@dataclass
class _FakeComponent:
    Name: str
    code: str
    Type: int = TYPE_STD_MODULE

    def __post_init__(self):
        self.CodeModule = _FakeCodeModule(self.code)


class _FakeProject:
    def __init__(self, components):
        self.VBComponents = components


class _FakeSession:
    def __init__(self, components):
        self.vb_project = _FakeProject(components)


@pytest.mark.unit
def test_file_and_live_workbook_loaders_share_project_resolution(tmp_path):
    constants_code = "Option Explicit\nPublic FileCount As Long\n"
    consumer_code = (
        "Option Explicit\n"
        "Public Sub UpdateFileCount()\n"
        "    FileCount = 1\n"
        "End Sub\n"
    )
    modules_dir = tmp_path / "modules"
    modules_dir.mkdir()
    (modules_dir / "modConstants.bas").write_text(constants_code, encoding="utf-8")
    (modules_dir / "modConsumer.bas").write_text(consumer_code, encoding="utf-8")
    disabled = _uv_only_disabled_rules()

    file_issues = lint_files(str(tmp_path), disabled_rules=disabled)
    live_issues = lint_workbook(
        str(tmp_path / "not-opened.xlsm"),
        disabled_rules=disabled,
        compile_test=False,
        _session=_FakeSession(
            [
                _FakeComponent("modConstants", constants_code),
                _FakeComponent("modConsumer", consumer_code),
            ]
        ),
    )

    def signature(issues):
        return [
            (issue.rule_id, issue.severity, issue.module, issue.line_num, issue.message)
            for issue in issues
        ]

    assert signature(file_issues) == signature(live_issues) == []
