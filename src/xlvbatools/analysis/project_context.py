"""Project-level context shared by every VBA linting input adapter.

The analyzer must see the whole VBProject before resolving names.  This module
keeps source acquisition (files or COM) separate from VBA scope semantics so
both paths produce the same results.
"""

from __future__ import annotations

from dataclasses import dataclass
from types import MappingProxyType
from typing import Iterable, Mapping, Sequence
import os
import re


STANDARD_MODULE = "standard_module"
CLASS_MODULE = "class_module"
DOCUMENT_MODULE = "document_module"
USERFORM = "userform"


@dataclass(frozen=True)
class VBAModuleSource:
    """Normalized VBA component source, independent of its acquisition path."""

    name: str
    rel_path: str
    module_kind: str
    lines: tuple[str, ...]

    @classmethod
    def create(
        cls,
        *,
        name: str,
        rel_path: str,
        module_kind: str,
        lines: Sequence[str],
    ) -> "VBAModuleSource":
        return cls(
            name=name,
            rel_path=rel_path.replace("\\", "/"),
            module_kind=module_kind,
            lines=tuple(lines),
        )


@dataclass(frozen=True)
class VBASymbol:
    """A module-level declaration available to the project resolver."""

    name: str
    kind: str
    visibility: str
    module_name: str
    module_path: str
    module_kind: str
    line_num: int

    @property
    def key(self) -> str:
        return self.name.casefold()

    @property
    def project_visible(self) -> bool:
        # Only standard modules contribute unqualified project-level names.
        # Public class, document, and form members require an object/module
        # qualifier and must not leak into the global namespace.
        return (
            self.module_kind == STANDARD_MODULE
            and self.visibility in {"public", "global"}
        )


@dataclass(frozen=True)
class VBAProjectIndex:
    """Immutable symbol index for one complete VBA project."""

    modules: tuple[VBAModuleSource, ...]
    symbols: tuple[VBASymbol, ...]
    _public_by_name: Mapping[str, tuple[VBASymbol, ...]]
    _symbols_by_module: Mapping[str, tuple[VBASymbol, ...]]
    _class_members: Mapping[str, frozenset[str]]

    @property
    def project_public_names(self) -> frozenset[str]:
        return frozenset(self._public_by_name)

    @property
    def class_registry(self) -> dict[str, set[str]]:
        # Existing rules intentionally receive ordinary containers.  Return a
        # copy so they cannot mutate the shared project index.
        return {name: set(members) for name, members in self._class_members.items()}

    def visible_names_for(self, rel_path: str) -> frozenset[str]:
        """Return names resolvable without qualification in ``rel_path``."""
        module_key = _path_key(rel_path)
        local_module_names = {
            symbol.key for symbol in self._symbols_by_module.get(module_key, ())
        }
        return frozenset(local_module_names | set(self._public_by_name))

    def duplicate_public_procedures(self) -> tuple[tuple[VBASymbol, VBASymbol], ...]:
        """Return cross-module public procedure collisions deterministically."""
        first_by_name: dict[str, VBASymbol] = {}
        duplicates: list[tuple[VBASymbol, VBASymbol]] = []
        for symbol in self.symbols:
            if symbol.kind not in {"sub", "function", "property"}:
                continue
            if not symbol.project_visible:
                continue
            previous = first_by_name.get(symbol.key)
            if previous is None:
                first_by_name[symbol.key] = symbol
            elif _path_key(previous.module_path) != _path_key(symbol.module_path):
                duplicates.append((previous, symbol))
        return tuple(duplicates)


def module_kind_from_path(rel_path: str) -> str:
    """Infer component kind for extracted source when no manifest is needed."""
    normalized = rel_path.replace("\\", "/").casefold()
    parts = normalized.split("/")
    parent = parts[-2] if len(parts) > 1 else ""
    if parent == "modules":
        return STANDARD_MODULE
    if parent == "classes":
        return CLASS_MODULE
    if parent == "sheets":
        return DOCUMENT_MODULE
    if parent == "forms":
        return USERFORM
    extension = os.path.splitext(normalized)[1]
    if extension == ".cls":
        return CLASS_MODULE
    if extension == ".frm":
        return USERFORM
    return STANDARD_MODULE


def build_project_index(modules: Iterable[VBAModuleSource]) -> VBAProjectIndex:
    """Parse module-level declarations from every component before linting."""
    normalized_modules = tuple(modules)
    symbols: list[VBASymbol] = []
    for module in normalized_modules:
        symbols.extend(_module_symbols(module))

    public_by_name: dict[str, list[VBASymbol]] = {}
    symbols_by_module: dict[str, list[VBASymbol]] = {}
    class_members: dict[str, set[str]] = {}
    for symbol in symbols:
        symbols_by_module.setdefault(_path_key(symbol.module_path), []).append(symbol)
        if symbol.project_visible:
            public_by_name.setdefault(symbol.key, []).append(symbol)
        if (
            symbol.module_kind == CLASS_MODULE
            and symbol.visibility in {"public", "friend"}
        ):
            class_members.setdefault(symbol.module_name.casefold(), set()).add(symbol.key)

    return VBAProjectIndex(
        modules=normalized_modules,
        symbols=tuple(symbols),
        _public_by_name=MappingProxyType(
            {name: tuple(items) for name, items in public_by_name.items()}
        ),
        _symbols_by_module=MappingProxyType(
            {name: tuple(items) for name, items in symbols_by_module.items()}
        ),
        _class_members=MappingProxyType(
            {name: frozenset(items) for name, items in class_members.items()}
        ),
    )


def lint_project_modules(
    modules: Iterable[VBAModuleSource],
    disabled_rules: Sequence[str] | None = None,
) -> tuple[list["VBAIssue"], VBAProjectIndex]:
    """Run module rules after indexing the complete project exactly once."""
    from xlvbatools.analysis.issue import VBAIssue
    from xlvbatools.analysis.rules import run_all_rules

    index = build_project_index(modules)
    issues: list[VBAIssue] = []
    for module in index.modules:
        issues.extend(
            run_all_rules(
                module.rel_path,
                list(module.lines),
                list(disabled_rules) if disabled_rules is not None else None,
                project_index=index,
            )
        )

    disabled = set(disabled_rules or ())
    if "DP001" not in disabled:
        for previous, duplicate in index.duplicate_public_procedures():
            issues.append(
                VBAIssue(
                    rule_id="DP001",
                    severity="ERROR",
                    module=duplicate.module_path,
                    line_num=duplicate.line_num,
                    message=(
                        f"Duplicate public procedure '{duplicate.name}' found in both "
                        f"'{duplicate.module_path}' and '{previous.module_path}' "
                        f"(L{previous.line_num}). VBA raises a compile error "
                        "(Ambiguous name detected) for duplicate public names."
                    ),
                )
            )

    return issues, index


def _module_symbols(module: VBAModuleSource) -> list[VBASymbol]:
    # Import privately to preserve one declaration grammar while avoiding a
    # module import cycle: rules accepts VBAProjectIndex as optional context.
    from xlvbatools.analysis.rules import (
        _PROC_END_RE,
        _PROC_START_RE,
        _extract_proc_name,
        _get_logical_lines,
        _parse_vba_declarations,
        _split_colon_statements,
    )

    symbols: list[VBASymbol] = []
    in_procedure = False
    for line_num, logical_line in _get_logical_lines(list(module.lines)):
        for raw_statement in _split_colon_statements(logical_line):
            statement = raw_statement.strip()
            if not statement or statement.startswith("'") or statement.casefold().startswith("rem "):
                continue

            if _PROC_START_RE.match(statement):
                in_procedure = True
                name = _extract_proc_name(statement)
                proc_kind = _procedure_kind(statement)
                visibility = _procedure_visibility(statement)
                if name:
                    symbols.append(
                        _symbol(module, name, proc_kind, visibility, line_num)
                    )
                continue
            if _PROC_END_RE.match(statement):
                in_procedure = False
                continue
            if in_procedure:
                continue

            visibility = _declaration_visibility(statement)
            declaration = statement
            if statement.casefold().startswith("global "):
                # The shared declaration parser treats Global as Public scope.
                declaration = "Public " + statement[len("Global "):]
            parsed = _parse_vba_declarations(declaration)
            if parsed:
                kind = "constant" if re.match(
                    r"^(?:(?:Public|Private|Global)\s+)?Const\s+",
                    statement,
                    re.IGNORECASE,
                ) else "variable"
                for name, _ in parsed:
                    symbols.append(
                        _symbol(module, name, kind, visibility, line_num)
                    )
                continue

            enum_or_type = re.match(
                r"^(?:(Public|Private|Friend)\s+)?(Enum|Type)\s+(\w+)",
                statement,
                re.IGNORECASE,
            )
            if enum_or_type:
                symbols.append(
                    _symbol(
                        module,
                        enum_or_type.group(3),
                        enum_or_type.group(2).casefold(),
                        (enum_or_type.group(1) or "public").casefold(),
                        line_num,
                    )
                )

    return symbols


def _symbol(
    module: VBAModuleSource,
    name: str,
    kind: str,
    visibility: str,
    line_num: int,
) -> VBASymbol:
    return VBASymbol(
        name=name,
        kind=kind,
        visibility=visibility,
        module_name=module.name,
        module_path=module.rel_path,
        module_kind=module.module_kind,
        line_num=line_num,
    )


def _declaration_visibility(statement: str) -> str:
    match = re.match(r"^(Public|Private|Friend|Global|Dim|Static|Const)\b", statement, re.IGNORECASE)
    if not match:
        return "private"
    access = match.group(1).casefold()
    if access in {"dim", "static", "const"}:
        return "private"
    return access


def _procedure_visibility(statement: str) -> str:
    match = re.match(r"^(Public|Private|Friend)\s+", statement, re.IGNORECASE)
    if match:
        return match.group(1).casefold()
    # VBA procedures default to Public; only standard-module public names are
    # subsequently admitted into the unqualified project namespace.
    return "public"


def _procedure_kind(statement: str) -> str:
    lowered = statement.casefold().lstrip()
    for prefix in ("public ", "private ", "friend ", "static "):
        if lowered.startswith(prefix):
            lowered = lowered[len(prefix):]
            break
    if lowered.startswith("function "):
        return "function"
    if lowered.startswith("property "):
        return "property"
    return "sub"


def _path_key(path: str) -> str:
    return path.replace("\\", "/").casefold()
