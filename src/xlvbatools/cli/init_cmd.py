"""
xlvba init -- Project scaffolding command
==========================================
Creates an xlvbatools.toml configuration file and optional directory structure
in the current working directory.
"""

import importlib.resources
import os
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Any


TEMPLATE_TOML = '''[xlvbatools]
workbook = "{workbook}"
vba_source = "vba_source"
snapshots_dir = "snapshots"
log_dir = "logs"
log_name = "xlvbatools"

[xlvbatools.snapshots]
rolling_limit = 10

[xlvbatools.backups]
limit = 5

[xlvbatools.lint]
protected_sheets = []
disabled_rules = []
'''


@dataclass(frozen=True)
class InitOutput:
    """Structured result of project initialization."""

    config_path: str
    workbook: str
    directories: tuple[str, ...]
    agents_status: str
    used_default_workbook: bool


@dataclass(frozen=True)
class AgentTemplateInstallOutput:
    """Files affected by a non-destructive agent-template installation."""

    destination: str
    installed: tuple[str, ...]
    skipped: tuple[str, ...]
    overwritten: tuple[str, ...]

    @property
    def status(self) -> str:
        if self.overwritten:
            return "updated"
        if self.installed:
            return "installed"
        return "skipped_existing"


def run_init(args) -> InitOutput:
    """Initialize a project without writing presentation text to stdout."""
    config_path = os.path.join(os.getcwd(), "xlvbatools.toml")

    if os.path.exists(config_path) and not getattr(args, "force", False):
        raise FileExistsError(
            f"xlvbatools.toml already exists at {config_path}; use --force to overwrite"
        )

    # Determine workbook path
    workbook = args.workbook or _find_workbook()

    # Write config file
    content = TEMPLATE_TOML.format(workbook=workbook.replace("\\", "/"))
    with open(config_path, "w", encoding="utf-8") as f:
        f.write(content)

    # Create standard directories
    directories = (
        "vba_source/modules",
        "vba_source/classes",
        "vba_source/sheets",
        "snapshots",
        "logs",
    )
    for dirname in directories:
        path = os.path.join(os.getcwd(), dirname)
        os.makedirs(path, exist_ok=True)

    # Optionally install .agents/ template
    agents_status = "not_requested"
    if args.agents:
        agents_status = install_agents_template().status

    return InitOutput(
        config_path=os.path.abspath(config_path),
        workbook=workbook,
        directories=tuple(os.path.abspath(path) for path in directories),
        agents_status=agents_status,
        used_default_workbook=workbook == "workbook.xlsm",
    )


def _find_workbook() -> str:
    """Look for a .xlsm file in the current directory."""
    for f in os.listdir("."):
        if f.endswith(".xlsm"):
            return f
    return "workbook.xlsm"


def install_agents_template(
    destination: str | os.PathLike[str] = ".agents",
    *,
    force: bool = False,
) -> AgentTemplateInstallOutput:
    """Copy packaged guidance without deleting project-specific agent files."""
    destination_path = Path(destination).resolve()
    installed: list[str] = []
    skipped: list[str] = []
    overwritten: list[str] = []
    try:
        templates_path = importlib.resources.files("xlvbatools").joinpath("templates/agents")

        def copy_resource_dir(resource_path: Any, relative: Path) -> None:
            for entry in resource_path.iterdir():
                if entry.name == "__pycache__":
                    continue
                entry_relative = relative / entry.name
                dest_path = destination_path / entry_relative
                if entry.is_dir():
                    copy_resource_dir(entry, entry_relative)
                elif entry.is_file():
                    relative_text = entry_relative.as_posix()
                    if dest_path.exists() and not force:
                        skipped.append(relative_text)
                        continue
                    existed = dest_path.exists()
                    dest_path.parent.mkdir(parents=True, exist_ok=True)
                    with entry.open("rb") as sf:
                        with dest_path.open("wb") as df:
                            shutil.copyfileobj(sf, df)
                    (overwritten if existed else installed).append(relative_text)

        copy_resource_dir(templates_path, Path())
        return AgentTemplateInstallOutput(
            destination=str(destination_path),
            installed=tuple(sorted(installed)),
            skipped=tuple(sorted(skipped)),
            overwritten=tuple(sorted(overwritten)),
        )
    except Exception as e:
        raise RuntimeError(f"Could not install .agents/ templates: {e}") from e
