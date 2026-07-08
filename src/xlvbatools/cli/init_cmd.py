"""
xlvba init -- Project scaffolding command
==========================================
Creates an xlvbatools.toml configuration file and optional directory structure
in the current working directory.
"""

import os
import sys
import shutil
import importlib.resources


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


def run_init(args):
    """Execute the `xlvba init` command."""
    config_path = os.path.join(os.getcwd(), "xlvbatools.toml")

    if os.path.exists(config_path) and not getattr(args, "force", False):
        print(f"xlvbatools.toml already exists at {config_path}")
        print("Use --force to overwrite")
        sys.exit(1)

    # Determine workbook path
    workbook = args.workbook or _find_workbook()

    # Write config file
    content = TEMPLATE_TOML.format(workbook=workbook.replace("\\", "/"))
    with open(config_path, "w", encoding="utf-8") as f:
        f.write(content)
    print(f"Created: {config_path}")

    # Create standard directories
    for dirname in ["vba_source/modules", "vba_source/classes", "vba_source/sheets",
                    "snapshots", "logs"]:
        path = os.path.join(os.getcwd(), dirname)
        os.makedirs(path, exist_ok=True)
        print(f"Created: {dirname}/")

    # Optionally install .agents/ template
    if args.agents:
        _install_agents_template()

    print(f"\nProject initialized. Edit xlvbatools.toml to configure your workbook path.")
    if workbook == "workbook.xlsm":
        print("  Hint: set 'workbook' to your actual .xlsm file path.")


def _find_workbook() -> str:
    """Look for a .xlsm file in the current directory."""
    for f in os.listdir("."):
        if f.endswith(".xlsm"):
            return f
    return "workbook.xlsm"


def _install_agents_template():
    """Copy the .agents/ template into the current project."""
    agents_dir = os.path.join(os.getcwd(), ".agents")
    if os.path.exists(agents_dir):
        print("  .agents/ already exists, skipping template installation")
        return

    try:
        templates_path = importlib.resources.files("xlvbatools").joinpath("templates/agents")
        
        def copy_resource_dir(resource_path, dest_dir):
            os.makedirs(dest_dir, exist_ok=True)
            for entry in resource_path.iterdir():
                dest_path = os.path.join(dest_dir, entry.name)
                if entry.name == "__pycache__":
                    continue
                if entry.is_dir():
                    copy_resource_dir(entry, dest_path)
                elif entry.is_file():
                    with entry.open("rb") as sf:
                        with open(dest_path, "wb") as df:
                            shutil.copyfileobj(sf, df)

        copy_resource_dir(templates_path, agents_dir)
        print("  Created .agents/ template successfully from package data.")
    except Exception as e:
        print(f"  Error installing agents template: {e}")
