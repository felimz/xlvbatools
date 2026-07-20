"""Installed package version and source provenance diagnostics."""

from __future__ import annotations

import json
import sys
from dataclasses import asdict, dataclass
from importlib import metadata
from pathlib import Path
from typing import Optional

from xlvbatools._version import __version__ as source_version
from xlvbatools.analysis.filtering import LINT_BASELINE_SCHEMA_VERSION
from xlvbatools.core.protocol import WORKER_PROTOCOL_VERSION
from xlvbatools.results import RESULT_SCHEMA_VERSION
from xlvbatools.workflow import WORKFLOW_SCHEMA_VERSION


@dataclass(frozen=True)
class VersionInfo:
    """Version, interpreter, install location, and VCS provenance."""

    version: str
    python_executable: str
    package_path: str
    source_url: Optional[str] = None
    commit_id: Optional[str] = None
    requested_revision: Optional[str] = None
    distribution_version: Optional[str] = None
    version_mismatch: bool = False
    result_schema_version: str = RESULT_SCHEMA_VERSION
    worker_protocol_version: str = WORKER_PROTOCOL_VERSION
    workflow_schema_version: str = WORKFLOW_SCHEMA_VERSION
    lint_baseline_schema_version: str = LINT_BASELINE_SCHEMA_VERSION

    def to_dict(self) -> dict:
        return asdict(self)


def get_version_info() -> VersionInfo:
    """Read package metadata without requiring Git or a source checkout."""
    package_path = str(Path(__file__).resolve().parent)
    try:
        distribution = metadata.distribution("xlvbatools")
        distribution_version = distribution.version
        direct_url_text = distribution.read_text("direct_url.json")
    except metadata.PackageNotFoundError:
        distribution_version = None
        direct_url_text = None

    source_url = None
    commit_id = None
    requested_revision = None
    if direct_url_text:
        try:
            direct_url = json.loads(direct_url_text)
            source_url = direct_url.get("url")
            vcs_info = direct_url.get("vcs_info") or {}
            commit_id = vcs_info.get("commit_id")
            requested_revision = vcs_info.get("requested_revision")
        except (TypeError, ValueError):
            pass

    return VersionInfo(
        version=source_version,
        python_executable=sys.executable,
        package_path=package_path,
        source_url=source_url,
        commit_id=commit_id,
        requested_revision=requested_revision,
        distribution_version=distribution_version,
        version_mismatch=(
            distribution_version is not None
            and distribution_version != source_version
        ),
    )
