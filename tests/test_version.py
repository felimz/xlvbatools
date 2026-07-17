"""Package version and Git provenance tests."""

import json
from unittest.mock import patch

import pytest


@pytest.mark.unit
def test_version_info_reads_git_commit_from_direct_url():
    from xlvbatools.version import get_version_info

    class Distribution:
        version = "1.2.3"

        @staticmethod
        def read_text(name):
            assert name == "direct_url.json"
            return json.dumps({
                "url": "https://github.com/example/xlvbatools.git",
                "vcs_info": {
                    "vcs": "git",
                    "commit_id": "abc123",
                    "requested_revision": "main",
                },
            })

    with patch("xlvbatools.version.metadata.distribution", return_value=Distribution()):
        info = get_version_info()

    assert info.version == "1.2.3"
    assert info.source_url == "https://github.com/example/xlvbatools.git"
    assert info.commit_id == "abc123"
    assert info.requested_revision == "main"
    assert info.result_schema_version == "1.1"
    assert info.worker_protocol_version == "2.0"


@pytest.mark.unit
def test_source_version_is_single_sourced():
    import xlvbatools
    from xlvbatools._version import __version__

    assert __version__.count(".") >= 2
    assert xlvbatools.__version__ == __version__
