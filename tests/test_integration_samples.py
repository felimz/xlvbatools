"""
Integration tests utilizing sample workbooks placed in the sample_workbooks/ directory.
These tests run dynamically if any sample workbooks are present in the directory.
"""

import os
import glob
import pytest
from xlvbatools.vba.extractor import extract_all
from xlvbatools.analysis.preflight import run_all_rules
from xlvbatools.vba.dependency import build_call_graph

# Find all xlsm/xlsb/xls files in sample_workbooks/ directory
SAMPLE_WORKBOOKS_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "sample_workbooks"
)


def get_sample_workbooks():
    if not os.path.exists(SAMPLE_WORKBOOKS_DIR):
        return []
    patterns = [
        os.path.join(SAMPLE_WORKBOOKS_DIR, "*.xlsm"),
        os.path.join(SAMPLE_WORKBOOKS_DIR, "*.xlsb"),
        os.path.join(SAMPLE_WORKBOOKS_DIR, "*.xls"),
    ]
    files = []
    for pat in patterns:
        for f in glob.glob(pat):
            # Ignore Excel temporary lock/owner files
            if not os.path.basename(f).startswith("~$"):
                files.append(f)
    return files


sample_files = get_sample_workbooks()

# Auto-skip all tests in this module if no sample workbooks are present
pytestmark = pytest.mark.skipif(
    not sample_files,
    reason="No sample workbooks found in sample_workbooks/ directory"
)


@pytest.fixture(scope="module", params=sample_files, ids=lambda x: os.path.basename(x))
def extracted_sample(request, tmp_path_factory):
    """Fixture to extract a sample workbook to a temporary directory."""
    wb_path = request.param
    basename = os.path.splitext(os.path.basename(wb_path))[0]
    out_dir = tmp_path_factory.mktemp(f"extracted_{basename}")
    
    if os.name != "nt":
        pytest.skip("Excel COM automation requires Windows")
        
    try:
        manifest = extract_all(wb_path, str(out_dir))
        assert manifest, f"Failed to extract VBA from {wb_path}"
    except Exception as e:
        pytest.fail(f"Exception during extraction of {wb_path}: {e}")
        
    return {
        "workbook": wb_path,
        "source_dir": out_dir
    }


@pytest.mark.com
@pytest.mark.e2e
def test_sample_linter(extracted_sample):
    """Run the offline linter on the extracted code and verify it completes without raising exceptions."""
    source_dir = extracted_sample["source_dir"]
    vba_files = []
    for root, _, files in os.walk(source_dir):
        for f in files:
            if f.endswith((".bas", ".cls", ".frm")):
                vba_files.append(os.path.join(root, f))
                
    issues_count = 0
    for fpath in vba_files:
        rel_path = os.path.relpath(fpath, source_dir)
        with open(fpath, "r", encoding="utf-8", errors="ignore") as f:
            lines = f.readlines()
        try:
            issues = run_all_rules(rel_path, lines)
            issues_count += len(issues)
        except Exception as e:
            pytest.fail(f"Linter crashed on file {rel_path} from {extracted_sample['workbook']}: {e}")
            
    print(f"\nLinter finished successfully. Found {issues_count} issue(s) in {len(vba_files)} file(s).")


@pytest.mark.com
@pytest.mark.e2e
def test_sample_call_graph(extracted_sample):
    """Build a call dependency graph from the extracted code and verify it constructs successfully."""
    source_dir = extracted_sample["source_dir"]
    try:
        cg = build_call_graph(str(source_dir))
        graph_dict = cg.to_dict()
        assert isinstance(graph_dict, dict)
    except Exception as e:
        pytest.fail(f"CallGraph construction crashed on {extracted_sample['workbook']}: {e}")
