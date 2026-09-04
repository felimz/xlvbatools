"""
Tests for xlvbatools.vba.search -- VBA source file search.
"""

import pytest


@pytest.mark.unit
class TestVBASearch:
    """Test search across VBA source files."""

    def test_literal_search(self, sample_bas_file, temp_vba_source):
        from xlvbatools.vba.search import search_vba
        results = search_vba(str(temp_vba_source), "Debug.Print")
        assert len(results) == 1
        assert results[0].line_num == 7
        assert results[0].module == "modTest"

    @pytest.mark.parametrize(
        ("pattern", "options", "expected"),
        [
            ("debug.print", {}, 1),
            ("debug.print", {"case_sensitive": True}, 0),
            (r"Dim\s+\w+\s+As\s+Double", {"regex": True}, 1),
            ("NONEXISTENT_PATTERN_XYZ", {}, 0),
            ("[invalid", {"regex": True}, 0),
        ],
        ids=(
            "case-insensitive", "case-sensitive", "regex", "no-match",
            "invalid-regex",
        ),
    )
    def test_search_modes(
        self, sample_bas_file, temp_vba_source, pattern, options, expected,
    ):
        from xlvbatools.vba.search import search_vba

        results = search_vba(str(temp_vba_source), pattern, **options)
        assert len(results) == expected

    def test_search_summary(self, sample_bas_file, temp_vba_source):
        from xlvbatools.vba.search import search_vba_summary
        summary = search_vba_summary(str(temp_vba_source), "Sub|Function", regex=True)
        assert summary["total_matches"] >= 2
        assert len(summary["files"]) >= 1

    def test_search_cls_files(self, sample_cls_file, temp_vba_source):
        from xlvbatools.vba.search import search_vba
        results = search_vba(str(temp_vba_source), "Property")
        assert len(results) >= 2  # Get and Let

    def test_nonexistent_dir(self):
        from xlvbatools.vba.search import search_vba
        results = search_vba("/nonexistent/path", "test")
        assert results == []
