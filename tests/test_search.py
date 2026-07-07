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

    def test_case_insensitive_default(self, sample_bas_file, temp_vba_source):
        from xlvbatools.vba.search import search_vba
        results = search_vba(str(temp_vba_source), "debug.print")
        assert len(results) == 1

    def test_case_sensitive(self, sample_bas_file, temp_vba_source):
        from xlvbatools.vba.search import search_vba
        results = search_vba(str(temp_vba_source), "debug.print", case_sensitive=True)
        assert len(results) == 0

    def test_regex_search(self, sample_bas_file, temp_vba_source):
        from xlvbatools.vba.search import search_vba
        results = search_vba(str(temp_vba_source), r"Dim\s+\w+\s+As\s+Double", regex=True)
        assert len(results) == 1

    def test_no_matches(self, sample_bas_file, temp_vba_source):
        from xlvbatools.vba.search import search_vba
        results = search_vba(str(temp_vba_source), "NONEXISTENT_PATTERN_XYZ")
        assert len(results) == 0

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

    def test_invalid_regex(self, sample_bas_file, temp_vba_source):
        from xlvbatools.vba.search import search_vba
        results = search_vba(str(temp_vba_source), "[invalid", regex=True)
        assert results == []
