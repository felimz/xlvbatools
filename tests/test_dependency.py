"""
Tests for xlvbatools.vba.dependency -- Call graph analyzer.
"""

import pytest


@pytest.mark.unit
class TestCallGraph:
    """Test call graph construction."""

    def test_build_from_fixture(self, temp_vba_source):
        from xlvbatools.vba.dependency import build_call_graph

        # Create two modules with cross-calls
        mod_a = temp_vba_source / "modules" / "modA.bas"
        mod_a.write_text(
            "Public Sub Main()\n    Call Helper\nEnd Sub\n"
            "Public Sub Helper()\n    Debug.Print 1\nEnd Sub\n",
            encoding="utf-8",
        )

        mod_b = temp_vba_source / "modules" / "modB.bas"
        mod_b.write_text(
            "Public Sub Init()\n    Call Main\nEnd Sub\n",
            encoding="utf-8",
        )

        graph = build_call_graph(str(temp_vba_source))

        assert len(graph.modules) >= 2
        assert len(graph.procedures) >= 3

        # Check edge from Main -> Helper
        edges_from_main = [(a, b) for a, b in graph.edges if "Main" in a]
        assert any("Helper" in b for _, b in edges_from_main)

    def test_render_mermaid(self, temp_vba_source):
        from xlvbatools.vba.dependency import build_call_graph, render_mermaid

        mod = temp_vba_source / "modules" / "modTest.bas"
        mod.write_text(
            "Public Sub A()\n    Call B\nEnd Sub\n"
            "Public Sub B()\nEnd Sub\n",
            encoding="utf-8",
        )

        graph = build_call_graph(str(temp_vba_source))
        mermaid = render_mermaid(graph)
        assert "graph TD" in mermaid
        assert "subgraph" in mermaid

    def test_render_dot(self, temp_vba_source):
        from xlvbatools.vba.dependency import build_call_graph, render_dot

        mod = temp_vba_source / "modules" / "modTest.bas"
        mod.write_text("Public Sub A()\nEnd Sub\n", encoding="utf-8")

        graph = build_call_graph(str(temp_vba_source))
        dot = render_dot(graph)
        assert "digraph" in dot
        assert "cluster_" in dot

    def test_to_dict(self, temp_vba_source):
        from xlvbatools.vba.dependency import build_call_graph

        mod = temp_vba_source / "modules" / "modTest.bas"
        mod.write_text(
            "Public Sub A()\n    Call B\nEnd Sub\n"
            "Private Function B() As Long\n    B = 42\nEnd Function\n",
            encoding="utf-8",
        )

        graph = build_call_graph(str(temp_vba_source))
        d = graph.to_dict()
        assert "procedures" in d
        assert "edges" in d
        assert d["node_count"] >= 2

    def test_empty_directory(self, tmp_path):
        from xlvbatools.vba.dependency import build_call_graph
        graph = build_call_graph(str(tmp_path))
        assert len(graph.procedures) == 0
        assert len(graph.edges) == 0
