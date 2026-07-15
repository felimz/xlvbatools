"""
Tests for xlvbatools.snapshot.manager -- Snapshot system.
"""

import pytest


@pytest.mark.unit
class TestSnapshotStore:
    """Test snapshot create/list/restore/prune lifecycle."""

    def _make_manager(self, tmp_path):
        from xlvbatools.snapshot.manager import _SnapshotStore

        # Create a fake workbook file
        wb = tmp_path / "test.xlsm"
        wb.write_bytes(b"PK\x03\x04fake_workbook_content")

        # Create a fake vba_source directory
        vba = tmp_path / "vba_source" / "modules"
        vba.mkdir(parents=True)
        (vba / "modTest.bas").write_text("Public Sub Test()\nEnd Sub\n", encoding="utf-8")

        snaps = tmp_path / "snapshots"
        return _SnapshotStore(
            workbook_path=str(wb),
            vba_source_dir=str(tmp_path / "vba_source"),
            snapshots_dir=str(snaps),
            rolling_limit=3,
        )

    def test_create_and_list(self, tmp_path):
        mgr = self._make_manager(tmp_path)
        sid = mgr.create(description="test snapshot")
        assert sid  # Non-empty string
        snapshots = mgr.list()
        assert len(snapshots) == 1
        assert snapshots[0]["description"] == "test snapshot"
        assert snapshots[0]["snapshot_id"] == sid

    def test_info(self, tmp_path):
        mgr = self._make_manager(tmp_path)
        sid = mgr.create(description="info test")
        info = mgr.info(sid)
        assert info is not None
        assert info["snapshot_id"] == sid

    def test_info_latest(self, tmp_path):
        import time
        mgr = self._make_manager(tmp_path)
        mgr.create(description="first")
        time.sleep(1.1)  # Ensure different timestamp
        sid2 = mgr.create(description="second")
        info = mgr.info("latest")
        assert info["snapshot_id"] == sid2

    def test_restore(self, tmp_path):
        mgr = self._make_manager(tmp_path)
        sid = mgr.create(description="before change")

        # Modify the workbook
        wb_path = tmp_path / "test.xlsm"
        wb_path.write_bytes(b"MODIFIED_CONTENT")

        # Restore
        success = mgr.restore(sid, safety_snapshot=False)
        assert success

        # Verify workbook was restored
        content = wb_path.read_bytes()
        assert content == b"PK\x03\x04fake_workbook_content"

    def test_auto_prune(self, tmp_path):
        import time
        mgr = self._make_manager(tmp_path)  # rolling_limit=3

        for i in range(5):
            mgr.create(description=f"snapshot {i}")
            time.sleep(1.1)

        snapshots = mgr.list()
        non_milestones = [s for s in snapshots if not s.get("milestone")]
        assert len(non_milestones) <= 3

    def test_milestone_not_pruned(self, tmp_path):
        import time
        mgr = self._make_manager(tmp_path)  # rolling_limit=3

        mgr.create(description="milestone", milestone=True)
        time.sleep(1.1)

        for i in range(5):
            mgr.create(description=f"rolling {i}")
            time.sleep(1.1)

        snapshots = mgr.list()
        milestones = [s for s in snapshots if s.get("milestone")]
        assert len(milestones) == 1

    def test_prune_manual(self, tmp_path):
        import time
        mgr = self._make_manager(tmp_path)
        for i in range(5):
            mgr.create(description=f"snap {i}")
            time.sleep(1.1)

        pruned = mgr.prune(keep=2)
        assert pruned > 0
        assert len(mgr.list()) == 2

    def test_diff(self, tmp_path):
        mgr = self._make_manager(tmp_path)
        sid = mgr.create(description="baseline")

        # Modify a VBA file
        mod_file = tmp_path / "vba_source" / "modules" / "modTest.bas"
        mod_file.write_text("Public Sub Test()\n    ' Modified\nEnd Sub\n", encoding="utf-8")

        result = mgr.diff(sid)
        assert isinstance(result, str)

    def test_not_found(self, tmp_path):
        mgr = self._make_manager(tmp_path)
        assert mgr.info("nonexistent") is None
        assert mgr.restore("nonexistent") is False

    def test_physical_directories_and_zip(self, tmp_path):
        mgr = self._make_manager(tmp_path)
        # Create rolling snapshot
        mgr.create(description="Rolling Snapshot Example", milestone=False)
        # Create milestone snapshot
        sid_m = mgr.create(description="Milestone Snapshot Example", milestone=True)

        # Check physical existence of files in separate directory structures
        rolling_dir = tmp_path / "snapshots" / "rolling"
        milestones_dir = tmp_path / "snapshots" / "milestones"

        assert rolling_dir.is_dir()
        assert milestones_dir.is_dir()

        # Filename should include the slug suffix tag
        assert any("rolling-snapshot-example" in f.name for f in rolling_dir.iterdir() if f.name.endswith(".xlsm"))
        assert any("milestone-snapshot-example" in f.name for f in milestones_dir.iterdir() if f.name.endswith(".xlsm"))

        # VBA source should be archived in a ZIP file and the raw folder removed
        assert any("rolling-snapshot-example" in f.name and f.name.endswith(".zip") for f in rolling_dir.iterdir())
        assert not any("rolling-snapshot-example" in f.name and f.is_dir() for f in rolling_dir.iterdir())

        # Test restore of zipped archive
        vba_file = tmp_path / "vba_source" / "modules" / "modTest.bas"
        vba_file.write_text("Public Sub Test()\n    ' Changed current state\nEnd Sub\n", encoding="utf-8")

        # Restore milestone (which was created when vba_file had original content)
        assert mgr.restore(sid_m, safety_snapshot=False)
        assert "Changed current state" not in vba_file.read_text(encoding="utf-8")

    def test_git_metadata(self, tmp_path):
        mgr = self._make_manager(tmp_path)
        sid = mgr.create(description="test git info")
        info = mgr.info(sid)
        assert "git" in info
        if info["git"] is not None:
            assert "branch" in info["git"]
            assert "commit" in info["git"]
            assert "dirty" in info["git"]



@pytest.mark.unit
class TestManifest:
    """Test manifest read/write."""

    def test_round_trip(self, tmp_path):
        from xlvbatools.vba.manifest import Manifest, ComponentInfo

        m = Manifest(
            workbook="test.xlsm",
            extracted_at="2026-01-01T00:00:00",
            components=[
                ComponentInfo(
                    name="modTest",
                    type_code=1,
                    type_name="standard_module",
                    file="modules/modTest.bas",
                    line_count=10,
                    sha256="abc123",
                ),
            ],
        )
        path = str(tmp_path / "manifest.json")
        m.save(path)

        loaded = Manifest.load(path)
        assert len(loaded.components) == 1
        assert loaded.components[0].name == "modTest"
        assert loaded.workbook == "test.xlsm"

    def test_find_component(self):
        from xlvbatools.vba.manifest import Manifest, ComponentInfo
        m = Manifest(components=[
            ComponentInfo("modA", 1, "standard_module", "modules/modA.bas"),
            ComponentInfo("modB", 1, "standard_module", "modules/modB.bas"),
        ])
        assert m.find("moda") is not None
        assert m.find("modB") is not None
        assert m.find("modC") is None

    def test_load_missing_file(self, tmp_path):
        from xlvbatools.vba.manifest import Manifest
        m = Manifest.load(str(tmp_path / "nonexistent.json"))
        assert len(m.components) == 0
