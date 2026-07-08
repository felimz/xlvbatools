"""
Snapshot Manager
=================
Timestamped checkpoint/rollback system for workbook and VBA source state.

Every snapshot is identified by its ISO-8601 compact timestamp (YYYYMMDDTHHMMSS).
Supports dual-layer snapshots: git commits for VBA source + binary .xlsm copies.

Usage:
    from xlvbatools.snapshot import SnapshotManager

    mgr = SnapshotManager("workbook.xlsm", "vba_source/", "snapshots/")
    sid = mgr.create(description="before refactor")
    mgr.list()
    mgr.restore("latest")
    mgr.prune(keep=10)
"""

import datetime
import hashlib
import json
import logging
import os
import shutil
from typing import Optional

logger = logging.getLogger(__name__)

TS_FORMAT = "%Y%m%dT%H%M%S"


class SnapshotManager:
    """
    Manages timestamped snapshots for workbook and VBA source state.

    Parameters
    ----------
    workbook_path : str
        Path to the .xlsm workbook.
    vba_source_dir : str
        Path to the vba_source/ directory.
    snapshots_dir : str
        Directory where snapshots are stored.
    rolling_limit : int
        Max number of rolling (non-milestone) snapshots before auto-pruning.
    """

    def __init__(
        self,
        workbook_path: str,
        vba_source_dir: str,
        snapshots_dir: str,
        rolling_limit: int = 10,
    ):
        self.workbook_path = os.path.abspath(workbook_path)
        self.vba_source_dir = os.path.abspath(vba_source_dir)
        self.snapshots_dir = os.path.abspath(snapshots_dir)
        self.rolling_limit = rolling_limit
        self._log_path = os.path.join(self.snapshots_dir, "snapshot_log.json")
        self._lock_depth = 0

    def create(
        self,
        description: str = "",
        milestone: bool = False,
    ) -> str:
        """
        Create a new snapshot.

        Returns the snapshot ID (timestamp string).
        """
        snapshot_id = datetime.datetime.now().strftime(TS_FORMAT)
        while os.path.exists(os.path.join(self.snapshots_dir, f"{snapshot_id}.xlsm")):
            import time
            time.sleep(0.2)
            snapshot_id = datetime.datetime.now().strftime(TS_FORMAT)
        os.makedirs(self.snapshots_dir, exist_ok=True)

        # Copy workbook binary
        wb_filename = f"{snapshot_id}.xlsm"
        wb_snapshot = os.path.join(self.snapshots_dir, wb_filename)
        wb_size = 0
        if os.path.exists(self.workbook_path):
            shutil.copy2(self.workbook_path, wb_snapshot)
            wb_size = os.path.getsize(self.workbook_path)

        # Copy VBA source
        vba_dir_name = f"{snapshot_id}_vba"
        vba_snapshot = os.path.join(self.snapshots_dir, vba_dir_name)
        if os.path.isdir(self.vba_source_dir):
            shutil.copytree(self.vba_source_dir, vba_snapshot, dirs_exist_ok=True)

        # Build entry
        entry = {
            "snapshot_id": snapshot_id,
            "timestamp": datetime.datetime.now().isoformat(timespec="seconds"),
            "description": description,
            "workbook_file": wb_filename,
            "workbook_hash": self._hash_file(self.workbook_path),
            "workbook_size_bytes": wb_size,
            "vba_source_dir": vba_dir_name if os.path.isdir(vba_snapshot) else None,
            "vba_hash": self._hash_vba_source(),
            "milestone": milestone,
        }

        log = self._load_log()

        # Auto-prune rolling snapshots
        if not milestone:
            non_milestones = [e for e in log if not e.get("milestone", False)]
            if len(non_milestones) >= self.rolling_limit:
                prune_count = len(non_milestones) - self.rolling_limit + 1
                for p in non_milestones[:prune_count]:
                    self._delete_snapshot_files(p)
                    log.remove(p)

        log.append(entry)
        self._save_log(log)
        logger.info(f"Snapshot created: {snapshot_id} ({description or 'no description'})")
        return snapshot_id

    def list(self) -> list[dict]:
        """List all snapshots sorted chronologically."""
        return self._load_log()

    def info(self, identifier: str) -> dict | None:
        """Get details for a specific snapshot."""
        return self._find(identifier)

    def restore(self, identifier: str, safety_snapshot: bool = True) -> bool:
        """
        Restore workbook and VBA source from a snapshot.

        Parameters
        ----------
        identifier : str
            Snapshot ID, "latest", prefix, description substring, or index.
        safety_snapshot : bool
            Whether to create a safety snapshot of current state first.
        """
        entry = self._find(identifier)
        if entry is None:
            logger.error(f"Snapshot not found: {identifier}")
            return False

        wb_snapshot = os.path.join(self.snapshots_dir, entry["workbook_file"])
        if not os.path.exists(wb_snapshot):
            logger.error(f"Snapshot workbook not found: {wb_snapshot}")
            return False

        # Safety snapshot
        if safety_snapshot:
            self.create(description=f"auto-safety before restoring to {entry['snapshot_id']}")

        # Restore workbook
        shutil.copy2(wb_snapshot, self.workbook_path)
        logger.info(f"Workbook restored from: {entry['workbook_file']}")

        # Restore VBA source
        vba_dir = entry.get("vba_source_dir")
        if vba_dir:
            vba_snapshot = os.path.join(self.snapshots_dir, vba_dir)
            if os.path.isdir(vba_snapshot):
                for subdir in ("modules", "classes", "sheets"):
                    target = os.path.join(self.vba_source_dir, subdir)
                    if os.path.isdir(target):
                        shutil.rmtree(target)
                shutil.copytree(vba_snapshot, self.vba_source_dir, dirs_exist_ok=True)
                logger.info("VBA source restored from snapshot")

        return True

    def diff(self, identifier: str) -> str:
        """Show VBA file differences between a snapshot and current state."""
        import difflib

        entry = self._find(identifier)
        if entry is None:
            return f"Snapshot not found: {identifier}"

        vba_dir = entry.get("vba_source_dir")
        if not vba_dir:
            return "No VBA source in this snapshot"

        vba_snapshot = os.path.join(self.snapshots_dir, vba_dir)
        if not os.path.isdir(vba_snapshot):
            return "Snapshot VBA directory not found"

        return self._diff_directories(vba_snapshot, self.vba_source_dir)

    def prune(self, keep: int = 10) -> int:
        """Remove old rolling snapshots, keeping the most recent `keep`."""
        log = self._load_log()
        non_milestones = [e for e in log if not e.get("milestone", False)]

        if len(non_milestones) <= keep:
            return 0

        to_remove = non_milestones[:-keep]
        for entry in to_remove:
            self._delete_snapshot_files(entry)
            log.remove(entry)

        self._save_log(log)
        logger.info(f"Pruned {len(to_remove)} old snapshots")
        return len(to_remove)

    # ── Internal Helpers ──

    def _find(self, identifier: str) -> dict | None:
        """Flexible snapshot lookup."""
        log = self._load_log()
        if not log:
            return None

        if identifier.lower() == "latest":
            return log[-1]

        for e in log:
            if e["snapshot_id"] == identifier:
                return e

        prefix = [e for e in log if e["snapshot_id"].startswith(identifier)]
        if len(prefix) == 1:
            return prefix[0]

        desc = [e for e in log if identifier.lower() in e.get("description", "").lower()]
        if len(desc) == 1:
            return desc[0]

        try:
            idx = int(identifier)
            if 0 <= idx < len(log):
                return log[idx]
        except ValueError:
            pass

        return None

    def _lock(self):
        """Reentrant file lock context manager for snapshot operations."""
        import contextlib
        import time

        @contextlib.contextmanager
        def _inner_lock():
            lock_path = self._log_path + ".lock"
            if getattr(self, "_lock_depth", 0) > 0:
                self._lock_depth += 1
                try:
                    yield
                finally:
                    self._lock_depth -= 1
                return

            os.makedirs(self.snapshots_dir, exist_ok=True)
            acquired = False
            for _ in range(50):  # Retry up to 5 seconds
                try:
                    fd = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                    os.close(fd)
                    acquired = True
                    break
                except FileExistsError:
                    time.sleep(0.1)

            if not acquired:
                logger.warning("Could not acquire snapshot log lock; proceeding without lock.")

            self._lock_depth = 1
            try:
                yield
            finally:
                self._lock_depth = 0
                if acquired:
                    try:
                        os.remove(lock_path)
                    except Exception:
                        pass
        return _inner_lock()

    def _load_log(self) -> list:
        import contextlib
        # Inline ContextDecorator helper to use generator with custom lock
        with self._lock():
            if not os.path.exists(self._log_path):
                return []
            with open(self._log_path, "r", encoding="utf-8") as f:
                entries = json.load(f)
            return sorted(entries, key=lambda e: e.get("snapshot_id", ""))

    def _save_log(self, entries: list):
        with self._lock():
            os.makedirs(self.snapshots_dir, exist_ok=True)
            entries.sort(key=lambda e: e.get("snapshot_id", ""))
            with open(self._log_path, "w", encoding="utf-8") as f:
                json.dump(entries, f, indent=2)

    def _delete_snapshot_files(self, entry: dict):
        wb = os.path.join(self.snapshots_dir, entry["workbook_file"])
        if os.path.exists(wb):
            os.remove(wb)
        vba = entry.get("vba_source_dir")
        if vba:
            vba_path = os.path.join(self.snapshots_dir, vba)
            if os.path.isdir(vba_path):
                shutil.rmtree(vba_path)

    def _hash_file(self, filepath: str) -> str:
        if not os.path.exists(filepath):
            return ""
        h = hashlib.sha256()
        with open(filepath, "rb") as f:
            while chunk := f.read(8192):
                h.update(chunk)
        return h.hexdigest()[:16]

    def _hash_vba_source(self) -> str:
        h = hashlib.sha256()
        if os.path.isdir(self.vba_source_dir):
            for root, _, files in os.walk(self.vba_source_dir):
                for fname in sorted(files):
                    if fname.endswith((".bas", ".cls", ".json")):
                        with open(os.path.join(root, fname), "rb") as f:
                            h.update(f.read())
        return h.hexdigest()[:16]

    def _diff_directories(self, dir_a: str, dir_b: str) -> str:
        import difflib

        files_a, files_b = set(), set()
        for root, _, files in os.walk(dir_a):
            for f in files:
                if f.endswith((".bas", ".cls")):
                    files_a.add(os.path.relpath(os.path.join(root, f), dir_a))
        for root, _, files in os.walk(dir_b):
            for f in files:
                if f.endswith((".bas", ".cls")):
                    files_b.add(os.path.relpath(os.path.join(root, f), dir_b))

        output = []
        for fname in sorted(files_a | files_b):
            if fname not in files_a:
                output.append(f"  + ADDED: {fname}")
                continue
            if fname not in files_b:
                output.append(f"  - REMOVED: {fname}")
                continue
            from xlvbatools.vba._io import read_vba_lines
            a_lines = read_vba_lines(os.path.join(dir_a, fname))
            b_lines = read_vba_lines(os.path.join(dir_b, fname))
            if a_lines != b_lines:
                diff = difflib.unified_diff(a_lines, b_lines,
                                            fromfile=f"snapshot/{fname}",
                                            tofile=f"current/{fname}")
                output.append("".join(diff))

        return "\n".join(output) if output else "No differences found"
