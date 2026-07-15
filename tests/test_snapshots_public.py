"""Public snapshot-service contract tests."""

from dataclasses import FrozenInstanceError

import pytest


def _service(tmp_path):
    from xlvbatools import SnapshotService

    workbook = tmp_path / "book.xlsm"
    workbook.write_bytes(b"workbook")
    source = tmp_path / "vba_source" / "modules"
    source.mkdir(parents=True)
    (source / "modMain.bas").write_text(
        "Public Sub Main()\nEnd Sub\n",
        encoding="utf-8",
    )
    return SnapshotService(
        workbook,
        tmp_path / "vba_source",
        tmp_path / "snapshots",
        rolling_limit=3,
    )


@pytest.mark.unit
def test_snapshot_service_returns_immutable_records(tmp_path):
    service = _service(tmp_path)

    record = service.create("baseline")

    assert service.info(record) == record
    assert service.list() == (record,)
    with pytest.raises(FrozenInstanceError):
        record.description = "changed"

    payload = record.to_dict()
    payload["description"] = "changed"
    assert record.description == "baseline"


@pytest.mark.unit
def test_snapshot_service_restores_records_and_raises_typed_not_found(tmp_path):
    from xlvbatools import SnapshotNotFoundError

    service = _service(tmp_path)
    record = service.create("baseline")
    workbook = tmp_path / "book.xlsm"
    workbook.write_bytes(b"changed")

    restored = service.restore(record, safety_snapshot=False)

    assert restored == record
    assert workbook.read_bytes() == b"workbook"
    with pytest.raises(SnapshotNotFoundError):
        service.restore("missing", safety_snapshot=False)


@pytest.mark.unit
def test_snapshot_service_validates_prune_limit(tmp_path):
    service = _service(tmp_path)

    with pytest.raises(ValueError, match="at least one"):
        service.prune(keep=0)


@pytest.mark.unit
def test_snapshot_service_translates_corrupt_metadata_to_public_error(tmp_path):
    from xlvbatools import SnapshotError

    service = _service(tmp_path)
    snapshots = tmp_path / "snapshots"
    snapshots.mkdir()
    (snapshots / "snapshot_log.json").write_text("not json", encoding="utf-8")

    with pytest.raises(SnapshotError, match="Could not list snapshot data"):
        service.list()
