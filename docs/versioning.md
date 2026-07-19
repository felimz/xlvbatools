# Versioning and releases

## Version domains

xlvbatools versions four contracts independently:

1. **Package version** — Semantic Versioning (`MAJOR.MINOR.PATCH`). The single
   source is `src/xlvbatools/_version.py`; setuptools reads that value when it
   builds a wheel.
2. **Result schema** — `RESULT_SCHEMA_VERSION` identifies serialized
   `OperationResult.to_dict()` payloads. It changes only when their machine-
   readable structure changes.
3. **Worker protocol** — `WORKER_PROTOCOL_VERSION` identifies the private
   parent/worker transport. It may evolve without changing the Python API.
4. **Workflow schema** — `WORKFLOW_SCHEMA_VERSION` identifies versioned
   `xlvba workflow` request files independently from result and transport
   schemas.

`xlvba version` reports all four identifiers plus installed-package
and VCS provenance.

## Public compatibility boundary

Only names listed in `xlvbatools.__all__` are public. Modules under `core`,
`vba`, `macro`, `workbook`, `analysis`, and `snapshot` implement the worker and
may change between minor releases. Applications should build their wrappers
around `Project` and `OperationResult`.

Starting with 1.0.0:

- incompatible public API changes require a major package version;
- backward-compatible functionality increments the minor version;
- backward-compatible fixes increment the patch version;
- result-schema and worker-protocol changes are called out in the changelog;
- supported public deprecations normally remain for one minor release before
  removal, except for unsafe behavior or security fixes.

The distribution metadata uses `Development Status :: 5 -
Production/Stable`. Pre-release work must use a PEP 440 suffix such as
`1.1.0rc1`; do not publish an unqualified stable version with beta metadata.

## Release procedure

1. Update `_version.py` and add the dated changelog entry.
2. Run unit, integration, sequential-COM, and real-workbook acceptance tests
   using the repository `.venv`.
3. Build the wheel with normal PEP 517 build isolation.
4. Install the wheel into a fresh virtual environment outside the source tree
   and verify the public API, package version, result schema, protocol, CLI
   discovery catalog, and packaged `.agents/` installation.
5. Commit the release, create an annotated `vMAJOR.MINOR.PATCH` Git tag, and
   push the commit and tag.
6. Publish the exact wheel built from that tagged commit and record its hash.

The release commit, tag, wheel metadata, `xlvba version`, and changelog
must all report the same package version. Do not derive the package version
from an editable checkout's branch name or local environment.

Downstream projects should prefer a pinned released wheel and artifact hash.
A full 40-character Git commit pin is appropriate for reviewing an unreleased
revision, but should be replaced with the released version after publication.
