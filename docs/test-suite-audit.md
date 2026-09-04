# Test suite inventory and LFL audit

This audit maps every test module to its lowest feasible layer (LFL). The
machine-enforced primary tiers are `unit`, `integration`, `excel`,
`distribution`, and `external`; `smoke` and `stress` are scheduling labels.
For this Windows library, live Excel tests are the application-level E2E tier.
There is no browser, HTTP service, database, Playwright, or Cypress surface.

The September 2026 review inspected every logical test definition and every
collected parameter case. It found no stale xlvbatools imports, API routes, or
UI selectors. The two URL-like literals in tests are an OOXML namespace and a
fake package-provenance URL, not application routes. The fast suite now checks
that every symbol imported from `xlvbatools` by a test still resolves.

## Inventory and decisions

| Test module | Logical tests | Current/LFL layer | Decision |
| --- | ---: | --- | --- |
| `test_cli.py` | 28 | Unit | Keep command parsing/presentation offline; delete duplicate live CLI-to-worker case. |
| `test_compat.py` | 4 | Unit | Keep platform boundary checks. |
| `test_config.py` | 7 | Unit | Consolidate invalid schema boundaries into a named parameter matrix. |
| `test_dependency.py` | 6 | Unit | Keep parser/rendering equivalence classes. |
| `test_differ.py` | 7 | Unit | Keep token-aware and raw comparison boundaries; replaces live diff permutation. |
| `test_distribution.py` | 1 | Distribution integration | Keep one fresh-wheel consumer contract. |
| `test_documentation.py` | 10 | Unit | Keep package/documentation consistency contracts. |
| `test_dumper.py` | 19 | Unit + Excel E2E | Keep rich-text and rendering logic offline; retain two unique live fidelity cases. |
| `test_execution.py` | 16 | Unit | Keep retry, timeout-budget, protocol, and failure-classification matrices offline. |
| `test_external_workbook.py` | 1 | Opt-in external E2E | Keep isolated from library CI and require an explicit workbook path. |
| `test_formatter.py` | 9 | Unit | Keep syntax/format/file boundaries. |
| `test_lint_filtering.py` | 8 | Unit | Keep filtering, baseline, multiset, and fail-closed boundaries. |
| `test_logging.py` | 4 | Unit | Keep local filesystem behavior. |
| `test_macro_runner.py` | 6 | Unit + Excel E2E | Delete duplicate success and timeout permutations; retain error, modal UI, and unrelated-process safety. |
| `test_preflight.py` | 78 | Unit | Keep rule-specific equivalence classes; separate names give better rule failure localization than one giant matrix. |
| `test_process.py` | 3 | Unit | Delete environment-dependent boolean tautology; retain exact-PID behavior and implementation guard. |
| `test_project_context.py` | 5 | Unit | Keep cross-module symbol-resolution boundaries. |
| `test_project.py` | 17 | Unit + Excel E2E | Delete duplicate inspect/run/diff/modify live paths; retain round-trip, startup safety, and compile contracts. |
| `test_repository_hygiene.py` | 5 | Unit | Batch ignore queries and add stale-import resolution gate. |
| `test_results.py` | 4 | Unit | Keep public serialization/error/cleanup contracts. |
| `test_search.py` | 5 | Unit | Consolidate search modes into a named parameter matrix. |
| `test_sequential_com.py` | 4 | Scheduled Excel stress | Keep out of PR CI; retain repeated native-lifecycle evidence. |
| `test_session.py` | 18 | Unit + Excel E2E | Keep exact-project compile, error-location, COM ownership, and teardown cases; these cannot be lowered faithfully. |
| `test_snapshot.py` | 15 | Unit | Keep internal store lifecycle and corruption boundaries. |
| `test_snapshots_public.py` | 4 | Unit | Keep public translation/immutability boundary; not redundant with store internals. |
| `test_test_runner.py` | 5 | Unit | Keep tier-routing safety contract. |
| `test_version.py` | 2 | Unit | Keep single-source and installed-provenance contracts. |
| `test_watchdog.py` | 20 | Unit | Consolidate dialog classification; keep PID-scoping and compile-controller behavior offline. |
| `test_worker.py` | 7 | Unit + offline integration | Retain two real subprocess protocol tests; mocked dispatch/creation behavior remains unit-level. |
| `test_workflow.py` | 11 | Unit | Keep typed public workflow and in-session primitive boundaries. |
| `test_workflow_live.py` | 6 | Excel E2E + scheduled stress | Delete duplicate CLI E2E; retain transaction, screenshot, failure, timeout, and repetition behavior. |
| `test_workflow_worker.py` | 6 | Unit | Keep session orchestration behind a fake session; no external boundary is claimed. |

## Executed migrations

Eight live tests were removed only after their distinct assertions were mapped
to lower-layer coverage:

- CLI named-range/save/visibility forwarding remains in `test_cli.py`; live
  worker execution is covered by Project and workflow acceptance.
- Basic macro success remains covered by public workflow and round-trip tests;
  the runner retains its distinct runtime-error and dialog cases.
- Infinite-loop cleanup remains covered by the workflow timeout and
  unrelated-Excel ownership tests.
- Basic Project inspect and macro cleanup remain covered by the combined live
  inspection test, workflow golden path, and session lifecycle tests.
- VBA case/spacing equivalence remains in the token-aware unit matrix; Project
  forwarding of comparison mode is separately unit-tested.
- Modify-then-inspect behavior remains covered in one-session workflow E2E and
  offline Project normalization tests.
- CLI workflow parsing/output remains fully offline; the live workflow test
  now tests Excel transaction semantics rather than repeating the CLI adapter.

The before baseline was 399 collected tests, 30 non-stress Excel tests, and a
388.67-second Excel acceptance run on the review machine. The refactor collects
391 tests and 22 non-stress Excel tests. The full acceptance rerun passed in
329.36 seconds, a 15.3% reduction despite two unusually slow fixture builds.
The offline tier improved from 7.75 to 6.93 seconds (10.6%). The PR smoke set is
three deliberately orthogonal cases using only the lightweight minimal
workbook: session startup and graceful close, public whole-project compile, and
combined data/screenshot inspection. It passed in 25.67 seconds; the prior
smoke cases consumed approximately 95.97 seconds from the same baseline timing
components, so routine live feedback is about 73% faster.

The collection hook enforces this design: smoke cases may use only the minimal
workbook fixture, cannot also be stress tests, and must remain in the Excel
tier. A future test cannot silently reintroduce the heavyweight macro or
compile-error fixture into the PR gate.

## CI scheduling

- Every pull request and push to `main` runs offline quality,
  unit/integration coverage, and the wheel contract. Feature branches do not
  also launch a duplicate push workflow while a PR is open.
- The three-case Excel smoke job joins PR/main CI only when an Excel runner is
  registered and `XLVBA_EXCEL_CI_ENABLED=true`; otherwise it skips immediately
  rather than blocking the PR in a queue.
- Nightly runs execute full non-stress Excel acceptance when native CI is
  enabled. Manual dispatch is always available for an intentionally active
  runner.
- The repeated stress tier remains opt-in on manual dispatch because it proves
  long-horizon lifecycle reliability rather than ordinary PR correctness.
- External consumer workbooks are never discovered automatically and remain
  the downstream project's responsibility.

This is an LFL policy, not a blanket preference for mocks. A test stays live
when Excel, COM ownership, VBE compilation, modal UI, rendering, or process
teardown is the behavior under test.
