# XL-18: graceful shutdown after successful Excel workflows

WA-OCEAN's repeated complete Excel and release profiles exposed successful
operations followed by forced termination of their owned Excel PID. The
operation ran once, no dialog was reported, and no process remained. Those
outcomes correctly failed `require_clean_shutdown()`.

## Shutdown defect

The previous sentinel-workbook path released COM and posted `WM_CLOSE` to
`Application.Hwnd`, then posted `WM_QUIT` to the captured thread if Excel was
still running. It never called `Application.Quit()`. If the HWND could no
longer be verified, it could skip even the window-close request.

Excel's SDI interface makes `Application.Hwnd` the active workbook window's
handle. Closing that window is not an explicit request to quit the automation
application. The old message sequence intermittently left the server alive
until the existing exact-PID termination deadline.

Simply calling `self.excel.Quit()` before releasing the application wrapper
was also insufficient: native live tests exposed `0x80010108` while finalizing
the wrapper after Excel disconnected. A successful pytest case status did not
make those native diagnostics acceptable.

## Corrected order

1. Suppress shutdown events and alerts, and keep a saved blank sentinel
   workbook open while closing and releasing the target workbook.
2. Retain the owned application's raw `IDispatch` interface and resolve `Quit`.
3. Release the Python application wrapper, its type-info proxies, and the
   sentinel's Python wrapper while Excel is still alive.
4. Invoke `Application.Quit` through the retained interface, then release it
   and the session-owned COM apartment.
5. Keep the owned-PID dialog watchdog active while waiting for the process
   handle to become signaled after natural exit.
   The existing forced-cleanup fallback remains an unclean result.

The public `Project` API, result schema, timeout budget, save behavior,
no-replay boundary, and strict cleanup assertions are unchanged. No global
process cleanup, registry changes, add-in changes, or outer retry is added.

Unit regressions require the quit request even without a usable HWND, assert
that Python workbook/application references are released before invoking it,
and preserve failed-quit diagnostics and exact-PID fallback. The existing
subprocess finalizer regression additionally requires graceful exit without
forced termination.

## References

- [Excel SDI behavior](https://learn.microsoft.com/en-us/office/vba/excel/concepts/programming-for-the-single-document-interface-in-excel)
- [Application.Quit](https://learn.microsoft.com/en-us/office/vba/api/excel.application.quit)
- [PyIDispatch.Invoke](https://mhammond.github.io/pywin32/PyIDispatch__Invoke_meth.html)
- [PyIDispatch.GetIDsOfNames](https://mhammond.github.io/pywin32/PyIDispatch__GetIDsOfNames_meth.html)

## Previous candidate qualification (superseded)

The runtime fix is commit `794088b5b066ce904cd494892c0264dbad2e091d`.
Windows testing used Python 3.14.4 and pywin32 312. The package version remains
1.2.3; this commit is a source candidate, not a new tagged release.

| Gate | Result | Retained evidence |
|---|---|---|
| Complete offline tier | 364 passed; 69.22% coverage | [JUnit](validation/xl18/upstream-offline.xml) |
| Fresh consumer wheel installation | 1 passed | [JUnit](validation/xl18/upstream-distribution.xml) |
| Complete live Excel tier | 22 passed | [JUnit](validation/xl18/upstream-excel.xml) |
| Complete lifecycle stress tier | 5 passed | [JUnit](validation/xl18/upstream-stress.xml) |

Each selected tier has zero failures, errors, and skips. The live and stress
logs contain no native COM finalizer signatures. Stress includes nested direct
COM cases, five rich-COM sessions, 50 public macro runs, repeated valid/invalid
compilation, and 25 dependent workflows. Deliberate timeout cases in the live
tier retain their exact-PID failure cleanup; ordinary operations still require
graceful shutdown.

One earlier stress run was interrupted by host Modern Standby from 19:27:21
to 19:32:57 on 2026-09-04 (America/Chicago). Its failed operation reported no
Excel PID and exceeded its 180-second budget during the host pause. That run
is excluded from acceptance. The complete stress tier above was rerun with a
temporary system/display keep-awake request and unchanged test timeouts.

## Process-exit race found by repeated downstream testing

The additional read-only audit rejected the second complete Excel profile on
candidate `794088b5b066ce904cd494892c0264dbad2e091d` (40 passed, 1 failed).
The exact process handle initially returned `WAIT_TIMEOUT`, while its exit
code was already 0 and its exit FILETIME was populated. The same retained
handle became signaled by the diagnostic observation 50 ms later. The audit
still failed the initial observation; no extra wait was used to accept it.

This is consistent with the documented
[ExitProcess sequence](https://learn.microsoft.com/en-us/windows/win32/api/processthreadsapi/nf-processthreadsapi-exitprocess):
Windows sets the exit status before signaling the process object. The old
`GetExitCodeProcess != STILL_ACTIVE` check could certify completion during
that gap. Candidate `6ecc57579f3ca4446f00ae2986eece75eefc19e7` checks
`WaitForSingleObject(handle, 0)` with `SYNCHRONIZE` access instead. Only
`WAIT_OBJECT_0` establishes completion; timeouts and unknown wait results
remain live. Existing grace deadlines and forced-cleanup rejection remain
unchanged. Regression cases cover signaled, pending, and failed waits.

The earlier integer-PID-only observation remains inconclusive because it
captured neither process identity nor exit-state information. It is not
claimed as proven PID reuse. Both failed profiles are retained and excluded
from acceptance. Qualification restarts with the final runtime candidate;
results for the superseded candidate above are historical only.

## Final candidate qualification

Runtime candidate: `6ecc57579f3ca4446f00ae2986eece75eefc19e7`.
Offline: 367 passed (69.23% coverage); distribution: 1 passed; native Excel:
22 passed. Ruff and mypy (18 source files) pass. Native log scans find no
COM finalizer signatures. All 5 lifecycle stress tests pass (636.267 seconds), with no native finalizer
signatures. Three consecutive complete WA-OCEAN Excel profiles pass 41 cases
each. Two consecutive complete downstream release profiles pass 289 cases each.
The source candidate is not a tagged release. Test runners temporarily inhibit
idle standby and release that request in a finally block. There is no outer
workflow retry or relaxed cleanup assertion.

### Host concurrency disclosure

This runner executes the downstream profiles sequentially, but it did not have
exclusive use of the host. The coordinated upstream task reported an old-main
all-tier suite overlapping the Excel profiles (390 passed in 886.08 seconds),
then an all-tier suite on runtime `6ecc575` with v1.2.4 metadata overlapping the
release profiles. These are separate owned-PID sessions. Results below report
observed passes and strict cleanup evidence, not exclusive-host timing or an
absence of other user/task-owned Excel instances.

## Retained final results

All gates below qualify runtime `6ecc57579f3ca4446f00ae2986eece75eefc19e7`.
WA-OCEAN was tested at `c10565b` with the exact Git dependency installed, not
an official v1.2.3 wheel. No runtime edits occurred during these runs.

| Gate | Passed | Seconds | JUnit |
|---|---:|---:|---|
| upstream-fast | 367 | 7.931 | [report](validation/xl18/xl18-complete-upstream-fast.xml) |
| upstream-distribution | 1 | 16.640 | [report](validation/xl18/xl18-complete-upstream-distribution.xml) |
| upstream-excel | 22 | 319.349 | [report](validation/xl18/xl18-complete-upstream-excel.xml) |
| upstream-stress | 5 | 636.267 | [report](validation/xl18/xl18-complete-upstream-stress.xml) |
| excel-1 | 41 | 330.582 | [report](validation/xl18/xl18-complete-excel-1.xml) |
| excel-2 | 41 | 335.513 | [report](validation/xl18/xl18-complete-excel-2.xml) |
| excel-3 | 41 | 337.968 | [report](validation/xl18/xl18-complete-excel-3.xml) |
| release-1 | 289 | 358.690 | [report](validation/xl18/xl18-complete-release-1.xml) |
| release-2 | 289 | 331.301 | [report](validation/xl18/xl18-complete-release-2.xml) |

All selected gates have zero failures, errors, and skips. The five downstream
profiles audit 26 unique public operations each, 130 total. Every audited
operation has graceful shutdown, no forced termination, a reaped worker, and
no residual owned PID. Each profile includes four intentional error-dialog
fixtures: negative module position, unreachable imperial CG shift, invalid
combination count, and duplicate VBA declaration. Those expected diagnostics
are preserved; unexpected dialogs are zero. Successful operations report no
dialogs. No cleanup assertion is relaxed for error fixtures.

The [manifest](validation/xl18/summary.json) records environment, source
provenance, counts, durations, and SHA-256 hashes. Each named JUnit report has
an adjacent full `.log`; downstream reports also have a `-cleanup.jsonl`.
The retained [audit](validation/xl18/cleanup_audit.py) delegates the public
strict assertion unchanged and independently probes only reported owned PIDs.
The earlier [excluded runs](validation/xl18/excluded-runs.json) remain visible.

Production and intentionally excluded legacy workbook SHA-256 hashes are
unchanged. No user workbook is included in these artifacts. This qualifies
the source fix; tagged-release packaging and subsequent consumer release
qualification are separate from these candidate results.
