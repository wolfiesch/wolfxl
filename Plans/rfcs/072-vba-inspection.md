# RFC-072 — VBA inspection: read-only `Workbook.vba_archive` accessor (Sprint 6 / G19)

> **Status**: Proposed
> **Owner**: Claude (S6 design)
> **Sprint**: S6 — External Links + VBA Inspection
> **Closes**: G19 (VBA inspection API, read-only) in the openpyxl parity program

## 1. Goal

Make `Workbook.vba_archive` return the actual VBA project bytes (when present) instead of always `None`, matching openpyxl's read-only inspection contract.

```python
wb = wolfxl.load_workbook("macro.xlsm", keep_vba=True)  # or any modify=True load
archive = wb.vba_archive
assert archive is not None                # bytes-like
assert b"vbaProject" in archive or len(archive) > 0
```

Today `python/wolfxl/_workbook.py:306-308` returns `None` unconditionally. `vba.preserve` is already `supported` (modify-save round-trips `xl/vbaProject.bin` byte-for-byte), so the underlying bytes are reachable; G19 only needs to plumb them.

## 2. Problem statement

1. `Workbook.vba_archive` is a stub returning `None`.
2. There is no oracle probe for VBA inspection — `vba.inspect` has `gap_id: G19` and no `probe` field. **The implementation must register a new probe** as part of the closure.
3. Modify-mode workbook holds the bytes through `XlsxPatcher::file_passthroughs` (existing). Write-mode workbook (`wolfxl.Workbook()`) has no VBA — that's correct since `keep_vba=True` is meaningless without a source file.

## 3. Public contract

```python
wb.vba_archive  # → bytes | None
```

- Returns the raw `xl/vbaProject.bin` bytes when present.
- Returns `None` when the workbook is xlsx (no VBA).
- Read-only — there is no setter in v1.0. Authoring is `G28`, decision-gated, separate sprint.

The choice of `bytes` over `zipfile.ZipFile` is deliberate:

- openpyxl returns a `ZipFile` because their internal model holds the source as a ZipFile reader. wolfxl's modify-mode patcher operates on raw bytes; a `bytes` return is honest.
- A user wanting structured inspection can do `zipfile.ZipFile(io.BytesIO(wb.vba_archive))` themselves.
- A future RFC can wrap the bytes in a richer accessor (`vba_archive.modules`, `vba_archive.signatures`) if real demand surfaces.

If this proves to be a parity issue with openpyxl users, a follow-up RFC adds a `_vba_zip` cached property returning `zipfile.ZipFile(io.BytesIO(self.vba_archive))`.

## 4. Reader design

On modify-mode load (`load_workbook(path, modify=True)`):

1. The patcher already captures `xl/vbaProject.bin` (and any `xl/vbaProjectSignature.bin`) under `file_passthroughs` for `.xlsm` files. No code change here.
2. Add a new PyO3 method on the patcher: `get_vba_archive_bytes(&self) -> Option<Vec<u8>>` that returns the captured `xl/vbaProject.bin` bytes when present.
3. Replace the stub `Workbook.vba_archive` property with a real one that calls into the patcher (modify-mode only) and returns `None` for write-mode workbooks.

For non-modify reads (`load_workbook(path)` without `modify=True`), the bytes are not currently retained. v1.0 scope: `vba_archive` returns `None` for those reads. Document the limitation; if users need it, a follow-up RFC can wire write-mode-load preservation.

Actually re-evaluating: the cleaner v1.0 contract is "modify-mode load returns bytes, other loads return None." That's narrow but honest. Read-only / data-only loads do not need to allocate a copy of the bin file just for inspection.

## 5. Save path

No changes. `xl/vbaProject.bin` already round-trips on modify-save; that's what the existing `vba.preserve = supported` row asserts. G19 only adds an inspection accessor.

## 6. Test plan

### 6.1 New oracle probe

Register `@_register("vba_inspect")` in `tests/test_openpyxl_compat_oracle.py`. Probe contract:

```python
@_register("vba_inspect")
def _probe_vba_inspect(tmp_path: Path) -> None:
    """Read-only VBA archive inspection. Tracked under G19 (S6).

    Loads a fixture .xlsm and asserts wb.vba_archive surfaces the
    underlying xl/vbaProject.bin bytes.
    """
    import wolfxl

    fixture = Path(__file__).parent / "fixtures" / "macro_basic.xlsm"
    if not fixture.exists():
        pytest.skip("vba inspection fixture not vendored")
    wb = wolfxl.load_workbook(fixture, modify=True)
    archive = wb.vba_archive
    assert archive is not None, "wb.vba_archive must surface bytes for .xlsm"
    assert isinstance(archive, (bytes, bytearray, memoryview))
    assert len(archive) > 0
```

If no fixture exists, the impl pod creates one. The minimal fixture: an `.xlsm` saved from Excel with a one-line `Sub Test()\nEnd Sub` module. ~2KB of bytes. Vendor under `tests/fixtures/macro_basic.xlsm`.

### 6.2 New focused tests

`tests/test_vba_inspect.py` (new):

- **xlsx returns None:** plain `wb = wolfxl.load_workbook("plain.xlsx", modify=True); assert wb.vba_archive is None`.
- **xlsm returns bytes:** macro fixture; assert bytes-like, non-empty.
- **Round-trip:** load xlsm modify-mode, save, reload, assert `vba_archive` still bytes-like (and ideally byte-identical to original — verify or document).
- **Write-mode:** `wb = wolfxl.Workbook(); assert wb.vba_archive is None`.

### 6.3 No cargo tests needed

Pure plumbing on existing Rust state — no new parser logic.

## 7. Acceptance criteria

1. New probe `vba_inspect` registered AND passes (oracle pass count rises by 1; total probes rise by 1).
2. New focused tests in §6.2 all pass.
3. `cargo test --workspace` stays green.
4. Compat-oracle pass count rises by 1.
5. Compat-matrix row `vba.inspect` flips `not_yet` → `supported`. Tracker row `landed`.
6. Spec entry adds `probe: "vba_inspect"`.

## 8. Out-of-scope (deferred to G28 / S11)

- Authoring new VBA modules.
- Modifying existing VBA modules.
- Returning a `ZipFile` instead of raw bytes (if real demand surfaces, follow-up RFC).
- Read-only inspection from `load_workbook(path)` (non-modify) — v1.0 only surfaces in modify-mode loads.

## 9. Risks

| # | Risk | Mitigation |
|---|------|-----------|
| 1 | Patcher's `file_passthroughs` may not capture `xl/vbaProject.bin` if the file isn't accessed during normal load. | Verify via grep + targeted test; if missing, add an explicit capture rule for the part. |
| 2 | Returning a `bytes` reference may force a clone on every property access. | Use `Cow<[u8]>` or memoize — first-access pays the clone, subsequent accesses are cheap. Premature opt; only optimise if benchmarks show it. |
| 3 | Probe fixture (`tests/fixtures/macro_basic.xlsm`) might not exist; need to create it. | Impl pod creates the fixture: a minimal Excel-authored .xlsm with one trivial Sub. ~2KB. Document the creation process in the test docstring so future fixture-rebuilds are reproducible. |

## 10. Implementation plan

1. RFC review.
2. Subagent handoff: register probe, create or vendor fixture, wire `Workbook.vba_archive` through patcher, focused tests.
3. Subagent verifies all six acceptance gates.
4. Central merge.
