# RFC-071 — External links: workbook-level collection + part preservation (Sprint 6 / G18)

> **Status**: Proposed
> **Owner**: Claude (S6 design)
> **Sprint**: S6 — External Links + VBA Inspection
> **Closes**: G18 (external links workbook-level collection + rels) in the openpyxl parity program

## 1. Goal

Make this work, end-to-end, no `xfail`:

```python
wb = wolfxl.Workbook()
wb.active["A1"] = "='[ext.xlsx]Sheet1'!$A$1"
out = tmp_path / "ext.xlsx"
wb.save(out)

wb2 = wolfxl.load_workbook(out)
links = wb2._external_links            # list[ExternalLink]
assert links is not None and len(links) >= 0
```

This is the verbatim contract of `tests/test_openpyxl_compat_oracle.py:823-841` (the `external_links_collection` probe). Closing it flips `external_links.workbook_collection` from `not_yet` (gap_id `G18`) to `supported`.

## 2. Problem statement

Surface scan found ZERO references to `externalLink` anywhere in `python/wolfxl/`, `src/wolfxl/`, or `crates/`. The probe currently `xfails` because:

1. There is no parser for `xl/externalLinks/externalLink{N}.xml` parts.
2. `Workbook._external_links` does not exist as a Python attribute.
3. The save path does not preserve external-link parts (the `vba.preserve` style claim in the spec notes is aspirational for external links — there is no patcher branch handling them).

## 3. Public contract

```python
class ExternalLink:
    file_link: ExternalFileLink   # ro: target workbook ref
    rid: str                       # ro: rels id "rId1"
    target: str                    # ro: linked filename, e.g. "ext.xlsx"
    sheet_names: list[str]         # ro: sheet names referenced (parsed from <sheetNames>)
    cached_data: dict[str, Any]    # ro: parsed <sheetDataSet> if present (else {})

class ExternalFileLink:
    target: str          # "ext.xlsx" — the linked workbook filename
    target_mode: str     # "External"
```

Workbook surface:

```python
wb._external_links  # list[ExternalLink]; empty list when no external refs
wb.external_links   # alias (no underscore) for openpyxl-shape compatibility
```

Both exposures point at the same backing list. The list is read-only in v1.0 — no `append` / `remove` API yet (deferred to a follow-up RFC).

## 4. Reader design

A minimal `quick-xml` parser (or string scan, given the bounded element set) on the load path:

For each rel in `xl/_rels/workbook.xml.rels` of type `…/relationships/externalLink`:

1. Resolve the target part (e.g. `xl/externalLinks/externalLink1.xml`).
2. Parse the part — extract `<externalBook r:id="...">` → `target_mode`, the linked workbook's rels target.
3. Walk `xl/externalLinks/_rels/externalLink{N}.xml.rels` for the `externalLinkPath` rel → that's the linked filename (`ext.xlsx`).
4. Parse `<sheetNames>/<sheetName val="..."/>` children into `sheet_names`.
5. Parse `<sheetDataSet>` (cached values, optional) into `cached_data` — keep this loose / dict-shaped in v1.0.

Build one `ExternalLink` per part; attach to `Workbook._external_links`. Done at load time, eagerly (not lazy — count is small).

Parser lives in: `src/wolfxl/external_links.rs` (new), PyO3-exported helpers `parse_external_link_part(xml: &[u8]) -> PyDict` and `parse_external_link_rels(xml: &[u8]) -> PyDict`. Python wrapping in `python/wolfxl/_external_links.py` (new).

## 5. Save path

External-link parts are **opaque preservation** in v1.0:

- On `load_workbook(modify=True)`, the patcher captures the existing `xl/externalLinks/externalLink{N}.xml` and `xl/externalLinks/_rels/externalLink{N}.xml.rels` bytes into `XlsxPatcher::file_passthroughs` (existing mechanism).
- On `save`, those files round-trip byte-for-byte. Workbook rels graph and content-types preserve their entries.
- Write-mode (`wolfxl.Workbook()`) creating a *new* external link is out of scope. The probe creates `wb.active["A1"] = "='[ext.xlsx]Sheet1'!$A$1"` with NO matching `xl/externalLinks/` part — that's a forward-reference formula with no cached link. wolfxl writes the formula string verbatim; on reload, no `xl/externalLinks/` parts exist and `_external_links` is `[]`. The probe asserts `len(links) >= 0` which is satisfied by an empty list.

This means the probe passes with: parser returning `[]` for files without external-link parts, AND parser returning a populated list for files that do. Both cases shipped together.

## 6. Test plan

### 6.1 Compat-oracle probe (existing)

`tests/test_openpyxl_compat_oracle.py:823-841` flips xfail → passed. No probe modification.

### 6.2 New focused tests

`tests/test_external_links.py` (new):

- **Empty case:** workbook with no external refs → `wb._external_links == []`.
- **Forward-ref formula:** the probe scenario verbatim; assert no `externalLinks/` parts and empty list.
- **Real external-link fixture:** load a hand-crafted xlsx with one external link; assert `len(_external_links) == 1`, `link.target == "ext.xlsx"`, `link.sheet_names == ["Sheet1"]`. Fixture lives at `tests/fixtures/external_links_basic.xlsx`.
- **Round-trip preserve:** load real fixture in `modify=True`, save to new path, reload, assert external links still present and shape-identical.
- **Alias:** `wb.external_links is wb._external_links`.

### 6.3 Cargo tests

`crates/wolfxl-reader/tests/external_links_parse.rs` (new): unit tests for the `parse_external_link_part` / `parse_external_link_rels` helpers on hand-crafted XML strings, including the no-rels-file edge case.

## 7. Acceptance criteria

1. `tests/test_openpyxl_compat_oracle.py::test_compat_oracle_probe[external_links.workbook_collection-...]` flips xfail → pass.
2. New focused tests in §6.2 all pass.
3. `cargo test --workspace` green; new parser unit tests pass.
4. Compat-oracle pass count rises by 1.
5. Compat-matrix row `external_links.workbook_collection` flips `not_yet` → `supported`. Tracker row `landed`.
6. README + `docs/trust/limitations.md` mention v1.0 scope: read-only inspection, opaque preservation on modify-save, no external-link authoring API.

## 8. Out-of-scope (deferred)

- Authoring new external links (`wb._external_links.append(...)`).
- Removing or rewriting existing external links.
- Following the link target (loading the linked workbook's data).
- Updating cached values in `<sheetDataSet>`.

## 9. Risks

| # | Risk | Mitigation |
|---|------|-----------|
| 1 | Modify-mode save may strip `xl/externalLinks/` parts because the patcher doesn't have a passthrough rule for them. | Add explicit passthrough enum entry in `XlsxPatcher` for `xl/externalLinks/**`; verify via fixture round-trip test §6.2. |
| 2 | Workbook rels graph may lose external-link rels on save. | The patcher already preserves unknown rels; verify via the round-trip test. If it fails, add an explicit preservation rule. |
| 3 | Formula `='[ext.xlsx]Sheet1'!$A$1` may not round-trip if the formula parser strips it. | The formula is a string — wolfxl stores cell formulas as strings on `cell.value`; round-trip is byte-level. Verify in the focused test. |

## 10. Implementation plan

1. RFC review (this document).
2. Subagent handoff: parser + `Workbook._external_links` accessor + load-time materialization + passthrough preservation + tests.
3. Subagent verifies all six acceptance gates.
4. Central merge into main; cleanup.
