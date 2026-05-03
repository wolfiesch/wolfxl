# RFC-069 — Combination charts: multi-family `<plotArea>` + secondary value axis (Sprint 4 / G15)

> **Status**: Proposed
> **Owner**: Claude (S4 design)
> **Sprint**: S4 — Charts beyond Singletons
> **Closes**: G15 (combination charts) in the openpyxl parity program
> **Depends on**: RFC-046 (chart `to_rust_dict()` shape)
> **Unblocks**: G16 (pivot-chart per-point overrides) — same emit module
> **Note on numbering**: an earlier draft session referred to this RFC as RFC-068. The tracker (`Plans/openpyxl-parity-program.md`) allocates RFC-068 to G08 (threaded comments, S2) and RFC-069 to G15. Going with the tracker.

## 1. Goal

Make `BarChart += LineChart` (and any other openpyxl-shaped chart-family composition) actually emit a combination chart whose `xl/charts/chartN.xml` has multiple chart-family elements inside `<plotArea>`, sharing a category axis and supporting an optional secondary value axis. Round-trip through openpyxl's reader so the compat-oracle probe `charts_combination` flips from xfail to pass.

The Sprint 0 hardening of that probe proved the existing path silently drops the secondary chart family: openpyxl's reload sees only one chart kind in the saved file.

## 2. Problem statement

The S0 trace pinpoints two breakages in the chart save path:

1. **Python serialiser drops the secondary chart family.** `ChartBase.__iadd__` appends the second chart to `self._charts: list[ChartBase] = [self]` (`python/wolfxl/chart/_chart.py:281–285`). But `ChartBase.to_rust_dict()` (`:363–430`) ignores `self._charts` entirely; it serialises only `self.kind`, `self.ser`, and `self.{x,y,z}_axis`. The line series and `<lineChart>` metadata never reach the Rust side.
2. **Rust emitter emits exactly one chart family.** `emit_plot_chart()` (`crates/wolfxl-writer/src/emit/charts/plot.rs:9–11`) writes a single `<c:{elem}>` element from `chart.kind`. The caller (`crates/wolfxl-writer/src/emit/charts.rs:101–120`) wraps it inside `<c:plotArea>` with a single shape header, then closes the plot area. There is no loop over chart families and no concept of secondary value axes.

Per ECMA-376 Part 1 §21.2.2.27, `<plotArea>` is a sequence type that admits one or more chart-family children (`<barChart>`, `<lineChart>`, `<scatterChart>`, ...). Each chart-family element references its axes by `<c:axId val="..."/>` children. Multiple families can share a category axis (same `axId`) while owning their own value axes. When a secondary value axis is present, it appears as a second `<c:valAx>` sibling of the primary one, typically with `<c:crosses val="max"/>` to render on the right side of the plot.

## 3. Public contract

The Python API stays openpyxl-shaped:

```python
bar = BarChart()
line = LineChart()
bar.add_data(Reference(...), titles_from_data=True)
line.add_data(Reference(...), titles_from_data=True)

# Default: line shares both axes with bar.
bar += line

# Secondary axis: caller signals intent by mutating the line chart's y_axis.
line.y_axis.crosses = "max"
line.y_axis.axId = 200
bar += line

ws.add_chart(bar, "E2")
wb.save(out)
```

After save, openpyxl's `load_workbook(out)` must see both `BarChart` and `LineChart` instances on `ws._charts`. (The probe asserts this exact predicate.)

Out of scope (deliberate): chart-anchored callouts, mixed 2D + 3D family combinations, more than three axes (no `valAx[3]`). These can ship later if a real user need surfaces.

## 4. Python design

### 4.1 `to_rust_dict()` extension

In `python/wolfxl/chart/_chart.py:363–430`, after the existing primary dict is built, iterate `self._charts[1:]` (the first entry is `self`) and serialise each as a sibling-chart dict, then attach to the payload under a new key:

```python
secondary_dicts = [
    secondary.to_rust_dict()
    for secondary in self._charts[1:]
]
if secondary_dicts:
    d["secondary_charts"] = secondary_dicts
```

Each secondary entry is a fully-formed chart dict (same shape as the primary) so the Rust side does not need to learn a half-shape. A secondary chart's `anchor`, `width_emu`, `height_emu`, `title`, `legend`, and outer-frame keys are ignored by the emitter — only the per-family fields (`kind`, `series_type`, `series`, type-specific keys, `y_axis`) are consumed.

### 4.2 Axis-id intent

The primary chart's `x_axis` and `y_axis` carry their `axId` (auto-allocated today). For each secondary chart:

- If `secondary.y_axis.axId == primary.y_axis.axId` (or unset → defaults to the same), the secondary shares the primary value axis.
- If `secondary.y_axis.axId` is distinct, the secondary owns its own value axis. If `secondary.y_axis.crosses == "max"`, the emitter renders it as the *secondary* axis (right side of plot).

`x_axis` (category axis) is always shared across all chart families in a single chartspace. wolfxl will refuse to emit two distinct category axes in one plotArea (fail loudly with a `ValueError` in `_validate_at_emit()` if a secondary chart's `x_axis.axId` differs from the primary).

### 4.3 No mutation of `self._charts`

`__iadd__` already does the right thing (`self._charts.append(other)`). We do not change its semantics. We do change `_validate_at_emit()` to walk the tree and assert each secondary is a permitted family (any non-Pie 2D chart kind for v1.0; Pie/Doughnut combinations are out of scope and raise).

## 5. Rust model + emit changes

### 5.1 `crates/wolfxl-writer/src/model/chart.rs`

Add to `Chart`:

```rust
pub struct Chart {
    // ... existing fields ...
    /// Sibling chart families to emit inside the same <plotArea>. Each entry
    /// is a fully-formed Chart whose outer-frame fields (title, legend, anchor)
    /// are ignored by the combination-chart emit path.
    pub secondary_charts: Vec<Chart>,
}
```

Default = empty vec; back-compat for every existing single-family caller.

### 5.2 `crates/wolfxl-writer/src/emit/charts.rs:101–120`

Replace the single-family emit with a multi-family emit:

```rust
out.push_str("<c:plotArea>");
if let Some(layout) = &chart.layout {
    emit_layout(&mut out, layout);
}

let primary_axes = (ax_id_a, ax_id_b);
emit_plot_chart(&mut out, chart, primary_axes.0, primary_axes.1);

let mut secondary_value_axes: Vec<&ValueAxis> = Vec::new();
for secondary in &chart.secondary_charts {
    let secondary_axes = secondary_axis_pair(
        secondary,
        primary_axes,
        &mut secondary_value_axes,
    );
    emit_plot_chart(&mut out, secondary, secondary_axes.0, secondary_axes.1);
}

if !chart.kind.is_axis_free() {
    if let Some(x) = &chart.x_axis { emit_axis(&mut out, x); }
    if let Some(y) = &chart.y_axis { emit_axis(&mut out, y); }
    for sv in &secondary_value_axes { emit_axis_as_secondary(&mut out, sv); }
}
out.push_str("</c:plotArea>");
```

`secondary_axis_pair()` returns `(cat_id, val_id)` for the secondary family:
- `cat_id = primary_axes.0` (always shared).
- `val_id = secondary.y_axis.ax_id` if distinct from primary; otherwise `primary_axes.1`.
- If a new `val_id` is allocated, push the secondary's `&ValueAxis` into `secondary_value_axes` so the emit phase writes a second `<c:valAx>` sibling.

`emit_axis_as_secondary()` is `emit_axis()` with one tweak: it forces `<c:crosses val="max"/>` when not explicitly set, and forces `<c:axPos val="r"/>`. This matches Excel's convention for secondary axes.

### 5.3 Validation

A secondary chart family that re-uses the primary `axId` for *both* axes is degenerate (it would render on top of the primary with no visual distinction). The Python side's `_validate_at_emit()` already rejects empty series; extend it to refuse a secondary chart whose `(x_axis.axId, y_axis.axId)` exactly equals the primary's `(x_axis.axId, y_axis.axId)` *and* the secondary has the same `kind` as the primary (likely a copy-paste bug). Fail loudly; do not silently dedupe.

## 6. OOXML reference (target shape)

For the probe input, the saved `xl/charts/chart1.xml` must contain (ECMA-376 Part 1 §21.2):

```xml
<c:chartSpace>
  <c:chart>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:varyColors val="1"/>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>...</c:ser>           <!-- bar series -->
        <c:axId val="10"/>
        <c:axId val="20"/>
      </c:barChart>
      <c:lineChart>
        <c:grouping val="standard"/>
        <c:ser>...</c:ser>           <!-- line series -->
        <c:axId val="10"/>           <!-- shared cat axis -->
        <c:axId val="200"/>          <!-- secondary val axis -->
      </c:lineChart>
      <c:catAx>
        <c:axId val="10"/>
        <c:axPos val="b"/>
        <c:crossAx val="20"/>
      </c:catAx>
      <c:valAx>
        <c:axId val="20"/>
        <c:axPos val="l"/>
        <c:crossAx val="10"/>
      </c:valAx>
      <c:valAx>
        <c:axId val="200"/>
        <c:axPos val="r"/>
        <c:crosses val="max"/>
        <c:crossAx val="10"/>
      </c:valAx>
    </c:plotArea>
  </c:chart>
</c:chartSpace>
```

This is the literal target the openpyxl reader recognises as a combination chart with secondary axis. The `<c:crossAx>` references inside each axis must point at the *other* axis on that family's `axId` pair.

## 7. Acceptance criteria

1. `tests/test_openpyxl_compat_oracle.py::test_compat_oracle_probe[charts.combination-...]` flips from xfail to passed.
2. New focused test `tests/test_combination_charts.py` covers:
   - Bar + Line, shared axes (no secondary).
   - Bar + Line, secondary y-axis with `crosses="max"`.
   - Bar + Line + Area (3-family). Document failure mode if not supported in this RFC; otherwise cover it.
   - Validation: refuse two families with identical kind + identical axIds.
3. `cargo test -p wolfxl-writer` chart tests stay green; new emit-side test asserts the multi-family `<plotArea>` shape matches §6.
4. `tests/test_external_oracle_preservation.py` with combo-chart fixtures stays green (LibreOffice + Excelize round-trip).
5. Compat-oracle pass count rises by exactly 1 (G15 closure).
6. No regression in the 30 existing passed probes; no regression in the 19 remaining xfailed (which are non-G15 gaps).
7. Compat-matrix row `charts.combination` flips from `not_yet` (gap_id `G15`) to `supported`. The G15 status row in the parity-program tracker is marked `landed` only after all six gates above pass.

## 8. Out-of-scope

- Combo charts containing 3D families. Excel's 3D combo support is fragile; defer.
- Pivot chart combos. G16 (pivot-chart per-point overrides) is the next-door RFC; combo on a pivot source is two RFCs out.
- Chartsheets (`<chartsheet>` standalone). Existing chartsheet path is single-family today; combo on a chartsheet is a separate gap.
- Surface, Bubble, Stock combinations. None are in the probe; none have user demand evidence yet.

## 9. Risks

| # | Risk | Mitigation |
|---|------|-----------|
| 1 | LibreOffice and Excel render `<axPos val="r"/>` differently for the secondary axis. | External-oracle preservation test pack covers both renderers; if outputs diverge, follow Excel's convention since it is the openpyxl reference. |
| 2 | Existing single-family chart tests share the `Chart` struct; adding `secondary_charts: Vec<Chart>` may break their construction. | Vec defaults to empty; struct construction sites use struct-update syntax or `..Default::default()`. Verify all build sites compile before landing. |
| 3 | `axId` collisions across multiple combos in one workbook. | Axis-id allocator at the workbook level (`workbook._chart_axis_allocator` or equivalent — confirm during impl) must be unique per workbook, not per chart. The handoff/impl pod must verify. |
| 4 | Pyright / mypy stubs for `BarChart += LineChart` may need explicit `__iadd__` return-type widening if subclasses care. | Keep `__iadd__: ChartBase -> ChartBase` as it is. The `_charts` list field already exists. |

## 10. Implementation plan (tentative)

1. RFC review (this document).
2. Codex handoff spec at `Plans/rfcs/handoffs/G15-combination-charts.md` derived from §4 + §5 — see if the work is small enough to delegate after the design is locked.
3. Pod work on `feat/parity-G15-combination-charts` worktree.
4. Acceptance gate per §7.
5. Mark G15 `landed`; flip spec; regenerate matrix; commit.

## 11. Open questions

- Should `secondary_charts` payload allow recursion (a secondary chart with its own `secondary_charts`)? Default proposal: **no**. Flat list of siblings under the primary; reject nested.
- For Pie + Bar combos (a niche openpyxl feature): does the probe care? Default proposal: **defer** — no probe today; revisit when one shows up.
- `axId` reservation strategy: today's allocator is per-chart. Secondary axes need workbook-unique IDs to avoid collisions across multiple combo charts. Confirm allocator scope during impl; possibly a separate small RFC if the allocator needs surgery.
