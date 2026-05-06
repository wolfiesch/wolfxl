# Public Evidence Status

This page is the truth source for what WolfXL currently claims publicly.

## Status

- WolfXL package version: `2.0.0`
- Repository under audit: `SynthGL/wolfxl`
- Last local verification pass: pending refresh for the next release
- Fresh release-artifact benchmark snapshot: `ExcelBench/results-release-2026-04-28/`
- Benchmark posture: historical wheel-backed evidence is available; refresh before making next-release claims.

## Verified Now

These points are current and can be repeated publicly with links to the supporting docs or commands.

| Claim | Status | Evidence |
|---|---|---|
| WolfXL ships read, write, and modify modes behind an openpyxl-style API. | Publishable | `README.md`, `docs/getting-started/quickstart.md` |
| WolfXL 2.0 adds pivot cache construction, pivot table construction, pivot-chart linkage, and pivot-bearing `copy_worksheet`. | Publishable | `docs/release-notes-2.0.md` |
| WolfXL preserves untouched workbook parts during modify-mode saves, including charts, images, macros, and pivot-bearing files. | Publishable with limitations | `README.md`, `docs/trust/limitations.md`, local test suite |
| Local verification is green for the Rust workspace and Python package. | Publishable with command output, not as a headline | `cargo test --workspace`, `uv run --no-sync pytest -q` |
| ExcelBench contains reproducible fidelity and performance harnesses, generated artifacts, and methodology docs. | Publishable | `https://github.com/SynthGL/ExcelBench`, `ExcelBench/METHODOLOGY.md` |
| The fresh WolfXL 2.0 wheel-backed ExcelBench rerun reaches `18/18` green features and `100%` pass rate. | Publishable with timestamp | `ExcelBench/results-release-2026-04-28/README.md` |

## Gated Claims

These points need a fresh artifact or clearer caveat before they should appear as headline launch copy.

| Claim Type | Current State | Action Before Headline Use |
|---|---|---|
| Universal WolfXL 2.0 speedup headline vs `openpyxl` | Still workload-specific | Cite the dated release snapshot and case studies instead of one universal number |
| "Only library" or "first library" ecosystem claims | Narrowly plausible but still comparative | Keep qualified wording and link the compatibility matrix |
| Public benchmark tables in `docs/performance/benchmark-results.md` | Historical v1.7 baseline | Treat as archived context, not current release proof |
| Older ExcelBench dashboard snapshots | Accurate for their timestamp, not current release evidence | Label as archived snapshots in README links and launch copy |
| Next release quality claims | Pending current verification | Run the release evidence checklist after atomic-save, package-safety, and fidelity gates are green |

## Release Evidence Checklist

Use this set before any post, PyPI release thread, HN launch, or comparison blog post.

1. `cargo test --workspace`
2. `uv run --no-sync pytest -q`
3. `uv run pytest tests/parity -q -x`
4. Fresh wheel build and clean install smoke
5. Fresh ExcelBench fidelity rerun against the WolfXL 2.0 release artifact
6. Fresh ExcelBench perf rerun against the WolfXL 2.0 release artifact
7. Update public timestamps in both READMEs and any dashboard summary pages

## Public Wording Policy

- Safe: "Rust-backed, openpyxl-compatible Excel automation with surgical modify mode."
- Safe: "Pivot tables are constructible from Python in WolfXL 2.0."
- Safe with timestamp: "On the 2026-04-29 wheel-backed ExcelBench release snapshot..."
- Avoid: broad `10x-100x` release-headline speed claims that ignore workload shape.
- Avoid until refreshed: presenting historical v1.7 benchmark tables as WolfXL 2.0 release proof.

## Cross-Repo Alignment Notes

- WolfXL is on `2.0.0` in `Cargo.toml` and release docs.
- ExcelBench now contains a fresh WolfXL 2.0 wheel-backed release snapshot in `results-release-2026-04-28/` alongside older historical snapshots.
- When in doubt, prefer explicit dates, exact commands, and raw artifact links over compressed marketing summaries.
