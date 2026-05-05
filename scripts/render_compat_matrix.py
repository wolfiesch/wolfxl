"""Render docs/migration/compatibility-matrix.md from the spec module.

Run: ``python scripts/render_compat_matrix.py``

The spec at ``docs/migration/_compat_spec.py`` is the source of truth. Edit
it, then regenerate the markdown so the public matrix and the in-repo data
stay in sync. The matrix's ``Last rendered`` line records the most recent
regen date; the renderer also emits a totals block (supported / partial /
not_yet / out_of_scope) that the parity-program tracker references.
"""

from __future__ import annotations

import datetime as _dt
import importlib.util
import sys
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
SPEC_PATH = REPO_ROOT / "docs" / "migration" / "_compat_spec.py"
OUT_PATH = REPO_ROOT / "docs" / "migration" / "compatibility-matrix.md"

STATUS_DISPLAY = {
    "supported": "✅ Supported",
    "partial": "🟡 Partial",
    "not_yet": "❌ Not Yet",
    "out_of_scope": "⛔ Out of Scope",
}


def _load_spec_module():
    spec = importlib.util.spec_from_file_location("_compat_spec", SPEC_PATH)
    if spec is None or spec.loader is None:
        raise SystemExit(f"failed to import {SPEC_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _render_table(entries: list[dict]) -> str:
    lines = [
        "| openpyxl | wolfxl | Status | Gap | Notes |",
        "|---|---|---|---|---|",
    ]
    for entry in entries:
        status = STATUS_DISPLAY.get(entry["status"], entry["status"])
        gap = entry.get("gap_id") or ""
        notes = (entry.get("notes") or "").strip().replace("\n", " ")
        lines.append(
            f"| `{entry['openpyxl']}` | `{entry['wolfxl']}` | {status} | {gap} | {notes} |"
        )
    return "\n".join(lines)


def render(spec_module) -> str:
    today = _dt.date.today().isoformat()
    categories = spec_module.CATEGORIES
    grouped = spec_module.entries_by_category()
    totals = spec_module.status_totals()
    total_count = sum(totals.values())

    parts: list[str] = []
    parts.append("# Compatibility Matrix")
    parts.append("")
    parts.append(
        f"_Last rendered: **{today}** from "
        "[`docs/migration/_compat_spec.py`](_compat_spec.py). Do not edit "
        "this file by hand; run `python scripts/render_compat_matrix.py`._"
    )
    parts.append("")
    parts.append(
        "This page is the public scoreboard for wolfxl's openpyxl-API "
        "compatibility. Each row maps an openpyxl idiom to its wolfxl "
        "equivalent and the current implementation status. Gap IDs (e.g. "
        "`G11`) link to rows in [`Plans/openpyxl-parity-program.md`]"
        "(../../Plans/openpyxl-parity-program.md), where the work to close "
        "them is sequenced into sprints."
    )
    parts.append("")
    parts.append("## Status legend")
    parts.append("")
    parts.append("| Symbol | Meaning |")
    parts.append("|---|---|")
    parts.append("| ✅ Supported | Implemented and covered by tests / fixtures. |")
    parts.append("| 🟡 Partial | Common case works; edge cases tracked under a gap ID. |")
    parts.append("| ❌ Not Yet | Not implemented; tracked under a gap ID. |")
    parts.append("| ⛔ Out of Scope | Explicitly excluded from the roadmap. |")
    parts.append("")
    parts.append("## Totals")
    parts.append("")
    parts.append(f"- ✅ Supported: **{totals.get('supported', 0)}** / {total_count}")
    parts.append(f"- 🟡 Partial: **{totals.get('partial', 0)}** / {total_count}")
    parts.append(f"- ❌ Not Yet: **{totals.get('not_yet', 0)}** / {total_count}")
    parts.append(f"- ⛔ Out of Scope: **{totals.get('out_of_scope', 0)}** / {total_count}")
    parts.append("")

    for cat in categories:
        entries = grouped.get(cat["id"], [])
        if not entries:
            continue
        parts.append(f"## {cat['title']}")
        parts.append("")
        parts.append(_render_table(entries))
        parts.append("")

    parts.append("## How this page is maintained")
    parts.append("")
    parts.append(
        "This file is generated from `_compat_spec.py`. To update a row, "
        "edit the spec module and run:"
    )
    parts.append("")
    parts.append("```bash")
    parts.append("python scripts/render_compat_matrix.py")
    parts.append("```")
    parts.append("")
    parts.append(
        "The accompanying `tests/test_openpyxl_compat_oracle.py` harness "
        "imports `ENTRIES` from the same module and uses each entry's "
        "`probe` field to drive a live test. Adding a new probe means "
        "adding a new entry in the spec and a matching probe function in "
        "the harness."
    )
    parts.append("")
    return "\n".join(parts) + "\n"


def main() -> int:
    spec_module = _load_spec_module()
    rendered = render(spec_module)
    OUT_PATH.write_text(rendered)
    print(f"wrote {OUT_PATH.relative_to(REPO_ROOT)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
