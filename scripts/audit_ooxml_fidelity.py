#!/usr/bin/env python3
"""Audit OOXML package fidelity between two workbook files.

This is intentionally package-level rather than API-level. It catches the
class of modify-save regressions where a workbook still opens, but an OOXML
dependency has been dropped, orphaned, or left pointing at a missing part.
"""

from __future__ import annotations

import argparse
import json
import posixpath
import re
import sys
import zipfile
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree

REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
CT_NS = "{http://schemas.openxmlformats.org/package/2006/content-types}"
MAIN_NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"

FEATURE_PART_PREFIXES = {
    "chart": ("xl/charts/",),
    "chart_style": ("xl/charts/style", "xl/charts/colors"),
    "conditional_formatting": ("xl/worksheets/", "xl/styles.xml"),
    "drawing": ("xl/drawings/",),
    "external_link": ("xl/externalLinks/",),
    "pivot": ("xl/pivotCache/", "xl/pivotTables/", "pivotCache/"),
    "slicer": ("xl/slicers/", "xl/slicerCaches/"),
    "table": ("xl/tables/",),
    "vba": ("xl/vbaProject.bin",),
}


@dataclass(frozen=True)
class Relationship:
    rels_part: str
    rel_id: str
    rel_type: str
    target: str
    target_mode: str | None
    resolved_target: str | None


@dataclass
class Snapshot:
    path: str
    parts: set[str]
    content_overrides: dict[str, str]
    relationships: list[Relationship]
    dxfs_count: int
    cf_dxf_refs: list[tuple[str, int]]
    feature_parts: dict[str, list[str]]


def snapshot(path: Path) -> Snapshot:
    with zipfile.ZipFile(path) as archive:
        parts = set(archive.namelist())
        return Snapshot(
            path=str(path),
            parts=parts,
            content_overrides=_read_content_overrides(archive),
            relationships=_read_relationships(archive),
            dxfs_count=_read_dxfs_count(archive),
            cf_dxf_refs=_read_cf_dxf_refs(archive),
            feature_parts=_classify_feature_parts(parts),
        )


def audit(before: Path, after: Path) -> dict:
    before_snapshot = snapshot(before)
    after_snapshot = snapshot(after)
    issues: list[dict[str, str]] = []

    missing_parts = sorted(before_snapshot.parts - after_snapshot.parts)
    for part in missing_parts:
        issues.append(
            {
                "severity": "error",
                "kind": "missing_part",
                "part": part,
                "message": f"Part existed before save and is missing after save: {part}",
            }
        )

    _audit_relationship_preservation(before_snapshot, after_snapshot, issues)
    _audit_dangling_relationships(after_snapshot, issues)
    _audit_content_type_preservation(before_snapshot, after_snapshot, issues)
    _audit_conditional_formatting_refs(after_snapshot, issues)
    _audit_feature_hotspots(before_snapshot, after_snapshot, issues)

    return {
        "before": _snapshot_summary(before_snapshot),
        "after": _snapshot_summary(after_snapshot),
        "issue_count": len(issues),
        "issues": issues,
    }


def _audit_relationship_preservation(
    before: Snapshot, after: Snapshot, issues: list[dict[str, str]]
) -> None:
    before_rels = {_relationship_key(rel): rel for rel in before.relationships}
    after_rels = {_relationship_key(rel): rel for rel in after.relationships}
    for key, rel in sorted(before_rels.items()):
        if key not in after_rels:
            issues.append(
                {
                    "severity": "error",
                    "kind": "missing_relationship",
                    "part": rel.rels_part,
                    "message": (
                        "Relationship existed before save and is missing after save: "
                        f"{rel.rels_part} {rel.rel_id} {rel.rel_type} -> {rel.target}"
                    ),
                }
            )


def _audit_dangling_relationships(snapshot_: Snapshot, issues: list[dict[str, str]]) -> None:
    for rel in snapshot_.relationships:
        if rel.resolved_target is None:
            continue
        if rel.resolved_target not in snapshot_.parts:
            issues.append(
                {
                    "severity": "error",
                    "kind": "dangling_relationship",
                    "part": rel.rels_part,
                    "message": (
                        f"{rel.rels_part} {rel.rel_id} points to missing "
                        f"{rel.resolved_target}"
                    ),
                }
            )


def _audit_content_type_preservation(
    before: Snapshot, after: Snapshot, issues: list[dict[str, str]]
) -> None:
    for part, content_type in sorted(before.content_overrides.items()):
        if part not in after.parts:
            continue
        after_content_type = after.content_overrides.get(part)
        if after_content_type != content_type:
            issues.append(
                {
                    "severity": "error",
                    "kind": "content_type_changed",
                    "part": part,
                    "message": (
                        f"Content type for {part} changed from {content_type!r} "
                        f"to {after_content_type!r}"
                    ),
                }
            )


def _audit_conditional_formatting_refs(
    snapshot_: Snapshot, issues: list[dict[str, str]]
) -> None:
    for sheet_part, dxf_id in snapshot_.cf_dxf_refs:
        if dxf_id >= snapshot_.dxfs_count:
            issues.append(
                {
                    "severity": "error",
                    "kind": "conditional_formatting_dxf_out_of_range",
                    "part": sheet_part,
                    "message": (
                        f"{sheet_part} references dxfId={dxf_id}, but styles.xml "
                        f"only has {snapshot_.dxfs_count} <dxf> entries"
                    ),
                }
            )


def _audit_feature_hotspots(
    before: Snapshot, after: Snapshot, issues: list[dict[str, str]]
) -> None:
    for feature, before_parts in sorted(before.feature_parts.items()):
        if not before_parts:
            continue
        after_parts = set(after.feature_parts.get(feature, []))
        missing = sorted(set(before_parts) - after_parts)
        if missing:
            issues.append(
                {
                    "severity": "error",
                    "kind": "feature_part_loss",
                    "part": feature,
                    "message": f"{feature} parts disappeared after save: {missing}",
                }
            )


def _read_content_overrides(archive: zipfile.ZipFile) -> dict[str, str]:
    try:
        xml = archive.read("[Content_Types].xml")
    except KeyError:
        return {}
    root = ElementTree.fromstring(xml)
    overrides: dict[str, str] = {}
    for node in root.findall(f"{CT_NS}Override"):
        part_name = node.attrib.get("PartName", "").lstrip("/")
        content_type = node.attrib.get("ContentType")
        if part_name and content_type:
            overrides[part_name] = content_type
    return overrides


def _read_relationships(archive: zipfile.ZipFile) -> list[Relationship]:
    out: list[Relationship] = []
    for rels_part in sorted(p for p in archive.namelist() if p.endswith(".rels")):
        root = ElementTree.fromstring(archive.read(rels_part))
        seen_ids: set[str] = set()
        for node in root.findall(f"{REL_NS}Relationship"):
            rel_id = node.attrib.get("Id", "")
            rel_type = node.attrib.get("Type", "")
            target = node.attrib.get("Target", "")
            target_mode = node.attrib.get("TargetMode")
            resolved = _resolve_relationship_target(rels_part, target, target_mode)
            if rel_id in seen_ids:
                rel_id = f"{rel_id}#duplicate"
            seen_ids.add(rel_id)
            out.append(
                Relationship(
                    rels_part=rels_part,
                    rel_id=rel_id,
                    rel_type=rel_type,
                    target=target,
                    target_mode=target_mode,
                    resolved_target=resolved,
                )
            )
    return out


def _resolve_relationship_target(
    rels_part: str, target: str, target_mode: str | None
) -> str | None:
    if not target or target_mode == "External":
        return None
    if target.startswith("/"):
        return posixpath.normpath(target.lstrip("/"))

    source_part = _source_part_for_rels(rels_part)
    source_dir = posixpath.dirname(source_part)
    return posixpath.normpath(posixpath.join(source_dir, target))


def _source_part_for_rels(rels_part: str) -> str:
    if rels_part == "_rels/.rels":
        return ""
    prefix, name = rels_part.rsplit("/_rels/", 1)
    return posixpath.join(prefix, name.removesuffix(".rels"))


def _read_dxfs_count(archive: zipfile.ZipFile) -> int:
    try:
        root = ElementTree.fromstring(archive.read("xl/styles.xml"))
    except (KeyError, ElementTree.ParseError):
        return 0
    dxfs = root.find(f"{MAIN_NS}dxfs")
    if dxfs is None:
        return 0
    return len(dxfs.findall(f"{MAIN_NS}dxf"))


def _read_cf_dxf_refs(archive: zipfile.ZipFile) -> list[tuple[str, int]]:
    refs: list[tuple[str, int]] = []
    for part in sorted(_worksheet_parts(archive.namelist())):
        try:
            root = ElementTree.fromstring(archive.read(part))
        except ElementTree.ParseError:
            continue
        for cf_rule in root.findall(f".//{MAIN_NS}cfRule"):
            raw = cf_rule.attrib.get("dxfId")
            if raw is not None and raw.isdigit():
                refs.append((part, int(raw)))
    return refs


def _worksheet_parts(parts: Iterable[str]) -> Iterable[str]:
    pattern = re.compile(r"^xl/worksheets/sheet\d+\.xml$")
    return (part for part in parts if pattern.match(part))


def _classify_feature_parts(parts: set[str]) -> dict[str, list[str]]:
    classified: dict[str, list[str]] = {}
    for feature, prefixes in FEATURE_PART_PREFIXES.items():
        classified[feature] = sorted(
            part for part in parts if any(part.startswith(prefix) for prefix in prefixes)
        )
    return classified


def _relationship_key(rel: Relationship) -> tuple[str, str, str, str | None]:
    return (rel.rels_part, rel.rel_type, rel.target, rel.target_mode)


def _snapshot_summary(snapshot_: Snapshot) -> dict:
    return {
        "path": snapshot_.path,
        "part_count": len(snapshot_.parts),
        "relationship_count": len(snapshot_.relationships),
        "content_override_count": len(snapshot_.content_overrides),
        "dxfs_count": snapshot_.dxfs_count,
        "cf_dxf_ref_count": len(snapshot_.cf_dxf_refs),
        "feature_part_counts": {
            feature: len(parts) for feature, parts in snapshot_.feature_parts.items()
        },
    }


def _json_default(value: object) -> object:
    if isinstance(value, set):
        return sorted(value)
    if hasattr(value, "__dataclass_fields__"):
        return asdict(value)
    raise TypeError(f"Object of type {type(value).__name__} is not JSON serializable")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("before", type=Path, help="Workbook before modify-save")
    parser.add_argument("after", type=Path, help="Workbook after modify-save")
    parser.add_argument("--json", action="store_true", help="Emit machine-readable JSON")
    args = parser.parse_args(argv)

    report = audit(args.before, args.after)
    if args.json:
        print(json.dumps(report, indent=2, default=_json_default, sort_keys=True))
    else:
        _print_text_report(report)
    return 1 if report["issues"] else 0


def _print_text_report(report: dict) -> None:
    print(f"Before parts: {report['before']['part_count']}")
    print(f"After parts:  {report['after']['part_count']}")
    print(f"Issues:       {report['issue_count']}")
    for issue in report["issues"]:
        print(f"- [{issue['severity']}] {issue['kind']}: {issue['message']}")


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
