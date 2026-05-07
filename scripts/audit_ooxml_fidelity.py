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
    "timeline": ("xl/timelines/", "xl/timelineCaches/"),
    "vba": ("xl/vbaProject.bin",),
}

CF_EXTENSION_NAMES = frozenset(
    {
        "conditionalFormatting",
        "conditionalFormattings",
        "cfRule",
        "colorScale",
        "dataBar",
        "iconSet",
        "pivotAreas",
    }
)

SLICER_EXTENSION_NAMES = frozenset({"slicerCaches", "slicerList"})
TIMELINE_EXTENSION_NAMES = frozenset({"timelineCacheRefs", "timelineRefs", "timelineList"})


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
    semantic_fingerprints: dict[str, dict[str, object]]


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
            semantic_fingerprints=_read_semantic_fingerprints(archive),
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
    _audit_semantic_fingerprints(before_snapshot, after_snapshot, issues)

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


def _audit_semantic_fingerprints(
    before: Snapshot, after: Snapshot, issues: list[dict[str, str]]
) -> None:
    for feature, before_fingerprint in sorted(before.semantic_fingerprints.items()):
        if not before_fingerprint:
            continue
        after_fingerprint = after.semantic_fingerprints.get(feature, {})
        if after_fingerprint != before_fingerprint:
            issues.append(
                {
                    "severity": "error",
                    "kind": f"{feature}_semantic_drift",
                    "part": feature,
                    "message": (
                        f"{feature} semantic fingerprint changed after save: "
                        f"before={before_fingerprint!r} after={after_fingerprint!r}"
                    ),
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


def _read_semantic_fingerprints(archive: zipfile.ZipFile) -> dict[str, dict[str, object]]:
    parts = set(archive.namelist())
    return {
        "charts": _chart_fingerprint(archive, parts),
        "chart_styles": _chart_style_fingerprint(archive, parts),
        "conditional_formatting": _conditional_formatting_fingerprint(archive, parts),
        "data_validations": _data_validation_fingerprint(archive, parts),
        "external_links": _external_link_fingerprint(archive, parts),
        "pivots": _pivot_fingerprint(archive, parts),
        "slicers": _slicer_fingerprint(archive, parts),
        "timelines": _timeline_fingerprint(archive, parts),
    }


def _chart_fingerprint(
    archive: zipfile.ZipFile, parts: set[str]
) -> dict[str, list[object]]:
    out: dict[str, list[object]] = {}
    rels_by_owner = _relationships_by_owner(archive)
    for part in sorted(_feature_xml_parts(parts, "xl/charts/", ".xml")):
        if _is_chart_style_part(part):
            continue
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        out[part] = [
            ("formulas", _texts_by_local(root, "f")),
            ("pivot_sources", _pivot_source_names(root)),
            ("dPt_count", len(_nodes_by_local(root, "dPt"))),
            ("style_vals", _vals_by_path(root, ("style",))),
            ("rels", rels_by_owner.get(part, [])),
        ]
    return out


def _chart_style_fingerprint(
    archive: zipfile.ZipFile, parts: set[str]
) -> dict[str, object]:
    out: dict[str, object] = {}
    for part in sorted(_feature_xml_parts(parts, "xl/charts/", ".xml")):
        if not _is_chart_style_part(part):
            continue
        root = _read_xml_or_none(archive, part)
        if root is not None:
            out[part] = _xml_tree_fingerprint(root)
    return out


def _conditional_formatting_fingerprint(
    archive: zipfile.ZipFile, parts: set[str]
) -> dict[str, list[object]]:
    out: dict[str, list[object]] = {}
    for part in sorted(_worksheet_parts(parts)):
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        blocks: list[object] = []
        for block in _nodes_by_local(root, "conditionalFormatting"):
            rules: list[object] = []
            for rule in _children_by_local(block, "cfRule"):
                rules.append(
                    (
                        _stable_attrs(rule, ("type", "priority", "operator", "dxfId")),
                        _texts_by_local(rule, "formula"),
                    )
                )
            blocks.append(
                (
                    _attr(block, "sqref"),
                    rules,
                    _extension_fingerprints(block, CF_EXTENSION_NAMES),
                )
            )
        extensions = _extension_fingerprints(root, CF_EXTENSION_NAMES)
        if blocks or extensions:
            out[part] = [("blocks", blocks), ("extensions", extensions)]
    return out


def _data_validation_fingerprint(
    archive: zipfile.ZipFile, parts: set[str]
) -> dict[str, list[object]]:
    out: dict[str, list[object]] = {}
    for part in sorted(_worksheet_parts(parts)):
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        validations: list[object] = []
        for validation in _nodes_by_local(root, "dataValidation"):
            validations.append(
                (
                    _stable_attrs(
                        validation,
                        (
                            "type",
                            "operator",
                            "allowBlank",
                            "showErrorMessage",
                            "showInputMessage",
                            "sqref",
                        ),
                    ),
                    _texts_by_local(validation, "formula1"),
                    _texts_by_local(validation, "formula2"),
                )
            )
        if validations:
            out[part] = validations
    return out


def _external_link_fingerprint(
    archive: zipfile.ZipFile, parts: set[str]
) -> dict[str, list[object]]:
    out: dict[str, list[object]] = {}
    rel_targets = _rels_target_lookup(archive)
    for part in sorted(_feature_xml_parts(parts, "xl/externalLinks/", ".xml")):
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        external_books: list[object] = []
        for book in _nodes_by_local(root, "externalBook"):
            rid = _relationship_id(book)
            external_books.append(
                (
                    rid,
                    rel_targets.get((part, rid)),
                    _external_sheet_names(book),
                    _defined_name_refs(book),
                    _external_sheet_data(book),
                )
            )
        out[part] = external_books
    workbook_formulas = _worksheet_formulas(archive, parts, external_only=True)
    if workbook_formulas:
        out["worksheet_formulas"] = workbook_formulas
    return out


def _pivot_fingerprint(
    archive: zipfile.ZipFile, parts: set[str]
) -> dict[str, list[object]]:
    out: dict[str, list[object]] = {}
    rels_by_owner = _relationships_by_owner(archive)
    for part in sorted(_feature_xml_parts(parts, "xl/pivotTables/", ".xml")):
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        out[part] = [
            ("attrs", _stable_attrs(root, ("name", "cacheId", "dataOnRows"))),
            ("rels", rels_by_owner.get(part, [])),
            ("data_fields", _pivot_data_fields(root)),
            ("row_fields", _pivot_field_indices(root, "rowFields", "field")),
            ("col_fields", _pivot_field_indices(root, "colFields", "field")),
            ("page_fields", _pivot_field_indices(root, "pageFields", "pageField")),
            ("calculated_items", _pivot_calculated_items(root)),
            ("formats", _pivot_formats(root)),
            ("conditional_formats", _pivot_conditional_formats(root)),
        ]
    for part in sorted(_feature_xml_parts(parts, "xl/pivotCache/", ".xml")):
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        source = _first_node_by_local(root, "worksheetSource")
        out[part] = [
            ("cacheSource", _stable_attrs(source, ("ref", "sheet", "name"))),
            ("refreshOnLoad", _attr(root, "refreshOnLoad")),
            ("rels", rels_by_owner.get(part, [])),
            ("fields", _pivot_cache_fields(root)),
            ("calculated_fields", _pivot_calculated_fields(root)),
            ("field_groups", _pivot_field_groups(root)),
        ]
    return out


def _slicer_fingerprint(
    archive: zipfile.ZipFile, parts: set[str]
) -> dict[str, list[object]]:
    out: dict[str, list[object]] = {}
    rels_by_owner = _relationships_by_owner(archive)
    workbook_root = _read_xml_or_none(archive, "xl/workbook.xml")
    if workbook_root is not None:
        extensions = _extension_fingerprints(workbook_root, SLICER_EXTENSION_NAMES)
        if extensions:
            out["xl/workbook.xml"] = [("extensions", extensions)]
    for part in sorted(_worksheet_parts(parts)):
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        extensions = _extension_fingerprints(root, SLICER_EXTENSION_NAMES)
        if extensions:
            out[part] = [("extensions", extensions)]
    for part in sorted(_feature_xml_parts(parts, "xl/slicerCaches/", ".xml")):
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        out[part] = [
            ("attrs", _stable_attrs(root, ("name", "pivotCacheId"))),
            ("rels", rels_by_owner.get(part, [])),
            ("data", _stable_attrs(_first_node_by_local(root, "data"), ("pivotCacheId",))),
            ("items", _slicer_items(root)),
        ]
    for part in sorted(_feature_xml_parts(parts, "xl/slicers/", ".xml")):
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        out[part] = [
            ("attrs", _stable_attrs(root, ("name", "cache", "caption", "style"))),
            ("rels", rels_by_owner.get(part, [])),
            (
                "slicers",
                [
                    _stable_attrs(node, ("name", "cache", "caption"))
                    for node in _nodes_by_local(root, "slicer")
                ],
            ),
        ]
    return out


def _timeline_fingerprint(
    archive: zipfile.ZipFile, parts: set[str]
) -> dict[str, list[object]]:
    out: dict[str, list[object]] = {}
    rels_by_owner = _relationships_by_owner(archive)
    workbook_root = _read_xml_or_none(archive, "xl/workbook.xml")
    if workbook_root is not None:
        extensions = _extension_fingerprints(workbook_root, TIMELINE_EXTENSION_NAMES)
        if extensions:
            out["xl/workbook.xml"] = [("extensions", extensions)]
    for part in sorted(_worksheet_parts(parts)):
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        extensions = _extension_fingerprints(root, TIMELINE_EXTENSION_NAMES)
        if extensions:
            out[part] = [("extensions", extensions)]
    for prefix in ("xl/timelineCaches/", "xl/timelines/"):
        for part in sorted(_feature_xml_parts(parts, prefix, ".xml")):
            root = _read_xml_or_none(archive, part)
            if root is not None:
                out[part] = [
                    ("attrs", _stable_attrs(root, ("name", "pivotCacheId", "cache"))),
                    ("rels", rels_by_owner.get(part, [])),
                    ("xml", _xml_tree_fingerprint(root)),
                ]
    return out


def _read_xml_or_none(archive: zipfile.ZipFile, part: str) -> ElementTree.Element | None:
    try:
        return ElementTree.fromstring(archive.read(part))
    except (KeyError, ElementTree.ParseError):
        return None


def _feature_xml_parts(parts: set[str], prefix: str, suffix: str) -> Iterable[str]:
    return (part for part in parts if part.startswith(prefix) and part.endswith(suffix))


def _is_chart_style_part(part: str) -> bool:
    name = posixpath.basename(part)
    return name.startswith("style") or name.startswith("colors")


def _nodes_by_local(root: ElementTree.Element, name: str) -> list[ElementTree.Element]:
    return [node for node in root.iter() if _local_name(node.tag) == name]


def _first_node_by_local(
    root: ElementTree.Element, name: str
) -> ElementTree.Element | None:
    nodes = _nodes_by_local(root, name)
    return nodes[0] if nodes else None


def _children_by_local(root: ElementTree.Element, name: str) -> list[ElementTree.Element]:
    return [node for node in list(root) if _local_name(node.tag) == name]


def _texts_by_local(root: ElementTree.Element, name: str) -> list[str]:
    return [text for node in _nodes_by_local(root, name) if (text := _text(node))]


def _vals_by_path(root: ElementTree.Element, names: tuple[str, ...]) -> list[str]:
    return [
        val
        for node in root.iter()
        if _local_name(node.tag) in names and (val := _attr(node, "val")) is not None
    ]


def _pivot_source_names(root: ElementTree.Element) -> list[str]:
    names: list[str] = []
    for pivot_source in _nodes_by_local(root, "pivotSource"):
        names.extend(_texts_by_local(pivot_source, "name"))
    return names


def _defined_name_refs(root: ElementTree.Element) -> list[tuple[str | None, str | None]]:
    return [
        (_attr(node, "name"), _attr(node, "refersTo"))
        for node in _nodes_by_local(root, "definedName")
    ]


def _external_sheet_names(root: ElementTree.Element) -> list[str | None]:
    return [_attr(node, "val") for node in _nodes_by_local(root, "sheetName")]


def _external_sheet_data(root: ElementTree.Element) -> list[object]:
    sheets: list[object] = []
    for sheet_data in _nodes_by_local(root, "sheetData"):
        rows: list[object] = []
        for row in _children_by_local(sheet_data, "row"):
            cells = [
                (_stable_attrs(cell, ("r", "t", "vm")), _texts_by_local(cell, "v"))
                for cell in _children_by_local(row, "cell")
            ]
            rows.append((_stable_attrs(row, ("r",)), cells))
        sheets.append(
            (
                _stable_attrs(sheet_data, ("sheetId", "refreshError")),
                rows,
            )
        )
    return sheets


def _worksheet_formulas(
    archive: zipfile.ZipFile, parts: set[str], *, external_only: bool = False
) -> dict[str, list[str]]:
    out: dict[str, list[str]] = {}
    for part in sorted(_worksheet_parts(parts)):
        root = _read_xml_or_none(archive, part)
        if root is None:
            continue
        formulas = _texts_by_local(root, "f")
        if external_only:
            formulas = [formula for formula in formulas if "[" in formula and "]" in formula]
        if formulas:
            out[part] = formulas
    return out


def _pivot_data_fields(root: ElementTree.Element) -> list[tuple[tuple[str, str | None], ...]]:
    return [
        _stable_attrs(node, ("name", "fld", "subtotal", "baseField", "baseItem"))
        for node in _nodes_by_local(root, "dataField")
    ]


def _pivot_field_indices(
    root: ElementTree.Element, container_name: str, child_name: str
) -> list[str | None]:
    container = _first_node_by_local(root, container_name)
    if container is None:
        return []
    return [
        _attr(child, "x") or _attr(child, "fld")
        for child in _children_by_local(container, child_name)
    ]


def _pivot_cache_fields(root: ElementTree.Element) -> list[tuple[tuple[str, str | None], ...]]:
    return [
        _stable_attrs(node, ("name", "numFmtId", "databaseField", "formula"))
        for node in _nodes_by_local(root, "cacheField")
    ]


def _pivot_calculated_fields(
    root: ElementTree.Element,
) -> list[tuple[tuple[str, str | None], ...]]:
    return [
        _stable_attrs(node, ("name", "formula", "hierarchy", "memberName", "mdx", "solveOrder"))
        for node in _nodes_by_local(root, "calculatedField")
    ]


def _pivot_calculated_items(root: ElementTree.Element) -> list[object]:
    return [
        (
            _stable_attrs(node, ("field", "formula")),
            _extension_fingerprints(node, CF_EXTENSION_NAMES),
        )
        for node in _nodes_by_local(root, "calculatedItem")
    ]


def _pivot_formats(root: ElementTree.Element) -> list[object]:
    return [
        (
            _stable_attrs(node, ("action", "dxfId")),
            [
                _stable_attrs(area, ("type", "field", "fieldPosition"))
                for area in _nodes_by_local(node, "pivotArea")
            ],
        )
        for node in _nodes_by_local(root, "format")
    ]


def _pivot_conditional_formats(root: ElementTree.Element) -> list[object]:
    return [
        (
            _stable_attrs(node, ("scope", "type", "priority")),
            [
                _stable_attrs(area, ("type", "field", "fieldPosition"))
                for area in _nodes_by_local(node, "pivotArea")
            ],
            _extension_fingerprints(node, CF_EXTENSION_NAMES),
        )
        for node in _nodes_by_local(root, "conditionalFormat")
    ]


def _pivot_field_groups(root: ElementTree.Element) -> list[object]:
    groups: list[object] = []
    for node in _nodes_by_local(root, "fieldGroup"):
        groups.append(_xml_tree_fingerprint(node))
    return groups


def _slicer_items(root: ElementTree.Element) -> list[tuple[tuple[str, str | None], ...]]:
    return [
        _stable_attrs(node, ("n", "c", "x", "s"))
        for node in _nodes_by_local(root, "i")
    ]


def _extension_fingerprints(
    root: ElementTree.Element, interesting_names: frozenset[str]
) -> list[object]:
    out: list[object] = []
    for ext in _nodes_by_local(root, "ext"):
        if any(_local_name(node.tag) in interesting_names for node in ext.iter()):
            out.append(_xml_tree_fingerprint(ext))
    return out


def _xml_tree_fingerprint(node: ElementTree.Element) -> object:
    return (
        _local_name(node.tag),
        _all_stable_attrs(node),
        _text(node),
        [_xml_tree_fingerprint(child) for child in list(node)],
    )


def _all_stable_attrs(node: ElementTree.Element) -> tuple[tuple[str, str], ...]:
    return tuple(sorted((_local_name(key), value) for key, value in node.attrib.items()))


def _relationships_by_owner(
    archive: zipfile.ZipFile,
) -> dict[str, list[tuple[str, str, str, str | None]]]:
    lookup: dict[str, list[tuple[str, str, str, str | None]]] = {}
    for rel in _read_relationships(archive):
        owner = _source_part_for_rels(rel.rels_part)
        if owner:
            lookup.setdefault(owner, []).append(
                (rel.rel_id, rel.rel_type, rel.target, rel.target_mode)
            )
    return lookup


def _rels_target_lookup(archive: zipfile.ZipFile) -> dict[tuple[str, str], str]:
    lookup: dict[tuple[str, str], str] = {}
    for rel in _read_relationships(archive):
        owner = _source_part_for_rels(rel.rels_part)
        if owner:
            lookup[(owner, rel.rel_id)] = rel.target
    return lookup


def _relationship_id(node: ElementTree.Element) -> str | None:
    for key, value in node.attrib.items():
        if key.endswith("}id") or key == "id":
            return value
    return None


def _stable_attrs(
    node: ElementTree.Element | None, names: Iterable[str]
) -> tuple[tuple[str, str | None], ...]:
    if node is None:
        return tuple((name, None) for name in names)
    return tuple((name, _attr(node, name)) for name in names)


def _attr(node: ElementTree.Element | None, name: str) -> str | None:
    if node is None:
        return None
    for key, value in node.attrib.items():
        if _local_name(key) == name:
            return value
    return None


def _text(node: ElementTree.Element) -> str | None:
    value = (node.text or "").strip()
    return value or None


def _local_name(name: str) -> str:
    return name.rsplit("}", 1)[-1] if "}" in name else name


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
        "semantic_fingerprint_counts": {
            feature: len(fingerprint)
            for feature, fingerprint in snapshot_.semantic_fingerprints.items()
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
