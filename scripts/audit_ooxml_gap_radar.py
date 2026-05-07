#!/usr/bin/env python3
"""Inventory unclassified OOXML surface area in fidelity fixture packs."""

from __future__ import annotations

import argparse
import json
import posixpath
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import audit_ooxml_fidelity  # noqa: E402
import run_ooxml_fidelity_mutations  # noqa: E402

REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
CT_NS = "{http://schemas.openxmlformats.org/package/2006/content-types}"

CORE_PART_PATTERNS = (
    re.compile(r"^\[Content_Types\]\.xml$"),
    re.compile(r"^_rels/\.rels$"),
    re.compile(r"^docProps/[^/]+\.xml$"),
    re.compile(r"^package/services/metadata/core-properties/[^/]+\.psmdcp$"),
    re.compile(r"^xl/workbook\.xml$"),
    re.compile(r"^xl/_rels/workbook\.xml\.rels$"),
    re.compile(r"^xl/worksheets/sheet\d+\.xml$"),
    re.compile(r"^xl/worksheets/_rels/sheet\d+\.xml\.rels$"),
    re.compile(r"^xl/styles\.xml$"),
    re.compile(r"^xl/sharedStrings\.xml$"),
    re.compile(r"^xl/theme/theme\d+\.xml$"),
    re.compile(r"^xl/metadata\.xml$"),
)

CORE_REL_TYPES = {
    "core-properties",
    "extended-properties",
    "custom-properties",
    "officeDocument",
    "worksheet",
    "styles",
    "theme",
    "sharedStrings",
    "metadata",
    "calcChain",
    "hyperlink",
}


def audit_gap_radar(fixture_dir: Path) -> dict:
    fixture_dir = fixture_dir.resolve()
    fixtures = []
    unknown_part_fixtures: dict[str, set[str]] = {}
    unknown_rel_fixtures: dict[str, set[str]] = {}
    unknown_content_type_fixtures: dict[str, set[str]] = {}

    for entry in run_ooxml_fidelity_mutations.discover_fixtures(fixture_dir):
        path = fixture_dir / entry.filename
        if not path.is_file():
            continue
        fixture_report = _fixture_unknowns(path)
        fixtures.append(
            {
                "filename": entry.filename,
                "fixture_id": entry.fixture_id,
                "tool": entry.tool,
                **fixture_report,
            }
        )
        for family in fixture_report["unknown_part_families"]:
            unknown_part_fixtures.setdefault(family, set()).add(entry.filename)
        for rel_type in fixture_report["unknown_relationship_types"]:
            unknown_rel_fixtures.setdefault(rel_type, set()).add(entry.filename)
        for content_type in fixture_report["unknown_content_types"]:
            unknown_content_type_fixtures.setdefault(content_type, set()).add(entry.filename)

    return {
        "fixture_dir": str(fixture_dir),
        "fixture_count": len(fixtures),
        "fixtures": fixtures,
        "unknown_part_families": _sorted_mapping(unknown_part_fixtures),
        "unknown_relationship_types": _sorted_mapping(unknown_rel_fixtures),
        "unknown_content_types": _sorted_mapping(unknown_content_type_fixtures),
        "unknown_part_family_count": len(unknown_part_fixtures),
        "unknown_relationship_type_count": len(unknown_rel_fixtures),
        "unknown_content_type_count": len(unknown_content_type_fixtures),
        "clear": not unknown_part_fixtures
        and not unknown_rel_fixtures
        and not unknown_content_type_fixtures,
    }


def _fixture_unknowns(path: Path) -> dict:
    with zipfile.ZipFile(path) as archive:
        parts = set(archive.namelist())
        feature_parts = audit_ooxml_fidelity._classify_feature_parts(parts)
        classified_parts = {part for values in feature_parts.values() for part in values}
        unknown_parts = sorted(
            part
            for part in parts
            if part not in classified_parts and not _is_core_part(part)
        )
        unknown_families = sorted({_part_family(part) for part in unknown_parts})
        unknown_relationship_types = sorted(
            {
                _rel_type_tail(rel_type)
                for rel_type in _relationship_types(archive)
                if not _is_known_relationship_type(rel_type)
            }
        )
        unknown_content_types = sorted(
            {
                content_type
                for part, content_type in _content_overrides(archive).items()
                if part in unknown_parts
            }
        )
    return {
        "unknown_parts": unknown_parts,
        "unknown_part_families": unknown_families,
        "unknown_relationship_types": unknown_relationship_types,
        "unknown_content_types": unknown_content_types,
    }


def _is_core_part(part: str) -> bool:
    return any(pattern.match(part) for pattern in CORE_PART_PATTERNS)


def _part_family(part: str) -> str:
    basename = posixpath.basename(part)
    normalized = re.sub(r"\d+", "{n}", basename)
    parent = posixpath.dirname(part)
    return f"{parent}/{normalized}" if parent else normalized


def _relationship_types(archive: zipfile.ZipFile) -> set[str]:
    out: set[str] = set()
    for rels_part in sorted(p for p in archive.namelist() if p.endswith(".rels")):
        try:
            root = ElementTree.fromstring(archive.read(rels_part))
        except ElementTree.ParseError:
            continue
        for node in root.findall(f"{REL_NS}Relationship"):
            if rel_type := node.attrib.get("Type"):
                out.add(rel_type)
    return out


def _content_overrides(archive: zipfile.ZipFile) -> dict[str, str]:
    try:
        root = ElementTree.fromstring(archive.read("[Content_Types].xml"))
    except (KeyError, ElementTree.ParseError):
        return {}
    out: dict[str, str] = {}
    for node in root.findall(f"{CT_NS}Override"):
        part_name = node.attrib.get("PartName", "").lstrip("/")
        content_type = node.attrib.get("ContentType")
        if part_name and content_type:
            out[part_name] = content_type
    return out


def _is_known_relationship_type(rel_type: str) -> bool:
    tail = _rel_type_tail(rel_type)
    if tail in CORE_REL_TYPES:
        return True
    return any(tail in prefixes for prefixes in _known_feature_relationship_tails())


def _known_feature_relationship_tails() -> tuple[tuple[str, ...], ...]:
    return (
        ("chart", "chartsheet", "chartUserShapes", "chartStyle", "chartColorStyle"),
        ("comments", "threadedComment", "person", "vmlDrawing"),
        ("drawing", "image", "printerSettings"),
        ("externalLink", "externalLinkPath", "xlPathMissing"),
        ("table",),
        ("pivotTable", "pivotCacheDefinition", "pivotCacheRecords"),
        ("slicer", "slicerCache", "timeline", "timelineCache"),
        ("vbaProject", "control", "ctrlProp", "activeXControl", "oleObject"),
        ("customXml", "customXmlProps"),
    )


def _rel_type_tail(rel_type: str) -> str:
    return rel_type.rstrip("/").rsplit("/", 1)[-1]


def _sorted_mapping(mapping: dict[str, set[str]]) -> dict[str, list[str]]:
    return {key: sorted(values) for key, values in sorted(mapping.items())}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("fixture_dir", type=Path)
    parser.add_argument("--json", action="store_true", help="Emit JSON")
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Exit non-zero when unclassified package surface is present.",
    )
    args = parser.parse_args(argv)

    report = audit_gap_radar(args.fixture_dir)
    if args.json:
        print(json.dumps(report, indent=2, sort_keys=True))
    else:
        print(f"Fixtures: {report['fixture_count']}")
        print(f"Unknown part families: {report['unknown_part_family_count']}")
        print(f"Unknown relationship types: {report['unknown_relationship_type_count']}")
        print(f"Unknown content types: {report['unknown_content_type_count']}")
        for family, fixtures in report["unknown_part_families"].items():
            print(f"- {family}: {', '.join(fixtures)}")
    return 1 if args.strict and not report["clear"] else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
