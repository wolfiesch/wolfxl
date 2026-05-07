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
    re.compile(r"^docProps/thumbnail\.wmf$"),
    re.compile(r"^package/services/metadata/core-properties/[^/]+\.psmdcp$"),
    re.compile(r"^xl/workbook\.xml$"),
    re.compile(r"^xl/_rels/workbook\.xml\.rels$"),
    re.compile(r"^xl/worksheets/sheet\d+\.xml$"),
    re.compile(r"^xl/worksheets/_rels/sheet\d+\.xml\.rels$"),
    re.compile(r"^xl/styles\.xml$"),
    re.compile(r"^xl/sharedStrings\.xml$"),
    re.compile(r"^xl/theme/theme(?:\d+)?\.xml$"),
    re.compile(r"^xl/metadata\.xml$"),
)

CORE_REL_TYPES = {
    "core-properties",
    "extended-properties",
    "custom-properties",
    "classificationlabels",
    "officeDocument",
    "worksheet",
    "styles",
    "theme",
    "thumbnail",
    "sharedStrings",
    "metadata",
    "calcChain",
    "connections",
    "hyperlink",
}

KNOWN_EXTENSION_URIS = {
    "{02D57815-91ED-43cb-92C2-25804820EDAC}",
    "{05A4C25C-085E-4340-85A3-A5531E510DB2}",
    "{0605FD5F-26C8-4aeb-8148-2DB25E43C511}",
    "{140A7094-0E35-4892-8432-C4D2E57EDEB5}",
    "{28A0092B-C50C-407E-A947-70E740481C1C}",
    "{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}",
    "{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}",
    "{46BE6895-7355-4a93-B00E-2C351335B9C9}",
    "{53640926-AAD7-44D8-BBD7-CCE9431645EC}",
    "{63B3BB69-23CF-44E3-9099-C40C66FF867C}",
    "{725AE2AE-9491-48be-B2B4-4EB974FC3084}",
    "{747A6164-185A-40DC-8AA5-F01512510D54}",
    "{7626C862-2A13-11E5-B345-FEFF819CDC9F}",
    "{781A3756-C4B2-4CAC-9D66-4F8BD8637D16}",
    "{78C0D931-6437-407d-A8EE-F0AAD7539E65}",
    "{79F54976-1DA5-4618-B147-4CDE4B953A38}",
    "{7E03D99C-DC04-49d9-9315-930204A7B6E9}",
    "{876F7934-8845-4945-9796-88D515C7AA90}",
    "{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}",
    "{91240B29-F687-4F45-9708-019B960494DF}",
    "{9260A510-F301-46a8-8635-F512D64BE5F5}",
    "{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}",
    "{A8765BA9-456A-4dab-B4F3-ACF838C121DE}",
    "{AF507438-7753-43E0-B8FC-AC1667EBCBE1}",
    "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}",
    "{B58B0392-4F1F-4190-BB64-5DF3571DCE5F}",
    "{B97F6D7D-B522-45F9-BDA1-12C45D357490}",
    "{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}",
    "{BBE1A952-AA13-448e-AADC-164F8A28A991}",
    "{C3380CC4-5D6E-409C-BE32-E72D297353CC}",
    "{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}",
    "{CE6537A1-D6FC-4f65-9D91-7224C49458BB}",
    "{D0CA8CA8-9F24-4464-BF8E-62219DCF47F9}",
    "{D14903EA-33C4-47F7-8F05-3474C54BE107}",
    "{DE250136-89BD-433C-8126-D09CA5730AF9}",
    "{E28EC0CA-F0BB-4C9C-879D-F8772B89E7AC}",
    "{E67621CE-5B39-4880-91FE-76760E9C1902}",
    "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}",
    "{F057638F-6D5F-4e77-A914-E7F072B9BCA8}",
    "{FCE2AD5D-F65C-4FA6-A056-5C36A1767C68}",
    "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}",
    "{56B9EC1D-385E-4148-901F-78D8002777C0}",
}


def audit_gap_radar(fixture_dir: Path, recursive: bool = False) -> dict:
    fixture_dir = fixture_dir.resolve()
    fixtures = []
    unknown_part_fixtures: dict[str, set[str]] = {}
    unknown_rel_fixtures: dict[str, set[str]] = {}
    unknown_content_type_fixtures: dict[str, set[str]] = {}
    unknown_extension_uri_fixtures: dict[str, set[str]] = {}

    for entry in run_ooxml_fidelity_mutations.discover_fixtures(
        fixture_dir, recursive=recursive
    ):
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
        for uri in fixture_report["unknown_extension_uris"]:
            unknown_extension_uri_fixtures.setdefault(uri, set()).add(entry.filename)

    return {
        "fixture_dir": str(fixture_dir),
        "fixture_count": len(fixtures),
        "recursive": recursive,
        "fixtures": fixtures,
        "unknown_part_families": _sorted_mapping(unknown_part_fixtures),
        "unknown_relationship_types": _sorted_mapping(unknown_rel_fixtures),
        "unknown_content_types": _sorted_mapping(unknown_content_type_fixtures),
        "unknown_extension_uris": _sorted_mapping(unknown_extension_uri_fixtures),
        "unknown_part_family_count": len(unknown_part_fixtures),
        "unknown_relationship_type_count": len(unknown_rel_fixtures),
        "unknown_content_type_count": len(unknown_content_type_fixtures),
        "unknown_extension_uri_count": len(unknown_extension_uri_fixtures),
        "clear": not unknown_part_fixtures
        and not unknown_rel_fixtures
        and not unknown_content_type_fixtures
        and not unknown_extension_uri_fixtures,
    }


def _fixture_unknowns(path: Path) -> dict:
    with zipfile.ZipFile(path) as archive:
        parts = set(archive.namelist())
        feature_parts = audit_ooxml_fidelity._classify_feature_parts(parts)
        classified_parts = {part for values in feature_parts.values() for part in values}
        unknown_parts = sorted(
            part
            for part in parts
            if not part.endswith("/")
            and part not in classified_parts
            and not _is_core_part(part)
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
        unknown_extension_uris = sorted(
            uri for uri in _extension_uris(archive) if uri not in KNOWN_EXTENSION_URIS
        )
    return {
        "unknown_parts": unknown_parts,
        "unknown_part_families": unknown_families,
        "unknown_relationship_types": unknown_relationship_types,
        "unknown_content_types": unknown_content_types,
        "unknown_extension_uris": unknown_extension_uris,
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


def _extension_uris(archive: zipfile.ZipFile) -> set[str]:
    out: set[str] = set()
    for part in sorted(p for p in archive.namelist() if p.endswith(".xml")):
        try:
            root = ElementTree.fromstring(archive.read(part))
        except ElementTree.ParseError:
            continue
        for node in root.iter():
            if _local_name(node.tag) == "ext":
                if uri := node.attrib.get("uri"):
                    out.add(uri)
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
        ("Python", "sheetMetadata"),
        ("jsaProject",),
        ("table",),
        ("pivotTable", "pivotCacheDefinition", "pivotCacheRecords"),
        ("powerPivotData", "model"),
        ("slicer", "slicerCache", "timeline", "timelineCache"),
        ("vbaProject", "control", "ctrlProp", "activeXControl", "oleObject"),
        ("customXml", "customXmlProps"),
    )


def _rel_type_tail(rel_type: str) -> str:
    return rel_type.rstrip("/").rsplit("/", 1)[-1]


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


def _sorted_mapping(mapping: dict[str, set[str]]) -> dict[str, list[str]]:
    return {key: sorted(values) for key, values in sorted(mapping.items())}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("fixture_dir", type=Path)
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Discover .xlsx fixtures recursively when no manifest.json is present.",
    )
    parser.add_argument("--json", action="store_true", help="Emit JSON")
    parser.add_argument(
        "--strict",
        action="store_true",
        help="Exit non-zero when unclassified package surface is present.",
    )
    args = parser.parse_args(argv)

    report = audit_gap_radar(args.fixture_dir, recursive=args.recursive)
    if args.json:
        print(json.dumps(report, indent=2, sort_keys=True))
    else:
        print(f"Fixtures: {report['fixture_count']}")
        print(f"Unknown part families: {report['unknown_part_family_count']}")
        print(f"Unknown relationship types: {report['unknown_relationship_type_count']}")
        print(f"Unknown content types: {report['unknown_content_type_count']}")
        print(f"Unknown extension URIs: {report['unknown_extension_uri_count']}")
        for family, fixtures in report["unknown_part_families"].items():
            print(f"- {family}: {', '.join(fixtures)}")
        for uri, fixtures in report["unknown_extension_uris"].items():
            print(f"- {uri}: {', '.join(fixtures)}")
    return 1 if args.strict and not report["clear"] else 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
