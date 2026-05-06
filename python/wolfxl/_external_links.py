"""External links surface (RFC-071 / G18).

This module wraps the Rust parsers
(``wolfxl._rust.parse_external_link_part`` /
``parse_external_link_rels``) into Python dataclasses that match
openpyxl's ``ExternalLink`` / ``ExternalFileLink`` shape just enough for
the compat oracle and downstream introspection:

* ``ExternalLink.target`` — linked workbook filename, e.g. ``ext.xlsx``.
* ``ExternalLink.sheet_names`` — sheet names referenced by formulas.
* ``ExternalLink.cached_data`` — a loose dict (sheetId -> [{r, v}]) of
  any cached cell values present in ``<sheetDataSet>``.
* ``ExternalLink.rid`` — the rels id under which the part is wired into
  ``xl/workbook.xml.rels``.
* ``ExternalLink.file_link`` — small companion holding ``target`` /
  ``target_mode`` so callers can check ``link.file_link.target_mode ==
  "External"``.

The collection is read-only in v1.0 (RFC-071 §5 / §8): we expose the
parsed list, but the patcher rewrites ``xl/externalLinks/`` parts
byte-for-byte on save. Authoring is deferred to a follow-up RFC.
"""

from __future__ import annotations

import zipfile
from dataclasses import dataclass, field
from typing import Any
from xml.etree import ElementTree as ET

from wolfxl._zip_safety import read_entry, validate_zipfile

_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_REL_TYPE_EXTERNAL_LINK = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink"
)


@dataclass(frozen=True)
class ExternalFileLink:
    """Reference to the linked workbook file.

    ``target`` is the relationship Target verbatim — e.g. ``ext.xlsx`` for
    a sibling file or ``../foo/data.xlsx`` for one in a parent folder. We
    deliberately do NOT URL-decode it because Open Packaging Conventions
    treat the Target as opaque.
    """

    target: str
    target_mode: str = "External"


@dataclass(frozen=True)
class ExternalLink:
    """One entry in :attr:`Workbook._external_links`.

    All fields are read-only in v1.0. ``cached_data`` shape is loose by
    design — see RFC-071 §3.
    """

    file_link: ExternalFileLink
    rid: str
    target: str
    sheet_names: list[str] = field(default_factory=list)
    cached_data: dict[str, list[dict[str, str]]] = field(default_factory=dict)


def load_external_links(source_path: str | None) -> list[ExternalLink]:
    """Load external-link parts from an xlsx at ``source_path``.

    Returns an empty list when:

    * ``source_path`` is ``None`` (workbook materialised from bytes with
      no tempfile path threaded through).
    * The xlsx has no ``xl/_rels/workbook.xml.rels`` (degenerate input).
    * No rel of type ``…/externalLink`` exists.

    Raises ``FileNotFoundError`` only on the hard case where
    ``source_path`` is set but the file is gone.
    """
    if not source_path:
        return []

    try:
        with zipfile.ZipFile(source_path, "r") as zf:
            validate_zipfile(zf)
            return _load_from_zip(zf)
    except FileNotFoundError:
        raise
    except (zipfile.BadZipFile, KeyError, OSError, ValueError):
        # A malformed source ZIP shouldn't kill the whole load. The
        # patcher will surface a clearer error on save if the file is
        # actually unusable.
        return []


def _load_from_zip(zf: zipfile.ZipFile) -> list[ExternalLink]:
    """Walk the workbook rels graph and parse every external link."""
    try:
        rels_bytes = read_entry(zf, "xl/_rels/workbook.xml.rels")
    except KeyError:
        return []

    rels = _parse_workbook_rels(rels_bytes)
    external_rels = [r for r in rels if r["type"] == _REL_TYPE_EXTERNAL_LINK]
    if not external_rels:
        return []

    out: list[ExternalLink] = []
    for rel in external_rels:
        # Target is relative to xl/_rels/workbook.xml.rels, so to xl/.
        target = rel["target"]
        part_path = _normalize_part_path("xl/", target)
        try:
            part_bytes = read_entry(zf, part_path)
        except KeyError:
            # Rel without a matching part — preserved on save by the
            # passthrough but not introspectable. Skip.
            continue

        parsed_part = _parse_part_xml(part_bytes)
        rels_path = _sibling_rels_path(part_path)
        try:
            part_rels_bytes = read_entry(zf, rels_path)
        except KeyError:
            part_rels_bytes = b""

        parsed_rels = _parse_part_rels_xml(part_rels_bytes)
        link_target = parsed_rels.get("target") or ""
        link_mode = parsed_rels.get("target_mode") or "External"

        file_link = ExternalFileLink(target=link_target, target_mode=link_mode)
        out.append(
            ExternalLink(
                file_link=file_link,
                rid=rel["id"],
                target=link_target,
                sheet_names=list(parsed_part.get("sheet_names", [])),
                cached_data=dict(parsed_part.get("cached_data", {})),
            )
        )

    return out


def _parse_workbook_rels(xml: bytes) -> list[dict[str, str]]:
    """Tiny rels parser specialised for the workbook rels graph.

    The Rust ``RelsGraph`` parser is overkill on the load path — we only
    need ``Id``, ``Type``, ``Target`` for one rel type. Using stdlib
    ElementTree keeps this hot path off the PyO3 boundary entirely for
    the empty case (no externalLink rels), which is the common case.
    """
    root = ET.fromstring(xml)
    out: list[dict[str, str]] = []
    for child in root:
        if not child.tag.endswith("}Relationship") and child.tag != "Relationship":
            continue
        out.append(
            {
                "id": child.attrib.get("Id", ""),
                "type": child.attrib.get("Type", ""),
                "target": child.attrib.get("Target", ""),
            }
        )
    return out


def _parse_part_xml(xml: bytes) -> dict[str, Any]:
    """Call the Rust parser, returning ``{}`` on import failure.

    The Rust extension is the production path; the import-failure
    fallback exists only so a half-built dev environment doesn't crash
    every load.
    """
    try:
        from wolfxl import _rust
    except ImportError:  # pragma: no cover - dev-only safety net
        return {"book_rid": None, "sheet_names": [], "cached_data": {}}
    return dict(_rust.parse_external_link_part(xml))


def _parse_part_rels_xml(xml: bytes) -> dict[str, Any]:
    """Call the Rust rels parser. Empty bytes -> empty dict."""
    if not xml:
        return {"target": None, "target_mode": None, "rid": None}
    try:
        from wolfxl import _rust
    except ImportError:  # pragma: no cover - dev-only safety net
        return {"target": None, "target_mode": None, "rid": None}
    return dict(_rust.parse_external_link_rels(xml))


def _normalize_part_path(base: str, target: str) -> str:
    """Resolve a rels target relative to ``base`` (`xl/`).

    Mirrors :func:`wolfxl._rust.ooxml_util::join_and_normalize` for the
    shapes external-link rels actually ship with: relative
    (``externalLinks/externalLink1.xml``) and parent-anchored
    (``../foo/x.xml``). Absolute targets (``/xl/...``) lose the leading
    slash and are taken as-is from the package root.
    """
    if target.startswith("/"):
        return target.lstrip("/")
    parts = (base + target).split("/")
    stack: list[str] = []
    for p in parts:
        if p == "" or p == ".":
            continue
        if p == "..":
            if stack:
                stack.pop()
            continue
        stack.append(p)
    return "/".join(stack)


def _sibling_rels_path(part_path: str) -> str:
    """For ``xl/externalLinks/externalLink1.xml`` return
    ``xl/externalLinks/_rels/externalLink1.xml.rels``."""
    if "/" not in part_path:
        return f"_rels/{part_path}.rels"
    parent, name = part_path.rsplit("/", 1)
    return f"{parent}/_rels/{name}.rels"
