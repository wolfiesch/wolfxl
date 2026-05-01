#!/usr/bin/env python3
"""Inventory BIFF12 records inside .xlsb package parts.

The script is intentionally read-only. It walks workbook, worksheet, table,
comments, styles, and shared-string binary parts and reports record counts by
part so native-reader parity work can start from observed fixture truth.
"""

from __future__ import annotations

import argparse
import collections
import zipfile
from pathlib import Path


RECORD_NAMES = {
    0x0000: "BrtRowHdr",
    0x0001: "BrtCellBlank",
    0x0002: "BrtCellRk",
    0x0003: "BrtCellError",
    0x0004: "BrtCellBool",
    0x0005: "BrtCellReal",
    0x0006: "BrtCellSt",
    0x0007: "BrtCellIsst",
    0x0008: "BrtFmlaString",
    0x0009: "BrtFmlaNum",
    0x000A: "BrtFmlaBool",
    0x000B: "BrtFmlaError",
    0x0013: "BrtSSTItem",
    0x002B: "BrtFont",
    0x002C: "BrtFmt",
    0x002D: "BrtFill",
    0x002E: "BrtBorder",
    0x002F: "BrtXF",
    0x003C: "BrtColInfo",
    0x0093: "BrtWsProp",
    0x0094: "BrtWsDim",
    0x0097: "BrtPane",
    0x0098: "BrtSel",
    0x009C: "BrtBundleSh",
    0x00A1: "BrtBeginAFilter",
    0x00A2: "BrtEndAFilter",
    0x00A3: "BrtBeginFilterColumn",
    0x00A4: "BrtEndFilterColumn",
    0x00A5: "BrtBeginFilters",
    0x00A6: "BrtEndFilters",
    0x00A7: "BrtFilter",
    0x00AA: "BrtTop10Filter",
    0x00AB: "BrtDynamicFilter",
    0x00AC: "BrtBeginCustomFilters",
    0x00AD: "BrtEndCustomFilters",
    0x00AE: "BrtCustomFilter",
    0x00AF: "BrtAFilterDateGroupItem",
    0x00B0: "BrtMergeCell",
    0x00B1: "BrtBeginMergeCells",
    0x00B2: "BrtEndMergeCells",
    0x0157: "BrtBeginList",
    0x0158: "BrtEndList",
    0x0159: "BrtBeginListCols",
    0x015A: "BrtEndListCols",
    0x015B: "BrtBeginListCol",
    0x015C: "BrtEndListCol",
    0x0186: "BrtBeginColInfos",
    0x0187: "BrtEndColInfos",
    0x01AC: "BrtTable",
    0x01DC: "BrtMargins",
    0x01DD: "BrtPrintOptions",
    0x01DE: "BrtPageSetup",
    0x01DF: "BrtBeginHeaderFooter",
    0x01E0: "BrtEndHeaderFooter",
    0x01E5: "BrtWsFmtInfo",
    0x01EE: "BrtHLink",
    0x0217: "BrtSheetProtection",
    0x0274: "BrtBeginComments",
    0x0275: "BrtEndComments",
    0x0276: "BrtBeginCommentAuthors",
    0x0277: "BrtEndCommentAuthors",
    0x0278: "BrtCommentAuthor",
    0x0279: "BrtBeginCommentList",
    0x027A: "BrtEndCommentList",
    0x027B: "BrtBeginComment",
    0x027C: "BrtEndComment",
    0x027D: "BrtCommentText",
    0x0294: "BrtBeginListParts",
    0x0295: "BrtListPart",
    0x0296: "BrtEndListParts",
    0x0400: "BrtArrFmla",
    0x0415: "BrtWsProp14",
}


def read_varint(data: bytes, offset: int, *, max_shift: int) -> tuple[int, int]:
    value = 0
    shift = 0
    while True:
        if offset >= len(data):
            raise ValueError("truncated varint")
        byte = data[offset]
        offset += 1
        value |= (byte & 0x7F) << shift
        if byte & 0x80 == 0:
            return value, offset
        shift += 7
        if shift > max_shift:
            raise ValueError("varint too long")


def records(data: bytes) -> list[tuple[int, int]]:
    out = []
    offset = 0
    while offset < len(data):
        record_id, offset = read_varint(data, offset, max_shift=14)
        payload_len, offset = read_varint(data, offset, max_shift=28)
        end = offset + payload_len
        if end > len(data):
            raise ValueError(f"truncated record 0x{record_id:04X}")
        out.append((record_id, payload_len))
        offset = end
    return out


def is_binary_part(name: str) -> bool:
    return (
        name.endswith(".bin")
        and (
            name == "xl/workbook.bin"
            or name in {"xl/styles.bin", "xl/sharedStrings.bin"}
            or name.startswith("xl/worksheets/")
            or name.startswith("xl/tables/")
            or name.startswith("xl/comments")
            or "/comments" in name
        )
    )


def inventory(path: Path) -> dict[str, collections.Counter[int]]:
    result: dict[str, collections.Counter[int]] = {}
    with zipfile.ZipFile(path) as archive:
        for name in sorted(archive.namelist()):
            if not is_binary_part(name):
                continue
            try:
                parsed = records(archive.read(name))
            except ValueError as exc:
                print(f"{path}:{name}: parse error: {exc}")
                continue
            result[name] = collections.Counter(record_id for record_id, _ in parsed)
    return result


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("paths", nargs="+", type=Path)
    parser.add_argument(
        "--unknown-only",
        action="store_true",
        help="only print records missing from the built-in name map",
    )
    args = parser.parse_args()

    for path in args.paths:
        print(f"== {path} ==")
        for part, counts in inventory(path).items():
            rows = []
            for record_id, count in sorted(counts.items()):
                name = RECORD_NAMES.get(record_id, "UNKNOWN")
                if args.unknown_only and name != "UNKNOWN":
                    continue
                rows.append(f"  0x{record_id:04X} {record_id:4d} {count:4d} {name}")
            if rows:
                print(part)
                print("\n".join(rows))


if __name__ == "__main__":
    main()
