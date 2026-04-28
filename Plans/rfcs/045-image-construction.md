# RFC-045: Image construction — `wolfxl.drawing.image.Image` (replace stub)

Status: Shipped 1.5 (Sprint Λ Pod-β) — writer `0ace8c5`, Image+add_image `d9cb569`, tests `a73737e`
Owner: Sprint Λ Pod-β
Phase: 5 (1.5)
Estimate: L
Depends-on: RFC-010 (rels graph), RFC-013 (patcher infra: `file_adds`, content-types ops, two-phase flush)
Unblocks: T3 closure for image construction; chart-construction prerequisites (RFC-Μ); openpyxl-parity surface for `Image()` callers

> **S** = ≤2 days; **M** = 3-5 days; **L** = 1-2 weeks; **XL** = 2+ weeks
> (calendar, with parallel subagent dispatch + review).

## 1. Background — Problem Statement

`python/wolfxl/drawing/image.py` is currently:

```python
"""Shim for ``openpyxl.drawing.image``."""

from __future__ import annotations
from wolfxl._compat import _make_stub

Image = _make_stub(
    "Image",
    "Images are preserved on modify-mode round-trip but cannot be added programmatically.",
)

__all__ = ["Image"]
```

`_make_stub` returns a class whose `__init__` raises
`NotImplementedError("Image is preserved on modify-mode round-trip
but cannot be added programmatically.")`. Every user code path that
constructs an `Image(...)` raises immediately:

```python
from wolfxl.drawing.image import Image
img = Image("logo.png")            # NotImplementedError today
ws.add_image(img, "B2")
```

Modify-mode workbooks **already** preserve existing image parts
(`xl/media/imageN.png` is copied verbatim during round-trip — RFC-013
content-types graph and rels graph carry the references through).
What's missing is the **construction** path: there's no way to add a
new image programmatically, in either write mode or modify mode.

**Target behaviour**: `Image(...)` is a real, openpyxl-shaped class.
`Worksheet.add_image(img, anchor)` accepts it and routes the bytes
through the appropriate emit pipeline (native writer in write mode,
patcher's `file_adds` in modify mode). Supports PNG, JPEG, GIF, and
BMP formats. Supports one-cell, two-cell, and absolute anchor types.
Width and height auto-detected from image headers (callers can
override). Round-trips end-to-end via wolfxl read → write → re-read,
and via openpyxl interop in both directions.

## 2. Architecture

The image lifecycle has three pieces — anchor + bytes + ZIP wiring.
The implementation splits across three layers (Python coordinator,
patcher, native writer) to mirror the existing pattern from RFC-035
§5.2-5.3 (which clones drawings as part of `copy_worksheet`).

### 2.1 Write mode (native writer emits drawingN.xml + media + rels)

`crates/wolfxl-writer` already emits the bones of `xl/drawings/`
during `copy_worksheet` write-mode (Sprint Θ Pod-C1). Pod-β extends
the writer's emit pipeline with a fresh "images" pass:

1. The Python `Worksheet._pending_images: list[(Image, Anchor)]`
   queue holds user-added images for the worksheet.
2. On `Workbook.save()`, the Python coordinator hands the queue to
   the writer via `NativeWorkbook.add_image_to_sheet(sheet_idx,
   bytes, ext, anchor_dict)`.
3. The writer's drawings emitter:
   * Allocates `xl/media/imageN.<ext>` via `PartIdAllocator` (the
     centralized allocator from RFC-035 §5.2).
   * Writes the image bytes verbatim into the output ZIP.
   * Allocates `xl/drawings/drawingN.xml` and emits the
     `<xdr:wsDr>` root with one anchor child per image on the sheet.
   * Allocates `xl/drawings/_rels/drawingN.xml.rels` with one
     `<Relationship Type=".../image">` per anchor.
   * Adds a `<drawing r:id="..."/>` child to the worksheet's
     `<sheetData>`-sibling block.
   * Registers `<Default Extension="png" ContentType="image/png"/>`
     in `[Content_Types].xml` (one per distinct extension).
   * Adds a `<Relationship Type=".../drawing">` to the sheet's rels
     graph pointing at the new drawingN.xml.

### 2.2 Modify mode (patcher routes new images through file_adds)

Modify mode reuses the same drawings-emit logic via a new patcher
PyMethod `queue_image_add`:

1. The Python coordinator calls `patcher.queue_image_add(sheet_path,
   bytes, ext, anchor_payload)` per `add_image` invocation.
2. A new Phase 2.5j (sequenced after Phase 2.5g comments / 2.5f
   tables and before 2.5c content-types aggregation) drains the
   queue per sheet:
   * Reuses `PartIdAllocator` to pick fresh `imageN.<ext>` and
     `drawingN.xml` numbers — collision-free against the source ZIP
     listing AND any in-flight `file_adds` from RFC-035 sheet copies
     or other features.
   * Emits the new `imageN.<ext>` part bytes into `file_adds`.
   * Builds or extends the sheet's `xl/drawings/drawingN.xml` —
     **extending** if the sheet already has a drawing part (just
     append a new `<xdr:twoCellAnchor>` / `<xdr:oneCellAnchor>` /
     `<xdr:absoluteAnchor>` to the existing `<xdr:wsDr>` root via
     a quick-xml splice); **creating** otherwise.
   * Emits the corresponding `xl/drawings/_rels/drawingN.xml.rels`
     entries for new image rels.
   * If a new drawing was created, splices a `<drawing
     r:id="..."/>` into the sheet XML and adds a
     `<Relationship Type=".../drawing">` entry to the sheet's rels
     graph.
   * Pushes content-type ops onto `queued_content_type_ops` for any
     new file extensions or new drawing parts.

The path **does not** mutate any existing image bytes or drawings —
new images are additive. A future "replace existing image" RFC can
build on this seam.

### 2.3 Anchor types

OOXML drawings support three anchor flavors (ECMA-376 §20.5.2.16):

| Anchor | XML element | Use case |
|---|---|---|
| One-cell | `<xdr:oneCellAnchor>` | Image pinned to a top-left cell, sized in EMUs (English Metric Units, 914,400 EMU = 1 inch). Resizes with column width but keeps its own dimensions. |
| Two-cell | `<xdr:twoCellAnchor>` | Image pinned between a top-left and a bottom-right cell. Resizes with both. Has two sub-modes: `editAs="oneCell"` and `editAs="twoCell"` (we default to `twoCell`). |
| Absolute | `<xdr:absoluteAnchor>` | Image pinned to absolute EMU coordinates, ignores cell layout. Used for floating logos. |

Each anchor uses `<xdr:from>` and (where applicable) `<xdr:to>`
markers, each carrying `<xdr:col>`, `<xdr:colOff>`, `<xdr:row>`,
`<xdr:rowOff>` children. Offsets are EMU values; the helpers
`pixels_to_EMU(...)` (already in `wolfxl.utils.units`) and
`points_to_EMU(...)` translate from the user-friendly units.

The default anchor (when `ws.add_image(img, "B2")` is called with a
plain coordinate string) is **one-cell at the top-left of B2**, with
image's own width/height as the size. This matches openpyxl's
`worksheet/_drawing.py` default.

## 3. Image format support

Magic-byte sniffing identifies the format from the first few bytes
of the image stream. Width and height are extracted from
format-specific header fields so callers don't have to specify them
explicitly.

| Format | Magic bytes | Header parser | Notes |
|---|---|---|---|
| **PNG** | `89 50 4E 47 0D 0A 1A 0A` | IHDR chunk at offset 16 — `width: u32 BE`, `height: u32 BE` | Most common; lossless. |
| **JPEG** | `FF D8 FF` | SOF0/SOF1/SOF2 marker scan — `height: u16 BE`, `width: u16 BE` after marker length + precision byte | Most common photo format. |
| **GIF** | `47 49 46 38 37 61` (GIF87a) or `47 49 46 38 39 61` (GIF89a) | LSD (Logical Screen Descriptor) at offset 6 — `width: u16 LE`, `height: u16 LE` | Animated GIF first-frame is what Excel renders. |
| **BMP** | `42 4D` (`BM`) | DIB header at offset 14 — `width: i32 LE`, `height: i32 LE` (height may be negative for top-down) | Rare in Excel but valid per spec. |

The sniffer lives in `crates/wolfxl-writer/src/image_meta.rs` (new
module) and is exposed via PyO3 as `wolfxl._rust.classify_image(bytes)
-> dict[str, int | str]` returning `{format, width, height}`. The
Python `Image` class consumes this on construction and caches the
metadata for the writer/patcher to read at flush time.

Unrecognized magic bytes raise `ValueError("Unsupported image format;
supported: PNG, JPEG, GIF, BMP. First 8 bytes: ...")`. WebP, TIFF,
and SVG are out of scope (see §8); the error message points users
at the supported list.

## 4. Public API

```python
from io import BytesIO
from wolfxl.drawing.image import Image
from wolfxl.drawing.spreadsheet_drawing import (
    OneCellAnchor, TwoCellAnchor, AbsoluteAnchor, AnchorMarker,
)
from wolfxl.drawing.xdr import XDRPoint2D, XDRPositiveSize2D

# Construct from a path
img = Image("logo.png")

# Construct from BytesIO
img = Image(BytesIO(open("logo.png", "rb").read()))

# Inspect auto-detected metadata
img.format        # → "png"
img.width         # → 320  (pixels)
img.height        # → 100  (pixels)

# Add to a sheet — three flavors
ws.add_image(img, "B2")                              # one-cell anchor at B2 (default)
ws.add_image(img, anchor=TwoCellAnchor(             # two-cell anchor
    _from=AnchorMarker(col=1, colOff=0, row=1, rowOff=0),
    to=AnchorMarker(col=5, colOff=0, row=10, rowOff=0),
))
ws.add_image(img, anchor=AbsoluteAnchor(             # absolute anchor (floating)
    pos=XDRPoint2D(x=914400, y=914400),              # 1 inch from top-left
    ext=XDRPositiveSize2D(cx=2438400, cy=914400),    # 2.67" wide × 1" tall
))
```

`Image.__init__(self, img: str | os.PathLike | BytesIO | bytes)` —
matches openpyxl's surface. Path-based construction reads the file
into memory eagerly (so the `Image` is independent of the source
file post-construction). BytesIO and bytes are sniffed in-place.

`Worksheet.add_image(img: Image, anchor: str | OneCellAnchor |
TwoCellAnchor | AbsoluteAnchor | None = None) -> None` — string
anchors are coordinate shorthand, parsed via the existing
`utils.cell.coordinate_to_tuple`.

The full openpyxl `Image()` surface is targeted for parity. Anchor
helper classes (`OneCellAnchor`, `TwoCellAnchor`, `AbsoluteAnchor`,
`AnchorMarker`, `XDRPoint2D`, `XDRPositiveSize2D`) live under
`python/wolfxl/drawing/spreadsheet_drawing.py` and
`python/wolfxl/drawing/xdr.py` to mirror openpyxl's module layout
(callers' `from openpyxl.drawing.spreadsheet_drawing import
TwoCellAnchor` becomes a one-line search-replace).

## 5. Part-id allocation

`crates/wolfxl-rels/src/part_id_allocator.rs` already centralizes
suffix allocation for `xl/{tables,comments,drawings,worksheets,
printerSettings}/*` (introduced in RFC-035 §5.2). Pod-β extends it
with `alloc_image(extension: &str) -> u32` and `alloc_drawing()
-> u32` (the latter already exists; image is new). The allocator's
"existing names" set is built from the source ZIP listing PLUS any
in-flight `file_adds` so multiple `add_image` calls in the same
save AND a concurrent RFC-035 sheet copy that clones drawings all
get collision-free numbers.

See RFC-035 §5.2 for the existing pattern. RFC-045 reuses the same
struct verbatim — no allocator refactor, just a new `alloc_image`
method and a per-extension counter (`xl/media/image1.png` and
`xl/media/image1.jpeg` are different files because Excel disambiguates
by extension; the allocator therefore tracks a counter per extension
to match Excel's emit conventions).

## 6. Content-types + rels

Two content-type updates per new image, dispatched through RFC-013's
`queued_content_type_ops`:

* `<Default Extension="png" ContentType="image/png"/>` (or `jpeg`,
  `gif`, `bmp`) — `ensure_default("png", "image/png")`. Idempotent
  per extension.
* `<Override PartName="/xl/drawings/drawingN.xml"
  ContentType="application/vnd.openxmlformats-officedocument.
  drawing+xml"/>` — `add_override(...)`. One per new drawing part
  (NOT per image — many images can share a drawing).

Two rels updates per new image:

* In the **drawing's** rels (`xl/drawings/_rels/drawingN.xml.rels`):
  ```xml
  <Relationship Id="rIdN" Type="http://schemas.openxmlformats.org/
    officeDocument/2006/relationships/image"
    Target="../media/imageN.png"/>
  ```

* In the **sheet's** rels (`xl/worksheets/_rels/sheetK.xml.rels`):
  ```xml
  <Relationship Id="rIdM" Type="http://schemas.openxmlformats.org/
    officeDocument/2006/relationships/drawing"
    Target="../drawings/drawingN.xml"/>
  ```
  Only one of these per sheet — multiple images on the same sheet
  share a single drawing part, so the sheet→drawing rel is
  per-sheet, not per-image.

The sheet's `<sheetData>` sibling gets a `<drawing r:id="rIdM"/>`
child (or already has one if a drawing existed; in that case we
extend the existing drawing rather than creating a second one).

## 7. Test plan

New test files:

* `tests/test_images_write.py` — write-mode coverage. Cases:
  * `Image(path)` constructs cleanly; `format`, `width`, `height`
    auto-detected for PNG / JPEG / GIF / BMP.
  * `Image(BytesIO)` and `Image(bytes)` work the same way.
  * `Image("garbage.bin")` raises `ValueError` with the supported-list
    hint.
  * `ws.add_image(img, "B2")` round-trips: write → re-read with
    wolfxl, assert image bytes match input.
  * Two images on the same sheet share a single `xl/drawings/
    drawing1.xml` part.
  * Two-cell and absolute anchors round-trip.
  * Mixed PNG + JPEG on the same sheet emits separate `<Default>`
    content-types for each extension.

* `tests/test_images_modify.py` — modify-mode coverage. Cases:
  * Open a fixture without images; add one; save; re-open; assert
    the image is present.
  * Open a fixture WITH existing images; add a new one; save;
    assert both old and new are present and intact.
  * Open a fixture; add an image to two different sheets; assert
    each sheet gets its own drawing part with collision-free numbers.
  * Compose with RFC-035 `copy_worksheet`: copy a sheet that has an
    image, then add a new image to the copy; assert the copy's
    drawing references both the aliased original image AND the new
    one.

* `tests/parity/test_images_parity.py` — openpyxl interop. Cases:
  * Build a workbook in wolfxl with `Image()` + `add_image()`,
    save, re-read with openpyxl: image bytes + dimensions + anchor
    coordinates all match.
  * Build a workbook in openpyxl with `Image()` + `add_image()`,
    save, re-read with wolfxl: same.

Test fixtures: `tests/fixtures/images/` houses the PNG / JPEG /
GIF / BMP samples (small, ~1 KB each, hand-crafted via Pillow at
test-build time so we don't bloat the repo).

Verification matrix:

| Layer | Coverage |
|---|---|
| 1. Rust unit tests | `cargo test -p wolfxl-writer image_meta` covers magic-byte sniffing + dim extraction for all four formats, plus malformed inputs. |
| 2. Golden round-trip (diffwriter) | `tests/diffwriter/cases/images.py` — `WOLFXL_TEST_EPOCH=0` golden for "1 PNG, two-cell anchor" and "PNG + JPEG mix". |
| 3. openpyxl parity | `tests/parity/test_images_parity.py` (above). |
| 4. LibreOffice cross-renderer | Manual: open the wolfxl-emitted file in LibreOffice, screenshot the image renders correctly. PR description. |
| 5. Cross-mode | `tests/test_images_write.py` + `tests/test_images_modify.py` cover both. |
| 6. Regression fixture | `tests/fixtures/images/source_with_logo.xlsx` (one PNG, two-cell anchor) — checked in. |

## 8. Divergences from openpyxl

**None expected for v1.5.0.** openpyxl's `Image` and anchor classes
are well-defined and well-tested; Pod-β targets behavioral parity at
the byte and Python-API levels. Anything Pod-β surfaces during
implementation should be added here as a divergence row before
merge.

**What's NOT shipped in 1.5** (out of scope):

* **WebP, TIFF, SVG, EMF, WMF.** Excel supports more formats than
  PNG/JPEG/GIF/BMP, but those four cover ~99% of real-world spreadsheet
  images. WebP support in particular is gated on Excel 2021+; SVG
  requires an embedded SVG parser. Tracked for a future RFC.
* **Image transformations** — rotation, cropping, brightness /
  contrast adjustments. These are encoded in `<a:blipFill>` /
  `<a:srcRect>` and `<a:lum>` / `<a:bright>` elements; preserving
  them on round-trip is straightforward but constructing them is a
  separate API surface.
* **Image hyperlinks** (`<a:hlinkClick>` on the picture) — clicking
  the image in Excel jumps to a URL or a sheet location. RFC-022's
  hyperlink infrastructure is the natural seam; deferred to a
  follow-up.
* **Replace / delete existing images.** v1.5.0 is **additive only** —
  `add_image` adds, but there's no `remove_image` or
  `replace_image` API. Tracked as a follow-up.
* **Picture metadata** — alt text, title, description (`<xdr:nvPicPr>`
  attrs). These are accessibility features; Excel exposes them in
  the right-click menu. Adding the full surface is straightforward
  but not in scope for the construction-stub-replacement slice.

## 9. SHA log

Sprint Λ Pod-β landed in three atomic commits on `feat/sprint-lambda-pod-beta`:

- `0ace8c5` `feat(writer): emit drawing parts + media for ws.add_image (RFC-045)`
- `d9cb569` `feat(images): real wolfxl.drawing.Image + ws.add_image for write & modify (RFC-045)`
- `a73737e` `test(images): write/modify/parity coverage + surface entries (RFC-045)`

Merged to `feat/native-writer` as merge commit `7dc00d2` on 2026-04-26.

## Acceptance

(Filled in after Pod-β merges.)

- Commit: `d9cb569` (Pod-β Image+add_image); merged via `7dc00d2` on `feat/native-writer`
- Verification: `pytest tests/test_images_write.py tests/test_images_modify.py tests/parity/test_images_parity.py` GREEN — 24 cases passing (requires `pip install pillow` for openpyxl-side parity assertions)
- Date: 2026-04-26
