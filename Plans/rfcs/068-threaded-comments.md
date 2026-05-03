# RFC-068 — Threaded comments: read + write + modify (Sprint 2 / G08)

> **Status**: Proposed
> **Owner**: Claude (S2 design)
> **Sprint**: S2 — Comments & Rich Text
> **Closes**: G08 (threaded comments) in the openpyxl parity program
> **Depends on**: RFC-023 (legacy comments + VML drawings, shipped)
> **Unblocks**: G09 (rich text in headers/footers) and G10 (rich text in chart labels) can land in parallel — they share the rich-text plumbing but not the comments path

## 1. Goal

Add full Python authoring + Rust write emit + read + modify-mode preservation for Excel 365's threaded comments (`xl/threadedComments/threadedCommentsN.xml`), so an openpyxl-shaped caller can:

```python
from wolfxl.comments import ThreadedComment, Person

wb = wolfxl.Workbook()
ws = wb.active
ws["A1"] = "topic"
alice = wb.persons.add(name="Alice", user_id="alice@example.com")
top = ThreadedComment(text="Looks wrong", person=alice)
reply = ThreadedComment(text="Agreed; investigating", person=alice, parent=top)
ws["A1"].threaded_comment = top
ws["A1"].threaded_comment_replies.append(reply)
wb.save(out)

wb2 = wolfxl.load_workbook(out)
assert wb2.active["A1"].threaded_comment.text == "Looks wrong"
assert wb2.active["A1"].threaded_comment_replies[0].text == "Agreed; investigating"
```

Plus the modify-mode contract: opening an Excel-authored file with threaded comments and saving it without touching the comments must preserve the threading exactly. Editing or deleting a threaded comment must rewrite the threading parts coherently with the legacy placeholder comment.

## 2. Problem statement

Wolfxl today supports legacy comments end-to-end (RFC-023, shipped) but treats threaded comments as opaque metadata:

- **Read**: the reader's `Comment` struct has a `threaded: bool` field at `crates/wolfxl-reader/src/lib.rs:351–356`, but it is always set to `false` — `xl/threadedCommentsN.xml` and `xl/persons/personN.xml` are not parsed.
- **Modify**: the patcher's `ExistingComment` struct (`src/wolfxl/comments.rs:70–71, 222`) preserves the `<extLst>` block on legacy comments, which contains the back-reference GUID Excel uses to find the threaded-comment payload. So **opening + re-saving** an Excel-authored file does not destroy threading metadata. But user-supplied comment edits clear `ext_lst` (line 491), which severs the back-reference — the threadedComments part still exists in the ZIP but is now orphaned.
- **Write**: there is no emit path for `xl/threadedCommentsN.xml` or `xl/persons/personN.xml`. The `[Threaded comment]` placeholder convention (legacy comment whose text is the literal string `[Threaded comment]` while the threadedComment payload carries the real text) is not implemented.
- **Python**: `wolfxl.comments` exposes `Comment` only; there is no `ThreadedComment` class. The compat-oracle probe at `tests/test_openpyxl_compat_oracle.py:819–832` xfails because the public class is missing.

The plan tags G08 "Claude-led" because the OOXML semantics are the hard part: the threadedComment / legacy-comment / extLst tripod must stay coherent across every save, and the person registry is workbook-scoped.

## 3. OOXML primer

### 3.1 Files involved

For a workbook with threaded comments:

```
xl/comments/comments1.xml              <-- legacy comments part (placeholder)
xl/drawings/vmlDrawing1.vml            <-- legacy VML anchors (still required)
xl/threadedComments/threadedComments1.xml   <-- NEW: threaded payload
xl/persons/personList.xml              <-- NEW: workbook-scoped person table
[Content_Types].xml                    <-- adds Override for the new parts
xl/_rels/workbook.xml.rels             <-- adds Relationship to personList
xl/worksheets/_rels/sheetN.xml.rels    <-- adds Relationship to threadedCommentsN
```

The legacy `comments1.xml` survives even when threaded comments are present. Each `<comment>` in it carries:

- The same `ref` cell anchor as its threaded counterpart.
- Text body equal to the literal `[Threaded comment]` (so legacy readers see something coherent).
- An `<extLst>` block whose `ext` element references the threaded-comment GUID (URI `{B7B6F4E0-...}` family).

The threaded payload at `xl/threadedComments/threadedComments1.xml` looks like (ECMA-376 1st-edition extension, schema URL `http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments`):

```xml
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment ref="A1"
                   dT="2024-09-12T15:31:01.42"
                   personId="{8B0E8A60-...}"
                   id="{A1B2-...}">
    <text>Looks wrong</text>
  </threadedComment>
  <threadedComment ref="A1"
                   dT="2024-09-12T15:33:00.00"
                   personId="{8B0E8A60-...}"
                   parentId="{A1B2-...}"
                   id="{C3D4-...}">
    <text>Agreed; investigating</text>
  </threadedComment>
</ThreadedComments>
```

The person registry at `xl/persons/personList.xml`:

```xml
<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <person displayName="Alice"
          id="{8B0E8A60-...}"
          userId="alice@example.com"
          providerId="None"/>
</personList>
```

### 3.2 Coherence rules

A wolfxl save must satisfy all of:

1. Every threadedComment has a non-empty `id` (GUID, lowercase + braces) and a `personId` that resolves in personList.
2. Every threadedComment with a `parentId` has a sibling top-level threadedComment whose `id` matches the `parentId`. Replies are flat siblings; the parent/child relationship is via GUID ref, not XML nesting.
3. For every `ref="<coord>"` in threadedComments, the legacy `comments.xml` has a matching `<comment ref="<coord>">` whose body is `[Threaded comment]` and whose `<extLst>` contains the topmost threaded GUID for that anchor.
4. Every `personId` referenced from any threadedComment exists in personList.
5. The legacy comment's author column references a synthetic placeholder author (e.g., `tc={threadGuid}`) per Excel's convention. This avoids forcing wolfxl to invent a free-text author when the real author is already in personList.

## 4. Public Python contract

### 4.1 New classes

```python
# python/wolfxl/comments/_threaded_comment.py (NEW)
@dataclass
class ThreadedComment:
    """Excel 365 threaded comment.

    Two-tier model:
    - top-level: parent is None; lives at ws[coord].threaded_comment
    - reply: parent is a ThreadedComment; lives at
      ws[coord].threaded_comment_replies (top-level's reply list)
    """
    text: str
    person: "Person"
    parent: "ThreadedComment | None" = None
    created: datetime | None = None  # auto-now if None at flush
    done: bool = False
    id: str | None = None  # auto-allocated GUID at flush if None
```

```python
# python/wolfxl/comments/_person.py (NEW)
@dataclass
class Person:
    name: str
    user_id: str = ""           # mirrors openpyxl
    provider_id: str = "None"   # the literal string "None" matches Excel default
    id: str | None = None       # auto-allocated GUID at registration
```

### 4.2 Workbook + Worksheet additions

- `wb.persons` — a registry exposing `add(name=..., user_id=...) -> Person`, indexing by `id`. Insertion-order preserving. Mirrors openpyxl's `wb._person_list`.
- `ws[coord].threaded_comment` — getter + setter. Setting to `None` removes the threaded comment and its replies. Setting to a `ThreadedComment` instance whose `parent` is non-None raises `ValueError("threaded_comment must be a top-level comment; reply via threaded_comment_replies")`.
- `ws[coord].threaded_comment_replies` — a list-like that proxies to the top-level's reply list. Adding to it without a top-level threaded_comment raises.

### 4.3 Coexistence with `cell.comment`

Setting `cell.threaded_comment` on a cell that already has a legacy `cell.comment` raises `ValueError("cell already has a legacy comment; remove it before adding a threaded comment")`. Going the other direction is also disallowed — Excel itself does not support both on the same cell coherently. (Decision: refuse the conflict explicitly. Open question §11.1 if user evidence later argues otherwise.)

### 4.4 Read model

`load_workbook(path)` populates:

- `wb.persons` from the parsed `personList.xml` (or empty if none).
- `ws[coord].threaded_comment` and `ws[coord].threaded_comment_replies` from the parsed `threadedCommentsN.xml`. Reply chains are reassembled by `parentId` GUID matching.
- `ws[coord].comment` is set to the legacy placeholder comment object **only if** the user opens with `read_only=True` or with a flag explicitly requesting placeholder visibility. Otherwise (default), the legacy placeholder is suppressed at the Python layer to match openpyxl's behaviour: wolfxl shows the threaded version and hides the placeholder.

## 5. Rust model + emit

### 5.1 New writer modules

- `crates/wolfxl-writer/src/model/threaded_comment.rs` (NEW) — `ThreadedComment`, `Person`, `PersonTable`. Field shapes mirror Python.
- `crates/wolfxl-writer/src/emit/threaded_comments_xml.rs` (NEW) — emits `xl/threadedComments/threadedCommentsN.xml`.
- `crates/wolfxl-writer/src/emit/persons_xml.rs` (NEW) — emits `xl/persons/personList.xml`. (Note: Excel always uses singular `personList.xml`, never numbered. The personList is workbook-scoped, not sheet-scoped.)

### 5.2 Existing modules to extend

- `crates/wolfxl-writer/src/emit/comments_xml.rs` — when emitting a legacy comment that has an associated threadedComment, write the placeholder body `[Threaded comment]` and an `<extLst>` block referencing the threaded GUID. Authors table must include the synthetic `tc={guid}` author.
- `crates/wolfxl-writer/src/emit/content_types_xml.rs` — register the two new content types when at least one threadedComment exists in the workbook.
- `crates/wolfxl-writer/src/emit/workbook_rels.rs` — add the workbook→personList relationship.
- `crates/wolfxl-writer/src/emit/sheet_rels.rs` — add the sheet→threadedCommentsN relationship per sheet that has any.

### 5.3 Reader changes

- `crates/wolfxl-reader/src/threaded_comments.rs` (NEW) — parse the two new XML parts. Populate `Comment.threaded = true` on the legacy comment that carries the placeholder, and emit a separate `ThreadedComment` payload list keyed by sheet + cell.
- `src/native_reader_backend.rs` — surface a new `read_threaded_comments(sheet) -> list[dict]` PyO3 entry point alongside `read_comments`. The Python Worksheet hydrates its threaded-comment view from this output.

### 5.4 Modify-mode (patcher)

The existing `src/wolfxl/comments.rs` already preserves `ext_lst` on legacy comments through round-trips that do not touch comments. Extend the patcher with:

- `queue_threaded_comment(sheet, cell, payload)` — analogous to `queue_comment`. Payload includes the GUID, parentId (if reply), personId, dT, text.
- `queue_threaded_comment_delete(sheet, cell)` — drops the threadedComment(s) at that cell and recomputes the legacy placeholder/extLst. Replies are deleted too (deleting the parent is the only way to delete a thread).
- `queue_person(payload)` — workbook-level; adds to personList.

The patcher merge logic must keep legacy + threaded coherent: if the user assigns a `Comment` (legacy) to a cell that previously had a threaded comment, the patcher drops the threadedComment and clears extLst. If the user assigns a `ThreadedComment` to a cell that previously had a legacy comment, the patcher rewrites the legacy comment to be a placeholder and adds the extLst back-reference.

## 6. Coherence invariants (enforced at flush time)

Wolfxl's flush layer (write mode and modify mode both) must validate before saving:

1. Every `parent` in a `ThreadedComment` is itself a top-level threadedComment in the same sheet at the same cell.
2. Every `person` in a `ThreadedComment` is registered in `wb.persons`.
3. No cell has both a legacy `Comment` and a `ThreadedComment` (raise at assignment time too — see §4.3).
4. Reply timestamps are non-decreasing relative to their parent's `created`. (Soft warning, not hard error — Excel does not strictly enforce this either.)

Failure of (1), (2), (3) raises `ValueError` at save time with a precise sheet/cell/issue message. (4) emits a `UserWarning`.

## 7. Acceptance criteria

1. Compat-oracle probe `comments_threaded` flips xfail → passed. The probe is **strengthened** by this RFC's impl PR to actually exercise add + reload + assert text equality (currently it only checks class existence; the strengthening also goes into this PR).
2. New focused test `tests/test_threaded_comments.py` covers:
   - Top-level threaded comment round-trip through wolfxl read.
   - Reply chain (one parent + two replies) round-trip; reply order preserved by `dT`.
   - Modify-mode preservation: open an Excel-authored fixture, save unchanged, assert byte-for-byte semantic equivalence (GUIDs preserved, persons preserved).
   - Modify-mode edit: load, change one threaded comment's text, save, reassert the rest of the threading metadata is intact.
   - Conflict guard: setting both `cell.comment` and `cell.threaded_comment` raises with a precise message.
3. New focused test `tests/test_threaded_comments_via_openpyxl.py` reloads wolfxl-saved files through openpyxl and asserts openpyxl exposes the same threaded comment + replies + persons.
4. `cargo test -p wolfxl-writer` covers the two new emitters; `cargo test -p wolfxl-reader` covers the new parser.
5. `tests/test_external_oracle_preservation.py` with threaded-comments fixtures stays green (LibreOffice + Excel round-trips).
6. Compat-oracle pass count rises by exactly 1 (G08 closure).
7. The compat-matrix row `comments.threaded` flips to `supported`. Any other rows tagged G08 (currently only this one) flip likewise.

## 8. Out-of-scope

- @-mentions inside threaded-comment text. Excel models mentions as a `<mentions>` block referencing a personId; openpyxl exposes them. This RFC defers them — text bodies are plain strings for v1. A follow-up small RFC adds mention parsing.
- Threaded-comment "resolved" / "done" UX semantics beyond round-tripping the `done` boolean. Excel renders "done" comments greyed out; wolfxl just preserves the flag.
- ProviderId values other than the default `"None"`. Microsoft accounts and AAD identities have richer providerId semantics; out of scope for v1.
- Bulk-conversion shortcut (`ws.upgrade_legacy_comments_to_threaded()`). If users need this, ship as a follow-up.
- Modifying threaded comments from the patcher in pre-1.0 versions of openpyxl that lack the `ThreadedComment` class. Wolfxl matches openpyxl 3.1+.

## 9. Risks

| # | Risk | Mitigation |
|---|------|-----------|
| 1 | The placeholder author convention `tc={guid}` is undocumented; Excel may stop honouring it. | Round-trip a bank of Excel-authored fixtures through wolfxl-modify and verify Excel still opens the result cleanly. The external-oracle preservation pack adds a threaded-comments slice in this RFC's PR. |
| 2 | Person GUIDs assigned by wolfxl drift across reload — wb.persons.add() returns a new Person with a fresh GUID even if "the same" person already exists. | Make `add()` idempotent on `(user_id, provider_id)` if both are non-empty; otherwise allocate a new GUID. Document the rule. |
| 3 | The patcher's existing `ext_lst` preservation might conflict with the new emit logic. | Make the patcher prefer wolfxl-controlled `<extLst>` whenever a threaded-comment write is queued for that cell. Existing preservation kicks in only when no edit is queued. |
| 4 | Sheet cloning (RFC-035) was designed before threaded comments existed. Cloning a sheet that has threaded comments must rewrite GUIDs. | Add a small RFC-035 patch in this PR: when copy_worksheet runs and threaded comments exist, regenerate every `id`/`parentId`/extLst back-ref so the clone gets fresh GUIDs while preserving the parent/child structure. |
| 5 | LibreOffice's threaded-comments support is partial; round-tripping through LibreOffice may strip threadedComments while preserving legacy. | Document this in `docs/trust/limitations.md`; the external-oracle preservation gate must accept "LibreOffice strips threading; legacy survives" as a tolerated diff. Excel round-trip is the authoritative gate. |
| 6 | `personList.xml` is workbook-scoped but the existing patcher has no concept of workbook-level XML parts beyond rels and shared strings. | Add a small workbook-scoped queue; mirror how `wb.security` flushes (RFC-058) — that pattern already exists. |

## 10. Implementation plan

1. RFC review + approval (this document).
2. Land the Python class layer (`ThreadedComment`, `Person`, `wb.persons`) behind an internal-only feature for testing — no Rust changes yet, no public re-export.
3. Land the writer model + emitters (`threaded_comments_xml.rs`, `persons_xml.rs`); plumb through `wb.save()` for write mode only. Add tests against a known-good Excel-authored fixture.
4. Land the reader parser. The reader unblocks round-trip tests through wolfxl alone.
5. Land the patcher modify-mode entry points. Modify-mode preservation tests come online here.
6. Add the openpyxl-via-reload tests; verify openpyxl reads wolfxl-saved files equivalently.
7. Strengthen the compat-oracle probe (in the same PR as step 6) to do real round-trip assertions, not class-existence.
8. RFC-035 patch for sheet cloning + threaded comments.
9. Flip spec; regenerate matrix; mark G08 `landed`.

A codex handoff under `Plans/rfcs/handoffs/G08-threaded-comments.md` may be derived from §4 + §5 once this RFC is approved — but the cross-cutting patcher coherence work likely keeps this gap Claude-led, with codex used only for the bounded emit + reader tasks (steps 3 + 4 of §10).

## 11. Open questions

1. Mixing legacy and threaded comments on the same cell: Excel itself does not allow this. Wolfxl's stance (refuse at assignment time) matches Excel and openpyxl. Confirm before impl.
2. ProviderId default: Microsoft uses `"None"` (literal string) when no identity provider is attached. Wolfxl mirrors. Confirm before impl.
3. Reply timestamp ordering: enforce monotone increase, or just preserve what the user provided? Default proposal: preserve, warn.
4. Do we expose `cell.threaded_comment_replies` as an alias for `cell.threaded_comment.replies`, or only the latter? openpyxl uses the latter — match it; drop the alias from §4.2.
