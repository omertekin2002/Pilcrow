# Coverage and design choices

Version 0.1 is a partial prose view. This is neither a visual comparison nor a
lossless Word-to-Markdown converter. There is no import path from prose to DOCX.
Only conventional Transitional OOXML packages with `word/document.xml` are
accepted. Strict OOXML, encrypted Office files, legacy `.doc`, and both single-file
Word XML dialects are unsupported. They must not be relabeled `.docx`.

## Supported view

| Content | v1 behavior |
| --- | --- |
| Paragraphs | Runs joined without artificial spaces; blank line between nonempty paragraphs |
| Headings | Outline level, inherited paragraph styles, and conventional Heading1–9 style IDs; capped at six Markdown levels; heading numbering (e.g. `Chapter 1`) is neither rendered nor recorded in list metadata |
| Lists | Item and indentation boundaries from direct/inherited `numPr`; generic `-` marker in prose; each used list's effective format per used level (`numFmt`, `lvlText`, start value, after `lvlOverride`/`startOverride`) recorded in metadata, together with which list and level each item belongs to, so bullet/number/start changes and moves between lists or levels make companions stale; the recorded level is the paragraph's own `ilvl`, else the level the numbering definition, after `lvlOverride`, associates with its paragraph style (`w:lvl/w:pStyle`), while prose indentation follows only the nearest `numPr`'s `ilvl`, so the two can differ; counter values and bullet glyphs are not reconstructed and trigger a warning; unresolvable formats (no numbering part, unknown list or level, `numStyleLink`) are recorded as `unknown` with a warning; custom formats such as `001, 002, …` record their pattern; bullet fonts, legal numbering, suffixes, and level restarts are not recorded |
| Tables | Explicit table, row, and cell markers, including empty cells; nested tables and content-control wrappers; grid spans/vertical merge markers retained, no rendered grid/layout |
| Footnotes/endnotes | Related parts with reference markers; separator notes omitted; IDs retained, so renumbering can produce diffs |
| Headers/footers | Related parts shown as separate sections by part name; no page/section placement mapping |
| Tracked text | Include `ins` and `moveTo`, omit `del` and `moveFrom`; separate kind, author, and text records, without dates or revision IDs |
| Comments | Text, author, and anchor in metadata, never in prose; `anchor` is the text between a comment's `w:commentRangeStart` and `w:commentRangeEnd` in the accepted view (insertions included, deletions omitted, tabs, line breaks, and paragraph boundaries as single spaces), matched by comment ID in the body, notes, headers/footers, and text boxes, and `anchor_part` is the part holding the range start, so moving, widening, narrowing, or re-attaching a range makes companions stale while Word's comment renumbering, run splits, proofing marks, and bookmarks do not; a reference without a range (or an empty range) records `""`, a comment never referenced in a rendered part records `null` with one coverage warning, and a range missing its start or end runs from the start or to the end of its part with one warning; a text box is a separate story, so a range around its anchor does not cover the box's text and a range inside a box runs at most to the box's edges; from the comments-extended part (`word/commentsExtended.xml`, reached through the document's relationship) each comment records `resolved` (`w15:done`) and `reply_to` (the 0-based position, in comments-part order, of the comment named by `w15:paraIdParent`), matched through the `w14:paraId` of the comment's last paragraph, so resolving, reopening, replying, and re-parenting make companions stale while Word's regenerated comment IDs and paraIds do not; both are `null` when there is no comments-extended part or matching `w15:commentEx`, and `reply_to` is `null` for a top-level comment or a parent that matches no comment; raw IDs, paraIds, and durable IDs are never recorded; dates, rich formatting, `commentsIds.xml`, and `commentsExtensible.xml` not compared |
| Hyperlinks | Accepted visible text in prose; label and destination in metadata, without fetching the URL |
| Fields/citations | Saved display text in prose; normalized instruction fingerprint and first command token in metadata; never recalculate fields or manage Zotero |
| Text boxes | Text inside modern (`mc:AlternateContent`) and legacy VML text boxes, rendered with the normal paragraph/table rules between `[Text box]` and `[/Text box]` markers right after the anchoring paragraph, in document order; one copy only (the first `mc:Choice` holding text, otherwise `mc:Fallback`); works in table cells, headers/footers, and nested boxes; position, anchoring, wrapping, and shape formatting are not compared and trigger one coverage warning; a shape with no text of its own, such as the continuation boxes of a linked-box chain (the first box holds the whole story), is not a text box here and keeps the placeholder and omission warnings in the next row, so how a story flows across linked boxes is not compared |
| Images | Placeholders (e.g. `[Unsupported drawing]`, `[Unsupported pict]`) in prose, as for other drawings; every DrawingML (`a:blip`, including its SVG copy) and legacy VML (`v:imagedata`) image in the body, headers/footers, notes, and text boxes is recorded in metadata by the SHA-256 and size of its media part (external links by target, never fetched), so swapping, adding, removing, or reordering images makes companions stale; renamed media parts with identical bytes do not; size, position, cropping, alt text, and images in comments are not compared and trigger one coverage warning; a missing image relationship or media part is recorded without a hash and warns |
| Drawings, charts, embedded objects, equations | Omission warnings and placeholders (e.g. `[Unsupported drawing]`) for common inline constructs, including a picture grouped with a text box; no rendering, OCR, object execution, or asset-content comparison (including the preview images of embedded objects) |
| Bold/italic, fonts, colors, other styles, pagination/layout | Not compared; always disclosed in coverage warning |

Tracked text inclusion is an **accept-insertions/omit-deletions view**, not Word's
complete Accept All Changes implementation. Paragraph-mark deletion, list/table
revisions, tracked formatting, moves expressed in unusual structures, and nesting
can differ from Word's result. These constructs are reported where detected;
review them in Word. Deleted table rows are omitted. Comment text is plain text;
hyperlinks/fields inside comments are not separately interpreted.

Hidden text is included and triggers a visibility warning; formatting is not used
to decide what Word would display. Unknown Word/XML constructs trigger conservative warnings. Extra package parts
are listed as unrendered; their bytes are not hashed. Media parts hashed as images
(see below) are not listed; media nothing references still is. Some save-only parts
(`docProps`, settings, font tables, theme) are deliberately ignored. Numbering
definitions are read only for the list levels that list items (not headings) use, so edits to
unused definitions and renumbered list IDs do not change companions. Counter
values (which number an item shows) and list layout can still change without a
companion update.
Unsupported assets can change while warning text remains identical. The `partial-prose`
coverage marker and CLI wording apply even when there are no specific warnings.
No claim of a complete/unchanged document is made. This is a modest coverage report,
not an audit of every OOXML relationship or layout semantic.

Metadata contains only schema/view/coverage, warnings, review text/authors, link
destinations, field fingerprints, list formats, and image hashes. `lists` has one entry per
list used by a rendered list paragraph, numbered 1, 2, … by first use (not by
Word's `numId`, which can change on save), with the part where it is first used
and the format, level text, and start value of each level its items use.
`list_runs` records which list and level each list item belongs to, as runs of
consecutive list items (other paragraphs in between are skipped) of the same list
and level in the same part, so moving an item between lists, or between levels
that its paragraph style selects without changing the prose indentation, is a
metadata change; inserting an item changes only one count. `images` has one entry per image reference in
traversal order with its story `part` and either `sha256`/`bytes` of the media
part (internal) or `target` (external link); a dangling reference has a null
`sha256`. Media file names are not recorded because Word renumbers them on save,
so order plus hash is the comparison, and re-encoded images are a real change.
The chosen `mc:AlternateContent` copy is read once: the text box branch, else
the first `mc:Choice`. Field fingerprints avoid huge Zotero JSON
payloads but only say a normalized code changed, not what that change means.
Arrays follow document traversal order; related parts and warnings sort
lexically. Internal part paths come from the package; no machine paths, save
counters, revision dates, or export timestamps are added. Hyperlink destinations
are retained literally and may themselves contain sensitive or local paths.

## Normalization and sentence lines

- UTF-8, LF endings, terminal newline. Unicode, punctuation, numbers, and negations
  are retained. No stemming, case folding, Unicode composition, or spellchecking.
- Word run boundaries and run formatting do not introduce text boundaries.
- After joining runs, whitespace (including tabs and nonbreaking spaces) collapses
  to ordinary single spaces; leading/trailing paragraph whitespace is removed.
  Explicit Word line breaks create separate lines. XML source newlines are spaces.
  Empty paragraphs are omitted. Thus spacing and blank-paragraph layout changes
  are intentionally not reviewed. Nonbreaking and soft hyphen characters remain.
- Sentence breaks occur at `.`, `!`, or `?` followed by whitespace; closing quotes
  or brackets stay with the sentence. A small English abbreviation list and
  single-letter/initialism heuristic suppress common false splits. Decimals such
  as `3.14` have no whitespace and stay intact. Question/exclamation sentences and
  ordinary Unicode text are retained.
- This is English-oriented, not universal sentence detection. Unknown abbreviations
  can split too soon; an abbreviation at sentence end can suppress a valid split.
  CJK text without whitespace stays on one line. Citation punctuation, initials,
  ellipses, and unusual scientific notation may need later tuning. An incorrect
  split changes line boundaries, not the underlying words.
- Heading text stays on one line. Each list item starts with `-`; subsequent
  sentences are indented. Table cells contain normal sentence/paragraph text.
  There are no per-sentence positional IDs or automatic counters.
- The `.md` is optimized for the ordinary text diff, not faithful Markdown rendering:
  literal Markdown metacharacters in Word are not escaped, and table/text box/unsupported
  markers can coincide with literal document text.

## Operational limits

Packages are limited to 4096 entries and 1 GiB total uncompressed (Word's own
documented ceiling for a document with graphics is 512 MB), which bounds
decompression work. Parts that Pilcrow parses into memory (XML and `.rels`) are
limited to 16 MiB each and 64 MiB in total, so a `word/document.xml` over 16 MiB
is rejected. Images are hashed in 1 MiB chunks without holding them in memory, so
large or numerous images count only toward the 1 GiB total. A package over a
limit is an error (exit 2) naming the part or the limit it exceeded. Parts must
use the stored or deflate ZIP methods, the only ones OPC allows and the only
ones read with bounded memory. Duplicate ZIP entries and XML entity expansion
are rejected.
Output/source symlink paths are rejected to avoid writing through unexpected
paths. Export checks source stat information before/after reading; it is not a
cross-process lock. Save completely before export, and rerun after a concurrent
save. No watcher retries or transactional multi-document snapshot is provided.

The output mapping preserves directory paths and the `.docx` extension. It does
not solve case-colliding filenames on case-insensitive filesystems; repositories
shared with Windows/macOS should avoid filenames differing only by case.
Git discovery requires an available working Git installation. A selected file
that disappears is an error; discovery treats deleted tracked files as absent and
reports any existing companions as orphans. Empty discovery can succeed with
“No DOCX documents selected”; verify the scope when that is unexpected.

## Revisit with real documents

1. Validate real Word saves on Windows/macOS and Desktop's companion diffs, including
   save-without-editing and Zotero documents, before claiming GUI compatibility.
2. Tune sentence segmentation against collaborators' languages and abbreviation
   conventions; keep fixtures for every regression.
3. Decide whether reconstructed list counters, comment dates, or field-code previews
   justify more metadata. Avoid a dump of all OOXML.
4. Improve tracked structural revisions, note-reference validation, complex tables,
   and relationship/path canonicalization only as actual documents require it.
5. Opt-in textconv exists (`pilcrow setup --textconv` plus a one-time
   `git config diff.pilcrow.textconv "pilcrow textconv --"`); it shows prose only,
   not metadata, and is untested in GitHub Desktop, where executable discovery
   needs its own validation. Consider refresh-only watching after the primary
   workflow is reliable.
6. If Word XML is wanted later, build a distinct dialect-detecting adapter. Do not
   treat arbitrary XML as Word input. Automatic document merging stays separate.
