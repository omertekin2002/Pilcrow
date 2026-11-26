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
| Legacy form fields | A `FORMCHECKBOX` or `FORMDROPDOWN` field's state, which Word keeps in the `w:ffData` of the field's begin `w:fldChar` and never writes as display text, is recorded in its field entry as `form`, so ticking a box or choosing another dropdown entry makes companions stale; a checkbox records its `name` and `checked` (`w:checked`, else `w:default`, else unchecked; an element without `w:val` is checked, as for any on/off value), and a dropdown its `name`, its `entries` (`w:listEntry`, in order) and the `selected` entry's text (the zero-based `w:result`, else `w:default`, else 0; `null` when that index names no entry); the state is not shown in prose; macros, help and status text, enabled state, and checkbox size are not compared and no longer warn; a text input (`FORMTEXT`) records no form state, because its value is already its display text, and its settings (`w:textInput`: type, default, maximum length, format) keep their unsupported-construct warnings; a form field inside a tracked deletion is not recorded, like any field there; content-control (`w:sdt`) checkboxes need no form state, because their glyph is prose text |
| Text boxes | Text inside modern (`mc:AlternateContent`) and legacy VML text boxes, rendered with the normal paragraph/table rules between `[Text box]` and `[/Text box]` markers right after the anchoring paragraph, in document order; one copy only (the first `mc:Choice` holding text, otherwise `mc:Fallback`); works in table cells, headers/footers, and nested boxes; position, anchoring, wrapping, and shape formatting are not compared and trigger one coverage warning; a shape with no text of its own, such as the continuation boxes of a linked-box chain (the first box holds the whole story), is not a text box here and keeps the placeholder and omission warnings in the next row, so how a story flows across linked boxes is not compared |
| Images | Placeholders (e.g. `[Unsupported drawing]`, `[Unsupported pict]`) in prose, as for other drawings; every DrawingML (`a:blip`, including its SVG copy) and legacy VML (`v:imagedata`) image in the body, headers/footers, notes, and text boxes is recorded in metadata by the SHA-256 and size of its media part (external links by target, never fetched), so swapping, adding, removing, or reordering images makes companions stale; renamed media parts with identical bytes do not; each image's alt text (`wp:docPr/@descr`, else the picture's own `pic:cNvPr/@descr`; a VML shape's `alt`) and alt title (`@title`, recorded apart) is recorded with it, so editing, adding, or removing alt text makes companions stale, while a group's alt text and alt text of charts, diagrams and other non-picture drawings are not recorded; size, position, cropping, and images in comments are not compared and trigger one coverage warning; a missing image relationship or media part is recorded without a hash and warns |
| Equations and symbols | Each equation is shown in prose as deterministic linear text, `[Math: …]` (see [Equations as linear text](#equations-as-linear-text)), in its paragraph (inline `m:oMath`/`m:oMathPara`) or as a paragraph of its own (math directly in the body or a table cell), with one coverage note that layout, fonts, and alignment are not compared; `w:sym` keeps the `[Unsupported sym]` placeholder and its omission warning; each outermost OMML equation (`m:oMathPara`, or an `m:oMath` outside one) and each `w:sym` in the body, headers/footers, notes, and text boxes is recorded in metadata, so editing, adding, removing, or reordering one makes companions stale; an equation is recorded by the SHA-256 of a canonical form of its accepted view (tracked deletions dropped, insertions kept; run formatting `w:rPr`, control formatting `m:ctrlPr`, `rsid*`/`w14:paraId`/`w14:textId` attributes, proofing marks, bookmarks, and comment markers left out; `m:t` text and the other `m:*Pr` structure kept), so the fingerprint covers OMML structure and text, not how Word draws it, and a formatting-only re-save stays fresh; a symbol is recorded by its font and character code; equations inside tracked deletions or comments are not recorded, and only the chosen `mc:AlternateContent` copy is read, as for images |
| Charts, SmartArt, embedded objects, imported chunks | Placeholders (`[Unsupported drawing]`, `[Unsupported object]`, `[Unsupported altChunk]`) and omission warnings in prose; each DrawingML chart (`c:chart`), SmartArt diagram (`dgm:relIds`), OLE object (`o:OLEObject`) and imported chunk (`w:altChunk`) in the body, headers/footers, notes, and text boxes is recorded in metadata by the SHA-256 of the raw bytes of the part it points to (the chart part, which holds the cached values a reader sees; the diagram's data part, `r:dm`; the embedding; the chunk), so editing, adding, removing, or reordering one makes companions stale; renaming a part with identical bytes does not; the fingerprint covers the stored part, not how Word draws it, so a Word re-save that rewrites the part without a visible edit also makes companions stale; a chart's embedded workbook, a diagram's layout, quick-style, and colour parts, and the preview images of embedded objects are not compared and keep their unrendered-part warnings; a linked OLE object (external relationship) is recorded by target and never fetched; a missing relationship or part is recorded without a hash and warns; objects in tracked deletions or comments are not recorded, and only the chosen `mc:AlternateContent` copy is read, as for images, so an Office 2016 chart (`cx:chart`, e.g. waterfall or sunburst) whose compatibility copy is a text box is not recorded and its part keeps its unrendered-part warning; a part that cannot be read (e.g. corrupt compressed data) fails the export, as for images |
| WordArt | Legacy VML WordArt (`v:textpath` with a `string`) shows as `[Unsupported pict]` in prose; its text is recorded in metadata, whitespace-normalised, in the body, headers/footers, notes, and text boxes, so editing, adding, or removing it makes companions stale; WordArt in tracked deletions or comments is not recorded, and only the chosen `mc:AlternateContent` copy is read, as for images; DrawingML WordArt is a text box and renders in prose |
| Drawings | Omission warnings and placeholders (e.g. `[Unsupported drawing]`) for common inline constructs, including a picture grouped with a text box; no rendering, OCR, object execution, or text extraction from charts, diagrams, or embedded objects |
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
destinations, field fingerprints and legacy form field state, list formats, image hashes, and equation, symbol, and object part fingerprints. `lists` has one entry per
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
`sha256`. An image with alt text also has `alt` (and `alt_title` for a title),
whitespace-normalised; the keys are present only when non-empty, so images
without alt text keep the metadata they had. Media file names are not recorded because Word renumbers them on save,
so order plus hash is the comparison, and re-encoded images are a real change.
The chosen `mc:AlternateContent` copy is read once: the text box branch, else
the equation branch, else the first `mc:Choice` (or `mc:Fallback`). `objects` lists equations, symbols, charts, diagrams,
embedded objects, imported chunks, and WordArt in one traversal order, each with its story
`part`: `{"kind": "equation", "sha256": …}` (the canonical accepted-view hash
above), `{"kind": "WordArt", "text": …}` (a VML `v:textpath` string), `{"kind": "symbol", "font": …, "char": …}` (`w:font` and `w:char`,
uppercased), or `{"kind": "chart" | "diagram" | "embedded object" | "imported
chunk", "sha256": …}` (the raw bytes of the part it points to; an embedded object
also records its `prog_id`, and a linked one `target` instead of `sha256`; a
dangling reference has a null `sha256`). Part names are not recorded, as for
images. The key is present only when the list is non-empty, so documents without
any of these keep the metadata they had. A field entry has a `form` key only for a
legacy checkbox or dropdown form field: `{"type": "checkbox", "name": …,
"checked": true | false}` or `{"type": "dropdown", "name": …, "entries": […],
"selected": … | null}`, so other fields keep the metadata they had. Field fingerprints avoid huge Zotero JSON
payloads but only say a normalized code changed, not what that change means.
Arrays follow document traversal order; related parts and warnings sort
lexically. Internal part paths come from the package; no machine paths, save
counters, revision dates, or export timestamps are added. Hyperlink destinations
are retained literally and may themselves contain sensitive or local paths.

## The review report

`pilcrow review` compares the same two views the companions hold, the prose and
the metadata above, and describes their differences; it adds no new extraction.
It claims no more than they do: formatting, layout, and unrendered content can
differ while the report says "No supported prose or review changes detected".
Metadata entries are matched by aligning their full content, so an entry that
moved is reported as removed and added. A field change inside an aligned block
(a link target, an image hash, a list level's format) is described at that
position. Comments are worded as resolved, reopened, re-anchored, or edited only
when every other recorded field, including author, text, anchor and the parent
of a reply, identifies exactly one comment on each side, and resolved or
reopened only when both states are known; otherwise they fall back to removed
and added, or to a notice that the category changed. A reply is matched by its
parent's author, text, and anchor rather than its recorded position, so inserting
an earlier comment does not change it; when several comments share that identity,
the position is compared too, so moving a reply between them is still reported.
Replies are then compared through their parents' pairing: once a parent is matched
or worded as an edit, a reply to it counts as having the same parent, so editing,
re-anchoring, or resolving a parent gives no line for an unchanged reply, and a
reply that changed too is worded as an edit (describing its new parent). A reply
whose parent is removed, added, or cannot be paired is still removed and added.
An image whose media part goes missing or appears is listed as removed and added.
An image's alt text change is worded with both texts, `~ image 1
(word/document.xml) alt text: "A red cat" → "A blue dog"`, with `(none)` for alt
text added or removed. WordArt is listed with charts, diagrams, and embedded
objects, and a text change reads `~ WordArt 1 (word/document.xml): "SALE" →
"CLOSED"`.
Equations and symbols are numbered by kind (equation 1, 2, … and symbol 1, 2, …);
an equation at the same position of an aligned block is worded as changed, and a
symbol as its old and new character and font. An equation's linear text is
prose, so an edit inside it also shows in the prose diff as `-`/`+` lines (with
`--word-diff`, the marks land on the changed math text); the linear text does not
show layout, fonts, or alignment, which only the fingerprint covers.
Charts, diagrams, embedded objects, and imported chunks have their own section,
numbered by kind the same way; one at the same position whose part's hash changed
is worded as changed (`~ chart 1 (word/document.xml) changed`, with an embedded
object's ProgID before the part); a part that goes missing or appears, or a link
whose target changes, is listed as removed and added. An equation or symbol that only
moved past a chart, diagram, embedded object, or imported chunk, while each
section's own entries stay the same, is also reported as `Objects: equations or
symbols moved relative to charts, diagrams or embedded objects`.
A legacy checkbox or dropdown form field at the same position whose state alone
changed is worded with its old and new state, `~ field 1 (FORMCHECKBOX "Check1")
in word/document.xml: unchecked → checked` or `~ field 2 (FORMDROPDOWN "Drop1")
in word/document.xml: "Yes" → "No"`, and a changed list of dropdown entries as
`entries changed`; a form field that was renamed or changed kind is listed as
removed and added, with its state.
Long texts are shortened to 80 characters for display.

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
  or brackets stay with the sentence. A small English abbreviation list (including
  `cf.` and `viz.`) and single-letter/initialism heuristic suppress common false
  splits. Decimals such as `3.14` have no whitespace and stay intact.
  Question/exclamation sentences and ordinary Unicode text are retained.
- There is no break when the next word starts with a lowercase letter (`Ph.D. in`,
  `Acme Inc. was`, `Max. load`, `Absolutely... fine`), which also covers unknown
  abbreviations. The trade-off: a sentence that really starts lowercase after
  `.`, `!`, or `?` (for example one starting with `iPhone`) joins the previous line.
- After `approx.`, `ca.`, `ch.`, `chap.`, `sec.`, `secs.`, `art.`, `arts.`,
  `para.`, `paras.`, `fn.`, `nos.`, or a month abbreviation (`Jan.` to `Dec.`,
  `Sept.`; any case) there is no break when the next word starts with a digit
  (`Art. 5`, `Sec. 3.2`, `Jan. 2020`). Before a capital letter these still end a
  sentence, so `It was Jan. Then it snowed.` stays two lines.
- Upgrading from a version without the two rules above: documents containing
  these forms report stale companions once, because their sentence lines change.
  Run `pilcrow export` and commit the refreshed companions. Documents without
  them produce byte-identical companions.
- This is English-oriented, not universal sentence detection. Unknown abbreviations
  can still split too soon before a capitalized word or a number; an abbreviation
  at sentence end can suppress a valid split.
  CJK text without whitespace stays on one line. Citation punctuation, initials,
  ellipses, and unusual scientific notation may need later tuning. An incorrect
  split changes line boundaries, not the underlying words.
- Heading text stays on one line. Each list item starts with `-`; subsequent
  sentences are indented. Table cells contain normal sentence/paragraph text.
  There are no per-sentence positional IDs or automatic counters.
- An equation's `[Math: …]` text is never split, even when it holds `. ` before a
  capital letter, `!`, or `?`; sentence breaks around it follow the rules above.
- The `.md` is optimized for the ordinary text diff, not faithful Markdown rendering:
  literal Markdown metacharacters in Word are not escaped, and table/text box/unsupported
  markers can coincide with literal document text.

## Equations as linear text

Each equation (`m:oMath`) is shown as `[Math: …]`: a deterministic linear text,
loosely after UnicodeMath, meant for reading a change, not for pasting into Word or
LaTeX. An `m:oMathPara` holding several equations shows each as its own `[Math: …]`,
separated by a space; an equation with no text shows `[Math: (empty)]`.

- It reads the accepted view, like the fingerprint: content in `w:del`/`w:moveFrom`
  is left out and `w:ins`/`w:moveTo` kept; formatting (`w:rPr`, `m:ctrlPr`, and the
  `m:*Pr` properties other than those named below) adds no text. Tracked changes
  inside an equation are not listed with the document's tracked changes, and comment
  markers inside an equation are not read, as before; a comment range around an
  equation anchors the same `[Math: …]` text.
- Text (`m:t`, and `w:t` in a `w:r` inside math) is shown as written; whitespace is
  collapsed. `w:tab` and `w:br` are a space, and `w:sym` is `[Unsupported sym]`.
- An **operand** (a script, a limit, a fraction side, a base) is wrapped in parentheses unless
  it is a single atom: one character, only letters, a number (`12`, `3.5`), or one
  delimiter with both its characters (`(a+b)` is not wrapped again). An empty
  argument shows nothing, and an empty script is left out with its `_` or `^`.

| OMML | Linear text |
| --- | --- |
| Fraction `m:f` | `num/den`: `a/b`, `(a+b)/(2c)`; `¦` instead of `/` for a `noBar` fraction |
| Scripts `m:sSup`, `m:sSub`, `m:sSubSup`, `m:sPre` | `x^2`, `x_i`, `x_i^2`, `_i^2 x`; `x^(n+1)` |
| Radical `m:rad` | `√(x)`; with a degree that `m:degHide` does not hide, `√(3&x)` |
| Delimiter `m:d` | `m:begChr`, the elements joined by `m:sepChr`, then `m:endChr` (defaults `(`, `\|`, `)`; an empty `m:val` means none): `(a\|b)`, `[0,1)` |
| N-ary `m:nary` | `m:chr` (default `∫`), then `_sub` and `^sup` unless `m:subHide`/`m:supHide` hide them, then a space and the operand: `∑_(i=1)^n x_i`; without limits, `∑x` |
| Function `m:func` | the name, a space, and the operand, never wrapped: `sin x`, `sin 2x`, `lim_(n→∞) a_n` |
| Limits `m:limLow`, `m:limUpp` | `lim_(n→∞)`, `x^k` |
| Accent `m:acc` | the operand, then `m:chr` (default U+0302, combining circumflex): `x̂` |
| Bar `m:bar` | the operand, then U+0305 (combining overline) for `m:pos` `top`, else U+0332 (combining low line) |
| Group character `m:groupChr` | the operand, then `m:chr` (default `⏟`) |
| `m:box`, `m:borderBox`, `m:phant` | their content (a phantom's content shows even when Word hides it) |
| Matrix `m:m` | `[a, b; c, d]`: cells joined by `, `, rows by `; ` |
| Equation array `m:eqArr` | rows joined by `; ` |

Any other math element shows its content's text, and each such element name adds
one coverage warning, for example `Unsupported math construct foo; rendered as its
text, structure not shown`. Layout, fonts, alignment, equation numbering, and
argument sizes are not shown. The equation fingerprint in `objects` stays the
safety net: an OMML change the linear text misses still makes companions stale.

## Operational limits

Packages are limited to 4096 entries and 1 GiB total uncompressed (Word's own
documented ceiling for a document with graphics is 512 MB), which bounds
decompression work. Parts that Pilcrow parses into memory (XML and `.rels`) are
limited to 16 MiB each and 64 MiB in total, so a `word/document.xml` over 16 MiB
is rejected. They are inflated no further than the size the package declares
for them, so a part that inflates past it fails its CRC check without using more
memory. Images are hashed in 1 MiB chunks without holding them in memory, so
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
