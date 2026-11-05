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
| Headings | Outline level, inherited paragraph styles, and conventional Heading1–9 style IDs; capped at six Markdown levels |
| Lists | Item and indentation boundaries from direct/inherited `numPr`; generic `-` marker; automatic numbers/bullet glyphs are not reconstructed and trigger a warning |
| Tables | Explicit table, row, and cell markers, including empty cells; nested tables and content-control wrappers; grid spans/vertical merge markers retained, no rendered grid/layout |
| Footnotes/endnotes | Related parts with reference markers; separator notes omitted; IDs retained, so renumbering can produce diffs |
| Headers/footers | Related parts shown as separate sections by part name; no page/section placement mapping |
| Tracked text | Include `ins` and `moveTo`, omit `del` and `moveFrom`; separate kind, author, and text records, without dates or revision IDs |
| Comments | Text and author in metadata; anchors, dates, threaded replies, rich formatting, and resolved state not compared |
| Hyperlinks | Accepted visible text in prose; label and destination in metadata, without fetching the URL |
| Fields/citations | Saved display text in prose; normalized instruction fingerprint and first command token in metadata; never recalculate fields or manage Zotero |
| Text boxes, drawings, images, embedded objects, equations | Omission warnings and placeholders for common inline constructs; no rendering, OCR, object execution, or asset-content comparison |
| Bold/italic, fonts, colors, other styles, pagination/layout | Not compared; always disclosed in coverage warning |

Tracked text inclusion is an **accept-insertions/omit-deletions view**, not Word's
complete Accept All Changes implementation. Paragraph-mark deletion, list/table
revisions, tracked formatting, moves expressed in unusual structures, and nesting
can differ from Word's result. These constructs are reported where detected;
review them in Word. Deleted table rows are omitted. Comment text is plain text;
hyperlinks/fields inside comments are not separately interpreted.

Hidden text is included and triggers a visibility warning; formatting is not used
to decide what Word would display. Unknown Word/XML constructs trigger conservative warnings. Extra package parts
are listed as unrendered; their bytes are not hashed. Some save-only parts
(`docProps`, settings, font tables, theme, numbering definitions) are deliberately
ignored. List values and layout can therefore change without a companion update.
Unsupported assets can change while warning text remains identical. The `partial-prose`
coverage marker and CLI wording apply even when there are no specific warnings.
No claim of a complete/unchanged document is made. This is a modest coverage report,
not an audit of every OOXML relationship or layout semantic.

Metadata contains only schema/view/coverage, warnings, review text/authors, link
destinations, and field fingerprints. Field fingerprints avoid huge Zotero JSON
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
  literal Markdown metacharacters in Word are not escaped, and table/unsupported
  markers can coincide with literal document text.

## Operational limits

Packages are limited to 4096 entries, 64 MiB total uncompressed, and 16 MiB per
part to bound memory. Large image-heavy documents may exceed these deliberately
conservative defaults. Duplicate ZIP entries and XML entity expansion are rejected.
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
3. Decide whether list values, comment anchors, field-code previews, or image hashes
   justify more metadata. Avoid a dump of all OOXML.
4. Improve tracked structural revisions, note-reference validation, complex tables,
   and relationship/path canonicalization only as actual documents require it.
5. Consider opt-in textconv and refresh-only watching after the primary workflow is
   reliable. Desktop executable discovery needs its own validation.
6. If Word XML is wanted later, build a distinct dialect-detecting adapter. Do not
   treat arbitrary XML as Word input. Automatic document merging stays separate.
