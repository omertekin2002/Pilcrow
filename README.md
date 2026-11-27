# Pilcrow

Pilcrow makes Word text changes reviewable in GitHub Desktop. Keep editing the
`.docx`; export readable companion files and commit both through your normal Git
workflow. **The DOCX is authoritative.** The prose is a generated review artifact,
not an editable replacement or a round-trip format.

This is an independent, small Python implementation, conceptually inspired by the
companion-file approach of stet. No source or tests were copied from that project.

## Install

With Node.js 22 or newer:

```sh
npx otekin pilcrow
pilcrow --help
```

The installer installs `pilcrow-docx` from PyPI for your user account, using uv
and downloading uv and Python if needed. It supports Linux, macOS, and Windows.
If it prints a directory to add to PATH, add that directory and reopen your
terminal. Run `npx otekin pilcrow --upgrade` to update Pilcrow.

If you already use uv, you can also install directly:

```sh
uv tool install pilcrow-docx
```

### Install from source

Requires Python 3.10 or later. Git is needed for repository discovery and the Git
workflow. Word is not required to export. From this project's directory:

```sh
python -m pip install .
pilcrow --help
```

Use `python3` instead of `python` on systems where needed. A virtual environment
is recommended (`python -m venv .venv`); activate it before installing:

- macOS/Linux: `source .venv/bin/activate`
- Windows PowerShell: `.venv\Scripts\Activate.ps1`
- Windows Command Prompt: `.venv\Scripts\activate.bat`

If activation is unavailable, invoke `.venv/bin/python` or
`.venv\Scripts\python.exe` directly with `-m pip install .`, then `-m pilcrow`.
For development, see [Development](#development).

Pilcrow uses Python's ZIP reader and `defusedxml` for XML parsing. There is no Word
process, cloud service, AI connection, watcher, or background service. Extraction
never writes to the DOCX.

## Quickstart

From a Word file already in Git to a readable review, then to a freshness gate.
Run these in your document repository; `report.docx` stands for your file.

**1. Review a committed change, with no setup.** Compare the last commit with
your working copy:

```sh
pilcrow review --rev HEAD "report.docx"
```

Without Git history, compare two files: `pilcrow review old.docx new.docx`.
Review is read-only and writes nothing, not even `.pilcrow/` companions. It exits
0 when no supported changes are found, 1 when it reports changes, and 2 on an
error. `--rev` first ships in the release after v0.1.7. More:
[Review a change in words](#review-a-change-in-words).

**2. Export a baseline**, then check it:

```sh
pilcrow export "report.docx" --root .
pilcrow check "report.docx" --root .
```

Confirm that `report.docx` is named in both outputs: a passing check that selected
no documents proves nothing. Commit the DOCX and its two companions under
`.pilcrow/` together. More: [Baseline, then review a change](#baseline-then-review-a-change).

**3. Review the next edit.** Save your edit in Word, export again, and read the
`.prose.md` diff in GitHub Desktop's **Changes** view or with Git (`--word-diff`
marks the changed words in each line):

```sh
pilcrow export "report.docx" --root .
git diff --word-diff -- ".pilcrow/report.docx.prose.md"
```

The Word and GitHub Desktop steps have not been tested; see
[manual validation](docs/manual-validation.md).

**4. Add a freshness gate.** A pre-commit hook or a GitHub Action runs
`pilcrow check --all` and fails on missing or stale companions. In
`.pre-commit-config.yaml`:

```yaml
repos:
  - repo: https://github.com/omertekin2002/Pilcrow
    rev: v0.1.8
    hooks:
      - id: pilcrow-check
```

Or as a step in a GitHub Actions job, after `actions/checkout`:

```yaml
      - uses: omertekin2002/Pilcrow@v0.1.8
```

More, including the optional review report in the job summary:
[Enforce freshness](#enforce-freshness).

## Baseline, then review a change

In **your document repository**, choose an explicit scope. For example:

```sh
pilcrow export --all --root .
pilcrow check --all --root .
```

Or select documents (paths are relative to your current directory):

```sh
pilcrow export "report.docx" "chapters/Methods and results.docx" --root .
```

Outputs retain the complete source-relative filename:

```text
report.docx
chapters/Methods and results.docx
.pilcrow/report.docx.prose.md
.pilcrow/report.docx.meta.json
.pilcrow/chapters/Methods and results.docx.prose.md
.pilcrow/chapters/Methods and results.docx.meta.json
```

1. Export a baseline. In GitHub Desktop, review and commit the DOCX and its two
   companion files together. Do not put `.pilcrow/` in `.gitignore`.
2. Edit and save the DOCX in Word, or use a separate document-editing tool.
3. Run the same `pilcrow export` command again.
4. In Desktop's **Changes** view, select the `.prose.md` file to read the text diff;
   inspect `.meta.json` for review/coverage changes. The DOCX may still say binary
   changed. Review formatting/images in Word when relevant.
5. Commit the DOCX and both companions together through Desktop, then push and
   branch normally. Desktop's **History** view can show committed companion diffs.

Pilcrow never stages, commits, pushes, switches branches, or merges. No DOCX
text converter is installed unless you opt in (see
[readable `git diff`](#readable-git-diff-optional)). The CLI and an ordinary Git diff have been
tested; **the Microsoft Word / GitHub Desktop GUI workflow has not been tested in
this Linux environment**. See [manual validation](docs/manual-validation.md).
See an [actual generated prose and Git diff example](docs/example.md).
Old commits without companions will not gain prose diffs automatically; there is
no backfill or history rewriting. The optional `git diff` driver below can show
their prose on the command line.

## Scope and discovery

- `--root` defaults to the **current directory**, never an automatically detected
  repository top. It controls both discovery and the location of `.pilcrow/`.
  Use the same root each time. From a subdirectory, pass `--root ..` deliberately
  if the output belongs at the parent. Explicit paths remain relative to the
  current directory, not to `--root`.
- Supply paths **or** `--all`. No arguments does not silently scan the machine.
  Passing a directory as a document is an error; use `--all --root directory`.
- Inside Git, `--all` uses `git ls-files` for tracked and nonignored untracked
  files within the selected root. Tracked files remain included even if an ignore
  rule matches. An explicitly selected ignored DOCX is allowed.
- Outside Git, `--all` recursively walks the selected directory. Git ignore rules
  only apply inside a Git worktree; there is no separate ignore syntax.
- Both modes exclude Word lock files (`~$…`), `.pilcrow`, `.git`, `.venv`, `venv`,
  `env`, `node_modules`, `__pycache__`, `.tox`, `.nox`, `build`, and `dist`.
  Extension matching is case-insensitive. Unicode and spaces are supported.
- Symlink files/paths are rejected; directory walking does not follow symlink
  directories. Submodule documents are not recursively discovered from the parent
  repository: select a root inside the submodule.
- `--all` reports orphan companions after source deletion/rename. Remove the old
  companion files manually if the deletion was intentional and commit those
  deletions together. Export never automatically deletes review history files.
  Explicit-file mode checks only selected inputs, not unrelated orphans.

## Check before committing

```sh
pilcrow check --all --root .
```

`check` is read-only, even when files are missing. It recomputes expected prose and
metadata and compares bytes, tolerating ZIP order/compression, run boundaries,
save timestamps, and other ignored Word metadata. It does not use a DOCX-byte hash.

| Exit | Meaning |
| --- | --- |
| 0 | Successful export, or companions match the supported normalized view |
| 1 | Missing/stale companions in `check`, or orphan companions in either command |
| 2 | Invalid selection, unreadable/malformed input, unsafe path, or write failure |

Errors identify the file; remaining selected documents are processed after a
per-document error. Retry after Word has finished saving if a file is temporarily
unreadable or changes during extraction. A fresh result means **no supported
prose/review changes detected**, never that the complete document is identical.
Warnings appear on stderr and in metadata even for fresh results. Known coverage
limitations do not themselves cause a nonzero exit.

Writes use a temporary file and atomic replacement in the destination directory.
Byte-identical outputs keep their existing modification times. Each file is
atomic, **not the whole pair or batch**: if interrupted between files, run `check`
and export again. Avoid concurrent exports into the same root.

## Enforce freshness

Pilcrow ships a [pre-commit](https://pre-commit.com) hook and a GitHub Action, so a
document repository can refuse stale companions locally and in CI. Both only run
`pilcrow check --all`: they never export, stage, or rewrite files. When either
fails, run `pilcrow export` yourself and commit the DOCX and its companions
together.

Pin both to a Pilcrow release tag (see the repository's Releases). The hook and
action first ship in the release after v0.1.2. The snippets in this README assume
`v0.1.8`, the release after v0.1.7, which adds `pilcrow review --rev` and the
action's `summary` input. Pin the same release as your installed
`pilcrow --version`: a newer release can record more in its companions, so
companions exported by one release can be stale to another.

`.pre-commit-config.yaml` in your document repository root (then run
`pre-commit install` once):

```yaml
repos:
  - repo: https://github.com/omertekin2002/Pilcrow
    rev: v0.1.8
    hooks:
      - id: pilcrow-check
```

The hook runs `pilcrow check --all --root .` from the repository root when a commit
adds or changes a DOCX or a file under `.pilcrow/`. pre-commit does not pass
deleted paths to hooks, so a commit that only deletes files (for example a DOCX
without its companions) skips the hook; the CI check below catches the resulting
orphans. To run the check on every commit instead, add `always_run: true` under
`- id: pilcrow-check`. The hook is for the pre-commit framework, run by
`git commit` once `pre-commit install` has set it up. It has not been tested with
GitHub Desktop.

`.github/workflows/pilcrow.yml` in your document repository:

```yaml
name: Pilcrow
on: [push, pull_request]
jobs:
  check:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v7
      - uses: omertekin2002/Pilcrow@v0.1.8
        with:
          root: .              # optional, default "."
          python-version: "3.12"  # optional, default "3.12"
```

The action installs Pilcrow from the tagged source into a private virtual
environment and fails with Pilcrow's exit code: 1 for missing, stale or orphan
companions, 2 for errors.

### Review report in the job summary (optional)

The action can also write a review of each DOCX the push or pull request changed
to the workflow run's job summary, so a reviewer can see what changed without
checking the branch out. It is off by default (`summary: none`): without the
input, the action only runs the check, exactly as above. The `summary` and `base`
inputs first ship in the release after v0.1.7; the snippet below assumes `v0.1.8`.

```yaml
    steps:
      - uses: actions/checkout@v7
        with:
          fetch-depth: 0       # the summary compares from the base's history, so fetch it
      - uses: omertekin2002/Pilcrow@v0.1.8
        with:
          summary: outline     # none (default), outline or full
          # base: <revision>   # optional; default: the pull request's base, or the push's previous commit
```

- `outline` publishes **no document text**. For each changed document it lists
  the path, its status (changed, no supported changes, added, removed, or could
  not be read; why a document could not be read goes only to a warning
  annotation) and the report sections that differ, each with its count of
  described lines, such as `` `report.docx` — changed: Prose changed, Comments (2) ``.
- `full` publishes each changed document's complete
  [`pilcrow review --format markdown`](#review-a-change-in-words) report, comparing
  the document at the merge base of `base` and `HEAD` with `HEAD`. **The report
  includes document text (prose, comments, tracked changes, link targets), and
  anyone who can view the workflow run can read it**, which for a public repository means anyone. The summary
  repeats this warning at the top. Choose `outline` when the text must not leave
  the repository.

The new side is always `HEAD`, and the old side is the merge base of `base` and
`HEAD`, the commit `HEAD` forked from, as in GitHub's Files changed view. So the
summary shows what the pull request changes whether the workflow checks out the
default merge commit or the pull request's head (`ref: ${{ github.event.pull_request.head.sha }}`),
and edits made on the base branch since the fork, or commits a force-push
replaced, are not reported as the pull request's own. When `base` is an ancestor
of `HEAD`, as for the merge commit or an ordinary push, the merge base is `base`
itself; otherwise the line under the summary's title names the merge base. Only
DOCX files within `root` are listed, with the same lock-file and excluded-folder
rules as `pilcrow check --all`. A renamed document is listed as removed and
added, and an added or removed one has no report.

The summary step is read-only and never fails the job (only a `summary` value
other than `none`, `outline` or `full` does): it runs even when the
check fails, and the action's outcome stays the check's. When the base is missing
(the first push of a new branch), not in the checkout (the default shallow
`fetch-depth: 1` clone), or shares no history with `HEAD` (unrelated histories,
or a shallow clone that stops before the merge base), the summary explains why
there is no review and the step adds a warning annotation. One unreadable document gets a warning while the
others are still reported. GitHub limits a job summary to 1 MiB, so once the
reports reach about 900 KB, the remaining documents are listed as `report
omitted: summary size limit`. The action's `summary-file` output is the path of
the Markdown file it wrote, for later steps that want to post or check it.

## Branches and conflicts

Opt in to conservative merge rules once, from your document repository root:

```sh
pilcrow setup --root .
```

This appends a managed block to `.gitattributes`, preserving other rules. Repeating
it is harmless. Review and commit the file through Desktop. Rules mark DOCX and
all `.pilcrow/` files `-merge` (Git's binary merge driver), and companion files as
`text eol=lf`. Both-sided modifications conflict instead of automatically merging
prose into a document that never existed. Normal one-sided changes still merge.
More-specific nested attributes can override these rules; inspect with
`git check-attr merge -- "report.docx" ".pilcrow/report.docx.prose.md"` if needed.

When branches conflict:

1. Preserve copies of both versions of the DOCX outside the conflicted path.
2. Choose the authoritative version or reconcile the documents manually in Word.
   If using Git commands to select a side, make sure you understand that side in
   the current merge/rebase. Pilcrow does not make this choice.
3. Put the resolved DOCX at its repository path. Regenerate its companions, even
   if Git did not flag them as conflicted. Never independently accept a merged
   prose artifact as the Word document.
4. Run `pilcrow check --all --root .`. Inspect the resulting text and Word file,
   then mark resolutions and commit using your normal Git workflow.

Git commit identity attributes a snapshot, not every contributor between saves.
For basic snapshot attribution:

```sh
git blame -- ".pilcrow/report.docx.prose.md"
```

Word tracked-change authors (included in review metadata) are separate from Git
authors. Neither proves an edit came from an AI. Companion metadata can contain
comment authors and deleted text: review it before sharing, just as you would the
DOCX. No automatic OOXML merging or sentence-lineage system is provided.

## Readable `git diff` (optional)

Git can also print a DOCX's prose itself in `git diff`, `git log -p` and
`git show`, instead of "Binary files differ". Opt in with two steps from your
document repository root:

```sh
pilcrow setup --textconv --root .
git config diff.pilcrow.textconv "pilcrow textconv --"
```

The first adds `*.[dD][oO][cC][xX] diff=pilcrow` to Pilcrow's managed block in
`.gitattributes` (commit it like the other rules); a later plain `pilcrow setup`
keeps that line. The second is a one-time, per-clone Git setting that Pilcrow
prints but never runs: Pilcrow does not edit Git config. Each collaborator who
wants the readable diff runs it in their own clone. Optionally add
`git config diff.pilcrow.cachetextconv true` so Git caches the converted text.
Until the config is set, Git treats DOCX files exactly as before. The `--`
lets file names that start with `-` convert too.

`pilcrow textconv <file>` prints exactly what the file's `.prose.md` companion
would contain and writes nothing. Coverage warnings go to stderr, so they can
appear in the terminal alongside the diff. Because Git converts each revision on
demand, this also works for history committed before you used Pilcrow, with no
companions. It shows prose only: changes recorded in `.meta.json`, such as
comments, list formats, images and their alt text, WordArt text, form
checkbox and dropdown state, and tracked-change authors, do not appear, so it does not replace companions,
`pilcrow check`, or review in Word. Whether
GitHub Desktop uses the driver is **untested**.

To opt out, delete the `diff=pilcrow` line from the managed block by hand, or
run `git config --unset diff.pilcrow.textconv`.

## Review a change in words

`pilcrow review OLD.docx NEW.docx` prints a plain-language report of what
changed between two DOCX files. It is read-only: it extracts both files in
memory and writes nothing (no `.pilcrow/` companions).

To review a committed change, name the revision with `--rev`, as in `git diff`:

```sh
pilcrow review --rev HEAD report.docx              # last commit → working copy
pilcrow review --rev HEAD~1 --rev HEAD report.docx # one commit → the next
pilcrow review --rev main --rev feature report.docx
```

With one `--rev`, the old side is `report.docx` as committed at that revision
and the new side is the file in your working copy. With two, both sides come
from Git and the working copy is not read. The header names both sides, such as
`Review: HEAD~1:report.docx → report.docx`. The path is relative to the current
directory, like `git show HEAD:./report.docx`. A folder inside the repository
that was committed as a folder and is now a symlink keeps its own history:
Pilcrow reads the committed file, not whatever the link points to. The old
side is extracted from the **committed DOCX itself**, not from companions
committed with it, so it needs no `.pilcrow/` files and works for history from
before you used Pilcrow.
Pilcrow reads it with read-only Git commands (`rev-parse` and `cat-file`): it
never stages, checks out, or writes a file, not even a temporary one, and it
checks the committed file's size against the package limit before reading it.

Alternatively, let `git difftool` hand it the two versions as temporary files:

```sh
git difftool -y -x 'pilcrow review' HEAD~1 -- report.docx
```

The report starts with a unified diff of the prose (the `.prose.md` text, capped
at 2,000 lines), then one section per kind of review metadata that differs:
comments, tracked changes, hyperlinks, fields, images, equations and symbols,
charts, diagrams and embedded objects, lists, and coverage warnings. Entries are matched by content, not by position, and described in
words, for example `~ resolved comment by Ada on "first sentence.": "Please
clarify."`, `~ link "Report": https://old.example → https://new.example`, or
`~ list 1 level 0: bullet → decimal`. An edited equation reads `~ equation 1
(word/document.xml) changed`, next to its `-`/`+` `[Math: …]` lines in the prose
diff, and a changed symbol `~ symbol 1
(word/document.xml): F061 (Symbol) → F062 (Symbol)`; an edited chart, SmartArt
diagram, embedded object, or imported chunk reads `~ chart 1 (word/document.xml)
changed` or `~ embedded object 1 (Excel.Sheet.12, word/document.xml) changed`,
and an edited diagram node also shows as its `-`/`+` lines in the `[Diagram]`
block of the prose diff.
Edited image alt text reads `~ image 1 (word/document.xml) alt text: "A red cat"
→ "A blue dog"`, and edited WordArt text `~ WordArt 1 (word/document.xml):
"SALE" → "CLOSED"`.
Ticking a legacy form checkbox reads `~ field 1 (FORMCHECKBOX "Check1")
in word/document.xml: unchecked → checked`, and choosing another dropdown entry
`~ field 2 (FORMDROPDOWN "Drop1") in word/document.xml: "Yes" → "No"`. A comment is called resolved, reopened,
re-anchored, or edited only when its author, text, and anchor identify exactly
one comment on each side; otherwise the report says it was removed and added.
When something changed, the coverage warnings of both files are listed once.

To paste the report into a pull request, an issue, or a chat that renders
Markdown, add `--format markdown` (the default, `--format text`, is the plain
report above). It works with both forms, `OLD NEW` and `--rev`, and says the
same things with the same exit code:

```sh
pilcrow review --format markdown --rev HEAD~1 --rev HEAD report.docx
```

For a reworded sentence and a resolved comment, it prints:

````markdown
## Review: `HEAD~1:report.docx` → `HEAD:report.docx`

### Prose

```diff
--- OLD
+++ NEW
@@ -1,4 +1,4 @@
 first sentence.
 
 Opening sentence.
-The cat sat on the mat.
+The dog sat on the mat.
```

### Comments

- `~` resolved comment by Ada on "first sentence."\: "Please clarify."

### Coverage warnings (either side)

- Comment text, authors, anchored text, replies, and resolved state are recorded; dates and rich formatting are not compared.
- Formatting and page layout are not compared; unchanged prose is not an identical document.

Only the prose and review metadata Pilcrow extracts are compared (not a complete document comparison).
````

The header becomes a heading with both sides in code spans, and the prose diff
sits unchanged in a fenced `diff` block whose fence is longer than any run of
backticks in the document. Each metadata section becomes a subheading with one
bullet per line of the text report, its `-`, `~`, or `+` marker kept in a code
span so the meaning does not depend on colour. Text from the document (comment
text, anchors, authors, link text and targets, warnings) is escaped, so it renders
literally: it cannot add headings, emphasis, links, images, HTML, or tables. A
renderer may still turn a bare URL, an `@name`, or a `#123` into a link
afterwards, as GitHub does in any comment.

To see which words changed inside a sentence, add `--word-diff`. It works with
both forms and both formats, and changes only how the prose diff is printed:

```sh
pilcrow review --word-diff --rev HEAD report.docx
```

A replaced sentence that pairs up with its new version is printed once, after
`~ `, with removed words in `[-…-]` and added words in `{+…+}`:

```text
@@ -1,2 +1,2 @@
~ The supplier [-shall-]{+may+} deliver the goods within [-30-]{+60+} days.
 Payment is due on receipt.
```

Sentences pair one-to-one, in order, only inside one block of replaced lines
with as many new lines as old ones, and only when, for every pair, at least half
of the two sentences' words and punctuation match. Otherwise, and whenever a
sentence has more than 200 words or already contains `[-`, `-]`, `{+` or `+}`,
the whole old and new lines are shown as `-` and `+`, exactly as without the
flag. Dropping the added words and the brackets gives the old sentence back
exactly; dropping the removed words and the braces gives the new one. Without
`--word-diff` the report is unchanged, and the flag changes no exit code,
companion, `textconv` output, or Action job summary.

The report shows only what Pilcrow's extraction supports: formatting, layout,
and anything else outside the prose and `.meta.json` are not compared, so "No
supported prose or review changes detected" is not a claim that the documents
are identical. Like `textconv`, it accepts any file name, because Git passes
temporary copies. Exit codes follow `diff`: 0 when nothing supported differs,
1 when the report lists differences, and 2 when an input is missing, unreadable,
or malformed (the error names that input, such as `error: HEAD~1:report.docx: …`,
and nothing is printed on stdout). With `--rev`, that also covers running
outside a Git repository, an unknown revision, and a revision whose name starts
with `-`. A file added or deleted in the commit has no other side to compare, so
it reports an error. `git difftool` ignores the exit code unless you pass
`--trust-exit-code`.

In CI, the GitHub Action can publish these reports for every changed document in
the job summary; see [Review report in the job summary](#review-report-in-the-job-summary-optional).

## Coverage and maintenance

See [coverage, normalization, and backlog](docs/limitations.md). The prose retains
common headings, list indentation, tables, notes, headers, and footers. Tracked
insertions/move destinations are included and deletions/move sources omitted.
List bullet/number formats are recorded in metadata, so changing them makes
companions stale; the counters themselves are not shown in prose, and numbered
headings are not compared.
Complex tracked structure/layout is limited and reported. Text box text is shown
between `[Text box]` and `[/Text box]` markers after its anchoring paragraph; the
box's position, wrapping, and shape formatting are not compared. Images, charts,
and embedded objects are not rendered; images are recorded in
metadata by content hash, so swapping, adding, removing, or reordering one makes
companions stale (size, position, and cropping are not compared). An image's alt
text (its description, and its title when it has one) is recorded with it, so
editing, adding, or removing alt text makes companions stale too; screen readers
read it, but it is not shown in the prose. Legacy VML WordArt is not rendered
either; its text is recorded in metadata, so editing it makes companions stale.
Equations are shown in prose as linear text, for example `[Math: E=mc^2]` or
`[Math: ∑_(i=1)^n x_i]` (the notation is in
[Equations as linear text](docs/limitations.md#equations-as-linear-text)), so an edit
inside an equation shows in the prose diff; layout, fonts, and alignment are not
shown. Equations and Word symbols (`w:sym`) are also recorded in metadata by a
fingerprint of their OMML structure and text (formatting and revision IDs left out)
and by font and character code, so editing, adding, or removing one makes
companions stale even where the linear text does not show the change. Symbols stay
`[Unsupported sym]` in prose.
SmartArt diagram text is shown as an outline after its anchoring paragraph, so an
edit to a node shows in the prose diff; layout, shapes, and formatting are not
shown:

```text
[Diagram]

- Plan
  - Gather requirements
  - Draft the outline
- Review

[/Diagram]
```

Charts, SmartArt diagrams, embedded (OLE) objects, and imported chunks
(`w:altChunk`) are recorded by the SHA-256 of the package part they point to (the
chart part, the diagram's data part, the embedding, the chunk), so editing a
chart's values, a diagram's text, an embedded workbook, or a chunk, or adding or
removing one, makes companions stale; renaming the part does not. The fingerprint
covers the stored part, not how Word draws it, so a diagram edit that leaves its
outline text unchanged (formatting stored in the data part, for example) still
makes companions stale; a chart's own workbook and a diagram's layout, style, and
colour parts are not compared. Charts and embedded objects are not rendered.
Legacy form fields (`FORMCHECKBOX`, `FORMDROPDOWN`), common in contract, HR, and
government templates, keep their state outside the prose: Word writes no display
text for them. Each checkbox's checked state and each dropdown's entries and
selected entry are recorded with the field in metadata, so ticking a box or
choosing another entry makes companions stale; the state is not shown in the
prose. A text form field's value is its display text, already in the prose.
Content-control checkboxes need nothing extra, because their glyph is prose text.
Comments stay out of the prose; their text, author, and anchored text (the
accepted text their range covers, and the part it is in) are recorded in metadata,
so editing a comment or moving, widening, or narrowing its range makes companions
stale. From Word's comments-extended part, each comment also records whether it is
`resolved` and, for a reply, `reply_to` (the position of its parent comment), so
resolving, reopening, or replying to a comment makes companions stale too; both
are `null` when the document does not say. Dates, rich formatting, and the
`commentsIds.xml` and `commentsExtensible.xml` parts are not compared. Formatting-only edits can
leave both companions unchanged and still matter in Word.

Upgrading from a version that left text box text out: documents with text boxes
report stale companions once, because their prose now includes that text. Run
`pilcrow export` and commit the refreshed companions.

Upgrading from a version that did not record list formats: every document reports
stale companions once, because `.meta.json` now always has `lists` and
`list_runs` keys (empty for documents without lists). Run `pilcrow export` and
commit the refreshed companions. A document whose numbering relationship points to
a missing part is now an error, like a missing styles part.

Upgrading from a version that did not record image hashes: every document reports
stale companions once, because `.meta.json` now always has an `images` key
(empty for documents without images). This ships with the list-format change
above, so one `pilcrow export` and commit covers both.

Upgrading from a version that did not record comment anchors: documents with
comments report stale companions once, because each comment in `.meta.json` now
has `anchor` and `anchor_part` keys. Run `pilcrow export` and commit the refreshed
companions. Documents without comments are unaffected.

Upgrading from a version that did not record comment threads: documents with
comments report stale companions once, because each comment in `.meta.json` now
has `resolved` and `reply_to` keys. Run `pilcrow export` and commit the refreshed
companions. Documents without comments are unaffected. A document whose
comments-extended relationship points to a missing part is now an error, like a
missing styles part.

Upgrading from a version that split sentences after forms such as `Art. 5`,
`Jan. 2020`, `Cf. Smith`, or `Ph.D. in`: documents containing them report stale
companions once, because those sentences now stay on one line (see
[Normalization and sentence lines](docs/limitations.md#normalization-and-sentence-lines)).
Run `pilcrow export` and commit the refreshed companions. Other documents are
unaffected.

Upgrading from a version that did not record equations and symbols: documents
containing an equation or a Word symbol report stale companions once, because
`.meta.json` now has an `objects` key for them. Run `pilcrow export` and commit
the refreshed companions. Documents without equations or symbols are unaffected.

Upgrading from a version that did not record charts, SmartArt diagrams, embedded
objects, and imported chunks: documents containing one report stale companions
once, because `.meta.json` now records them in `objects` (and no longer warns
that their parts are not compared). Run `pilcrow export` and commit the refreshed
companions. Documents without them are unaffected.

Upgrading from a version that did not record legacy form field state: documents
containing a checkbox or dropdown form field report stale companions once,
because their field entries in `.meta.json` now have a `form` key (and documents
with any legacy form field no longer warn about its `w:ffData`). Run `pilcrow
export` and commit the refreshed companions. Documents without legacy form
fields are unaffected.

Upgrading from a version that did not record image alt text and WordArt text:
documents containing an image with alt text, or legacy VML WordArt, report stale
companions once, because their image entries in `.meta.json` now have an `alt`
(or `alt_title`) key, and WordArt is recorded in `objects`. Run `pilcrow export`
and commit the refreshed companions. Documents without them are unaffected.

Upgrading from a version that showed equations as `[Unsupported oMath]`:
documents containing an equation report stale companions once, because their
prose now shows each equation as `[Math: …]` text and their coverage warnings
changed. Run `pilcrow export` and commit the refreshed companions. Documents
without equations are unaffected.

Upgrading from a version that showed SmartArt diagrams as `[Unsupported drawing]`:
documents containing a SmartArt diagram report stale companions once, because
their prose now shows the diagram's text in a `[Diagram]` block and their coverage
warnings changed. Run `pilcrow export` and commit the refreshed companions. The
diagram's `objects` entry is unchanged, and documents without SmartArt are
unaffected.

Run the generated-fixture suite:

```sh
python -m unittest discover -s tests -v
```

The implementation has five small modules: `prose.py` (segmentation), `extract.py`
(OOXML coverage), `files.py` (scope and writes), `review.py` (the change report), and
`cli.py` (commands). Tests
include temporary Git repositories and failure cases. Generated fixtures do not
establish complete Microsoft Word compatibility. No real/private document corpus
was available for this implementation.

## Development

Install Pilcrow in editable mode with the pinned lint and type-check tools, then
run both checks from the project root (they read their settings from
`pyproject.toml`):

```sh
python -m pip install -e '.[dev]'
ruff check .
mypy
```

`ruff check .` lints `src/`, `tests/` and `scripts/`; `mypy` type-checks
`src/pilcrow` in strict mode.

To look for slow paths, `python scripts/benchmark.py` generates a 500-page DOCX
(lists, tables, comments and tracked changes) and one very long paragraph in a
temporary directory, then prints how long extraction takes for each. It uses only the
standard library; `--help` lists the size options. It is a manual tool, not a CI gate.

## Releases

Every pull request runs `.github/workflows/ci.yml`: the test suite on Linux, macOS
and Windows with Python 3.10 and 3.14 (plus 3.11–3.13 on Linux), a lint job that
runs `ruff check .` and strict `mypy`, and a packaging check that builds the sdist
and wheel, runs `twine check`, and runs the installed wheel. An `enforce` job
builds a fixture document repository with `scripts/make_fixture_repo.py` and checks
that the action (`uses: ./`) and the pre-commit hook (`pre-commit try-repo`) pass
with fresh companions and fail after the DOCX changes without re-export. The release workflow
below still runs the suite again before publishing.

Every push to `main` runs `.github/workflows/publish.yml`. Release runs are queued
and each checks out the current `main`, including previous release version bumps.
The workflow increments the patch version in `pyproject.toml` and
`src/pilcrow/__init__.py`, builds and validates the distributions, runs the test
suite, and verifies the installed CLI version. It then commits the version bump, pushes
the commit and tag atomically, and creates a GitHub release with the wheel and
source archive. Release commits use `chore(release): ... [skip ci]` to avoid loops.
If `main` advances during a build, the push fails safely before publishing; the
queued run checks out the newer `main`.

A separate job uploads those same release files to PyPI using Trusted Publishing.
Configure the PyPI publisher for owner `omertekin2002`, repository `Pilcrow`,
workflow `publish.yml`, and GitHub environment `pypi`. No PyPI API token is needed.
If only the PyPI upload fails, fix its configuration and use **Re-run failed jobs**
on that workflow run to retry the existing release without bumping the version.
