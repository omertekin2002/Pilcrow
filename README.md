# Pilcrow

Pilcrow makes Word text changes reviewable in GitHub Desktop. Keep editing the
`.docx`; export readable companion files and commit both through your normal Git
workflow. **The DOCX is authoritative.** The prose is a generated review artifact,
not an editable replacement or a round-trip format.

This is an independent, small Python implementation, conceptually inspired by the
companion-file approach of stet. No source or tests were copied from that project.

## Install locally

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
For development, use `python -m pip install -e .`.

Pilcrow uses Python's ZIP reader and `defusedxml` for XML parsing. There is no Word
process, cloud service, AI connection, watcher, or background service. Extraction
never writes to the DOCX.

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

Pilcrow never stages, commits, pushes, switches branches, or merges. No direct
DOCX text converter is installed. The CLI and an ordinary Git diff have been
tested; **the Microsoft Word / GitHub Desktop GUI workflow has not been tested in
this Linux environment**. See [manual validation](docs/manual-validation.md).
See an [actual generated prose and Git diff example](docs/example.md).
Old commits without companions will not gain prose diffs automatically; there is
no backfill or history rewriting.

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

## Coverage and maintenance

See [coverage, normalization, and backlog](docs/limitations.md). The prose retains
common headings, list indentation, tables, notes, headers, and footers. Tracked
insertions/move destinations are included and deletions/move sources omitted.
Complex tracked structure/layout is limited and reported. Images, equations,
text boxes, and embedded objects are not rendered. Formatting-only edits can
leave both companions unchanged and still matter in Word.

Run the generated-fixture suite:

```sh
python -m unittest discover -s tests -v
```

The implementation has four small modules: `prose.py` (segmentation), `extract.py`
(OOXML coverage), `files.py` (scope and writes), and `cli.py` (commands). Tests
include temporary Git repositories and failure cases. Generated fixtures do not
establish complete Microsoft Word compatibility. No real/private document corpus
was available for this implementation.
