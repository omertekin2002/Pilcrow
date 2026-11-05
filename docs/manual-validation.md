# Validate with Word and GitHub Desktop

Automated tests use generated ZIP/XML fixtures and a temporary Git repository.
They validate extraction and ordinary Git diffs, not Microsoft Word rendering or
GitHub Desktop itself. Neither Word nor Desktop was available for this validation;
no real document corpus was provided. Run this checklist on a copy of a document
you are allowed to share. Do not use the only copy of important work.

1. Put the copy in a disposable Git repository opened in Desktop. Install Pilcrow
   in your Python environment, then export with `pilcrow export --all --root .`.
   Read stderr and `.meta.json` warnings. Confirm headings, paragraphs, list items,
   and table cells are understandable in `.prose.md`.
2. Commit the DOCX and companions in Desktop. Export again: the companions should
   remain unchanged. Run `pilcrow check --all --root .`: expect exit 0.
3. Open/save/close the file in Word **without editing**. Export again. The DOCX may
   change as a package, but ordinary prose and review metadata should stay stable.
   Investigate any companion churn; keep a sanitized reproducible case if possible.
4. Change one sentence in a multi-sentence paragraph, save, and run `check`: expect
   exit 1. Export, select the companion in Desktop Changes, and verify that only
   the affected sentence changes. Confirm the original DOCX still opens normally.
5. Add a heading, a list item, a two-cell table row, a note, and a hyperlink. Export
   and review boundaries. Change only the hyperlink URL: inspect the metadata diff.
6. Insert/delete text with Track Changes on, add a comment, then export. Prose
   should include inserted text and omit deleted text. Review text/authors should
   appear in metadata. Inspect complex tracked structures in Word directly.
7. Make a bold-only change and change an image. Export. No prose update is required;
   coverage warnings should still make the limitation clear. Inspect in Word.
8. If using Zotero/fields, update citations/bibliography in Word before saving.
   Confirm saved display text appears in prose and code changes can affect the
   field fingerprints. Pilcrow does not refresh or verify the bibliography.
9. Commit the edited DOCX and companions together. Select the companion in Desktop
   History and inspect the committed difference. Confirm branch/index/history
   remain under your control.
10. In the disposable repository only, try two branches with changes to the same
    document after committing `pilcrow setup` rules. Merge through your usual
    workflow. Choose/reconcile the DOCX, regenerate companions, check, and complete
    the merge. Never resolve Word content by editing only the companion.

Record OS, Word version, Desktop version, Python version, and sanitized symptoms
for failures. These manual results are needed before extending compatibility claims.
