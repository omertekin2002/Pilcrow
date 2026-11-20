"""Describe prose and review-metadata changes between two extractions in words.

Read-only: it compares two in-memory extractions and never touches files or Git.
Entries are matched by content alignment, never by position alone, and a change is
worded as an edit to one entry only when the metadata identifies that entry.
"""

from __future__ import annotations

import difflib
import json
from collections import Counter
from collections.abc import Callable, Sequence
from typing import Any

from .extract import Extraction

MAX_DIFF_LINES = 2000
MAX_TEXT = 80
BODY = "word/document.xml"
CAVEAT = "No supported prose or review changes detected (not a complete document comparison)."
TRACKED = {"ins": "tracked insertion", "del": "tracked deletion",
           "moveFrom": "tracked move source", "moveTo": "tracked move destination"}

Entry = dict[str, Any]
# (old entry, new entry, old position, new position) -> a "~ ..." line, or None if not one change.
Change = Callable[[Entry, Entry, int, int], str | None]


def canonical(entry: Any) -> str:
    return json.dumps(entry, ensure_ascii=False, sort_keys=True)


def short(text: str | None) -> str:
    """Quote ``text`` on one line, cut to MAX_TEXT characters."""
    value = " ".join((text or "").split())
    if len(value) > MAX_TEXT:
        value = value[:MAX_TEXT - 1] + "…"
    return f'"{value}"'


def author(entry: Entry) -> str:
    return entry.get("author") or "(no author)"


def where(part: str | None) -> str:
    return "" if part in (None, BODY) else f" in {part}"


def prose_section(old: str, new: str) -> list[str]:
    if old == new:
        return ["Prose: no changes"]
    diff = list(difflib.unified_diff(old.splitlines(), new.splitlines(), "OLD", "NEW", n=3, lineterm=""))
    lines = ["Prose:", *diff[:MAX_DIFF_LINES]]
    if len(diff) > MAX_DIFF_LINES:
        lines.append(f"… {len(diff) - MAX_DIFF_LINES} more diff lines omitted")
    return lines


def aligned(old: Sequence[Entry], new: Sequence[Entry],
            key: Callable[[Entry], str] = canonical) -> list[tuple[list[int], list[int]]]:
    """The unmatched (old positions, new positions) of each differing block of an alignment."""
    matcher = difflib.SequenceMatcher(None, [key(e) for e in old], [key(e) for e in new], autojunk=False)
    return [(list(range(i1, i2)), list(range(j1, j2)))
            for tag, i1, i2, j1, j2 in matcher.get_opcodes() if tag != "equal"]


def entries_section(old: Sequence[Entry], new: Sequence[Entry], describe: Callable[[Entry, int], str],
                    change: Change | None = None) -> list[str]:
    """Removed/added lines, with ``change`` wording pairs in the same position of a replaced block."""
    lines = []
    for olds, news in aligned(old, new):
        changes: list[str] = []
        unpaired_old, unpaired_new = list(olds), list(news)
        if change is not None:
            for i, j in zip(olds, news, strict=False):  # positional only within one block the alignment replaced
                line = change(old[i], new[j], i, j)
                if line is not None:
                    changes.append(line)
                    unpaired_old.remove(i)
                    unpaired_new.remove(j)
        lines += [f"- removed {describe(old[i], i)}" for i in unpaired_old]
        lines += changes
        lines += [f"+ added {describe(new[j], j)}" for j in unpaired_new]
    return lines


def only_differs(old: Entry, new: Entry, *fields: str) -> bool:
    """True when ``old`` and ``new`` differ, and only in ``fields``."""
    rest = [k for k in {*old, *new} if k not in fields]
    return old != new and all(old.get(k) == new.get(k) for k in rest)


# Comments


def comment_view(comments: Sequence[Entry]) -> list[Entry]:
    """Comments with ``reply_to`` replaced by the parent's author and text.

    A reply's raw ``reply_to`` is a position, so inserting an earlier comment would
    change every later reply; the parent's identity does not.
    """
    result = []
    for entry in comments:
        view = {k: v for k, v in entry.items() if k != "reply_to"}
        parent = entry.get("reply_to")
        if isinstance(parent, int) and 0 <= parent < len(comments):
            view["parent"] = [author(comments[parent]), comments[parent].get("text")]
        else:
            view["parent"] = None
        result.append(view)
    return result


def anchor(entry: Entry) -> str:
    value = entry.get("anchor")
    if value is None:
        return "with no anchor"
    if value == "":
        return "at a point" + where(entry.get("anchor_part"))
    return "on " + short(value) + where(entry.get("anchor_part"))


def describe_comment(entry: Entry) -> str:
    resolved = entry.get("resolved")
    state = " (resolved)" if resolved is True else " (open)" if resolved is False else ""
    if entry.get("parent") is not None:
        parent_author, parent_text = entry["parent"]
        return f"reply by {author(entry)} to {parent_author}'s comment {short(parent_text)}: {short(entry.get('text'))}{state}"
    return f"comment by {author(entry)} {anchor(entry)}: {short(entry.get('text'))}{state}"


# Field groups that one comment edit changes, and how to word the edit.
COMMENT_EDITS: list[tuple[tuple[str, ...], Callable[[Entry, Entry], str | None]]] = [
    # Resolved state is worded only when both sides know it; null means the document does not say.
    (("resolved",), lambda old, new: None if None in (old.get("resolved"), new.get("resolved")) else
        f"~ {'resolved' if new['resolved'] else 'reopened'} {describe_comment({**new, 'resolved': None})}"),
    (("anchor", "anchor_part"), lambda old, new:
        f"~ re-anchored {describe_comment({**new, 'resolved': None, 'anchor': old.get('anchor'), 'anchor_part': old.get('anchor_part')})}"
        f" → now {anchor(new)}"),
    (("text",), lambda old, new:
        f"~ edited text of {describe_comment({**old, 'resolved': None})} → {short(new.get('text'))}"),
]


def comments_section(old_raw: Sequence[Entry], new_raw: Sequence[Entry]) -> list[str]:
    old, new = comment_view(old_raw), comment_view(new_raw)
    blocks = aligned(old, new)
    unmatched_old = {i for olds, _ in blocks for i in olds}
    unmatched_new = {j for _, news in blocks for j in news}
    edits: dict[int, tuple[int, str]] = {}  # old position -> (new position, line)
    for fields, word in COMMENT_EDITS:
        def key(entry: Entry, fields: tuple[str, ...] = fields) -> str:
            return canonical({k: v for k, v in entry.items() if k not in fields})
        old_keys, new_keys = Counter(map(key, old)), Counter(map(key, new))
        new_position = {key(entry): j for j, entry in enumerate(new)}
        paired_new = {j for j, _ in edits.values()}
        for i in sorted(unmatched_old - edits.keys()):
            k = key(old[i])
            # The metadata must identify this comment on both sides: exactly one each.
            if old_keys[k] != 1 or new_keys[k] != 1:
                continue
            j = new_position[k]
            if j not in unmatched_new or j in paired_new or not only_differs(old[i], new[j], *fields):
                continue
            line = word(old[i], new[j])
            if line is not None:
                edits[i] = (j, line)
                paired_new.add(j)
    paired_new = {j for j, _ in edits.values()}
    lines = []
    for olds, news in blocks:
        for i in olds:
            lines.append(edits[i][1] if i in edits else f"- removed {describe_comment(old[i])}")
        lines += [f"+ added {describe_comment(new[j])}" for j in news if j not in paired_new]
    return lines


# Other categories


def describe_tracked(entry: Entry, _: int) -> str:
    kind = TRACKED.get(entry.get("kind", ""), f"tracked {entry.get('kind')}")
    return f"{kind} by {author(entry)}{where(entry.get('part'))}: {short(entry.get('text'))}"


def describe_link(entry: Entry, _: int) -> str:
    return f"link {short(entry.get('text'))} → {entry.get('target')}{where(entry.get('part'))}"


def change_link(old: Entry, new: Entry, _: int, __: int) -> str | None:
    if only_differs(old, new, "target"):
        return f"~ link {short(new.get('text'))}{where(new.get('part'))}: {old.get('target')} → {new.get('target')}"
    if only_differs(old, new, "text"):
        return f"~ link → {new.get('target')}{where(new.get('part'))}: label {short(old.get('text'))} → {short(new.get('text'))}"
    return None


def describe_field(entry: Entry, position: int) -> str:
    return f"field {position + 1} ({entry.get('command') or 'empty'}) in {entry.get('part')}"


def change_field(old: Entry, new: Entry, _: int, j: int) -> str | None:
    if only_differs(old, new, "sha256"):
        return f"~ field {j + 1} ({new.get('command') or 'empty'}) in {new.get('part')}: instruction changed"
    return None


def describe_image(entry: Entry, position: int) -> str:
    if "target" in entry:
        source = f"linked to {entry['target']}"
    elif entry.get("sha256") is None:
        source = "with a missing media part"
    else:
        source = f"{entry.get('bytes')} bytes, sha256 {entry['sha256'][:12]}"
    return f"image {position + 1} ({entry.get('part')}): {source}"


def change_image(old: Entry, new: Entry, _: int, j: int) -> str | None:
    if old.get("part") != new.get("part"):
        return None
    if "target" in old and "target" in new:
        return f"~ image {j + 1} ({new.get('part')}) link changed: {old['target']} → {new['target']}"
    return f"~ image {j + 1} ({new.get('part')}) replaced"


def describe_level(level: Entry) -> str:
    return f'{level.get("format")} "{level.get("text")}" from {level.get("start")}'


def describe_list(entry: Entry, _: int) -> str:
    levels = "; ".join(f"level {lv.get('level')} {describe_level(lv)}" for lv in entry.get("levels", []))
    return f"list {entry.get('list')} ({entry.get('part')}): {levels}"


def change_list(old: Entry, new: Entry, _: int, __: int) -> str | None:
    if not only_differs(old, new, "levels"):
        return None
    before = {lv.get("level"): lv for lv in old.get("levels", [])}
    after = {lv.get("level"): lv for lv in new.get("levels", [])}
    details = []
    for number in sorted(before.keys() | after.keys()):
        a, b = before.get(number), after.get(number)
        if a == b:
            continue
        if a is None or b is None:
            details.append(f"level {number} {'added' if a is None else 'removed'}: {describe_level(b or a or {})}")
            continue
        parts = []
        if a.get("format") != b.get("format"):
            parts.append(f"{a.get('format')} → {b.get('format')}")
        if a.get("text") != b.get("text"):
            parts.append(f'level text "{a.get("text")}" → "{b.get("text")}"')
        if a.get("start") != b.get("start"):
            parts.append(f"start {a.get('start')} → {b.get('start')}")
        details.append(f"level {number}: " + "; ".join(parts))
    return f"~ list {new.get('list')} " + "; ".join(details)


def describe_run(entry: Entry, _: int) -> str:
    items = entry.get("items")
    return (f"{items} item{'s' if items != 1 else ''} in list {entry.get('list')} level {entry.get('level')}"
            f" ({entry.get('part')})")


def change_run(old: Entry, new: Entry, _: int, __: int) -> str | None:
    if only_differs(old, new, "items"):
        return f"~ list {new.get('list')} level {new.get('level')} ({new.get('part')}): {old.get('items')} → {new.get('items')} items"
    if only_differs(old, new, "level"):
        return (f"~ {new.get('items')} item{'s' if new.get('items') != 1 else ''} of list {new.get('list')}"
                f" ({new.get('part')}) moved from level {old.get('level')} to level {new.get('level')}")
    return None


def notice(title: str, lines: list[str], differs: bool) -> list[str]:
    if lines:
        return [f"{title}:", *lines]
    # The data differs but no entry could be described: never report "no changes".
    return [f"{title}: metadata changed (could not match entries one-to-one)"] if differs else []


def review(old: Extraction, new: Extraction) -> tuple[str, bool]:
    """Return the report body (everything after the header line) and whether anything differs."""
    a, b = old.metadata, new.metadata
    old_review, new_review = a.get("review", []), b.get("review", [])
    old_comments = [e for e in old_review if e.get("kind") == "comment"]
    new_comments = [e for e in new_review if e.get("kind") == "comment"]
    old_tracked = [e for e in old_review if e.get("kind") != "comment"]
    new_tracked = [e for e in new_review if e.get("kind") != "comment"]
    old_warnings, new_warnings = a.get("warnings", []), b.get("warnings", [])
    sections = [
        notice("Comments", comments_section(old_comments, new_comments), old_comments != new_comments),
        notice("Tracked changes", entries_section(old_tracked, new_tracked, describe_tracked),
               old_tracked != new_tracked),
        notice("Hyperlinks", entries_section(a.get("hyperlinks", []), b.get("hyperlinks", []),
                                             describe_link, change_link),
               a.get("hyperlinks") != b.get("hyperlinks")),
        notice("Fields", entries_section(a.get("fields", []), b.get("fields", []), describe_field, change_field),
               a.get("fields") != b.get("fields")),
        notice("Images", entries_section(a.get("images", []), b.get("images", []), describe_image, change_image),
               a.get("images") != b.get("images")),
        notice("Lists", entries_section(a.get("lists", []), b.get("lists", []), describe_list, change_list)
               + entries_section(a.get("list_runs", []), b.get("list_runs", []), describe_run, change_run),
               (a.get("lists"), a.get("list_runs")) != (b.get("lists"), b.get("list_runs"))),
        notice("Warnings", [f"- removed: {w}" for w in old_warnings if w not in new_warnings]
               + [f"+ added: {w}" for w in new_warnings if w not in old_warnings],
               old_warnings != new_warnings),
    ]
    known = {"review", "hyperlinks", "fields", "images", "lists", "list_runs", "warnings"}
    sections += [[f"Metadata: {key} changed"] for key in sorted((a.keys() | b.keys()) - known)
                 if a.get(key) != b.get(key)]
    sections = [s for s in sections if s]
    prose = prose_section(old.prose, new.prose)
    changed = bool(sections) or old.prose != new.prose
    blocks = [prose, *sections]
    if changed:
        warnings = sorted(set(old_warnings) | set(new_warnings))
        if warnings:
            blocks.append(["Coverage warnings (either side):", *(f"  {w}" for w in warnings)])
        blocks.append(["Only the prose and review metadata Pilcrow extracts are compared"
                       " (not a complete document comparison)."])
    else:
        blocks.append([CAVEAT])
    return "\n\n".join("\n".join(block) for block in blocks) + "\n", changed
