"""Describe prose and review-metadata changes between two extractions in words.

Read-only: it compares two in-memory extractions and never touches files or Git.
Entries are matched by content alignment, never by position alone, and a change is
worded as an edit to one entry only when the metadata identifies that entry.
"""

from __future__ import annotations

import difflib
import json
import re
from collections import Counter
from collections.abc import Callable, Sequence
from dataclasses import dataclass
from typing import Any

from .extract import Extraction

MAX_DIFF_LINES = 2000
MAX_TEXT = 80
BODY = "word/document.xml"
CAVEAT = "No supported prose or review changes detected (not a complete document comparison)."
CLOSING = "Only the prose and review metadata Pilcrow extracts are compared (not a complete document comparison)."
COVERAGE = "Coverage warnings (either side)"
FORMATS = ("text", "markdown")
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


# Word marks (--word-diff)

# A replaced line pair is marked word by word only when its word tokens are at least this similar.
MIN_WORD_SIMILARITY = 0.5
# Lines with more word tokens than this are never paired: the whole lines are shown instead.
MAX_WORD_TOKENS = 200
# Lines that already contain one of these could not be read back unambiguously, so are never marked.
MARKS = ("[-", "-]", "{+", "+}")
# Runs of word characters, runs of whitespace, and single other characters: joined, they are the line.
TOKEN = re.compile(r"\w+|\s+|[^\w\s]")


def tokens(line: str) -> list[str]:
    return TOKEN.findall(line)


def words(line_tokens: list[str]) -> list[str]:
    return [t for t in line_tokens if not t.isspace()]


def word_marks(old: str, new: str) -> str:
    """``old`` turned into ``new``, with removed tokens as [-…-] and added tokens as {+…+}.

    Changes separated only by unchanged whitespace merge into one mark, so a changed
    phrase reads as one edit. Dropping every {+…+} and the brackets of every [-…-]
    gives back ``old``; dropping every [-…-] and the braces of every {+…+} gives ``new``.
    """
    a, b = tokens(old), tokens(new)
    # Whitespace never starts a match (it would make long lines slow to align), but
    # unchanged whitespace next to a matched word is still part of the match.
    opcodes = difflib.SequenceMatcher(str.isspace, a, b, autojunk=False).get_opcodes()
    tags = [tag for tag, *_ in opcodes]
    spans = [[i1, i2, j1, j2] for _, i1, i2, j1, j2 in opcodes]
    # An inserted or deleted " word" can equally be read as "word "; always take the
    # second, so an added word reads "deliver {+all +}the" wherever the matcher put it.
    # Opcodes alternate between equal and changed, so a change's neighbours are equal.
    for n in range(1, len(opcodes) - 1):
        if tags[n] not in ("insert", "delete"):
            continue
        before, change, after = spans[n - 1:n + 2]
        side, lo, hi = (b, change[2], change[3]) if tags[n] == "insert" else (a, change[0], change[1])
        while after[1] - after[0] > 1 and side[lo].isspace() and side[lo] == side[hi]:
            before[1], before[3] = before[1] + 1, before[3] + 1
            change[:] = [k + 1 for k in change]
            after[0], after[2] = after[0] + 1, after[2] + 1
            lo, hi = lo + 1, hi + 1
    # An equal opcode that is neither first nor last sits between two changes; when it
    # is only whitespace, it joins them into one mark.
    runs: list[tuple[bool, list[int]]] = []
    for n, (tag, span) in enumerate(zip(tags, spans, strict=True)):
        changed = tag != "equal" or (0 < n < len(tags) - 1 and all(t.isspace() for t in a[span[0]:span[1]]))
        if changed and runs and runs[-1][0]:
            runs[-1][1][1], runs[-1][1][3] = span[1], span[3]
        else:
            runs.append((changed, list(span)))
    result = []
    for changed, (i1, i2, j1, j2) in runs:
        removed, added = "".join(a[i1:i2]), "".join(b[j1:j2])
        if not changed:
            result.append(removed)
            continue
        result.append((f"[-{removed}-]" if removed else "") + (f"{{+{added}+}}" if added else ""))
    return "".join(result)


def pairable(old: str, new: str) -> bool:
    """True when ``old`` and ``new`` can be shown as one marked line (see MIN_WORD_SIMILARITY)."""
    if any(mark in line for line in (old, new) for mark in MARKS):
        return False
    a, b = words(tokens(old)), words(tokens(new))
    if len(a) > MAX_WORD_TOKENS or len(b) > MAX_WORD_TOKENS:
        return False
    return difflib.SequenceMatcher(None, a, b, autojunk=False).ratio() >= MIN_WORD_SIMILARITY


def paired(olds: list[str], news: list[str], shown: int | None = None) -> list[str] | None:
    """One line per pair of a replaced block, or None when the block must be shown as whole lines.

    Lines pair 1:1 by position only when the block replaces as many lines as it adds and
    every pair is pairable; an unchanged pair (a line the alignment did not match) is context.
    Only the first ``shown`` lines (all by default) are worded; the rest, past the
    MAX_DIFF_LINES cap, are only counted, so they are left empty.
    """
    if len(olds) != len(news) or not all(pairable(a, b) for a, b in zip(olds, news, strict=True)):
        return None
    return ["" if shown is not None and n >= shown else f" {a}" if a == b else f"~ {word_marks(a, b)}"
            for n, (a, b) in enumerate(zip(olds, news, strict=True))]


def span(start: int, stop: int) -> str:
    """A unified diff hunk range, as difflib writes it."""
    length = stop - start
    if length == 1:
        return f"{start + 1}"
    return f"{start if length == 0 else start + 1},{length}"


def word_diff(old: list[str], new: list[str]) -> list[str]:
    """The unified diff of difflib.unified_diff, with each pairable replaced block as marked lines."""
    lines = ["--- OLD", "+++ NEW"]
    for group in difflib.SequenceMatcher(None, old, new).get_grouped_opcodes(3):
        first, last = group[0], group[-1]
        lines.append(f"@@ -{span(first[1], last[2])} +{span(first[3], last[4])} @@")
        for tag, i1, i2, j1, j2 in group:
            if tag == "equal":
                lines += [f" {line}" for line in old[i1:i2]]
                continue
            marked = paired(old[i1:i2], new[j1:j2], MAX_DIFF_LINES - len(lines)) if tag == "replace" else None
            if marked is not None:
                lines += marked
                continue
            lines += [f"-{line}" for line in old[i1:i2]]
            lines += [f"+{line}" for line in new[j1:j2]]
    return lines


def prose_diff(old: str, new: str, marks: bool = False) -> list[str]:
    """The unified prose diff, cut at MAX_DIFF_LINES with a note; empty when the prose is unchanged.

    With ``marks``, replaced lines that pair up are shown once, with word marks (see ``paired``).
    """
    if old == new:
        return []
    if marks:
        diff = word_diff(old.splitlines(), new.splitlines())
    else:
        diff = list(difflib.unified_diff(old.splitlines(), new.splitlines(), "OLD", "NEW", n=3, lineterm=""))
    lines = diff[:MAX_DIFF_LINES]
    if len(diff) > MAX_DIFF_LINES:
        lines.append(f"… {len(diff) - MAX_DIFF_LINES} more diff lines omitted")
    return lines


def aligned(old: Sequence[Entry], new: Sequence[Entry],
            key: Callable[[Entry], str] = canonical) -> list[tuple[list[int], list[int]]]:
    """The unmatched (old positions, new positions) of each differing block of an alignment."""
    a, b = [key(e) for e in old], [key(e) for e in new]
    # Trim the common ends first: aligning many repeated entries is quadratic, and most
    # reviews change a few entries of a long, otherwise identical sequence.
    head = 0
    while head < min(len(a), len(b)) and a[head] == b[head]:
        head += 1
    tail = 0
    while tail < min(len(a), len(b)) - head and a[-1 - tail] == b[-1 - tail]:
        tail += 1
    if head + tail == len(a) == len(b):
        return []
    matcher = difflib.SequenceMatcher(None, a[head:len(a) - tail], b[head:len(b) - tail], autojunk=False)
    return [(list(range(head + i1, head + i2)), list(range(head + j1, head + j2)))
            for tag, i1, i2, j1, j2 in matcher.get_opcodes() if tag != "equal"]


def entries_section(old: Sequence[Entry], new: Sequence[Entry], describe: Callable[[Entry, int], str],
                    change: Change | None = None, key: Callable[[Entry], str] = canonical) -> list[str]:
    """Removed/added lines, with ``change`` wording pairs in the same position of a replaced block."""
    lines = []
    for olds, news in aligned(old, new, key):
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
    """Comments with ``reply_to`` replaced by the parent's identity.

    A reply's raw ``reply_to`` is a position, so inserting an earlier comment would
    change every later reply; the parent's author, text and anchor do not. When that
    identity is shared by several comments, the position is kept too, so moving a
    reply between look-alike parents is still a difference.
    """
    def identity(entry: Entry) -> list[Any]:
        return [author(entry), entry.get("text"), entry.get("anchor"), entry.get("anchor_part")]
    shared = Counter(canonical(identity(entry)) for entry in comments)
    result = []
    for entry in comments:
        view = {k: v for k, v in entry.items() if k != "reply_to"}
        parent = entry.get("reply_to")
        if isinstance(parent, int) and 0 <= parent < len(comments):
            view["parent"] = identity(comments[parent])
            if shared[canonical(view["parent"])] > 1:
                view["parent"].append(parent)
        else:
            view["parent"] = parent  # None, or a position that names no comment
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
    parent = entry.get("parent")
    if isinstance(parent, list):
        parent_author, parent_text, parent_anchor, parent_part = parent[:4]
        which = f" (comment {parent[4] + 1})" if len(parent) > 4 else ""  # look-alike parents
        on = anchor({"anchor": parent_anchor, "anchor_part": parent_part})
        return (f"reply by {author(entry)} to {parent_author}'s comment {short(parent_text)} {on}{which}:"
                f" {short(entry.get('text'))}{state}")
    if parent is not None:
        return f"reply by {author(entry)} to a missing comment {parent + 1}: {short(entry.get('text'))}{state}"
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
        f"~ edited text of {describe_comment({**old, 'parent': new.get('parent'), 'resolved': None})} → {short(new.get('text'))}"),
]


def valid_parent(raw: Sequence[Entry], index: int) -> int | None:
    """The parent comment position of ``raw[index]`` when it names an existing comment, else None."""
    parent = raw[index].get("reply_to")
    return parent if isinstance(parent, int) and 0 <= parent < len(raw) else None


def through_pairs(view: Sequence[Entry], raw: Sequence[Entry], pairs: dict[int, int], side: str) -> list[Entry]:
    """``view`` with each reply's parent replaced by its parent's pairing, or an unpaired marker.

    ``pairs`` maps a paired parent's position to the position of the new comment it is
    paired with, so two replies whose parents are paired with each other compare equal
    however the parent itself changed, while a reply whose parent is not paired cannot
    match through its parent's raw identity.
    """
    result = []
    for index, entry in enumerate(view):
        parent = valid_parent(raw, index)
        if parent is None:
            result.append(entry)
        elif parent in pairs:
            result.append({**entry, "parent": {"paired": pairs[parent]}})
        else:
            result.append({**entry, "parent": {"unpaired": side, "index": parent}})
    return result


def comments_section(old_raw: Sequence[Entry], new_raw: Sequence[Entry]) -> list[str]:
    old, new = comment_view(old_raw), comment_view(new_raw)
    blocks = aligned(old, new)
    unmatched_old = {i for olds, _ in blocks for i in olds}
    unmatched_new = {j for _, news in blocks for j in news}
    # Entries outside the differing blocks are equal and pair up in order.
    pairs = dict(zip((i for i in range(len(old)) if i not in unmatched_old),
                     (j for j in range(len(new)) if j not in unmatched_new), strict=True))
    # A matched reply only counts as unchanged when its parent is paired with its new parent;
    # hold back any others until a later round pairs their parents, or leave them unmatched.
    pruned: set[tuple[int, int]] = set()
    while True:
        drop = [(i, j) for i, j in pairs.items()
                if (op := valid_parent(old_raw, i)) is not None
                and pairs.get(op) != valid_parent(new_raw, j)]
        if not drop:
            break
        for i, j in drop:
            del pairs[i]
            pruned.add((i, j))
    if pruned:
        anchors = [*sorted(pairs.items()), (len(old), len(new))]
        blocks, prev_i, prev_j = [], 0, 0
        for i, j in anchors:
            if i > prev_i or j > prev_j:
                blocks.append((list(range(prev_i, i)), list(range(prev_j, j))))
            prev_i, prev_j = i + 1, j + 1
        unmatched_old = {i for olds, _ in blocks for i in olds}
        unmatched_new = {j for _, news in blocks for j in news}
    edits: dict[int, tuple[int, str | None]] = {}  # old position -> (new position, line or None for no change)
    # Pair parents before their replies: each round compares replies through the pairs of the
    # rounds before it, and the rounds stop once one pairs nothing, even on a reply_to cycle.
    found = True
    while found:
        found = False
        new_paired = set(pairs.values())
        old_now = through_pairs(old, old_raw, pairs, "old")
        new_now = through_pairs(new, new_raw, {j: j for j in new_paired}, "new")
        # Exact matches through a paired parent: pair in order even for identical replies.
        by_key: dict[str, list[int]] = {}
        for j in sorted(unmatched_new - new_paired):
            by_key.setdefault(canonical(new_now[j]), []).append(j)
        for i in sorted(unmatched_old - pairs.keys()):
            candidates = by_key.get(canonical(old_now[i]))
            if not candidates:
                continue
            # Match a reply whose parent changed, or one aligned in place before its parent paired;
            # an entry equal on its own that moved stays removed and added.
            j = next((c for c in candidates if old[i] != new[c] or (i, c) in pruned), None)
            if j is None:
                continue
            candidates.remove(j)
            pairs[i] = j
            edits[i] = (j, None)
            new_paired.add(j)
            found = True
        for fields, word in COMMENT_EDITS:
            def key(entry: Entry, fields: tuple[str, ...] = fields) -> str:
                return canonical({k: v for k, v in entry.items() if k not in fields})
            old_keys, new_keys = Counter(map(key, old_now)), Counter(map(key, new_now))
            new_position = {key(entry): j for j, entry in enumerate(new_now)}
            for i in sorted(unmatched_old - pairs.keys()):
                k = key(old_now[i])
                # The metadata must identify this comment on both sides: exactly one each.
                if old_keys[k] != 1 or new_keys[k] != 1:
                    continue
                j = new_position[k]
                if j not in unmatched_new or j in new_paired or not only_differs(old_now[i], new_now[j], *fields):
                    continue
                line = word(old[i], new[j])
                if line is None:
                    continue
                pairs[i] = j
                edits[i] = (j, line)
                new_paired.add(j)
                found = True
    paired_new = {j for j, _ in edits.values()}
    lines = []
    for olds, news in blocks:
        for i in olds:
            line = edits[i][1] if i in edits else f"- removed {describe_comment(old[i])}"
            if line is not None:
                lines.append(line)
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


def form_of(entry: Entry) -> Entry:
    """A field's legacy form state (checkbox or dropdown), or {} for any other field."""
    form = entry.get("form")
    return form if isinstance(form, dict) else {}


def field_label(entry: Entry) -> str:
    """"FORMCHECKBOX "Check1"" for a named form field, else the field's command."""
    command = entry.get("command") or "empty"
    name = form_of(entry).get("name")
    return f"{command} {short(name)}" if name else command


def form_state(form: Entry) -> str:
    """A checkbox's state or a dropdown's selected entry, as the review words it."""
    if form.get("type") == "checkbox":
        return "checked" if form.get("checked") else "unchecked"
    selected = form.get("selected")
    return short(selected) if isinstance(selected, str) else "(no entry selected)"


def describe_field(entry: Entry, position: int) -> str:
    form = form_of(entry)
    state = f": {form_state(form)}" if form.get("type") in ("checkbox", "dropdown") else ""
    return f"field {position + 1} ({field_label(entry)}) in {entry.get('part')}{state}"


def change_field(old: Entry, new: Entry, _: int, j: int) -> str | None:
    if only_differs(old, new, "sha256"):
        return f"~ field {j + 1} ({new.get('command') or 'empty'}) in {new.get('part')}: instruction changed"
    if only_differs(old, new, "form"):
        a, b = form_of(old), form_of(new)
        # A form field whose kind or name changed, or that gained or lost its form data, is removed and added.
        if not only_differs(a, b, "checked", "entries", "selected") or a.get("type") not in ("checkbox", "dropdown"):
            return None
        details = []
        if a.get("entries") != b.get("entries"):
            details.append("entries changed")
        if (a.get("checked"), a.get("selected")) != (b.get("checked"), b.get("selected")):
            details.append(f"{form_state(a)} → {form_state(b)}")
        if details:
            return f"~ field {j + 1} ({field_label(new)}) in {new.get('part')}: " + "; ".join(details)
    return None


# Alt text metadata keys and how the review names them.
ALT_TEXT = (("alt", "alt text"), ("alt_title", "alt title"))


def describe_image(entry: Entry, position: int) -> str:
    # An image can hold both an embedded copy (sha256/bytes) and an external link (target).
    sources = []
    if "sha256" in entry:
        sha = entry["sha256"]
        sources.append("with a missing media part" if sha is None else f"{entry.get('bytes')} bytes, sha256 {sha[:12]}")
    if "target" in entry:
        sources.append(f"linked to {entry['target']}")
    labels = [f"{name} {short(entry[key])}" for key, name in ALT_TEXT if entry.get(key)]
    return f"image {position + 1} ({entry.get('part')}): {', '.join(sources) or 'no source'}" + "".join(
        f", {label}" for label in labels)


def alt_value(entry: Entry, key: str) -> str:
    value = entry.get(key)
    return short(value) if value else "(none)"


def change_image(old: Entry, new: Entry, _: int, j: int) -> str | None:
    """Word a hash, link or alt text change of one image; anything else falls back to removed/added."""
    if old.get("part") != new.get("part") or ("sha256" in old) != ("sha256" in new):
        return None
    details = []
    if (old.get("sha256"), old.get("bytes")) != (new.get("sha256"), new.get("bytes")):
        if None in (old.get("sha256"), new.get("sha256")):
            return None  # a media part went missing or appeared: describe both sides in full
        details.append("replaced")
    if old.get("target") != new.get("target"):
        details.append(f"link changed: {old.get('target') or '(none)'} → {new.get('target') or '(none)'}")
    details += [f"{name}: {alt_value(old, key)} → {alt_value(new, key)}"
                for key, name in ALT_TEXT if old.get(key) != new.get(key)]
    return f"~ image {j + 1} ({new.get('part')}) " + "; ".join(details) if details else None


# Object kinds shown under "Equations and symbols"; every other kind (chart, diagram,
# embedded object, imported chunk, WordArt) is shown under "Charts, diagrams and embedded objects".
FORMULA_KINDS = frozenset({"equation", "symbol"})


def objects_of(metadata: Entry, formulas: bool) -> list[Entry]:
    """The ``objects`` entries of one section: equations and symbols, or every other kind."""
    return [e for e in metadata.get("objects", []) if (e.get("kind") in FORMULA_KINDS) == formulas]


def numbered_objects(objects: Sequence[Entry]) -> list[Entry]:
    """Objects, each with its 1-based number among its own kind."""
    counts: Counter[str] = Counter()
    result = []
    for entry in objects:
        counts[entry.get("kind", "")] += 1
        result.append({**entry, "number": counts[entry.get("kind", "")]})
    return result


def unnumbered(entry: Entry) -> str:
    """An object's identity for alignment: its number shifts when an earlier one is added."""
    return canonical({k: v for k, v in entry.items() if k != "number"})


def symbol_char(entry: Entry) -> str:
    return f"{entry.get('char') or '(no char)'} ({entry.get('font') or 'no font'})"


def describe_object(entry: Entry, _: int) -> str:
    kind = entry.get("kind")
    label = f"{kind} {entry.get('number')} ({entry.get('part')})"
    if kind == "symbol":
        return f"{label}: {symbol_char(entry)}"
    if kind == "equation":
        return f"{label}, sha256 {str(entry.get('sha256'))[:12]}"
    return label


def one_spaced(value: Any) -> str:
    """Document-controlled text on one line, so it cannot start a line of its own in a text report."""
    return " ".join(str(value).split())


def part_label(entry: Entry) -> str:
    """"chart 1 (word/document.xml)", with an embedded object's ProgID before the part."""
    prog_id = one_spaced(entry.get("prog_id") or "")
    place = f"{prog_id}, {entry.get('part')}" if prog_id else entry.get("part")
    return f"{entry.get('kind')} {entry.get('number')} ({place})"


def describe_part_object(entry: Entry, _: int) -> str:
    if entry.get("kind") == "WordArt":
        return f"{part_label(entry)}: {short(entry.get('text'))}"
    sources = []
    if "sha256" in entry:
        sha = entry["sha256"]
        sources.append("with a missing part" if sha is None else f"sha256 {sha[:12]}")
    if "target" in entry:
        sources.append(f"linked to {one_spaced(entry['target'])}")
    return f"{part_label(entry)}: {', '.join(sources) or 'no source'}"


def change_part_object(old: Entry, new: Entry, _: int, __: int) -> str | None:
    """Word a content change of one chart, diagram, embedded object or imported chunk.

    Only a hash change between two present parts, or a WordArt text change, is an edit;
    anything else falls back to removed/added.
    """
    if old.get("kind") == new.get("kind") == "WordArt" and only_differs(old, new, "text", "number"):
        return f"~ {part_label(new)}: {short(old.get('text'))} → {short(new.get('text'))}"
    if None in (old.get("sha256"), new.get("sha256")) or not only_differs(old, new, "sha256", "number"):
        return None
    return f"~ {part_label(new)} changed"


def change_object(old: Entry, new: Entry, _: int, __: int) -> str | None:
    """Word an edit to one equation or symbol; anything else falls back to removed/added."""
    if old.get("kind") != new.get("kind") or old.get("part") != new.get("part"):
        return None
    label = f"{new.get('kind')} {new.get('number')} ({new.get('part')})"
    if new.get("kind") == "equation" and only_differs(old, new, "sha256", "number"):
        return f"~ {label} changed"
    if new.get("kind") == "symbol" and only_differs(old, new, "font", "char", "number"):
        return f"~ {label}: {symbol_char(old)} → {symbol_char(new)}"
    return None


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


@dataclass
class Section:
    """One metadata section: a title and its lines, each a self-contained statement."""

    title: str
    lines: list[str]
    inline: bool = False  # a single summary line, printed as "Title: line" in text


@dataclass
class Report:
    """What changed, independent of how it is printed."""

    diff: list[str]  # the prose diff with its cap note; empty when the prose is unchanged
    sections: list[Section]
    warnings: list[str]  # coverage warnings of either side, sorted
    changed: bool


def notice(title: str, lines: list[str], differs: bool) -> Section | None:
    if lines:
        return Section(title, lines)
    # The data differs but no entry could be described: never report "no changes".
    return Section(title, ["metadata changed (could not match entries one-to-one)"], inline=True) if differs else None


def build(old: Extraction, new: Extraction, marks: bool = False) -> Report:
    """Compare two extractions; render the result with ``text`` or ``markdown``.

    ``marks`` marks the changed words of replaced prose lines that pair up (``--word-diff``).
    """
    a, b = old.metadata, new.metadata
    old_review, new_review = a.get("review", []), b.get("review", [])
    old_comments = [e for e in old_review if e.get("kind") == "comment"]
    new_comments = [e for e in new_review if e.get("kind") == "comment"]
    old_tracked = [e for e in old_review if e.get("kind") != "comment"]
    new_tracked = [e for e in new_review if e.get("kind") != "comment"]
    old_warnings, new_warnings = a.get("warnings", []), b.get("warnings", [])
    # Both object sections can be unchanged while their entries moved past each other, e.g. a
    # block-level equation moved across a chart: never report "no changes" then.
    objects_moved = (a.get("objects", []) != b.get("objects", [])
                     and all(objects_of(a, kind) == objects_of(b, kind) for kind in (True, False)))
    found = [
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
        notice("Equations and symbols",
               entries_section(numbered_objects(objects_of(a, True)), numbered_objects(objects_of(b, True)),
                               describe_object, change_object, unnumbered),
               objects_of(a, True) != objects_of(b, True)),
        notice("Charts, diagrams and embedded objects",
               entries_section(numbered_objects(objects_of(a, False)), numbered_objects(objects_of(b, False)),
                               describe_part_object, change_part_object, unnumbered),
               objects_of(a, False) != objects_of(b, False)),
        Section("Objects", ["equations or symbols moved relative to charts, diagrams or embedded objects"],
                inline=True) if objects_moved else None,
        notice("Lists", entries_section(a.get("lists", []), b.get("lists", []), describe_list, change_list)
               + entries_section(a.get("list_runs", []), b.get("list_runs", []), describe_run, change_run),
               (a.get("lists"), a.get("list_runs")) != (b.get("lists"), b.get("list_runs"))),
        notice("Warnings", [f"- removed: {w}" for w in old_warnings if w not in new_warnings]
               + [f"+ added: {w}" for w in new_warnings if w not in old_warnings],
               old_warnings != new_warnings),
    ]
    known = {"review", "hyperlinks", "fields", "images", "objects", "lists", "list_runs", "warnings"}
    sections = [s for s in found if s is not None]
    sections += [Section("Metadata", [f"{key} changed"], inline=True)
                 for key in sorted((a.keys() | b.keys()) - known) if a.get(key) != b.get(key)]
    changed = bool(sections) or old.prose != new.prose
    return Report(prose_diff(old.prose, new.prose, marks), sections,
                  sorted(set(old_warnings) | set(new_warnings)), changed)


def text(report: Report) -> str:
    """The plain-text report body (everything after the header line)."""
    blocks = [["Prose:", *report.diff] if report.diff else ["Prose: no changes"]]
    blocks += [[f"{s.title}: {s.lines[0]}"] if s.inline else [f"{s.title}:", *s.lines] for s in report.sections]
    if report.changed:
        if report.warnings:
            blocks.append([f"{COVERAGE}:", *(f"  {w}" for w in report.warnings)])
        blocks.append([CLOSING])
    else:
        blocks.append([CAVEAT])
    return "\n\n".join("\n".join(block) for block in blocks) + "\n"


def review(old: Extraction, new: Extraction) -> tuple[str, bool]:
    """Return the plain-text report body (everything after the header line) and whether anything differs."""
    report = build(old, new)
    return text(report), report.changed


# Markdown

# Characters escaped wherever they appear: emphasis, code, links and images, headings
# (including a closing "#" sequence), tables, GFM strikethrough, and math. ":" and "@"
# (and the dot of "www.") stop GFM's autolink extension from taking a URL or address
# with its escapes still in it; GitHub then links the unescaped text it renders.
MARKDOWN_SPECIAL = frozenset("\\`*_[]!#|~$:@")
# GFM links a "www." at the start or after whitespace, "*", "_", "~" or "(". Escape its dot
# after anything but a letter or digit: "_" is a word character, so "\b" would miss "_www.".
WWW = re.compile(r"(?<![^\W_])(www)\.", re.IGNORECASE)
# Raw HTML and entity references are written as entities.
MARKDOWN_ENTITIES = {"<": "&lt;", ">": "&gt;", "&": "&amp;"}
LINE_BREAK = re.compile(r"\r\n?|\n")


def one_line(value: str) -> str:
    """``value`` with CommonMark line endings replaced by spaces, so it cannot start a new block."""
    return LINE_BREAK.sub(" ", value)


def escape(value: str) -> str:
    """``value`` as inline Markdown text that renders literally on one line."""
    value = one_line(value).strip()
    result = "".join(MARKDOWN_ENTITIES.get(c) or ("\\" + c if c in MARKDOWN_SPECIAL else c) for c in value)
    result = WWW.sub(r"\1\\.", result)
    # A leading marker would start a list, quote or setext underline: escape it too.
    return re.sub(r"^([-+=>])|^(\d+)([.)])", lambda m: f"\\{m[1]}" if m[1] else f"{m[2]}\\{m[3]}", result)


def longest_backticks(value: str) -> int:
    return max((len(run) for run in re.findall(r"`+", value)), default=0)


def code(value: str) -> str:
    """``value`` as a code span: delimited by a longer backtick run than any inside it."""
    value = one_line(value)
    ticks = "`" * (longest_backticks(value) + 1)
    pad = " " if value[:1] in ("`", " ") or value[-1:] in ("`", " ") else ""
    return f"{ticks}{pad}{value}{pad}{ticks}"


def bullet(line: str) -> str:
    """A list item for one text-format line, with its -/~/+ marker kept visible."""
    if line[:2] in ("- ", "~ ", "+ "):
        return f"- `{line[0]}` {escape(line[2:])}"
    return f"- {escape(line)}"


def markdown(report: Report, old: str, new: str) -> str:
    """The whole report, header included, as CommonMark that renders document text literally."""
    blocks = [f"## Review: {code(old)} → {code(new)}"]
    if report.diff:
        # A fence longer than any backtick run inside cannot be closed by document text.
        fence = "`" * max(3, longest_backticks("\n".join(report.diff)) + 1)
        blocks += ["### Prose", "\n".join([f"{fence}diff", *report.diff, fence])]
    else:
        blocks.append("Prose: no changes")
    for section in report.sections:
        blocks += [f"### {escape(section.title)}", "\n".join(map(bullet, section.lines))]
    if report.changed:
        if report.warnings:
            blocks += [f"### {escape(COVERAGE)}", "\n".join(f"- {escape(w)}" for w in report.warnings)]
        blocks.append(escape(CLOSING))
    else:
        blocks.append(escape(CAVEAT))
    return "\n\n".join(blocks) + "\n"
