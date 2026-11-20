"""Deliberately small English-oriented sentence segmentation, not an NLP model."""

import re

ABBREVIATIONS = frozenset(
    "mr mrs ms dr prof sr jr st vs etc e.g i.e fig figs eq eqs ref refs vol no pp al".split()
)
# Split only after terminal punctuation followed by whitespace. Keep closing quotes
# and reference brackets with the sentence. Decimal points have no following space.
BOUNDARY = re.compile(r'''[.!?]+[”’"')\]]*\s+''')
# One character of the token a full stop ends; a boundary ends in whitespace, so a
# token never reaches back past the previous boundary.
TOKEN_CHAR = re.compile(r"[\w.]")
INITIALISM = re.compile(r"(?:[A-Za-z]\.)*[A-Za-z]")


def normalize(text: str) -> str:
    return " ".join(text.split())


def abbreviation_candidate(line: str, stop: int) -> str:
    """The run of word characters and dots just before a full stop at ``stop``, or "".

    Scans back only over that run, so splitting a line costs time linear in its length.
    """
    if line[stop] != ".":
        return ""
    first = stop
    while first > 0 and TOKEN_CHAR.match(line, first - 1):
        first -= 1
    return line[first:stop]


def sentences(text: str) -> list[str]:
    """Keep explicit Word line breaks; normalize other whitespace after joining runs."""
    result = []
    for line in text.split("\n"):
        line = normalize(line)
        start = 0
        for match in BOUNDARY.finditer(line):
            end = match.end()
            token = abbreviation_candidate(line, match.start())
            if token and (token.lower() in ABBREVIATIONS or INITIALISM.fullmatch(token)):
                continue
            result.append(line[start: end].strip())
            start = end
        if line[start:]:
            result.append(line[start:])
    return result
