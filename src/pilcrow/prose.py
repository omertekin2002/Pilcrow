"""Deliberately small English-oriented sentence segmentation, not an NLP model."""

import re

ABBREVIATIONS = frozenset(
    "mr mrs ms dr prof sr jr st vs etc e.g i.e fig figs eq eqs ref refs vol no pp al".split()
)
# Split only after terminal punctuation followed by whitespace. Keep closing quotes
# and reference brackets with the sentence. Decimal points have no following space.
BOUNDARY = re.compile(r'''[.!?]+[”’"')\]]*\s+''')


def normalize(text: str) -> str:
    return " ".join(text.split())


def sentences(text: str) -> list[str]:
    """Keep explicit Word line breaks; normalize other whitespace after joining runs."""
    result = []
    for line in text.split("\n"):
        line = normalize(line)
        start = 0
        for match in BOUNDARY.finditer(line):
            end = match.end()
            prefix = line[:match.start() + 1]
            token = re.search(r"([\w.]+)\.$", prefix)
            if token and (token[1].lower() in ABBREVIATIONS
                          or re.fullmatch(r"(?:[A-Za-z]\.)*[A-Za-z]", token[1])):
                continue
            result.append(line[start: end].strip())
            start = end
        if line[start:]:
            result.append(line[start:])
    return result
