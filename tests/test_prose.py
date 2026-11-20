"""Sentence segmentation: identical output to the original algorithm, in linear time.

Run with ``python -m unittest discover -s tests -v`` from the project root.
"""

from __future__ import annotations

import random
import re
import time
import unittest

from pilcrow.prose import ABBREVIATIONS, BOUNDARY, normalize, sentences


def reference_sentences(text: str) -> list[str]:
    """The algorithm before issue #26, kept verbatim as an oracle.

    It is quadratic in the length of one line: every boundary copies the line up to
    that point and searches it for the token before the full stop.
    """
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


CORPUS = [
    "",
    " ",
    "\n",
    "\n\n",
    ".",
    ". ",
    " . ",
    "..",
    "... ",
    "!",
    "? !",
    "One.",
    "One. Two.",
    "One.  Two.\tThree. Four.",
    "  Leading and trailing whitespace.   Second sentence.  ",
    "First line. Still first line.\nSecond line. Its second sentence.",
    "Line one.\n\nLine three after an empty line.\n",
    "Ask e.g. a friend. Or i.e. nobody. Then stop.",
    "See Dr. Smith and Mr. Jones vs. Mrs. Brown. They agree.",
    "Smith et al. found it. Others did not.",
    "See pp. 4-5 and fig. 2. Also Figs. 3 and Eq. 4. Done.",
    "Ref. 12 and refs. 13-14 and vol. 2, no. 3. Next.",
    "St. Ives is near. Sr. and Jr. too. Prof. X agrees. Etc. etc. The end.",
    "J. R. R. Tolkien wrote it. C. S. Lewis read it.",
    "Born in the U.S.A. in 1970. Moved to the U.K. later.",
    "Pi is 3.14 and e is 2.71. Both are irrational.",
    "Version 1.2.3. Version 4.5.6 is next.",
    "Really? Yes! Absolutely... Fine.",
    "Wait!? What?! Okay.",
    "He said “Stop.” Then he left.",
    "She said ‘go.’ He went.",
    "He said \"Stop.\" Then he left.",
    "She said 'go.' He went.",
    "A claim (see below.) Another claim.",
    "A citation.[1] Another.[2] Last.",
    "A claim [see 3.] Another.",
    "Nested.”) Quotes and brackets.”] End.",
    "Café au lait. Naïve résumé. Ünïcödé wörds. Straße.",
    "Ébauche é. Ω. Ж. Next sentence.",
    "日本語の文です。 English after it. Ok.",
    "under_score_word. digits123. 456. Next.",
    "Trailing abbreviation e.g.",
    "Trailing initial A.",
    "a. b. c. d. Done.",
    "A.B.C. Next. X.Y. Next.",
    "Mixed.dots.word. Next.",
    "....leading dots. Next.",
    ". leading boundary. Next.",
    "No terminal punctuation at all",
    "Tabs\tinside\ta sentence. And em spaces. Ok.",
    "Line break after abbreviation e.g.\nNew line starts here. Ok.",
    "MR. SHOUTING. DR. WHO. VS. THEM.",
    "Al. Gore. AL. Gore.",
]

FRAGMENTS = [
    "word", "Word", "words", "Another", "x", "A", "b", "J", "R", "Tolkien", "U.S.A", "U.K",
    "e.g", "i.e", "E.g", "Dr", "dr", "Mr", "Mrs", "Ms", "Prof", "St", "vs", "etc", "et al",
    "fig", "Figs", "eq", "Refs", "vol", "no", "pp", "3.14", "1.2.3", "42", "café", "naïve",
    "Ω", "Жук", "日本", "snake_case", "a.b", "...", "", "“", "”", "‘", "’", "\"", "'", "(", ")",
    "[", "]", "[1]",
]
PUNCTUATION = ["", "", ".", ".", ".", "!", "?", "...", "?!", ".”", ".’", '."', ".'", ".)", ".]",
               "!)", "?”", ".”)", ","]
WHITESPACE = [" ", " ", " ", "  ", "\t", "\n", " ", " \n ", ""]


def random_text(rng: random.Random) -> str:
    pieces = []
    for _ in range(rng.randrange(0, 40)):
        pieces.append(rng.choice(FRAGMENTS))
        pieces.append(rng.choice(PUNCTUATION))
        pieces.append(rng.choice(WHITESPACE))
    if rng.random() < 0.3:
        pieces.insert(0, rng.choice(WHITESPACE))
    return "".join(pieces)


class SentenceTests(unittest.TestCase):
    def test_matches_reference_on_corpus(self) -> None:
        for text in CORPUS:
            with self.subTest(text=text):
                self.assertEqual(sentences(text), reference_sentences(text))

    def test_matches_reference_on_seeded_random_text(self) -> None:
        rng = random.Random(20261120)
        for case in range(5000):
            text = random_text(rng)
            with self.subTest(case=case, text=text):
                self.assertEqual(sentences(text), reference_sentences(text))

    def test_matches_reference_on_abbreviation_list(self) -> None:
        for abbreviation in sorted(ABBREVIATIONS):
            for variant in (abbreviation, abbreviation.upper(), abbreviation.capitalize()):
                text = f"Before {variant}. after. Then {variant}.” more! End {variant}."
                with self.subTest(text=text):
                    self.assertEqual(sentences(text), reference_sentences(text))

    def test_long_paragraph_splits_in_linear_time(self) -> None:
        # The quadratic algorithm needs well over ten minutes for this line; the
        # linear one needs a fraction of a second. The bound leaves room for slow CI.
        count = 20_000
        text = " ".join(
            f"Sentence {n} cites Dr. Smith et al. and J. R. R. Tolkien, e.g. on pp. {n}-{n + 1}."
            for n in range(count))
        started = time.perf_counter()
        result = sentences(text)
        elapsed = time.perf_counter() - started
        self.assertEqual(len(result), count)
        self.assertEqual(result[0], "Sentence 0 cites Dr. Smith et al. and J. R. R. Tolkien, e.g. on pp. 0-1.")
        self.assertEqual(result[-1], f"Sentence {count - 1} cites Dr. Smith et al. and J. R. R. Tolkien, "
                                     f"e.g. on pp. {count - 1}-{count}.")
        self.assertLess(elapsed, 5.0)


if __name__ == "__main__":
    unittest.main()
