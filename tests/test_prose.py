"""Sentence segmentation: matches a naive oracle, in linear time.

Run with ``python -m unittest discover -s tests -v`` from the project root.
"""

from __future__ import annotations

import random
import re
import time
import unittest

from pilcrow.prose import (
    ABBREVIATIONS,
    BOUNDARY,
    NUMBER_ABBREVIATIONS,
    normalize,
    sentences,
)

# The abbreviation list before issue #54 added "cf" and "viz".
LEGACY_ABBREVIATIONS = ABBREVIATIONS - {"cf", "viz"}


def legacy_sentences(text: str) -> list[str]:
    """The algorithm before issue #26, kept verbatim with its abbreviation list.

    It is quadratic in the length of one line: every boundary copies the line up to
    that point and searches it for the token before the full stop. Its output is what
    companions exported before issue #54 contain.
    """
    result = []
    for line in text.split("\n"):
        line = normalize(line)
        start = 0
        for match in BOUNDARY.finditer(line):
            end = match.end()
            prefix = line[:match.start() + 1]
            token = re.search(r"([\w.]+)\.$", prefix)
            if token and (token[1].lower() in LEGACY_ABBREVIATIONS
                          or re.fullmatch(r"(?:[A-Za-z]\.)*[A-Za-z]", token[1])):
                continue
            result.append(line[start: end].strip())
            start = end
        if line[start:]:
            result.append(line[start:])
    return result


def reference_sentences(text: str) -> list[str]:
    """``legacy_sentences`` plus the rules of issue #54, written naively as an oracle.

    Each rule is a regex over the whole line before or after the boundary, not the
    linear scan ``sentences`` uses, so the two implementations stay independent.
    """
    result = []
    for line in text.split("\n"):
        line = normalize(line)
        start = 0
        for match in BOUNDARY.finditer(line):
            end = match.end()
            prefix = line[:match.start() + 1]
            suffix = line[end:]
            token = re.search(r"([\w.]+)\.$", prefix)
            # Rule 1: the next word starts lowercase.
            if re.match(r"\S", suffix) and suffix[0].islower():
                continue
            if token and (token[1].lower() in LEGACY_ABBREVIATIONS
                          or re.fullmatch(r"(?:[A-Za-z]\.)*[A-Za-z]", token[1])
                          # Rule 2: always joined.
                          or re.fullmatch(r"cf|viz", token[1], re.IGNORECASE)
                          # Rule 3: joined before a digit.
                          or (re.fullmatch(r"approx|ca|chap|ch|secs?|arts?|paras?|fn|nos|"
                                           r"jan|feb|mar|apr|jun|jul|aug|sept?|oct|nov|dec",
                                           token[1], re.IGNORECASE)
                              and re.match(r"\d", suffix))):
                continue
            result.append(line[start: end].strip())
            start = end
        if line[start:]:
            result.append(line[start:])
    return result


# Issue #54: each of these is one sentence (the legacy algorithm split each of them).
JOINED = [
    "See Art. 5 of the Act.",
    "As in para. 12 above.",
    "Under Sec. 3.2 we agree.",
    "See ch. 4 for more.",
    "Due Sept. 4 at noon.",
    "Since Jan. 2020 it grew.",
    "It is approx. 3 m long.",
    "There were ca. 40 people.",
    "Cf. Smith 2020 for details.",
    "She has a Ph.D. in physics.",
    "Acme Inc. was founded.",
    "Max. load is 5 kg.",
    "See p. 5 and nos. 3 and 4.",
    "Compare viz. Table 2 and CF. Jones.",
    "See Fn. 12 and chap. 3, secs. 4 and arts. 5-6, paras. 7 and 8.",
    "Dated Feb. 1, Mar. 2, Apr. 3, Jun. 4, Jul. 5, Aug. 6, Sep. 7, Oct. 8, Nov. 9 and DEC. 10.",
    "Une phrase. élégante suite.",
]

# Exact output, including the splits issue #54 keeps.
EXPECTED = {
    "Acme Inc. She left.": ["Acme Inc.", "She left."],
    "It was Jan. Then it snowed.": ["It was Jan.", "Then it snowed."],
    "See Art. The rest.": ["See Art.", "The rest."],
    "One. Two.": ["One.", "Two."],
    "Really? Yes! Absolutely... Fine.": ["Really?", "Yes!", "Absolutely...", "Fine."],
    "He said “Stop.” Then he left.": ["He said “Stop.”", "Then he left."],
    "J. R. R. Tolkien wrote it. C. S. Lewis read it.": ["J. R. R. Tolkien wrote it.", "C. S. Lewis read it."],
    "Born in the U.S.A. in 1970. Moved to the U.K. later.": ["Born in the U.S.A. in 1970.", "Moved to the U.K. later."],
    "Absolutely... fine. Wait! what? Okay.": ["Absolutely... fine.", "Wait! what?", "Okay."],
    "Max. 5 kg. The rest.": ["Max.", "5 kg.", "The rest."],
    "Read ch. 4. Then ch. 5.": ["Read ch. 4.", "Then ch. 5."],
    "In Art. “5” and Sec. (3).": ["In Art.", "“5” and Sec.", "(3)."],
    "Done in Dec. 2. And May. 3 more.": ["Done in Dec. 2.", "And May.", "3 more."],
    "The unit is \u212a. Next.": ["The unit is \u212a.", "Next."],
    **{text: [text] for text in JOINED},
}

# CORPUS entries whose output issue #54 changes; every other entry is unchanged.
CHANGED_BY_54 = {
    "under_score_word. digits123. 456. Next.",
    ". leading boundary. Next.",
}

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
    "under_score_word. digits123. 456. Next.",
    ". leading boundary. Next.",
    *EXPECTED,
]

FRAGMENTS = [
    "word", "Word", "words", "Another", "x", "A", "b", "J", "R", "Tolkien", "U.S.A", "U.K",
    "e.g", "i.e", "E.g", "Dr", "dr", "Mr", "Mrs", "Ms", "Prof", "St", "vs", "etc", "et al",
    "fig", "Figs", "eq", "Refs", "vol", "no", "pp", "3.14", "1.2.3", "42", "café", "naïve",
    "Ω", "Жук", "日本", "snake_case", "a.b", "...", "", "“", "”", "‘", "’", "\"", "'", "(", ")",
    "[", "]", "[1]", "cf", "Cf", "VIZ", "approx", "ca", "Ch", "chap", "Sec", "secs", "Art",
    "arts", "para", "Paras", "fn", "nos", "Jan", "sept", "Sep", "DEC", "May", "Inc", "Ph.D",
    "5", "2020", "3.2", "élan", "in", "ß",
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
    def test_expected_sentences(self) -> None:
        for text, expected in EXPECTED.items():
            with self.subTest(text=text):
                self.assertEqual(sentences(text), expected)

    def test_joined_cases_were_split_before(self) -> None:
        for text in JOINED:
            with self.subTest(text=text):
                self.assertGreater(len(legacy_sentences(text)), 1)

    def test_corpus_unaffected_by_54_is_unchanged(self) -> None:
        # Companions of documents without the forms issue #54 joins stay byte-identical.
        for text in CORPUS:
            if text in EXPECTED:
                continue
            with self.subTest(text=text):
                if text in CHANGED_BY_54:
                    self.assertNotEqual(sentences(text), legacy_sentences(text))
                else:
                    self.assertEqual(sentences(text), legacy_sentences(text))

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

    def test_number_abbreviations_join_only_before_a_digit(self) -> None:
        for abbreviation in sorted(NUMBER_ABBREVIATIONS):
            for variant in (abbreviation, abbreviation.upper(), abbreviation.capitalize()):
                text = f"See {variant}. 12 here. Then {variant}. Next. Or {variant}. (3) too."
                with self.subTest(text=text):
                    self.assertEqual(sentences(text), [f"See {variant}. 12 here.", f"Then {variant}.",
                                                       "Next.", f"Or {variant}.", "(3) too."])
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
