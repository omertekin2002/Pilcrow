"""Time extraction of a generated large document and of one very long paragraph.

Usage (Pilcrow must be importable by this Python, e.g. after ``pip install .``)::

    python scripts/benchmark.py                      # 500 pages, a 4,000-sentence paragraph
    python scripts/benchmark.py --pages 1000 --sentences 20000
    python scripts/benchmark.py --keep DIR           # also leave the generated DOCX files in DIR

The documents are generated deterministically in a temporary directory with the standard
library only. This is a manual tool for spotting slow paths, not a CI gate: timings
depend on the machine.
"""

import argparse
import tempfile
import time
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

from pilcrow.extract import extract

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_R = "http://schemas.openxmlformats.org/package/2006/relationships"
PARAGRAPHS_PER_PAGE = 13
SENTENCE = ("Section {n} cites Dr. Smith et al. and J. R. R. Tolkien, e.g. on pp. {n}-{m}, "
            "and the U.S.A. figures of 3.14 percent. Is that right? It is!")


def run(text: str) -> str:
    return f'<w:r><w:t xml:space="preserve">{escape(text)}</w:t></w:r>'


def paragraph(content: str, properties: str = "") -> str:
    return f"<w:p>{properties}{content}</w:p>"


def page(number: int) -> str:
    """Thirteen body paragraphs: a heading, prose, a list, a table, a comment and tracked changes."""
    n = number * 100
    text = SENTENCE.format(n=n, m=n + 1)
    comment = str(number)
    blocks = [
        paragraph(run(f"Chapter {number}"), '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>'),
        paragraph(run(text) + run(" " + SENTENCE.format(n=n + 2, m=n + 3))),
        paragraph(f'<w:commentRangeStart w:id="{comment}"/>' + run(text)
                  + f'<w:commentRangeEnd w:id="{comment}"/><w:r><w:commentReference w:id="{comment}"/></w:r>'),
        paragraph(run("Kept words. ")
                  + f'<w:ins w:id="{number * 2}" w:author="Ada" w:date="2024-01-02T03:04:05Z">{run("Inserted words. ")}</w:ins>'
                  + f'<w:del w:id="{number * 2 + 1}" w:author="Ben" w:date="2024-01-03T03:04:05Z">'
                    '<w:r><w:delText>Deleted words.</w:delText></w:r></w:del>'),
    ]
    for item in range(3):
        blocks.append(paragraph(run(f"List item {item}. {text}"),
                                f'<w:pPr><w:numPr><w:ilvl w:val="{item % 2}"/><w:numId w:val="1"/></w:numPr></w:pPr>'))
    rows = "".join(
        "<w:tr>" + "".join(f"<w:tc>{paragraph(run(f'Cell {row}-{col}. Value {n + col}.'))}</w:tc>"
                           for col in range(2)) + "</w:tr>"
        for row in range(2))
    blocks.append(f"<w:tbl>{rows}</w:tbl>")
    while len(blocks) < PARAGRAPHS_PER_PAGE:
        blocks.append(paragraph(run(text)))
    return "".join(blocks)


def write_docx(path: Path, body: str, comments: int = 0) -> None:
    relationships = (f'<Relationship Id="styles" Type="{R}/styles" Target="styles.xml"/>'
                     f'<Relationship Id="numbering" Type="{R}/numbering" Target="numbering.xml"/>'
                     f'<Relationship Id="comments" Type="{R}/comments" Target="comments.xml"/>')
    parts = {
        "[Content_Types].xml": (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
            '</Types>'
        ),
        "_rels/.rels": (f'<Relationships xmlns="{PKG_R}"><Relationship Id="rId1" '
                        f'Type="{R}/officeDocument" Target="word/document.xml"/></Relationships>'),
        "word/document.xml": ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                              f'<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>{body}</w:body></w:document>'),
        "word/_rels/document.xml.rels": f'<Relationships xmlns="{PKG_R}">{relationships}</Relationships>',
        "word/styles.xml": (f'<w:styles xmlns:w="{W}"><w:style w:type="paragraph" w:styleId="Heading1">'
                            '<w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style></w:styles>'),
        "word/numbering.xml": (f'<w:numbering xmlns:w="{W}"><w:abstractNum w:abstractNumId="0">'
                               '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>'
                               '<w:lvl w:ilvl="1"><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl>'
                               '</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>'),
        "word/comments.xml": (f'<w:comments xmlns:w="{W}">' + "".join(
            f'<w:comment w:id="{i}" w:author="Ada">{paragraph(run(f"Comment {i}. Please check e.g. the figures."))}</w:comment>'
            for i in range(comments)) + "</w:comments>"),
    }
    with ZipFile(path, "w", ZIP_DEFLATED) as archive:
        for name, content in sorted(parts.items()):
            info = ZipInfo(name, date_time=(2020, 1, 2, 3, 4, 6))
            info.compress_type = ZIP_DEFLATED
            archive.writestr(info, content.encode("utf-8"))


def timed(label: str, path: Path) -> None:
    started = time.perf_counter()
    result = extract(path)
    elapsed = time.perf_counter() - started
    lines = result.prose.count("\n")
    print(f"{label}: {elapsed:.2f} s ({path.stat().st_size / 1024:.0f} KiB DOCX, {lines} prose lines)")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    parser.add_argument("--pages", type=int, default=500, help="pages in the large document (default 500)")
    parser.add_argument("--sentences", type=int, default=4000,
                        help="sentences in the single long paragraph (default 4000)")
    parser.add_argument("--keep", type=Path, help="write the DOCX files to this directory and keep them")
    args = parser.parse_args()

    with tempfile.TemporaryDirectory() as temporary:
        directory = args.keep or Path(temporary)
        directory.mkdir(parents=True, exist_ok=True)
        large = directory / f"large-{args.pages}-pages.docx"
        write_docx(large, "".join(page(number) for number in range(args.pages)), comments=args.pages)
        long = directory / f"long-paragraph-{args.sentences}-sentences.docx"
        text = " ".join(SENTENCE.format(n=n, m=n + 1) for n in range((args.sentences + 2) // 3))
        write_docx(long, paragraph(run(text)))

        timed(f"{args.pages}-page document ({args.pages * PARAGRAPHS_PER_PAGE} body blocks)", large)
        timed(f"one paragraph of about {args.sentences} sentences ({len(text)} characters)", long)


if __name__ == "__main__":
    main()
