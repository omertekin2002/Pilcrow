"""Package safety limits: what is rejected, and image-heavy documents that are not.

Large parts are built from zeros or repeated text and compressed, so fixtures stay small
on disk. Run with ``python -m unittest discover -s tests -v`` from the project root.
"""

from __future__ import annotations

import json
import struct
import tempfile
import tracemalloc
import unittest
from pathlib import Path
from zipfile import ZIP_BZIP2, ZIP_DEFLATED, ZIP_LZMA, ZipFile, ZipInfo

try:
    from . import test_cli as cli_tests
    from .test_cli import (
        PKG_R,
        R,
        W,
        anchored,
        image_entry,
        image_rel,
        make_docx,
        paragraph,
        picture,
    )
except ImportError:  # discovered with -s tests: the directory itself is on sys.path
    import test_cli as cli_tests  # type: ignore[no-redef]
    from test_cli import (  # type: ignore[no-redef]
        PKG_R,
        R,
        W,
        anchored,
        image_entry,
        image_rel,
        make_docx,
        paragraph,
        picture,
    )

MIB = 1024 * 1024
GIB = 1024 * MIB
HEADER_REL = '<Relationship Id="{rid}" Type="{R}/header" Target="{target}"/>'


def header(text: str) -> str:
    return f'<w:hdr xmlns:w="{W}" xmlns:r="{R}">{paragraph(text)}</w:hdr>'


def declare_sizes(path: Path, sizes: dict[str, int]) -> None:
    """Overwrite the uncompressed sizes the central directory declares for some parts.

    Only the declared size changes, as in a hostile package; the stored data does not.
    """
    data = bytearray(path.read_bytes())
    with ZipFile(path) as archive:
        position = archive.start_dir  # type: ignore[attr-defined]
    found = set()
    while data[position:position + 4] == b"PK\x01\x02":
        name_length, extra_length, comment_length = struct.unpack_from("<HHH", data, position + 28)
        name = data[position + 46:position + 46 + name_length].decode("utf-8")
        if name in sizes:
            struct.pack_into("<I", data, position + 24, sizes[name])
            found.add(name)
        position += 46 + name_length + extra_length + comment_length
    assert found == set(sizes), (found, sizes)
    path.write_bytes(bytes(data))


class LimitTestCase(unittest.TestCase):
    """A temporary root and the CLI helpers of CLITests, without inheriting (or rerunning) its tests."""

    cli = cli_tests.CLITests.cli
    companions = cli_tests.CLITests.companions
    exported = cli_tests.CLITests.exported

    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(prefix="pilcrow-test-")
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)


class ImageHeavyDocumentTests(LimitTestCase):
    """The limits bound what Pilcrow parses into memory, not the media it only hashes."""

    def images_docx(self, sizes: list[int]) -> list[str]:
        """essay.docx with one referenced image of zeros per size; returns each image's data."""
        images = ["\0" * size for size in sizes]
        body = paragraph("Caption.") + "".join(anchored(f"Figure {n}:", picture(f"img{n}")) for n in range(len(images)))
        make_docx(self.root / "essay.docx", body,
                  extra={f"word/media/image{n}.png": data for n, data in enumerate(images)},
                  relationships="".join(image_rel(f"img{n}", f"media/image{n}.png") for n in range(len(images))))
        return images

    def test_an_image_over_16_mib_exports_and_checks_fresh(self) -> None:
        (image,) = self.images_docx([20 * MIB])
        self.cli("export", "essay.docx")
        meta = json.loads(self.exported()[1])
        self.assertEqual(meta["images"], [image_entry(image)])
        self.assertEqual(meta["images"][0]["bytes"], 20 * MIB)
        self.cli("check", "essay.docx")

    def test_images_totalling_over_64_mib_export_and_check_fresh(self) -> None:
        images = self.images_docx([25 * MIB] * 3)
        self.cli("export", "essay.docx")
        meta = json.loads(self.exported()[1])
        self.assertEqual(meta["images"], [image_entry(data) for data in images])
        self.cli("check", "essay.docx")

    def test_hashing_a_large_image_does_not_hold_it_in_memory(self) -> None:
        from pilcrow.extract import extract
        (image,) = self.images_docx([20 * MIB])
        del image
        tracemalloc.start()
        try:
            result = extract(self.root / "essay.docx")
            peak = tracemalloc.get_traced_memory()[1]
        finally:
            tracemalloc.stop()
        self.assertEqual(result.metadata["images"][0]["bytes"], 20 * MIB)
        self.assertLess(peak, 8 * MIB, f"peak traced memory {peak} bytes while hashing a 20 MiB image")


class SafetyLimitTests(LimitTestCase):
    """Each limit still rejects its package with exit 2 and a message naming what exceeded it."""

    def rejected(self, *parts: str) -> str:
        result = self.cli("export", "essay.docx", code=2)
        self.assertNotIn("Traceback", result.stderr)
        for part in parts:
            self.assertIn(part, result.stderr)
        self.assertFalse(self.companions()[0].exists())
        self.cli("check", "essay.docx", code=2)
        return result.stderr

    def test_a_parsed_part_over_16_mib_is_rejected(self) -> None:
        make_docx(self.root / "essay.docx", paragraph("a" * (17 * MIB)))
        self.rejected("error: essay.docx: word/document.xml is ", "16 MiB")

    def test_a_diagram_data_part_over_16_mib_is_rejected(self) -> None:
        # A SmartArt data part is parsed for its text, so the parsed-part cap applies to it.
        make_docx(self.root / "essay.docx", anchored("Steps:", cli_tests.smartart("dgm")),
                  extra={"word/diagrams/data1.xml": "<a/>" + " " * (17 * MIB)},
                  relationships=cli_tests.rel("dgm", "diagramData", "diagrams/data1.xml"))
        self.rejected("word/diagrams/data1.xml is ", "16 MiB")

    def test_a_relationships_part_over_16_mib_is_rejected(self) -> None:
        make_docx(self.root / "essay.docx", paragraph("Prose."), relationships=" " * (17 * MIB))
        self.rejected("word/_rels/document.xml.rels is ", "16 MiB")

    def test_a_part_inflating_past_its_declared_size_is_rejected_without_inflating_it(self) -> None:
        from pilcrow.extract import DocumentError, extract
        make_docx(self.root / "essay.docx", paragraph("a" * (16 * MIB)))
        # A hostile header declares 64 KiB for a stream that inflates to 16 MiB.
        declare_sizes(self.root / "essay.docx", {"word/document.xml": 64 * 1024})
        tracemalloc.start()
        try:
            with self.assertRaisesRegex(DocumentError, "cannot parse word/document.xml: Bad CRC-32"):
                extract(self.root / "essay.docx")
            peak = tracemalloc.get_traced_memory()[1]
        finally:
            tracemalloc.stop()
        self.assertLess(peak, 8 * MIB, f"peak traced memory {peak} bytes reading a part declared as 64 KiB")
        self.rejected("cannot parse word/document.xml")

    def test_parsed_parts_over_64_mib_in_total_are_rejected(self) -> None:
        headers = {f"word/header{n}.xml": header(f"Header {n}.") for n in range(5)}
        rels = "".join(HEADER_REL.format(rid=f"h{n}", R=R, target=f"header{n}.xml") for n in range(5))
        make_docx(self.root / "essay.docx", paragraph("Prose."), extra=headers, relationships=rels)
        # Each header declares 15 MiB: under the per-part cap, 75 MiB together.
        declare_sizes(self.root / "essay.docx", dict.fromkeys(headers, 15 * MIB))
        self.rejected("cannot parse word/header", "64 MiB total")

    def test_a_story_parts_relationships_count_once_however_many_notes_it_has(self) -> None:
        notes = "".join(f'<w:footnote w:id="{n}">{paragraph(f"Note {n}.")}</w:footnote>' for n in range(1, 1101))
        # 64 KiB of relationships, which 1100 separate parses would charge as 70 MiB.
        note_rels = f'<Relationships xmlns="{PKG_R}">{image_rel("unused", "media/none.png")}{" " * 64 * 1024}</Relationships>'
        make_docx(self.root / "essay.docx", paragraph("Prose."),
                  extra={"word/footnotes.xml": f'<w:footnotes xmlns:w="{W}">{notes}</w:footnotes>',
                         "word/_rels/footnotes.xml.rels": note_rels},
                  relationships=f'<Relationship Id="footnotes" Type="{R}/footnotes" Target="footnotes.xml"/>')
        self.cli("export", "essay.docx")
        self.assertIn(b"Note 1100.", self.exported()[0])

    def test_a_package_over_1_gib_uncompressed_is_rejected(self) -> None:
        make_docx(self.root / "essay.docx", anchored("Figure:", picture("rId5")),
                  extra={"word/media/image1.png": "picture"}, relationships=image_rel("rId5", "media/image1.png"))
        declare_sizes(self.root / "essay.docx", {"word/media/image1.png": GIB})
        self.rejected("package is ", "1 GiB")

    def test_a_package_with_more_than_4096_entries_is_rejected(self) -> None:
        make_docx(self.root / "essay.docx", paragraph("Prose."))
        with ZipFile(self.root / "essay.docx", "a", ZIP_DEFLATED) as archive:
            for n in range(4096 - len(archive.namelist()) + 1):
                archive.writestr(f"customXml/item{n}.xml", "")
        self.rejected("package has 4097 entries", "4096")

    def test_compression_methods_without_bounded_output_are_rejected(self) -> None:
        for method, module in ((ZIP_BZIP2, "bz2"), (ZIP_LZMA, "lzma")):
            with self.subTest(module=module):
                try:
                    __import__(module)
                except ImportError:
                    self.skipTest(f"{module} unavailable in this Python")
                make_docx(self.root / "essay.docx", anchored("Figure:", picture("rId5")),
                          relationships=image_rel("rId5", "media/image1.png"))
                with ZipFile(self.root / "essay.docx", "a") as archive:
                    info = ZipInfo("word/media/image1.png", date_time=(2020, 1, 2, 3, 4, 6))
                    info.compress_type = method
                    archive.writestr(info, b"\0" * MIB)
                self.rejected(f"word/media/image1.png uses ZIP compression method {method}", "stored or deflate")

    def test_packages_at_the_limits_are_accepted(self) -> None:
        make_docx(self.root / "essay.docx", paragraph("Prose."), extra={"word/media/image1.png": "unused"})
        with ZipFile(self.root / "essay.docx", "a", ZIP_DEFLATED) as archive:
            for n in range(4096 - len(archive.namelist())):
                archive.writestr(f"customXml/item{n}.xml", "")
        # The total includes an unreferenced media part that declares the rest of 1 GiB.
        with ZipFile(self.root / "essay.docx") as archive:
            rest = GIB - sum(i.file_size for i in archive.infolist() if i.filename != "word/media/image1.png")
        declare_sizes(self.root / "essay.docx", {"word/media/image1.png": rest})
        self.cli("export", "essay.docx")


if __name__ == "__main__":
    unittest.main()
