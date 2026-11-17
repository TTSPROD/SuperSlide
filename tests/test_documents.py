import unittest
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from support import temporary_directory
from superslide.documents import DocumentError, load_documents


class DocumentTests(unittest.TestCase):
    def setUp(self):
        self.directory = temporary_directory()
        self.addCleanup(self.directory.cleanup)
        self.root = Path(self.directory.name)

    def test_utf8_bom_and_stable_source_ids(self):
        first = self.root / "тезисы.md"
        first.write_text("Рост выручки на 12%.", encoding="utf-8-sig")
        second = self.root / "other.txt"
        second.write_text("Second source", encoding="utf-8")
        bundle = load_documents([first, second])
        self.assertEqual(bundle.chunks[0].id, "doc1:text")
        self.assertEqual(bundle.chunks[1].id, "doc2:text")
        self.assertEqual(bundle.chunks[0].text, "Рост выручки на 12%.")
        self.assertIn("[doc1:text]", bundle.context_text())
        self.assertEqual(bundle.to_dict()["chunks"][0]["document"], first.name)

    def test_limits_do_not_silently_truncate(self):
        path = self.root / "a.txt"
        path.write_text("1234567890", encoding="utf-8")
        with self.assertRaisesRegex(DocumentError, "символов"):
            load_documents([path], max_file_chars=9)
        with self.assertRaisesRegex(DocumentError, "байт"):
            load_documents([path], max_file_bytes=9)
        other = self.root / "b.txt"
        other.write_text("abcdefghij", encoding="utf-8")
        with self.assertRaisesRegex(DocumentError, "Общий объём"):
            load_documents([path, other], max_total_chars=19)
        with self.assertRaisesRegex(DocumentError, "не более"):
            load_documents([path, other], max_files=1)

    def test_missing_duplicate_unsupported_and_encoding(self):
        with self.assertRaisesRegex(DocumentError, "не найден"):
            load_documents([self.root / "missing.txt"])
        path = self.root / "a.txt"
        path.write_text("text", encoding="utf-8")
        with self.assertRaisesRegex(DocumentError, "повторно"):
            load_documents([path, path])
        path = self.root / "a.csv"
        path.write_text("a,b", encoding="utf-8")
        with self.assertRaisesRegex(DocumentError, "Неподдерживаемый"):
            load_documents([path])
        path = self.root / "a.txt"
        path.write_bytes(b"\xff\xfe\x80\x80")
        with self.assertRaisesRegex(DocumentError, "UTF-8"):
            load_documents([path])

    def test_docx_paragraphs_and_tables_preserve_order(self):
        from docx import Document
        document = Document()
        document.add_paragraph("Before table")
        table = document.add_table(rows=2, cols=2)
        table.cell(0, 0).text = "Year"
        table.cell(0, 1).text = "Revenue"
        table.cell(1, 0).text = "2025"
        table.cell(1, 1).text = "125"
        document.add_paragraph("After table")
        path = self.root / "report.docx"
        document.save(path)
        bundle = load_documents([path])
        self.assertEqual([chunk.id for chunk in bundle.chunks], [
            "doc1:paragraph1", "doc1:table1:row1", "doc1:table1:row2", "doc1:paragraph2",
        ])
        self.assertEqual(bundle.chunks[2].text, "2025 | 125")
        self.assertTrue(bundle.warnings)

    def test_compressed_docx_expansion_is_bounded(self):
        path = self.root / "large.docx"
        with ZipFile(path, "w", ZIP_DEFLATED) as archive:
            archive.writestr("word/document.xml", "a" * 20_000)
        with self.assertRaisesRegex(DocumentError, "Распакованный"):
            load_documents([path], max_file_bytes=1000)

    def test_pdf_text_and_missing_text_page_diagnostic(self):
        from pypdf import PdfWriter
        from pypdf.generic import DecodedStreamObject, DictionaryObject, NameObject
        writer = PdfWriter()
        page = writer.add_blank_page(width=300, height=300)
        font = DictionaryObject({
            NameObject("/Type"): NameObject("/Font"),
            NameObject("/Subtype"): NameObject("/Type1"),
            NameObject("/BaseFont"): NameObject("/Helvetica"),
        })
        page[NameObject("/Resources")] = DictionaryObject({
            NameObject("/Font"): DictionaryObject({NameObject("/F1"): writer._add_object(font)}),
        })
        stream = DecodedStreamObject()
        stream.set_data(b"BT /F1 12 Tf 20 250 Td (Revenue 125) Tj ET")
        page[NameObject("/Contents")] = writer._add_object(stream)
        writer.add_blank_page(width=300, height=300)
        path = self.root / "report.pdf"
        with path.open("wb") as handle:
            writer.write(handle)
        bundle = load_documents([path])
        self.assertEqual(bundle.chunks[0].id, "doc1:p1")
        self.assertIn("Revenue 125", bundle.chunks[0].text)
        self.assertTrue(any("страница 2" in warning for warning in bundle.warnings))

    def test_scan_only_pdf_requires_local_ocr(self):
        from pypdf import PdfWriter
        writer = PdfWriter()
        writer.add_blank_page(width=300, height=300)
        path = self.root / "scan.pdf"
        with path.open("wb") as handle:
            writer.write(handle)
        with self.assertRaisesRegex(DocumentError, "OCR"):
            load_documents([path])


if __name__ == "__main__":
    unittest.main()
