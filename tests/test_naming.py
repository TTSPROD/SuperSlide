"""Presentation downloads retain their generated titles as portable names."""

import unittest
import unicodedata

from superslide.naming import presentation_filename


class PresentationFilenameTests(unittest.TestCase):
    def test_descriptive_unicode_title_keeps_words_and_spaces(self):
        self.assertEqual(presentation_filename("Итоги проекта Северсталь 2026"),
                         "Итоги проекта Северсталь 2026.pptx")
        self.assertEqual(presentation_filename("Résumé — план запуска"),
                         "Résumé — план запуска.pptx")

    def test_existing_extension_is_not_duplicated(self):
        for title in ("Итоги.pptx", "Итоги.PPTX", "Итоги.pptx.pptx", "  Итоги.PPTX.  "):
            with self.subTest(title=title):
                self.assertEqual(presentation_filename(title), "Итоги.pptx")

    def test_path_and_header_characters_are_removed_from_a_single_filename(self):
        result = presentation_filename('  ../Итоги\\проект: "Склад" <2D>|3D?*\r\nX-Test: yes  ')
        self.assertEqual(result, "Итоги проект Склад 2D 3D X-Test yes.pptx")
        self.assertFalse(any(character in result for character in '/\\:*?"<>|\r\n'))
        self.assertFalse(any(unicodedata.category(character) in {"Cc", "Cf", "Cs"}
                             for character in result))

    def test_invisible_bidi_controls_cannot_disguise_the_name(self):
        self.assertEqual(presentation_filename("Отч\u202eёт\u2066 за\u200b год"), "Отчёт за год.pptx")
        self.assertEqual(presentation_filename("Итоги\ud800 проекта"), "Итоги проекта.pptx")

    def test_windows_reserved_stems_are_safe_including_stems_with_dots(self):
        for title in ("CON", "aux", "NUL.notes", "COM1", "LPT9", "PRN .notes",
                      "CON.pptx", "com¹", "CONOUT$"):
            with self.subTest(title=title):
                self.assertTrue(presentation_filename(title).startswith("_"))
        self.assertEqual(presentation_filename("CON отчет"), "CON отчет.pptx")

    def test_long_titles_fit_character_and_utf8_limits(self):
        for title in ("a" * 500, "Производство " * 100, "🚀" * 200, "CON." + "я" * 500):
            with self.subTest(title=title[:20]):
                filename = presentation_filename(title)
                self.assertTrue(filename.endswith(".pptx"))
                stem = filename[:-5]
                self.assertTrue(stem)
                self.assertLessEqual(len(stem), 96)
                self.assertLessEqual(len(stem.encode("utf-8")), 180)
                self.assertFalse(stem.endswith((" ", ".")))

    def test_blank_or_entirely_invalid_titles_use_a_readable_fallback(self):
        for title in ("", " \t\r\n ", '../\\:*?"<>|...', "\u202e\u2066", ".PPTX"):
            with self.subTest(title=repr(title)):
                self.assertEqual(presentation_filename(title), "Презентация.pptx")


if __name__ == "__main__":
    unittest.main()
