"""Default artifacts stay together even when the CLI starts outside the checkout."""

import io
import json
import unittest
from contextlib import chdir, redirect_stderr, redirect_stdout
from pathlib import Path
from unittest.mock import patch
from zipfile import ZipFile

from support import temporary_directory
from superslide import cli
from superslide.llm import GenerationResult, LLMError
from superslide.paths import OUTPUT_DIR, PROJECT_ROOT


class OutputPathTests(unittest.TestCase):
    def setUp(self):
        temporary = temporary_directory()
        self.addCleanup(temporary.cleanup)
        self.root = Path(temporary.name)
        self.working = self.root / "caller"
        self.working.mkdir()
        self.output = self.root / "output"
        self.source = self.working / "facts.txt"
        self.source.write_text("Выручка составила 125 млн рублей.", encoding="utf-8")
        self.spec = {"template_version": "2.0", "title": "Отчёт", "slides": [
            {"layout_id": "title", "variant_id": "slide_002", "updates": []},
        ]}

    def invoke(self, arguments):
        stdout, stderr = io.StringIO(), io.StringIO()
        with chdir(self.working), patch.object(cli, "OUTPUT_DIR", self.output), \
                redirect_stdout(stdout), redirect_stderr(stderr):
            result = cli.main(arguments)
        return result, stdout.getvalue(), stderr.getvalue()

    def test_default_output_is_anchored_to_checkout(self):
        with chdir(self.working):
            self.assertEqual(OUTPUT_DIR, PROJECT_ROOT / "output")
            self.assertTrue(OUTPUT_DIR.is_absolute())

    def test_build_defaults_to_input_name_and_report_in_output(self):
        (self.working / "quarter.json").write_text(json.dumps(self.spec), encoding="utf-8")
        result, _, errors = self.invoke(["build", "quarter.json"])
        self.assertEqual(result, 0, errors)
        with ZipFile(self.output / "quarter.pptx") as archive:
            self.assertIn("ppt/presentation.xml", archive.namelist())
        report = json.loads((self.output / "quarter.report.json").read_text(encoding="utf-8"))
        self.assertEqual(report["slide_count"], 1)
        self.assertFalse((self.working / "quarter.pptx").exists())

    def test_extract_defaults_to_output_but_explicit_path_remains_relative_to_caller(self):
        result, _, errors = self.invoke(["extract", "facts.txt"])
        self.assertEqual(result, 0, errors)
        default = self.output / "sources.json"
        self.assertTrue(json.loads(default.read_text(encoding="utf-8"))["chunks"])
        result, _, errors = self.invoke(["extract", "facts.txt", "--output", "custom/sources.json"])
        self.assertEqual(result, 0, errors)
        explicit = self.working / "custom" / "sources.json"
        self.assertEqual(json.loads(default.read_text(encoding="utf-8")),
                         json.loads(explicit.read_text(encoding="utf-8")))

    def draft_arguments(self):
        config = self.working / "llm.json"
        config.write_text(json.dumps({"base_url": "https://mock.invalid/v1", "model": "test"}),
                          encoding="utf-8")
        return ["draft", "--task", "Отчёт", "--config", "llm.json", "--documents", "facts.txt"]

    def test_draft_and_sidecar_files_default_to_output(self):
        with patch.object(cli, "generate_draft", return_value=GenerationResult(self.spec, 1)):
            result, _, errors = self.invoke(self.draft_arguments())
        self.assertEqual(result, 0, errors)
        self.assertEqual({path.name for path in self.output.iterdir()},
                         {"draft.json", "draft.sources.json", "draft.generation.json"})
        self.assertEqual(json.loads((self.output / "draft.json").read_text(encoding="utf-8")), self.spec)

    def test_failed_draft_diagnostics_default_to_output(self):
        with patch.object(cli, "generate_draft", side_effect=LLMError("Test failure", code="test")):
            result, _, _ = self.invoke(self.draft_arguments())
        self.assertEqual(result, 2)
        self.assertEqual({path.name for path in self.output.iterdir()},
                         {"draft.sources.json", "draft.error.json"})


if __name__ == "__main__":
    unittest.main()
