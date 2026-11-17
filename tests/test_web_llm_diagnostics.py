"""Failures retain field-level evidence, while credentials/responses stay private."""
import json
import os
import threading
import time
import unittest
from pathlib import Path
from unittest.mock import patch

from support import temporary_directory
from superslide.llm import LLMError
from superslide.web import ROOT, WebState, _safe_failure, _failure_message


class WebDiagnosticsTests(unittest.TestCase):
    def test_unfilled_content_group_is_reported_without_source_text(self):
        exc = LLMError("private model text", code="validation", diagnostics=[{
            "validation_issues": [{"slide_index": 2, "reason": "unfilled_group", "limit": 4, "actual": 2}],
        }])
        details = _safe_failure(exc, {"stage": "slides", "slide_start": 2})
        message = _failure_message(exc, details)
        self.assertIn("не заполнены все блоки", message)
        self.assertIn("Слайд 2", message)
        self.assertNotIn("private model text", message)

    def setUp(self):
        self.temporary = temporary_directory()
        self.addCleanup(self.temporary.cleanup)
        self.config = Path(self.temporary.name) / "config.json"
        self.config.write_text(json.dumps({"base_url": "https://private.invalid/v1", "model": "test", "api_key_env": "SUPERSLIDE_PROBE_TEST_KEY", "timeout_seconds": 240}), encoding="utf-8")
        self.state = WebState(self.config)
        self.state.directory = Path(self.temporary.name) / "jobs"
        self.addCleanup(lambda: self.state.executor.shutdown(wait=True, cancel_futures=True))
        self.addCleanup(lambda: self.state.probe_executor.shutdown(wait=True, cancel_futures=True))

    def poll(self, job_id):
        for _ in range(100):
            job = self.state.public_job(job_id)
            if job["status"] in ("succeeded", "failed"):
                return job
            time.sleep(0.02)
        self.fail("Job did not finish")

    def test_timeout_retains_stage_and_partial_slides_without_raw_diagnostics(self):
        partial = {"template_version": "2.0", "title": "Готовая часть", "slides": [{"layout_id": "title", "variant_id": "slide_003", "updates": []}]}

        def work(job_id):
            self.state.progress(job_id, {"stage": "slides", "slide_start": 2, "total_slides": 3, "timeout_seconds": 240, "raw": "SECRET"})
            exc = LLMError("SECRET", code="timeout", diagnostics=[{"response_excerpt": "SECRET"}])
            exc.partial_draft = partial
            raise exc

        job = self.poll(self.state.submit("draft", work)["job_id"])
        self.assertIn("240 с", job["error"])
        self.assertIn("Слайд 2", job["error"])
        self.assertIn("Сохранены готовые слайды: 1", job["error"])
        self.assertNotIn("SECRET", json.dumps(job))
        saved = self.state.directory / "jobs" / job["id"]
        self.assertTrue((saved / "partial-draft.json").is_file())
        self.assertNotIn("SECRET", (saved / "diagnostics.json").read_text(encoding="utf-8"))

    def test_structure_issue_publishes_only_location_and_reason(self):
        error = LLMError("SECRET", code="validation", diagnostics=[{
            "response_excerpt": "SECRET", "validation_details": [{"slide_index": 2, "field_id": "slide_text_13", "reason": "text_length", "limit": 50, "actual": 60, "input": "SECRET"}],
        }])
        details = _safe_failure(error, {"stage": "repair", "slide_start": 2})
        message = _failure_message(error, details)
        self.assertIn("slide_text_13", message)
        self.assertIn("текст длиннее лимита", message)
        self.assertNotIn("SECRET", json.dumps(details))

    def test_connection_probe_does_not_wait_for_presentation_queue(self):
        block = threading.Event()
        started = threading.Event()

        def busy(job_id):
            started.set()
            block.wait(5)
            return {}

        self.state.submit("draft", busy)
        started.wait(1)
        try:
            with patch.dict(os.environ, {"SUPERSLIDE_PROBE_TEST_KEY": "test-secret"}), patch("superslide.llm._request_content", return_value="ОК") as request:
                job = self.poll(self.state.check_connection()["job_id"])
                self.assertTrue(job["connection_ok"])
                payload, settings, key = request.call_args.args
                self.assertEqual(key, "test-secret")
                self.assertLessEqual(settings.timeout_seconds, 60)
                self.assertEqual(payload["messages"][0]["content"], "Проверка связи. Ответь одним словом: ОК")
                self.assertNotIn("test-secret", json.dumps(job))
        finally:
            block.set()

    def test_schema_feedback_keeps_field_path_without_response_text(self):
        error = LLMError("SECRET", code="validation", diagnostics=[{
            "validation_details": {"reason": "schema", "fields": [{"location": ["slides", 1, "updates", 0, "text"], "type": "string_too_long", "input": "SECRET"}]},
            "response_excerpt": "SECRET",
        }])
        details = _safe_failure(error, {"stage": "slides", "slide_start": 2})
        self.assertEqual(details["issues"][0]["slide_index"], 2)
        self.assertIn("slides.1.updates.0.text", _failure_message(error, details))
        self.assertNotIn("SECRET", json.dumps(details))

    def test_last_attempt_only_and_batch_range_are_shown(self):
        error = LLMError("failed", code="validation", diagnostics=[
            {"validation_details": {"slide_index": 2, "field_id": "slide_text_22", "reason": "text_length", "limit": 224, "actual": 284}},
            {"validation_issues": [{"slide_index": 2, "field_id": "slide_text_27", "reason": "text_length", "limit": 116, "actual": 146}]},
        ])
        details = _safe_failure(error, {"stage": "slides", "slide_start": 1, "slide_end": 2})
        message = _failure_message(error, details)
        self.assertIn("Слайды 1–2", message)
        self.assertNotIn("slide_text_22", message)
        self.assertIn("slide_text_27", message)


if __name__ == "__main__":
    unittest.main()
