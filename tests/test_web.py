"""Exercise the real local HTTP API without contacting a language model."""

import copy
import http.client
import json
import os
import tempfile
import threading
import time
import unittest
import uuid
from pathlib import Path
from unittest.mock import patch
from urllib.parse import quote

from pptx import Presentation

from superslide.llm import GenerationResult, LLMError
from superslide.web import MAX_JSON, MAX_UPLOAD, ROOT, create_server


class WebTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.directory = tempfile.TemporaryDirectory(dir=ROOT)
        cls.addClassCleanup(cls.directory.cleanup)
        cls.config = Path(cls.directory.name) / "llm.json"
        cls.config.write_text(json.dumps({"base_url": "https://internal.invalid/v1", "model": "test-model",
                                          "api_key_env": "SUPERSLIDE_WEB_TEST_KEY"}), encoding="utf-8")
        cls.server = create_server(port=0, config_path=cls.config)
        cls.server.state.directory = Path(cls.directory.name) / "web"
        cls.thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.thread.start()
        cls.port = cls.server.server_address[1]
        cls.example = json.loads((ROOT / "examples/variety.json").read_text(encoding="utf-8"))

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()
        cls.thread.join(timeout=2)

    def setUp(self):
        self.environment = patch.dict(os.environ, {"SUPERSLIDE_WEB_TEST_KEY": "test-local-secret"})
        self.environment.start()
        self.addCleanup(self.environment.stop)

    def request(self, method, path, body=None, headers=None, *, json_body=False):
        request_headers = {"X-SuperSlide-Request": "1"}
        if json_body:
            body = json.dumps(body, ensure_ascii=False).encode("utf-8")
            request_headers["Content-Type"] = "application/json"
        if headers:
            request_headers.update(headers)
        connection = http.client.HTTPConnection("127.0.0.1", self.port, timeout=5)
        try:
            connection.request(method, path, body, request_headers)
            response = connection.getresponse()
            raw = response.read()
            data = json.loads(raw) if "application/json" in response.getheader("Content-Type", "") else raw
            return response.status, data, dict(response.getheaders())
        finally:
            connection.close()

    def poll(self, job_id):
        deadline = time.monotonic() + 10
        while time.monotonic() < deadline:
            status, job, _ = self.request("GET", f"/api/jobs/{job_id}")
            self.assertEqual(status, 200)
            if job["status"] in {"succeeded", "failed"}:
                return job
            time.sleep(0.025)
        self.fail("Local job did not finish")

    def upload(self, filename="Заметки.txt", content=b"Local source"):
        return self.request("POST", "/api/uploads?filename=" + quote(filename), content,
                            {"Content-Type": "application/octet-stream"})

    def test_bootstrap_exposes_catalog_without_configuration_secrets(self):
        status, data, headers = self.request("GET", "/api/bootstrap")
        self.assertEqual(status, 200)
        self.assertTrue(data["llm_configured"])
        self.assertTrue(data["key_configured"])
        self.assertEqual(data["model"], "test-model")
        self.assertEqual(len(data["layouts"]), 14)
        self.assertEqual(data["design_count"], 31)
        raw = json.dumps(data)
        self.assertNotIn("internal.invalid", raw)
        self.assertNotIn("test-local-secret", raw)
        self.assertNotIn("Access-Control-Allow-Origin", headers)
        self.assertIn("frame-ancestors 'none'", headers["Content-Security-Policy"])

    def test_example_builds_native_editable_pptx_and_allowlisted_downloads(self):
        status, example, _ = self.request("GET", "/api/example")
        self.assertEqual(status, 200)
        self.assertEqual(len(example["draft"]["slides"]), 14)
        status, created, _ = self.request("POST", "/api/build", {"draft": example["draft"]}, json_body=True)
        self.assertEqual(status, 202)
        job = self.poll(created["job_id"])
        self.assertEqual(job["status"], "succeeded", job)
        status, raw, headers = self.request("GET", job["downloads"]["pptx"])
        self.assertEqual(status, 200)
        self.assertTrue(raw.startswith(b"PK"))
        self.assertIn('filename="presentation.pptx"', headers["Content-Disposition"])
        path = Path(self.directory.name) / "download.pptx"
        path.write_bytes(raw)
        deck = Presentation(path)
        self.assertEqual(len(deck.slides), 14)
        self.assertEqual(sum(shape.has_chart for slide in deck.slides for shape in slide.shapes), 1)
        self.assertEqual(sum(shape.has_table for slide in deck.slides for shape in slide.shapes), 1)
        status, report, _ = self.request("GET", job["downloads"]["report"])
        self.assertEqual(status, 200)
        self.assertEqual(report["output"], "presentation.pptx")
        for name in ("sources.json", "../llm.json", "presentation.pptx/../sources.json"):
            status, _, _ = self.request("GET", f"/api/downloads/{job['id']}/{name}")
            self.assertEqual(status, 404)

    def test_draft_extracts_uploaded_documents_and_preserves_source_job(self):
        status, upload, _ = self.upload(content="Проверенный локальный факт".encode("utf-8"))
        self.assertEqual(status, 201)
        payload = copy.deepcopy(self.example)
        payload["slides"] = payload["slides"][:3]

        def generate(task, bundle, settings, validate, **kwargs):
            self.assertIn("ровно 3 слайда", task)
            self.assertIn("Проверенный локальный факт", bundle.context_text())
            self.assertEqual(settings.model, "test-model")
            self.assertEqual(kwargs["template_version"], "1.2")
            self.assertTrue(kwargs["image_catalog"])
            payload["slides"][2]["source_refs"] = [bundle.chunks[0].id]
            validate(payload)
            return GenerationResult(payload, 1)

        with patch("superslide.web.generate_draft", side_effect=generate) as generate_mock:
            status, created, _ = self.request("POST", "/api/drafts", {
                "task": "Тестовая презентация", "document_ids": [upload["id"]], "slide_count": 3}, json_body=True)
            self.assertEqual(status, 202)
            job = self.poll(created["job_id"])
            self.assertEqual(job["status"], "succeeded", job)
            self.assertEqual(generate_mock.call_count, 1)
        self.assertNotIn("_bundle", job)
        self.assertNotIn("_path", json.dumps(job))
        status, created, _ = self.request("POST", "/api/build", {
            "draft": job["draft"], "source_job_id": job["id"]}, json_body=True)
        self.assertEqual(status, 202)
        built = self.poll(created["job_id"])
        self.assertEqual(built["status"], "succeeded", built)
        self.assertEqual(self.request("DELETE", f"/api/uploads/{upload['id']}")[0], 200)

    def test_import_normalizes_draft_and_round_trips_into_build(self):
        status, example, _ = self.request("GET", "/api/example")
        self.assertEqual(status, 200)
        draft = copy.deepcopy(example["draft"])
        draft["slides"] = draft["slides"][:4]
        draft["slides"][0]["variant_id"] = "blue_minimal"
        status, imported, _ = self.request("POST", "/api/import", {"draft": draft}, json_body=True)
        self.assertEqual(status, 200)
        self.assertEqual(len(imported["draft"]["slides"]), 4)
        self.assertEqual(imported["draft"]["slides"][0]["variant_id"], "blue_minimal")
        self.assertEqual(imported["draft"]["template_version"], "1.2")
        status, created, _ = self.request("POST", "/api/build", {"draft": imported["draft"]}, json_body=True)
        self.assertEqual(status, 202)
        self.assertEqual(self.poll(created["job_id"])["status"], "succeeded")

    def test_import_rejects_invalid_drafts_extra_fields_and_foreign_requests(self):
        bad = copy.deepcopy(self.example)
        bad["slides"][0]["variant_id"] = "invented"
        self.assertEqual(self.request("POST", "/api/import", {"draft": bad}, json_body=True)[0], 400)
        self.assertEqual(self.request("POST", "/api/import", {"draft": "not-an-object"}, json_body=True)[0], 400)
        self.assertEqual(self.request("POST", "/api/import", {}, json_body=True)[0], 400)
        self.assertEqual(self.request("POST", "/api/import", {"draft": self.example, "extra": 1}, json_body=True)[0], 400)
        self.assertEqual(self.request("POST", "/api/import", {"draft": self.example}, json_body=True,
                                      headers={"X-SuperSlide-Request": ""})[0], 403)

    def test_overflowing_draft_is_reported_for_model_shortening(self):
        long_bullet = ("Подробное описание направления развития производства с указанием "
                       "конкретных мероприятий, сроков и ответственных подразделений " * 2)[:130]

        def base(bullets):
            return {"template_version": "1.2", "title": "Проверка", "slides": [
                {"layout_id": "title_slide", "title": "Отчёт", "subtitle": "Демонстрация"},
                {"layout_id": "bullet_list", "title": "Направления", "bullets": bullets},
                {"layout_id": "closing", "title": "Спасибо", "subtitle": "Вопросы"},
            ]}

        overflowing, fitting, errors = base([long_bullet] * 6), base(["Короткий тезис"]), []

        def generate(task, bundle, settings, validate, **kwargs):
            # Emulate the repair loop: the overflowing draft fails validation,
            # the shortened second attempt is accepted.
            with self.assertRaises(ValueError) as caught:
                validate(overflowing)
            errors.append(str(caught.exception))
            validate(fitting)
            return GenerationResult(fitting, 2)

        with patch("superslide.web.generate_draft", side_effect=generate):
            status, created, _ = self.request("POST", "/api/drafts", {"task": "Test", "slide_count": 3}, json_body=True)
            self.assertEqual(status, 202)
            job = self.poll(created["job_id"])
        self.assertEqual(job["status"], "succeeded", job)
        self.assertEqual(len(job["draft"]["slides"]), 3)
        self.assertTrue(errors)
        self.assertIn("Сократи или суммаризуй", errors[0])
        self.assertIn("Слайд 2, поле bullets", errors[0])

    def test_api_key_is_memory_only_and_never_returned(self):
        key = "temporary-secret-should-not-escape"
        before = set(Path(self.directory.name).rglob("*"))
        status, data, _ = self.request("POST", "/api/settings", {"api_key": key}, json_body=True)
        self.assertEqual(status, 200)
        self.assertEqual(data, {"key_configured": True})
        self.assertEqual(os.environ["SUPERSLIDE_WEB_TEST_KEY"], key)
        self.assertEqual(set(Path(self.directory.name).rglob("*")), before)
        self.assertNotIn(key, json.dumps(self.request("GET", "/api/bootstrap")[1]))
        for invalid in ("", "a\nb", "a" * 513):
            self.assertEqual(self.request("POST", "/api/settings", {"api_key": invalid}, json_body=True)[0], 400)

    def test_failed_llm_job_redacts_raw_exception_and_diagnostics(self):
        secret = "private-document-secret-with-key"
        with patch("superslide.web.generate_draft", side_effect=LLMError(secret, code="connection", diagnostics=[{"raw": secret}])):
            status, created, _ = self.request("POST", "/api/drafts", {"task": "Test", "slide_count": 3}, json_body=True)
            self.assertEqual(status, 202)
            job = self.poll(created["job_id"])
        self.assertEqual(job["status"], "failed")
        self.assertNotIn(secret, json.dumps(job))
        self.assertNotIn("diagnostics", job)

    def test_unknown_assets_variants_and_sources_are_rejected_before_build(self):
        for mutate in (
            lambda d: d["slides"][2]["image"].update(asset_id="unknown_asset"),
            lambda d: d["slides"][0].update(variant_id="invented"),
        ):
            draft = copy.deepcopy(self.example)
            mutate(draft)
            status, data, _ = self.request("POST", "/api/build", {"draft": draft}, json_body=True)
            self.assertEqual(status, 400)
            self.assertNotIn("job_id", data)
        self.assertEqual(self.request("POST", "/api/build", {"draft": self.example,
                                                              "source_job_id": str(uuid.uuid4())}, json_body=True)[0], 404)

    def test_paths_host_origin_and_mutator_header_are_guarded(self):
        for path in ("/../config/llm.local.json", "/src/superslide/config.py", "/assets/../llm.local.json"):
            self.assertEqual(self.request("GET", path)[0], 404)
        self.assertEqual(self.request("GET", "/api/bootstrap", headers={"Host": "evil.invalid"})[0], 403)
        self.assertEqual(self.request("GET", "/api/bootstrap", headers={"Origin": "https://evil.invalid"})[0], 403)
        self.assertEqual(self.request("POST", "/api/settings", {"api_key": "test"}, json_body=True,
                                      headers={"X-SuperSlide-Request": ""})[0], 403)
        self.assertEqual(self.request("POST", "/api/build", {}, json_body=True,
                                      headers={"Origin": "null"})[0], 403)
        self.assertEqual(self.request("GET", f"/api/jobs/{uuid.uuid4()}")[0], 404)

    def test_upload_filename_types_and_size_limits(self):
        for filename in ("../test.txt", "C:\\test.txt", "/test.txt", "CON.txt", "test.txt.", "a:b.txt"):
            self.assertEqual(self.upload(filename)[0], 400, filename)
        self.assertEqual(self.upload("script.exe")[0], 415)
        self.assertEqual(self.upload("empty.txt", b"")[0], 400)
        self.assertEqual(self.upload("large.txt", b"",)[0], 400)
        status, _, _ = self.request("POST", "/api/uploads?filename=large.txt", b"",
                                    {"Content-Type": "application/octet-stream", "Content-Length": str(MAX_UPLOAD + 1)})
        self.assertEqual(status, 413)
        self.assertEqual(self.request("POST", "/api/build", b"", {
            "Content-Type": "application/json", "Content-Length": str(MAX_JSON + 1)})[0], 413)
        self.assertEqual(self.request("POST", "/api/build", b"not-json", {"Content-Type": "application/json"})[0], 400)

    def test_missing_key_and_unknown_uploads_do_not_start_generation(self):
        with patch.dict(os.environ, {"SUPERSLIDE_WEB_TEST_KEY": ""}):
            self.assertEqual(self.request("POST", "/api/drafts", {"task": "Test"}, json_body=True)[0], 400)
        status, _, _ = self.request("POST", "/api/drafts", {"task": "Test", "document_ids": [str(uuid.uuid4())]}, json_body=True)
        self.assertEqual(status, 404)
        for count in (2, 21, True):
            self.assertEqual(self.request("POST", "/api/drafts", {"task": "Test", "slide_count": count}, json_body=True)[0], 400)

    def test_upload_count_and_aggregate_size_are_enforced(self):
        ids = []
        try:
            for n in range(10):
                status, uploaded, _ = self.upload(f"source-{n}.txt", b"Local")
                self.assertEqual(status, 201)
                ids.append(uploaded["id"])
            self.assertEqual(self.upload("extra.txt")[0], 413)
        finally:
            for upload_id in ids:
                self.request("DELETE", f"/api/uploads/{upload_id}")
        ids = []
        try:
            with patch("superslide.web.MAX_TOTAL_UPLOAD", 10):
                status, uploaded, _ = self.upload("small.txt", b"123456")
                self.assertEqual(status, 201)
                ids.append(uploaded["id"])
                self.assertEqual(self.upload("second.txt", b"123456")[0], 413)
        finally:
            for upload_id in ids:
                self.request("DELETE", f"/api/uploads/{upload_id}")

    def test_upload_cannot_be_deleted_while_generation_reads_it(self):
        _, uploaded, _ = self.upload("in-use.txt", b"Source")
        started, release = threading.Event(), threading.Event()
        payload = copy.deepcopy(self.example)
        payload["slides"] = payload["slides"][:3]

        def generate(*args, **kwargs):
            started.set()
            release.wait(3)
            kwargs["validate"](payload)
            return GenerationResult(payload, 1)

        try:
            with patch("superslide.web.generate_draft", side_effect=generate):
                status, created, _ = self.request("POST", "/api/drafts", {
                    "task": "Test", "slide_count": 3, "document_ids": [uploaded["id"]]}, json_body=True)
                self.assertEqual(status, 202)
                self.assertTrue(started.wait(2))
                self.assertEqual(self.request("DELETE", f"/api/uploads/{uploaded['id']}")[0], 409)
                release.set()
                self.assertEqual(self.poll(created["job_id"])["status"], "succeeded")
        finally:
            release.set()
            self.request("DELETE", f"/api/uploads/{uploaded['id']}")

    def test_public_binding_is_not_allowed(self):
        with self.assertRaises(ValueError):
            create_server(host="0.0.0.0", port=0, config_path=self.config)


if __name__ == "__main__":
    unittest.main()
