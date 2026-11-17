import io
import json
import os
import socket
import unittest
import urllib.error
import urllib.request
from pathlib import Path
from unittest.mock import MagicMock, patch

from support import temporary_directory
from superslide.config import ConfigurationError, LLMSettings, load_llm_settings
from superslide.documents import DocumentBundle, SourceChunk
from superslide.llm import LLMError, _NoRedirects, _strict_response_schema, generate_draft, list_models


def api_response(content, *, finish_reason="stop"):
    response = MagicMock()
    response.__enter__.return_value = response
    response.read.return_value = json.dumps({"choices": [{
        "message": {"content": content}, "finish_reason": finish_reason,
    }]}).encode("utf-8")
    return response


def validate(draft):
    if not isinstance(draft.get("title"), str) or not draft["title"]:
        raise ValueError("title должен быть непустой строкой")


class LLMTests(unittest.TestCase):
    def setUp(self):
        self.settings = LLMSettings("https://internal.example/v1", "corporate", temperature=None)
        self.bundle = DocumentBundle([SourceChunk("doc1:text", "facts.txt", "text", "Revenue 125")])
        self.schema = {"type": "object", "properties": {"title": {"type": "string"}}}

    def generate(self, settings=None):
        return generate_draft("Составь презентацию", self.bundle, settings or self.settings, validate,
                              self.schema, {"bullet_list": {"max_bullets": 5}})

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_configured_endpoint_payload_and_authorization(self, request):
        request.return_value = api_response('{"title":"Отчёт"}')
        with patch.dict(os.environ, {"SUPERSLIDE_LLM_API_KEY": "secret-token"}):
            result = self.generate()
        self.assertEqual(result.draft["title"], "Отчёт")
        self.assertEqual(result.attempts, 1)
        call = request.call_args
        sent = call.args[0]
        self.assertEqual(sent.full_url, "https://internal.example/v1/chat/completions")
        self.assertEqual(sent.get_header("Authorization"), "Bearer secret-token")
        self.assertEqual(call.kwargs["timeout"], 120)
        body = json.loads(sent.data)
        self.assertEqual(body["model"], "corporate")
        self.assertEqual(body["response_format"], {"type": "json_object"})
        self.assertEqual(body["max_tokens"], 6000)
        self.assertNotIn("temperature", body)
        self.assertNotIn("secret-token", json.dumps(body))
        self.assertIn("doc1:text", body["messages"][1]["content"])

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_json_and_schema_failures_are_repaired(self, request):
        request.side_effect = [api_response("not json"), api_response('{"title":42}'),
                               api_response('{"title":"Исправлено"}')]
        result = self.generate()
        self.assertEqual(result.attempts, 3)
        self.assertEqual(len(result.validation_errors), 2)
        self.assertEqual(request.call_count, 3)
        last = json.loads(request.call_args.args[0].data)
        self.assertEqual(len(last["messages"]), 6)

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_response_format_modes(self, request):
        request.return_value = api_response('{"title":"OK"}')
        self.generate(LLMSettings("http://local/v1", "local", response_format="json_schema",
                                  token_limit_parameter="max_completion_tokens"))
        body = json.loads(request.call_args.args[0].data)
        self.assertEqual(body["response_format"]["json_schema"]["schema"], {
            **self.schema, "required": ["title"], "additionalProperties": False,
        })
        self.assertIn("max_completion_tokens", body)
        self.generate(LLMSettings("http://local/v1", "local", response_format="none"))
        self.assertNotIn("response_format", json.loads(request.call_args.args[0].data))

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_strict_request_schema_normalizes_real_models_without_mutating_contract(self, request):
        from superslide.models import presentation_schema
        schema = presentation_schema()
        original = json.dumps(schema, ensure_ascii=False, sort_keys=True)
        request.return_value = api_response('{"title":"OK"}')
        generate_draft("Задание", self.bundle,
                       LLMSettings("http://local/v1", "local", response_format="json_schema"),
                       validate, schema, {})
        sent_schema = json.loads(request.call_args.args[0].data)["response_format"]["json_schema"]["schema"]

        def check(node):
            if isinstance(node, list):
                for item in node:
                    check(item)
            elif isinstance(node, dict):
                self.assertNotIn("default", node)
                self.assertNotIn("discriminator", node)
                self.assertNotIn("oneOf", node)
                if node.get("type") == "object" or "properties" in node:
                    self.assertEqual(node["required"], list(node.get("properties", {})))
                    self.assertIs(node["additionalProperties"], False)
                for child in node.values():
                    check(child)

        check(sent_schema)
        self.assertIn("anyOf", sent_schema["properties"]["slides"]["items"])
        self.assertIn("notes", sent_schema["$defs"]["BulletSlide"]["required"])
        self.assertEqual(sent_schema["$defs"]["BulletSlide"]["properties"]["bullets"]["maxItems"], 6)
        self.assertEqual(json.dumps(schema, ensure_ascii=False, sort_keys=True), original)
        self.assertEqual(presentation_schema(), schema)

    def test_strict_normalization_preserves_schema_field_names_and_literal_values(self):
        schema = {"type": "object", "properties": {
            "default": {"type": "string", "default": "original"},
            "discriminator": {"const": {"default": "literal data"}},
        }}
        normalized = _strict_response_schema(schema)
        self.assertIn("default", normalized["properties"])
        self.assertNotIn("default", normalized["properties"]["default"])
        self.assertEqual(normalized["properties"]["discriminator"]["const"], {"default": "literal data"})
        self.assertEqual(schema["properties"]["default"]["default"], "original")

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_image_prompt_uses_descriptions_and_filters_unavailable_layouts(self, request):
        request.return_value = api_response('{"title":"OK"}')
        layouts = {name: {} for name in ("bullet_list", "photo_text", "image_spotlight", "gallery", "kpi_cards", "timeline")}
        for images, expected_image_layouts in (
            (None, set()),
            ([{"id": "photo_01", "description": "Завод", "tags": ["производство"], "file": "private/photo.png"}],
             {"photo_text", "image_spotlight"}),
            ([{"id": f"photo_{number}", "description": "Завод", "tags": []} for number in range(3)],
             {"photo_text", "image_spotlight", "gallery"}),
        ):
            with self.subTest(images=images):
                generate_draft("Задание", self.bundle, self.settings, validate, self.schema, layouts,
                               image_catalog=images, template_version="1.1")
                body = json.loads(request.call_args.args[0].data)
                prompt = json.loads(body["messages"][1]["content"])
                selected = set(prompt["layout_catalog"]) & {"photo_text", "image_spotlight", "gallery"}
                self.assertEqual(selected, expected_image_layouts)
                self.assertIn("template_version='1.1'", body["messages"][0]["content"])
                self.assertNotIn("private/photo.png", json.dumps(body))
                self.assertEqual(set(layouts), {"bullet_list", "photo_text", "image_spotlight", "gallery", "kpi_cards", "timeline"})
                if images:
                    self.assertEqual(set(prompt["image_catalog"][0]), {"id", "description", "tags"})

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_explicit_legacy_template_version_is_preserved(self, request):
        request.return_value = api_response('{"title":"OK"}')
        generate_draft("Задание", self.bundle, self.settings, validate, self.schema, {}, template_version="1.0")
        body = json.loads(request.call_args.args[0].data)
        self.assertIn("template_version='1.0'", body["messages"][0]["content"])

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_variant_catalog_is_compact_and_preserves_design_descriptions(self, request):
        request.return_value = api_response('{"title":"OK"}')
        layouts = {"bullet_list": {
            "layout_name": "bullet_list", "description": "Базовая композиция",
            "fields": {"title": {"type": "text", "max_lines": 2, "idx": 0, "x": 123,
                                   "font_size_pt": 32}},
            "variants": {"accent_panel": {
                "layout_name": "bullet_list__accent_panel", "description": "Акцентная панель слева",
                "source_reference": {"slide": 41},
                "fields": {"title": {"type": "text", "max_lines": 2, "idx": 0, "x": 456}},
            }},
        }}
        original = json.loads(json.dumps(layouts))
        generate_draft("Задание", self.bundle, self.settings, validate, self.schema, layouts)
        body = json.loads(request.call_args.args[0].data)
        self.assertIn("template_version='1.2'", body["messages"][0]["content"])
        self.assertIn("variant_id", body["messages"][0]["content"])
        catalog = json.loads(body["messages"][1]["content"])["layout_catalog"]
        self.assertEqual(set(catalog["bullet_list"]["variants"]), {"default", "accent_panel"})
        accent = catalog["bullet_list"]["variants"]["accent_panel"]
        self.assertEqual(accent["description"], "Акцентная панель слева")
        self.assertEqual(accent["source_reference"], {"slide": 41})
        self.assertEqual(accent["fields"]["title"], {"type": "text", "max_lines": 2})
        self.assertNotIn("bullet_list__accent_panel", json.dumps(catalog))
        self.assertEqual(layouts, original)

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_legacy_generation_hides_variants_and_title_photo_requires_assets(self, request):
        request.return_value = api_response('{"title":"OK"}')
        layouts = {"title_slide": {"variants": {"photo": {"description": "Фотография"},
                                               "blue_minimal": {"description": "Синий фон"}}}}
        for version, images, expected in (
            ("1.0", None, None), ("1.1", None, None),
            ("1.2", None, {"default", "blue_minimal"}),
            ("1.2", [{"id": "worker", "description": "Рабочий"}], {"default", "blue_minimal", "photo"}),
        ):
            with self.subTest(version=version, images=images):
                generate_draft("Задание", self.bundle, self.settings, validate, self.schema, layouts,
                               template_version=version, image_catalog=images)
                body = json.loads(request.call_args.args[0].data)
                info = json.loads(body["messages"][1]["content"])["layout_catalog"]["title_slide"]
                if expected is None:
                    self.assertNotIn("variants", info)
                else:
                    self.assertEqual(set(info["variants"]), expected)
        self.assertIn("photo", layouts["title_slide"]["variants"])

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_unsupported_template_version_fails_before_network(self, request):
        with self.assertRaises(LLMError) as failure:
            generate_draft("Задание", self.bundle, self.settings, validate, self.schema, {},
                           template_version="3.0")
        self.assertEqual(failure.exception.code, "configuration")
        request.assert_not_called()

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_validation_diagnostics_do_not_expose_auth_or_urls(self, request):
        request.return_value = api_response('{"title":42,"api_key":"secret-token","url":"https://private/token"}')
        with patch.dict(os.environ, {"SUPERSLIDE_LLM_API_KEY": "secret-token"}):
            with self.assertRaises(LLMError) as failure:
                self.generate(LLMSettings("http://local/v1", "local", max_repair_attempts=0))
        self.assertEqual(failure.exception.code, "validation")
        diagnostics = json.dumps(failure.exception.diagnostics)
        self.assertNotIn("secret-token", diagnostics)
        self.assertNotIn("https://private", diagnostics)
        self.assertNotIn("secret-token", str(failure.exception))

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_http_auth_error_is_safe_and_never_retried(self, request):
        request.side_effect = urllib.error.HTTPError(
            "https://private/secret", 401, "secret-token", {}, io.BytesIO(b"secret-token"),
        )
        with self.assertRaises(LLMError) as failure:
            self.generate()
        self.assertEqual(failure.exception.code, "http_401")
        self.assertEqual(request.call_count, 1)
        self.assertNotIn("private", str(failure.exception))
        self.assertNotIn("secret-token", str(failure.exception))

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_timeout_and_connection_errors_are_safe(self, request):
        for exception, code in ((socket.timeout("secret"), "timeout"),
                                (urllib.error.URLError(socket.timeout("secret")), "timeout"),
                                (urllib.error.URLError("https://private?token=secret"), "connection")):
            with self.subTest(code=code):
                request.side_effect = exception
                with self.assertRaises(LLMError) as failure:
                    self.generate()
                self.assertEqual(failure.exception.code, code)
                self.assertNotIn("secret", str(failure.exception))

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_limits_before_network_and_truncated_provider_output(self, request):
        with self.assertRaises(LLMError) as failure:
            self.generate(LLMSettings("http://local/v1", "local", max_prompt_chars=10))
        self.assertEqual(failure.exception.code, "input_limit")
        request.assert_not_called()
        with self.assertRaises(LLMError) as failure:
            self.generate(LLMSettings("http://local/v1", "local", max_output_tokens=10, context_window_tokens=20))
        self.assertEqual(failure.exception.code, "context_limit")
        request.assert_not_called()
        request.return_value = api_response('{"title":"OK"}', finish_reason="length")
        with self.assertRaises(LLMError) as failure:
            self.generate()
        self.assertEqual(failure.exception.code, "output_limit")

    def test_redirects_are_not_followed(self):
        self.assertIsNone(_NoRedirects().redirect_request(None, None, 302, "", {}, "http://other"))

    def test_config_load_and_secret_rejection(self):
        with temporary_directory() as directory:
            path = Path(directory) / "settings.json"
            path.write_text('{"base_url":"http://local/v1","model":"m","response_format":"none"}', encoding="utf-8-sig")
            with patch.dict(os.environ, {"SUPERSLIDE_LLM_MODEL": "override"}):
                settings = load_llm_settings(path)
            self.assertEqual(settings.model, "override")
            self.assertEqual(settings.response_format, "none")
            path.write_text('{"api_key":"secret"}', encoding="utf-8")
            with self.assertRaises(ConfigurationError):
                load_llm_settings(path)
        for invalid in ("https://user:password@local/v1", "https://local/v1?token=secret", "file:///tmp/api"):
            with self.subTest(url=invalid), self.assertRaises(ConfigurationError):
                LLMSettings(invalid, "model")
        for invalid in ("false", 0, 1, None):
            with self.subTest(proxy=invalid), self.assertRaises(ConfigurationError):
                LLMSettings("http://internal/v1", "model", trust_env_proxy=invalid)

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_proxy_routing_option_does_not_disable_tls_verification(self, request):
        request.return_value = api_response('{"title":"OK"}')
        original_build_opener = urllib.request.build_opener
        for trust_env_proxy in (True, False):
            with self.subTest(trust_env_proxy=trust_env_proxy):
                with patch("superslide.llm.urllib.request.build_opener", wraps=original_build_opener) as build:
                    self.generate(LLMSettings("https://internal/v1", "model", trust_env_proxy=trust_env_proxy))
                handlers = build.call_args.args
                proxy_handlers = [handler for handler in handlers if isinstance(handler, urllib.request.ProxyHandler)]
                self.assertEqual(len(proxy_handlers), 0 if trust_env_proxy else 1)
                if proxy_handlers:
                    self.assertEqual(proxy_handlers[0].proxies, {})
                self.assertTrue(any(isinstance(handler, _NoRedirects) for handler in handlers))
                # No replacement HTTPS handler or SSL context is supplied.
                self.assertFalse(any(isinstance(handler, urllib.request.HTTPSHandler) for handler in handlers))


class ModelDiscoveryTests(unittest.TestCase):
    @staticmethod
    def response(payload):
        response = MagicMock()
        response.__enter__.return_value = response
        response.read.return_value = json.dumps(payload).encode("utf-8")
        return response

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_model_aliases_get_endpoint_env_auth_and_deduplication(self, request):
        request.return_value = self.response({"object": "list", "data": [
            {"id": "corporate-a"}, {"id": "corporate-b"}, {"id": "corporate-a"},
        ]})
        with patch.dict(os.environ, {"DISCOVERY_TEST_TOKEN": "test-secret"}):
            aliases = list_models("https://internal.example/prefix/v1/", "DISCOVERY_TEST_TOKEN", 7)
        self.assertEqual(aliases, ["corporate-a", "corporate-b"])
        sent_request = request.call_args.args[0]
        self.assertEqual(sent_request.method, "GET")
        self.assertEqual(sent_request.full_url, "https://internal.example/prefix/v1/models")
        self.assertEqual(sent_request.get_header("Authorization"), "Bearer test-secret")
        self.assertIsNone(sent_request.data)
        self.assertEqual(request.call_args.kwargs["timeout"], 7)
        request.return_value.read.assert_called_once_with(2 * 1024 * 1024 + 1)

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_empty_model_list_without_auth_is_valid(self, request):
        request.return_value = self.response({"data": []})
        with patch.dict(os.environ, {}, clear=True):
            self.assertEqual(list_models("http://internal/v1"), [])
        self.assertIsNone(request.call_args.args[0].get_header("Authorization"))

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_model_response_shape_and_size_limits(self, request):
        for payload in ({"models": []}, {"data": [{"name": "alias"}]}, {"data": [{"id": ""}]},
                        {"data": [{"id": "line\nbreak"}]}, {"data": [{"id": 12}]}):
            with self.subTest(payload=payload):
                request.return_value = self.response(payload)
                with self.assertRaises(LLMError) as failure:
                    list_models("http://internal/v1")
                self.assertEqual(failure.exception.code, "response_shape")
        request.return_value.read.return_value = b"x" * (2 * 1024 * 1024 + 1)
        with self.assertRaises(LLMError) as failure:
            list_models("http://internal/v1")
        self.assertEqual(failure.exception.code, "response_limit")

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_discovery_errors_and_redirects_are_safe(self, request):
        for error, code in (
            (urllib.error.HTTPError("https://private/token", 403, "test-secret", {}, io.BytesIO(b"test-secret")), "http_403"),
            (urllib.error.HTTPError("https://private/token", 302, "test-secret", {}, io.BytesIO()), "http_302"),
            (socket.timeout("test-secret"), "timeout"),
            (urllib.error.URLError("https://private/token test-secret"), "connection"),
        ):
            with self.subTest(code=code):
                request.side_effect = error
                with self.assertRaises(LLMError) as failure:
                    list_models("http://internal/v1")
                self.assertEqual(failure.exception.code, code)
                self.assertNotIn("test-secret", str(failure.exception))
                self.assertNotIn("https://private", str(failure.exception))

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_discovery_invalid_configuration_does_not_send_request(self, request):
        for base_url in ("file:///tmp/api", "https://user:secret@internal/v1", "https://internal/v1?token=secret",
                         "https://internal/\n"):
            with self.subTest(base_url=base_url), self.assertRaises(LLMError) as failure:
                list_models(base_url)
            self.assertEqual(failure.exception.code, "configuration")
        request.assert_not_called()

    @patch("superslide.llm.urllib.request.OpenerDirector.open")
    def test_model_discovery_supports_explicit_direct_routing(self, request):
        request.return_value = self.response({"data": []})
        original_build_opener = urllib.request.build_opener
        with patch("superslide.llm.urllib.request.build_opener", wraps=original_build_opener) as build:
            self.assertEqual(list_models("https://internal/v1", trust_env_proxy=False), [])
        proxies = [handler for handler in build.call_args.args if isinstance(handler, urllib.request.ProxyHandler)]
        self.assertEqual(len(proxies), 1)
        self.assertEqual(proxies[0].proxies, {})
        self.assertFalse(any(isinstance(handler, urllib.request.HTTPSHandler) for handler in build.call_args.args))


if __name__ == "__main__":
    unittest.main()
