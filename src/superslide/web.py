"""Loopback-only local web UI. Documents and generated files stay in this project."""

from __future__ import annotations

import copy
import json
import mimetypes
import os
import re
import threading
import uuid
from concurrent.futures import ThreadPoolExecutor
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path, PureWindowsPath
from urllib.parse import parse_qs, urlsplit

from pydantic import ValidationError

from .assets import AssetError, MAX_IMAGE_BYTES, load_asset_catalog
from .config import ConfigurationError, load_llm_settings
from .documents import DocumentBundle, DocumentError, load_documents
from .llm import LLMError, generate_draft
from .models import iter_image_refs, presentation_schema, validate_presentation
from .render import LayoutOverflowError, build_presentation, check_text_fit
from .template import load_template, select_layout


ROOT = Path(__file__).resolve().parents[2]
STATIC = Path(__file__).resolve().parent / "static"
MAX_JSON = 2 * 1024 * 1024
MAX_UPLOAD = 25 * 1024 * 1024
MAX_TOTAL_UPLOAD = 50 * 1024 * 1024
ALLOWED_EXTENSIONS = {".txt", ".md", ".docx", ".pdf"}
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg"}
MAX_SESSION_IMAGES = 10
MAX_SESSION_IMAGE_BYTES = 100 * 1024 * 1024
DOWNLOADS = {"presentation.pptx", "draft.json", "report.json"}
LAYOUT_NAMES = {
    "title_slide": "Обложка", "bullet_list": "Тезисы", "two_columns": "Две колонки",
    "quote": "Цитата", "table": "Таблица", "bar_chart": "Диаграмма",
    "section_divider": "Разделитель", "photo_text": "Фото и текст",
    "image_spotlight": "Крупное фото", "gallery": "Галерея", "kpi_cards": "Показатели",
    "timeline": "Этапы проекта", "comparison": "Сравнение", "closing": "Финальный слайд",
}
VARIANT_NAMES = {
    "default": "Основной", "blue_minimal": "Синий минимальный", "blue_graphic": "Синий с графикой",
    "photo": "Фотообложка", "accent_panel": "Акцентная плашка", "two_columns": "Две колонки",
    "blue": "Синий фон", "accent": "Красный акцент", "circle_pattern": "Круговой паттерн",
    "diagonal_pattern": "Диагональный паттерн", "photo_right": "Фото справа",
    "photo_band": "Фото полосой", "white_caption": "Белая подпись", "side_panel": "Боковая панель",
    "plain": "Без плашек", "white_graphic": "Белый с графикой",
}


class APIError(Exception):
    def __init__(self, status: int, code: str, message: str):
        self.status, self.code, self.message = status, code, message


def _write_json(path: Path, value) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def _identifier(value: str) -> str:
    try:
        if str(uuid.UUID(value)) != value:
            raise ValueError
    except (ValueError, TypeError, AttributeError):
        raise APIError(404, "not_found", "Объект не найден.") from None
    return value


def _filename(value: str, allowed_extensions, type_message: str) -> str:
    if (not value or len(value) > 180 or value != value.strip() or value.endswith((".", " "))
            or any(ord(c) < 32 or c in '<>:"/\\|?*' for c in value)
            or Path(value).name != value or PureWindowsPath(value).name != value
            or PureWindowsPath(value).is_absolute() or value in {".", ".."}):
        raise APIError(400, "filename", "Укажите имя файла без пути и специальных символов.")
    stem = value.split(".", 1)[0].upper()
    if stem in {"CON", "PRN", "AUX", "NUL", *[f"{p}{n}" for p in ("COM", "LPT") for n in range(1, 10)]}:
        raise APIError(400, "filename", "Это имя файла зарезервировано системой.")
    if Path(value).suffix.lower() not in allowed_extensions:
        raise APIError(415, "file_type", type_message)
    return value


def _safe_error(exc: Exception) -> str:
    # Never send exception text, model responses, API URLs, keys or document contents to the UI.
    if isinstance(exc, LLMError):
        return {
            "authentication": "API отклонил ключ. Проверьте ключ доступа в настройках.",
            "http_401": "API отклонил ключ. Проверьте ключ доступа в настройках.",
            "timeout": "LLM не ответила вовремя. Повторите запрос позже.",
            "validation": "Ответ LLM не прошёл проверку структуры. Уточните задание и повторите.",
        }.get(exc.code, "Не удалось получить ответ корпоративной LLM. Проверьте подключение и настройки.")
    if isinstance(exc, ConfigurationError):
        return "Корпоративная LLM не настроена. Проверьте локальный файл конфигурации."
    if isinstance(exc, DocumentError):
        return "Не удалось прочитать документ. Проверьте формат, размер и наличие текстового слоя."
    if isinstance(exc, LayoutOverflowError):
        return "Текст не помещается в выбранную композицию. Сократите текст или выберите другой вариант."
    if isinstance(exc, ValidationError):
        fields = [".".join(map(str, error["loc"])) for error in exc.errors(include_input=False)[:3]]
        return "Проверьте поля презентации: " + ", ".join(fields) + "."
    return "Не удалось выполнить операцию. Проверьте содержание и выбранные макеты."


class WebState:
    def __init__(self, config_path: Path | None):
        self.config_path = config_path
        self.template = ROOT / "templates" / "corporate_base.pptx"
        self.manifest_path = ROOT / "templates" / "template_manifest.json"
        _, self.manifest, _ = load_template(self.template, self.manifest_path)
        self.catalog = load_asset_catalog(ROOT / "assets" / "images" / "catalog.json")
        self.directory = ROOT / "output" / "web"
        self.image_directory = self.directory / "images"
        self.image_sizes: dict[str, int] = {}
        self.lock = threading.RLock()
        self.uploads: dict[str, dict] = {}
        self.jobs: dict[str, dict] = {}
        self.executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="superslide-web")
        self.previews: dict[tuple[str, str], Path] = {}
        variants = ROOT / "examples" / "variants.json"
        if variants.exists():
            records = json.loads(variants.read_text(encoding="utf-8-sig"))["slides"]
            for n, slide in enumerate(records, 1):
                image = ROOT / "output" / "variants-powerpoint" / f"slide-{n:02d}.png"
                if image.is_file():
                    self.previews[(slide["layout_id"], slide.get("variant_id", "default"))] = image

    def settings(self):
        return load_llm_settings(self.config_path)

    def image_list(self):
        return [{"id": record["id"], "description": record.get("description", ""),
                 "tags": list(record.get("tags", [])), "url": f"/api/images/{record['id']}/file"}
                for record in self.catalog.records]

    def add_image(self, filename: str, content: bytes) -> dict:
        if not content:
            raise APIError(400, "empty_file", "Файл пуст.")
        if len(content) > MAX_IMAGE_BYTES:
            raise APIError(413, "too_large", "Изображение должно быть не больше 15 МиБ.")
        with self.lock:
            if len(self.image_sizes) >= MAX_SESSION_IMAGES:
                raise APIError(429, "image_limit", "Допустимо до 10 загруженных фотографий за сеанс.")
            if sum(self.image_sizes.values()) + len(content) > MAX_SESSION_IMAGE_BYTES:
                raise APIError(413, "image_limit", "Общий размер загруженных фотографий превышает 100 МиБ.")
            asset_id = "upload_" + uuid.uuid4().hex[:12]
            destination = self.image_directory / asset_id / filename
            destination.parent.mkdir(parents=True, exist_ok=True)
            destination.write_bytes(content)
            description = f"Загруженное изображение: {filename}"
            try:
                self.catalog.register(asset_id, destination, description)
            except AssetError as exc:
                destination.unlink(missing_ok=True)
                try:
                    destination.parent.rmdir()
                except OSError:
                    pass
                raise APIError(400, "image", str(exc)) from None
            self.image_sizes[asset_id] = len(content)
            return {"id": asset_id, "description": description, "url": f"/api/images/{asset_id}/file"}

    def bootstrap(self):
        try:
            settings = self.settings()
            model, configured, key = settings.model, True, bool(os.environ.get(settings.api_key_env))
        except ConfigurationError:
            model, configured, key = None, False, False
        layouts = []
        for layout_id, info in self.manifest["layouts"].items():
            variants = []
            for variant_id, variant in [("default", info), *info.get("variants", {}).items()]:
                variants.append({"id": variant_id, "name": VARIANT_NAMES.get(variant_id, variant_id),
                                 "description": variant.get("description", "Фирменная композиция"),
                                 "preview_url": f"/api/previews/{layout_id}/{variant_id}.png"
                                 if (layout_id, variant_id) in self.previews else None})
            layouts.append({"id": layout_id, "name": LAYOUT_NAMES.get(layout_id, layout_id), "variants": variants})
        return {"model": model, "llm_configured": configured, "key_configured": key,
                "template_version": self.manifest["template_version"], "layouts": layouts,
                "design_count": sum(len(item["variants"]) for item in layouts)}

    def validate(self, draft, known_sources=None):
        if not isinstance(draft, dict):
            raise APIError(400, "draft", "Презентация должна быть JSON-объектом.")
        spec = validate_presentation(draft, known_sources)
        for slide in spec.slides:
            select_layout(self.manifest, slide.layout_id, slide.variant_id)
        for image in iter_image_refs(spec):
            self.catalog.resolve(image.asset_id)
        return spec

    def public_job(self, job_id):
        with self.lock:
            job = self.jobs.get(_identifier(job_id))
            if not job:
                raise APIError(404, "not_found", "Задание не найдено.")
            return copy.deepcopy({k: v for k, v in job.items() if not k.startswith("_")})

    def submit(self, kind, work):
        with self.lock:
            if len(self.jobs) >= 100 or sum(j["status"] in {"queued", "running"} for j in self.jobs.values()) >= 8:
                raise APIError(429, "queue_full", "Очередь заполнена. Дождитесь завершения текущих заданий.")
            job_id = str(uuid.uuid4())
            self.jobs[job_id] = {"id": job_id, "kind": kind, "status": "queued", "message": "В очереди"}

        def run():
            with self.lock:
                self.jobs[job_id].update(status="running", message="Генерация структуры" if kind == "draft" else "Сборка PowerPoint")
            try:
                result = work(job_id)
                with self.lock:
                    self.jobs[job_id].update(result, status="succeeded", message="Готово")
            except Exception as exc:
                with self.lock:
                    self.jobs[job_id].update(status="failed", message="Не удалось завершить", error=_safe_error(exc))

        self.executor.submit(run)
        return {"job_id": job_id}

    def draft(self, payload):
        _fields(payload, {"task", "document_ids", "slide_count"})
        task = payload.get("task")
        count = payload.get("slide_count", 8)
        ids = payload.get("document_ids", [])
        if not isinstance(task, str) or not task.strip() or len(task) > 20_000:
            raise APIError(400, "task", "Опишите задачу: от 1 до 20 000 символов.")
        if isinstance(count, bool) or not isinstance(count, int) or not 3 <= count <= 20:
            raise APIError(400, "slide_count", "Количество слайдов: от 3 до 20.")
        if not isinstance(ids, list) or len(ids) > 10 or len(set(str(i) for i in ids)) != len(ids):
            raise APIError(400, "documents", "Выберите не более 10 различных документов.")
        # Fail early, before creating a job or sending documents.
        settings = self.settings()
        if not os.environ.get(settings.api_key_env):
            raise APIError(400, "api_key", "Укажите API-ключ в настройках подключения.")
        with self.lock:
            paths = []
            for upload_id in ids:
                item = self.uploads.get(_identifier(upload_id))
                if not item:
                    raise APIError(404, "not_found", "Загруженный документ не найден.")
                paths.append(item["_path"])
            # Pin selected uploads while an asynchronous job reads them.
            for upload_id in ids:
                self.uploads[upload_id]["_readers"] += 1

        def work(job_id):
            try:
                bundle = load_documents(paths, max_total_chars=160_000) if paths else DocumentBundle()
                known = {chunk.id for chunk in bundle.chunks}

                def validate_generated(value):
                    spec = self.validate(value, known)
                    if len(spec.slides) != count:
                        raise ValueError(f"Ожидается ровно {count} слайдов.")
                    try:
                        # Dry-run render: overlong slide text is reported back so the model can shorten it.
                        check_text_fit(spec, self.template, self.manifest_path, asset_catalog=self.catalog)
                    except LayoutOverflowError as exc:
                        raise ValueError(
                            "Текст не помещается в поля слайдов. Сократи или суммаризуй текст проблемных полей, "
                            "сохранив ключевые факты, числа, source_refs и число слайдов.\n" + str(exc)
                        ) from None
                    return spec

                slide_word = "слайда" if count in (3, 4) else "слайдов"
                result = generate_draft(
                    task.strip() + f"\n\nПодготовь ровно {count} {slide_word}, включая обложку и завершение.",
                    bundle, settings, validate=validate_generated,
                    schema=presentation_schema(), layout_catalog=self.manifest["layouts"],
                    image_catalog=self.catalog.prompt_catalog(), template_version=self.manifest["template_version"],
                )
                validate_generated(result.draft)
                directory = self.directory / "jobs" / job_id
                _write_json(directory / "draft.json", result.draft)
                _write_json(directory / "sources.json", bundle.to_dict())
                return {"draft": result.draft, "warnings": bundle.warnings, "_bundle": bundle}
            finally:
                with self.lock:
                    for upload_id in ids:
                        self.uploads[upload_id]["_readers"] -= 1

        try:
            return self.submit("draft", work)
        except Exception:
            with self.lock:
                for upload_id in ids:
                    self.uploads[upload_id]["_readers"] -= 1
            raise

    def build(self, payload):
        _fields(payload, {"draft", "source_job_id"})
        source_job_id = payload.get("source_job_id")
        bundle = None
        if source_job_id is not None:
            with self.lock:
                source = self.jobs.get(_identifier(source_job_id))
                if not source or source["kind"] != "draft" or source["status"] != "succeeded":
                    raise APIError(404, "not_found", "Завершённое задание с источниками не найдено.")
                bundle = source.get("_bundle")
        lookup = {chunk.id: chunk.to_dict() for chunk in bundle.chunks} if bundle else None
        spec = self.validate(payload.get("draft"), set(lookup) if lookup is not None else None)
        draft = spec.model_dump(mode="json")

        def work(job_id):
            directory = self.directory / "jobs" / job_id
            directory.mkdir(parents=True, exist_ok=True)
            report = build_presentation(spec, self.template, self.manifest_path, directory / "presentation.pptx",
                                        source_lookup=lookup, asset_catalog=self.catalog)
            _write_json(directory / "draft.json", draft)
            # The downloadable report needs no absolute local output path.
            report["output"] = "presentation.pptx"
            _write_json(directory / "report.json", report)
            return {"warnings": report["warnings"], "downloads": {
                "pptx": f"/api/downloads/{job_id}/presentation.pptx", "json": f"/api/downloads/{job_id}/draft.json",
                "report": f"/api/downloads/{job_id}/report.json"}}

        return self.submit("build", work)

    def import_draft(self, payload):
        _fields(payload, {"draft"})
        spec = self.validate(payload.get("draft"))
        # Normalize through the contract so the editor starts from canonical JSON.
        return {"draft": spec.model_dump(mode="json")}


def _fields(payload, allowed):
    if not isinstance(payload, dict) or set(payload) - allowed:
        raise APIError(400, "payload", "Некорректные поля запроса.")


class WebServer(ThreadingHTTPServer):
    daemon_threads = True

    def server_close(self):
        super().server_close()
        if hasattr(self, "state"):
            self.state.executor.shutdown(wait=False, cancel_futures=True)


class Handler(BaseHTTPRequestHandler):
    server_version = "SuperSlide"
    sys_version = ""

    def setup(self):
        super().setup()
        self.connection.settimeout(30)
        self._body_consumed = False

    def log_message(self, *args):
        pass  # Upload names, prompts and query parameters must not enter HTTP logs.

    @property
    def state(self):
        return self.server.state

    def _guard(self, mutating=False):
        port = self.server.server_address[1]
        authorities = {f"127.0.0.1:{port}", f"localhost:{port}", f"[::1]:{port}"}
        if self.headers.get("Host", "").lower() not in authorities:
            raise APIError(403, "host", "Недопустимый адрес сервера.")
        origin = self.headers.get("Origin")
        if origin is not None and origin.lower() not in {f"http://{host}" for host in authorities}:
            raise APIError(403, "origin", "Запрос разрешён только из локального интерфейса.")
        if mutating and self.headers.get("X-SuperSlide-Request") != "1":
            raise APIError(403, "request_header", "Откройте локальный интерфейс SuperSlide.")

    def _headers(self, status, content_type, size, *, attachment=None):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(size))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Referrer-Policy", "no-referrer")
        self.send_header("Content-Security-Policy", "default-src 'self'; img-src 'self' data:; style-src 'self'; script-src 'self'; connect-src 'self'; object-src 'none'; base-uri 'none'; frame-ancestors 'none'")
        if attachment:
            self.send_header("Content-Disposition", f'attachment; filename="{attachment}"')
        self.end_headers()

    def _json(self, value, status=200):
        raw = json.dumps(value, ensure_ascii=False).encode("utf-8")
        self._headers(status, "application/json; charset=utf-8", len(raw))
        self.wfile.write(raw)

    def _file(self, path, attachment=None):
        if not path.is_file():
            raise APIError(404, "not_found", "Файл не найден.")
        mime = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        if path.suffix == ".js":
            mime = "text/javascript"
        self._headers(200, mime, path.stat().st_size, attachment=attachment)
        with path.open("rb") as stream:
            while chunk := stream.read(64 * 1024):
                self.wfile.write(chunk)

    def _body(self, limit):
        if self.headers.get("Transfer-Encoding") is not None:
            raise APIError(400, "encoding", "Потоковая кодировка запроса не поддерживается.")
        try:
            size = int(self.headers.get("Content-Length", "-1"))
        except ValueError:
            size = -1
        if size < 0:
            raise APIError(411, "length", "Укажите размер запроса.")
        if size > limit:
            raise APIError(413, "too_large", "Запрос превышает допустимый размер.")
        raw = self.rfile.read(size)
        self._body_consumed = True
        if len(raw) != size:
            raise APIError(400, "body", "Запрос получен не полностью.")
        return raw

    def _payload(self):
        if self.headers.get_content_type() != "application/json":
            raise APIError(415, "content_type", "Ожидается application/json.")
        try:
            return json.loads(self._body(MAX_JSON).decode("utf-8"))
        except (ValueError, UnicodeError):
            raise APIError(400, "json", "Некорректный JSON.") from None

    def _dispatch(self):
        mutating = self.command != "GET"
        self._guard(mutating)
        url = urlsplit(self.path)
        path = url.path
        if self.command == "GET":
            if path == "/api/bootstrap":
                return self._json(self.state.bootstrap())
            if path == "/api/example":
                draft = json.loads((ROOT / "examples" / "variety.json").read_text(encoding="utf-8-sig"))
                draft["template_version"] = self.state.manifest["template_version"]
                self.state.validate(draft)
                return self._json({"draft": draft, "warnings": ["Демонстрационный пример: все показатели и планы вымышлены."]})
            if match := re.fullmatch(r"/api/previews/([a-z_]+)/([a-z_]+)\.png", path):
                image = self.state.previews.get(match.groups())
                if image:
                    return self._file(image)
            if path == "/api/images":
                return self._json({"images": self.state.image_list()})
            if match := re.fullmatch(r"/api/images/([A-Za-z0-9][A-Za-z0-9_.-]{0,119})/file", path):
                try:
                    image_path = self.state.catalog.resolve(match[1])
                except AssetError:
                    raise APIError(404, "not_found", "Изображение не найдено.") from None
                return self._file(image_path)
            if match := re.fullmatch(r"/api/jobs/([a-f0-9-]+)", path):
                return self._json(self.state.public_job(match[1]))
            if match := re.fullmatch(r"/api/downloads/([a-f0-9-]+)/([^/]+)", path):
                job_id, filename = match.groups()
                job = self.state.public_job(job_id)
                if filename in DOWNLOADS and job["status"] == "succeeded" and job["kind"] == "build":
                    return self._file(self.state.directory / "jobs" / job_id / filename, attachment=filename)
            static = {"/": "index.html", "/index.html": "index.html", "/app.js": "app.js",
                      "/styles.css": "styles.css", "/assets/logo.svg": "assets/logo.svg"}
            if path in static:
                return self._file(STATIC / static[path])
        elif self.command == "POST":
            if path == "/api/settings":
                payload = self._payload()
                _fields(payload, {"api_key"})
                key = payload.get("api_key")
                if (not isinstance(key, str) or not key.strip() or len(key) > 512
                        or any(ord(c) < 32 or ord(c) == 127 for c in key)):
                    raise APIError(400, "api_key", "Ключ должен содержать от 1 до 512 символов без переносов строк.")
                settings = self.state.settings()
                os.environ[settings.api_key_env] = key.strip()
                return self._json({"key_configured": True})
            if path == "/api/uploads":
                query = parse_qs(url.query, keep_blank_values=True)
                if set(query) != {"filename"} or len(query["filename"]) != 1:
                    raise APIError(400, "filename", "Укажите одно имя файла.")
                filename = _filename(query["filename"][0], ALLOWED_EXTENSIONS, "Поддерживаются TXT, MD, DOCX и PDF.")
                if self.headers.get_content_type() != "application/octet-stream":
                    raise APIError(415, "content_type", "Ожидается бинарное содержимое файла.")
                content = self._body(MAX_UPLOAD)
                if not content:
                    raise APIError(400, "empty_file", "Файл пуст.")
                with self.state.lock:
                    if len(self.state.uploads) >= 10 or sum(i["size"] for i in self.state.uploads.values()) + len(content) > MAX_TOTAL_UPLOAD:
                        raise APIError(413, "upload_limit", "Допустимо до 10 файлов общим размером до 50 МБ.")
                    upload_id = str(uuid.uuid4())
                    destination = self.state.directory / "uploads" / upload_id / filename
                    destination.parent.mkdir(parents=True, exist_ok=True)
                    destination.write_bytes(content)
                    record = {"id": upload_id, "name": filename, "size": len(content)}
                    self.state.uploads[upload_id] = {**record, "_path": destination, "_readers": 0}
                return self._json(record, 201)
            if path == "/api/images":
                query = parse_qs(url.query, keep_blank_values=True)
                if set(query) != {"filename"} or len(query["filename"]) != 1:
                    raise APIError(400, "filename", "Укажите одно имя файла.")
                filename = _filename(query["filename"][0], IMAGE_EXTENSIONS, "Поддерживаются PNG и JPEG.")
                if self.headers.get_content_type() != "application/octet-stream":
                    raise APIError(415, "content_type", "Ожидается бинарное содержимое файла.")
                content = self._body(MAX_IMAGE_BYTES + 1)
                return self._json(self.state.add_image(filename, content), 201)
            if path == "/api/drafts":
                return self._json(self.state.draft(self._payload()), 202)
            if path == "/api/build":
                return self._json(self.state.build(self._payload()), 202)
            if path == "/api/import":
                return self._json(self.state.import_draft(self._payload()))
        elif self.command == "DELETE":
            if match := re.fullmatch(r"/api/uploads/([a-f0-9-]+)", path):
                upload_id = _identifier(match[1])
                with self.state.lock:
                    item = self.state.uploads.get(upload_id)
                    if not item:
                        raise APIError(404, "not_found", "Документ не найден.")
                    if item["_readers"]:
                        raise APIError(409, "in_use", "Документ используется текущим заданием.")
                    item["_path"].unlink(missing_ok=True)
                    item["_path"].parent.rmdir()
                    del self.state.uploads[upload_id]
                return self._json({"deleted": True})
        raise APIError(404, "not_found", "Ресурс не найден.")

    def _handle(self):
        # Do not emit internal exceptions to the browser or log.
        try:
            self._dispatch()
        except APIError as exc:
            self._discard_small_body()
            self._json({"error": {"code": exc.code, "message": exc.message}}, exc.status)
        except (ValidationError, ValueError, ConfigurationError) as exc:
            self._json({"error": {"code": "validation", "message": _safe_error(exc)}}, 400)
        except (BrokenPipeError, ConnectionResetError, TimeoutError):
            pass
        except Exception:
            self._json({"error": {"code": "internal", "message": "Не удалось выполнить запрос."}}, 500)

    def _discard_small_body(self):
        # Drain ordinary rejected requests before closing: Windows may otherwise reset
        # the connection with unread request bytes and hide the useful JSON error.
        if self._body_consumed or self.headers.get("Transfer-Encoding") is not None:
            return
        try:
            size = int(self.headers.get("Content-Length", "0"))
            if 0 < size <= MAX_JSON:
                self.connection.settimeout(1)
                self.rfile.read(size)
        except (ValueError, OSError):
            pass

    do_GET = _handle
    do_POST = _handle
    do_DELETE = _handle
    do_OPTIONS = _handle


def create_server(host="127.0.0.1", port=8765, config_path=ROOT / "config" / "llm.local.json"):
    """Create a loopback server. Port 0 is useful for local integration tests."""
    if host not in {"127.0.0.1", "localhost"}:
        raise ValueError("Веб-интерфейс разрешено запускать только на 127.0.0.1 или localhost.")
    if isinstance(port, bool) or not isinstance(port, int) or not 0 <= port <= 65535:
        raise ValueError("Некорректный порт веб-интерфейса.")
    state = WebState(Path(config_path) if config_path is not None else None)
    try:
        server = WebServer((host, port), Handler)
    except Exception:
        state.executor.shutdown(wait=False, cancel_futures=True)
        raise
    server.state = state
    return server


def serve(host="127.0.0.1", port=8765, config_path=ROOT / "config" / "llm.local.json"):
    server = create_server(host, port, config_path)
    print(f"SuperSlide: http://{host}:{server.server_address[1]}")
    print("Остановка: Ctrl+C. API-ключ хранится только в памяти процесса.")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
