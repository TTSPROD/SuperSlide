"""Small local CLI: inspect -> extract -> draft -> validate -> build."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from pydantic import ValidationError

from .assets import load_asset_catalog
from .config import ConfigurationError, load_llm_settings
from .documents import DocumentError, DocumentBundle, load_documents
from .llm import LLMError, generate_draft, list_models
from .models import iter_image_refs, load_presentation, presentation_schema, validate_presentation
from .paths import (
    PROJECT_ROOT, OUTPUT_DIR, DEFAULT_TEMPLATE, DEFAULT_MANIFEST,
    DEFAULT_ASSET_CATALOG, SOURCE_PHOTO_CATALOG, SOURCE_ICON_CATALOG,
)
from .render import LayoutOverflowError, build_presentation, check_text_fit
from .template import inspect_template, load_template, select_layout


def _json_text(value):
    return json.dumps(value, ensure_ascii=False, indent=2) + "\n"


def _write_json(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(_json_text(value), encoding="utf-8")


def _source_lookup(path):
    if not path:
        return None
    path = Path(path)
    if path.stat().st_size > 5_000_000:
        raise ValueError("Файл источников превышает 5 МБ")
    data = json.loads(path.read_text(encoding="utf-8-sig"))
    if not isinstance(data, dict) or not isinstance(data.get("chunks"), list):
        raise ValueError("Файл источников должен содержать массив chunks; используйте extract")
    lookup = {}
    for item in data["chunks"]:
        if not isinstance(item, dict) or any(not isinstance(item.get(k), str) for k in ("id", "document", "location", "text")):
            raise ValueError("Некорректная запись в файле источников")
        if item["id"] in lookup:
            raise ValueError("Повторяющийся идентификатор источника")
        lookup[item["id"]] = item
    return lookup


def _template_flags(parser):
    parser.add_argument("--template", type=Path, default=DEFAULT_TEMPLATE)
    parser.add_argument("--manifest", type=Path, default=DEFAULT_MANIFEST)


def _asset_flag(parser):
    parser.add_argument("--asset-catalog", type=Path,
                        help="Каталог локальных PNG/JPEG с описаниями")
    parser.add_argument("--icon-catalog", type=Path, default=SOURCE_ICON_CATALOG,
                        help="Каталог иконок из исходной библиотеки")


def _assets(path):
    path = path or DEFAULT_ASSET_CATALOG
    if path == DEFAULT_ASSET_CATALOG and not path.exists():
        return None
    return load_asset_catalog(path)


def _read_json(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def _source_mode(manifest_path):
    return _read_json(manifest_path).get("template_version") == "2.0"


def _source_catalog(path, kind):
    """Read the extraction catalogs without imposing the legacy image contract."""
    data = _read_json(path)
    if isinstance(data, list):
        return [item for item in data if item.get("selectable", True)]
    for key in (kind, "images" if kind == "photos" else kind):
        if isinstance(data.get(key), list):
            return [item for item in data[key] if item.get("selectable", True)]
    raise ValueError(f"Каталог {path} должен содержать массив {kind}")


def _validate_source_assets(spec, photo_path, icon_path, manifest=None):
    from .source_template import validate_source_icon_styles
    photo_ids = {item["id"] for item in _source_catalog(photo_path, "photos")}
    icons = _source_catalog(icon_path, "icons")
    icon_ids = {item["id"] for item in icons}
    for slide in spec.slides:
        for update in slide.updates:
            values = update.model_dump() if hasattr(update, "model_dump") else update
            for key, allowed in (("photo_id", photo_ids), ("icon_id", icon_ids)):
                if values.get(key) is not None and values[key] not in allowed:
                    raise ValueError(f"Неизвестный {key}: {values[key]}")
    return validate_source_icon_styles(spec, icons, manifest)


def _source_load(input_path, manifest_path, sources, photo_path, icon_path):
    from .source_template import load_source_manifest, validate_source_presentation
    manifest = load_source_manifest(manifest_path)
    return _validate_source_assets(validate_source_presentation(
        _read_json(input_path), manifest,
        set(sources) if sources is not None else None), photo_path, icon_path, manifest)


def _validate_assets(spec, catalog):
    images = list(iter_image_refs(spec))
    if images and catalog is None:
        raise ValueError("Для слайдов с изображениями нужен --asset-catalog")
    for image in images:
        catalog.resolve(image.asset_id)
    return spec


def _catalog_paths(catalog):
    return [catalog.resolve(record["id"]) for record in catalog.records] if catalog else []


def _validate_designs(spec, manifest):
    for slide in spec.slides:
        select_layout(manifest, slide.layout_id, slide.variant_id)
    return spec


def parser():
    cli = argparse.ArgumentParser(prog="superslide", description="Локальный генератор корпоративных презентаций")
    sub = cli.add_subparsers(dest="command", required=True)
    web = sub.add_parser("serve", help="Открыть локальный веб-интерфейс")
    web.add_argument("--host", choices=("127.0.0.1", "localhost"), default="127.0.0.1")
    web.add_argument("--port", type=int, default=8765)
    web.add_argument("--config", type=Path, default=PROJECT_ROOT / "config" / "llm.local.json")
    web.add_argument("--open", action="store_true", help="Открыть интерфейс в браузере")
    inspect = sub.add_parser("inspect-template", help="Проверить макеты и привязки полей")
    _template_flags(inspect)
    validate = sub.add_parser("validate", help="Проверить JSON презентации")
    validate.add_argument("input", type=Path)
    validate.add_argument("--sources", type=Path)
    _template_flags(validate)
    _asset_flag(validate)
    build = sub.add_parser("build", help="Собрать нативный PPTX из JSON")
    build.add_argument("input", type=Path)
    build.add_argument("--output", type=Path, help="PPTX; по умолчанию output/<имя входного файла>.pptx")
    build.add_argument("--report", type=Path)
    build.add_argument("--sources", type=Path)
    _template_flags(build)
    _asset_flag(build)
    assets = sub.add_parser("assets", help="Проверить и показать локальный каталог изображений")
    _asset_flag(assets)
    extract = sub.add_parser("extract", help="Извлечь текст и источники из TXT/MD/DOCX/PDF")
    extract.add_argument("documents", type=Path, nargs="+")
    extract.add_argument("--output", type=Path, help="JSON источников; по умолчанию output/sources.json")
    extract.add_argument("--max-total-chars", type=int, default=160_000)
    draft = sub.add_parser("draft", help="Получить JSON от настроенной корпоративной LLM")
    task = draft.add_mutually_exclusive_group(required=True)
    task.add_argument("--task", help="Описание задачи")
    task.add_argument("--task-file", type=Path)
    draft.add_argument("--documents", type=Path, nargs="*", default=[])
    draft.add_argument("--config", type=Path)
    draft.add_argument("--output", type=Path, help="JSON презентации; по умолчанию output/draft.json")
    draft.add_argument("--sources-output", type=Path)
    draft.add_argument("--max-total-chars", type=int, default=160_000)
    _template_flags(draft)
    _asset_flag(draft)
    schema = sub.add_parser("schema", help="Вывести JSON Schema контракта презентации")
    schema.add_argument("--output", type=Path)
    _template_flags(schema)
    models = sub.add_parser("models", help="Получить доступные model aliases из внутреннего API")
    address = models.add_mutually_exclusive_group(required=True)
    address.add_argument("--base-url")
    address.add_argument("--config", type=Path)
    models.add_argument("--api-key-env", default="SUPERSLIDE_LLM_API_KEY")
    models.add_argument("--no-proxy", action="store_true", help="Прямое соединение без прокси окружения")
    return cli


def _check_paths(outputs, protected):
    outputs = [Path(p).resolve() for p in outputs if p]
    protected = {Path(p).resolve() for p in protected if p}
    if len(set(outputs)) != len(outputs):
        raise ValueError("Пути выходных файлов должны различаться")
    if any(p in protected for p in outputs):
        raise ValueError("Путь результата совпадает с входным файлом или шаблоном")


def _run(args):
    if args.command in {"build", "extract", "draft"} and args.output is None:
        filename = (f"{args.input.stem}.pptx" if args.command == "build"
                    else {"extract": "sources.json", "draft": "draft.json"}[args.command])
        args.output = OUTPUT_DIR / filename
    if args.command == "serve":
        from .web import create_server
        import threading
        import webbrowser
        if not 1 <= args.port <= 65535:
            raise ValueError("Порт должен быть от 1 до 65535")
        server = create_server(args.host, args.port, args.config)
        address = f"http://{args.host}:{server.server_address[1]}"
        print(f"SuperSlide: {address}", flush=True)
        print("Для остановки нажмите Ctrl+C.", flush=True)
        if args.open:
            threading.Timer(0.3, lambda: webbrowser.open(address)).start()
        try:
            server.serve_forever()
        except KeyboardInterrupt:
            print("\nSuperSlide остановлен.")
        finally:
            server.server_close()
    elif args.command == "models":
        if args.config:
            settings = load_llm_settings(args.config)
            aliases = list_models(settings.base_url, settings.api_key_env, settings.timeout_seconds,
                                  trust_env_proxy=settings.trust_env_proxy and not args.no_proxy)
        else:
            aliases = list_models(args.base_url, args.api_key_env, trust_env_proxy=not args.no_proxy)
        print(_json_text({"models": aliases}), end="")
    elif args.command == "schema":
        if _source_mode(args.manifest):
            from .source_template import source_presentation_schema
            schema = source_presentation_schema()
        else:
            schema = presentation_schema()
        if args.output:
            _write_json(args.output, schema)
            print(f"JSON Schema: {args.output.resolve()}")
        else:
            print(_json_text(schema), end="")
    elif args.command == "inspect-template":
        if _source_mode(args.manifest):
            from .source_template import inspect_source_template
            result = inspect_source_template(args.template, args.manifest)
        else:
            result = inspect_template(args.template, args.manifest)
        print(_json_text(result), end="")
    elif args.command == "assets":
        photo_path = args.asset_catalog or SOURCE_PHOTO_CATALOG
        if photo_path == SOURCE_PHOTO_CATALOG:
            print(_json_text({"photos": _source_catalog(photo_path, "photos"),
                              "icons": _source_catalog(args.icon_catalog, "icons")}), end="")
        else:
            catalog = _assets(photo_path)
            print(_json_text(catalog.to_dict() if catalog else {"images": []}), end="")
    elif args.command == "validate":
        sources = _source_lookup(args.sources)
        if _read_json(args.input).get("template_version") == "2.0":
            spec = _source_load(args.input, args.manifest, sources,
                                args.asset_catalog or SOURCE_PHOTO_CATALOG, args.icon_catalog)
        else:
            spec = load_presentation(args.input, set(sources) if sources is not None else None)
            _validate_assets(spec, _assets(args.asset_catalog))
        print(f"JSON корректен: {len(spec.slides)} слайдов, шаблон {spec.template_version}")
    elif args.command == "build":
        report_path = args.report or args.output.with_suffix(".report.json")
        _check_paths([args.output, report_path], [args.input, args.template, args.manifest, args.sources, args.asset_catalog])
        sources = _source_lookup(args.sources)
        if _source_mode(args.manifest):
            from .source_template import build_source_presentation
            photo_path = args.asset_catalog or SOURCE_PHOTO_CATALOG
            _check_paths([args.output, report_path], [photo_path, args.icon_catalog])
            spec = _source_load(args.input, args.manifest, sources, photo_path, args.icon_catalog)
            report = build_source_presentation(spec, args.template, args.manifest, args.output,
                                               source_lookup=sources, asset_catalog_path=photo_path,
                                               icon_catalog_path=args.icon_catalog)
            _write_json(report_path, report)
        else:
            spec = load_presentation(args.input, set(sources) if sources is not None else None)
            catalog = _assets(args.asset_catalog)
            _check_paths([args.output, report_path], _catalog_paths(catalog))
            report = build_presentation(spec, args.template, args.manifest, args.output,
                                        source_lookup=sources, report_path=report_path, asset_catalog=catalog)
        print(f"Создано: {args.output.resolve()} ({len(spec.slides)} слайдов)")
        print(f"Отчёт: {report_path.resolve()}")
        for warning in report.get("warnings", []):
            print(f"Проверка: {warning}")
    elif args.command == "extract":
        _check_paths([args.output], args.documents)
        bundle = load_documents(args.documents, max_total_chars=args.max_total_chars)
        _write_json(args.output, bundle.to_dict())
        print(f"Извлечено {len(bundle.chunks)} фрагментов: {args.output.resolve()}")
        for warning in bundle.warnings:
            print(f"Проверка: {warning}")
    elif args.command == "draft":
        sources_path = args.sources_output or args.output.with_suffix(".sources.json")
        generation_path = args.output.with_suffix(".generation.json")
        error_path = args.output.with_suffix(".error.json")
        _check_paths([args.output, sources_path, generation_path, error_path],
                     [*args.documents, args.task_file, args.config, args.template, args.manifest, args.asset_catalog])
        # Check all local prerequisites before sending documents to the configured endpoint.
        source_mode = _source_mode(args.manifest)
        if source_mode:
            from .source_template import (inspect_source_template, load_source_manifest,
                                          validate_source_presentation, source_presentation_schema,
                                          source_prompt_catalog, check_source_text_fit)
            inspect_source_template(args.template, args.manifest)
            manifest = load_source_manifest(args.manifest)
            photo_path = args.asset_catalog or SOURCE_PHOTO_CATALOG
            photos = _source_catalog(photo_path, "photos")
            icons = _source_catalog(args.icon_catalog, "icons")
            _check_paths([args.output, sources_path, generation_path, error_path], [photo_path, args.icon_catalog])
            catalog = None
        else:
            _, manifest, _ = load_template(args.template, args.manifest)
            catalog = _assets(args.asset_catalog)
            _check_paths([args.output, sources_path, generation_path, error_path], _catalog_paths(catalog))
        settings = load_llm_settings(args.config)
        task = args.task
        if args.task_file:
            if args.task_file.stat().st_size > 40_000:
                raise ValueError("Файл задания превышает 40 КБ")
            task = args.task_file.read_text(encoding="utf-8-sig")
        bundle = load_documents(args.documents, max_total_chars=args.max_total_chars) if args.documents else DocumentBundle(chunks=[], warnings=[])
        known = {chunk.id for chunk in bundle.chunks}
        _write_json(sources_path, bundle.to_dict())
        try:
            def validate_with_fit(value):
                if source_mode:
                    spec = _validate_source_assets(validate_source_presentation(value, manifest, known),
                                                   photo_path, args.icon_catalog, manifest)
                    check_source_text_fit(spec, args.template, args.manifest)
                    return spec
                spec = _validate_designs(_validate_assets(validate_presentation(value, known), catalog), manifest)
                try:
                    # Dry-run render: overlong slide text is reported back so the model can shorten it.
                    check_text_fit(spec, args.template, args.manifest, asset_catalog=catalog)
                except LayoutOverflowError as exc:
                    raise ValueError(
                        "Текст не помещается в поля слайдов. Сократи или суммаризуй текст проблемных полей, "
                        "сохранив ключевые факты, числа, source_refs и число слайдов.\n" + str(exc)
                    ) from None
                return spec

            result = generate_draft(
                task, bundle, settings, validate=validate_with_fit,
                schema=source_presentation_schema() if source_mode else presentation_schema(),
                layout_catalog=source_prompt_catalog(manifest) if source_mode else manifest["layouts"],
                image_catalog=photos if source_mode else catalog.prompt_catalog() if catalog else [],
                icon_catalog=icons if source_mode else [],
                template_version=manifest["template_version"],
            )
        except LLMError as exc:
            _write_json(error_path, {"status": "failed", "code": exc.code,
                                     "message": str(exc), "attempts": exc.diagnostics})
            print(f"Диагностика: {error_path.resolve()}", file=sys.stderr)
            raise
        _write_json(args.output, result.draft)
        _write_json(generation_path, {"status": "drafted", "attempts": result.attempts,
                                      "validation_errors": result.validation_errors,
                                      "source_chunks": len(bundle.chunks)})
        error_path.unlink(missing_ok=True)
        print(f"Черновик: {args.output.resolve()}; попыток: {result.attempts}")
        print(f"Источники: {sources_path.resolve()}")
        print("Проверьте содержание JSON и выполните build для сборки PPTX.")


def main(argv=None):
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8")
    args = parser().parse_args(argv)
    try:
        _run(args)
        return 0
    except ValidationError as exc:
        # Show locations and constraints, without echoing full documents or model content.
        errors = [{"loc": list(e["loc"]), "msg": e["msg"]} for e in exc.errors(include_input=False, include_url=False)]
        print("Ошибка структуры JSON:\n" + _json_text(errors), file=sys.stderr)
    except (ConfigurationError, DocumentError, LLMError, ValueError, OSError, KeyError) as exc:
        print(f"Ошибка: {exc}", file=sys.stderr)
    return 2
