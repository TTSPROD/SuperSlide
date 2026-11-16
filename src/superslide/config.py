"""Explicit configuration for an internal OpenAI-compatible text endpoint."""

from __future__ import annotations

import json
import os
from dataclasses import dataclass, fields
from pathlib import Path
from urllib.parse import urlsplit


class ConfigurationError(ValueError):
    pass


@dataclass(frozen=True)
class LLMSettings:
    base_url: str
    model: str
    api_key_env: str = "SUPERSLIDE_LLM_API_KEY"
    timeout_seconds: float = 120.0
    max_output_tokens: int = 6000
    response_format: str = "json_object"
    temperature: float | None = 0.2
    max_repair_attempts: int = 2
    max_prompt_chars: int = 220_000
    context_window_tokens: int | None = None
    token_limit_parameter: str = "max_tokens"
    reasoning_effort: str | None = None
    trust_env_proxy: bool = True

    def __post_init__(self) -> None:
        if not isinstance(self.base_url, str) or not isinstance(self.model, str) or not self.model.strip():
            raise ConfigurationError("Укажите base_url и model корпоративной LLM.")
        try:
            url = urlsplit(self.base_url)
            valid_url = url.scheme in {"http", "https"} and url.hostname and url.port != 0
        except ValueError:
            valid_url = False
            url = None
        if not valid_url or url is None or url.username or url.password or url.query or url.fragment:
            raise ConfigurationError("base_url должен быть HTTP(S) URL без логина, пароля, query и fragment.")
        if not isinstance(self.api_key_env, str) or not self.api_key_env or not self.api_key_env.replace("_", "a").isalnum():
            raise ConfigurationError("api_key_env должен содержать имя переменной окружения.")
        for name in ("max_output_tokens", "max_prompt_chars"):
            value = getattr(self, name)
            if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
                raise ConfigurationError(f"{name} должен быть положительным целым числом.")
        if isinstance(self.timeout_seconds, bool) or not isinstance(self.timeout_seconds, (int, float)) or not 0 < self.timeout_seconds <= 3600:
            raise ConfigurationError("timeout_seconds должен быть от 0 до 3600 секунд.")
        if self.response_format not in {"json_schema", "json_object", "none"}:
            raise ConfigurationError("response_format: json_schema, json_object или none.")
        if isinstance(self.max_repair_attempts, bool) or not isinstance(self.max_repair_attempts, int) or not 0 <= self.max_repair_attempts <= 2:
            raise ConfigurationError("max_repair_attempts должен быть от 0 до 2.")
        if self.context_window_tokens is not None and (
            isinstance(self.context_window_tokens, bool) or not isinstance(self.context_window_tokens, int)
            or self.context_window_tokens <= self.max_output_tokens
        ):
            raise ConfigurationError("context_window_tokens должен превышать max_output_tokens.")
        if self.temperature is not None and (
            isinstance(self.temperature, bool) or not isinstance(self.temperature, (int, float))
            or not 0 <= self.temperature <= 2
        ):
            raise ConfigurationError("temperature должен быть от 0 до 2 либо null.")
        if self.token_limit_parameter not in {"max_tokens", "max_completion_tokens"}:
            raise ConfigurationError("token_limit_parameter: max_tokens или max_completion_tokens.")
        if self.reasoning_effort is not None and self.reasoning_effort not in {"none", "minimal", "low", "medium", "high", "xhigh"}:
            raise ConfigurationError("Неподдерживаемое значение reasoning_effort.")
        if not isinstance(self.trust_env_proxy, bool):
            raise ConfigurationError("trust_env_proxy должен быть true или false.")


def load_llm_settings(config_path: str | Path | None = None) -> LLMSettings:
    values: dict = {}
    if config_path is not None:
        try:
            values = json.loads(Path(config_path).read_text(encoding="utf-8-sig"))
        except (OSError, UnicodeError, json.JSONDecodeError) as exc:
            raise ConfigurationError("Не удалось прочитать JSON-конфигурацию LLM.") from exc
        if not isinstance(values, dict):
            raise ConfigurationError("Конфигурация LLM должна быть JSON-объектом.")
        allowed = {item.name for item in fields(LLMSettings)}
        if set(values) - allowed:
            raise ConfigurationError("Конфигурация LLM содержит неподдерживаемые поля. Секрет храните в окружении.")
    for name, env_name in (("base_url", "SUPERSLIDE_LLM_BASE_URL"), ("model", "SUPERSLIDE_LLM_MODEL")):
        if os.environ.get(env_name):
            values[name] = os.environ[env_name]
    if not values.get("base_url") or not values.get("model"):
        raise ConfigurationError("Укажите base_url/model в конфигурации или SUPERSLIDE_LLM_BASE_URL/SUPERSLIDE_LLM_MODEL.")
    try:
        return LLMSettings(**values)
    except TypeError as exc:
        raise ConfigurationError("Проверьте типы полей конфигурации LLM.") from exc
