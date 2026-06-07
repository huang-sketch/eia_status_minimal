import json
import http.client
import os
import random
import re
import ssl
import time
import urllib.error
import urllib.request
from enum import Enum
from typing import Any, Callable, Dict, List, Optional, TypeVar

import certifi

DEFAULT_ENDPOINT = "https://api.openai.com/v1/chat/completions"
DEFAULT_MODEL = "gpt-4o-mini"
LLM_TIMEOUT_SECONDS = int(os.getenv("EIA_LLM_TIMEOUT_SECONDS", "25"))
LLM_TEXT_TIMEOUT_SECONDS = int(os.getenv("EIA_LLM_TEXT_TIMEOUT_SECONDS", "120"))
LLM_MAX_RETRIES = int(os.getenv("EIA_LLM_MAX_RETRIES", "3"))
LLM_TEXT_MAX_RETRIES = int(os.getenv("EIA_LLM_TEXT_MAX_RETRIES", "8"))
LLM_RETRY_DELAY_SECONDS = 2
LLM_TEXT_RETRY_DELAY_SECONDS = int(os.getenv("EIA_LLM_TEXT_RETRY_DELAY_SECONDS", "4"))
LLM_TEXT_RECOVERY_ATTEMPTS = int(os.getenv("EIA_LLM_TEXT_RECOVERY_ATTEMPTS", "3"))
LLM_TEXT_RECOVERY_DELAY_SECONDS = int(os.getenv("EIA_LLM_TEXT_RECOVERY_DELAY_SECONDS", "15"))
LLM_TEXT_BATCH_PAUSE_SECONDS = float(os.getenv("EIA_LLM_TEXT_BATCH_PAUSE_SECONDS", "2"))
SSL_CONTEXT = ssl.create_default_context(cafile=certifi.where())

RETRIABLE_HTTP_STATUS_CODES = {429, 500, 502, 503, 504}


def is_network_error(exc: BaseException) -> bool:
    if isinstance(exc, (urllib.error.URLError, TimeoutError, ConnectionResetError, ConnectionAbortedError)):
        return True
    if isinstance(exc, http.client.RemoteDisconnected):
        return True
    if isinstance(exc, RuntimeError) and "failed after all retries" in str(exc):
        nested = str(exc).split(": ", 1)[-1]
        return is_network_error_from_message(nested)
    return is_network_error_from_message(str(exc))


def is_network_error_from_message(text: str) -> bool:
    lowered = str(text or "").lower()
    markers = (
        "10054",
        "10060",
        "connection reset",
        "remote host",
        "remote disconnected",
        "timed out",
        "timeout",
        "connection aborted",
        "broken pipe",
        "temporarily unavailable",
        "connection refused",
    )
    return any(marker in lowered for marker in markers)


def describe_llm_error(exc: BaseException) -> str:
    if is_network_error(exc):
        return "LLM 网络连接异常，重试后仍失败，已使用规则文本"
    return f"LLM 调用失败，已使用规则文本：{exc}"


def build_rule_text_fallback_validation(exc: BaseException) -> Dict[str, Any]:
    return {
        "used_llm": True,
        "valid": True,
        "llm_applied": False,
        "fallback": "rule_texts",
        "error_type": "network" if is_network_error(exc) else "llm_error",
        "warnings": [describe_llm_error(exc)],
    }
RETRIABLE_ERRORS = (
    urllib.error.URLError,
    http.client.RemoteDisconnected,
    http.client.IncompleteRead,
    ConnectionResetError,
    ConnectionAbortedError,
    BrokenPipeError,
    OSError,
    TimeoutError,
    json.JSONDecodeError,
    KeyError,
    ValueError,
)


class LlmProfile(str, Enum):
    extraction = "extraction"
    text_polish = "text_polish"
    project_overview = "project_overview"
    fast = "fast"


T = TypeVar("T")


def resolve_max_retries(profile: LlmProfile, max_retries: Optional[int]) -> int:
    if max_retries is not None:
        return max(1, max_retries)
    if profile in (LlmProfile.text_polish, LlmProfile.project_overview):
        return max(1, LLM_TEXT_MAX_RETRIES)
    return max(1, LLM_MAX_RETRIES)


def retry_delay_seconds(attempt: int, profile: LlmProfile) -> float:
    if profile in (LlmProfile.text_polish, LlmProfile.project_overview):
        delay = LLM_TEXT_RETRY_DELAY_SECONDS * (2 ** (attempt - 1))
        delay = min(delay, 60)
        return delay + random.uniform(0, 2)
    return LLM_RETRY_DELAY_SECONDS * attempt


def llm_enabled() -> bool:
    return bool(os.getenv("EIA_LLM_API_KEY")) and os.getenv("EIA_LLM_DISABLE") != "1"


def resolve_profile_settings(
    profile: LlmProfile,
    *,
    model: Optional[str] = None,
    temperature: Optional[float] = None,
    max_tokens: Optional[int] = None,
) -> Dict[str, Any]:
    if profile == LlmProfile.extraction:
        resolved_model = model or os.getenv("EIA_LLM_MODEL", DEFAULT_MODEL)
        resolved_temperature = 0 if temperature is None else temperature
        resolved_max_tokens = int(
            max_tokens if max_tokens is not None else os.getenv("EIA_LLM_MAX_TOKENS", "2048")
        )
    elif profile == LlmProfile.text_polish:
        resolved_model = model or os.getenv("EIA_LLM_MODEL", "qwen-plus")
        resolved_temperature = 0.2 if temperature is None else temperature
        resolved_max_tokens = int(
            max_tokens if max_tokens is not None else os.getenv("EIA_LLM_TEXT_MAX_TOKENS", "1800")
        )
    elif profile == LlmProfile.project_overview:
        resolved_model = model or os.getenv("EIA_LLM_MODEL", "qwen-plus")
        resolved_temperature = 0.2 if temperature is None else temperature
        resolved_max_tokens = int(
            max_tokens
            if max_tokens is not None
            else os.getenv("EIA_PROJECT_OVERVIEW_MAX_TOKENS", "3500")
        )
    elif profile == LlmProfile.fast:
        resolved_model = model or os.getenv("EIA_LLM_FAST_MODEL", "qwen-flash")
        resolved_temperature = 0 if temperature is None else temperature
        resolved_max_tokens = int(
            max_tokens if max_tokens is not None else os.getenv("EIA_LLM_FAST_MAX_TOKENS", "4096")
        )
    else:
        raise ValueError(f"Unsupported LLM profile: {profile}")

    return {
        "model": resolved_model,
        "temperature": resolved_temperature,
        "max_tokens": resolved_max_tokens,
    }


def resolve_request_timeout(profile: LlmProfile, timeout: Optional[int]) -> int:
    if timeout is not None:
        return timeout
    if profile in (LlmProfile.text_polish, LlmProfile.project_overview):
        return LLM_TEXT_TIMEOUT_SECONDS
    return LLM_TIMEOUT_SECONDS


def chat_completion(
    messages: List[Dict[str, str]],
    *,
    profile: LlmProfile = LlmProfile.extraction,
    model: Optional[str] = None,
    temperature: Optional[float] = None,
    max_tokens: Optional[int] = None,
    timeout: Optional[int] = None,
    max_retries: Optional[int] = None,
    label: str = "llm",
    logger: Optional[Any] = None,
    logger_context: Optional[Dict[str, Any]] = None,
) -> str:
    return _chat_completion_with_retry(
        messages,
        profile=profile,
        model=model,
        temperature=temperature,
        max_tokens=max_tokens,
        timeout=timeout,
        max_retries=max_retries,
        label=label,
        logger=logger,
        logger_context=logger_context,
        parser=None,
    )


def chat_completion_json_object(
    messages: List[Dict[str, str]],
    *,
    profile: LlmProfile = LlmProfile.text_polish,
    model: Optional[str] = None,
    temperature: Optional[float] = None,
    max_tokens: Optional[int] = None,
    timeout: Optional[int] = None,
    max_retries: Optional[int] = None,
    label: str = "llm_json_object",
    logger: Optional[Any] = None,
    logger_context: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    return _chat_completion_with_retry(
        messages,
        profile=profile,
        model=model,
        temperature=temperature,
        max_tokens=max_tokens,
        timeout=timeout,
        max_retries=max_retries,
        label=label,
        logger=logger,
        logger_context=logger_context,
        parser=parse_json_object,
    )


def chat_completion_json_object_with_recovery(
    messages: List[Dict[str, str]],
    *,
    profile: LlmProfile = LlmProfile.text_polish,
    model: Optional[str] = None,
    temperature: Optional[float] = None,
    max_tokens: Optional[int] = None,
    timeout: Optional[int] = None,
    max_retries: Optional[int] = None,
    label: str = "llm_json_object",
    recovery_attempts: Optional[int] = None,
    logger: Optional[Any] = None,
    logger_context: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    total_recovery = max(1, recovery_attempts if recovery_attempts is not None else LLM_TEXT_RECOVERY_ATTEMPTS)
    last_error: Optional[BaseException] = None
    for round_idx in range(1, total_recovery + 1):
        round_label = label if round_idx == 1 else f"{label}_recovery_{round_idx}"
        try:
            return chat_completion_json_object(
                messages,
                profile=profile,
                model=model,
                temperature=temperature,
                max_tokens=max_tokens,
                timeout=timeout,
                max_retries=max_retries,
                label=round_label,
                logger=logger,
                logger_context=logger_context,
            )
        except (RuntimeError, urllib.error.URLError, TimeoutError, OSError) as exc:
            last_error = exc
            if round_idx >= total_recovery or not is_network_error(exc):
                raise
            delay = LLM_TEXT_RECOVERY_DELAY_SECONDS * round_idx
            print(
                f"{label} network recovery wait {delay:.0f}s "
                f"({round_idx}/{total_recovery})",
                flush=True,
            )
            time.sleep(delay)
    if last_error:
        raise last_error
    raise RuntimeError(f"{label} failed after recovery attempts")


def chat_completion_json_array(
    messages: List[Dict[str, str]],
    *,
    profile: LlmProfile = LlmProfile.extraction,
    model: Optional[str] = None,
    temperature: Optional[float] = None,
    max_tokens: Optional[int] = None,
    timeout: Optional[int] = None,
    max_retries: Optional[int] = None,
    label: str = "llm_json_array",
    logger: Optional[Any] = None,
    logger_context: Optional[Dict[str, Any]] = None,
) -> List[Dict[str, Any]]:
    return _chat_completion_with_retry(
        messages,
        profile=profile,
        model=model,
        temperature=temperature,
        max_tokens=max_tokens,
        timeout=timeout,
        max_retries=max_retries,
        label=label,
        logger=logger,
        logger_context=logger_context,
        parser=parse_json_array,
    )


def _chat_completion_with_retry(
    messages: List[Dict[str, str]],
    *,
    profile: LlmProfile,
    model: Optional[str],
    temperature: Optional[float],
    max_tokens: Optional[int],
    timeout: Optional[int],
    max_retries: Optional[int],
    label: str,
    logger: Optional[Any],
    logger_context: Optional[Dict[str, Any]],
    parser: Optional[Callable[[str], T]],
) -> Any:
    settings = resolve_profile_settings(
        profile,
        model=model,
        temperature=temperature,
        max_tokens=max_tokens,
    )
    endpoint = os.getenv("EIA_LLM_ENDPOINT", DEFAULT_ENDPOINT)
    request_timeout = resolve_request_timeout(profile, timeout)
    total_attempts = resolve_max_retries(profile, max_retries)
    payload = {
        "model": settings["model"],
        "temperature": settings["temperature"],
        "max_tokens": settings["max_tokens"],
        "messages": messages,
    }

    last_error: Optional[Exception] = None
    for attempt in range(1, total_attempts + 1):
        started = time.time()
        request = urllib.request.Request(
            endpoint,
            data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            headers={
                "Authorization": f"Bearer {os.getenv('EIA_LLM_API_KEY')}",
                "Content-Type": "application/json",
                "Accept": "application/json",
                "User-Agent": "eia-status-minimal/1.0",
            },
            method="POST",
        )
        try:
            print(f"{label} attempt {attempt}/{total_attempts}", flush=True)
            with urllib.request.urlopen(
                request,
                timeout=request_timeout,
                context=SSL_CONTEXT,
            ) as response:
                data = json.loads(response.read().decode("utf-8"))
            content = data["choices"][0]["message"]["content"]
            result: Any = parser(content) if parser else content
            event = {
                "attempt": attempt,
                "model": settings["model"],
                "success": True,
                "latency_ms": int((time.time() - started) * 1000),
                "error": None,
            }
            if isinstance(result, list):
                event["record_count"] = len(result)
            if logger_context:
                event.update(logger_context)
            if logger:
                logger.log_llm_call(event)
            print(f"{label} succeeded", flush=True)
            return result
        except urllib.error.HTTPError as exc:
            if exc.code not in RETRIABLE_HTTP_STATUS_CODES:
                raise
            last_error = exc
            event = {
                "attempt": attempt,
                "model": settings["model"],
                "success": False,
                "latency_ms": int((time.time() - started) * 1000),
                "record_count": 0,
                "error": str(exc),
            }
            if logger_context:
                event.update(logger_context)
            if logger:
                logger.log_llm_call(event)
            print(f"{label} failed on attempt {attempt}/{total_attempts}: {exc}", flush=True)
            if attempt < total_attempts:
                time.sleep(retry_delay_seconds(attempt, profile))
        except RETRIABLE_ERRORS as exc:
            last_error = exc
            event = {
                "attempt": attempt,
                "model": settings["model"],
                "success": False,
                "latency_ms": int((time.time() - started) * 1000),
                "record_count": 0,
                "error": str(exc),
            }
            if logger_context:
                event.update(logger_context)
            if logger:
                logger.log_llm_call(event)
            print(f"{label} failed on attempt {attempt}/{total_attempts}: {exc}", flush=True)
            if attempt < total_attempts:
                time.sleep(retry_delay_seconds(attempt, profile))

    raise RuntimeError(f"{label} failed after all retries: {last_error}")


def parse_json_object(content: str) -> Dict[str, Any]:
    text = str(content or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    start = text.find("{")
    end = text.rfind("}")
    if start == -1 or end == -1 or end <= start:
        raise ValueError("LLM response must be a JSON object")
    payload = json.loads(text[start : end + 1])
    if not isinstance(payload, dict):
        raise ValueError("LLM response JSON is not an object")
    return payload


def parse_json_array(content: str) -> List[Dict[str, Any]]:
    text = str(content or "").strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?", "", text).strip()
        text = re.sub(r"```$", "", text).strip()

    parsed = json.loads(text)
    if isinstance(parsed, dict):
        if parsed.get("contains_monitoring_data") is False:
            return []
        parsed = parsed.get("records", [])
    if not isinstance(parsed, list):
        raise ValueError("LLM response must be a JSON array or an object with records.")
    return [item for item in parsed if isinstance(item, dict)]
