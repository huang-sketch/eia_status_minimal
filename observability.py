import json
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional


class RunLogger:
    def __init__(self, root: Path = Path("runs"), enabled: bool = True) -> None:
        self.enabled = enabled
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.run_dir = root / timestamp
        if self.enabled:
            self.run_dir.mkdir(parents=True, exist_ok=True)

    def log_chunk(self, chunk: Dict[str, Any], details: Optional[Dict[str, Any]] = None) -> None:
        payload = {
            "chunk_id": chunk.get("chunk_id"),
            "source_file": chunk.get("source_file"),
            "kind": chunk.get("kind"),
            "char_count": len(chunk.get("text") or ""),
            "preview": (chunk.get("text") or "")[:300],
        }
        if details:
            payload.update(details)
        self.write_jsonl("chunks.jsonl", payload)

    def log_part(self, chunk: Dict[str, Any], details: Optional[Dict[str, Any]] = None) -> None:
        self.log_chunk(chunk, details=details)

    def log_llm_call(self, event: Dict[str, Any]) -> None:
        self.write_jsonl("llm_calls.jsonl", event)

    def log_failed_chunk(
        self,
        chunk: Dict[str, Any],
        stage: str,
        error: Exception | str,
        fallback_used: bool = False,
    ) -> None:
        self.write_jsonl(
            "failed_chunks.jsonl",
            {
                "chunk_id": chunk.get("chunk_id"),
                "source_file": chunk.get("source_file"),
                "stage": stage,
                "error_type": type(error).__name__ if isinstance(error, Exception) else "Error",
                "error_message": str(error),
                "fallback_used": fallback_used,
            },
        )

    def log_warning(
        self,
        chunk: Dict[str, Any],
        warning: str,
        details: Optional[Dict[str, Any]] = None,
    ) -> None:
        payload = {
            "chunk_id": chunk.get("chunk_id"),
            "source_file": chunk.get("source_file"),
            "warning": warning,
        }
        if details:
            payload.update(details)
        self.write_jsonl("warnings.jsonl", payload)

    def write_jsonl(self, filename: str, payload: Dict[str, Any]) -> None:
        if not self.enabled:
            return
        path = self.run_dir / filename
        with path.open("a", encoding="utf-8") as file:
            file.write(json.dumps(payload, ensure_ascii=False) + "\n")
