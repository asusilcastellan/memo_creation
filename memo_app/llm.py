"""Thin wrapper around the OpenAI Responses API.

    extract(schema, instructions, text, images)  -> pydantic object (structured outputs)
    write(system, user)                         -> plain text
    web_search(prompt, schema)                  -> pydantic object, using the hosted web_search tool

Swap models in resources/config/settings.yaml or with OPENAI_MODEL_EXTRACT / OPENAI_MODEL_WRITE in .env.
"""
from __future__ import annotations

import threading
from typing import Callable, Type, TypeVar

from openai import OpenAI
from pydantic import BaseModel

from .settings import Settings

T = TypeVar("T", bound=BaseModel)
Log = Callable[[str], None]


class LLM:
    def __init__(self, settings: Settings, log: Log = print):
        self.s = settings
        self.log = log
        self.client = OpenAI(
            api_key=settings.openai_api_key,
            timeout=float(settings.get("openai.request_timeout_seconds", 300)),
            max_retries=int(settings.get("openai.max_retries", 3)),
        )
        self.usage = {"input_tokens": 0, "output_tokens": 0, "calls": 0}
        self._lock = threading.Lock()

    # ------------------------------------------------------------------ helpers
    @staticmethod
    def _reasoning(effort: str | None) -> dict:
        return {"reasoning": {"effort": effort}} if effort else {}

    def _track(self, resp) -> None:
        u = getattr(resp, "usage", None)
        with self._lock:
            self.usage["calls"] += 1
            if u is not None:
                self.usage["input_tokens"] += getattr(u, "input_tokens", 0) or 0
                self.usage["output_tokens"] += getattr(u, "output_tokens", 0) or 0

    @staticmethod
    def _content(text: str | None, images_b64: list[str] | None) -> list[dict]:
        content: list[dict] = []
        if text:
            content.append({"type": "input_text", "text": text})
        for b64 in images_b64 or []:
            content.append({"type": "input_image", "image_url": f"data:image/png;base64,{b64}", "detail": "high"})
        return content

    # ------------------------------------------------------------------ API
    def extract(self, schema: Type[T], instructions: str, text: str | None = None,
                images_b64: list[str] | None = None) -> T:
        resp = self.client.responses.parse(
            model=self.s.extract_model,
            instructions=instructions,
            input=[{"role": "user", "content": self._content(text, images_b64)}],
            text_format=schema,
            **self._reasoning(self.s.get("openai.extract_reasoning_effort")),
        )
        self._track(resp)
        if resp.output_parsed is None:
            raise RuntimeError(f"Model returned no parsable {schema.__name__}")
        return resp.output_parsed

    def write(self, system: str, user: str) -> str:
        resp = self.client.responses.create(
            model=self.s.write_model,
            instructions=system,
            input=user,
            **self._reasoning(self.s.get("openai.write_reasoning_effort")),
        )
        self._track(resp)
        return (resp.output_text or "").strip()

    def web_search(self, schema: Type[T], prompt: str) -> T:
        resp = self.client.responses.parse(
            model=self.s.extract_model,
            input=prompt,
            tools=[{"type": "web_search"}],
            text_format=schema,
        )
        self._track(resp)
        return resp.output_parsed
