"""The HTTP client used for the NY Fed / FRED / Microsoft Graph calls.

openai<3 installs `httpx`; openai>=3 installs the API-compatible fork `httpx2`. Import whichever is present
so no extra dependency is needed:   from .net import httpx
"""
from __future__ import annotations

try:
    import httpx  # noqa: F401
except ImportError:  # pragma: no cover - depends on the installed openai version
    import httpx2 as httpx  # type: ignore[no-redef]  # noqa: F401
