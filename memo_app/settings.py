"""Load secrets from .env and options from resources/config/settings.yaml."""
from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import Any

import yaml
from dotenv import load_dotenv

from . import paths


@dataclass
class Settings:
    openai_api_key: str = ""
    dropbox_app_key: str = ""
    dropbox_app_secret: str = ""
    dropbox_refresh_token: str = ""
    dropbox_access_token: str = ""
    dropbox_use_team_root: bool = True
    fred_api_key: str = ""              # optional: live WSJ Prime from FRED (series DPRIME)
    graph_tenant_id: str = ""           # Microsoft Graph (Outlook email chain), app-only auth
    graph_client_id: str = ""
    graph_client_secret: str = ""
    graph_mailboxes: list[str] = field(default_factory=list)
    raw: dict[str, Any] = field(default_factory=dict)

    # convenience accessors -------------------------------------------------
    def get(self, dotted: str, default: Any = None) -> Any:
        node: Any = self.raw
        for part in dotted.split("."):
            if not isinstance(node, dict) or part not in node:
                return default
            node = node[part]
        return node

    @property
    def extract_model(self) -> str:
        return os.getenv("OPENAI_MODEL_EXTRACT") or self.get("openai.extract_model", "gpt-5.4-mini")

    @property
    def write_model(self) -> str:
        return os.getenv("OPENAI_MODEL_WRITE") or self.get("openai.write_model", "gpt-5.5")

    @property
    def has_dropbox(self) -> bool:
        return bool(self.dropbox_access_token or (self.dropbox_refresh_token and self.dropbox_app_key))

    @property
    def has_graph(self) -> bool:
        return bool(self.graph_tenant_id and self.graph_client_id and self.graph_client_secret and self.graph_mailboxes)

    def missing_secrets(self) -> list[str]:
        missing = []
        if not self.openai_api_key:
            missing.append("OPENAI_API_KEY")
        return missing


def _truthy(v: str | None, default: bool) -> bool:
    if v is None or v == "":
        return default
    return v.strip().lower() in {"1", "true", "yes", "y", "on"}


def load_settings() -> Settings:
    load_dotenv(paths.env_file(), override=False)
    load_dotenv(override=False)  # also a .env in the current working directory
    with open(paths.resource("config", "settings.yaml"), encoding="utf-8") as fh:
        raw = yaml.safe_load(fh) or {}
    return Settings(
        openai_api_key=os.getenv("OPENAI_API_KEY", "").strip(),
        dropbox_app_key=os.getenv("DROPBOX_APP_KEY", "").strip(),
        dropbox_app_secret=os.getenv("DROPBOX_APP_SECRET", "").strip(),
        dropbox_refresh_token=os.getenv("DROPBOX_REFRESH_TOKEN", "").strip(),
        dropbox_access_token=os.getenv("DROPBOX_ACCESS_TOKEN", "").strip(),
        dropbox_use_team_root=_truthy(os.getenv("DROPBOX_USE_TEAM_ROOT"), True),
        fred_api_key=os.getenv("FRED_API_KEY", "").strip(),
        graph_tenant_id=os.getenv("GRAPH_TENANT_ID", "").strip(),
        graph_client_id=os.getenv("GRAPH_CLIENT_ID", "").strip(),
        graph_client_secret=os.getenv("GRAPH_CLIENT_SECRET", "").strip(),
        graph_mailboxes=[m.strip() for m in os.getenv("GRAPH_MAILBOXES", "").split(",") if m.strip()],
        raw=raw,
    )


def load_yaml_resource(*parts: str) -> dict:
    with open(paths.resource(*parts), encoding="utf-8") as fh:
        return yaml.safe_load(fh) or {}
