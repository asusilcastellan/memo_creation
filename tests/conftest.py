"""Every test uses its own temporary cache folder (never %LOCALAPPDATA%\\CastellanMemo)."""
import pytest


@pytest.fixture(autouse=True)
def _tmp_cache(tmp_path, monkeypatch):
    cache = tmp_path / "_cache"
    monkeypatch.setenv("CASTELLAN_CACHE_DIR", str(cache))
    yield cache
