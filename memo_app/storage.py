"""Read a deal folder either from the local disk (Dropbox desktop sync) or via the Dropbox API.

Both back-ends expose the same operations:
    walk(root)      -> every file/folder under the deal folder, with modified dates
    fetch(entry)    -> a local Path to the file contents (downloaded + cached for Dropbox)
    list_dir(path)  -> the direct children of one folder only (cheap; used for the prior-deal search)
    parent(path)    -> the folder that contains `path`
"""
from __future__ import annotations

import hashlib
import os
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path, PurePosixPath
from typing import Callable

from .settings import Settings

Log = Callable[[str], None]


@dataclass
class FileEntry:
    path: str            # full path in the storage back-end
    rel: str             # path relative to the deal root, using "/" separators
    name: str
    is_dir: bool
    modified: datetime   # what Explorer shows as "Date modified"
    size: int = 0

    @property
    def parent_rel(self) -> str:
        return str(PurePosixPath(self.rel).parent) if "/" in self.rel else "."

    @property
    def suffix(self) -> str:
        return PurePosixPath(self.name).suffix.lower()


class Storage:
    kind = "base"

    def walk(self, root: str) -> list[FileEntry]:
        raise NotImplementedError

    def fetch(self, entry: FileEntry) -> Path:
        raise NotImplementedError

    def list_dir(self, path: str) -> list[FileEntry]:
        raise NotImplementedError

    def parent(self, path: str) -> str:
        raise NotImplementedError

    def display_root(self, root: str) -> str:
        return root


# --------------------------------------------------------------------------- #
class LocalStorage(Storage):
    kind = "local"

    def walk(self, root: str) -> list[FileEntry]:
        base = Path(root)
        out: list[FileEntry] = []
        for dirpath, dirnames, filenames in os.walk(base):
            for name in dirnames + filenames:
                p = Path(dirpath) / name
                try:
                    st = p.stat()
                except OSError:
                    continue
                out.append(FileEntry(
                    path=str(p),
                    rel=p.relative_to(base).as_posix(),
                    name=name,
                    is_dir=p.is_dir(),
                    modified=datetime.fromtimestamp(st.st_mtime, tz=timezone.utc),
                    size=st.st_size if p.is_file() else 0,
                ))
        return out

    def fetch(self, entry: FileEntry) -> Path:
        return Path(entry.path)

    def list_dir(self, path: str) -> list[FileEntry]:
        out: list[FileEntry] = []
        with os.scandir(Path(path).expanduser()) as it:
            for d in it:
                try:
                    st = d.stat()
                    is_dir = d.is_dir()
                except OSError:
                    continue
                out.append(FileEntry(d.path, d.name, d.name, is_dir,
                                     datetime.fromtimestamp(st.st_mtime, tz=timezone.utc), 0 if is_dir else st.st_size))
        return out

    def parent(self, path: str) -> str:
        return str(Path(path).expanduser().resolve().parent)


# --------------------------------------------------------------------------- #
class DropboxStorage(Storage):
    kind = "dropbox"

    def __init__(self, settings: Settings, log: Log = print):
        import dropbox  # imported lazily so local-only use does not need network

        self.log = log
        if settings.dropbox_refresh_token and settings.dropbox_app_key:
            dbx = dropbox.Dropbox(
                oauth2_refresh_token=settings.dropbox_refresh_token,
                app_key=settings.dropbox_app_key,
                app_secret=settings.dropbox_app_secret or None,
                timeout=120,
            )
        elif settings.dropbox_access_token:
            dbx = dropbox.Dropbox(oauth2_access_token=settings.dropbox_access_token, timeout=120)
        else:
            raise RuntimeError(
                "No Dropbox credentials. Set DROPBOX_REFRESH_TOKEN + DROPBOX_APP_KEY + DROPBOX_APP_SECRET "
                "(recommended) or DROPBOX_ACCESS_TOKEN in the .env file."
            )
        if settings.dropbox_use_team_root:
            # Team spaces: list paths relative to the team root, like the desktop app shows them.
            try:
                acct = dbx.users_get_current_account()
                root_ns = acct.root_info.root_namespace_id
                dbx = dbx.with_path_root(dropbox.common.PathRoot.root(root_ns))
            except Exception as exc:  # pragma: no cover - depends on account type
                log(f"  (Dropbox team root not applied: {exc})")
        self.dbx = dbx
        from .cache import downloads_dir
        self.cache = downloads_dir()             # cleaned up after debug.cache_retention_days

    @staticmethod
    def normalise(root: str) -> str:
        root = root.strip().replace("\\", "/")
        if not root.startswith("/"):
            root = "/" + root
        return root.rstrip("/")

    def walk(self, root: str) -> list[FileEntry]:
        import dropbox

        root = self.normalise(root)
        out: list[FileEntry] = []
        res = self.dbx.files_list_folder(root, recursive=True, limit=2000)
        while True:
            for m in res.entries:
                if m.path_display.rstrip("/").lower() == root.lower():
                    continue
                rel = m.path_display[len(root):].lstrip("/")
                if isinstance(m, dropbox.files.FolderMetadata):
                    out.append(FileEntry(m.path_display, rel, m.name, True,
                                         datetime.min.replace(tzinfo=timezone.utc)))
                elif isinstance(m, dropbox.files.FileMetadata):
                    mod = m.client_modified or m.server_modified
                    out.append(FileEntry(m.path_display, rel, m.name, False,
                                         mod.replace(tzinfo=timezone.utc), m.size))
            if not res.has_more:
                break
            res = self.dbx.files_list_folder_continue(res.cursor)
        return out

    def list_dir(self, path: str) -> list[FileEntry]:
        import dropbox

        out: list[FileEntry] = []
        res = self.dbx.files_list_folder(self.normalise(path), recursive=False, limit=2000)
        while True:
            for m in res.entries:
                if isinstance(m, dropbox.files.FolderMetadata):
                    out.append(FileEntry(m.path_display, m.name, m.name, True, datetime.min.replace(tzinfo=timezone.utc)))
                elif isinstance(m, dropbox.files.FileMetadata):
                    mod = m.client_modified or m.server_modified
                    out.append(FileEntry(m.path_display, m.name, m.name, False, mod.replace(tzinfo=timezone.utc), m.size))
            if not res.has_more:
                break
            res = self.dbx.files_list_folder_continue(res.cursor)
        return out

    def parent(self, path: str) -> str:
        return str(PurePosixPath(self.normalise(path)).parent)

    def download_path(self, path: str) -> Path:
        """Download a single Dropbox file (given its full path) to the cache and return it."""
        p = self.normalise(path)
        md = self.dbx.files_get_metadata(p)
        mod = (getattr(md, "client_modified", None) or getattr(md, "server_modified", None)
               or datetime.min.replace(tzinfo=timezone.utc))
        entry = FileEntry(md.path_display, md.name, md.name, False,
                          mod.replace(tzinfo=timezone.utc), getattr(md, "size", 0))
        return self.fetch(entry)

    def fetch(self, entry: FileEntry) -> Path:
        key = hashlib.sha1(f"{entry.path}|{entry.modified.isoformat()}|{entry.size}".encode()).hexdigest()[:16]
        local = self.cache / f"{key}{entry.suffix}"
        if not local.exists():
            self.log(f"  downloading {entry.rel}")
            self.dbx.files_download_to_file(str(local), entry.path)
        else:
            os.utime(local)                     # recently used -> kept by the cache cleanup
        return local


# --------------------------------------------------------------------------- #
def open_storage(root: str, settings: Settings, log: Log = print) -> Storage:
    """A path that exists on this computer is read locally; anything else goes to the Dropbox API."""
    if root and Path(root).expanduser().is_dir():
        return LocalStorage()
    if not settings.has_dropbox:
        raise RuntimeError(
            f"'{root}' is not a folder on this computer and no Dropbox API credentials are configured."
        )
    return DropboxStorage(settings, log)
