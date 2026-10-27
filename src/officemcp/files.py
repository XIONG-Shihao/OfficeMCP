"""Where OfficeMCP may read and write, file revisions, and writes that land whole or not at all."""

import hashlib
import json
import os
import shutil
import tempfile
from dataclasses import dataclass, field
from pathlib import Path

from .errors import Refusal

FORMATS = {".docx": "docx", ".xlsx": "xlsx", ".pptx": "pptx", ".pdf": "pdf"}
LIBREOFFICE_PATHS = (
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/opt/homebrew/bin/soffice",
    "/usr/local/bin/soffice",
    "/usr/bin/soffice",
)


@dataclass
class Config:
    """What a host allows: the folders OfficeMCP may read and write, optional engines and limits."""

    roots: list[Path]
    libreoffice: Path | None = None
    max_bytes: int = 200 * 1024 * 1024
    max_cells: int = 5000
    extra: dict = field(default_factory=dict)

    @classmethod
    def load(cls, path: str | os.PathLike) -> "Config":
        data = json.loads(Path(path).read_text(encoding="utf-8"))
        return cls.from_dict(data)

    @classmethod
    def from_dict(cls, data: dict) -> "Config":
        roots = data.get("roots")
        if not isinstance(roots, list) or not roots or not all(isinstance(r, str) and r for r in roots):
            raise ValueError("The configuration needs roots: a list of folders")
        engines = data.get("engines") or {}
        limits = data.get("limits") or {}
        office = engines.get("libreOffice")
        found = Path(office) if office else next((Path(p) for p in LIBREOFFICE_PATHS if Path(p).is_file()), None)
        return cls(
            roots=[Path(r).resolve(strict=True) for r in roots],
            libreoffice=found if found and found.is_file() else None,
            max_bytes=int(limits.get("maxBytes", cls.max_bytes)),
            max_cells=int(limits.get("maxCells", cls.max_cells)),
            extra=data,
        )

    def _inside(self, path: str, resolved: Path) -> Path:
        if not any(resolved.is_relative_to(root) for root in self.roots):
            raise Refusal("path_not_allowed", f"{path} is outside the folders OfficeMCP may use")
        return resolved

    def input(self, path: str) -> Path:
        """An existing file inside a root. A relative path is relative to the first root."""
        if not isinstance(path, str) or not path:
            raise Refusal("invalid_request", "A path is required")
        candidate = Path(path)
        if not candidate.is_absolute():
            candidate = self.roots[0] / candidate
        try:
            resolved = self._inside(path, candidate.resolve(strict=True))
        except FileNotFoundError:
            raise Refusal("not_found", f"No file at {path}") from None
        if not resolved.is_file():
            raise Refusal("not_found", f"{path} is not a file")
        if resolved.stat().st_size > self.max_bytes:
            raise Refusal("too_large", f"{path} is larger than {self.max_bytes} bytes")
        return resolved

    def new_file(self, path: str, extension: str) -> Path:
        """A path for a new file inside a root, in an existing folder; an existing file is refused."""
        if not isinstance(path, str) or not path or "\x00" in path:
            raise Refusal("invalid_request", "A path is required")
        candidate = Path(path)
        if not candidate.is_absolute():
            candidate = self.roots[0] / candidate
        if candidate.suffix.lower() != extension:
            raise Refusal("invalid_request", f"The new file's name must end in {extension}")
        try:
            folder = self._inside(path, candidate.parent.resolve(strict=True))
        except FileNotFoundError:
            raise Refusal("not_found", f"The folder for {path} does not exist") from None
        target = folder / candidate.name
        if target.exists() or target.is_symlink():
            raise Refusal("output_exists", f"{path} already exists", hint="Choose another path, or edit that file")
        return target

    def beside(self, source: Path, name: str) -> Path:
        """A new file next to the source; a taken name is numbered."""
        stem, suffix = os.path.splitext(name)
        target, n = source.parent / name, 2
        while target.exists() or target.is_symlink():
            target, n = source.parent / f"{stem}-{n}{suffix}", n + 1
        return target


def format_of(path: Path) -> str:
    kind = FORMATS.get(path.suffix.lower())
    if not kind:
        raise Refusal("unsupported_format", f"{path.name} is not a PDF, DOCX, XLSX or PPTX file")
    return kind


def revision(path: Path) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return "sha256:" + digest.hexdigest()


def _staged(draft: Path, target: Path) -> Path:
    """A copy of the draft beside the target, on its file system, ready to be renamed."""
    staged = target.with_name(f".{target.name}.officemcp-{os.urandom(6).hex()}")
    shutil.copyfile(draft, staged)
    with open(staged, "rb+") as handle:
        os.fsync(handle.fileno())
    return staged


def write_new(draft: Path, target: Path) -> None:
    """Put a finished draft at a new path; fails if the path was taken meanwhile."""
    staged = _staged(draft, target)
    try:
        os.link(staged, target)
    except FileExistsError:
        raise Refusal("output_exists", f"{target.name} appeared meanwhile") from None
    finally:
        staged.unlink(missing_ok=True)


def replace(draft: Path, target: Path, expected: str) -> None:
    """Replace a file with a finished draft in one rename, if it is still at the expected revision."""
    staged = _staged(draft, target)
    try:
        os.chmod(staged, target.stat().st_mode & 0o7777)
        if revision(target) != expected:
            raise Refusal("revision_mismatch", "The file changed while the edit was prepared; nothing was written",
                          hint="Inspect or read it again", revision=revision(target))
        os.replace(staged, target)
    finally:
        staged.unlink(missing_ok=True)


class Scratch:
    """A private temporary folder in the system's temporary folder."""

    def __init__(self, config: Config):
        self.path = Path(tempfile.mkdtemp(prefix="officemcp-"))

    def file(self, name: str) -> Path:
        return self.path / name

    def __enter__(self):
        return self

    def __exit__(self, *_):
        shutil.rmtree(self.path, ignore_errors=True)
