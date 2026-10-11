"""Local files: bounded sources inside attachment_root/upload_root and atomic downloads in node storage."""

from datetime import datetime, timezone
import hashlib
import os
from pathlib import Path
import re
import secrets
import shutil
import stat

from . import logic

DOWNLOAD_FOLDER = re.compile(r"[0-9]{8}T[0-9]{6}\.[0-9]{6}Z-[0-9a-f]{8}\Z")


class LocalFile:
    """One opened regular file; reads are positional and the file must not change while it is sent."""

    def __init__(self, fd, info, name):
        self.fd, self.size, self.name = fd, info.st_size, name
        self.identity = (info.st_dev, info.st_ino, info.st_size, info.st_mtime_ns)

    def read(self, offset, length):
        chunks, done = [], 0
        while done < length:
            piece = os.pread(self.fd, length - done, offset + done)
            if not piece:
                logic.fail("path", "The local file changed while it was read; nothing more was sent.")
            chunks.append(piece)
            done += len(piece)
        return b"".join(chunks)

    def read_all(self):
        data = self.read(0, self.size)
        self.check()
        return data

    def check(self):
        info = os.fstat(self.fd)
        if (info.st_dev, info.st_ino, info.st_size, info.st_mtime_ns) != self.identity or os.pread(self.fd, 1, self.size):
            logic.fail("path", "The local file changed while it was read; nothing more was sent.")

    def sha256(self):
        digest, offset = hashlib.sha256(), 0
        while offset < self.size:
            piece = self.read(offset, min(1048576, self.size - offset))
            digest.update(piece)
            offset += len(piece)
        return digest.hexdigest()

    def close(self):
        if self.fd is not None:
            os.close(self.fd)
            self.fd = None


def open_inside(value, root_value, limit, *, root_code, too_large):
    """Open one regular file inside a configured root without following a swapped final symlink."""
    if not root_value:
        logic.fail(root_code)
    try:
        root = Path(root_value).resolve(strict=True)
    except (OSError, RuntimeError):
        logic.fail("path", "The configured root folder does not exist; nothing was sent.")
    if not root.is_dir():
        logic.fail("path", "The configured root must be a directory; nothing was sent.")
    candidate = Path(value).expanduser()
    try:
        resolved = (candidate if candidate.is_absolute() else root / candidate).resolve(strict=True)
    except (OSError, RuntimeError):
        logic.fail("path", "The local file does not exist; nothing was sent.")
    if not resolved.is_relative_to(root):
        logic.fail("path_denied")
    try:
        fd = os.open(resolved, os.O_RDONLY | os.O_NOFOLLOW | os.O_CLOEXEC | os.O_NONBLOCK)
    except OSError:
        logic.fail("path", "The local file cannot be opened; nothing was sent.")
    try:
        info = os.fstat(fd)
        current = os.stat(resolved, follow_symlinks=False)
        if not stat.S_ISREG(info.st_mode) or (current.st_dev, current.st_ino) != (info.st_dev, info.st_ino):
            logic.fail("path", "Only regular files can be sent; nothing was sent.")
        if info.st_size > limit:
            logic.fail(too_large)
        return LocalFile(fd, info, resolved.name)
    except BaseException:
        os.close(fd)
        raise


def read_attachments(paths, config):
    """Read mail attachments inside attachment_root: each and all together bounded before any request."""
    result, total = [], 0
    for value in paths:
        local = open_inside(value, config["attachment_root"], config["max_attachment_bytes"],
                            root_code="attachment_root_required", too_large="file_too_large")
        try:
            total += local.size
            if total > logic.SEND_ATTACHMENTS_LIMIT:
                logic.fail("file_too_large", "Attachments exceed 3 MiB in total, the inline sendMail limit; nothing was sent.")
            result.append((local.name, local.read_all()))
        finally:
            local.close()
    return result


def open_upload(value, config):
    return open_inside(value, config["upload_root"], config["max_upload_bytes"], root_code="upload_root_required",
                       too_large="file_too_large")


def storage_directory(storage):
    """Resolve the node storage directory before any request, so a missing scope sends nothing."""
    if not callable(storage):
        logic.fail("storage")
    try:
        directory = Path(storage())
    except Exception:
        raise logic.GraphError("storage") from None
    if not directory.is_absolute() or directory.is_symlink() or not directory.is_dir():
        logic.fail("storage", "The block storage directory is not a real directory.")
    return directory


def reserve(directory):
    """Create a unique folder under downloads/ and an exclusive temporary file inside it."""
    downloads = directory / "downloads"
    try:
        downloads.mkdir(mode=0o700, exist_ok=True)
    except OSError:
        logic.fail("storage", "The downloads folder cannot be created.")
    if downloads.is_symlink() or not downloads.is_dir():
        logic.fail("storage", "The downloads folder is not a real directory.")
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%S.%fZ")
    folder = downloads / (stamp + "-" + secrets.token_hex(4))
    try:
        folder.mkdir(mode=0o700)
    except OSError:
        logic.fail("storage", "The download folder cannot be created.")
    temporary = folder / (".partial-" + secrets.token_hex(4))
    try:
        fd = os.open(temporary, os.O_RDWR | os.O_CREAT | os.O_EXCL | os.O_NOFOLLOW | os.O_CLOEXEC, 0o600)
    except OSError:
        shutil.rmtree(folder, ignore_errors=True)
        logic.fail("storage", "The download file cannot be created.")
    return {"folder": folder, "temporary": temporary, "fd": fd, "downloads": downloads}


def finish(target, name, keep):
    """Close, rename atomically, sync the folder, then prune older download folders."""
    os.fsync(target["fd"])
    os.close(target["fd"])
    target["fd"] = None
    final = target["folder"] / name
    os.replace(target["temporary"], final)
    directory = os.open(target["folder"], os.O_RDONLY | os.O_DIRECTORY)
    try:
        os.fsync(directory)
    finally:
        os.close(directory)
    prune(target["downloads"], keep, target["folder"])
    return final


def discard(target):
    """Remove a reserved download after any failure; never leaves a partial file behind."""
    if target.get("fd") is not None:
        try:
            os.close(target["fd"])
        except OSError:
            pass
        target["fd"] = None
    shutil.rmtree(target["folder"], ignore_errors=True)


def prune(downloads, keep, current):
    """Keep the current and newest download folders created by this block; never follow links."""
    others = sorted(entry for entry in downloads.iterdir() if entry != current
                    and DOWNLOAD_FOLDER.fullmatch(entry.name) and entry.is_dir() and not entry.is_symlink())
    for entry in others[:max(0, len(others) - (keep - 1))]:
        shutil.rmtree(entry, ignore_errors=True)
