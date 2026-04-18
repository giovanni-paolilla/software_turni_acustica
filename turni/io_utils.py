"""File I/O atomico e locking distribuito per scritture concorrenti."""
from __future__ import annotations
import json
import os
import tempfile
import time

from turni.constants import LOCK_STALE_SECONDS


def _pid_is_alive(pid: int) -> bool:
    if pid <= 0:
        return False
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    except OSError:
        return False
    else:
        return True


def _read_lock_metadata(lock_path: str) -> dict | None:
    try:
        with open(lock_path, encoding="utf-8") as f:
            payload = json.load(f)
    except (OSError, json.JSONDecodeError, TypeError, ValueError):
        return None
    return payload if isinstance(payload, dict) else None


def _get_process_start_token(pid: int) -> str | None:
    if pid <= 0:
        return None
    proc_stat = f"/proc/{pid}/stat"
    try:
        with open(proc_stat, encoding="utf-8") as f:
            stat = f.read()
    except OSError:
        return None

    try:
        after_comm = stat.rsplit(")", 1)[1].strip()
        fields = after_comm.split()
        return fields[19]
    except (IndexError, ValueError):
        return None


def _is_stale_lock(lock_path: str, *, stale_after_seconds: int = LOCK_STALE_SECONDS) -> bool:
    metadata = _read_lock_metadata(lock_path)
    now = time.time()

    if metadata is not None:
        pid = metadata.get("pid")
        created_at = metadata.get("created_at")
        proc_start_token = metadata.get("proc_start_token")

        if isinstance(pid, int):
            if not _pid_is_alive(pid):
                return True

            current_token = _get_process_start_token(pid)
            if (proc_start_token is not None and current_token is not None
                    and str(proc_start_token) != str(current_token)):
                return True

            return False

        if isinstance(created_at, (int, float)) and now - float(created_at) > stale_after_seconds:
            return True
        return False

    try:
        mtime = os.path.getmtime(lock_path)
    except OSError:
        return False
    return now - mtime > stale_after_seconds


class _TargetFileLock:
    """Lock file per coordinare scritture concorrenti sullo stesso target."""

    def __init__(self, target_path: str) -> None:
        self.target_path = os.path.abspath(target_path)
        self.lock_path = self.target_path + ".lock"
        self._fd: int | None = None

    def _try_acquire(self) -> bool:
        try:
            fd = os.open(self.lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        except FileExistsError:
            return False

        try:
            payload = json.dumps({
                "pid": os.getpid(),
                "created_at": time.time(),
                "proc_start_token": _get_process_start_token(os.getpid()),
            })
            os.write(fd, payload.encode())
            os.fsync(fd)
        except OSError:
            try:
                os.close(fd)
            except OSError:
                pass
            try:
                os.unlink(self.lock_path)
            except OSError:
                pass
            raise

        self._fd = fd
        return True

    def __enter__(self) -> _TargetFileLock:
        if self._try_acquire():
            return self

        if _is_stale_lock(self.lock_path):
            try:
                os.unlink(self.lock_path)
            except FileNotFoundError:
                pass
            except OSError:
                pass
            if self._try_acquire():
                return self

        raise BlockingIOError(
            f"File occupato da un'altra istanza: '{self.target_path}'."
        )

    def __exit__(self, exc_type, exc, tb) -> None:
        if self._fd is not None:
            try:
                os.close(self._fd)
            finally:
                self._fd = None
        try:
            os.unlink(self.lock_path)
        except FileNotFoundError:
            pass


def _write_text_file_atomic(path: str, data: str, newline: str | None = None) -> None:
    """Scrive testo in modo atomico tramite file temporaneo + os.replace()."""
    directory = os.path.dirname(os.path.abspath(path)) or "."
    with _TargetFileLock(path):
        fd, tmp_path = tempfile.mkstemp(prefix=".tmp_turni_", dir=directory)
        try:
            with os.fdopen(fd, "w", encoding="utf-8", newline=newline) as f:
                f.write(data)
                f.flush()
                os.fsync(f.fileno())
            os.replace(tmp_path, path)
        except OSError:
            try:
                os.unlink(tmp_path)
            except OSError:
                pass
            raise
