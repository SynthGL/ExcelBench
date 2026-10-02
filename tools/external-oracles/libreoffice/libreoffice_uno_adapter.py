#!/usr/bin/env python3
"""Host launcher for the LibreOffice Calc JSON-model adapter helper.

Reads one ExcelBench helper request (JSON) on stdin, runs the
``uno/excelbench_uno.py`` macro inside ``soffice --headless`` (LibreOffice's
own embedded Python interpreter), and prints the macro's JSON response on
stdout. Only the Python standard library is used on the host side; every
spreadsheet operation happens inside soffice through the UNO API.

Why a macro instead of a UNO socket bridge: on macOS the bundled standalone
``Contents/Resources/python`` is SIGKILLed by the OS and importing ``uno`` from
a foreign Python segfaults. A Python macro stored in the user profile runs in
soffice's embedded interpreter and needs neither.
"""

from __future__ import annotations

import json
import os
import shutil
import signal
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any

JSONDict = dict[str, Any]

HELPER_DIR = Path(__file__).resolve().parent
MACRO_SOURCE = HELPER_DIR / "uno" / "excelbench_uno.py"
MACRO_NAME = "excelbench_uno.py"
MACRO_URL = f"vnd.sun.star.script:{MACRO_NAME}$main?language=Python&location=user"
REQUEST_ENV = "EXCELBENCH_UNO_REQUEST"
RESPONSE_ENV = "EXCELBENCH_UNO_RESPONSE"
DEFAULT_TIMEOUT_SECONDS = 150.0
STDERR_TAIL_CHARS = 4000


def main() -> int:
    try:
        request = json.load(sys.stdin)
        if not isinstance(request, dict):
            raise ValueError("request must be a JSON object")
        response = run_request(request)
    except Exception as exc:  # pragma: no cover - last-resort CLI guard
        response = {
            "error": "libreoffice_failed",
            "message": f"{type(exc).__name__}: {exc}",
        }
    print(json.dumps(response, sort_keys=True))
    return 1 if response.get("error") else 0


def resolve_soffice() -> str | None:
    """Find soffice the same way the Python adapter does."""
    for candidate in (
        os.environ.get("LIBREOFFICE_BIN"),
        shutil.which("soffice"),
        shutil.which("libreoffice"),
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ):
        if candidate and Path(candidate).exists():
            return str(candidate)
    return None


def profile_root() -> Path:
    """Return the reusable LibreOffice profile directory (outside the repo)."""
    override = os.environ.get("EXCELBENCH_LIBREOFFICE_PROFILE")
    if override:
        return Path(override).expanduser().resolve()
    cache = os.environ.get("XDG_CACHE_HOME") or str(Path.home() / ".cache")
    return Path(cache).expanduser().resolve() / "excelbench" / "libreoffice-uno" / "profile"


class ProfileLock:
    """Exclusive lock serializing soffice runs that share one user profile.

    Two soffice processes on one ``UserInstallation`` do not run side by side:
    the second hands its arguments to the first over IPC and exits, so a
    concurrent request would silently never execute. The lock prevents that.
    """

    def __init__(self, path: Path) -> None:
        self.path = path
        self._handle: Any = None

    def __enter__(self) -> ProfileLock:
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self._handle = open(self.path, "a+b")  # noqa: SIM115 - held for the lock lifetime
        if os.name == "nt":
            import msvcrt

            self._handle.seek(0)
            msvcrt.locking(self._handle.fileno(), msvcrt.LK_LOCK, 1)
        else:
            import fcntl

            fcntl.flock(self._handle.fileno(), fcntl.LOCK_EX)
        return self

    def __exit__(self, *_exc: object) -> None:
        if self._handle is None:
            return
        if os.name == "nt":
            import msvcrt

            self._handle.seek(0)
            msvcrt.locking(self._handle.fileno(), msvcrt.LK_UNLCK, 1)
        else:
            import fcntl

            fcntl.flock(self._handle.fileno(), fcntl.LOCK_UN)
        self._handle.close()
        self._handle = None


def install_macro(profile: Path) -> None:
    """Copy the macro into ``<profile>/user/Scripts/python`` for this run."""
    target_dir = profile / "user" / "Scripts" / "python"
    target_dir.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(MACRO_SOURCE, target_dir / MACRO_NAME)


def timeout_seconds() -> float:
    raw = os.environ.get("EXCELBENCH_LIBREOFFICE_TIMEOUT")
    if raw:
        try:
            return max(1.0, float(raw))
        except ValueError:
            pass
    return DEFAULT_TIMEOUT_SECONDS


def _stderr_tail(stderr: str) -> str:
    stderr = stderr.strip()
    return stderr[-STDERR_TAIL_CHARS:] if len(stderr) > STDERR_TAIL_CHARS else stderr


def _kill_group(process: subprocess.Popen[str]) -> None:
    """Kill soffice and anything it spawned so the profile is released."""
    try:
        if os.name == "nt":
            process.kill()
        else:
            os.killpg(process.pid, signal.SIGKILL)
    except (ProcessLookupError, PermissionError):
        pass


def run_request(request: JSONDict) -> JSONDict:
    if not MACRO_SOURCE.is_file():
        return {
            "error": "libreoffice_failed",
            "message": f"macro not found: {MACRO_SOURCE}",
        }
    soffice = resolve_soffice()
    if soffice is None:
        return {
            "error": "libreoffice_failed",
            "message": "LibreOffice soffice executable not found "
            "(set LIBREOFFICE_BIN or put soffice on PATH)",
        }

    profile = profile_root()
    limit = timeout_seconds()
    with tempfile.TemporaryDirectory(prefix="excelbench-uno-") as tmp:
        request_path = Path(tmp) / "request.json"
        response_path = Path(tmp) / "response.json"
        request_path.write_text(json.dumps(request), encoding="utf-8")
        env = os.environ.copy()
        env[REQUEST_ENV] = str(request_path)
        env[RESPONSE_ENV] = str(response_path)
        command = [
            soffice,
            f"-env:UserInstallation={profile.as_uri()}",
            "--headless",
            "--invisible",
            "--norestore",
            "--nologo",
            "--nodefault",
            "--nolockcheck",
            MACRO_URL,
        ]
        with ProfileLock(profile.parent / f"{profile.name}.lock"):
            install_macro(profile)
            process = subprocess.Popen(
                command,
                stdin=subprocess.DEVNULL,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                env=env,
                start_new_session=os.name != "nt",
            )
            try:
                stdout, stderr = process.communicate(timeout=limit)
            except subprocess.TimeoutExpired:
                _kill_group(process)
                stdout, stderr = process.communicate()
                return {
                    "error": "libreoffice_failed",
                    "message": f"soffice did not finish {request.get('operation')!r} "
                    f"within {limit:g}s",
                    "stderr": _stderr_tail(stderr or ""),
                }

        if response_path.is_file():
            try:
                response = json.loads(response_path.read_text(encoding="utf-8"))
            except json.JSONDecodeError as exc:
                return {
                    "error": "libreoffice_failed",
                    "message": f"macro wrote invalid JSON: {exc}",
                    "stderr": _stderr_tail(stderr or ""),
                }
            if isinstance(response, dict):
                return response
            return {
                "error": "libreoffice_failed",
                "message": "macro response is not an object",
            }

        return {
            "error": "libreoffice_failed",
            "message": "soffice exited without running the ExcelBench macro "
            f"(returncode {process.returncode})",
            "stdout": (stdout or "").strip()[-STDERR_TAIL_CHARS:],
            "stderr": _stderr_tail(stderr or ""),
        }


if __name__ == "__main__":
    raise SystemExit(main())
