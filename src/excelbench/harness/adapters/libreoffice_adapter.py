"""Adapter for LibreOffice Calc driven headlessly through its UNO API.

The helper runs a Python macro inside ``soffice --headless`` (LibreOffice's
own embedded Python), so every read and write goes through Calc's import and
export filters exactly as a server-side LibreOffice conversion would.
"""

from __future__ import annotations

import os
import shutil
import sys
from pathlib import Path

from excelbench.harness.adapters.json_model_adapter import JsonModelAdapter
from excelbench.harness.external_oracles import ExternalOracleTool

_HELPER_DIR = Path(__file__).resolve().parents[4] / "tools" / "external-oracles" / "libreoffice"


def soffice_binary() -> str | None:
    """Resolve the soffice executable the helper will launch."""
    for candidate in (
        os.environ.get("LIBREOFFICE_BIN"),
        shutil.which("soffice"),
        shutil.which("libreoffice"),
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ):
        if candidate and Path(candidate).exists():
            return candidate
    return None


class LibreofficeAdapter(JsonModelAdapter):
    """LibreOffice Calc read/write/modify via headless UNO."""

    LIBRARY_NAME = "libreoffice"
    LANGUAGE = "cli"
    TIMEOUT_SECONDS = 180.0

    @classmethod
    def tool(cls) -> ExternalOracleTool:
        return ExternalOracleTool(
            name="libreoffice-uno",
            command=(sys.executable, str(_HELPER_DIR / "libreoffice_uno_adapter.py")),
            language="cli",
            homepage="https://www.libreoffice.org/",
            capabilities=frozenset({"read", "write", "modify"}),
            cwd=_HELPER_DIR,
            required_paths=(_HELPER_DIR / "uno" / "excelbench_uno.py",),
        )

    @classmethod
    def is_available(cls) -> bool:
        return super().is_available() and soffice_binary() is not None
