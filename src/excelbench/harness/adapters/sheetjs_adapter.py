"""Adapter for SheetJS Community Edition (npm ``xlsx``) via a Node helper."""

from __future__ import annotations

from pathlib import Path

from excelbench.harness.adapters.json_model_adapter import JsonModelAdapter
from excelbench.harness.external_oracles import ExternalOracleTool

_HELPER_DIR = Path(__file__).resolve().parents[4] / "tools" / "external-oracles" / "sheetjs"


class SheetjsAdapter(JsonModelAdapter):
    """SheetJS CE read/write; modify means read the package, then write it back."""

    LIBRARY_NAME = "sheetjs"
    LANGUAGE = "javascript"

    @classmethod
    def tool(cls) -> ExternalOracleTool:
        return ExternalOracleTool(
            name="sheetjs",
            command=("node", str(_HELPER_DIR / "sheetjs-adapter.cjs")),
            language="node",
            homepage="https://docs.sheetjs.com/",
            capabilities=frozenset({"read", "write", "modify"}),
            cwd=_HELPER_DIR,
            required_paths=(
                _HELPER_DIR / "sheetjs-adapter.cjs",
                _HELPER_DIR / "node_modules" / "xlsx" / "package.json",
            ),
        )
