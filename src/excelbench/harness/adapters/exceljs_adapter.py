"""Adapter for ExcelJS (npm ``exceljs``) via the ExcelJS external helper."""

from __future__ import annotations

from pathlib import Path

from excelbench.harness.adapters.json_model_adapter import JsonModelAdapter
from excelbench.harness.external_oracles import (
    ExternalOracleTool,
    external_oracle_catalog,
)


class ExceljsAdapter(JsonModelAdapter):
    """ExcelJS read/write/modify through ``tools/external-oracles/exceljs``."""

    LIBRARY_NAME = "exceljs"
    LANGUAGE = "javascript"

    @classmethod
    def tool(cls) -> ExternalOracleTool:
        return external_oracle_catalog(repo_root=Path(__file__).resolve().parents[4])["exceljs"]
