"""Subprocess contract for optional external spreadsheet oracles."""

from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
from collections.abc import Mapping
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

JSONDict = dict[str, Any]


@dataclass(frozen=True)
class ExternalOracleTool:
    """Descriptor for an optional non-Python spreadsheet oracle.

    Args:
        name: Stable identifier used in reports and diagnostics.
        command: Command used to launch the helper. The first item must be an
            executable name or absolute path.
        language: Runtime ecosystem for the helper.
        homepage: Documentation or project URL for the underlying tool.
        capabilities: Coarse capability labels such as ``write`` or
            ``open_save_validate``.
        env: Extra environment variables for the subprocess.
        cwd: Optional working directory for helper commands that are checked
            into this repository.
    """

    name: str
    command: tuple[str, ...]
    language: str
    homepage: str
    capabilities: frozenset[str]
    env: Mapping[str, str] = field(default_factory=dict)
    cwd: Path | None = None

    def resolve_executable(self) -> str | None:
        """Resolve the command executable without running it."""
        if not self.command:
            return None
        executable = self.command[0]
        if os.path.isabs(executable):
            return executable if Path(executable).exists() else None
        return shutil.which(executable)

    def is_available(self) -> bool:
        """Return whether the helper command is available on this machine."""
        return self.resolve_executable() is not None

    def resolved_command(self) -> tuple[str, ...] | None:
        """Return the command with the executable resolved to an absolute path."""
        executable = self.resolve_executable()
        if executable is None:
            return None
        return (executable, *self.command[1:])


@dataclass(frozen=True)
class ExternalOracleRequest:
    """JSON request passed to an external oracle helper.

    Args:
        fixture_id: Stable fixture or scenario identifier.
        operation: Helper operation, for example ``write_fixture`` or
            ``open_save_validate``.
        payload: Operation-specific JSON payload.
        input_path: Optional workbook path for read/validate operations.
        output_path: Optional workbook path for generated output.
    """

    fixture_id: str
    operation: str
    payload: JSONDict
    input_path: Path | None = None
    output_path: Path | None = None

    def to_json_dict(self) -> JSONDict:
        """Convert the request to a JSON-serializable dictionary."""
        body: JSONDict = {
            "fixture_id": self.fixture_id,
            "operation": self.operation,
            "payload": self.payload,
        }
        if self.input_path is not None:
            body["input_path"] = str(self.input_path)
        if self.output_path is not None:
            body["output_path"] = str(self.output_path)
        return body

    def to_json(self) -> str:
        """Serialize the request for subprocess stdin."""
        return json.dumps(self.to_json_dict(), sort_keys=True)


@dataclass(frozen=True)
class ExternalOracleResult:
    """Structured result returned by ``run_external_oracle``."""

    tool_name: str
    passed: bool
    skipped: bool
    returncode: int | None
    stdout: str
    stderr: str
    payload: JSONDict
    notes: str | None = None


def external_oracle_catalog(repo_root: Path | None = None) -> dict[str, ExternalOracleTool]:
    """Return the planned external-oracle helper catalog.

    The commands are helper entrypoints, not raw runtime commands. They are
    intentionally absent from normal installs until each oracle is implemented
    and audited.
    """
    excelize_command: tuple[str, ...] = ("excelbench-excelize-oracle",)
    excelize_cwd = None
    if repo_root is not None:
        excelize_command = ("go", "run", ".")
        excelize_cwd = repo_root / "tools" / "external-oracles" / "excelize"
    libreoffice_command: tuple[str, ...] = ("excelbench-libreoffice-oracle",)
    if repo_root is not None:
        libreoffice_command = (
            sys.executable,
            str(repo_root / "tools" / "external-oracles" / "libreoffice" / "libreoffice_oracle.py"),
        )
    closedxml_command: tuple[str, ...] = ("excelbench-closedxml-oracle",)
    if repo_root is not None:
        closedxml_command = (
            "dotnet",
            "run",
            "--project",
            str(repo_root / "tools" / "external-oracles" / "closedxml" / "closedxml-oracle.csproj"),
            "--configuration",
            "Release",
            "--no-launch-profile",
            "--verbosity",
            "quiet",
            "--",
        )

    return {
        "excelize": ExternalOracleTool(
            name="excelize",
            command=excelize_command,
            language="go",
            homepage="https://github.com/qax-os/excelize",
            capabilities=frozenset({"read", "write", "charts", "pivots", "slicers"}),
            cwd=excelize_cwd,
        ),
        "libreoffice": ExternalOracleTool(
            name="libreoffice",
            command=libreoffice_command,
            language="cli",
            homepage="https://www.libreoffice.org/",
            capabilities=frozenset({"open_save_validate", "render_validate"}),
        ),
        "apache-poi": ExternalOracleTool(
            name="apache-poi",
            command=("excelbench-poi-oracle",),
            language="java",
            homepage="https://poi.apache.org/components/spreadsheet/",
            capabilities=frozenset({"read", "write", "charts", "pivots"}),
        ),
        "closedxml": ExternalOracleTool(
            name="closedxml",
            command=closedxml_command,
            language="dotnet",
            homepage="https://docs.closedxml.io/",
            capabilities=frozenset({"read", "write", "pivots", "conditional_formatting"}),
        ),
    }


def run_external_oracle(
    tool: ExternalOracleTool,
    request: ExternalOracleRequest,
    *,
    timeout_seconds: float = 60.0,
    cwd: Path | None = None,
) -> ExternalOracleResult:
    """Run an optional external oracle helper with a JSON request.

    Missing helper commands return a skipped result instead of raising. This
    keeps Go/Java/.NET/LibreOffice dependencies out of the core test suite
    while preserving a stable contract for local pre-release oracle passes.
    """
    command = tool.resolved_command()
    if command is None:
        return ExternalOracleResult(
            tool_name=tool.name,
            passed=False,
            skipped=True,
            returncode=None,
            stdout="",
            stderr="",
            payload={},
            notes=f"External oracle helper not found: {tool.command[0] if tool.command else ''}",
        )

    env = os.environ.copy()
    env.update(tool.env)
    try:
        completed = subprocess.run(
            command,
            input=request.to_json(),
            text=True,
            capture_output=True,
            check=False,
            timeout=timeout_seconds,
            cwd=cwd if cwd is not None else tool.cwd,
            env=env,
        )
    except subprocess.TimeoutExpired as exc:
        return ExternalOracleResult(
            tool_name=tool.name,
            passed=False,
            skipped=False,
            returncode=None,
            stdout=exc.stdout if isinstance(exc.stdout, str) else "",
            stderr=exc.stderr if isinstance(exc.stderr, str) else "",
            payload={},
            notes=f"External oracle timed out after {timeout_seconds:g}s",
        )

    payload = _parse_stdout_payload(completed.stdout)
    skipped = bool(payload.get("skipped", False))
    passed = completed.returncode == 0 and not skipped and not payload.get("error")
    notes = payload.get("notes")
    return ExternalOracleResult(
        tool_name=tool.name,
        passed=passed,
        skipped=skipped,
        returncode=completed.returncode,
        stdout=completed.stdout,
        stderr=completed.stderr,
        payload=payload,
        notes=notes if isinstance(notes, str) else None,
    )


def _parse_stdout_payload(stdout: str) -> JSONDict:
    """Parse helper stdout as JSON, preserving invalid output as diagnostics."""
    stripped = stdout.strip()
    if not stripped:
        return {}
    try:
        parsed = json.loads(stripped)
    except json.JSONDecodeError as exc:
        return {
            "error": "invalid_json_stdout",
            "message": str(exc),
            "stdout": stdout,
        }
    if isinstance(parsed, dict):
        return parsed
    return {"payload": parsed}
