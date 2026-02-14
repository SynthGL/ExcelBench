from __future__ import annotations

import re
from pathlib import Path


def _extract_rust_umya_methods() -> set[str]:
    umya_dir = Path("rust/excelbench_rust/src/umya")
    methods: set[str] = set()
    for path in sorted(umya_dir.glob("*.rs")):
        txt = path.read_text(encoding="utf-8")
        for m in re.finditer(r"(?m)^\s*pub\s+fn\s+([A-Za-z_][A-Za-z0-9_]*)\s*\(", txt):
            methods.add(m.group(1))
    return methods


def _iter_python_code_blocks(md_text: str) -> list[str]:
    blocks: list[str] = []
    for m in re.finditer(r"```(?:python)?\s*\n(.*?)\n```", md_text, flags=re.DOTALL):
        code = m.group(1)
        if "UmyaBook" in code:
            blocks.append(code)
    return blocks


def test_pyumya_docs_only_reference_existing_umya_methods() -> None:
    rust_methods = _extract_rust_umya_methods()
    assert rust_methods, "Failed to discover any UmyaBook methods from rust sources"

    docs_dir = Path("docs/pyumya/docs")
    offenders: dict[str, set[str]] = {}

    for md_path in sorted(docs_dir.rglob("*.md")):
        md = md_path.read_text(encoding="utf-8")
        referenced: set[str] = set()

        for code in _iter_python_code_blocks(md):
            # Find variables that are UmyaBook instances.
            vars_: set[str] = set()
            vars_.update(re.findall(r"(?m)^\s*(\w+)\s*=\s*UmyaBook\s*\(", code))
            vars_.update(re.findall(r"(?m)^\s*(\w+)\s*=\s*UmyaBook\.open\s*\(", code))

            # Also treat `book` as a conventional instance name if present.
            if re.search(r"\bbook\s*\.", code):
                vars_.add("book")

            for var in vars_:
                for meth in re.findall(
                    rf"\b{re.escape(var)}\.([A-Za-z_][A-Za-z0-9_]*)\s*\(",
                    code,
                ):
                    referenced.add(meth)

            for meth in re.findall(r"\bUmyaBook\.([A-Za-z_][A-Za-z0-9_]*)\s*\(", code):
                referenced.add(meth)

        missing = {m for m in referenced if m not in rust_methods and m != "__init__"}
        if missing:
            offenders[str(md_path)] = missing

    assert not offenders, f"Docs reference missing UmyaBook methods: {offenders}"
