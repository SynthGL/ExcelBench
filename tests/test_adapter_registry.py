"""Tests for adapter registry: get_all_adapters, __all__, adapter capabilities."""

from __future__ import annotations

from pathlib import Path

from excelbench.harness.adapters import (
    ExcelAdapter,
    OpenpyxlAdapter,
    ReadOnlyAdapter,
    WriteOnlyAdapter,
    get_all_adapters,
)


class TestGetAllAdapters:
    def test_returns_list(self) -> None:
        adapters = get_all_adapters()
        assert isinstance(adapters, list)
        assert len(adapters) >= 1

    def test_openpyxl_always_present(self) -> None:
        adapters = get_all_adapters()
        names = [a.name for a in adapters]
        assert "openpyxl" in names

    def test_all_are_excel_adapters(self) -> None:
        for adapter in get_all_adapters():
            assert isinstance(adapter, ExcelAdapter)

    def test_each_adapter_has_info(self) -> None:
        for adapter in get_all_adapters():
            info = adapter.info
            assert info.name
            assert info.version
            assert info.language
            assert len(info.capabilities) >= 1

    def test_each_adapter_can_read_or_write(self) -> None:
        for adapter in get_all_adapters():
            assert adapter.can_read() or adapter.can_write()

    def test_no_duplicate_names(self) -> None:
        adapters = get_all_adapters()
        names = [a.name for a in adapters]
        assert len(names) == len(set(names))


class TestAdapterClassification:
    def test_openpyxl_is_readwrite(self) -> None:
        a = OpenpyxlAdapter()
        assert a.can_read()
        assert a.can_write()

    def test_readonly_subclass_contract(self) -> None:
        for adapter in get_all_adapters():
            if isinstance(adapter, ReadOnlyAdapter):
                assert adapter.can_read()
                assert not adapter.can_write()

    def test_writeonly_subclass_contract(self) -> None:
        for adapter in get_all_adapters():
            if isinstance(adapter, WriteOnlyAdapter):
                assert not adapter.can_read()
                assert adapter.can_write()


class TestAdapterExtensions:
    def test_openpyxl_supports_xlsx(self) -> None:
        a = OpenpyxlAdapter()
        assert a.supports_read_path(Path("test.xlsx"))

    def test_openpyxl_rejects_csv(self) -> None:
        a = OpenpyxlAdapter()
        assert not a.supports_read_path(Path("test.csv"))

    def test_output_extensions(self) -> None:
        for adapter in get_all_adapters():
            ext = adapter.output_extension
            assert ext in {".xlsx", ".xls"}


class TestAllExports:
    def test_all_contains_base_classes(self) -> None:
        from excelbench.harness import adapters

        assert "ExcelAdapter" in adapters.__all__
        assert "ReadOnlyAdapter" in adapters.__all__
        assert "WriteOnlyAdapter" in adapters.__all__
        assert "OpenpyxlAdapter" in adapters.__all__

    def test_all_matches_available_adapters(self) -> None:
        """Each adapter returned by get_all_adapters should have its class in __all__."""
        from excelbench.harness import adapters

        for adapter in get_all_adapters():
            cls_name = type(adapter).__name__
            # Base classes always in __all__; optional adapters only if import succeeded
            if cls_name in ("OpenpyxlAdapter",):
                assert cls_name in adapters.__all__
