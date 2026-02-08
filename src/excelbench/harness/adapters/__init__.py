"""Excel library adapters."""

from typing import TypeAlias

from excelbench.harness.adapters.base import ExcelAdapter, ReadOnlyAdapter, WriteOnlyAdapter
from excelbench.harness.adapters.openpyxl_adapter import OpenpyxlAdapter

AdapterClass: TypeAlias = type[ExcelAdapter]

try:
    from excelbench.harness.adapters.xlsxwriter_adapter import (
        XlsxwriterAdapter as _XlsxwriterAdapter,
    )
except ImportError:  # Optional dependency
    XlsxwriterAdapter: AdapterClass | None = None
else:
    XlsxwriterAdapter = _XlsxwriterAdapter
try:
    from excelbench.harness.adapters.calamine_adapter import CalamineAdapter as _CalamineAdapter
except ImportError:
    CalamineAdapter: AdapterClass | None = None
else:
    CalamineAdapter = _CalamineAdapter

try:
    from excelbench.harness.adapters.rust_calamine_adapter import (
        RustCalamineAdapter as _RustCalamineAdapter,
    )
except ImportError:
    RustCalamineAdapter: AdapterClass | None = None
else:
    RustCalamineAdapter = _RustCalamineAdapter

try:
    from excelbench.harness.adapters.rust_xlsxwriter_adapter import (
        RustXlsxWriterAdapter as _RustXlsxWriterAdapter,
    )
except ImportError:
    RustXlsxWriterAdapter: AdapterClass | None = None
else:
    RustXlsxWriterAdapter = _RustXlsxWriterAdapter

try:
    from excelbench.harness.adapters.umya_adapter import UmyaAdapter as _UmyaAdapter
except ImportError:
    UmyaAdapter: AdapterClass | None = None
else:
    UmyaAdapter = _UmyaAdapter

try:
    from excelbench.harness.adapters.pylightxl_adapter import PylightxlAdapter as _PylightxlAdapter
except ImportError:
    PylightxlAdapter: AdapterClass | None = None
else:
    PylightxlAdapter = _PylightxlAdapter
try:
    from excelbench.harness.adapters.xlrd_adapter import XlrdAdapter as _XlrdAdapter
except ImportError:
    XlrdAdapter: AdapterClass | None = None
else:
    XlrdAdapter = _XlrdAdapter
try:
    from excelbench.harness.adapters.pyexcel_adapter import PyexcelAdapter as _PyexcelAdapter
except ImportError:
    PyexcelAdapter: AdapterClass | None = None
else:
    PyexcelAdapter = _PyexcelAdapter
try:
    from excelbench.harness.adapters.xlwt_adapter import XlwtAdapter as _XlwtAdapter
except ImportError:
    XlwtAdapter: AdapterClass | None = None
else:
    XlwtAdapter = _XlwtAdapter
try:
    from excelbench.harness.adapters.pandas_adapter import PandasAdapter as _PandasAdapter
except ImportError:
    PandasAdapter: AdapterClass | None = None
else:
    PandasAdapter = _PandasAdapter

try:
    from excelbench.harness.adapters.xlsxwriter_constmem_adapter import (
        XlsxwriterConstmemAdapter as _XlsxwriterConstmemAdapter,
    )
except ImportError:
    XlsxwriterConstmemAdapter: AdapterClass | None = None
else:
    XlsxwriterConstmemAdapter = _XlsxwriterConstmemAdapter
try:
    from excelbench.harness.adapters.openpyxl_readonly_adapter import (
        OpenpyxlReadonlyAdapter as _OpenpyxlReadonlyAdapter,
    )
except ImportError:
    OpenpyxlReadonlyAdapter: AdapterClass | None = None
else:
    OpenpyxlReadonlyAdapter = _OpenpyxlReadonlyAdapter
try:
    from excelbench.harness.adapters.polars_adapter import PolarsAdapter as _PolarsAdapter
except ImportError:
    PolarsAdapter: AdapterClass | None = None
else:
    PolarsAdapter = _PolarsAdapter
try:
    from excelbench.harness.adapters.tablib_adapter import TablibAdapter as _TablibAdapter
except ImportError:
    TablibAdapter: AdapterClass | None = None
else:
    TablibAdapter = _TablibAdapter

try:
    from excelbench.harness.adapters.xlwings_oracle_adapter import (
        ExcelOracleAdapter as _ExcelOracleAdapter,
    )
except ImportError:
    ExcelOracleAdapter: AdapterClass | None = None
else:
    ExcelOracleAdapter = _ExcelOracleAdapter

__all__ = [
    "ExcelAdapter",
    "ReadOnlyAdapter",
    "WriteOnlyAdapter",
    "OpenpyxlAdapter",
]
if ExcelOracleAdapter is not None:
    __all__.append("ExcelOracleAdapter")
if XlsxwriterAdapter is not None:
    __all__.append("XlsxwriterAdapter")
if CalamineAdapter is not None:
    __all__.append("CalamineAdapter")
if RustCalamineAdapter is not None:
    __all__.append("RustCalamineAdapter")
if RustXlsxWriterAdapter is not None:
    __all__.append("RustXlsxWriterAdapter")
if UmyaAdapter is not None:
    __all__.append("UmyaAdapter")
if PylightxlAdapter is not None:
    __all__.append("PylightxlAdapter")
if XlrdAdapter is not None:
    __all__.append("XlrdAdapter")
if PyexcelAdapter is not None:
    __all__.append("PyexcelAdapter")
if XlwtAdapter is not None:
    __all__.append("XlwtAdapter")
if PandasAdapter is not None:
    __all__.append("PandasAdapter")
if XlsxwriterConstmemAdapter is not None:
    __all__.append("XlsxwriterConstmemAdapter")
if OpenpyxlReadonlyAdapter is not None:
    __all__.append("OpenpyxlReadonlyAdapter")
if PolarsAdapter is not None:
    __all__.append("PolarsAdapter")
if TablibAdapter is not None:
    __all__.append("TablibAdapter")


def get_all_adapters() -> list[ExcelAdapter]:
    """Get all available adapters."""
    adapters: list[ExcelAdapter] = [OpenpyxlAdapter()]
    if XlsxwriterAdapter is not None:
        adapters.append(XlsxwriterAdapter())
    if CalamineAdapter is not None:
        adapters.append(CalamineAdapter())
    if RustCalamineAdapter is not None:
        adapters.append(RustCalamineAdapter())
    if RustXlsxWriterAdapter is not None:
        adapters.append(RustXlsxWriterAdapter())
    if UmyaAdapter is not None:
        adapters.append(UmyaAdapter())
    if PylightxlAdapter is not None:
        adapters.append(PylightxlAdapter())
    if XlrdAdapter is not None:
        adapters.append(XlrdAdapter())
    if PyexcelAdapter is not None:
        adapters.append(PyexcelAdapter())
    if XlwtAdapter is not None:
        adapters.append(XlwtAdapter())
    if PandasAdapter is not None:
        adapters.append(PandasAdapter())
    if XlsxwriterConstmemAdapter is not None:
        adapters.append(XlsxwriterConstmemAdapter())
    if OpenpyxlReadonlyAdapter is not None:
        adapters.append(OpenpyxlReadonlyAdapter())
    if PolarsAdapter is not None:
        adapters.append(PolarsAdapter())
    if TablibAdapter is not None:
        adapters.append(TablibAdapter())
    return adapters
