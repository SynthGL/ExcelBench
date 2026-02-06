"""Excel library adapters."""

from excelbench.harness.adapters.base import ExcelAdapter, ReadOnlyAdapter, WriteOnlyAdapter
from excelbench.harness.adapters.openpyxl_adapter import OpenpyxlAdapter

try:
    from excelbench.harness.adapters.xlsxwriter_adapter import XlsxwriterAdapter
except ImportError:  # Optional dependency
    XlsxwriterAdapter = None
try:
    from excelbench.harness.adapters.calamine_adapter import CalamineAdapter
except ImportError:
    CalamineAdapter = None
try:
    from excelbench.harness.adapters.pylightxl_adapter import PylightxlAdapter
except ImportError:
    PylightxlAdapter = None
try:
    from excelbench.harness.adapters.xlrd_adapter import XlrdAdapter
except ImportError:
    XlrdAdapter = None
try:
    from excelbench.harness.adapters.pyexcel_adapter import PyexcelAdapter
except ImportError:
    PyexcelAdapter = None
try:
    from excelbench.harness.adapters.xlwt_adapter import XlwtAdapter
except ImportError:
    XlwtAdapter = None
from excelbench.harness.adapters.xlwings_oracle_adapter import ExcelOracleAdapter

__all__ = [
    "ExcelAdapter",
    "ReadOnlyAdapter",
    "WriteOnlyAdapter",
    "OpenpyxlAdapter",
    "ExcelOracleAdapter",
]
if XlsxwriterAdapter is not None:
    __all__.append("XlsxwriterAdapter")
if CalamineAdapter is not None:
    __all__.append("CalamineAdapter")
if PylightxlAdapter is not None:
    __all__.append("PylightxlAdapter")
if XlrdAdapter is not None:
    __all__.append("XlrdAdapter")
if PyexcelAdapter is not None:
    __all__.append("PyexcelAdapter")
if XlwtAdapter is not None:
    __all__.append("XlwtAdapter")


def get_all_adapters() -> list[ExcelAdapter]:
    """Get all available adapters."""
    adapters: list[ExcelAdapter] = [OpenpyxlAdapter()]
    if XlsxwriterAdapter is not None:
        adapters.append(XlsxwriterAdapter())
    if CalamineAdapter is not None:
        adapters.append(CalamineAdapter())
    if PylightxlAdapter is not None:
        adapters.append(PylightxlAdapter())
    if XlrdAdapter is not None:
        adapters.append(XlrdAdapter())
    if PyexcelAdapter is not None:
        adapters.append(PyexcelAdapter())
    if XlwtAdapter is not None:
        adapters.append(XlwtAdapter())
    return adapters
