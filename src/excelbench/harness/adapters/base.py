"""Base adapter protocol for Excel libraries."""

from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any

from excelbench.models import BorderInfo, CellFormat, CellValue, LibraryInfo


class ExcelAdapter(ABC):
    """Abstract base class for Excel library adapters.

    Each adapter wraps a specific Excel library and provides a
    unified interface for reading and writing cell data.
    """

    @property
    @abstractmethod
    def info(self) -> LibraryInfo:
        """Get information about this library."""
        ...

    @property
    def name(self) -> str:
        """Library name."""
        return self.info.name

    @property
    def capabilities(self) -> set[str]:
        """Library capabilities (read, write)."""
        return self.info.capabilities

    def can_read(self) -> bool:
        """Check if this adapter supports reading."""
        return "read" in self.capabilities

    def can_write(self) -> bool:
        """Check if this adapter supports writing."""
        return "write" in self.capabilities

    # =========================================================================
    # Read Operations
    # =========================================================================

    @abstractmethod
    def open_workbook(self, path: Path) -> Any:
        """Open a workbook for reading.

        Args:
            path: Path to the Excel file.

        Returns:
            Library-specific workbook object.
        """
        ...

    @abstractmethod
    def close_workbook(self, workbook: Any) -> None:
        """Close an opened workbook.

        Args:
            workbook: The workbook object to close.
        """
        ...

    @abstractmethod
    def get_sheet_names(self, workbook: Any) -> list[str]:
        """Get list of sheet names in a workbook.

        Args:
            workbook: The workbook object.

        Returns:
            List of sheet names.
        """
        ...

    @abstractmethod
    def read_cell_value(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
    ) -> CellValue:
        """Read the value of a cell.

        Args:
            workbook: The workbook object.
            sheet: Sheet name.
            cell: Cell reference (e.g., "A1", "B2").

        Returns:
            CellValue with type and value.
        """
        ...

    @abstractmethod
    def read_cell_format(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
    ) -> CellFormat:
        """Read the formatting of a cell.

        Args:
            workbook: The workbook object.
            sheet: Sheet name.
            cell: Cell reference.

        Returns:
            CellFormat with formatting properties.
        """
        ...

    @abstractmethod
    def read_cell_border(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
    ) -> BorderInfo:
        """Read the border information of a cell.

        Args:
            workbook: The workbook object.
            sheet: Sheet name.
            cell: Cell reference.

        Returns:
            BorderInfo with border properties.
        """
        ...

    @abstractmethod
    def read_row_height(
        self,
        workbook: Any,
        sheet: str,
        row: int,
    ) -> float | None:
        """Read the height of a row (1-indexed)."""
        ...

    @abstractmethod
    def read_column_width(
        self,
        workbook: Any,
        sheet: str,
        column: str,
    ) -> float | None:
        """Read the width of a column by letter (e.g., "A")."""
        ...

    # =========================================================================
    # Tier 2 Read Operations
    # =========================================================================

    @abstractmethod
    def read_merged_ranges(self, workbook: Any, sheet: str) -> list[str]:
        """Read merged cell ranges in a sheet (e.g., ["A1:C1"])."""
        ...

    @abstractmethod
    def read_conditional_formats(self, workbook: Any, sheet: str) -> list[dict]:
        """Read conditional formatting rules in a sheet."""
        ...

    @abstractmethod
    def read_data_validations(self, workbook: Any, sheet: str) -> list[dict]:
        """Read data validation rules in a sheet."""
        ...

    @abstractmethod
    def read_hyperlinks(self, workbook: Any, sheet: str) -> list[dict]:
        """Read hyperlinks in a sheet."""
        ...

    @abstractmethod
    def read_images(self, workbook: Any, sheet: str) -> list[dict]:
        """Read images/embedded objects in a sheet."""
        ...

    @abstractmethod
    def read_pivot_tables(self, workbook: Any, sheet: str) -> list[dict]:
        """Read pivot table definitions in a sheet."""
        ...

    @abstractmethod
    def read_comments(self, workbook: Any, sheet: str) -> list[dict]:
        """Read comments/notes in a sheet."""
        ...

    @abstractmethod
    def read_freeze_panes(self, workbook: Any, sheet: str) -> dict:
        """Read freeze/split pane settings in a sheet."""
        ...

    # =========================================================================
    # Write Operations
    # =========================================================================

    @abstractmethod
    def create_workbook(self) -> Any:
        """Create a new workbook.

        Returns:
            Library-specific workbook object.
        """
        ...

    @abstractmethod
    def add_sheet(self, workbook: Any, name: str) -> None:
        """Add a new sheet to a workbook.

        Args:
            workbook: The workbook object.
            name: Name for the new sheet.
        """
        ...

    @abstractmethod
    def write_cell_value(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
        value: CellValue,
    ) -> None:
        """Write a value to a cell.

        Args:
            workbook: The workbook object.
            sheet: Sheet name.
            cell: Cell reference.
            value: The value to write.
        """
        ...

    @abstractmethod
    def write_cell_format(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
        format: CellFormat,
    ) -> None:
        """Apply formatting to a cell.

        Args:
            workbook: The workbook object.
            sheet: Sheet name.
            cell: Cell reference.
            format: The formatting to apply.
        """
        ...

    @abstractmethod
    def write_cell_border(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
        border: BorderInfo,
    ) -> None:
        """Apply border to a cell.

        Args:
            workbook: The workbook object.
            sheet: Sheet name.
            cell: Cell reference.
            border: The border to apply.
        """
        ...

    @abstractmethod
    def set_row_height(
        self,
        workbook: Any,
        sheet: str,
        row: int,
        height: float,
    ) -> None:
        """Set the height of a row (1-indexed)."""
        ...

    @abstractmethod
    def set_column_width(
        self,
        workbook: Any,
        sheet: str,
        column: str,
        width: float,
    ) -> None:
        """Set the width of a column by letter (e.g., "A")."""
        ...

    # =========================================================================
    # Tier 2 Write Operations
    # =========================================================================

    @abstractmethod
    def merge_cells(self, workbook: Any, sheet: str, cell_range: str) -> None:
        """Merge a range of cells (e.g., "A1:C1")."""
        ...

    @abstractmethod
    def add_conditional_format(self, workbook: Any, sheet: str, rule: dict) -> None:
        """Add a conditional formatting rule to a sheet."""
        ...

    @abstractmethod
    def add_data_validation(self, workbook: Any, sheet: str, validation: dict) -> None:
        """Add a data validation rule to a sheet."""
        ...

    @abstractmethod
    def add_hyperlink(self, workbook: Any, sheet: str, link: dict) -> None:
        """Add a hyperlink to a sheet."""
        ...

    @abstractmethod
    def add_image(self, workbook: Any, sheet: str, image: dict) -> None:
        """Add an image/embedded object to a sheet."""
        ...

    @abstractmethod
    def add_pivot_table(self, workbook: Any, sheet: str, pivot: dict) -> None:
        """Add a pivot table to a sheet."""
        ...

    @abstractmethod
    def add_comment(self, workbook: Any, sheet: str, comment: dict) -> None:
        """Add a comment/note to a sheet."""
        ...

    @abstractmethod
    def set_freeze_panes(self, workbook: Any, sheet: str, settings: dict) -> None:
        """Set freeze/split pane settings in a sheet."""
        ...

    @abstractmethod
    def save_workbook(self, workbook: Any, path: Path) -> None:
        """Save a workbook to a file.

        Args:
            workbook: The workbook object.
            path: Path to save to.
        """
        ...


class ReadOnlyAdapter(ExcelAdapter):
    """Base class for read-only adapters.

    Provides default implementations that raise NotImplementedError
    for all write operations.
    """

    def create_workbook(self) -> Any:
        raise NotImplementedError(f"{self.name} is read-only")

    def add_sheet(self, workbook: Any, name: str) -> None:
        raise NotImplementedError(f"{self.name} is read-only")

    def write_cell_value(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
        value: CellValue,
    ) -> None:
        raise NotImplementedError(f"{self.name} is read-only")

    def write_cell_format(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
        format: CellFormat,
    ) -> None:
        raise NotImplementedError(f"{self.name} is read-only")

    def write_cell_border(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
        border: BorderInfo,
    ) -> None:
        raise NotImplementedError(f"{self.name} is read-only")

    def save_workbook(self, workbook: Any, path: Path) -> None:
        raise NotImplementedError(f"{self.name} is read-only")

    def set_row_height(
        self,
        workbook: Any,
        sheet: str,
        row: int,
        height: float,
    ) -> None:
        raise NotImplementedError(f"{self.name} is read-only")

    def set_column_width(
        self,
        workbook: Any,
        sheet: str,
        column: str,
        width: float,
    ) -> None:
        raise NotImplementedError(f"{self.name} is read-only")

    def merge_cells(self, workbook: Any, sheet: str, cell_range: str) -> None:
        raise NotImplementedError(f"{self.name} is read-only")

    def add_conditional_format(self, workbook: Any, sheet: str, rule: dict) -> None:
        raise NotImplementedError(f"{self.name} is read-only")

    def add_data_validation(self, workbook: Any, sheet: str, validation: dict) -> None:
        raise NotImplementedError(f"{self.name} is read-only")

    def add_hyperlink(self, workbook: Any, sheet: str, link: dict) -> None:
        raise NotImplementedError(f"{self.name} is read-only")

    def add_image(self, workbook: Any, sheet: str, image: dict) -> None:
        raise NotImplementedError(f"{self.name} is read-only")

    def add_pivot_table(self, workbook: Any, sheet: str, pivot: dict) -> None:
        raise NotImplementedError(f"{self.name} is read-only")

    def add_comment(self, workbook: Any, sheet: str, comment: dict) -> None:
        raise NotImplementedError(f"{self.name} is read-only")

    def set_freeze_panes(self, workbook: Any, sheet: str, settings: dict) -> None:
        raise NotImplementedError(f"{self.name} is read-only")


class WriteOnlyAdapter(ExcelAdapter):
    """Base class for write-only adapters.

    Provides default implementations that raise NotImplementedError
    for all read operations.
    """

    def open_workbook(self, path: Path) -> Any:
        raise NotImplementedError(f"{self.name} is write-only")

    def close_workbook(self, workbook: Any) -> None:
        pass  # Nothing to close for write-only

    def get_sheet_names(self, workbook: Any) -> list[str]:
        raise NotImplementedError(f"{self.name} is write-only")

    def read_cell_value(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
    ) -> CellValue:
        raise NotImplementedError(f"{self.name} is write-only")

    def read_cell_format(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
    ) -> CellFormat:
        raise NotImplementedError(f"{self.name} is write-only")

    def read_cell_border(
        self,
        workbook: Any,
        sheet: str,
        cell: str,
    ) -> BorderInfo:
        raise NotImplementedError(f"{self.name} is write-only")

    def read_row_height(
        self,
        workbook: Any,
        sheet: str,
        row: int,
    ) -> float | None:
        raise NotImplementedError(f"{self.name} is write-only")

    def read_column_width(
        self,
        workbook: Any,
        sheet: str,
        column: str,
    ) -> float | None:
        raise NotImplementedError(f"{self.name} is write-only")

    def read_merged_ranges(self, workbook: Any, sheet: str) -> list[str]:
        raise NotImplementedError(f"{self.name} is write-only")

    def read_conditional_formats(self, workbook: Any, sheet: str) -> list[dict]:
        raise NotImplementedError(f"{self.name} is write-only")

    def read_data_validations(self, workbook: Any, sheet: str) -> list[dict]:
        raise NotImplementedError(f"{self.name} is write-only")

    def read_hyperlinks(self, workbook: Any, sheet: str) -> list[dict]:
        raise NotImplementedError(f"{self.name} is write-only")

    def read_images(self, workbook: Any, sheet: str) -> list[dict]:
        raise NotImplementedError(f"{self.name} is write-only")

    def read_pivot_tables(self, workbook: Any, sheet: str) -> list[dict]:
        raise NotImplementedError(f"{self.name} is write-only")

    def read_comments(self, workbook: Any, sheet: str) -> list[dict]:
        raise NotImplementedError(f"{self.name} is write-only")

    def read_freeze_panes(self, workbook: Any, sheet: str) -> dict:
        raise NotImplementedError(f"{self.name} is write-only")
