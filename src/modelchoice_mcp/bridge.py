"""Excel bridge for ModelChoice — read decision trees from a workbook.

Phase 0 is read-only. ModelChoice exposes no COM object model, but the
full tree model is persisted as JSON in the very-hidden ``_MC_Store``
worksheet (see :mod:`modelchoice_mcp.store`). We read that sheet's row-1
content (chunked across columns) via xlwings, parse each tree, and roll
it back in pure Python — so the add-in need not even be loaded.

The bridge deliberately does NOT unhide or modify ``_MC_Store``; it
reads cell values only.
"""

from __future__ import annotations

from typing import Any

from modelchoice_mcp.store import STORE_SHEET_NAME, parse_store, reassemble_chunks
from modelchoice_mcp.tree import DecisionTree, RollupResult, parse_model, rollup


class ModelChoiceNotFoundError(RuntimeError):
    """The workbook contains no ModelChoice tree store (`_MC_Store`)."""


class ExcelNotRunningError(RuntimeError):
    """No running Excel instance could be attached."""


class ModelChoiceBridge:
    """Attach to a running Excel and read ModelChoice trees from a
    workbook's very-hidden ``_MC_Store`` sheet."""

    def __init__(self) -> None:
        self._xw: Any = None

    def _load_xw(self) -> Any:
        if self._xw is None:
            import xlwings  # imported lazily so the package is importable without Excel

            self._xw = xlwings
        return self._xw

    def _book(self, workbook: str | None) -> Any:
        xw = self._load_xw()
        app = xw.apps.active
        if app is None:
            raise ExcelNotRunningError(
                "No running Excel instance found. Open the workbook in Excel first."
            )
        if workbook is None:
            return app.books.active
        for b in app.books:
            if b.name == workbook:
                return b
        raise ModelChoiceNotFoundError(f"Workbook {workbook!r} is not open.")

    def read_store_raw(self, workbook: str | None = None) -> str:
        """Return the reassembled ``_MC_Store`` A1 payload, or '' if the
        sheet is absent."""
        book = self._book(workbook)
        try:
            sheet = book.sheets[STORE_SHEET_NAME]
        except Exception:
            return ""
        # Read row 1 across enough columns to cover the chunk limit.
        row = sheet.range((1, 1), (1, 100)).value
        if not isinstance(row, list):
            row = [row]
        cells = [None if v is None else str(v) for v in row]
        return reassemble_chunks(cells)

    def list_trees(self, workbook: str | None = None) -> dict[str, str]:
        """Return a mapping of tree-sheet-name → model JSON string."""
        raw = self.read_store_raw(workbook)
        if not raw:
            raise ModelChoiceNotFoundError(
                "Workbook has no ModelChoice tree store (_MC_Store sheet)."
            )
        return parse_store(raw)

    def get_tree(self, sheet_name: str | None = None, workbook: str | None = None) -> DecisionTree:
        """Parse one tree (by its sheet name, or the first/only tree)."""
        trees = self.list_trees(workbook)
        if not trees:
            raise ModelChoiceNotFoundError("No trees stored in this workbook.")
        if sheet_name is None:
            sheet_name = next(iter(trees))
        if sheet_name not in trees:
            raise ModelChoiceNotFoundError(
                f"Tree {sheet_name!r} not found. Available: {', '.join(trees)}."
            )
        return parse_model(trees[sheet_name])

    def roll_up(self, sheet_name: str | None = None, workbook: str | None = None) -> RollupResult:
        """Read a tree and return its rolled-back EV + optimal policy."""
        return rollup(self.get_tree(sheet_name, workbook))
