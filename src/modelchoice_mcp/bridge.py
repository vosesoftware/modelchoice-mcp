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

    def read_node_values(self, workbook: str | None = None) -> dict[str, float]:
        """Read the rolled-back expected values ModelChoice itself wrote
        into the ``MC_V_<nodeId>`` named ranges on the rendered tree
        sheet(s). Returns ``{nodeId: value}``. These are the add-in's own
        rollback numbers — comparing them to our Python rollback proves
        the two agree. Returns an empty dict if the tree hasn't been
        rendered (named ranges only exist after a render)."""
        book = self._book(workbook)
        prefix = "MC_V_"
        out: dict[str, float] = {}

        name_objs: list[Any] = []
        try:
            name_objs.extend(list(book.names))
        except Exception:
            pass
        for sht in book.sheets:
            try:
                name_objs.extend(list(sht.names))
            except Exception:
                pass

        for nm in name_objs:
            try:
                full = str(nm.name)  # may be "MC_V_D" or "Sheet!MC_V_D"
            except Exception:
                continue
            idx = full.find(prefix)
            if idx < 0:
                continue
            node_id = full[idx + len(prefix):]
            try:
                val = nm.refers_to_range.value
            except Exception:
                continue
            if isinstance(val, (int, float)) and not isinstance(val, bool):
                out[node_id] = float(val)
        return out

    def run_evpi(self, workbook: str | None = None) -> dict[str, Any]:
        """Drive ModelChoice's headless EVPI command and read its result.

        Activates the workbook, calls ``Application.Run("MC_EVPI_Auto")``
        — which computes full-tree EVPI for the active tree and writes the
        ``MC_EVPI`` sheet — then reads that sheet back. Requires the
        ModelChoice add-in to be loaded in Excel and a tree to be active.
        Raises ``ModelChoiceNotFoundError`` if the command produced no
        result sheet (add-in not loaded, or no active tree)."""
        book = self._book(workbook)
        try:
            book.activate()
        except Exception:
            pass
        try:
            book.app.api.Run("MC_EVPI_Auto")
        except Exception as exc:
            raise ModelChoiceNotFoundError(
                "MC_EVPI_Auto could not run — is the ModelChoice add-in loaded "
                "in Excel, with a decision tree open?"
            ) from exc
        try:
            sheet = book.sheets["MC_EVPI"]
        except Exception as exc:
            raise ModelChoiceNotFoundError(
                "EVPI ran but wrote no MC_EVPI sheet — check that a tree is the "
                "active ModelChoice model."
            ) from exc

        def _num(cell: str) -> float | None:
            v = sheet.range(cell).value
            return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None

        return {
            "model_name": sheet.range("C4").value,
            "objective": sheet.range("C5").value,
            "optimal_ev": _num("C6"),
            "evpi": _num("C7"),
            "value_with_perfect_info": _num("C8"),
        }
