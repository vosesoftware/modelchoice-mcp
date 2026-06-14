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

import json
import os
from typing import Any, ClassVar

from modelchoice_mcp.store import STORE_SHEET_NAME, parse_store, reassemble_chunks
from modelchoice_mcp.tree import DecisionTree, RollupResult, parse_model, rollup

# Excel cells cap at 32,767 chars; ModelChoice chunks the store at 32,000.
_MAX_CELL_CHARS = 32000


class ModelChoiceNotFoundError(RuntimeError):
    """The workbook contains no ModelChoice tree store (`_MC_Store`)."""


class ExcelNotRunningError(RuntimeError):
    """No running Excel instance could be attached."""


class LicenseRequiredError(RuntimeError):
    """A licence-gated action was attempted without a fully licensed
    ModelChoice. Reading trees is unaffected; building/analysis actions need
    a full licence."""


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
        self._harden_attach(app)
        if workbook is None:
            return app.books.active
        for b in app.books:
            if b.name == workbook:
                return b
        raise ModelChoiceNotFoundError(f"Workbook {workbook!r} is not open.")

    def open_workbook(self, path: str) -> dict[str, Any]:
        """Open a workbook from disk in the running Excel and report its
        sheets + any ModelChoice trees it contains. If a workbook with the
        same file name is already open, that one is used (Excel won't open two
        with the same name). Raises if Excel isn't running, the file is
        missing, or Excel can't open it."""
        xw = self._load_xw()
        app = xw.apps.active
        if app is None:
            raise ExcelNotRunningError(
                "No running Excel instance found. Open Excel (with the ModelChoice "
                "add-in) first, then open the workbook."
            )
        self._harden_attach(app)
        if not os.path.isfile(path):
            raise ModelChoiceNotFoundError(f"File not found: {path!r}.")
        name = os.path.basename(path)
        book = None
        for b in app.books:
            try:
                if str(b.name).lower() == name.lower():
                    book = b
                    break
            except Exception:
                continue
        if book is None:
            book = self._open_no_prompts(app, path)
        wb_name = str(book.name)
        try:
            sheets = [s.name for s in book.sheets]
        except Exception:
            sheets = []
        # ModelChoice tree sheets, if this is a ModelChoice workbook.
        try:
            raw = self.read_store_raw(wb_name)
            trees = list(parse_store(raw)) if raw else []
        except Exception:
            trees = []
        return {"workbook": wb_name, "sheets": sheets, "trees": trees}

    @staticmethod
    def _open_no_prompts(app: Any, path: str) -> Any:
        """Workbooks.Open with the interactive prompts suppressed, so a headless
        open never hangs on a dialog: no Update-Links, read-only-recommended, or
        file-in-use prompts. ``update_links=False`` also means external links are
        NOT refreshed on open (cell values stay as last saved). DisplayAlerts and
        AskToUpdateLinks are toggled off around the call and restored after."""
        saved: dict[str, Any] = {}
        for attr in ("DisplayAlerts", "AskToUpdateLinks"):
            try:
                saved[attr] = getattr(app.api, attr)
                setattr(app.api, attr, False)
            except Exception:
                pass
        try:
            return app.books.open(
                path,
                update_links=False,
                ignore_read_only_recommended=True,
                notify=False,
                add_to_mru=False,
            )
        except Exception as exc:
            raise ModelChoiceNotFoundError(f"Excel could not open {path!r}: {exc}") from exc
        finally:
            for attr, val in saved.items():
                try:
                    setattr(app.api, attr, val)
                except Exception:
                    pass

    def close_workbook(self, workbook: str, save: bool = False) -> dict[str, Any]:
        """Close an open workbook by file name. If ``save`` is False (default)
        UNSAVED CHANGES ARE DISCARDED; pass save=True to write them first.
        Raises if Excel isn't running or the workbook isn't open."""
        book = self._book(workbook)  # raises if not open
        name = str(book.name)
        try:
            book.api.Close(SaveChanges=bool(save))
        except Exception as exc:
            raise ModelChoiceNotFoundError(f"Could not close {workbook!r}: {exc}") from exc
        return {"closed": name, "saved": bool(save)}

    @staticmethod
    def _harden_attach(app: Any) -> None:
        """Make the COM attach robust against a loaded ModelChoice add-in.

        xlwings binds to Excel by walking window handles
        (``AccessibleObjectFromWindow`` → ``.Application``). Once ModelChoice's
        Excel-DNA add-in is loaded, that path fails with OLE ``0x800A01A8``
        (and can hang the call entirely). The ROT-based
        ``GetActiveObject('Excel.Application')`` dispatch is reliable, so we
        inject it as the app's COM object before any worksheet access. Done
        per call so an Excel restart can't leave a stale handle. No-ops off
        Windows or if win32com isn't importable."""
        try:
            import win32com.client  # type: ignore[import-untyped]

            app.impl._xl = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            pass  # fall back to xlwings' default attach

    def license_status(self, workbook: str | None = None) -> dict[str, Any]:
        """Read the ModelChoice add-in's licence state via the headless
        ``MC_LicenseStatus_Auto`` command. Returns a dict with keys like
        ``isComplete`` / ``isTrial`` / ``isExpired`` / ``isNotActivated`` /
        ``daysLeft`` / ``statusText``. Returns ``{}`` if the command isn't
        available (add-in not loaded, or older than this command) or Excel
        couldn't run it — the caller decides how to treat an unknown state."""
        try:
            book = self._book(workbook)
            raw = book.app.api.Run("MC_LicenseStatus_Auto")
        except Exception:
            return {}
        if not isinstance(raw, str) or not raw.strip():
            return {}
        try:
            data = json.loads(raw)
        except Exception:
            return {}
        return data if isinstance(data, dict) else {}

    def _require_license(self, workbook: str | None = None) -> None:
        """Gate a licence-bound ACTION on a FULL ModelChoice licence
        (``isComplete``). Pure reads never call this. Fail-closed: if the
        status can't be read (add-in missing/old/blocked), the action is
        refused with an actionable message — building/analysis require a
        licensed ModelChoice, reading does not."""
        status = self.license_status(workbook)
        if status.get("isComplete") is True:
            return
        if not status:
            raise LicenseRequiredError(
                "Could not verify the ModelChoice licence (is the add-in loaded "
                "and up to date — i.e. does it provide MC_LicenseStatus_Auto?). "
                "Building and analysis actions require a licensed ModelChoice; "
                "reading trees does not."
            )
        if status.get("isTrial"):
            state = "trial"
        elif status.get("isExpired"):
            state = "expired"
        elif status.get("isNotActivated"):
            state = "not activated"
        else:
            state = "unlicensed"
        raise LicenseRequiredError(
            f"This action requires a fully licensed ModelChoice (current: {state}). "
            "Reading trees works without a licence; building and analysis do not."
        )

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
        self._require_license(workbook)
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

    _UTILITY_FUNCTIONS: ClassVar[dict[str, str]] = {
        "exponential": "Exponential",
        "logarithmic": "Logarithmic",
        "log": "Logarithmic",
        "linear": "Linear",
        "sqrt": "Square Root",
        "square_root": "Square Root",
        "squareroot": "Square Root",
    }

    def run_utility(
        self,
        function: str = "exponential",
        risk_tolerance: float = 1.0,
        workbook: str | None = None,
    ) -> dict[str, Any]:
        """Drive ModelChoice's headless risk-attitude (utility) rollback.

        Maps a friendly function name to ModelChoice's type, calls
        ``Application.Run("MC_Utility_Auto", specJson)`` — which writes the
        ``MC_Utility`` sheet (EV, certainty equivalent, risk premium, optimal
        decision under EV vs utility) — then reads it back. Requires the
        add-in loaded with a tree open (a build that includes MC_Utility_Auto).
        Not defined for MCDA models."""
        ftype = self._UTILITY_FUNCTIONS.get(function.strip().lower())
        if ftype is None:
            raise ValueError(
                f"Unknown utility function {function!r}. Choose from: "
                f"{', '.join(sorted(set(self._UTILITY_FUNCTIONS)))}."
            )
        self._require_license(workbook)
        book = self._book(workbook)
        try:
            book.activate()
        except Exception:
            pass
        spec = json.dumps({"functionType": ftype, "riskTolerance": risk_tolerance})
        try:
            book.app.api.Run("MC_Utility_Auto", spec)
        except Exception as exc:
            raise ModelChoiceNotFoundError(
                "MC_Utility_Auto could not run — is the ModelChoice add-in loaded "
                "in Excel, with a decision tree open? (Needs a build with the utility command.)"
            ) from exc
        try:
            sheet = book.sheets["MC_Utility"]
        except Exception as exc:
            raise ModelChoiceNotFoundError(
                "Utility ran but wrote no MC_Utility sheet — check a tree is the active "
                "model and (for log/sqrt utilities) that payoffs are positive."
            ) from exc

        def _num(cell: str) -> float | None:
            v = sheet.range(cell).value
            return float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else None

        return {
            "function": sheet.range("C5").value,
            "risk_tolerance": _num("C6"),
            "expected_value": _num("C7"),
            "certainty_equivalent": _num("C8"),
            "risk_premium": _num("C9"),
            "optimal_decision_ev": sheet.range("C10").value,
            "optimal_decision_utility": sheet.range("C11").value,
        }

    def run_evii(
        self,
        chance_node: str,
        likelihoods: list[list[float]],
        test_cost: float = 0.0,
        signals: list[str] | None = None,
        test_name: str = "Test",
        workbook: str | None = None,
    ) -> dict[str, Any]:
        """Drive ModelChoice's headless EVII command and read its result.

        Builds a JSON test specification (target chance node, likelihood
        matrix P(signal|state), test cost, optional signal names) and calls
        ``Application.Run("MC_EVII_Auto", specJson)`` — which writes the
        ``MC_EVII`` report sheet — then reads that sheet back. Requires the
        ModelChoice add-in loaded with a tree open. Raises
        ``ModelChoiceNotFoundError`` if the command produced no result sheet
        (add-in not loaded, no active tree, or the chance node wasn't found)."""
        self._require_license(workbook)
        book = self._book(workbook)
        try:
            book.activate()
        except Exception:
            pass

        spec: dict[str, Any] = {
            "chanceNode": chance_node,
            "likelihoods": likelihoods,
            "testCost": test_cost,
            "testName": test_name,
        }
        if signals:
            spec["signals"] = signals

        try:
            book.app.api.Run("MC_EVII_Auto", json.dumps(spec))
        except Exception as exc:
            raise ModelChoiceNotFoundError(
                "MC_EVII_Auto could not run — is the ModelChoice add-in loaded "
                "in Excel, with a decision tree open?"
            ) from exc
        try:
            book.sheets["MC_EVII"]
        except Exception as exc:
            raise ModelChoiceNotFoundError(
                f"EVII ran but wrote no MC_EVII sheet — check that '{chance_node}' is "
                "a chance node in the active tree and the likelihood matrix matches "
                "its number of states."
            ) from exc

        rows = self.read_sheet("MC_EVII", workbook)
        after = [s.name for s in book.sheets]
        return {"rows": rows, "sheets": after}

    def build_control_panel(
        self, sheet_name: str | None = None, workbook: str | None = None
    ) -> dict[str, Any]:
        """Drive ModelChoice's headless control-panel command.

        Activates the target tree sheet (or the active one), calls
        ``Application.Run("MC_BuildControlPanel_Auto")`` — which lifts the
        tree's probability/value inputs into a labelled panel at the top of
        the sheet and links each tree cell to its panel cell — then reads the
        panel block back. Requires the ModelChoice add-in loaded with a
        rendered tree. Returns ``{sheet, rows}``."""
        self._require_license(workbook)
        book = self._book(workbook)
        try:
            book.activate()
        except Exception:
            pass

        if sheet_name is not None:
            try:
                book.sheets[sheet_name].activate()
            except Exception as exc:
                raise ModelChoiceNotFoundError(
                    f"No sheet named {sheet_name!r} to build a control panel on."
                ) from exc
            target = sheet_name
        else:
            try:
                target = book.sheets.active.name
            except Exception:
                target = None

        try:
            book.app.api.Run("MC_BuildControlPanel_Auto")
        except Exception as exc:
            raise ModelChoiceNotFoundError(
                "MC_BuildControlPanel_Auto could not run — is the ModelChoice add-in "
                "loaded in Excel, with a rendered decision tree on the active sheet?"
            ) from exc

        rows: list[list[Any]] = []
        if target is not None:
            rows = self.read_sheet(target, workbook, max_rows=60, max_cols=3)
        return {"sheet": target, "rows": rows}

    def apply_mcda(self, spec_json: str, workbook: str | None = None) -> None:
        """Turn the active (rendered) tree into an MCDA model by driving
        ``Application.Run("MC_ApplyMcda_Auto", spec_json)`` — sets the criteria,
        weights, aggregation, and per-terminal scores, then re-renders. Requires
        the add-in loaded with the active tree built (a build that includes the
        MCDA command)."""
        self._require_license(workbook)
        book = self._book(workbook)
        try:
            book.activate()
        except Exception:
            pass
        try:
            book.app.api.Run("MC_ApplyMcda_Auto", spec_json)
        except Exception as exc:
            raise ModelChoiceNotFoundError(
                "MC_ApplyMcda_Auto could not run — is the ModelChoice add-in loaded "
                "with the active tree built? (Needs a build with the MCDA command.)"
            ) from exc

    def import_precisiontree(self, file_path: str) -> dict[str, Any]:
        """Drive ModelChoice's headless PrecisionTree import. Calls
        ``Application.Run("MC_ImportPrecisionTree_Auto", file_path)`` — which
        converts a copy of the workbook (originals untouched) and opens it as
        the active workbook — then reads the converted workbook's trees.
        Requires the add-in loaded with a build that includes the import
        command."""
        self._require_license()
        xw = self._load_xw()
        app = xw.apps.active
        if app is None:
            raise ExcelNotRunningError(
                "No running Excel instance found. Open Excel first."
            )
        try:
            app.api.Run("MC_ImportPrecisionTree_Auto", file_path)
        except Exception as exc:
            raise ModelChoiceNotFoundError(
                "MC_ImportPrecisionTree_Auto could not run — is the ModelChoice add-in "
                "loaded in Excel? (Needs a build that includes the import command.)"
            ) from exc

        name: str | None = None
        try:
            name = app.books.active.name
        except Exception:
            pass
        trees: list[str] = []
        if name:
            try:
                trees = list(self.list_trees(name).keys())
            except Exception:
                trees = []
        return {"workbook": name, "trees": trees}

    def run_analysis(self, command_name: str, workbook: str | None = None) -> dict[str, Any]:
        """Drive a ModelChoice headless analysis command (an ``MC_*_Auto``
        ExcelCommand) via ``Application.Run`` and report which result
        sheets it produced. Activates the workbook first; raises
        ``ModelChoiceNotFoundError`` if the command can't run (add-in not
        loaded, or no active tree)."""
        self._require_license(workbook)
        book = self._book(workbook)
        try:
            book.activate()
        except Exception:
            pass
        before = {s.name for s in book.sheets}
        try:
            book.app.api.Run(command_name)
        except Exception as exc:
            raise ModelChoiceNotFoundError(
                f"{command_name} could not run — is the ModelChoice add-in loaded "
                "in Excel, with a decision tree open?"
            ) from exc
        after = [s.name for s in book.sheets]
        new = [n for n in after if n not in before]
        return {"command": command_name, "new_sheets": new, "sheets": after}

    def write_tree(
        self, model_json: str, sheet_name: str | None = None, workbook: str | None = None
    ) -> str:
        """Write a tree's model JSON into the workbook's ``_MC_Store``
        (v2 envelope, merged with any existing trees, chunked across
        columns) under a tree sheet name, creating the very-hidden store
        sheet if needed. Returns the sheet name used. The add-in is not
        required to store; call :meth:`render_tree` to draw it."""
        self._require_license(workbook)
        book = self._book(workbook)
        raw = self.read_store_raw(workbook)
        trees = parse_store(raw) if raw else {}

        if sheet_name is None:
            existing = {s.name for s in book.sheets} | set(trees)
            i = 1
            while f"MC_Tree_{i}" in existing:
                i += 1
            sheet_name = f"MC_Tree_{i}"

        trees[sheet_name] = model_json
        envelope = json.dumps({"_v": 2, "trees": trees})

        try:
            store = book.sheets[STORE_SHEET_NAME]
        except Exception:
            store = book.sheets.add(STORE_SHEET_NAME)
        # Write chunked across row 1; clear leftover columns.
        col = 1
        pos = 0
        while pos < len(envelope) or col == 1:
            chunk = envelope[pos:pos + _MAX_CELL_CHARS]
            store.range((1, col)).value = chunk
            pos += _MAX_CELL_CHARS
            col += 1
            if pos >= len(envelope):
                break
        for clear_col in range(col, col + 100):
            cell = store.range((1, clear_col))
            if cell.value in (None, ""):
                break
            cell.value = None
        try:
            store.api.Visible = 2  # xlSheetVeryHidden
        except Exception:
            pass
        return sheet_name

    def render_tree(self, sheet_name: str, workbook: str | None = None) -> None:
        """Draw a stored tree by name via ModelChoice's renderer
        (``MC_RenderStoredTree``). Requires the add-in loaded."""
        self._require_license(workbook)
        book = self._book(workbook)
        try:
            book.activate()
        except Exception:
            pass
        try:
            book.app.api.Run("MC_RenderStoredTree", sheet_name)
        except Exception as exc:
            raise ModelChoiceNotFoundError(
                "MC_RenderStoredTree could not run — is the ModelChoice add-in "
                "loaded? (The tree JSON was stored regardless.)"
            ) from exc

    def set_input_formula(
        self,
        named_range: str,
        formula: str,
        sheet_name: str | None = None,
        workbook: str | None = None,
    ) -> str:
        """Assign a cell formula (e.g. a ``=VoseNormal(...)`` distribution) to a
        tree input named range, the way the UI lets you type a distribution into
        a branch value/probability cell. Records it in the stored model's
        ``UserFormulas`` (so it survives re-renders) and re-renders the tree so
        the formula lands in the cell. Returns the tree sheet name.

        Pure ``_MC_Store`` edit + render — no special add-in command beyond the
        existing MC_RenderStoredTree. ModelRisk (via modelrisk-mcp) then samples
        the formula during simulation."""
        self._require_license(workbook)
        raw = self.read_store_raw(workbook)
        if not raw:
            raise ModelChoiceNotFoundError("Workbook has no ModelChoice tree store.")
        trees = parse_store(raw)
        if sheet_name is None:
            sheet_name = next(iter(trees))
        if sheet_name not in trees:
            raise ModelChoiceNotFoundError(
                f"Tree {sheet_name!r} not found. Available: {', '.join(trees)}."
            )
        model = json.loads(trees[sheet_name])
        uf = model.get("UserFormulas")
        if not isinstance(uf, dict):
            uf = {}
        uf[named_range] = formula
        model["UserFormulas"] = uf
        self.write_tree(json.dumps(model), sheet_name=sheet_name, workbook=workbook)
        self.render_tree(sheet_name, workbook)
        return sheet_name

    def read_sheet(
        self,
        sheet_name: str,
        workbook: str | None = None,
        max_rows: int = 200,
        max_cols: int = 20,
    ) -> list[list[Any]]:
        """Read a result sheet's used range (capped) as a list of rows."""
        book = self._book(workbook)
        try:
            sheet = book.sheets[sheet_name]
        except Exception as exc:
            raise ModelChoiceNotFoundError(
                f"No sheet named {sheet_name!r} in the workbook."
            ) from exc
        used = sheet.used_range
        nrows = min(int(used.rows.count), max_rows)
        ncols = min(int(used.columns.count), max_cols)
        if nrows == 0 or ncols == 0:
            return []
        raw = sheet.range((1, 1), (nrows, ncols)).value
        if not isinstance(raw, list):
            return [[raw]]
        if raw and not isinstance(raw[0], list):
            return [raw] if ncols == 1 else [[v] for v in raw]
        return [list(row) for row in raw]
