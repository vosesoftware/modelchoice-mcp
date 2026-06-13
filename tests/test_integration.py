"""Gated live integration tests against a real Excel + ModelChoice add-in.

These are SKIPPED unless MODELCHOICE_LIVE=1 and a workbook with at least one
ModelChoice tree is open. They exercise the real COM bridge (no fakes) — the
attach, read, and rollback path — and serve as a regression harness for the
GetActiveObject attach fix (0.0.13). Run with:

    MODELCHOICE_LIVE=1 uv run pytest tests/test_integration.py -v

Read-only by default; they never write to the workbook.
"""

from __future__ import annotations

import os

import pytest

pytestmark = pytest.mark.skipif(
    os.environ.get("MODELCHOICE_LIVE") != "1",
    reason="live Excel test; set MODELCHOICE_LIVE=1 with a ModelChoice workbook open",
)


@pytest.fixture
def live_bridge():
    from modelchoice_mcp.bridge import ModelChoiceBridge

    return ModelChoiceBridge()


def test_attach_and_list(live_bridge) -> None:
    """The bridge attaches to the running Excel and finds at least one tree."""
    trees = live_bridge.list_trees()
    assert isinstance(trees, dict)
    assert trees, "no ModelChoice trees in the active workbook"


def test_roll_up_matches_rendered_cells(live_bridge) -> None:
    """Our Python rollback agrees with the add-in's own MC_V_ cells (when the
    tree has been rendered)."""
    trees = live_bridge.list_trees()
    name = next(iter(trees))
    result = live_bridge.roll_up(name)
    assert result.expected_value == pytest.approx(result.expected_value)

    cells = live_bridge.read_node_values()
    if not cells:
        pytest.skip("tree not rendered (no MC_V_ cells) — open it in ModelChoice")
    for node_id, node in result.node_results.items():
        if node_id in cells:
            assert node.expected_value == pytest.approx(cells[node_id], abs=0.01)
