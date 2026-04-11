"""Focused tests for real-grading vs self-test evidence modes (bpt.1/bpt.2).

These tests verify:
  - Normal grading (TestCaseGrader) searches only agent outputs for canaries
    and returns None when no agent structured data is found.
  - Self-test grading (SelfTestGrader) searches input files for canaries
    and falls back to expected data for correctness.
"""

from __future__ import annotations

import json
from pathlib import Path

from scoring.auto_grader import SelfTestGrader, TestCaseGrader

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _write_gold(gold_dir: Path, tc_id: str, gold: dict) -> Path:
    """Write a gold JSON file and return its path."""
    gold_dir.mkdir(parents=True, exist_ok=True)
    path = gold_dir / f"{tc_id}_gold.json"
    path.write_text(json.dumps(gold, indent=2))
    return path


def _make_text_file(directory: Path, name: str, content: str) -> Path:
    """Create a text file under *directory*."""
    directory.mkdir(parents=True, exist_ok=True)
    p = directory / name
    p.write_text(content)
    return p


GOLD_WITH_CANARY = {
    "expected_outputs": {"total_revenue": 100000},
    "canary_verification": {"tb_file": "XK7P2M9Q"},
    "error_detection": {},
}

GOLD_NO_CANARY = {
    "expected_outputs": {"total_revenue": 100000},
    "canary_verification": {},
    "error_detection": {},
}


# ---------------------------------------------------------------------------
# bpt.1 — Canary search scope
# ---------------------------------------------------------------------------

class TestCanarySearchScope:
    """Normal grading must search only agent outputs; self-test searches input files."""

    def test_normal_grader_ignores_input_files(self, tmp_path: Path) -> None:
        """Canary in input files but NOT in agent output → canary check fails."""
        suite = tmp_path / "suite"
        tc_id = "TC-99"

        gold_path = _write_gold(suite / "gold_standards", tc_id, GOLD_WITH_CANARY)

        # Put canary in input files (should be ignored by normal grader)
        _make_text_file(
            suite / "test_cases" / tc_id / "input_files",
            "trial_balance.txt",
            "Revenue row canary=XK7P2M9Q total 100000",
        )

        # Agent output has NO canary
        agent_out = tmp_path / "agent_output" / tc_id
        _make_text_file(agent_out, "result.txt", "Revenue total is 100000")

        grader = TestCaseGrader(
            test_case_id=tc_id,
            gold_standard_path=gold_path,
            agent_output_path=agent_out,
            suite_dir=suite,
        )
        result = grader.verify_canaries()

        assert result.score < 3, (
            "Normal grader should fail canary when canary is only in input files"
        )
        assert any(c.status == "fail" for c in result.checks)

    def test_normal_grader_passes_when_canary_in_agent_output(self, tmp_path: Path) -> None:
        """Canary present in agent output → canary check passes."""
        suite = tmp_path / "suite"
        tc_id = "TC-99"

        gold_path = _write_gold(suite / "gold_standards", tc_id, GOLD_WITH_CANARY)

        # Agent output contains the canary
        agent_out = tmp_path / "agent_output" / tc_id
        _make_text_file(agent_out, "result.txt", "Found canary XK7P2M9Q in TB")

        grader = TestCaseGrader(
            test_case_id=tc_id,
            gold_standard_path=gold_path,
            agent_output_path=agent_out,
            suite_dir=suite,
        )
        result = grader.verify_canaries()

        assert result.score == 3
        assert all(c.status == "pass" for c in result.checks)

    def test_self_test_grader_searches_input_files(self, tmp_path: Path) -> None:
        """SelfTestGrader finds canaries in generated input files."""
        suite = tmp_path / "suite"
        tc_id = "TC-99"

        gold_path = _write_gold(suite / "gold_standards", tc_id, GOLD_WITH_CANARY)

        # Canary lives in input files
        _make_text_file(
            suite / "test_cases" / tc_id / "input_files",
            "trial_balance.txt",
            "Revenue row canary=XK7P2M9Q total 100000",
        )

        # SelfTestGrader uses the TC dir as agent_output_path
        tc_dir = suite / "test_cases" / tc_id
        grader = SelfTestGrader(
            test_case_id=tc_id,
            gold_standard_path=gold_path,
            agent_output_path=tc_dir,
            suite_dir=suite,
        )
        result = grader.verify_canaries()

        assert result.score == 3
        assert all(c.status == "pass" for c in result.checks)

    def test_self_test_grader_fails_when_canary_missing(self, tmp_path: Path) -> None:
        """SelfTestGrader fails when canary is absent from input files."""
        suite = tmp_path / "suite"
        tc_id = "TC-99"

        gold_path = _write_gold(suite / "gold_standards", tc_id, GOLD_WITH_CANARY)

        # Input files exist but without the canary
        _make_text_file(
            suite / "test_cases" / tc_id / "input_files",
            "trial_balance.txt",
            "Revenue total 100000",
        )

        tc_dir = suite / "test_cases" / tc_id
        grader = SelfTestGrader(
            test_case_id=tc_id,
            gold_standard_path=gold_path,
            agent_output_path=tc_dir,
            suite_dir=suite,
        )
        result = grader.verify_canaries()

        assert result.score < 3


# ---------------------------------------------------------------------------
# bpt.2 — Agent data fallback
# ---------------------------------------------------------------------------

class TestAgentDataFallback:
    """Normal grading must NOT fall back to expected data when agent output is absent."""

    def test_normal_grader_returns_none_without_agent_json(self, tmp_path: Path) -> None:
        """No JSON in agent output → _load_agent_data returns None."""
        suite = tmp_path / "suite"
        tc_id = "TC-99"

        gold_path = _write_gold(suite / "gold_standards", tc_id, GOLD_NO_CANARY)

        # Agent output with only a text file (no JSON)
        agent_out = tmp_path / "agent_output" / tc_id
        _make_text_file(agent_out, "notes.txt", "No structured data here")

        grader = TestCaseGrader(
            test_case_id=tc_id,
            gold_standard_path=gold_path,
            agent_output_path=agent_out,
            suite_dir=suite,
        )
        assert grader._load_agent_data() is None

    def test_normal_grader_correctness_fails_without_agent_data(self, tmp_path: Path) -> None:
        """Correctness must score 1 when agent has no structured output."""
        suite = tmp_path / "suite"
        tc_id = "TC-99"

        gold_path = _write_gold(suite / "gold_standards", tc_id, GOLD_NO_CANARY)

        # Agent output: no JSON files
        agent_out = tmp_path / "agent_output" / tc_id
        _make_text_file(agent_out, "notes.txt", "No structured data here")

        grader = TestCaseGrader(
            test_case_id=tc_id,
            gold_standard_path=gold_path,
            agent_output_path=agent_out,
            suite_dir=suite,
        )
        result = grader.grade_correctness()

        assert result.score == 1, (
            "Normal grader should score correctness=1 when no agent data found"
        )
        assert any(c.status == "fail" and "agent output" in c.detail.lower()
                    for c in result.checks)

    def test_self_test_grader_uses_expected_data(self, tmp_path: Path) -> None:
        """SelfTestGrader._load_agent_data returns the expected data from gold."""
        suite = tmp_path / "suite"
        tc_id = "TC-99"

        gold_path = _write_gold(suite / "gold_standards", tc_id, GOLD_NO_CANARY)

        tc_dir = suite / "test_cases" / tc_id
        tc_dir.mkdir(parents=True, exist_ok=True)

        grader = SelfTestGrader(
            test_case_id=tc_id,
            gold_standard_path=gold_path,
            agent_output_path=tc_dir,
            suite_dir=suite,
        )
        data = grader._load_agent_data()

        assert data is not None
        assert data == {"total_revenue": 100000}

    def test_self_test_grader_correctness_perfect(self, tmp_path: Path) -> None:
        """SelfTestGrader scores correctness=3 (gold compared to itself)."""
        suite = tmp_path / "suite"
        tc_id = "TC-99"

        gold_path = _write_gold(suite / "gold_standards", tc_id, GOLD_NO_CANARY)

        tc_dir = suite / "test_cases" / tc_id
        tc_dir.mkdir(parents=True, exist_ok=True)

        grader = SelfTestGrader(
            test_case_id=tc_id,
            gold_standard_path=gold_path,
            agent_output_path=tc_dir,
            suite_dir=suite,
        )
        result = grader.grade_correctness()

        assert result.score == 3
