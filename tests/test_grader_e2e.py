"""End-to-end grader regression tests for real vs self-test evidence paths (bpt.5).

These tests exercise the actual scoring entrypoint/pipeline — _run_self_test(),
_grade_test_case(), and main() — not individual grader helper methods.

Fixtures cover:
  - Canary present only in generated input files (not agent output)
  - Canary present in agent output
  - No agent structured output (correctness must fail)
  - Malformed/absent JSON in agent output
  - SelfTestGrader path that validates generated input embedding
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from scoring.auto_grader import _grade_test_case, _run_self_test, main

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _build_suite(base: Path, tc_ids: list[str], golds: dict[str, dict],
                 input_texts: dict[str, dict[str, str]] | None = None) -> Path:
    """Build a minimal test suite directory structure.

    Parameters
    ----------
    base : root for the suite
    tc_ids : test case IDs to create
    golds : {tc_id: gold_dict}
    input_texts : {tc_id: {filename: content}} — files placed in input_files/
    """
    gold_dir = base / "gold_standards"
    gold_dir.mkdir(parents=True, exist_ok=True)

    for tc_id in tc_ids:
        # Gold standard
        gold_path = gold_dir / f"{tc_id}_gold.json"
        gold_path.write_text(json.dumps(golds[tc_id], indent=2))

        # Test case directory
        tc_dir = base / "test_cases" / tc_id
        tc_dir.mkdir(parents=True, exist_ok=True)

        # Input files
        if input_texts and tc_id in input_texts:
            inputs_dir = tc_dir / "input_files"
            inputs_dir.mkdir(parents=True, exist_ok=True)
            for fname, content in input_texts[tc_id].items():
                (inputs_dir / fname).write_text(content)

    return base


def _build_agent_output(base: Path, tc_id: str, files: dict[str, str]) -> Path:
    """Create agent output directory with the given files."""
    agent_dir = base / tc_id
    agent_dir.mkdir(parents=True, exist_ok=True)
    for fname, content in files.items():
        (agent_dir / fname).write_text(content)
    return agent_dir


# Reusable gold standard fixtures
GOLD_SIMPLE = {
    "expected_outputs": {"total_revenue": 100000, "net_income": 25000},
    "canary_verification": {"tb_file": "XK7P2M9Q"},
    "error_detection": {"ERR-001": "Revenue misstatement of $5,000 in Q3 schedule"},
}

GOLD_NO_ERRORS = {
    "expected_outputs": {"total_assets": 500000},
    "canary_verification": {"bs_file": "AB3CD4EF"},
    "error_detection": {},
}


# ---------------------------------------------------------------------------
# Self-test pipeline (exercises _run_self_test → SelfTestGrader.grade)
# ---------------------------------------------------------------------------

class TestSelfTestPipeline:
    """End-to-end tests through _run_self_test, which is the path loop.sh uses."""

    def test_self_test_passes_when_canaries_in_inputs(self, tmp_path: Path) -> None:
        """Self-test pipeline scores 3/3/3/3/3 when canaries are in input files."""
        suite = _build_suite(
            tmp_path / "suite",
            ["TC-01"],
            {"TC-01": GOLD_SIMPLE},
            {"TC-01": {"trial_balance.txt": "Revenue XK7P2M9Q total 100000"}},
        )

        ok = _run_self_test(suite)
        assert ok is True

    def test_self_test_fails_when_canary_missing_from_inputs(self, tmp_path: Path) -> None:
        """Self-test fails when a canary is not found in input files."""
        suite = _build_suite(
            tmp_path / "suite",
            ["TC-01"],
            {"TC-01": GOLD_SIMPLE},
            {"TC-01": {"trial_balance.txt": "Revenue total 100000 (no canary)"}},
        )

        ok = _run_self_test(suite)
        assert ok is False

    def test_self_test_multiple_tcs_all_must_pass(self, tmp_path: Path) -> None:
        """Self-test with multiple TCs: all must score perfect for overall pass."""
        suite = _build_suite(
            tmp_path / "suite",
            ["TC-01", "TC-02"],
            {"TC-01": GOLD_SIMPLE, "TC-02": GOLD_NO_ERRORS},
            {
                "TC-01": {"trial_balance.txt": "Revenue XK7P2M9Q total 100000"},
                "TC-02": {"balance_sheet.txt": "Assets AB3CD4EF total 500000"},
            },
        )

        ok = _run_self_test(suite)
        assert ok is True

    def test_self_test_one_tc_fails_overall_fails(self, tmp_path: Path) -> None:
        """If one TC fails, _run_self_test returns False."""
        suite = _build_suite(
            tmp_path / "suite",
            ["TC-01", "TC-02"],
            {"TC-01": GOLD_SIMPLE, "TC-02": GOLD_NO_ERRORS},
            {
                "TC-01": {"trial_balance.txt": "Revenue XK7P2M9Q total 100000"},
                # TC-02 canary missing
                "TC-02": {"balance_sheet.txt": "Assets total 500000"},
            },
        )

        ok = _run_self_test(suite)
        assert ok is False

    def test_self_test_writes_aggregate_report(self, tmp_path: Path) -> None:
        """_run_self_test writes a scoring/self_test_report.json file."""
        suite = _build_suite(
            tmp_path / "suite",
            ["TC-01"],
            {"TC-01": GOLD_NO_ERRORS},
            {"TC-01": {"balance_sheet.txt": "Assets AB3CD4EF total 500000"}},
        )

        _run_self_test(suite)

        report_path = suite / "scoring" / "self_test_report.json"
        assert report_path.exists()
        report = json.loads(report_path.read_text())
        assert report["total_test_cases"] == 1
        assert "per_dimension" in report


# ---------------------------------------------------------------------------
# Normal grading pipeline (exercises _grade_test_case → TestCaseGrader.grade)
# ---------------------------------------------------------------------------

class TestNormalGradingPipeline:
    """End-to-end tests through _grade_test_case, exercising the real grader path."""

    def test_correctness_fails_without_agent_structured_data(self, tmp_path: Path) -> None:
        """No JSON in agent output → correctness=1 through the full pipeline.

        This is the regression for the old over-credit behavior where the grader
        would fall back to expected data and give a false 3/3 on correctness.
        """
        suite = _build_suite(
            tmp_path / "suite",
            ["TC-01"],
            {"TC-01": GOLD_SIMPLE},
            {"TC-01": {"trial_balance.txt": "Revenue XK7P2M9Q total 100000"}},
        )
        gold_path = suite / "gold_standards" / "TC-01_gold.json"

        # Agent output: only a text file, no JSON
        agent_out = tmp_path / "agent_output" / "TC-01"
        agent_out.mkdir(parents=True, exist_ok=True)
        (agent_out / "notes.txt").write_text("I looked at the files.")

        report = _grade_test_case("TC-01", gold_path, agent_out, suite, None)

        assert report.dimensions["correctness"].score == 1

    def test_correctness_succeeds_with_matching_agent_json(self, tmp_path: Path) -> None:
        """Agent JSON matching gold → correctness=3 through the full pipeline."""
        suite = _build_suite(
            tmp_path / "suite",
            ["TC-01"],
            {"TC-01": GOLD_SIMPLE},
        )
        gold_path = suite / "gold_standards" / "TC-01_gold.json"

        agent_out = tmp_path / "agent_output" / "TC-01"
        agent_out.mkdir(parents=True, exist_ok=True)
        (agent_out / "result.json").write_text(json.dumps({
            "total_revenue": 100000,
            "net_income": 25000,
        }))

        report = _grade_test_case("TC-01", gold_path, agent_out, suite, None)

        assert report.dimensions["correctness"].score == 3

    def test_canary_fails_when_only_in_input_files(self, tmp_path: Path) -> None:
        """Canary in input files but not agent output → canary dimension fails.

        Regression: old grader searched the entire suite_dir, giving false passes.
        """
        suite = _build_suite(
            tmp_path / "suite",
            ["TC-01"],
            {"TC-01": GOLD_SIMPLE},
            {"TC-01": {"trial_balance.txt": "Revenue XK7P2M9Q total 100000"}},
        )
        gold_path = suite / "gold_standards" / "TC-01_gold.json"

        # Agent output does NOT contain the canary
        agent_out = tmp_path / "agent_output" / "TC-01"
        agent_out.mkdir(parents=True, exist_ok=True)
        (agent_out / "result.json").write_text(json.dumps({
            "total_revenue": 100000,
            "net_income": 25000,
        }))

        report = _grade_test_case("TC-01", gold_path, agent_out, suite, None)

        assert report.dimensions["canaries"].score < 3

    def test_canary_passes_when_in_agent_output(self, tmp_path: Path) -> None:
        """Canary in agent output → canary dimension passes."""
        suite = _build_suite(
            tmp_path / "suite",
            ["TC-01"],
            {"TC-01": GOLD_SIMPLE},
        )
        gold_path = suite / "gold_standards" / "TC-01_gold.json"

        agent_out = tmp_path / "agent_output" / "TC-01"
        agent_out.mkdir(parents=True, exist_ok=True)
        (agent_out / "result.txt").write_text("Found canary XK7P2M9Q in trial balance")

        report = _grade_test_case("TC-01", gold_path, agent_out, suite, None)

        assert report.dimensions["canaries"].score == 3

    def test_malformed_json_in_agent_output(self, tmp_path: Path) -> None:
        """Malformed JSON in agent output → correctness=1 (not a crash)."""
        suite = _build_suite(
            tmp_path / "suite",
            ["TC-01"],
            {"TC-01": GOLD_SIMPLE},
        )
        gold_path = suite / "gold_standards" / "TC-01_gold.json"

        agent_out = tmp_path / "agent_output" / "TC-01"
        agent_out.mkdir(parents=True, exist_ok=True)
        (agent_out / "result.json").write_text("{broken json: [}")

        report = _grade_test_case("TC-01", gold_path, agent_out, suite, None)

        # Should degrade gracefully, not crash
        assert report.dimensions["correctness"].score == 1

    def test_writes_report_json_when_path_given(self, tmp_path: Path) -> None:
        """_grade_test_case writes a report file when output_path is provided."""
        suite = _build_suite(
            tmp_path / "suite",
            ["TC-01"],
            {"TC-01": GOLD_NO_ERRORS},
        )
        gold_path = suite / "gold_standards" / "TC-01_gold.json"

        agent_out = tmp_path / "agent_output" / "TC-01"
        agent_out.mkdir(parents=True, exist_ok=True)
        (agent_out / "result.json").write_text(json.dumps({"total_assets": 500000}))

        report_path = tmp_path / "reports" / "TC-01_report.json"
        _grade_test_case("TC-01", gold_path, agent_out, suite, report_path)

        assert report_path.exists()
        report_data = json.loads(report_path.read_text())
        assert report_data["test_case"] == "TC-01"
        assert "scores" in report_data


# ---------------------------------------------------------------------------
# CLI entrypoint (exercises main())
# ---------------------------------------------------------------------------

class TestMainEntrypoint:
    """End-to-end tests through main(), matching how benchmark users invoke the grader."""

    def test_main_self_test_exits_zero_on_pass(self, tmp_path: Path) -> None:
        """main(--self-test --suite-dir ...) exits 0 when all TCs pass."""
        suite = _build_suite(
            tmp_path / "suite",
            ["TC-01"],
            {"TC-01": GOLD_NO_ERRORS},
            {"TC-01": {"balance_sheet.txt": "Assets AB3CD4EF total 500000"}},
        )

        # main() calls sys.exit(0) on success
        with pytest.raises(SystemExit) as exc_info:
            main(["--self-test", "--suite-dir", str(suite)])

        assert exc_info.value.code == 0

    def test_main_self_test_exits_nonzero_on_fail(self, tmp_path: Path) -> None:
        """main(--self-test --suite-dir ...) exits 1 when a TC fails."""
        suite = _build_suite(
            tmp_path / "suite",
            ["TC-01"],
            {"TC-01": GOLD_SIMPLE},
            # Canary missing from input files
            {"TC-01": {"trial_balance.txt": "Revenue total 100000"}},
        )

        with pytest.raises(SystemExit) as exc_info:
            main(["--self-test", "--suite-dir", str(suite)])

        assert exc_info.value.code == 1

    def test_main_single_tc_grading(self, tmp_path: Path, capsys) -> None:
        """main(--tc --gold --agent-output) prints scores and returns normally."""
        suite = _build_suite(
            tmp_path / "suite",
            ["TC-01"],
            {"TC-01": GOLD_NO_ERRORS},
        )
        gold_path = suite / "gold_standards" / "TC-01_gold.json"

        agent_out = tmp_path / "agent_output" / "TC-01"
        agent_out.mkdir(parents=True, exist_ok=True)
        (agent_out / "result.json").write_text(json.dumps({"total_assets": 500000}))
        (agent_out / "notes.txt").write_text("Found canary AB3CD4EF in balance sheet")

        # Should not raise SystemExit
        main([
            "--tc", "TC-01",
            "--gold", str(gold_path),
            "--agent-output", str(agent_out),
            "--suite-dir", str(suite),
        ])

        captured = capsys.readouterr()
        assert "TC-01:" in captured.out

    def test_main_batch_grading(self, tmp_path: Path, capsys) -> None:
        """main(--suite-dir --agent-output-dir) grades all TCs with agent output."""
        suite = _build_suite(
            tmp_path / "suite",
            ["TC-01", "TC-02"],
            {"TC-01": GOLD_NO_ERRORS, "TC-02": GOLD_NO_ERRORS},
        )

        agent_root = tmp_path / "agent_output"
        for tc_id in ["TC-01", "TC-02"]:
            _build_agent_output(agent_root, tc_id, {
                "result.json": json.dumps({"total_assets": 500000}),
                "notes.txt": "Found canary AB3CD4EF",
            })

        report_dir = tmp_path / "reports"
        main([
            "--suite-dir", str(suite),
            "--agent-output-dir", str(agent_root),
            "--report-dir", str(report_dir),
        ])

        captured = capsys.readouterr()
        assert "TC-01:" in captured.out
        assert "TC-02:" in captured.out

        # Aggregate report written
        agg_path = report_dir / "aggregate_report.json"
        assert agg_path.exists()

    def test_main_no_args_exits_with_help(self, tmp_path: Path) -> None:
        """main() with no arguments exits 1 (prints help)."""
        with pytest.raises(SystemExit) as exc_info:
            main([])

        assert exc_info.value.code == 1


# ---------------------------------------------------------------------------
# Regression: over-credit behavior
# ---------------------------------------------------------------------------

class TestOverCreditRegression:
    """Regression tests ensuring the old over-credit bug stays fixed.

    The old grader would:
    1. Fall back to expected data when agent JSON was absent → false correctness=3
    2. Search input files (not agent output) for canaries → false canary=3
    """

    def test_empty_agent_dir_does_not_score_perfect(self, tmp_path: Path) -> None:
        """Empty agent output directory must NOT score 3 on correctness or canaries."""
        suite = _build_suite(
            tmp_path / "suite",
            ["TC-01"],
            {"TC-01": GOLD_SIMPLE},
            {"TC-01": {"trial_balance.txt": "Revenue XK7P2M9Q total 100000"}},
        )
        gold_path = suite / "gold_standards" / "TC-01_gold.json"

        # Empty agent output
        agent_out = tmp_path / "agent_output" / "TC-01"
        agent_out.mkdir(parents=True, exist_ok=True)

        report = _grade_test_case("TC-01", gold_path, agent_out, suite, None)

        assert report.dimensions["correctness"].score == 1, \
            "Empty agent output must not get perfect correctness"
        assert report.dimensions["canaries"].score < 3, \
            "Empty agent output must not get perfect canaries"

    def test_self_test_path_still_scores_perfect(self, tmp_path: Path) -> None:
        """Self-test path (no agent output, designed to validate inputs) still gets 3/3/3/3/3.

        This ensures the fix for over-credit didn't break the legitimate self-test flow.
        """
        suite = _build_suite(
            tmp_path / "suite",
            ["TC-01"],
            {"TC-01": GOLD_SIMPLE},
            {"TC-01": {"trial_balance.txt": "Revenue XK7P2M9Q total 100000"}},
        )

        ok = _run_self_test(suite)
        assert ok is True, "Self-test must still pass after the over-credit fix"
