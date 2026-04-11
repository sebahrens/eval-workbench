"""Tests for evidence grading and real-grading vs self-test evidence modes.

These tests verify:
  - Normal grading (TestCaseGrader) searches only agent outputs for canaries
    and returns None when no agent structured data is found.
  - Self-test grading (SelfTestGrader) searches input files for canaries
    and falls back to expected data for correctness.
  - Evidence expectation checks: full credit, partial/fail, and backward compat.
"""

from __future__ import annotations

import json
from pathlib import Path

from scoring.auto_grader import (
    SelfTestGrader,
    TestCaseGrader,
    aggregate_reports,
)

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


# ---------------------------------------------------------------------------
# Evidence expectation checks (ups.12)
# ---------------------------------------------------------------------------

EVIDENCE_GOLD = {
    "expected_outputs": {"total_revenue": 100000},
    "canary_verification": {},
    "error_detection": {},
    "evidence_expectations": {
        "risk_change_of_control": {
            "required_sources": ["contract_lctr_001"],
            "primary_source_required": True,
            "acceptable_terms": ["change of control", "assignment"],
        },
        "risk_mfn_contradiction": {
            "required_sources": ["contract_lctr_003", "amendment_amd_002"],
            "primary_source_required": True,
            "acceptable_terms": ["most-favored-nation", "MFN"],
        },
    },
}


class TestEvidenceCheckFullCredit:
    """Evidence check passes when agent output cites all sources and terms."""

    def test_full_credit(self, tmp_path: Path) -> None:
        suite = tmp_path / "suite"
        tc_id = "TC-99"
        gold_path = _write_gold(suite / "gold_standards", tc_id, EVIDENCE_GOLD)

        agent_out = tmp_path / "agent_output" / tc_id
        _make_text_file(
            agent_out, "findings.txt",
            "Per contract_lctr_001, change of control clause triggers consent.\n"
            "contract_lctr_003 and amendment_amd_002 show MFN contradiction.\n"
            "Revenue: 100000",
        )

        grader = TestCaseGrader(
            test_case_id=tc_id,
            gold_standard_path=gold_path,
            agent_output_path=agent_out,
            suite_dir=suite,
        )
        result = grader.check_evidence()

        assert result.dimension == "evidence"
        assert result.score == 3
        assert all(c.status == "pass" for c in result.checks)

    def test_full_credit_case_insensitive(self, tmp_path: Path) -> None:
        """Term matching is case-insensitive."""
        suite = tmp_path / "suite"
        tc_id = "TC-99"
        gold_path = _write_gold(suite / "gold_standards", tc_id, EVIDENCE_GOLD)

        agent_out = tmp_path / "agent_output" / tc_id
        _make_text_file(
            agent_out, "findings.txt",
            "Per CONTRACT_LCTR_001, Change Of Control clause triggers consent.\n"
            "CONTRACT_LCTR_003 and AMENDMENT_AMD_002 show mfn contradiction.\n",
        )

        grader = TestCaseGrader(
            test_case_id=tc_id,
            gold_standard_path=gold_path,
            agent_output_path=agent_out,
            suite_dir=suite,
        )
        result = grader.check_evidence()
        assert result.score == 3


class TestEvidenceCheckPartialFail:
    """Evidence check fails partially when some findings lack citations."""

    def test_missing_source_fails_finding(self, tmp_path: Path) -> None:
        """One finding has sources, other is missing → score < 3."""
        suite = tmp_path / "suite"
        tc_id = "TC-99"
        gold_path = _write_gold(suite / "gold_standards", tc_id, EVIDENCE_GOLD)

        agent_out = tmp_path / "agent_output" / tc_id
        _make_text_file(
            agent_out, "findings.txt",
            "Per contract_lctr_001, change of control clause triggers consent.\n"
            "MFN clause may be problematic.\n",  # Missing sources for 2nd finding
        )

        grader = TestCaseGrader(
            test_case_id=tc_id,
            gold_standard_path=gold_path,
            agent_output_path=agent_out,
            suite_dir=suite,
        )
        result = grader.check_evidence()

        assert result.score == 2  # 1 of 2 passed → ≥50%
        failed = [c for c in result.checks if c.status == "fail"]
        assert len(failed) == 1
        assert "mfn_contradiction" in failed[0].name

    def test_no_sources_no_terms_fails_all(self, tmp_path: Path) -> None:
        """Agent output has no relevant citations → score 1."""
        suite = tmp_path / "suite"
        tc_id = "TC-99"
        gold_path = _write_gold(suite / "gold_standards", tc_id, EVIDENCE_GOLD)

        agent_out = tmp_path / "agent_output" / tc_id
        _make_text_file(agent_out, "notes.txt", "Nothing relevant here.")

        grader = TestCaseGrader(
            test_case_id=tc_id,
            gold_standard_path=gold_path,
            agent_output_path=agent_out,
            suite_dir=suite,
        )
        result = grader.check_evidence()

        assert result.score == 1
        assert all(c.status == "fail" for c in result.checks)

    def test_terms_present_but_sources_missing(self, tmp_path: Path) -> None:
        """Terms found but primary source not cited → still fails."""
        suite = tmp_path / "suite"
        tc_id = "TC-99"
        gold_path = _write_gold(suite / "gold_standards", tc_id, EVIDENCE_GOLD)

        agent_out = tmp_path / "agent_output" / tc_id
        _make_text_file(
            agent_out, "findings.txt",
            "Change of control and assignment clauses exist.\n"
            "MFN pricing threshold noted.\n",
        )

        grader = TestCaseGrader(
            test_case_id=tc_id,
            gold_standard_path=gold_path,
            agent_output_path=agent_out,
            suite_dir=suite,
        )
        result = grader.check_evidence()

        assert result.score == 1  # 0 of 2 findings pass (sources missing)


class TestEvidenceBackwardCompat:
    """Golds without evidence_expectations must not regress."""

    def test_no_evidence_in_gold_scores_3(self, tmp_path: Path) -> None:
        """Old golds (no evidence_expectations) → evidence score 3, not in report."""
        suite = tmp_path / "suite"
        tc_id = "TC-99"
        gold_path = _write_gold(suite / "gold_standards", tc_id, GOLD_NO_CANARY)

        agent_out = tmp_path / "agent_output" / tc_id
        _make_text_file(agent_out, "result.json", '{"total_revenue": 100000}')

        grader = TestCaseGrader(
            test_case_id=tc_id,
            gold_standard_path=gold_path,
            agent_output_path=agent_out,
            suite_dir=suite,
        )
        result = grader.check_evidence()
        assert result.score == 3

        # Full grade should NOT include evidence dimension for old golds
        report = grader.grade()
        assert "evidence" not in report.dimensions

    def test_grade_includes_evidence_when_present(self, tmp_path: Path) -> None:
        """Golds with evidence_expectations include evidence in grade report."""
        suite = tmp_path / "suite"
        tc_id = "TC-99"
        gold_path = _write_gold(suite / "gold_standards", tc_id, EVIDENCE_GOLD)

        agent_out = tmp_path / "agent_output" / tc_id
        _make_text_file(
            agent_out, "findings.txt",
            "contract_lctr_001 change of control. "
            "contract_lctr_003 amendment_amd_002 MFN.",
        )

        grader = TestCaseGrader(
            test_case_id=tc_id,
            gold_standard_path=gold_path,
            agent_output_path=agent_out,
            suite_dir=suite,
        )
        report = grader.grade()
        assert "evidence" in report.dimensions
        assert report.dimensions["evidence"].score == 3


class TestEvidenceSelfTest:
    """SelfTestGrader validates evidence spec structure, not agent output."""

    def test_well_formed_evidence_passes(self, tmp_path: Path) -> None:
        suite = tmp_path / "suite"
        tc_id = "TC-99"
        gold_path = _write_gold(suite / "gold_standards", tc_id, EVIDENCE_GOLD)

        tc_dir = suite / "test_cases" / tc_id
        tc_dir.mkdir(parents=True, exist_ok=True)

        grader = SelfTestGrader(
            test_case_id=tc_id,
            gold_standard_path=gold_path,
            agent_output_path=tc_dir,
            suite_dir=suite,
        )
        result = grader.check_evidence()
        assert result.score == 3
        assert all(c.status == "pass" for c in result.checks)

    def test_malformed_evidence_fails(self, tmp_path: Path) -> None:
        bad_gold = {
            "expected_outputs": {},
            "canary_verification": {},
            "error_detection": {},
            "evidence_expectations": {
                "bad_finding": {
                    "required_sources": [],  # empty → fail
                    "acceptable_terms": ["some term"],
                },
            },
        }
        suite = tmp_path / "suite"
        tc_id = "TC-99"
        gold_path = _write_gold(suite / "gold_standards", tc_id, bad_gold)

        tc_dir = suite / "test_cases" / tc_id
        tc_dir.mkdir(parents=True, exist_ok=True)

        grader = SelfTestGrader(
            test_case_id=tc_id,
            gold_standard_path=gold_path,
            agent_output_path=tc_dir,
            suite_dir=suite,
        )
        result = grader.check_evidence()
        assert result.score < 3

    def test_no_evidence_in_self_test_scores_3(self, tmp_path: Path) -> None:
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
        result = grader.check_evidence()
        assert result.score == 3


class TestEvidenceInAggregate:
    """Aggregate reports include evidence when present."""

    def test_aggregate_includes_evidence(self, tmp_path: Path) -> None:
        suite = tmp_path / "suite"

        # TC with evidence
        gold_path = _write_gold(
            suite / "gold_standards", "TC-99", EVIDENCE_GOLD,
        )
        agent_out = tmp_path / "agent_output" / "TC-99"
        _make_text_file(
            agent_out, "findings.txt",
            "contract_lctr_001 change of control. "
            "contract_lctr_003 amendment_amd_002 MFN.",
        )
        grader1 = TestCaseGrader(
            test_case_id="TC-99",
            gold_standard_path=gold_path,
            agent_output_path=agent_out,
            suite_dir=suite,
        )
        report1 = grader1.grade()

        # TC without evidence
        gold_path2 = _write_gold(
            suite / "gold_standards", "TC-98", GOLD_NO_CANARY,
        )
        agent_out2 = tmp_path / "agent_output" / "TC-98"
        _make_text_file(agent_out2, "result.json", '{"total_revenue": 100000}')
        grader2 = TestCaseGrader(
            test_case_id="TC-98",
            gold_standard_path=gold_path2,
            agent_output_path=agent_out2,
            suite_dir=suite,
        )
        report2 = grader2.grade()

        agg = aggregate_reports([report1, report2])

        # Evidence dimension should appear in aggregate with 1 TC
        assert "evidence" in agg["per_dimension"]
        assert agg["per_dimension"]["evidence"]["total"] == 1
        assert agg["per_dimension"]["evidence"]["perfect_count"] == 1

        # TC-99 has evidence in per_test_case, TC-98 does not
        assert "evidence" in agg["per_test_case"]["TC-99"]
        assert "evidence" not in agg["per_test_case"]["TC-98"]
