"""Tests for generator.model.legal — contracts, clauses, amendments, diligence issues."""

from decimal import Decimal

from generator.model.legal import (
    CONTRACT_AMENDMENTS,
    CONTRACT_CLAUSES,
    LEGAL_CONTRACTS,
    LEGAL_DILIGENCE_ISSUES,
    amendments_for_contract,
    change_of_control_clauses,
    clauses_for_contract,
    contracts_by_entity,
    high_risk_clauses,
    issues_by_severity,
    issues_for_contract,
    missing_consent_issues,
    stale_summary_issues,
    total_revenue_at_risk,
)


class TestLegalContracts:
    def test_contract_count(self):
        assert len(LEGAL_CONTRACTS) >= 8

    def test_ids_unique(self):
        ids = [c.contract_id for c in LEGAL_CONTRACTS]
        assert len(ids) == len(set(ids))

    def test_id_convention(self):
        for c in LEGAL_CONTRACTS:
            assert c.contract_id.startswith("LCTR-")

    def test_entity_codes_valid(self):
        valid = {"CI", "PC", "AM", "DS"}
        for c in LEGAL_CONTRACTS:
            assert c.entity_code in valid

    def test_counterparty_ids_reference_existing(self):
        """All counterparty IDs must follow CUST-NNN or VEND-NNN format."""
        for c in LEGAL_CONTRACTS:
            assert c.counterparty_id.startswith(("CUST-", "VEND-")), (
                f"{c.contract_id} has invalid counterparty_id {c.counterparty_id}"
            )

    def test_acme_contract_present(self):
        """Acme Manufacturing (CUST-001) must have a legal contract."""
        acme = [c for c in LEGAL_CONTRACTS if c.counterparty_id == "CUST-001"]
        assert len(acme) == 1
        assert acme[0].annual_value == Decimal("36_000_000")

    def test_contracts_by_entity(self):
        pc = contracts_by_entity("PC")
        assert len(pc) >= 2
        for c in pc:
            assert c.entity_code == "PC"


class TestContractClauses:
    def test_clause_count(self):
        assert len(CONTRACT_CLAUSES) >= 5

    def test_ids_unique(self):
        ids = [cl.clause_id for cl in CONTRACT_CLAUSES]
        assert len(ids) == len(set(ids))

    def test_id_convention(self):
        for cl in CONTRACT_CLAUSES:
            assert cl.clause_id.startswith("CLS-")

    def test_contract_refs_valid(self):
        """All clause contract_id values must reference existing contracts."""
        contract_ids = {c.contract_id for c in LEGAL_CONTRACTS}
        for cl in CONTRACT_CLAUSES:
            assert cl.contract_id in contract_ids, (
                f"{cl.clause_id} references unknown contract {cl.contract_id}"
            )

    def test_risk_levels_valid(self):
        valid = {"high", "medium", "low"}
        for cl in CONTRACT_CLAUSES:
            assert cl.risk_level in valid

    def test_change_of_control_exists(self):
        """At least one change-of-control clause (TC-19 required trap)."""
        coc = change_of_control_clauses()
        assert len(coc) >= 1
        assert coc[0].risk_level == "high"

    def test_mfn_clause_exists(self):
        """At least one MFN clause (TC-19 required trap)."""
        mfn = [cl for cl in CONTRACT_CLAUSES if cl.clause_type == "mfn"]
        assert len(mfn) >= 1

    def test_exclusivity_clause_exists(self):
        """At least one exclusivity clause (TC-19 required trap)."""
        exc = [cl for cl in CONTRACT_CLAUSES if cl.clause_type == "exclusivity"]
        assert len(exc) >= 1

    def test_high_risk_clauses(self):
        high = high_risk_clauses()
        assert len(high) >= 2
        for cl in high:
            assert cl.risk_level == "high"

    def test_clauses_for_contract(self):
        acme_clauses = clauses_for_contract("LCTR-001")
        assert len(acme_clauses) >= 1
        for cl in acme_clauses:
            assert cl.contract_id == "LCTR-001"


class TestContractAmendments:
    def test_amendment_count(self):
        """Spec requires at least 2 amendments or side letters."""
        assert len(CONTRACT_AMENDMENTS) >= 2

    def test_ids_unique(self):
        ids = [a.amendment_id for a in CONTRACT_AMENDMENTS]
        assert len(ids) == len(set(ids))

    def test_id_convention(self):
        for a in CONTRACT_AMENDMENTS:
            assert a.amendment_id.startswith("AMD-")

    def test_contract_refs_valid(self):
        """All amendment contract_id values must reference existing contracts."""
        contract_ids = {c.contract_id for c in LEGAL_CONTRACTS}
        for a in CONTRACT_AMENDMENTS:
            assert a.contract_id in contract_ids, (
                f"{a.amendment_id} references unknown contract {a.contract_id}"
            )

    def test_clause_refs_valid(self):
        """Amendment clause references must reference existing clauses."""
        clause_ids = {cl.clause_id for cl in CONTRACT_CLAUSES}
        for a in CONTRACT_AMENDMENTS:
            for cid in a.changes_clause_ids:
                assert cid in clause_ids, (
                    f"{a.amendment_id} references unknown clause {cid}"
                )

    def test_amendments_for_contract(self):
        acme_amds = amendments_for_contract("LCTR-001")
        assert len(acme_amds) >= 1
        for a in acme_amds:
            assert a.contract_id == "LCTR-001"


class TestLegalDiligenceIssues:
    def test_issue_count(self):
        assert len(LEGAL_DILIGENCE_ISSUES) >= 4

    def test_ids_unique(self):
        ids = [i.issue_id for i in LEGAL_DILIGENCE_ISSUES]
        assert len(ids) == len(set(ids))

    def test_id_convention(self):
        for i in LEGAL_DILIGENCE_ISSUES:
            assert i.issue_id.startswith("LDI-")

    def test_contract_refs_valid(self):
        """All issue contract_id values must reference existing contracts."""
        contract_ids = {c.contract_id for c in LEGAL_CONTRACTS}
        for i in LEGAL_DILIGENCE_ISSUES:
            assert i.contract_id in contract_ids, (
                f"{i.issue_id} references unknown contract {i.contract_id}"
            )

    def test_clause_refs_valid(self):
        """Issue clause_id values (when present) must reference existing clauses."""
        clause_ids = {cl.clause_id for cl in CONTRACT_CLAUSES}
        for i in LEGAL_DILIGENCE_ISSUES:
            if i.clause_id is not None:
                assert i.clause_id in clause_ids, (
                    f"{i.issue_id} references unknown clause {i.clause_id}"
                )

    def test_severity_levels_valid(self):
        valid = {"high", "medium", "low"}
        for i in LEGAL_DILIGENCE_ISSUES:
            assert i.severity in valid

    def test_change_of_control_risk(self):
        """At least one change-of-control consent issue (TC-19 required trap)."""
        coc_issues = [
            i for i in LEGAL_DILIGENCE_ISSUES
            if i.issue_type == "missing_consent"
            and i.clause_id is not None
            and any(
                cl.clause_type == "change_of_control"
                for cl in CONTRACT_CLAUSES
                if cl.clause_id == i.clause_id
            )
        ]
        assert len(coc_issues) >= 1

    def test_stale_summary_contradiction(self):
        """At least one contradicts_summary issue (TC-19 required trap)."""
        stale = [i for i in LEGAL_DILIGENCE_ISSUES if i.issue_type == "contradicts_summary"]
        assert len(stale) >= 1

    def test_missing_amendment_gap(self):
        """At least one stale_document issue for missing amendment (TC-19 required trap)."""
        stale = [i for i in LEGAL_DILIGENCE_ISSUES if i.issue_type == "stale_document"]
        assert len(stale) >= 1

    def test_issues_for_contract(self):
        acme_issues = issues_for_contract("LCTR-001")
        assert len(acme_issues) >= 1
        for i in acme_issues:
            assert i.contract_id == "LCTR-001"

    def test_issues_by_severity(self):
        high = issues_by_severity("high")
        assert len(high) >= 2
        for i in high:
            assert i.severity == "high"

    def test_source_refs_nonempty_for_high(self):
        """High-severity issues must have source references (TC-19 required trap)."""
        for i in LEGAL_DILIGENCE_ISSUES:
            if i.severity == "high":
                assert len(i.source_refs) >= 1, (
                    f"{i.issue_id} high-severity issue has no source refs"
                )


class TestQueryHelpers:
    def test_total_revenue_at_risk(self):
        total = total_revenue_at_risk()
        assert total > 0
        # Acme alone is $36M, so total should be at least that
        assert total >= Decimal("36_000_000")

    def test_stale_summary_issues(self):
        stale = stale_summary_issues()
        assert len(stale) >= 2  # contradicts_summary + stale_document

    def test_missing_consent_issues(self):
        missing = missing_consent_issues()
        assert len(missing) >= 1
