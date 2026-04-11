"""Tests that CascadeModel includes legal and HR diligence fields."""

from __future__ import annotations

import pytest

from generator.model.build import CascadeModel, build_model
from generator.model.hr_diligence import (
    CONTRACTOR_CLASSIFICATION_SIGNALS,
    DILIGENCE_REQUESTS,
    EMPLOYMENT_AGREEMENTS,
    RETENTION_AWARDS,
    SEVERANCE_EXPOSURES,
    ContractorClassificationSignal,
    DiligenceRequest,
    EmploymentAgreement,
    RetentionAward,
    SeveranceExposure,
)
from generator.model.legal import (
    CONTRACT_AMENDMENTS,
    CONTRACT_CLAUSES,
    LEGAL_CONTRACTS,
    LEGAL_DILIGENCE_ISSUES,
    ContractAmendment,
    ContractClause,
    LegalContract,
    LegalDiligenceIssue,
)


@pytest.fixture(scope="module")
def model() -> CascadeModel:
    return build_model(seed=42)


# ── Legal diligence fields ─────────────────────────────────────────────────


class TestLegalDiligenceOnModel:
    def test_legal_contracts_present(self, model: CascadeModel) -> None:
        assert len(model.legal_contracts) > 0
        assert model.legal_contracts is LEGAL_CONTRACTS
        assert all(isinstance(c, LegalContract) for c in model.legal_contracts)

    def test_contract_clauses_present(self, model: CascadeModel) -> None:
        assert len(model.contract_clauses) > 0
        assert model.contract_clauses is CONTRACT_CLAUSES
        assert all(isinstance(c, ContractClause) for c in model.contract_clauses)

    def test_contract_amendments_present(self, model: CascadeModel) -> None:
        assert len(model.contract_amendments) > 0
        assert model.contract_amendments is CONTRACT_AMENDMENTS
        assert all(isinstance(a, ContractAmendment) for a in model.contract_amendments)

    def test_legal_diligence_issues_present(self, model: CascadeModel) -> None:
        assert len(model.legal_diligence_issues) > 0
        assert model.legal_diligence_issues is LEGAL_DILIGENCE_ISSUES
        assert all(
            isinstance(i, LegalDiligenceIssue) for i in model.legal_diligence_issues
        )


# ── HR diligence fields ───────────────────────────────────────────────────


class TestHRDiligenceOnModel:
    def test_employment_agreements_present(self, model: CascadeModel) -> None:
        assert len(model.employment_agreements) > 0
        assert model.employment_agreements is EMPLOYMENT_AGREEMENTS
        assert all(
            isinstance(ea, EmploymentAgreement) for ea in model.employment_agreements
        )

    def test_retention_awards_present(self, model: CascadeModel) -> None:
        assert len(model.retention_awards) > 0
        assert model.retention_awards is RETENTION_AWARDS
        assert all(isinstance(r, RetentionAward) for r in model.retention_awards)

    def test_severance_exposures_present(self, model: CascadeModel) -> None:
        assert len(model.severance_exposures) > 0
        assert model.severance_exposures is SEVERANCE_EXPOSURES
        assert all(
            isinstance(s, SeveranceExposure) for s in model.severance_exposures
        )

    def test_contractor_signals_present(self, model: CascadeModel) -> None:
        assert len(model.contractor_signals) > 0
        assert model.contractor_signals is CONTRACTOR_CLASSIFICATION_SIGNALS
        assert all(
            isinstance(c, ContractorClassificationSignal)
            for c in model.contractor_signals
        )

    def test_diligence_requests_present(self, model: CascadeModel) -> None:
        assert len(model.diligence_requests) > 0
        assert model.diligence_requests is DILIGENCE_REQUESTS
        assert all(isinstance(d, DiligenceRequest) for d in model.diligence_requests)


# ── Defaults for backwards compatibility ───────────────────────────────────


class TestDefaultsAreEmptyTuples:
    """CascadeModel fields default to empty tuples so existing code that
    constructs CascadeModel without legal/HR args still works."""

    def test_legal_defaults(self) -> None:
        from generator.model.gl import Ledger

        m = CascadeModel(
            ledger=Ledger(),
            revenue_records=[],
            employees=[],
            opex_records=[],
            assets=[],
            leases=[],
            lease_schedules=[],
            tax_provisions={},
        )
        assert m.legal_contracts == ()
        assert m.contract_clauses == ()
        assert m.contract_amendments == ()
        assert m.legal_diligence_issues == ()

    def test_hr_defaults(self) -> None:
        from generator.model.gl import Ledger

        m = CascadeModel(
            ledger=Ledger(),
            revenue_records=[],
            employees=[],
            opex_records=[],
            assets=[],
            leases=[],
            lease_schedules=[],
            tax_provisions={},
        )
        assert m.employment_agreements == ()
        assert m.retention_awards == ()
        assert m.severance_exposures == ()
        assert m.contractor_signals == ()
        assert m.diligence_requests == ()
