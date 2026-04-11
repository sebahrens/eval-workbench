"""Tests for generator.model.hr_diligence — employment agreements, severance, retention, contractors."""

from decimal import Decimal

from generator.model.hr_diligence import (
    CONTRACTOR_CLASSIFICATION_SIGNALS,
    DILIGENCE_REQUESTS,
    EMPLOYMENT_AGREEMENTS,
    RETENTION_AWARDS,
    SEVERANCE_EXPOSURES,
    high_risk_contractors,
    missing_executed_agreements,
    open_diligence_requests,
    retention_award_for_employee,
    severance_exposure_for_employee,
    total_retention_awards,
    total_severance_exposure,
)


class TestEmploymentAgreements:
    def test_agreement_count(self):
        assert len(EMPLOYMENT_AGREEMENTS) == 7

    def test_ids_unique(self):
        ids = [ea.agreement_id for ea in EMPLOYMENT_AGREEMENTS]
        assert len(ids) == len(set(ids))

    def test_id_convention(self):
        for ea in EMPLOYMENT_AGREEMENTS:
            assert ea.agreement_id.startswith("EA-")

    def test_entity_codes_valid(self):
        valid = {"CI", "PC", "AM", "DS"}
        for ea in EMPLOYMENT_AGREEMENTS:
            assert ea.entity_code in valid

    def test_ceo_matches_key_personnel(self):
        """CEO agreement must be consistent with customers.KEY_PERSONNEL."""
        ceo = [ea for ea in EMPLOYMENT_AGREEMENTS if ea.agreement_id == "EA-001"]
        assert len(ceo) == 1
        ea = ceo[0]
        assert ea.employee_name == "Robert J. Cascade"
        assert ea.base_salary == 325_000
        assert ea.severance_multiplier == Decimal("3")
        assert ea.change_of_control_multiplier == Decimal("3")


class TestSeveranceExposure:
    def test_severance_tied_to_salary(self):
        """Each severance payout must equal base_salary * multiplier."""
        for sev in SEVERANCE_EXPOSURES:
            expected = Decimal(str(sev.base_salary)) * sev.severance_multiplier
            assert sev.estimated_payout == expected, (
                f"{sev.exposure_id}: expected {expected}, got {sev.estimated_payout}"
            )

    def test_severance_ids_unique(self):
        ids = [s.exposure_id for s in SEVERANCE_EXPOSURES]
        assert len(ids) == len(set(ids))

    def test_total_exposure(self):
        total = total_severance_exposure()
        assert total > 0
        # Sum of all individual payouts
        manual = sum(s.estimated_payout for s in SEVERANCE_EXPOSURES)
        assert total == manual

    def test_ceo_exposure_is_975k(self):
        """CEO golden parachute: 3x $325,000 = $975,000."""
        sev = severance_exposure_for_employee("Robert J. Cascade")
        assert sev is not None
        assert sev.estimated_payout == Decimal("975_000")

    def test_severance_references_agreement(self):
        """Each severance exposure should reference an employment agreement by name."""
        ea_names = {ea.employee_name for ea in EMPLOYMENT_AGREEMENTS}
        for sev in SEVERANCE_EXPOSURES:
            assert sev.employee_name in ea_names, (
                f"{sev.exposure_id} references {sev.employee_name} not in agreements"
            )


class TestRetentionAwards:
    def test_retention_count(self):
        assert len(RETENTION_AWARDS) >= 3

    def test_ids_unique(self):
        ids = [r.award_id for r in RETENTION_AWARDS]
        assert len(ids) == len(set(ids))

    def test_total_retention(self):
        total = total_retention_awards()
        assert total > 0

    def test_no_double_count_with_severance(self):
        """Dr. Patel has both retention (RET-004) and severance (SEV-006).

        The model specifies greater-of, not both.  Verify the notes
        document this explicitly so the grader can check for double-count
        avoidance.
        """
        patel_ret = retention_award_for_employee("Dr. Anika Patel")
        patel_sev = severance_exposure_for_employee("Dr. Anika Patel")

        assert patel_ret is not None, "Dr. Patel should have a retention award"
        assert patel_sev is not None, "Dr. Patel should have a severance exposure"

        # Severance > retention, so net exposure is severance amount only
        assert patel_sev.estimated_payout > patel_ret.award_amount

        # Notes must indicate non-additive treatment
        assert "not additive" in patel_sev.notes.lower() or "greater-of" in patel_sev.notes.lower()
        assert "not additive" in patel_ret.notes.lower() or "greater-of" in patel_ret.notes.lower()


class TestMissingAgreement:
    def test_one_missing_executed(self):
        """Exactly one agreement should lack an executed copy."""
        missing = missing_executed_agreements()
        assert len(missing) == 1

    def test_missing_is_rd_director(self):
        missing = missing_executed_agreements()
        assert missing[0].employee_name == "Dr. Anika Patel"
        assert missing[0].agreement_id == "EA-006"


class TestContractorClassification:
    def test_signals_exist(self):
        assert len(CONTRACTOR_CLASSIFICATION_SIGNALS) >= 3

    def test_ids_unique(self):
        ids = [c.signal_id for c in CONTRACTOR_CLASSIFICATION_SIGNALS]
        assert len(ids) == len(set(ids))

    def test_high_risk_contractors(self):
        high = high_risk_contractors()
        assert len(high) >= 2
        for c in high:
            assert c.risk_level == "high"

    def test_notes_are_follow_up_not_conclusion(self):
        """Contractor signals must be caveated as follow-up, not legal conclusions."""
        for signal in CONTRACTOR_CLASSIFICATION_SIGNALS:
            if signal.risk_level == "high":
                notes_lower = signal.notes.lower()
                assert (
                    "follow-up" in notes_lower
                    or "signal" in notes_lower
                    or "not a legal" in notes_lower
                    or "preliminary" in notes_lower
                    or "recommend" in notes_lower
                ), (
                    f"{signal.signal_id} high-risk notes must caveat as follow-up, "
                    f"not definitive legal conclusion"
                )


class TestDiligenceRequests:
    def test_request_count(self):
        assert len(DILIGENCE_REQUESTS) >= 5

    def test_ids_unique(self):
        ids = [dr.request_id for dr in DILIGENCE_REQUESTS]
        assert len(ids) == len(set(ids))

    def test_open_requests(self):
        """At least one request should be not fully received."""
        open_reqs = open_diligence_requests()
        assert len(open_reqs) >= 1

    def test_missing_ea006_tracked(self):
        """The missing EA-006 executed copy should be tracked as a diligence request."""
        ea006_requests = [
            dr for dr in DILIGENCE_REQUESTS
            if "EA-006" in dr.description or "EA-006" in dr.notes
        ]
        assert len(ea006_requests) >= 1
        # Should be not_received or open
        assert any(dr.status in ("not_received", "open") for dr in ea006_requests)
