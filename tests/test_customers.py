"""Tests for generator.model.customers — contracts, concentration, key people, litigation."""

from decimal import Decimal

from generator.model.customers import (
    CONTRACTS,
    KEY_PERSONNEL,
    LITIGATION,
    compute_customer_concentration,
    contracts_expiring_within,
    contracts_with_change_of_control,
    data_room_red_flags,
    key_personnel_with_coc,
    top_n_concentration,
    total_coc_exposure,
    total_litigation_exposure,
)


class TestContracts:
    def test_eight_contracts(self):
        assert len(CONTRACTS) == 8

    def test_acme_has_change_of_control(self):
        acme = [c for c in CONTRACTS if c.customer_id == "CUST-001"]
        assert len(acme) == 1
        assert acme[0].change_of_control_clause is True
        assert acme[0].customer_name == "Acme Manufacturing Corp"

    def test_only_one_coc_contract(self):
        coc = contracts_with_change_of_control()
        assert len(coc) == 1
        assert coc[0].contract_id == "CTR-001"

    def test_contracts_expiring_within_12_months(self):
        expiring = contracts_expiring_within(months=12)
        # Should catch contracts expiring by end of 2026
        assert len(expiring) >= 2  # Acme (2025-12-31), Columbia (2025-06-30), etc.
        ids = {c.contract_id for c in expiring}
        # Acme expires 2025-12-31, within 12 months of 2025-12-31
        assert "CTR-001" in ids

    def test_contract_ids_unique(self):
        ids = [c.contract_id for c in CONTRACTS]
        assert len(ids) == len(set(ids))


class TestConcentration:
    def test_top_customer_is_acme(self):
        conc = compute_customer_concentration()
        assert conc[0].customer_id == "CUST-001"
        assert conc[0].customer_name == "Acme Manufacturing Corp"

    def test_acme_is_18_pct_of_consolidated(self):
        conc = compute_customer_concentration()
        acme = conc[0]
        # 37.9% of $95M = $36.005M / $200M = 18.0025%
        assert acme.pct_of_consolidated >= Decimal("0.17")
        assert acme.pct_of_consolidated <= Decimal("0.19")

    def test_top_10_concentration(self):
        top10_pct = top_n_concentration(n=10)
        # Top 10 should represent a significant but not 100% share
        assert top10_pct > Decimal("0.50")
        assert top10_pct < Decimal("1.00")

    def test_all_customers_included(self):
        conc = compute_customer_concentration()
        assert len(conc) == 26  # 10 PC + 8 AM + 8 DS

    def test_sorted_descending_by_consolidated_share(self):
        conc = compute_customer_concentration()
        for i in range(len(conc) - 1):
            assert conc[i].pct_of_consolidated >= conc[i + 1].pct_of_consolidated


class TestKeyPersonnel:
    def test_three_executives(self):
        assert len(KEY_PERSONNEL) == 3

    def test_ceo_golden_parachute_3x(self):
        ceo = [kp for kp in KEY_PERSONNEL if "Executive Officer" in kp.title]
        assert len(ceo) == 1
        assert ceo[0].change_of_control_multiplier == Decimal("3")
        assert ceo[0].base_salary == 325_000

    def test_cto_present(self):
        cto = [kp for kp in KEY_PERSONNEL if "Technology Officer" in kp.title]
        assert len(cto) == 1

    def test_cfo_present(self):
        cfo = [kp for kp in KEY_PERSONNEL if "Financial Officer" in kp.title]
        assert len(cfo) == 1

    def test_all_have_coc(self):
        coc_people = key_personnel_with_coc()
        assert len(coc_people) == 3  # All three have multiplier > 0

    def test_total_coc_exposure(self):
        total = total_coc_exposure()
        # CEO: 325K×3 = 975K, CFO: 260K×2 = 520K, CTO: 280K×2 = 560K
        expected = Decimal("2_055_000")
        assert total == expected


class TestLitigation:
    def test_one_pending_matter(self):
        assert len(LITIGATION) == 1

    def test_product_liability(self):
        matter = LITIGATION[0]
        assert matter.case_type == "Product Liability"
        assert matter.potential_exposure == Decimal("2_500_000")

    def test_total_exposure(self):
        assert total_litigation_exposure() == Decimal("2_500_000")


class TestRedFlags:
    def test_red_flags_generated(self):
        flags = data_room_red_flags()
        assert len(flags) >= 3  # litigation, CoC contract, CEO parachute

    def test_high_severity_flags(self):
        flags = data_room_red_flags()
        high = [f for f in flags if f.severity == "high"]
        assert len(high) >= 3  # litigation + CoC contract + CEO parachute

    def test_sorted_by_severity(self):
        flags = data_room_red_flags()
        severity_order = {"high": 0, "medium": 1, "low": 2}
        for i in range(len(flags) - 1):
            assert severity_order[flags[i].severity] <= severity_order[flags[i + 1].severity]

    def test_categories_present(self):
        flags = data_room_red_flags()
        categories = {f.category for f in flags}
        assert "litigation" in categories
        assert "contract" in categories
        assert "personnel" in categories
