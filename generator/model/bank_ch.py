"""Swiss multi-currency bank model for TC-23.

Generates deterministic bank transactions for 3 UBS accounts
(CHF/EUR/USD) for Cascade Precision Instruments AG ("CPI"), Zurich.

Gold-standard reconciliation (all CHF values computed from SNB rates):
  CHF account: adjusted bank = adjusted book = CHF 2,817,905
  EUR account: adjusted bank EUR 1,186,650 / 0.9387 = CHF 1,264,142
               GL CHF = 1,270,466 (stale FX error on opening balance)
  USD account: adjusted bank USD 474,550 / 0.8845 = CHF 536,518
  Consolidated total: CHF 4,618,565

Planted error ERR-CH-002: stale_data -- EUR month-end revaluation of
the opening balance (EUR 540,400) uses FY2024 rate (0.9285) instead of
FY2025 closing rate (0.9387), overstating CHF equivalent by 6,324.
The error is localised to the opening balance portion because December
transactions were booked at spot rates near the closing rate.
"""

from __future__ import annotations

import datetime
import random
from dataclasses import dataclass, field
from decimal import ROUND_HALF_UP, Decimal

# ── Gold-standard targets ────────────────────────────────────────────────

# FX rates (SNB convention: foreign currency per 1 CHF)
FX_CHF_EUR_CLOSING = Decimal("0.9387")   # 1 CHF = 0.9387 EUR (current)
FX_CHF_USD_CLOSING = Decimal("0.8845")   # 1 CHF = 0.8845 USD (current)
FX_CHF_EUR_STALE = Decimal("0.9285")     # FY2024 rate (stale)

# Derived: foreign-to-CHF rates (1 EUR/USD = X CHF)
# 1 EUR = 1/0.9387 CHF ≈ 1.0653 CHF
# 1 USD = 1/0.8845 CHF ≈ 1.1306 CHF
# 1 EUR (stale) = 1/0.9285 CHF ≈ 1.0770 CHF


def _to_chf(amount: Decimal, snb_rate: Decimal) -> Decimal:
    """Convert foreign currency to CHF using SNB rate (FC per 1 CHF).

    CHF = amount / snb_rate, rounded to nearest integer.
    """
    return (amount / snb_rate).quantize(Decimal("1"), rounding=ROUND_HALF_UP)


# ── CHF Account ──────────────────────────────────────────────────────────

CHF_IBAN = "CH39 0027 9279 1234 5678 0"
CHF_BANK_ENDING = Decimal("2847290")
CHF_OUTSTANDING_CHECKS_TOTAL = Decimal("47815")
CHF_DEPOSITS_IN_TRANSIT_TOTAL = Decimal("18430")
CHF_ADJUSTED_BANK = CHF_BANK_ENDING + CHF_DEPOSITS_IN_TRANSIT_TOTAL - CHF_OUTSTANDING_CHECKS_TOTAL
assert CHF_ADJUSTED_BANK == Decimal("2817905"), f"CHF adjusted bank: {CHF_ADJUSTED_BANK}"

CHF_GL_ENDING = Decimal("2815280")
CHF_BANK_CHARGES = Decimal("1375")
CHF_BANK_INTEREST = Decimal("4000")
CHF_ADJUSTED_BOOK = CHF_GL_ENDING + CHF_BANK_INTEREST - CHF_BANK_CHARGES
assert CHF_ADJUSTED_BOOK == Decimal("2817905"), f"CHF adjusted book: {CHF_ADJUSTED_BOOK}"
assert CHF_ADJUSTED_BANK == CHF_ADJUSTED_BOOK

CHF_OUTSTANDING_CHECKS: tuple[tuple[str, Decimal, datetime.date], ...] = (
    ("Scheck Nr. 8847 — Lieferant: Müller Maschinenbau AG", Decimal("28415"), datetime.date(2025, 12, 29)),
    ("Scheck Nr. 8852 — Lieferant: Feinmechanik Bern GmbH", Decimal("19400"), datetime.date(2025, 12, 30)),
)
assert sum(oc[1] for oc in CHF_OUTSTANDING_CHECKS) == CHF_OUTSTANDING_CHECKS_TOTAL

CHF_DEPOSITS_IN_TRANSIT: tuple[tuple[str, Decimal, datetime.date], ...] = (
    ("Einzahlung — Kunde: Schindler Aufzüge AG", Decimal("18430"), datetime.date(2025, 12, 31)),
)
assert sum(d[1] for d in CHF_DEPOSITS_IN_TRANSIT) == CHF_DEPOSITS_IN_TRANSIT_TOTAL

CHF_BANK_STARTING = Decimal("2512340")
CHF_TX_COUNT = 185

# ── EUR Account ──────────────────────────────────────────────────────────

EUR_IBAN = "CH56 0027 9279 2345 6789 0"
EUR_BANK_ENDING = Decimal("1215400")
EUR_OUTSTANDING_SEPA_TOTAL = Decimal("28750")
EUR_ADJUSTED_BANK = EUR_BANK_ENDING - EUR_OUTSTANDING_SEPA_TOTAL
assert EUR_ADJUSTED_BANK == Decimal("1186650"), f"EUR adjusted bank: {EUR_ADJUSTED_BANK}"

# CHF equivalent computed from SNB rate: EUR / SNB_rate = CHF
EUR_ADJUSTED_BANK_CHF = _to_chf(EUR_ADJUSTED_BANK, FX_CHF_EUR_CLOSING)
assert EUR_ADJUSTED_BANK_CHF == Decimal("1264142"), f"EUR adj CHF: {EUR_ADJUSTED_BANK_CHF}"

# ERR-CH-002: stale FX rate on opening balance revaluation.
# The bookkeeper revalued the opening EUR balance (EUR 540,400) at the
# FY2024 closing rate instead of the FY2025 rate.  December transactions
# were booked at spot rates near the closing rate and are unaffected.
EUR_REVAL_BALANCE = Decimal("540400")  # Opening EUR balance subject to stale-rate error
EUR_FX_ERROR = (
    _to_chf(EUR_REVAL_BALANCE, FX_CHF_EUR_STALE)
    - _to_chf(EUR_REVAL_BALANCE, FX_CHF_EUR_CLOSING)
)
assert EUR_FX_ERROR == Decimal("6324"), f"EUR FX error: {EUR_FX_ERROR}"

EUR_GL_ENDING_CHF = EUR_ADJUSTED_BANK_CHF + EUR_FX_ERROR
assert EUR_GL_ENDING_CHF == Decimal("1270466"), f"EUR GL CHF: {EUR_GL_ENDING_CHF}"

EUR_OUTSTANDING_SEPA: tuple[tuple[str, Decimal, datetime.date], ...] = (
    (
        "SEPA-Überweisung Ref SEPA-2025-4891 — Empfänger: Siemens AG München",
        Decimal("28750"),
        datetime.date(2025, 12, 30),
    ),
)
assert sum(s[1] for s in EUR_OUTSTANDING_SEPA) == EUR_OUTSTANDING_SEPA_TOTAL

EUR_BANK_STARTING = Decimal("1089200")
EUR_TX_COUNT = 62

# ── USD Account ──────────────────────────────────────────────────────────

USD_IBAN = "CH83 0027 9279 3456 7890 0"
USD_BANK_ENDING = Decimal("489750")
USD_OUTSTANDING_WIRE_TOTAL = Decimal("15200")
USD_ADJUSTED_BANK = USD_BANK_ENDING - USD_OUTSTANDING_WIRE_TOTAL
assert USD_ADJUSTED_BANK == Decimal("474550"), f"USD adjusted bank: {USD_ADJUSTED_BANK}"

# CHF equivalent computed from SNB rate: USD / SNB_rate = CHF
USD_ADJUSTED_BANK_CHF = _to_chf(USD_ADJUSTED_BANK, FX_CHF_USD_CLOSING)
assert USD_ADJUSTED_BANK_CHF == Decimal("536518"), f"USD adj CHF: {USD_ADJUSTED_BANK_CHF}"

USD_GL_ENDING_CHF = USD_ADJUSTED_BANK_CHF  # No error on USD
assert USD_GL_ENDING_CHF == USD_ADJUSTED_BANK_CHF

USD_OUTSTANDING_WIRE: tuple[tuple[str, Decimal, datetime.date], ...] = (
    ("SWIFT MT103 Ref TRF-2025-7832 — Beneficiary: Honeywell Inc.", Decimal("15200"), datetime.date(2025, 12, 31)),
)
assert sum(w[1] for w in USD_OUTSTANDING_WIRE) == USD_OUTSTANDING_WIRE_TOTAL

USD_BANK_STARTING = Decimal("462800")
USD_TX_COUNT = 28

# ── Consolidated ─────────────────────────────────────────────────────────

CONSOLIDATED_CHF = CHF_ADJUSTED_BANK + EUR_ADJUSTED_BANK_CHF + USD_ADJUSTED_BANK_CHF
assert CONSOLIDATED_CHF == Decimal("4618565"), f"Consolidated: {CONSOLIDATED_CHF}"

# ── Entity profile ───────────────────────────────────────────────────────

ENTITY_NAME = "Cascade Precision Instruments AG"
ENTITY_SHORT = "CPI"
ENTITY_CITY = "Zürich"
ENTITY_COUNTRY = "Schweiz"
BANK_NAME = "UBS Switzerland AG"
CREDIT_FACILITY = Decimal("5000000")  # CHF 5M revolving, undrawn

# ── Transaction templates ────────────────────────────────────────────────


@dataclass(frozen=True)
class TransactionTemplate:
    """Template for generating bank/GL transaction pairs."""

    category: str
    bank_desc_prefix: str
    gl_desc_prefix: str
    min_amount: int
    max_amount: int
    is_debit: bool  # True = money out, False = money in


# CHF templates (Swiss domestic)
_CHF_INFLOW_TEMPLATES: tuple[TransactionTemplate, ...] = (
    TransactionTemplate("lsv_collection", "LSV+ Einzug", "LSV+ Einzug —", 5000, 85000, False),
    TransactionTemplate("sic_inflow", "SIC Gutschrift", "SIC Zahlungseingang —", 10000, 120000, False),
    TransactionTemplate("deposit", "Bareinzahlung", "Bareinzahlung", 5000, 35000, False),
    TransactionTemplate("dta_inflow", "DTA Gutschrift", "DTA Zahlungseingang —", 3000, 45000, False),
    TransactionTemplate("qr_payment", "QR-Rechnung Einzahlung", "QR-Zahlung erhalten —", 2000, 25000, False),
)

_CHF_OUTFLOW_TEMPLATES: tuple[TransactionTemplate, ...] = (
    TransactionTemplate("dta_payment", "DTA-Zahlung", "DTA-Zahlung —", 3000, 50000, True),
    TransactionTemplate("check", "Scheck Nr.", "Scheck —", 2000, 30000, True),
    TransactionTemplate("lsv_debit", "LSV+ Lastschrift", "LSV+ Lastschrift —", 1000, 15000, True),
    TransactionTemplate("salary", "Lohnzahlung", "Lohnzahlung — Monatslohn", 180000, 280000, True),
    TransactionTemplate("tax_payment", "Steuerzahlung", "Steuerzahlung —", 5000, 45000, True),
    TransactionTemplate("rent", "Miete Büroräume", "Miete —", 8000, 25000, True),
    TransactionTemplate("insurance", "Versicherungsprämie", "Versicherung —", 2000, 12000, True),
)

# EUR templates (SEPA / EU)
_EUR_INFLOW_TEMPLATES: tuple[TransactionTemplate, ...] = (
    TransactionTemplate("sepa_credit", "SEPA Gutschrift", "SEPA Zahlungseingang —", 8000, 95000, False),
    TransactionTemplate("swift_in", "SWIFT Gutschrift MT103", "SWIFT Eingang —", 15000, 80000, False),
)

_EUR_OUTFLOW_TEMPLATES: tuple[TransactionTemplate, ...] = (
    TransactionTemplate("sepa_debit", "SEPA Überweisung", "SEPA-Zahlung —", 5000, 60000, True),
    TransactionTemplate("swift_out", "SWIFT Überweisung MT103", "SWIFT Zahlung —", 10000, 50000, True),
)

# USD templates (international wire)
_USD_INFLOW_TEMPLATES: tuple[TransactionTemplate, ...] = (
    TransactionTemplate("swift_in", "SWIFT Credit MT103", "SWIFT Receipt —", 10000, 65000, False),
    TransactionTemplate("wire_in", "Wire Transfer Credit", "Wire Received —", 8000, 40000, False),
)

_USD_OUTFLOW_TEMPLATES: tuple[TransactionTemplate, ...] = (
    TransactionTemplate("swift_out", "SWIFT Debit MT103", "SWIFT Payment —", 8000, 45000, True),
    TransactionTemplate("wire_out", "Wire Transfer Debit", "Wire Payment —", 5000, 30000, True),
)

# ── Swiss vendor/customer names ──────────────────────────────────────────

_CH_VENDOR_NAMES: tuple[str, ...] = (
    "Müller Maschinenbau AG", "Feinmechanik Bern GmbH", "Zürcher Metallwerke AG",
    "Bühler Präzisionstechnik", "Schweizer Elektronik AG", "Basler Chemie AG",
    "Luzerner Werkzeugbau", "Aargauer Stahlhandel", "Thurgauer Kunststoff GmbH",
    "Solothurner Oberflächentechnik", "Graubündner Logistik AG", "St. Galler Verpackung",
    "Winterthurer Druckerei", "Schaffhauser Metallbau", "Berner Bürobedarf AG",
)

_CH_CUSTOMER_NAMES: tuple[str, ...] = (
    "Schindler Aufzüge AG", "ABB Schweiz AG", "Hilti AG",
    "Georg Fischer AG", "Sulzer AG", "Oerlikon Metco AG",
    "Bühler Group AG", "Endress+Hauser AG", "Mettler-Toledo GmbH",
    "Rieter AG", "Burckhardt Compression AG", "Stadler Rail AG",
    "Bobst Group SA", "Sika AG", "Sonova Holding AG",
)

_EU_VENDOR_NAMES: tuple[str, ...] = (
    "Siemens AG München", "Bosch GmbH Stuttgart", "KUKA Roboter GmbH",
    "Trumpf GmbH Ditzingen", "Festo SE Esslingen", "Schneider Electric SA",
    "Dassault Systèmes SE", "Atlas Copco AB", "SKF Group AB",
    "Philips NV Eindhoven",
)

_EU_CUSTOMER_NAMES: tuple[str, ...] = (
    "BMW AG München", "Airbus SE Toulouse", "Thales Group SA",
    "Volvo AB Göteborg", "Ericsson AB Stockholm", "Nokia Oyj Helsinki",
    "ASML NV Veldhoven", "Infineon AG München", "Continental AG Hannover",
    "Rheinmetall AG Düsseldorf",
)

_US_VENDOR_NAMES: tuple[str, ...] = (
    "Honeywell Inc.", "Parker Hannifin Corp", "Emerson Electric Co",
    "Rockwell Automation", "Illinois Tool Works", "Eaton Corporation",
    "Dover Corporation", "Fortive Corporation",
)

_US_CUSTOMER_NAMES: tuple[str, ...] = (
    "General Electric Co", "Caterpillar Inc.", "John Deere & Co",
    "3M Company", "Raytheon Technologies", "Northrop Grumman Corp",
    "L3Harris Technologies", "Textron Inc.",
)


# ── Data classes ─────────────────────────────────────────────────────────

@dataclass
class BankTransactionCH:
    """A single transaction on a Swiss bank statement."""

    date: datetime.date
    description: str
    amount: Decimal  # Positive = credit (inflow), negative = debit (outflow)
    running_balance: Decimal
    currency: str = "CHF"
    category: str = ""
    matched: bool = True


@dataclass
class GLCashEntryCH:
    """A single line in the GL cash detail for a Swiss account.

    For CHF account: amounts in CHF.
    For EUR/USD accounts: amounts in CHF (functional currency) at spot rate.
    """

    date: datetime.date
    reference: str  # e.g., "BU-2025-1234"
    description: str
    debit: Decimal  # Soll (money in)
    credit: Decimal  # Haben (money out)
    running_balance: Decimal  # Saldo
    currency: str = "CHF"  # Always CHF for GL
    original_currency: str = "CHF"
    original_amount: Decimal = Decimal("0")
    fx_rate: Decimal = Decimal("1")  # SNB rate used for conversion
    category: str = ""
    matched: bool = True


@dataclass
class AccountReconciliation:
    """Reconciliation data for a single account."""

    currency: str
    iban: str
    bank_ending: Decimal
    outstanding_items: list[tuple[str, Decimal, datetime.date]]
    outstanding_total: Decimal
    deposits_in_transit: list[tuple[str, Decimal, datetime.date]]
    deposits_in_transit_total: Decimal
    adjusted_bank: Decimal
    adjusted_bank_chf: Decimal
    gl_ending_chf: Decimal
    bank_charges: Decimal = Decimal("0")
    bank_interest: Decimal = Decimal("0")
    fx_rate: Decimal = Decimal("1")
    fx_error: Decimal = Decimal("0")


@dataclass
class BankModelCH:
    """Complete Swiss multi-currency bank model for TC-23."""

    # CHF account
    chf_bank_transactions: list[BankTransactionCH] = field(default_factory=list)
    chf_gl_entries: list[GLCashEntryCH] = field(default_factory=list)
    chf_bank_starting: Decimal = CHF_BANK_STARTING
    chf_bank_ending: Decimal = CHF_BANK_ENDING
    chf_gl_starting: Decimal = Decimal("0")
    chf_gl_ending: Decimal = CHF_GL_ENDING

    # EUR account
    eur_bank_transactions: list[BankTransactionCH] = field(default_factory=list)
    eur_gl_entries: list[GLCashEntryCH] = field(default_factory=list)
    eur_bank_starting: Decimal = EUR_BANK_STARTING
    eur_bank_ending: Decimal = EUR_BANK_ENDING
    eur_gl_starting_chf: Decimal = Decimal("0")
    eur_gl_ending_chf: Decimal = EUR_GL_ENDING_CHF

    # USD account
    usd_bank_transactions: list[BankTransactionCH] = field(default_factory=list)
    usd_gl_entries: list[GLCashEntryCH] = field(default_factory=list)
    usd_bank_starting: Decimal = USD_BANK_STARTING
    usd_bank_ending: Decimal = USD_BANK_ENDING
    usd_gl_starting_chf: Decimal = Decimal("0")
    usd_gl_ending_chf: Decimal = USD_GL_ENDING_CHF

    # Reconciliation summaries
    chf_recon: AccountReconciliation | None = None
    eur_recon: AccountReconciliation | None = None
    usd_recon: AccountReconciliation | None = None

    # Consolidated
    consolidated_chf: Decimal = CONSOLIDATED_CHF

    # FX rates
    fx_chf_eur_closing: Decimal = FX_CHF_EUR_CLOSING
    fx_chf_usd_closing: Decimal = FX_CHF_USD_CLOSING
    fx_chf_eur_stale: Decimal = FX_CHF_EUR_STALE

    # Planted error
    eur_fx_error: Decimal = EUR_FX_ERROR


# ── Helpers ──────────────────────────────────────────────────────────────

def _business_days_offset(date: datetime.date, offset: int) -> datetime.date:
    """Shift a date by offset business days (skipping weekends)."""
    direction = 1 if offset >= 0 else -1
    remaining = abs(offset)
    current = date
    while remaining > 0:
        current += datetime.timedelta(days=direction)
        if current.weekday() < 5:
            remaining -= 1
    return current


def _dec_business_days() -> list[datetime.date]:
    """Return all business days in December 2025."""
    days: list[datetime.date] = []
    d = datetime.date(2025, 12, 1)
    while d <= datetime.date(2025, 12, 31):
        if d.weekday() < 5:
            days.append(d)
        d += datetime.timedelta(days=1)
    return days


def _approximate_daily_rate(
    month_day: int,
    closing_snb: Decimal,
    *,
    start_snb: Decimal | None = None,
) -> Decimal:
    """Generate a plausible intra-month SNB rate by linear interpolation.

    Interpolates between start_snb (Dec 1) and closing_snb (Dec 31).
    If start_snb is None, uses a rate ~0.5% away from closing.
    """
    if start_snb is None:
        start_snb = closing_snb * Decimal("1.005")
    # Linear interpolation: day 1 = start, day 31 = closing
    fraction = Decimal(str((month_day - 1) / 30))
    rate = start_snb + (closing_snb - start_snb) * fraction
    return rate.quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP)


# ── Generator: CHF account ───────────────────────────────────────────────

def _generate_chf_account(
    rng: random.Random,
    model: BankModelCH,
) -> None:
    """Generate 185 bank transactions and matching GL for the CHF account."""
    business_days = _dec_business_days()

    # Bank-only items: interest and charges on Dec 31
    n_bank_only = 2  # interest + charges
    n_matched = CHF_TX_COUNT - n_bank_only  # 183

    # Net of matched transactions: ending - starting - interest + charges
    net_matched = (
        CHF_BANK_ENDING
        - CHF_BANK_STARTING
        - CHF_BANK_INTEREST
        + CHF_BANK_CHARGES
    )

    # Generate matched amounts
    matched_amounts: list[Decimal] = []
    matched_categories: list[str] = []
    matched_is_debit: list[bool] = []

    # One salary payment
    salary = Decimal("235480")
    matched_amounts.append(-salary)
    matched_categories.append("salary")
    matched_is_debit.append(True)

    remaining_count = n_matched - len(matched_amounts)
    remaining_net = net_matched - sum(matched_amounts)

    n_inflows = int(remaining_count * 0.55)
    n_outflows = remaining_count - n_inflows

    inflows: list[tuple[Decimal, str]] = []
    inflow_total = Decimal("0")
    for _ in range(n_inflows):
        tmpl = _CHF_INFLOW_TEMPLATES[rng.randint(0, len(_CHF_INFLOW_TEMPLATES) - 1)]
        amount = Decimal(str(rng.randint(tmpl.min_amount, tmpl.max_amount)))
        inflows.append((amount, tmpl.category))
        inflow_total += amount

    outflows: list[tuple[Decimal, str]] = []
    outflow_total = Decimal("0")
    for _ in range(n_outflows):
        tmpl = _CHF_OUTFLOW_TEMPLATES[rng.randint(0, len(_CHF_OUTFLOW_TEMPLATES) - 1)]
        if tmpl.category == "salary":
            tmpl = _CHF_OUTFLOW_TEMPLATES[0]  # Use DTA instead
        amount = Decimal(str(rng.randint(tmpl.min_amount, tmpl.max_amount)))
        outflows.append((amount, tmpl.category))
        outflow_total += amount

    generated_net = inflow_total - outflow_total
    adjustment = remaining_net - generated_net

    if inflows and adjustment != Decimal("0"):
        total_inflow = sum(a for a, _ in inflows)
        if total_inflow > 0:
            remaining_adj = adjustment
            for i in range(len(inflows)):
                if i == len(inflows) - 1:
                    share = remaining_adj
                else:
                    share = (inflows[i][0] * adjustment / total_inflow).to_integral_value()
                    remaining_adj -= share
                old_amount, cat = inflows[i]
                inflows[i] = (old_amount + share, cat)

    for amount, cat in inflows:
        matched_amounts.append(amount)
        matched_categories.append(cat)
        matched_is_debit.append(False)

    for amount, cat in outflows:
        matched_amounts.append(-amount)
        matched_categories.append(cat)
        matched_is_debit.append(True)

    assert sum(matched_amounts) == net_matched, (
        f"CHF net mismatch: {sum(matched_amounts)} != {net_matched}"
    )

    # Assign dates
    bank_dates: list[datetime.date] = []
    for i in range(n_matched):
        if i == 0:
            bank_dates.append(datetime.date(2025, 12, 25))  # Salary
        else:
            bank_dates.append(business_days[rng.randint(0, len(business_days) - 1)])

    combined = list(zip(bank_dates, matched_amounts, matched_categories, matched_is_debit))
    combined.sort(key=lambda x: (x[0], x[2], str(x[1])))
    bank_dates = [c[0] for c in combined]
    matched_amounts = [c[1] for c in combined]
    matched_categories = [c[2] for c in combined]
    matched_is_debit = [c[3] for c in combined]

    # Build bank statement
    running = CHF_BANK_STARTING
    interest_inserted = False
    vendor_idx = 0
    customer_idx = 0
    bu_counter = 2025_12_001  # Beleg-Nr (voucher number)

    for i in range(n_matched):
        amount = matched_amounts[i]
        category = matched_categories[i]
        is_debit = matched_is_debit[i]
        bank_date = bank_dates[i]

        if bank_date >= datetime.date(2025, 12, 31) and not interest_inserted:
            # Interest credit
            running += CHF_BANK_INTEREST
            model.chf_bank_transactions.append(BankTransactionCH(
                date=datetime.date(2025, 12, 31),
                description="Zinsgutschrift Kontokorrent",
                amount=CHF_BANK_INTEREST,
                running_balance=running,
                currency="CHF",
                category="interest",
                matched=False,
            ))
            # Bank charges
            running -= CHF_BANK_CHARGES
            model.chf_bank_transactions.append(BankTransactionCH(
                date=datetime.date(2025, 12, 31),
                description="Kontoführungsgebühr Dezember 2025",
                amount=-CHF_BANK_CHARGES,
                running_balance=running,
                currency="CHF",
                category="service_charge",
                matched=False,
            ))
            interest_inserted = True

        # Bank description
        bank_desc = _make_chf_bank_desc(
            category, is_debit, rng, vendor_idx, customer_idx,
        )
        if is_debit and category in ("dta_payment", "check"):
            vendor_idx = (vendor_idx + 1) % len(_CH_VENDOR_NAMES)
        elif not is_debit and category in ("lsv_collection", "sic_inflow", "dta_inflow", "qr_payment"):
            customer_idx = (customer_idx + 1) % len(_CH_CUSTOMER_NAMES)

        running += amount
        model.chf_bank_transactions.append(BankTransactionCH(
            date=bank_date,
            description=bank_desc,
            amount=amount,
            running_balance=running,
            currency="CHF",
            category=category,
            matched=True,
        ))

        # GL entry
        offset_days = rng.choice([0, 0, 0, 0, 0, -1, -1, 1, 1, -2, 2])
        gl_date = _business_days_offset(bank_date, offset_days)
        gl_date = max(gl_date, datetime.date(2025, 12, 1))
        gl_date = min(gl_date, datetime.date(2025, 12, 31))

        gl_desc = _make_chf_gl_desc(
            category, is_debit, rng,
            (vendor_idx - 1) % len(_CH_VENDOR_NAMES) if is_debit else vendor_idx,
            (customer_idx - 1) % len(_CH_CUSTOMER_NAMES) if not is_debit else customer_idx,
        )

        ref = f"BU-{bu_counter}"
        bu_counter += 1

        if is_debit:
            model.chf_gl_entries.append(GLCashEntryCH(
                date=gl_date, reference=ref, description=gl_desc,
                debit=Decimal("0"), credit=abs(amount),
                running_balance=Decimal("0"),
                currency="CHF", original_currency="CHF",
                original_amount=abs(amount), fx_rate=Decimal("1"),
                category=category, matched=True,
            ))
        else:
            model.chf_gl_entries.append(GLCashEntryCH(
                date=gl_date, reference=ref, description=gl_desc,
                debit=amount, credit=Decimal("0"),
                running_balance=Decimal("0"),
                currency="CHF", original_currency="CHF",
                original_amount=amount, fx_rate=Decimal("1"),
                category=category, matched=True,
            ))

    # Insert bank-only if not yet done
    if not interest_inserted:
        running += CHF_BANK_INTEREST
        model.chf_bank_transactions.append(BankTransactionCH(
            date=datetime.date(2025, 12, 31),
            description="Zinsgutschrift Kontokorrent",
            amount=CHF_BANK_INTEREST,
            running_balance=running,
            currency="CHF",
            category="interest",
            matched=False,
        ))
        running -= CHF_BANK_CHARGES
        model.chf_bank_transactions.append(BankTransactionCH(
            date=datetime.date(2025, 12, 31),
            description="Kontoführungsgebühr Dezember 2025",
            amount=-CHF_BANK_CHARGES,
            running_balance=running,
            currency="CHF",
            category="service_charge",
            matched=False,
        ))

    assert running == CHF_BANK_ENDING, f"CHF bank ending: {running} != {CHF_BANK_ENDING}"
    assert len(model.chf_bank_transactions) == CHF_TX_COUNT, (
        f"CHF tx count: {len(model.chf_bank_transactions)} != {CHF_TX_COUNT}"
    )

    # Outstanding checks → GL only
    for desc, amount, date in CHF_OUTSTANDING_CHECKS:
        ref = f"BU-{bu_counter}"
        bu_counter += 1
        model.chf_gl_entries.append(GLCashEntryCH(
            date=date, reference=ref, description=desc,
            debit=Decimal("0"), credit=amount,
            running_balance=Decimal("0"),
            currency="CHF", original_currency="CHF",
            original_amount=amount, fx_rate=Decimal("1"),
            category="outstanding_check", matched=False,
        ))

    # Deposits in transit → GL only
    for desc, amount, date in CHF_DEPOSITS_IN_TRANSIT:
        ref = f"BU-{bu_counter}"
        bu_counter += 1
        model.chf_gl_entries.append(GLCashEntryCH(
            date=date, reference=ref, description=desc,
            debit=amount, credit=Decimal("0"),
            running_balance=Decimal("0"),
            currency="CHF", original_currency="CHF",
            original_amount=amount, fx_rate=Decimal("1"),
            category="deposit_in_transit", matched=False,
        ))

    # Sort GL and compute running balances
    model.chf_gl_entries.sort(key=lambda e: (e.date, e.reference))
    gl_net = sum(e.debit - e.credit for e in model.chf_gl_entries)
    model.chf_gl_starting = CHF_GL_ENDING - gl_net

    gl_running = model.chf_gl_starting
    for entry in model.chf_gl_entries:
        gl_running += entry.debit - entry.credit
        entry.running_balance = gl_running

    assert gl_running == CHF_GL_ENDING, f"CHF GL ending: {gl_running} != {CHF_GL_ENDING}"


# ── Generator: EUR account ───────────────────────────────────────────────

def _generate_eur_account(
    rng: random.Random,
    model: BankModelCH,
) -> None:
    """Generate 62 EUR bank transactions and matching GL (in CHF) entries.

    GL entries use intra-month spot rates for transaction-date conversions,
    plus a month-end revaluation entry. The revaluation intentionally uses
    the stale FY2024 rate (ERR-CH-002).
    """
    business_days = _dec_business_days()

    # EUR bank: 62 transactions, 1 outstanding SEPA
    n_matched = EUR_TX_COUNT  # All 62 appear on bank statement

    net_matched = EUR_BANK_ENDING - EUR_BANK_STARTING

    matched_amounts: list[Decimal] = []
    matched_categories: list[str] = []
    matched_is_debit: list[bool] = []

    remaining_count = n_matched
    n_inflows = int(remaining_count * 0.55)
    n_outflows = remaining_count - n_inflows

    inflows: list[tuple[Decimal, str]] = []
    inflow_total = Decimal("0")
    for _ in range(n_inflows):
        tmpl = _EUR_INFLOW_TEMPLATES[rng.randint(0, len(_EUR_INFLOW_TEMPLATES) - 1)]
        amount = Decimal(str(rng.randint(tmpl.min_amount, tmpl.max_amount)))
        inflows.append((amount, tmpl.category))
        inflow_total += amount

    outflows: list[tuple[Decimal, str]] = []
    outflow_total = Decimal("0")
    for _ in range(n_outflows):
        tmpl = _EUR_OUTFLOW_TEMPLATES[rng.randint(0, len(_EUR_OUTFLOW_TEMPLATES) - 1)]
        amount = Decimal(str(rng.randint(tmpl.min_amount, tmpl.max_amount)))
        outflows.append((amount, tmpl.category))
        outflow_total += amount

    generated_net = inflow_total - outflow_total
    adjustment = net_matched - generated_net

    if inflows and adjustment != Decimal("0"):
        total_inflow = sum(a for a, _ in inflows)
        if total_inflow > 0:
            remaining_adj = adjustment
            for i in range(len(inflows)):
                if i == len(inflows) - 1:
                    share = remaining_adj
                else:
                    share = (inflows[i][0] * adjustment / total_inflow).to_integral_value()
                    remaining_adj -= share
                old_amount, cat = inflows[i]
                inflows[i] = (old_amount + share, cat)

    for amount, cat in inflows:
        matched_amounts.append(amount)
        matched_categories.append(cat)
        matched_is_debit.append(False)

    for amount, cat in outflows:
        matched_amounts.append(-amount)
        matched_categories.append(cat)
        matched_is_debit.append(True)

    assert sum(matched_amounts) == net_matched, (
        f"EUR net mismatch: {sum(matched_amounts)} != {net_matched}"
    )

    # Assign dates
    bank_dates: list[datetime.date] = [
        business_days[rng.randint(0, len(business_days) - 1)]
        for _ in range(n_matched)
    ]

    combined = list(zip(bank_dates, matched_amounts, matched_categories, matched_is_debit))
    combined.sort(key=lambda x: (x[0], x[2], str(x[1])))
    bank_dates = [c[0] for c in combined]
    matched_amounts = [c[1] for c in combined]
    matched_categories = [c[2] for c in combined]
    matched_is_debit = [c[3] for c in combined]

    # Build bank statement
    running = EUR_BANK_STARTING
    vendor_idx = 0
    customer_idx = 0
    bu_counter = 2025_12_501  # Separate range for EUR

    # Track CHF equivalents for GL running balance
    gl_entries_pre_reval: list[GLCashEntryCH] = []

    # Start rate for EUR (slightly different from closing for realistic interpolation)
    eur_start_snb = FX_CHF_EUR_CLOSING * Decimal("1.004")

    for i in range(n_matched):
        amount = matched_amounts[i]
        category = matched_categories[i]
        is_debit = matched_is_debit[i]
        bank_date = bank_dates[i]

        bank_desc = _make_eur_bank_desc(
            category, is_debit, rng, vendor_idx, customer_idx,
        )
        if is_debit:
            vendor_idx = (vendor_idx + 1) % len(_EU_VENDOR_NAMES)
        else:
            customer_idx = (customer_idx + 1) % len(_EU_CUSTOMER_NAMES)

        running += amount
        model.eur_bank_transactions.append(BankTransactionCH(
            date=bank_date,
            description=bank_desc,
            amount=amount,
            running_balance=running,
            currency="EUR",
            category=category,
            matched=True,
        ))

        # GL entry in CHF at transaction-date spot rate
        offset_days = rng.choice([0, 0, 0, 0, 0, -1, 1])
        gl_date = _business_days_offset(bank_date, offset_days)
        gl_date = max(gl_date, datetime.date(2025, 12, 1))
        gl_date = min(gl_date, datetime.date(2025, 12, 31))

        # Spot rate for this date
        spot_snb = _approximate_daily_rate(
            bank_date.day, FX_CHF_EUR_CLOSING, start_snb=eur_start_snb,
        )
        chf_amount = _to_chf(abs(amount), spot_snb)

        gl_desc = _make_eur_gl_desc(
            category, is_debit, rng,
            (vendor_idx - 1) % len(_EU_VENDOR_NAMES),
            (customer_idx - 1) % len(_EU_CUSTOMER_NAMES),
        )

        ref = f"BU-{bu_counter}"
        bu_counter += 1

        entry = GLCashEntryCH(
            date=gl_date, reference=ref, description=gl_desc,
            debit=Decimal("0") if is_debit else chf_amount,
            credit=chf_amount if is_debit else Decimal("0"),
            running_balance=Decimal("0"),
            currency="CHF", original_currency="EUR",
            original_amount=abs(amount), fx_rate=spot_snb,
            category=category, matched=True,
        )
        gl_entries_pre_reval.append(entry)

    assert running == EUR_BANK_ENDING, f"EUR bank ending: {running} != {EUR_BANK_ENDING}"
    assert len(model.eur_bank_transactions) == EUR_TX_COUNT, (
        f"EUR tx count: {len(model.eur_bank_transactions)} != {EUR_TX_COUNT}"
    )

    # Outstanding SEPA → GL only (in CHF at spot rate near end of month)
    for desc, amount, date in EUR_OUTSTANDING_SEPA:
        ref = f"BU-{bu_counter}"
        bu_counter += 1
        spot_snb = _approximate_daily_rate(
            date.day, FX_CHF_EUR_CLOSING, start_snb=eur_start_snb,
        )
        chf_amount = _to_chf(amount, spot_snb)
        gl_entries_pre_reval.append(GLCashEntryCH(
            date=date, reference=ref, description=desc,
            debit=Decimal("0"), credit=chf_amount,
            running_balance=Decimal("0"),
            currency="CHF", original_currency="EUR",
            original_amount=amount, fx_rate=spot_snb,
            category="outstanding_sepa", matched=False,
        ))

    # Compute pre-revaluation GL balance
    gl_net_pre_reval = sum(e.debit - e.credit for e in gl_entries_pre_reval)

    # Month-end FX revaluation entry:
    # The EUR account has EUR_ADJUSTED_BANK (1,186,650) at month end.
    # Correct CHF = 1,264,142 (closing rate).
    # ERR-CH-002: the bookkeeper revalued the opening balance (EUR 540,400)
    # at the stale FY2024 rate 0.9285, adding CHF 6,324 overstatement.
    # GL ending target = 1,264,142 + 6,324 = 1,270,466.
    # The revaluation entry is the plug to make GL ending = 1,270,466.

    # EUR starting balance in CHF: approximate EUR_BANK_STARTING at a reasonable rate
    eur_gl_starting_approx = _to_chf(EUR_BANK_STARTING, FX_CHF_EUR_STALE)  # Opening at old rate

    # reval = gl_ending - gl_starting - gl_net_pre_reval
    reval_amount = EUR_GL_ENDING_CHF - eur_gl_starting_approx - gl_net_pre_reval

    # Add revaluation entry (using stale rate — this IS the planted error)
    reval_ref = f"BU-{bu_counter}"
    bu_counter += 1
    reval_desc = "FX-Neubewertung EUR per 31.12.2025 (Kurs CHF/EUR 0.9285)"
    if reval_amount >= 0:
        gl_entries_pre_reval.append(GLCashEntryCH(
            date=datetime.date(2025, 12, 31),
            reference=reval_ref,
            description=reval_desc,
            debit=reval_amount,
            credit=Decimal("0"),
            running_balance=Decimal("0"),
            currency="CHF", original_currency="EUR",
            original_amount=Decimal("0"),
            fx_rate=FX_CHF_EUR_STALE,
            category="fx_revaluation", matched=False,
        ))
    else:
        gl_entries_pre_reval.append(GLCashEntryCH(
            date=datetime.date(2025, 12, 31),
            reference=reval_ref,
            description=reval_desc,
            debit=Decimal("0"),
            credit=abs(reval_amount),
            running_balance=Decimal("0"),
            currency="CHF", original_currency="EUR",
            original_amount=Decimal("0"),
            fx_rate=FX_CHF_EUR_STALE,
            category="fx_revaluation", matched=False,
        ))

    # Sort and compute running balance
    gl_entries_pre_reval.sort(key=lambda e: (e.date, e.reference))
    model.eur_gl_entries = gl_entries_pre_reval
    model.eur_gl_starting_chf = eur_gl_starting_approx

    gl_running = model.eur_gl_starting_chf
    for entry in model.eur_gl_entries:
        gl_running += entry.debit - entry.credit
        entry.running_balance = gl_running

    assert gl_running == EUR_GL_ENDING_CHF, (
        f"EUR GL ending: {gl_running} != {EUR_GL_ENDING_CHF}"
    )


# ── Generator: USD account ───────────────────────────────────────────────

def _generate_usd_account(
    rng: random.Random,
    model: BankModelCH,
) -> None:
    """Generate 28 USD bank transactions and matching GL (in CHF) entries."""
    business_days = _dec_business_days()

    n_matched = USD_TX_COUNT  # 28

    net_matched = USD_BANK_ENDING - USD_BANK_STARTING

    matched_amounts: list[Decimal] = []
    matched_categories: list[str] = []
    matched_is_debit: list[bool] = []

    n_inflows = int(n_matched * 0.55)
    n_outflows = n_matched - n_inflows

    inflows: list[tuple[Decimal, str]] = []
    inflow_total = Decimal("0")
    for _ in range(n_inflows):
        tmpl = _USD_INFLOW_TEMPLATES[rng.randint(0, len(_USD_INFLOW_TEMPLATES) - 1)]
        amount = Decimal(str(rng.randint(tmpl.min_amount, tmpl.max_amount)))
        inflows.append((amount, tmpl.category))
        inflow_total += amount

    outflows: list[tuple[Decimal, str]] = []
    outflow_total = Decimal("0")
    for _ in range(n_outflows):
        tmpl = _USD_OUTFLOW_TEMPLATES[rng.randint(0, len(_USD_OUTFLOW_TEMPLATES) - 1)]
        amount = Decimal(str(rng.randint(tmpl.min_amount, tmpl.max_amount)))
        outflows.append((amount, tmpl.category))
        outflow_total += amount

    generated_net = inflow_total - outflow_total
    adjustment = net_matched - generated_net

    if inflows and adjustment != Decimal("0"):
        total_inflow = sum(a for a, _ in inflows)
        if total_inflow > 0:
            remaining_adj = adjustment
            for i in range(len(inflows)):
                if i == len(inflows) - 1:
                    share = remaining_adj
                else:
                    share = (inflows[i][0] * adjustment / total_inflow).to_integral_value()
                    remaining_adj -= share
                old_amount, cat = inflows[i]
                inflows[i] = (old_amount + share, cat)

    for amount, cat in inflows:
        matched_amounts.append(amount)
        matched_categories.append(cat)
        matched_is_debit.append(False)

    for amount, cat in outflows:
        matched_amounts.append(-amount)
        matched_categories.append(cat)
        matched_is_debit.append(True)

    assert sum(matched_amounts) == net_matched, (
        f"USD net mismatch: {sum(matched_amounts)} != {net_matched}"
    )

    # Assign dates
    bank_dates: list[datetime.date] = [
        business_days[rng.randint(0, len(business_days) - 1)]
        for _ in range(n_matched)
    ]

    combined = list(zip(bank_dates, matched_amounts, matched_categories, matched_is_debit))
    combined.sort(key=lambda x: (x[0], x[2], str(x[1])))
    bank_dates = [c[0] for c in combined]
    matched_amounts = [c[1] for c in combined]
    matched_categories = [c[2] for c in combined]
    matched_is_debit = [c[3] for c in combined]

    # Build bank statement
    running = USD_BANK_STARTING
    vendor_idx = 0
    customer_idx = 0
    bu_counter = 2025_12_801  # Separate range for USD

    gl_entries_pre_reval: list[GLCashEntryCH] = []
    usd_start_snb = FX_CHF_USD_CLOSING * Decimal("1.003")

    for i in range(n_matched):
        amount = matched_amounts[i]
        category = matched_categories[i]
        is_debit = matched_is_debit[i]
        bank_date = bank_dates[i]

        bank_desc = _make_usd_bank_desc(
            category, is_debit, rng, vendor_idx, customer_idx,
        )
        if is_debit:
            vendor_idx = (vendor_idx + 1) % len(_US_VENDOR_NAMES)
        else:
            customer_idx = (customer_idx + 1) % len(_US_CUSTOMER_NAMES)

        running += amount
        model.usd_bank_transactions.append(BankTransactionCH(
            date=bank_date,
            description=bank_desc,
            amount=amount,
            running_balance=running,
            currency="USD",
            category=category,
            matched=True,
        ))

        # GL entry in CHF
        offset_days = rng.choice([0, 0, 0, 0, 0, -1, 1])
        gl_date = _business_days_offset(bank_date, offset_days)
        gl_date = max(gl_date, datetime.date(2025, 12, 1))
        gl_date = min(gl_date, datetime.date(2025, 12, 31))

        spot_snb = _approximate_daily_rate(
            bank_date.day, FX_CHF_USD_CLOSING, start_snb=usd_start_snb,
        )
        chf_amount = _to_chf(abs(amount), spot_snb)

        gl_desc = _make_usd_gl_desc(
            category, is_debit, rng,
            (vendor_idx - 1) % len(_US_VENDOR_NAMES),
            (customer_idx - 1) % len(_US_CUSTOMER_NAMES),
        )

        ref = f"BU-{bu_counter}"
        bu_counter += 1

        entry = GLCashEntryCH(
            date=gl_date, reference=ref, description=gl_desc,
            debit=Decimal("0") if is_debit else chf_amount,
            credit=chf_amount if is_debit else Decimal("0"),
            running_balance=Decimal("0"),
            currency="CHF", original_currency="USD",
            original_amount=abs(amount), fx_rate=spot_snb,
            category=category, matched=True,
        )
        gl_entries_pre_reval.append(entry)

    assert running == USD_BANK_ENDING, f"USD bank ending: {running} != {USD_BANK_ENDING}"
    assert len(model.usd_bank_transactions) == USD_TX_COUNT, (
        f"USD tx count: {len(model.usd_bank_transactions)} != {USD_TX_COUNT}"
    )

    # Outstanding wire → GL only
    for desc, amount, date in USD_OUTSTANDING_WIRE:
        ref = f"BU-{bu_counter}"
        bu_counter += 1
        spot_snb = _approximate_daily_rate(
            date.day, FX_CHF_USD_CLOSING, start_snb=usd_start_snb,
        )
        chf_amount = _to_chf(amount, spot_snb)
        gl_entries_pre_reval.append(GLCashEntryCH(
            date=date, reference=ref, description=desc,
            debit=Decimal("0"), credit=chf_amount,
            running_balance=Decimal("0"),
            currency="CHF", original_currency="USD",
            original_amount=amount, fx_rate=spot_snb,
            category="outstanding_wire", matched=False,
        ))

    # Pre-revaluation GL net
    gl_net_pre_reval = sum(e.debit - e.credit for e in gl_entries_pre_reval)

    # USD GL starting: approximate at start-of-month rate
    usd_gl_starting_approx = _to_chf(USD_BANK_STARTING, usd_start_snb)

    # Revaluation to make GL ending = USD_GL_ENDING_CHF (correct rate, no error)
    reval_amount = USD_GL_ENDING_CHF - usd_gl_starting_approx - gl_net_pre_reval

    reval_ref = f"BU-{bu_counter}"
    bu_counter += 1
    reval_desc = "FX-Neubewertung USD per 31.12.2025 (Kurs CHF/USD 0.8845)"
    if reval_amount >= 0:
        gl_entries_pre_reval.append(GLCashEntryCH(
            date=datetime.date(2025, 12, 31),
            reference=reval_ref,
            description=reval_desc,
            debit=reval_amount,
            credit=Decimal("0"),
            running_balance=Decimal("0"),
            currency="CHF", original_currency="USD",
            original_amount=Decimal("0"),
            fx_rate=FX_CHF_USD_CLOSING,
            category="fx_revaluation", matched=False,
        ))
    else:
        gl_entries_pre_reval.append(GLCashEntryCH(
            date=datetime.date(2025, 12, 31),
            reference=reval_ref,
            description=reval_desc,
            debit=Decimal("0"),
            credit=abs(reval_amount),
            running_balance=Decimal("0"),
            currency="CHF", original_currency="USD",
            original_amount=Decimal("0"),
            fx_rate=FX_CHF_USD_CLOSING,
            category="fx_revaluation", matched=False,
        ))

    gl_entries_pre_reval.sort(key=lambda e: (e.date, e.reference))
    model.usd_gl_entries = gl_entries_pre_reval
    model.usd_gl_starting_chf = usd_gl_starting_approx

    gl_running = model.usd_gl_starting_chf
    for entry in model.usd_gl_entries:
        gl_running += entry.debit - entry.credit
        entry.running_balance = gl_running

    assert gl_running == USD_GL_ENDING_CHF, (
        f"USD GL ending: {gl_running} != {USD_GL_ENDING_CHF}"
    )


# ── Description generators ───────────────────────────────────────────────

def _make_chf_bank_desc(
    category: str,
    is_debit: bool,
    rng: random.Random,
    vendor_idx: int,
    customer_idx: int,
) -> str:
    """Generate Swiss German bank-style description for CHF transactions."""
    if category == "salary":
        return f"Lohnzahlung CPI AG Ref {rng.randint(1000, 9999)}"
    if category == "lsv_collection":
        name = _CH_CUSTOMER_NAMES[customer_idx % len(_CH_CUSTOMER_NAMES)]
        return f"LSV+ Einzug Ref LSV-2025-{rng.randint(100, 999)}"
    if category == "sic_inflow":
        name = _CH_CUSTOMER_NAMES[customer_idx % len(_CH_CUSTOMER_NAMES)]
        abbr = name[:15].upper()
        return f"SIC Gutschrift {abbr}"
    if category == "deposit":
        return f"Bareinzahlung Ref {rng.randint(1000, 9999)}"
    if category == "dta_inflow":
        return f"DTA Gutschrift Ref DTA-2025-{rng.randint(1000, 9999)}"
    if category == "qr_payment":
        return f"QR-Rechnung Einzahlung Ref QR-{rng.randint(100000, 999999)}"
    if category == "dta_payment":
        return f"DTA-Zahlung Ref 2025-12-{rng.randint(1000, 9999)}"
    if category == "check":
        return f"Scheck Nr. {rng.randint(8800, 8899)}"
    if category == "lsv_debit":
        return f"LSV+ Lastschrift Ref LSV-2025-{rng.randint(100, 999)}"
    if category == "tax_payment":
        cantons = ["ZH", "BE", "BS", "AG"]
        return f"Steuerzahlung Kanton {rng.choice(cantons)} {rng.randint(1000, 9999)}"
    if category == "rent":
        return f"Miete Büroräume Zürich Ref {rng.randint(100, 999)}"
    if category == "insurance":
        insurers = ["Zurich Versicherung", "Swiss Re", "Helvetia", "Mobiliar"]
        return f"Versicherungsprämie {rng.choice(insurers)}"
    return f"{'Belastung' if is_debit else 'Gutschrift'} Ref {rng.randint(1000, 9999)}"


def _make_chf_gl_desc(
    category: str,
    is_debit: bool,
    rng: random.Random,
    vendor_idx: int,
    customer_idx: int,
) -> str:
    """Generate company GL description for CHF transactions."""
    if category == "salary":
        return "Lohnzahlung — Monatslohn Dezember 2025"
    if category == "lsv_collection":
        name = _CH_CUSTOMER_NAMES[customer_idx % len(_CH_CUSTOMER_NAMES)]
        return f"LSV+ Einzug — {name}"
    if category == "sic_inflow":
        name = _CH_CUSTOMER_NAMES[customer_idx % len(_CH_CUSTOMER_NAMES)]
        return f"SIC Zahlungseingang — {name}"
    if category == "deposit":
        return "Bareinzahlung"
    if category == "dta_inflow":
        name = _CH_CUSTOMER_NAMES[customer_idx % len(_CH_CUSTOMER_NAMES)]
        return f"DTA Zahlungseingang — {name}"
    if category == "qr_payment":
        name = _CH_CUSTOMER_NAMES[customer_idx % len(_CH_CUSTOMER_NAMES)]
        return f"QR-Zahlung erhalten — {name}"
    if category == "dta_payment":
        name = _CH_VENDOR_NAMES[vendor_idx % len(_CH_VENDOR_NAMES)]
        return f"DTA-Zahlung — {name}"
    if category == "check":
        name = _CH_VENDOR_NAMES[vendor_idx % len(_CH_VENDOR_NAMES)]
        return f"Scheck — {name}"
    if category == "lsv_debit":
        name = _CH_VENDOR_NAMES[vendor_idx % len(_CH_VENDOR_NAMES)]
        return f"LSV+ Lastschrift — {name}"
    if category == "tax_payment":
        return "Steuerzahlung — Kanton/Bund"
    if category == "rent":
        return "Miete — Büroräume Zürich"
    if category == "insurance":
        return "Versicherungsprämie"
    return "Barbewegung" if is_debit else "Bareingang"


def _make_eur_bank_desc(
    category: str,
    is_debit: bool,
    rng: random.Random,
    vendor_idx: int,
    customer_idx: int,
) -> str:
    """Generate bank description for EUR transactions."""
    if category == "sepa_credit":
        name = _EU_CUSTOMER_NAMES[customer_idx % len(_EU_CUSTOMER_NAMES)]
        abbr = name[:15].upper()
        return f"SEPA Gutschrift {abbr} Ref SEPA-{rng.randint(100000, 999999)}"
    if category == "swift_in":
        name = _EU_CUSTOMER_NAMES[customer_idx % len(_EU_CUSTOMER_NAMES)]
        abbr = name[:12].upper()
        return f"SWIFT Gutschrift MT103 {abbr} REF{rng.randint(10000, 99999)}"
    if category == "sepa_debit":
        name = _EU_VENDOR_NAMES[vendor_idx % len(_EU_VENDOR_NAMES)]
        abbr = name[:15].upper()
        return f"SEPA Überweisung {abbr} Ref SEPA-{rng.randint(100000, 999999)}"
    if category == "swift_out":
        name = _EU_VENDOR_NAMES[vendor_idx % len(_EU_VENDOR_NAMES)]
        abbr = name[:12].upper()
        return f"SWIFT Überweisung MT103 {abbr} REF{rng.randint(10000, 99999)}"
    return f"{'Belastung' if is_debit else 'Gutschrift'} EUR {rng.randint(1000, 9999)}"


def _make_eur_gl_desc(
    category: str,
    is_debit: bool,
    rng: random.Random,
    vendor_idx: int,
    customer_idx: int,
) -> str:
    """Generate GL description for EUR transactions (in CHF)."""
    if category == "sepa_credit":
        name = _EU_CUSTOMER_NAMES[customer_idx % len(_EU_CUSTOMER_NAMES)]
        return f"SEPA Zahlungseingang — {name}"
    if category == "swift_in":
        name = _EU_CUSTOMER_NAMES[customer_idx % len(_EU_CUSTOMER_NAMES)]
        return f"SWIFT Eingang — {name}"
    if category == "sepa_debit":
        name = _EU_VENDOR_NAMES[vendor_idx % len(_EU_VENDOR_NAMES)]
        return f"SEPA-Zahlung — {name}"
    if category == "swift_out":
        name = _EU_VENDOR_NAMES[vendor_idx % len(_EU_VENDOR_NAMES)]
        return f"SWIFT Zahlung — {name}"
    return "EUR Buchung"


def _make_usd_bank_desc(
    category: str,
    is_debit: bool,
    rng: random.Random,
    vendor_idx: int,
    customer_idx: int,
) -> str:
    """Generate bank description for USD transactions."""
    if category == "swift_in":
        name = _US_CUSTOMER_NAMES[customer_idx % len(_US_CUSTOMER_NAMES)]
        abbr = name[:12].upper()
        return f"SWIFT Credit MT103 {abbr} REF{rng.randint(10000, 99999)}"
    if category == "wire_in":
        name = _US_CUSTOMER_NAMES[customer_idx % len(_US_CUSTOMER_NAMES)]
        abbr = name[:12].upper()
        return f"Wire Transfer Credit {abbr} REF{rng.randint(10000, 99999)}"
    if category == "swift_out":
        name = _US_VENDOR_NAMES[vendor_idx % len(_US_VENDOR_NAMES)]
        abbr = name[:12].upper()
        return f"SWIFT Debit MT103 {abbr} REF{rng.randint(10000, 99999)}"
    if category == "wire_out":
        name = _US_VENDOR_NAMES[vendor_idx % len(_US_VENDOR_NAMES)]
        abbr = name[:12].upper()
        return f"Wire Transfer Debit {abbr} REF{rng.randint(10000, 99999)}"
    return f"{'Debit' if is_debit else 'Credit'} USD {rng.randint(1000, 9999)}"


def _make_usd_gl_desc(
    category: str,
    is_debit: bool,
    rng: random.Random,
    vendor_idx: int,
    customer_idx: int,
) -> str:
    """Generate GL description for USD transactions (in CHF)."""
    if category == "swift_in":
        name = _US_CUSTOMER_NAMES[customer_idx % len(_US_CUSTOMER_NAMES)]
        return f"SWIFT Receipt — {name}"
    if category == "wire_in":
        name = _US_CUSTOMER_NAMES[customer_idx % len(_US_CUSTOMER_NAMES)]
        return f"Wire Received — {name}"
    if category == "swift_out":
        name = _US_VENDOR_NAMES[vendor_idx % len(_US_VENDOR_NAMES)]
        return f"SWIFT Payment — {name}"
    if category == "wire_out":
        name = _US_VENDOR_NAMES[vendor_idx % len(_US_VENDOR_NAMES)]
        return f"Wire Payment — {name}"
    return "USD Buchung"


# ── Main generator ───────────────────────────────────────────────────────

def generate_bank_ch_model(rng: random.Random) -> BankModelCH:
    """Generate the complete Swiss multi-currency bank model for TC-23.

    Produces:
    - CHF: 185 bank transactions, matching GL entries, 2 outstanding checks,
      1 deposit in transit, bank interest and charges
    - EUR: 62 bank transactions, matching GL in CHF, 1 outstanding SEPA,
      month-end FX revaluation with stale rate (ERR-CH-002)
    - USD: 28 bank transactions, matching GL in CHF, 1 outstanding wire,
      month-end FX revaluation with correct rate
    - Reconciliation summaries for all 3 accounts
    - Consolidated total: CHF 4,618,565

    Returns a fully populated BankModelCH with gold standard reconciliation data.
    """
    model = BankModelCH()

    _generate_chf_account(rng, model)
    _generate_eur_account(rng, model)
    _generate_usd_account(rng, model)

    # Build reconciliation summaries
    model.chf_recon = AccountReconciliation(
        currency="CHF",
        iban=CHF_IBAN,
        bank_ending=CHF_BANK_ENDING,
        outstanding_items=[(d, a, dt) for d, a, dt in CHF_OUTSTANDING_CHECKS],
        outstanding_total=CHF_OUTSTANDING_CHECKS_TOTAL,
        deposits_in_transit=[(d, a, dt) for d, a, dt in CHF_DEPOSITS_IN_TRANSIT],
        deposits_in_transit_total=CHF_DEPOSITS_IN_TRANSIT_TOTAL,
        adjusted_bank=CHF_ADJUSTED_BANK,
        adjusted_bank_chf=CHF_ADJUSTED_BANK,
        gl_ending_chf=CHF_GL_ENDING,
        bank_charges=CHF_BANK_CHARGES,
        bank_interest=CHF_BANK_INTEREST,
    )

    model.eur_recon = AccountReconciliation(
        currency="EUR",
        iban=EUR_IBAN,
        bank_ending=EUR_BANK_ENDING,
        outstanding_items=[(d, a, dt) for d, a, dt in EUR_OUTSTANDING_SEPA],
        outstanding_total=EUR_OUTSTANDING_SEPA_TOTAL,
        deposits_in_transit=[],
        deposits_in_transit_total=Decimal("0"),
        adjusted_bank=EUR_ADJUSTED_BANK,
        adjusted_bank_chf=EUR_ADJUSTED_BANK_CHF,
        gl_ending_chf=EUR_GL_ENDING_CHF,
        fx_rate=FX_CHF_EUR_CLOSING,
        fx_error=EUR_FX_ERROR,
    )

    model.usd_recon = AccountReconciliation(
        currency="USD",
        iban=USD_IBAN,
        bank_ending=USD_BANK_ENDING,
        outstanding_items=[(d, a, dt) for d, a, dt in USD_OUTSTANDING_WIRE],
        outstanding_total=USD_OUTSTANDING_WIRE_TOTAL,
        deposits_in_transit=[],
        deposits_in_transit_total=Decimal("0"),
        adjusted_bank=USD_ADJUSTED_BANK,
        adjusted_bank_chf=USD_ADJUSTED_BANK_CHF,
        gl_ending_chf=USD_GL_ENDING_CHF,
        fx_rate=FX_CHF_USD_CLOSING,
    )

    return model


# ── Validation ───────────────────────────────────────────────────────────

def validate_reconciliation_ch(model: BankModelCH) -> list[str]:
    """Validate the Swiss model's reconciliation against gold standard values.

    Returns a list of errors (empty = valid).
    """
    errors: list[str] = []

    # CHF account
    chf_bank_end = model.chf_bank_transactions[-1].running_balance
    if chf_bank_end != CHF_BANK_ENDING:
        errors.append(f"CHF bank ending {chf_bank_end} != {CHF_BANK_ENDING}")

    chf_adj = chf_bank_end + CHF_DEPOSITS_IN_TRANSIT_TOTAL - CHF_OUTSTANDING_CHECKS_TOTAL
    if chf_adj != CHF_ADJUSTED_BANK:
        errors.append(f"CHF adjusted bank {chf_adj} != {CHF_ADJUSTED_BANK}")

    chf_gl_end = model.chf_gl_entries[-1].running_balance
    if chf_gl_end != CHF_GL_ENDING:
        errors.append(f"CHF GL ending {chf_gl_end} != {CHF_GL_ENDING}")

    chf_adj_book = chf_gl_end + CHF_BANK_INTEREST - CHF_BANK_CHARGES
    if chf_adj_book != CHF_ADJUSTED_BOOK:
        errors.append(f"CHF adjusted book {chf_adj_book} != {CHF_ADJUSTED_BOOK}")

    if len(model.chf_bank_transactions) != CHF_TX_COUNT:
        errors.append(f"CHF tx count {len(model.chf_bank_transactions)} != {CHF_TX_COUNT}")

    # EUR account
    eur_bank_end = model.eur_bank_transactions[-1].running_balance
    if eur_bank_end != EUR_BANK_ENDING:
        errors.append(f"EUR bank ending {eur_bank_end} != {EUR_BANK_ENDING}")

    eur_gl_end = model.eur_gl_entries[-1].running_balance
    if eur_gl_end != EUR_GL_ENDING_CHF:
        errors.append(f"EUR GL ending CHF {eur_gl_end} != {EUR_GL_ENDING_CHF}")

    if len(model.eur_bank_transactions) != EUR_TX_COUNT:
        errors.append(f"EUR tx count {len(model.eur_bank_transactions)} != {EUR_TX_COUNT}")

    # USD account
    usd_bank_end = model.usd_bank_transactions[-1].running_balance
    if usd_bank_end != USD_BANK_ENDING:
        errors.append(f"USD bank ending {usd_bank_end} != {USD_BANK_ENDING}")

    usd_gl_end = model.usd_gl_entries[-1].running_balance
    if usd_gl_end != USD_GL_ENDING_CHF:
        errors.append(f"USD GL ending CHF {usd_gl_end} != {USD_GL_ENDING_CHF}")

    if len(model.usd_bank_transactions) != USD_TX_COUNT:
        errors.append(f"USD tx count {len(model.usd_bank_transactions)} != {USD_TX_COUNT}")

    # Consolidated
    total = CHF_ADJUSTED_BANK + EUR_ADJUSTED_BANK_CHF + USD_ADJUSTED_BANK_CHF
    if total != CONSOLIDATED_CHF:
        errors.append(f"Consolidated {total} != {CONSOLIDATED_CHF}")

    return errors
