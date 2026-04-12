# European Standards Reference

Status: Active
Last updated: 2026-04-12
Tracker bead: `synth-data-eu.3`

## Purpose

This document lists the authoritative standards and official guidance that anchor the European variants in the `cascade_europe_ifrs` scenario pack. Each entry includes a source URL, scope note, affected test cases, and a non-goal caveat.

**Non-goal (applies to all entries below):** The synth-data suite models benchmark scenarios for evaluating AI agent performance. It does not provide legal, tax, or accounting advice, and generated documents are not substitutes for professional judgment or compliance with any jurisdiction's requirements.

---

## Accounting Standards

### IFRS 16 — Leases

| Field | Value |
|---|---|
| **Source** | IFRS Foundation |
| **URL** | https://www.ifrs.org/issued-standards/list-of-standards/ifrs-16-leases/ |
| **Scope note** | Governs lessee accounting for right-of-use assets and lease liabilities. Replaces IAS 17. The European variant uses IFRS 16 classification (no operating/finance split for lessees) instead of ASC 842's dual-model approach. Key differences exercised: single lessee model, discount rate selection (incremental borrowing rate emphasis), and lease modification treatment. |
| **Affected TCs** | TC-04-EU (IFRS 16 lease extraction), TC-18-EU (IFRS/ISA rollforward — lease schedules) |
| **Non-goal** | The suite does not model every IFRS 16 practical expedient or transition election. Scenarios use a steady-state IFRS 16 portfolio, not a first-time adoption scenario. |

### IAS 12 — Income Taxes

| Field | Value |
|---|---|
| **Source** | IFRS Foundation |
| **URL** | https://www.ifrs.org/issued-standards/list-of-standards/ias-12-income-taxes/ |
| **Scope note** | Governs current and deferred tax accounting under IFRS. The European variant replaces ASC 740 concepts (valuation allowance, indefinite reversal exception) with IAS 12 equivalents (recoverability assessment of deferred tax assets, no indefinite reversal exception for subsidiaries). Exercises multi-jurisdiction tax provision across NL, DE, FR, and UK entities. |
| **Affected TCs** | TC-06-EU (IAS 12 income tax provision), TC-18-EU (IFRS/ISA rollforward — tax provision schedules) |
| **Non-goal** | The suite does not model IAS 12 amendments related to Pillar Two (the Cascade Europe group is below the €750M GloBE threshold). A judgment-trap question may test whether an agent incorrectly applies Pillar Two to a sub-threshold group. |

### IAS 21 — The Effects of Changes in Foreign Exchange Rates

| Field | Value |
|---|---|
| **Source** | IFRS Foundation |
| **URL** | https://www.ifrs.org/issued-standards/list-of-standards/ias-21-the-effects-of-changes-in-foreign-exchange-rates/ |
| **Scope note** | Governs translation of foreign currency transactions and foreign operations. Used for GBP→EUR translation of the UK subsidiary (Cascade Distribution Services Ltd). Translation method: closing rate for balance sheet, average rate for P&L. |
| **Affected TCs** | TC-06-EU (consolidated tax provision includes FX effects), TC-07-EU (investment reporting with multi-currency entities), TC-18-EU (rollforward with FX translation) |
| **Non-goal** | The suite uses fixed deterministic FX rates (budget rate GBP/EUR 1.17). It does not model hedge accounting or IAS 21 hyperinflationary economy provisions. |

---

## Tax and Transfer Pricing

### OECD Transfer Pricing Guidelines for Multinational Enterprises and Tax Administrations

| Field | Value |
|---|---|
| **Source** | OECD |
| **URL** | https://www.oecd.org/en/topics/transfer-pricing.html |
| **Scope note** | Provides the arm's-length principle framework for intercompany pricing. The European variant uses OECD-aligned master file/local file documentation (per BEPS Action 13) and exercises TNMM for routine distributors and cost-plus for contract manufacturing. High-risk transactions include the R&D royalty (CM→CP) and intercompany loan (CE→CM). |
| **Affected TCs** | TC-09-EU (OECD transfer pricing documentation), TC-12-EU (data room includes TP documentation) |
| **Non-goal** | The suite does not model advance pricing agreements (APAs), mutual agreement procedures (MAP), or dispute resolution. Country-by-country reporting is included as a design element despite the group being below the €750M threshold. |

### OECD Pillar Two — Global Anti-Base Erosion (GloBE) Model Rules

| Field | Value |
|---|---|
| **Source** | OECD |
| **URL** | https://www.oecd.org/en/topics/pillar-two.html |
| **Scope note** | Establishes a 15% global minimum tax for multinational groups with consolidated revenue ≥ €750M. The Cascade Europe group (~€120M revenue) is below this threshold and is therefore out of scope. Referenced only as a judgment trap in TC-06-EU where an agent might incorrectly apply GloBE rules. |
| **Affected TCs** | TC-06-EU (judgment-trap question only) |
| **Non-goal** | No Pillar Two calculations, top-up tax computations, or qualified domestic minimum top-up tax (QDMTT) modeling. The reference exists solely to test whether an agent correctly scopes the rule. |

### France — Crédit d'Impôt Recherche (CIR)

| Field | Value |
|---|---|
| **Source** | French Ministry of Research / Ministère de l'Enseignement supérieur et de la Recherche |
| **URL** | https://www.enseignementsup-recherche.gouv.fr/fr/le-credit-d-impot-recherche-cir-46351 |
| **Scope note** | France's primary R&D tax incentive: 30% credit on the first €100M of eligible R&D expenditure. Eligible costs include researcher salaries, consumables, and subcontracted R&D (capped). Cascade Matériaux Avancés SAS (Lyon) is the primary CIR beneficiary with 12% of revenue in R&D. |
| **Affected TCs** | TC-08-EU (R&D incentive variant — primary regime) |
| **Non-goal** | The suite does not model CIR audit defense, the innovation tax credit (CII), or the young innovative enterprise (JEI) status. Scenarios use the standard CIR rate without reduced rates for spend above €100M. |

### Germany — Forschungszulagengesetz (Research Allowance Act)

| Field | Value |
|---|---|
| **Source** | German Federal Ministry of Finance (BMF) |
| **URL** | https://www.bundesfinanzministerium.de/Web/EN/Issues/Taxation/research-allowance.html |
| **Scope note** | Germany's research allowance: 25% of eligible personnel costs up to a €2M/year assessment basis (maximum benefit €500k). Simpler than the French CIR but lower benefit ceiling. Cascade Präzisionsteile GmbH (Munich) qualifies with 4% of revenue in R&D. |
| **Affected TCs** | TC-08-EU (R&D incentive variant — secondary regime) |
| **Non-goal** | The suite does not model the expanded 2024+ assessment basis increases or contract research eligibility beyond the basic rule. |

---

## VAT and Indirect Tax

### European Commission — VAT Directive (Council Directive 2006/112/EC)

| Field | Value |
|---|---|
| **Source** | European Commission |
| **URL** | https://taxation-customs.ec.europa.eu/taxation/vat_en |
| **Scope note** | Governs VAT across EU member states. Key provisions exercised: intra-Community supply zero-rating (Art. 138), reverse charge on B2B services (Art. 196), and the distinction between intra-EU and third-country (UK post-Brexit) transactions. VAT return frequencies: monthly (DE, FR), quarterly (NL, UK). |
| **Affected TCs** | TC-10-EU (VAT and cross-border tax position), TC-12-EU (data room includes VAT registrations), TC-16-EU (engagement letter covers VAT compliance scope) |
| **Non-goal** | The suite does not model VAT grouping, partial exemption calculations, or member-state-specific derogations beyond DE, FR, and NL standard rates. UK VAT is modeled as a third-country import VAT scenario, not under the EU VAT Directive. |

---

## Auditing Standards

### IAASB — International Standards on Auditing (ISA)

| Field | Value |
|---|---|
| **Source** | International Auditing and Assurance Standards Board (IAASB) |
| **URL** | https://www.iaasb.org/standards |
| **Scope note** | ISAs provide the audit framework used internationally (outside the US PCAOB regime). Key standards referenced: ISA 315 (Revised 2019) for risk assessment, ISA 330 for audit procedures responding to assessed risks, and ISA 700/701 for the auditor's report format. European variant engagement letters and workpapers use ISA terminology instead of PCAOB/AICPA framing. |
| **Affected TCs** | TC-16-EU (European engagement letter — ISA framing), TC-18-EU (IFRS/ISA rollforward — ISA workpaper terminology) |
| **Non-goal** | The suite does not model full ISA audit methodology, quality management (ISQM 1/2), or key audit matters (KAM) reporting. ISA references are limited to terminology and document structure differences from US GAAS/PCAOB standards. |

---

## Corporate and Company Law

### EU Company Legal Forms

| Field | Value |
|---|---|
| **Source** | National company registries (KvK Netherlands, Handelsregister Germany, RCS France, Companies House UK) |
| **URL** | N/A (jurisdiction-specific registries) |
| **Scope note** | The Cascade Europe group uses jurisdiction-specific legal forms: B.V. (Netherlands), GmbH (Germany), SAS (France), Ltd (UK). These forms determine statutory filing requirements, minimum capital rules, and director liability structures referenced in engagement letters and data room documents. |
| **Affected TCs** | TC-12-EU (data room — corporate documents), TC-16-EU (engagement letter — entity identification) |
| **Non-goal** | The suite does not model company formation, statutory filing deadlines, or corporate governance requirements beyond what appears in generated documents. Legal forms are used for realistic document headers, not compliance testing. |

---

## Cross-Reference: Standards by Test Case

| Test Case | Standards Referenced |
|---|---|
| TC-04-EU | IFRS 16 |
| TC-06-EU | IAS 12, OECD Pillar Two (judgment trap only), IAS 21 |
| TC-07-EU | IAS 21 |
| TC-08-EU | France CIR, Germany Forschungszulage |
| TC-09-EU | OECD Transfer Pricing Guidelines |
| TC-10-EU | EU VAT Directive |
| TC-12-EU | OECD TP Guidelines, EU VAT Directive, EU Company Legal Forms |
| TC-16-EU | IAASB ISA, EU VAT Directive, EU Company Legal Forms |
| TC-18-EU | IFRS 16, IAS 12, IAS 21, IAASB ISA |
