# European Professional Services Variants

Status: Draft
Last updated: 2026-04-12
Primary tracker epic: `synth-data-eu`

## Summary

Add European/IFRS/OECD variants for test cases that currently depend on US-specific company forms, accounting standards, tax forms, or state tax regimes.

Initial affected cases:

- TC-04: ASC 842 lease extraction -> IFRS 16 lease extraction
- TC-06: ASC 740 income tax provision -> IAS 12 income taxes, with optional Pillar Two scope decision
- TC-07: K-1/Form 1120 extraction -> selected European partnership/investment reporting analogue
- TC-08: IRC Section 41 R&D credit -> selected European country R&D incentive regime
- TC-09: transfer pricing -> OECD-style transfer pricing documentation/local-file framing
- TC-10: US multi-state apportionment -> EU VAT and cross-border tax position analysis
- TC-12: US-heavy data room -> European diligence data room
- TC-16: US federal/state engagement letter -> European IFRS/VAT/CIT engagement letter
- TC-18: US-framed rollforward -> IFRS/ISA rollforward variant

## Source Standards And Guidance

See [`specs/european-standards-reference.md`](european-standards-reference.md) for the full reference with URLs, scope notes, affected TCs, and non-goal caveats.

Summary of referenced standards:

- IFRS Foundation: IFRS 16 Leases, IAS 12 Income Taxes, IAS 21 Foreign Exchange
- OECD: Transfer Pricing Guidelines, Pillar Two model rules (judgment trap only)
- France: Crédit d'Impôt Recherche (CIR)
- Germany: Forschungszulagengesetz (Research Allowance)
- European Commission: VAT Directive (Council Directive 2006/112/EC)
- IAASB: International Standards on Auditing (ISA 315, 330, 700/701)
- EU company legal forms: B.V., GmbH, SAS, Ltd

These sources are anchors for synthetic benchmark design, not legal or tax advice.

## Scope Boundary

Europe is not one tax jurisdiction. Every European variant must explicitly state its jurisdiction profile before implementation. The default recommendation is an IFRS-reporting group with selected European subsidiaries and a clearly documented decision on whether UK is included as a non-EU European comparator.

Do not replace existing US cases in place. European variants should be optional scenario-pack cases or explicit variants, preserving current benchmark continuity.

## Beads

The detailed implementation backlog is filed under `synth-data-eu`:

- `synth-data-eu.1`: Design European scenario-pack and jurisdiction strategy
- `synth-data-eu.2`: Design European company and jurisdiction profile
- `synth-data-eu.3`: Add European standards reference spec
- `synth-data-eu.4`: Design TC-04-EU IFRS 16 lease extraction variant
- `synth-data-eu.5`: Implement TC-04-EU IFRS 16 model, formatter, gold, and rubric
- `synth-data-eu.6`: Design TC-06-EU IAS 12 income tax provision variant
- `synth-data-eu.7`: Implement TC-06-EU IAS 12 model, formatter, gold, and rubric
- `synth-data-eu.8`: Design TC-07-EU partnership investment reporting variant
- `synth-data-eu.9`: Implement TC-07-EU partnership investment reporting variant
- `synth-data-eu.10`: Design TC-08-EU R&D incentive variant
- `synth-data-eu.11`: Implement TC-08-EU R&D incentive model, formatter, gold, and rubric
- `synth-data-eu.12`: Design TC-09-EU OECD transfer pricing documentation variant
- `synth-data-eu.13`: Implement TC-09-EU OECD transfer pricing documentation variant
- `synth-data-eu.14`: Design TC-10-EU VAT and cross-border tax position variant
- `synth-data-eu.15`: Implement TC-10-EU VAT and cross-border tax position variant
- `synth-data-eu.16`: Design TC-12-EU European diligence data room variant
- `synth-data-eu.17`: Implement TC-12-EU European diligence data room variant
- `synth-data-eu.18`: Design TC-16-EU European engagement letter variant
- `synth-data-eu.19`: Implement TC-16-EU European engagement letter variant
- `synth-data-eu.20`: Design TC-18-EU IFRS/ISA rollforward variant
- `synth-data-eu.21`: Implement TC-18-EU IFRS/ISA rollforward variant
- `synth-data-eu.22`: Add European variants pack end-to-end self-test and determinism gate
- `synth-data-eu.23`: Document European variants and source-standard assumptions
