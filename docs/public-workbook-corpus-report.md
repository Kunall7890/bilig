---
title: Public workbook corpus report
published: true
description: Generated Bilig public workbook corpus scorecard with workbook counts, formula cells, license evidence, and explicit limitations.
tags: xlsx, public-data, formulas, compatibility, workpaper
canonical_url: https://proompteng.github.io/bilig/public-workbook-corpus-report.html
cover_image: https://raw.githubusercontent.com/proompteng/bilig/main/docs/assets/github-social-preview.png
image: /assets/github-social-preview.png
---

# Public Workbook Corpus Report

Status: generated public evidence from `packages/benchmarks/baselines/public-workbook-corpus-scorecard.json`.

This report is intentionally scoped. It publishes the checked-in public workbook
corpus baseline that ships with the repository: `22` cached
public workbooks selected from `602` source candidates. It is not the
larger active financial or 10,000-workbook corpus goal.

## Reproduce The Report

```sh
pnpm public-workbook-corpus:status
pnpm public:evidence:check
```

For a no-key package smoke outside a repo clone, run:

```sh
npm exec --yes --package @bilig/xlsx-formula-recalc@latest -- bilig-evaluate --door xlsx-cache --json
```

The package smoke proves the stale-XLSX recalculation door. The corpus commands
above prove this checked repository scorecard and generated report.

## Scorecard Summary

| Metric | Value |
| --- | ---: |
| Source candidates | `602` |
| Cached public workbooks | `22` |
| Imported workbooks | `22` |
| Passed workbooks | `19` |
| Resource-limited unsupported workbooks | `3` |
| Failed workbooks | `0` |
| Error workbooks | `0` |
| Formula-oracle matches | `6/6` |
| Structural smoke runs | `6` |
| All cached workbooks passed gate | `true` |
| Remaining to checked target | `0` |
| Generated at | `2026-05-07T10:07:15.182Z` |

## Workbook Shape

| Feature | Count |
| --- | ---: |
| Workbook bytes | `7,080,426` |
| Sheets | `221` |
| Cells | `812,868` |
| Formula cells | `46,211` |
| Defined names | `213` |
| Tables | `2` |
| Charts | `3` |
| Pivots | `2` |
| Merges | `1,281` |
| Conditional formats | `29` |
| Warnings | `3` |

## Largest Cases

| Workbook | Status | Cells | Formula cells | License |
| --- | --- | ---: | ---: | --- |
| noibyfarmsize_fr.xlsx | `unsupported` | `342,986` | `46,205` | `ON-OGLO` |
| sfgsme-efcpme-tables-2023-eng.xlsx | `unsupported` | `246,084` | `0` | `CA-OGL-LGO` |
| 0nContract_Register_Jan_2026.xlsx | `unsupported` | `211,653` | `0` | `UK-OGL` |
| monies_ot_part_iv_and_xi.xlsx | `passed` | `2,325` | `0` | `CC-BY-2.5` |
| monies_rt_part_iv_and_xi.xlsx | `passed` | `1,593` | `0` | `CC-BY-2.5` |
| monies_rt_part_x.xlsx | `passed` | `1,494` | `0` | `CC-BY-2.5` |

## Licenses

| SPDX | Workbooks | Evidence |
| --- | ---: | --- |
| `Apache-2.0` | `2` | [Apache License 2.0](https://github.com/apache/poi/blob/trunk/LICENSE) |
| `CA-OGL-LGO` | `1` | [Open Government Licence - Canada](https://open.canada.ca/en/open-government-licence-canada) |
| `CC-BY-2.5` | `17` | [Creative Commons Attribution 2.5 Australia](http://creativecommons.org/licenses/by/2.5/au/) |
| `ON-OGLO` | `1` | [Open Government Licence - Ontario](https://www.ontario.ca/page/open-government-licence-ontario) |
| `UK-OGL` | `1` | [UK Open Government Licence (OGL)](https://www.nationalarchives.gov.uk/doc/open-government-licence/version/3/) |

## Resource-Limited Cases

The scorecard classifies `3` workbooks as unsupported
because their footprint exceeds the configured
round-trip or structural-smoke resource budget. Those rows still have source,
license, hash, import, and classification evidence.

| Workbook | Cells | Classification |
| --- | ---: | --- |
| 0nContract_Register_Jan_2026.xlsx | `211,653` | `xlsx.publicCorpus.resourceLimit:preflightRoundTripBudget>100000cells`<br>`xlsx.publicCorpus.resourceLimit:preflightStructuralSmokeBudget>100000cells` |
| noibyfarmsize_fr.xlsx | `342,986` | `xlsx.publicCorpus.resourceLimit:preflightFormulaOracleBudget>2000formulas`<br>`xlsx.publicCorpus.resourceLimit:preflightRoundTripBudget>100000cells`<br>`xlsx.publicCorpus.resourceLimit:preflightStructuralSmokeBudget>100000cells` |
| sfgsme-efcpme-tables-2023-eng.xlsx | `246,084` | `xlsx.publicCorpus.resourceLimit:preflightRoundTripBudget>100000cells`<br>`xlsx.publicCorpus.resourceLimit:preflightStructuralSmokeBudget>100000cells` |

## What This Proves

- The checked public baseline imports `22` cached public workbook
  files without failures or errors.
- It covers `812,868` cells, `46,211`
  formula cells, `221` sheets, and `5` source
  license families.
- Formula-oracle rows in this baseline match cached workbook evidence
  `6/6`.
- Resource-budget skips are visible instead of being counted as silent passes.

## What This Does Not Prove

- This is the checked-in 22-workbook public scorecard, not the broader 10,000-workbook or 5,000-financial-workbook objective.
- Unsupported rows are resource-budget classifications with evidence, not hidden failures.
- Formula-oracle matches only cover workbooks with meaningful cached formula comparisons in this scorecard.
- Run the commands above before using this report as current release evidence.

Read [where Bilig is not Excel-compatible yet](where-bilig-is-not-excel-compatible-yet.md)
before treating any corpus scorecard as a blanket Excel-parity claim. Use the
[XLSX corpus verifier walkthrough](xlsx-corpus-verifier-walkthrough.md) when
you need to run the same boundary against private workbooks.
