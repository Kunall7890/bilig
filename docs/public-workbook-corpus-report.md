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
corpus baseline that ships with the repository: `52` cached
public workbooks selected from `102` source candidates. It is not the
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
| Source candidates | `102` |
| Cached public workbooks | `52` |
| Imported workbooks | `52` |
| Passed workbooks | `38` |
| Resource-limited unsupported workbooks | `14` |
| Failed workbooks | `0` |
| Error workbooks | `0` |
| Formula-oracle matches | `529/529` |
| Structural smoke runs | `50` |
| All cached workbooks passed gate | `true` |
| Remaining to checked target | `0` |
| Generated at | `2026-06-05T17:48:03.593Z` |

## Workbook Shape

| Feature | Count |
| --- | ---: |
| Workbook bytes | `4,044,346` |
| Sheets | `217` |
| Cells | `280,480` |
| Formula cells | `529` |
| Defined names | `392` |
| Tables | `1` |
| Charts | `3` |
| Pivots | `1` |
| Merges | `1,844` |
| Conditional formats | `27` |
| Warnings | `15` |

## Largest Cases

| Workbook | Status | Cells | Formula cells | License |
| --- | --- | ---: | ---: | --- |
| sfgsme-efcpme-tables-2023-eng.xlsx | `unsupported` | `246,084` | `0` | `CA-OGL-LGO` |
| Online-data-1.1-Taxation-Revenue-GG-Budget-Update-2023-24.xlsx | `unsupported` | `13,508` | `0` | `CC-BY` |
| ganumberoffinancialstatementsandnotestoaccountsproducedbyagency201314.xlsx | `passed` | `3,733` | `0` | `CC-BY` |
| Macroeconomic-data-2024-25-Budget-Update.xlsx | `unsupported` | `2,016` | `422` | `CC-BY` |
| Online-data-1.2-Taxation-Revenue-Qtrly-Budget-Update-2023-24.xlsx | `passed` | `1,071` | `0` | `CC-BY` |
| incidence-economique-des-entreprises-soutenues-par-linitiative-de-catalyse-du-capital-de-risque.xlsx | `unsupported` | `980` | `1` | `CA-OGL-LGO` |

## Licenses

| SPDX | Workbooks | Evidence |
| --- | ---: | --- |
| `Apache-2.0` | `1` | [Apache License 2.0](https://github.com/apache/poi/blob/trunk/LICENSE) |
| `CA-OGL-LGO` | `5` | [Open Government Licence - Canada](https://open.canada.ca/en/open-government-licence-canada) |
| `CC-BY` | `46` | [Creative Commons Attribution 3.0 Australia](http://creativecommons.org/licenses/by/3.0/au/) |

## Resource-Limited Cases

The scorecard classifies `14` workbooks as unsupported
because their footprint exceeds the configured
round-trip or structural-smoke resource budget. Those rows still have source,
license, hash, import, and classification evidence.

| Workbook | Cells | Classification |
| --- | ---: | --- |
| DataVic-dataset-VMIA-Five-year-summary-of-financial-results-2020-2024.XLSX | `54` | `xlsx.import.warning:Some cell styles were ignored during XLSX import.` |
| DataVic-dataset-VMIA-Five-year-summary-of-financial-results-2019-2023.xlsx | `54` | `xlsx.import.warning:Some cell styles were ignored during XLSX import.` |
| DataVic-dataset-VMIA-Five-year-summary-of-financial-results-2018-2022.xlsx | `54` | `xlsx.import.warning:Some cell styles were ignored during XLSX import.` |
| DataVic-dataset-VMIA-Five-year-summary-of-financial-results-2017-2021_0.xlsx | `54` | `xlsx.import.warning:Some cell styles were ignored during XLSX import.` |
| DataVic-dataset-VMIA-Five-year-summary-of-financial-results-2016-2020_0.xlsx | `54` | `xlsx.import.warning:Some cell styles were ignored during XLSX import.` |
| DataVic-dataset-VMIA-Five-year-summary-of-financial-results-2014-2018.XLSX | `54` | `xlsx.import.warning:Some cell styles were ignored during XLSX import.` |
| Macroeconomic-data-2024-25-Budget-Update.xlsx | `2,016` | `xlsx.import.warning:Some cell styles were ignored during XLSX import.` |
| sagovtprofilingdata091011121314.xlsx | `809` | `xlsx.import.warning:Manual calculation mode is preserved during XLSX import; cached formula values may be stale.` |
| Online-data-1.1-Taxation-Revenue-GG-Budget-Update-2023-24.xlsx | `13,508` | `xlsx.externalLinks.workbookReferencesPreserved`<br>`xlsx.import.warning:Defined-name formulas contain volatile, external, or unsupported Excel semantics; dependent cached formula values may change during recalculation.`<br>`xlsx.import.warning:External workbook links were preserved but not recalculated during XLSX import.` |
| performance-metrics-for-venture-capital-catalyst-initiative-2017-december-31-2024.xlsx | `908` | `xlsx.import.warning:Some cell styles were ignored during XLSX import.` |
| indicateurs-de-performance-de-linitiative-de-catalyse-du-capital-de-risque-2017-31-decembre-202.xlsx | `902` | `xlsx.import.warning:Some cell styles were ignored during XLSX import.` |
| economic-impact-of-companies-supported-by-venture-capital-catalyst-initiative-2017-december-31-.xlsx | `978` | `xlsx.import.warning:Some cell styles were ignored during XLSX import.` |
| incidence-economique-des-entreprises-soutenues-par-linitiative-de-catalyse-du-capital-de-risque.xlsx | `980` | `xlsx.import.warning:Some cell styles were ignored during XLSX import.` |
| sfgsme-efcpme-tables-2023-eng.xlsx | `246,084` | `xlsx.publicCorpus.resourceLimit:preflightRoundTripBudget>100000cells` |

## What This Proves

- The checked public baseline imports `52` cached public workbook
  files without failures or errors.
- It covers `280,480` cells, `529`
  formula cells, `217` sheets, and `3` source
  license families.
- Formula-oracle rows in this baseline match cached workbook evidence
  `529/529`.
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
