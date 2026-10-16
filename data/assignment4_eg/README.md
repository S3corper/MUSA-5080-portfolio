# Assignment 4 data and results

These files support `assignments/assignment4_eg.qmd`. The document executes from the portfolio root, as configured in `_quarto.yml`.

## Cached inputs

- `acs_2023.rds`: Philadelphia tract estimates, margins of error, and geometry from the 2019–2023 ACS five-year release.
- `burglary_2023.rds`: residential burglary records with dispatch dates in calendar year 2023.
- `shooting_victims_2023.rds`: shooting-victim records dated in calendar year 2023. Multiple victims can belong to one incident.
- `source_manifest.json`: source descriptions, filters, cache timestamps, and SHA-256 hashes of the raw inputs.

The QMD contains the complete download queries and the cleaning steps. The cached RDS files preserve the downloaded records before analysis. Historical event records may contain later revisions. Removing an input cache makes the next render download it again and can change the results.

## Generated outputs

- `model_log.csv`: mean ten-fold RMSE and improvements relative to each model's stated parent.
- `tract_diagnostics.csv`: outcome estimates, uncertainty, fitted residuals, and held-out errors for the fixed sample.
- `analysis_results.rds`: selected formula and fitted model, training-fold indices, ordered tract IDs, sector summaries, and model log.

From the portfolio root, run:

```sh
quarto render assignments/assignment4_eg.qmd --to html
```

The rendered report is `docs/assignments/assignment4_eg.html`. It embeds its figures. Packages used include tidyverse, tidycensus, sf, units, caret, and knitr; the report records `sessionInfo()`. A configured Census API key is used for a fresh ACS download but is not included in these files.
