# Previewing and publishing Assignment 5

Run commands in the RStudio Terminal from the portfolio folder.

To preview the existing website quickly:

```sh
quarto preview
```

To rebuild Assignment 5 after editing it:

```sh
quarto render assignments/assignment5.qmd
```

The first render runs the models and builds the cache. Later renders reuse
results when the R code and input CSVs are unchanged. The Census download is
saved at `data/assignment5/acs_tracts_income_2023.rds` for reuse. Run only one
render or preview at a time.

To force the models to run again (for example, after updating R packages):

```sh
quarto render assignments/assignment5.qmd --cache-refresh
```

After changing the input CSVs, use the Assignment 5 render command above;
project-level frozen output alone does not detect data-file changes.

The assignments listing automatically includes `assignment*.qmd`. To rebuild
just that menu without running models:

```sh
quarto render assignments/index.qmd
```

GitHub Pages serves the generated `docs/` directory. Commit and push the changed
source files, `docs/`, and `_freeze/` together. Uploading only a `.qmd` file does
not update the published HTML. Local `*_cache/` folders are ignored by Git.
