# tirole-data-pipeline

A small ETL pipeline that turns the Tirole et al. (2022/2023)
*Double-slit time diffraction at optical frequencies* supplementary
spreadsheets into a tidy, queryable SQLite database — and a verifier that
proves the database is a **100% faithful image of the source workbooks**.

```
source_data/*.xlsx ──split──▶ pipeline/output/<wb>/<sheet>.xlsx
                   ──tidy───▶ pipeline/tidy_output/<group>/<fig>_tidy.csv
                   ──load───▶ output/double_slit.sqlite3   (one table per figure)
                   ──verify─▶ "FIDELITY: PASS"
```

## Quick start

```bash
# 0) dependencies (one self-contained virtualenv)
python -m venv .venv
.venv/bin/pip install -r requirements.txt
# activate it, or just call .venv/bin/python below

# 1) split workbooks + tidy CSVs + normalized DB + report   (writes under pipeline/)
cd pipeline
DATA_DIR=../source_data ../.venv/bin/python build_and_verify_pipeline.py run
#   ^ DATA_DIR defaults to ../source_data, so you can omit it

# 2) build the canonical per-figure DB                       (writes output/double_slit.sqlite3)
cd ../loader
../.venv/bin/python build_db_from_tidy_csv.py \
    --tidy_dir ../pipeline/tidy_output \
    --db       ../output/double_slit.sqlite3

# 3) prove 100% fidelity of the generated XLSX + SQLite vs the source XLSX
cd ..
.venv/bin/python verify/verify_fidelity.py              # core checks (1, 3, 4)
.venv/bin/python verify/verify_fidelity.py --reproduce  # + re-run the whole pipeline and diff (check 2)
```

> Paths are CWD-relative in the original scripts, so run each stage from the
> directory shown. The tidy CSVs land in **`pipeline/tidy_output/`** (not
> `output/tidy_output/`), which is why step 2 points `--tidy_dir` there.

Everything under `output/`, every `*.sqlite3`, and every `tidy_output/` is
git-ignored: a fresh clone ships only `source_data/` and the scripts, and you
regenerate the rest with the steps above.

## Layout

```
source_data/                         Tirole supplementary XLSX + the paper PDF
  Extended_Data_Fig_1_Source data.xlsx     (consumed; note the space in the name)
  Extended_Data_Fig_1_Source_data.xlsx     (unused byte-duplicate, underscore name)
  Extended_Data_Fig_2_Source_Data.xlsx
  Extended_Data_Fig_3_Source_data.xlsx
  Fig_1_Source_data.xlsx
  Fig_2_Source_data.xlsx
  double-slit-optical-2206.04362v2.pdf

pipeline/                            XLSX -> per-sheet XLSX -> tidy CSV
  build_and_verify_pipeline.py       orchestrator (split | tidy | load | verify | summary | run)
  ex2csv.py                          split each workbook into one XLSX per sheet
  extend-fig1-tidy-files.py          tidy Extended Fig 1a/1b/1c
  extend-fig2-tidy-files.py          tidy Extended Fig 2a/2b
  extend-fig3tidy-files.py           tidy Extended Fig 3a-3d
  extract_tidy_fig1d_and_Fig2a_2g.py tidy Fig 1d (and a standalone Fig 2 path)
  tidy-2a-2g.py                      tidy Fig 2a-2g

loader/
  build_db_from_tidy_csv.py          tidy CSV -> canonical per-figure SQLite

verify/
  verify_fidelity.py                 source XLSX <-> SQLite fidelity proof

output/                              generated, git-ignored
  double_slit.sqlite3                canonical DB: 17 figure tables + dataset_registry
```

## The scripts

All pipeline scripts are imported and driven by the orchestrator; each is also
runnable on its own. Run pipeline scripts from `pipeline/`, the loader from
`loader/`, and the verifier from the repo root.

### `pipeline/build_and_verify_pipeline.py` — orchestrator

```bash
python build_and_verify_pipeline.py <command>
```

| command   | what it does |
|-----------|--------------|
| `check`   | resolve and print `DATA_DIR` (the source-XLSX directory) |
| `split`   | run `ex2csv` -> `pipeline/output/<wb>/<sheet>.xlsx` |
| `tidy`    | run every tidy converter -> `pipeline/tidy_output/<group>/<fig>_tidy.csv` |
| `load`    | build the **normalized** DB `pipeline/double_slit.sqlite3` (see *SQLite output*) |
| `verify`  | row-count + first-row check of that DB vs the tidy CSVs -> `verification_report.csv` |
| `summary` | per-figure coverage counts -> `db_summary.csv` |
| `run`     | `split` -> `tidy` -> `load` -> `verify` |

`DATA_DIR` selects the source-XLSX directory; if unset it defaults to the
bundled `../source_data`. The five workbooks it consumes are listed in the
script's `EXPECTED_FILES`.

### `pipeline/ex2csv.py` — workbook splitter

`split_workbooks(src_dir, workbooks, output_dir)` opens each workbook and writes
one `.xlsx` per sheet under `output/<workbook_stem>/<sheet>.xlsx`. The only
change to cell content is cosmetic: column-name spaces become underscores
(`"Time (fs)"` -> `Time_(fs)`). Data values are copied unchanged. Imported by
the orchestrator; the module-level `workbooks = [...]` list is a commented
example.

### `pipeline/extend-fig1-tidy-files.py` — Extended Fig 1

`tidy_extended_fig1(src_dir, output_dir, make_plots=False)` reshapes the three
Extended Fig 1 panels, each with its own layout:
- **1a** delay vs reflectivity, two series (`actual`, `modelled`);
- **1b** delay x slit-separation reflectivity matrix, melted to long form;
- **1c** intensity vs reflectivity for two pump conditions.

Standalone: `python extend-fig1-tidy-files.py` (uses default relative paths;
needs the split files to exist first). Imports `matplotlib` at module top.

### `pipeline/extend-fig2-tidy-files.py` — Extended Fig 2

`tidy_extended_fig2(src_dir, output_dir)` passes Extended Fig 2a/2b through as
wide tables (drop unnamed columns, normalize column names).
*Note:* the module runs `tidy_extended_fig2()` once at import with default
paths — harmless (the orchestrator immediately calls it again with the right
paths), but it means importing the module touches the filesystem.

### `pipeline/extend-fig3tidy-files.py` — Extended Fig 3

`tidy_extended_fig3(src_dir, output_dir)` passes Extended Fig 3a-3d through as
wide tables. Fig 3d keeps its native **wide** layout (alternating
`Frequency - <rise> fs` / `Counts - <rise> fs` columns); the script's
long-format branch is not taken because the splitter already renamed
`fs rise time` to `fs_rise_time`, and downstream readers expect the wide form.
*Note:* like extend-fig2, this module calls `tidy_extended_fig3()` at import.

### `pipeline/extract_tidy_fig1d_and_Fig2a_2g.py` — Fig 1d

`tidy_fig_workbook(file_path, output_dir, make_plots=False)`. The orchestrator
uses it for **Fig 1d** (intensity vs reflectivity). The module also carries a
Fig 2 delay x slit-separation melt used by its own `__main__` demo. Imports
`matplotlib` at module top.

### `pipeline/tidy-2a-2g.py` — Fig 2a–2g

`tidy_fig2(file_path, output_dir)` handles the four Fig 2 layouts:
- **2a–2d** frequency-series spectra (`raw`, `smooth`, `model`);
- **2e** slit-separation vs oscillation period (`raw`, `model`, `fit`);
- **2f** slit-separation x frequency intensity matrix;
- **2g** slit-separation x delay matrix. **Declared unit transform:** the
  source delay axis is in seconds and is stored as femtoseconds
  (`Delay_fs = seconds * 1e15`). This is the only value transform anywhere in
  the pipeline, and the verifier inverts it.

### `loader/build_db_from_tidy_csv.py` — canonical DB builder

```bash
python build_db_from_tidy_csv.py --tidy_dir ../pipeline/tidy_output \
                                 --db ../output/double_slit.sqlite3
```

Loads each `*_tidy.csv` verbatim with `df.to_sql` into a table named after the
figure, and records `(dataset, relpath, nrows)` in a `dataset_registry` table.
Because the load is verbatim, every table is a 1:1 image of its tidy CSV.

### `verify/verify_fidelity.py` — fidelity proof

```bash
python verify/verify_fidelity.py [--db output/double_slit.sqlite3] [--reproduce]
```

See *Fidelity verification* below.

## SQLite output

The pipeline produces **two** databases with different purposes.

### Canonical — `output/double_slit.sqlite3` (the one to use)

Built by `loader/build_db_from_tidy_csv.py`. **18 tables = 17 figure tables +
`dataset_registry`**, each figure table a verbatim copy of its tidy CSV:

```
Fig_1d
Fig_2a Fig_2b Fig_2c Fig_2d Fig_2e Fig_2f Fig_2g
Extended_Fig_1a Extended_Fig_1b Extended_Fig_1c
Extended_Fig_2a Extended_Fig_2b
Extended_Fig_3a Extended_Fig_3b Extended_Fig_3c Extended_Fig_3d
dataset_registry
```

`Extended_Fig_3d` is the rise-time table (15 rows x 171 cols), paired
`Frequency`/`Counts` columns labelled by rise time, e.g.
*Counts – 3.6 fs rise time (MHz)*. Downstream code reads this wide layout
directly.

### Normalized — `pipeline/double_slit.sqlite3` (analytical view)

Built by the orchestrator's `load` stage. **3 tables** — `experiments`,
`spectra`, `time_domain` — that re-project the tidy CSVs into a long schema
with per-experiment slit separations (some inferred from the paper text). This
is a derived analytical view, **not** a 1:1 image of the source, so the 100%
fidelity proof below targets the canonical DB. The orchestrator's own `verify`
stage row-count-checks this normalized DB against the tidy CSVs.

## Fidelity verification

`verify/verify_fidelity.py` proves the generated Excel splits and the canonical
SQLite DB carry the source-workbook numbers with no loss or alteration. Run it
after the pipeline:

```bash
python verify/verify_fidelity.py              # checks 1, 3, 4 (pandas/openpyxl/numpy)
python verify/verify_fidelity.py --reproduce  # + check 2 (re-runs the pipeline; needs matplotlib)
```

It performs four checks:

| # | name | link | guarantee |
|---|------|------|-----------|
| 1 | **SPLIT** | source sheet <-> `pipeline/output/<wb>/<sheet>.xlsx` | every source cell appears in the split file, exact |
| 3 | **LOAD** | tidy CSV <-> table in `output/double_slit.sqlite3` | each DB table equals its tidy CSV cell-for-cell; `dataset_registry.nrows` agrees |
| 4 | **CONSERVE** | source workbook -> DB table | every numeric DB value traces to a source cell within double-precision tolerance, after inverting the declared Fig 2g `*1e15` |
| 2 | **REPRODUCE** *(opt.)* | committed DB <-> freshly re-run DB | re-running the whole pipeline in a temp tree reproduces the DB table-for-table |

The tidy step reshapes wide source tables into long form, so checks 1->3 are
not shape-preserving; check 4 is the value-level link across the reshape.
Together they establish **DB values are a subset of split values, which equal
source values**. The only declared value transform is the Fig 2g delay unit
change (seconds -> femtoseconds); every other value is conserved up to
spreadsheet/CSV float round-trip (largest observed relative deviation ~1e-12).

Latest run:

```
==================================================================
  CHECK 1 SPLIT        PASS
  CHECK 2 REPRODUCE    PASS
  CHECK 3 LOAD         PASS
  CHECK 4 CONSERVE     PASS
------------------------------------------------------------------
  FIDELITY: PASS  (68/68 items, source XLSX -> SQLite, full double precision)
==================================================================
```

A per-item CSV is written to `verify/fidelity_report.csv` (git-ignored).

## Source

Tirole, Vezzoli, Galiffi, Robertson, Maurice, Tilmann, Maier, Pendry,
Sapienza (2023), *Double-slit time diffraction at optical frequencies*,
**Nature Physics** 19, 999–1002.
DOI: [10.1038/s41567-023-01993-w](https://doi.org/10.1038/s41567-023-01993-w);
arXiv:2206.04362v2. The supplementary file `41567_2023_1993_MOESM6_ESM.xlsx`
is the upstream of the `source_data/` workbooks.

## What this directory is not

- Not a simulator and not a comparison harness — it produces the dataset only.
- No Lean4 or analysis content. Downstream consumers (e.g. the companion
  non-Hermitian double-slit simulator) read `output/double_slit.sqlite3`
  separately.
