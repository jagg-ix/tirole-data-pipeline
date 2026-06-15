# tirole-data-pipeline

XLSX → tidy CSV → SQLite loader for the Tirole et al. (2022)
*Double-slit time diffraction at optical frequencies* supplementary
data.

Extracted from `entropic-time/multiphysics/catsim/` so the Python
loader does not sit next to an older Lean4 formalisation in the same
tree. No application logic (simulator, comparison harness) is
included here — this directory is the data pipeline only.

## Layout

```
source_data/                                 ← Tirole supplementary XLSX + PDF
  Extended_Data_Fig_{1,2,3}_Source_data.xlsx
  Fig_{1,2}_Source_data.xlsx
  double-slit-optical-2206.04362v2.pdf

pipeline/                                    ← XLSX → tidy CSV
  build_and_verify_pipeline.py               ← orchestrator
  ex2csv.py                                  ← XLSX → per-sheet split
  extend-fig{1,2,3}-tidy-files.py            ← per-figure tidy
  tidy-2a-2g.py
  extract_tidy_fig1d_and_Fig2a_2g.py

loader/                                      ← tidy CSV → SQLite
  build_db_from_tidy_csv.py

output/                                      ← generated artifacts (gitignored)
  output/<workbook_stem>/<sheet>.xlsx        ← from ex2csv split
  tidy_output/                               ← per-figure tidy CSVs
  double_slit.sqlite3                        ← final SQLite, 21 tables
```

The pipeline scripts are unchanged from their original
`entropic-time/multiphysics/catsim/` location. Path resolution is
CWD-relative, so invoke from `pipeline/`.

## Source

Tirole, Vezzoli, Galiffi, Robertson, Maurice, Tilmann, Maier,
Pendry, Sapienza (2022/2023), *Double-slit time diffraction at
optical frequencies*, **Nature Physics**.
DOI: [10.1038/s41567-023-01993-w](https://doi.org/10.1038/s41567-023-01993-w).
Supplementary file `41567_2023_1993_MOESM6_ESM.xlsx` is the upstream
of the source-data XLSX files in `source_data/`.

## Usage

```bash
# 1) install deps
pip install -r requirements.txt

# 2) run the pipeline from the pipeline/ directory
cd pipeline
python build_and_verify_pipeline.py            # full pipeline

# 3) build SQLite from the tidy CSVs
cd ../loader
python build_db_from_tidy_csv.py \
    --tidy_dir ../output/tidy_output \
    --db       ../output/double_slit.sqlite3
```

The orchestrator uses `DATA_DIR` env var to locate `source_data/`;
default is the bundled `../source_data` relative to the script.

## SQLite output

`output/double_slit.sqlite3` contains 21 tables, one per Tirole
figure plus a `dataset_registry`:

```
Fig_1d
Fig_2a, Fig_2b, Fig_2c, Fig_2d, Fig_2e, Fig_2f, Fig_2g
Extended_Fig_1{a,b,c}
Extended_Fig_2{a,b}
Extended_Fig_3{a,b,c,d}
dataset_registry
```

The most relevant table for rise-time-resolved verification is
`Extended_Fig_3d` (15 rows × 171 cols, paired Frequency/Counts rows
labelled by rise time, e.g. *Counts – 3.6 fs rise time (MHz)*).

## What this directory is not

* Not a simulator. No `time_double_slit.py` or comparable code.
* Not a comparison harness. No CAT-vs-data residual analysis.
* No Lean4 content. The pipeline produces SQLite; downstream
  formal-verification consumers read that SQLite separately.
