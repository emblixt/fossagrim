# fossagrim

Scripts that prepare input to, and handle results from, the forest growth simulator
[Heureka](https://www.heurekaslu.se) for Fossagrim's project forests.

The scripts read a project's **Forvaltningsplan** (aka *Bestandsutvalg*, an Excel file with one row per forest
stand). They compute one area weighted *average stand* per tree species (Spruce, Pine, Birch) over the active stands,
and write a stand register that Heureka can import. Each average stand is simulated twice in Heureka:

* **BAU**: *business as usual*, with clear cutting and planting
* **PRES**: *preservation*

The simulation results are pasted back into an Excel file, and the scripts rearrange them and combine the tree
species, weighted by area, into one result per method.

## Installation

Requires Python 3.9 or newer.

```bash
pip install -e .            # pandas, numpy, matplotlib, openpyxl, xlsxwriter, pypdf
pip install -e .[test]      # + pytest
pip install -e .[plots]     # + bokeh, for plot_collected_stand_data()
```

## Workflow

All scripts can be run from any folder.

### 0. Quality check the input (optional)

```bash
python Fossagrim/qc_input_pdfs.py <forvaltningsplan.xlsx> <hovedtallsrapport.pdf> <sumtallsrapport.pdf> <project_name>
```

This compares the Forvaltningsplan with the Hovedtallsrapport and the Sumtallsrapport: productive area, volume and
selected stands. It also checks for harvest in MIS figures and in hogstklasse < IV, and for numbers that were written
as text. It writes `qc_log <project_name>.txt`, and `WARNING <project_name>.txt` if it finds problems, next to the
Forvaltningsplan.

### 1. Create the Heureka input

```bash
python Fossagrim/create_heureka_input_files.py <project_name> <project_folder> <forvaltningsplan.xlsx> [<unique_id>]
```

This creates two files in `project_folder`:

| File | Content |
|---|---|
| `<project_name> Averaged stand data.csv` | The stand register to import into Heureka, with one average stand per tree species. `unique_id` (optional) is written to the `UniqueID` column so that the Woods platform can tell versions apart |
| `<project_name> Heureka results.xlsx` | An empty workbook with one sheet per species and method, e.g. `FHF23-007 v01 Spruce BAU` |

Project names of 20 characters or more are cut to their first 10 characters in sheet names, because Excel allows
at most 31 characters.

### 2. Simulate in Heureka

Import the csv file and run BAU and PRES for each average stand. Then paste each result into its sheet of the
results workbook. The result must include the variable `Year`.

### 3. Rearrange the results

```bash
python Fossagrim/rearrange_heureka_results.py <project_name> <project_folder> <forvaltningsplan.xlsx> <heureka_results.xlsx>
```

This adds the sheet `Rearranged results`, which holds one table per simulation (one row per 5 year period) and one
area weighted combination of the tree species per method. If the sheet already exists, nothing is written. Delete
the sheet first to redo the step.

### Using the project settings file

`Fossagrim/utils/projects_new.py` runs steps 0, 1 and 3 from the project settings workbook
(`ProjectForestsSettings.xlsx`) instead of from command line arguments. Edit the variables at the bottom of the
file and run it.

## Code overview

| Module | Purpose |
|---|---|
| `create_heureka_input_files.py`, `rearrange_heureka_results.py`, `qc_input_pdfs.py` | Command line scripts (see above) |
| `io/common.py` | Read Excel files, and read/write the semicolon separated Heureka csv files |
| `io/stands.py` | Forvaltningsplan → Heureka: active stands, averaging, unit conversion, export |
| `io/heureka_results.py` | Read, rearrange and combine Heureka results; nature and climate effects (PRES − BAU) |
| `io/monetization.py` | The monetization Excel workbook (currently only used by the legacy workflow) |
| `io/fossagrim_io.py` | Backwards compatible entry point that re-exports all of the above |
| `utils/definitions.py` | Column names, Heureka keys, sheet names, averaging rules, colors |
| `utils/project_setup.py` | File and sheet naming conventions shared by the scripts |
| `utils/monetization_parameters.py` | Excel formulas of the monetization workbook |
| `plotting/misc_plots.py` | QC plots of stand data and Heureka results |
| `legacy/` | Code that is no longer used, kept for reference (see `legacy/__init__.py`) |

### How stands are averaged

`io/stands.py: average_over_stands()` treats each Forvaltningsplan column according to the lists in
`utils/definitions.py`:

* `fossagrim_keys_to_sum_over` (area and volumes): summed
* `fossagrim_keys_to_average_over` (age, height, density, ...): weighted by productive area (`Prod.areal`)
* `fossagrim_keys_to_most_of` (soil and vegetation codes): the value of the stand closest to the area weighted
  average
* All other columns: copied from the first stand

`arrange_export_of_one_stand()` then converts the units to Heureka's (daa → ha, m3fub/daa → m3sk/ha,
per daa → per ha).

## Tests

```bash
python -m pytest
```

The tests are in `Fossagrim/unit_tests`. Most of them use small synthetic files that are created on the fly
(see `conftest.py`). Tests in `test_real_data.py` use the real Forvaltningsplan in `Fossagrim/test_data/`, which is
client data and is ignored by git. These tests are skipped when the file is missing.

`Fossagrim/legacy/manual_checks.py` contains older manual checks that use files on the Fossagrim OneDrive. They are
not part of the automatic test suite.
