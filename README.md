# CAT Modeling CSIC

Given supplied data, creates `Account.csv` and `Location.csv` files for input to a CAT model. Also exports `WindModeling_Output.xlsx` file exhibiting renamed/derived columns for input verification.


## What it does

Starting from one `.xlsx` in `data/`, the script:

1. Renames/derives columns the required for CAT model input.
2. Builds `Account` and `Location` tables across four scenarios:
   - `2pct` SCS (severe convective storm)
   - `2pct`, `3pct`, and `5pct` WS (wind/hurricane) site deductibles
3. Writes results to `output/`:
   - `WindModeling_Output.xlsx` — the original workbook copied with `ModelingSOV`, `Account`, and `Location`
     sheets appended; model input columns are highlighted yellow.
   - `Account.csv` and `Location.csv` — model-ready imports.

## Repository layout

| Path | Description |
| --- | --- |
| `catmodeling.py` | Main script. |
| `/Legacy Files` | Files with similar functionality written by summer interns. |
| `catmodeling.ipynb` | Notebook version used for development. |
| `data/` | Drop the single source `.xlsx` here. |
| `output/` | Generated workbook and CSVs will output here. |

## Requirements

Runs with dependencies:

- `pandas`
- `numpy`
- `openpyxl`
- `python-calamine`

## Usage

Place **exactly one** `.xlsx` source file in `data/`, then run:

```bash
python catmodeling.py
```

Outputs are written to `output/`.

## Notes

- Notebook outputs are stripped on commit via `nbstripout`, so no read data is pushed to the remote. After cloning, run `nbstripout --install` to enable the same filter locally.
