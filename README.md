# Architecture Renaming

Utilities for preparing, cleaning, renaming, auditing, and rolling back an EMKP architecture asset dataset.

The scripts are written for a Windows file layout and currently use fixed paths such as `E:\8-14_updated files`, `E:\16_Guide to the Dataset.xlsx`, and `E:\EMKP_work`. Update those constants in the scripts before running them on a different machine or dataset.

## Requirements

- Python 3.10 or newer
- `openpyxl` for reading the Excel guide
- `pandas` for `rename.py`
- `opencv-python` for MP4 duration reporting in `finalize_emkp_assets.py`

Install the dependencies with:

```bash
pip install openpyxl pandas opencv-python
```

## Directory Files

| File | Purpose |
| --- | --- |
| `finalize_emkp_assets.py` | Main final cleanup workflow. Reads the Excel guide, maps folder names to item IDs, deletes excluded file types and the `15_EXTRAS` folder, renames retained assets in place, writes an audit plan CSV, and writes a rollback CSV before making real changes. |
| `backup_assest.py` | Earlier/fixed-count version of the final cleanup workflow. It validates expected totals before cleanup and rename. The filename appears to preserve the original spelling, `assest`. |
| `rename_in_place.py` | In-place renaming workflow only. It reads the Excel guide, creates a folder-to-item mapping, generates final names using the `2024G07` prefix, writes a rollback CSV, and performs a two-phase rename to avoid filename collisions. |
| `rollback_rename.py` | Rollback utility for restoring original filenames from a rollback CSV. It defaults to dry-run mode and uses a two-phase rollback process to avoid collisions. |
| `rename.py` | CSV helper that reads the rollback CSV and adds filename-only columns derived from the renamed path. |
| `__pycache__/` | Python bytecode cache generated automatically when scripts run. This directory is ignored by Git. |

## Main Workflow

1. Review and update the path constants near the top of the script you plan to run.
2. Keep `DRY_RUN = True` for the first pass when using a destructive workflow.
3. Review the printed plan and generated CSV audit files.
4. Change `DRY_RUN = False` only after the plan is confirmed.
5. Preserve the rollback CSV. It can restore renamed filenames, but it cannot restore files that were permanently deleted.

## Output Files

The final cleanup scripts write audit files into `E:\EMKP_work`:

- `EMKP_final_plan.csv`: complete action plan containing rename and delete records.
- `EMKP_final_rename_rollback.csv`: rollback map for renamed files.
- `EMKP_final_rename_rollback_with_filenames.csv`: optional helper output from `rename.py`.

## Safety Notes

- `finalize_emkp_assets.py` and `backup_assest.py` can permanently delete files when `DRY_RUN = False`.
- `rollback_rename.py` can restore renamed filenames only if the renamed files and rollback CSV are still present.
- None of the rename workflows convert file contents; they only rename or delete files.
