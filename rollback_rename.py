from pathlib import Path
import csv
import sys

ROLLBACK_CSV = Path(r"E:\rename_rollback.csv")

# True  = preview rollback only
# False = actually restore original filenames
DRY_RUN = True


def main():
    print()
    print("=" * 78)
    print("ROLLBACK DATASET FILENAMES")
    print("=" * 78)
    print(f"Rollback CSV: {ROLLBACK_CSV}")
    print(f"Dry run:      {DRY_RUN}")
    print()

    if not ROLLBACK_CSV.exists():
        print(f"ERROR: Rollback CSV not found: {ROLLBACK_CSV}")
        sys.exit(1)

    rows = []

    with ROLLBACK_CSV.open(
        "r",
        newline="",
        encoding="utf-8-sig"
    ) as f:
        reader = csv.DictReader(f)

        for row in reader:
            original = Path(row["Original Path"])
            renamed = Path(row["Renamed Path"])

            rows.append({
                "original": original,
                "renamed": renamed,
                "sequence": row["Sequence"],
            })

    print(f"Rollback entries: {len(rows):,}")
    print()

    errors = []

    for row in rows:
        if not row["renamed"].exists():
            errors.append(
                f"RENAMED FILE NOT FOUND: {row['renamed']}"
            )

        if row["original"].exists():
            errors.append(
                f"ORIGINAL PATH ALREADY EXISTS: {row['original']}"
            )

    if errors:
        print("ROLLBACK PRECHECK FAILED")
        print("-" * 78)
        for error in errors[:50]:
            print(error)

        if len(errors) > 50:
            print(f"... and {len(errors)-50} more errors.")

        print("\nNo rollback was performed.")
        sys.exit(1)

    for row in rows[:25]:
        print(
            f"{row['renamed'].name} "
            f"--> "
            f"{row['original'].name}"
        )

    if DRY_RUN:
        print()
        print("DRY RUN COMPLETE.")
        print("No files were changed.")
        print()
        print("To perform the rollback, change:")
        print("    DRY_RUN = True")
        print("to:")
        print("    DRY_RUN = False")
        return

    # --------------------------------------------------------
    # Two-phase rollback to prevent filename collisions
    # --------------------------------------------------------

    temp_rows = []

    try:
        # Phase 1: renamed -> temporary
        for i, row in enumerate(rows, start=1):
            renamed = row["renamed"]

            temp_path = renamed.with_name(
                f".__rollback_tmp__{i:06d}__{renamed.name}"
            )

            if temp_path.exists():
                raise FileExistsError(
                    f"Temporary rollback path exists: {temp_path}"
                )

            renamed.rename(temp_path)

            temp_rows.append({
                "temp": temp_path,
                "original": row["original"],
            })

            if i % 100 == 0 or i == len(rows):
                print(
                    f"Temporary rollback phase: "
                    f"{i:,}/{len(rows):,}"
                )

        # Phase 2: temporary -> original
        for i, row in enumerate(temp_rows, start=1):
            temp_path = row["temp"]
            original = row["original"]

            if original.exists():
                raise FileExistsError(
                    f"Original target unexpectedly exists: {original}"
                )

            temp_path.rename(original)

            if i % 100 == 0 or i == len(temp_rows):
                print(
                    f"Restore phase: "
                    f"{i:,}/{len(temp_rows):,}"
                )

    except Exception as error:
        print()
        print("ERROR DURING ROLLBACK:")
        print(error)
        print()
        print(
            "Some files may still have temporary names. "
            "Do not delete the rollback CSV."
        )
        sys.exit(2)

    print()
    print("=" * 78)
    print("ROLLBACK COMPLETE")
    print("=" * 78)
    print(f"Restored files: {len(rows):,}")
    print()


if __name__ == "__main__":
    main()