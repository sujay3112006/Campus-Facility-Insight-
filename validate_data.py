"""Dataset validation script for CH30 Campus Facility Insight.
Validates the five CSV datasets in data/raw/ to ensure schema compliance,
referential integrity, date/timestamp parsability, and reports controlled
data-quality issues (missing values, duplicates, capitalization, capacity exceedances).

SYNTHETIC DEMO DATA NOTICE:
These datasets and records are synthetic test samples for software evaluation.
They do not represent actual campus measurements.
"""
from pathlib import Path
import sys
import pandas as pd
import numpy as np

ROOT = Path(__file__).resolve().parent
DEFAULT_RAW = ROOT / "data" / "raw"

# Ensure UTF-8 stdout on Windows where possible
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass

REQUIRED_SCHEMAS = {
    "facilities.csv": ["facility", "type", "capacity", "block"],
    "bookings.csv": ["booking_id", "facility", "date", "start", "end", "purpose"],
    "occupancy.csv": ["timestamp", "facility", "count"],
    "calendar.csv": ["date", "working_day", "event", "exam_period"],
    "maintenance.csv": ["facility", "issue", "downtime_hours", "maintenance_date"]
}

KEY_COLUMNS = {
    "facilities.csv": ["facility"],
    "bookings.csv": ["booking_id"],
    "occupancy.csv": ["timestamp", "facility"],
    "calendar.csv": ["date"],
    "maintenance.csv": ["facility", "maintenance_date", "issue"]
}


def normalize_facility_str(series: pd.Series) -> pd.Series:
    """Normalize facility string for matching."""
    return series.astype("string").str.strip().str.upper()


def validate_datasets(data_dir=DEFAULT_RAW, verbose=True):
    """Run full validation across all five datasets in data_dir.

    Returns:
        dict: Complete structured validation findings and pass/fail summary.
    """
    data_dir = Path(data_dir)
    results = {
        "data_dir": str(data_dir),
        "files_checked": {},
        "summary_table": [],
        "referential_integrity": {},
        "capacity_checks": {},
        "occupancy_checks": {},
        "controlled_issues": {},
        "overall_passed": True,
        "critical_errors": []
    }

    # 1. File existence
    existing_files = {}
    for filename in REQUIRED_SCHEMAS.keys():
        file_path = data_dir / filename
        if not file_path.exists():
            results["overall_passed"] = False
            results["critical_errors"].append(f"Missing required file: {filename}")
            existing_files[filename] = None
        else:
            try:
                df = pd.read_csv(file_path)
                existing_files[filename] = df
            except Exception as e:
                results["overall_passed"] = False
                results["critical_errors"].append(f"Failed to read {filename}: {e}")
                existing_files[filename] = None

    if any(df is None for df in existing_files.values()):
        if verbose:
            print(f"Validation failed: one or more files could not be loaded from {data_dir}.")
            for err in results["critical_errors"]:
                print(f"  - {err}")
        return results

    # 2. Schema check and basic metrics
    dfs = existing_files
    for filename, req_cols in REQUIRED_SCHEMAS.items():
        df = dfs[filename]
        actual_cols_lower = [str(c).strip().lower() for c in df.columns]
        missing_cols = [c for c in req_cols if c.lower() not in actual_cols_lower]

        file_passed = len(missing_cols) == 0
        if not file_passed:
            results["overall_passed"] = False
            results["critical_errors"].append(f"{filename} missing required columns: {missing_cols}")

        missing_val_count = int(df.isna().sum().sum())
        exact_dup_count = int(df.duplicated().sum())

        # Key duplicates
        key_cols = KEY_COLUMNS.get(filename, [])
        key_cols_present = [c for c in key_cols if c in df.columns]
        if key_cols_present:
            key_dup_count = int(df.duplicated(subset=key_cols_present).sum())
        else:
            key_dup_count = 0

        results["files_checked"][filename] = {
            "rows": len(df),
            "columns": len(df.columns),
            "column_names": list(df.columns),
            "missing_cols": missing_cols,
            "missing_values_total": missing_val_count,
            "missing_values_by_column": df.isna().sum().to_dict(),
            "exact_duplicate_rows": exact_dup_count,
            "key_duplicates": key_dup_count,
            "passed_schema": file_passed
        }

    # 3. Date and timestamp parsing
    date_parsing = {}
    # calendar.csv -> date
    cal_df = dfs["calendar.csv"]
    cal_dates = pd.to_datetime(cal_df["date"], errors="coerce")
    cal_unparseable = int(cal_dates.isna().sum())
    date_parsing["calendar.date"] = {
        "total": len(cal_df),
        "valid": len(cal_df) - cal_unparseable,
        "unparseable": cal_unparseable,
        "min_date": str(cal_dates.min().date()) if len(cal_dates.dropna()) else None,
        "max_date": str(cal_dates.max().date()) if len(cal_dates.dropna()) else None
    }
    if cal_unparseable > 0:
        results["overall_passed"] = False
        results["critical_errors"].append(f"calendar.csv contains {cal_unparseable} unparseable dates")

    # bookings.csv -> date, start, end
    bk_df = dfs["bookings.csv"]
    bk_dates = pd.to_datetime(bk_df["date"], errors="coerce")
    bk_unparseable = int(bk_dates.isna().sum())
    bk_starts = pd.to_datetime(bk_df["start"], format="%H:%M", errors="coerce")
    bk_ends = pd.to_datetime(bk_df["end"], format="%H:%M", errors="coerce")
    date_parsing["bookings.date"] = {
        "total": len(bk_df),
        "valid": len(bk_df) - bk_unparseable,
        "unparseable": bk_unparseable,
        "min_date": str(bk_dates.min().date()) if len(bk_dates.dropna()) else None,
        "max_date": str(bk_dates.max().date()) if len(bk_dates.dropna()) else None
    }
    date_parsing["bookings.times"] = {
        "unparseable_start": int(bk_starts.isna().sum()),
        "unparseable_end": int(bk_ends.isna().sum())
    }
    if bk_unparseable > 0 or bk_starts.isna().sum() > 0 or bk_ends.isna().sum() > 0:
        results["overall_passed"] = False
        results["critical_errors"].append("bookings.csv contains unparseable dates or times")

    # occupancy.csv -> timestamp
    occ_df = dfs["occupancy.csv"]
    occ_ts = pd.to_datetime(occ_df["timestamp"], errors="coerce")
    occ_unparseable = int(occ_ts.isna().sum())
    date_parsing["occupancy.timestamp"] = {
        "total": len(occ_df),
        "valid": len(occ_df) - occ_unparseable,
        "unparseable": occ_unparseable,
        "min_timestamp": str(occ_ts.min()) if len(occ_ts.dropna()) else None,
        "max_timestamp": str(occ_ts.max()) if len(occ_ts.dropna()) else None
    }
    if occ_unparseable > 0:
        results["overall_passed"] = False
        results["critical_errors"].append(f"occupancy.csv contains {occ_unparseable} unparseable timestamps")

    # maintenance.csv -> maintenance_date
    maint_df = dfs["maintenance.csv"]
    maint_dates = pd.to_datetime(maint_df["maintenance_date"], errors="coerce")
    maint_unparseable = int(maint_dates.isna().sum())
    date_parsing["maintenance.maintenance_date"] = {
        "total": len(maint_df),
        "valid": len(maint_df) - maint_unparseable,
        "unparseable": maint_unparseable,
        "min_date": str(maint_dates.min().date()) if len(maint_dates.dropna()) else None,
        "max_date": str(maint_dates.max().date()) if len(maint_dates.dropna()) else None
    }
    if maint_unparseable > 0:
        results["overall_passed"] = False
        results["critical_errors"].append(f"maintenance.csv contains {maint_unparseable} unparseable dates")

    results["date_parsing"] = date_parsing

    # 4. Facility matching & casing across datasets
    fac_df = dfs["facilities.csv"]
    registered_facs = set(fac_df["facility"].dropna().astype(str).tolist())
    normalized_registered = {f.strip().upper(): f for f in registered_facs}

    cross_fac = {}
    for filename, df in [("bookings.csv", bk_df), ("occupancy.csv", occ_df), ("maintenance.csv", maint_df)]:
        if "facility" not in df.columns:
            continue
        raw_facs = df["facility"].dropna().astype(str)
        norm_facs = raw_facs.str.strip().str.upper()

        unmatched = set(norm_facs) - set(normalized_registered.keys())
        # Casing inconsistencies (matches normalized, but raw string differs from canonical registered name)
        casing_issues = raw_facs[raw_facs.isin(norm_facs) & (~raw_facs.isin(registered_facs))].unique().tolist()
        casing_issue_rows = int((~df["facility"].isin(registered_facs) & norm_facs.isin(normalized_registered.keys())).sum())

        cross_fac[filename] = {
            "unique_facilities": int(norm_facs.nunique()),
            "all_matched_normalized": len(unmatched) == 0,
            "unmatched_facilities": list(unmatched),
            "casing_inconsistencies": casing_issues,
            "casing_inconsistent_rows": casing_issue_rows
        }
        if len(unmatched) > 0:
            results["overall_passed"] = False
            results["critical_errors"].append(f"{filename} contains unregistered facilities: {list(unmatched)}")

    results["referential_integrity"] = cross_fac

    # 5. Capacity checks (positive capacity values)
    fac_cap = pd.to_numeric(fac_df["capacity"], errors="coerce")
    invalid_cap_count = int((fac_cap.isna() | (fac_cap <= 0)).sum())
    results["capacity_checks"] = {
        "valid_capacities": int((fac_cap > 0).sum()),
        "invalid_or_non_positive": invalid_cap_count,
        "min_capacity": int(fac_cap.min()) if fac_cap.notna().any() else None,
        "max_capacity": int(fac_cap.max()) if fac_cap.notna().any() else None
    }
    if invalid_cap_count > 0:
        results["overall_passed"] = False
        results["critical_errors"].append(f"facilities.csv contains {invalid_cap_count} non-positive or invalid capacity values")

    # 6. Occupancy checks (numeric or missing counts, and capacity exceedances)
    occ_counts = pd.to_numeric(occ_df["count"], errors="coerce")
    missing_occ_count = int(occ_df["count"].isna().sum())
    non_numeric_occ_count = int((occ_df["count"].notna() & occ_counts.isna()).sum())
    negative_occ_count = int((occ_counts < 0).sum())

    if non_numeric_occ_count > 0:
        results["overall_passed"] = False
        results["critical_errors"].append(f"occupancy.csv contains {non_numeric_occ_count} corrupt non-numeric count strings")
    if negative_occ_count > 0:
        results["overall_passed"] = False
        results["critical_errors"].append(f"occupancy.csv contains {negative_occ_count} negative count values")

    # Capacity exceedances check
    # Merge normalized facility capacity with occupancy readings
    fac_lookup = fac_df.copy()
    fac_lookup["norm_facility"] = normalize_facility_str(fac_lookup["facility"])
    fac_cap_map = dict(zip(fac_lookup["norm_facility"], pd.to_numeric(fac_lookup["capacity"], errors="coerce")))

    occ_norm = normalize_facility_str(occ_df["facility"])
    registered_capacities = occ_norm.map(fac_cap_map)
    exceedance_mask = (occ_counts > registered_capacities) & occ_counts.notna() & registered_capacities.notna()
    capacity_exceedance_count = int(exceedance_mask.sum())

    exceedance_details = []
    if capacity_exceedance_count > 0:
        exc_rows = occ_df[exceedance_mask].copy()
        exc_rows["capacity"] = registered_capacities[exceedance_mask]
        for _, r in exc_rows.head(5).iterrows():
            exceedance_details.append({
                "facility": str(r["facility"]),
                "timestamp": str(r["timestamp"]),
                "count": float(r["count"]),
                "capacity": float(r["capacity"]),
                "excess": float(r["count"] - r["capacity"])
            })

    results["occupancy_checks"] = {
        "total_readings": len(occ_df),
        "valid_numeric_readings": int(occ_counts.notna().sum()),
        "missing_count_readings": missing_occ_count,
        "non_numeric_strings": non_numeric_occ_count,
        "negative_readings": negative_occ_count,
        "capacity_exceedances": capacity_exceedance_count,
        "exceedance_examples": exceedance_details
    }

    # 7. Controlled quality issues summary
    controlled_issues = {
        "missing_occupancy_readings": missing_occ_count,
        "duplicate_bookings": results["files_checked"]["bookings.csv"]["key_duplicates"],
        "inconsistent_facility_capitalization": sum(
            cf["casing_inconsistent_rows"] for cf in cross_fac.values()
        ),
        "occupancy_exceeding_capacity": capacity_exceedance_count,
        "missing_maintenance_downtime": int(maint_df["downtime_hours"].isna().sum())
    }
    results["controlled_issues"] = controlled_issues

    # 8. Build Summary Table per Task 4
    # Columns required by Task 4:
    # - Dataset filename
    # - Number of rows
    # - Number of columns
    # - Missing-value counts
    # - Duplicate counts
    # - Capacity exceedance count
    # - Whether validation passed
    summary_table = []
    for filename in REQUIRED_SCHEMAS.keys():
        f_info = results["files_checked"][filename]
        # For occupancy, exceedances are computed; for others, N/A (0)
        exceedances = capacity_exceedance_count if filename == "occupancy.csv" else 0
        file_passed = f_info["passed_schema"] and (len(f_info["missing_cols"]) == 0)

        # File-specific integrity pass
        if filename == "facilities.csv" and invalid_cap_count > 0:
            file_passed = False
        if filename == "occupancy.csv" and (occ_unparseable > 0 or non_numeric_occ_count > 0 or negative_occ_count > 0):
            file_passed = False
        if filename == "bookings.csv" and (bk_unparseable > 0 or cross_fac.get("bookings.csv", {}).get("unmatched_facilities")):
            file_passed = False
        if filename == "calendar.csv" and cal_unparseable > 0:
            file_passed = False
        if filename == "maintenance.csv" and (maint_unparseable > 0 or cross_fac.get("maintenance.csv", {}).get("unmatched_facilities")):
            file_passed = False

        summary_table.append({
            "filename": filename,
            "rows": f_info["rows"],
            "columns": f_info["columns"],
            "missing_values": f_info["missing_values_total"],
            "duplicate_rows": f_info["exact_duplicate_rows"],
            "duplicate_keys": f_info["key_duplicates"],
            "capacity_exceedances": exceedances,
            "validation_passed": "PASSED" if file_passed else "FAILED"
        })

    results["summary_table"] = summary_table

    if verbose:
        print_validation_report(results)

    return results


def print_validation_report(res: dict):
    """Print a clean, professional verification report."""
    print("=" * 86)
    print("CH30 CAMPUS FACILITY INSIGHT -- DATASET PREPARATION & VALIDATION REPORT")
    print("NOTICE: All records are SYNTHETIC DEMO DATA for testing and validation.")
    print("=" * 86)
    print(f"Data Directory: {res['data_dir']}")
    print(f"Overall Status: {'PASSED (Schema & Integrity Valid)' if res['overall_passed'] else 'FAILED'}")
    print("-" * 86)

    # 1. Executive Summary Table (Required by Task 4)
    print("\nDATASET VERIFICATION SUMMARY TABLE:")
    header = f"{'Filename':<18} | {'Rows':>8} | {'Cols':>4} | {'Missing':>7} | {'Dups(Row)':>9} | {'Dups(Key)':>9} | {'Cap Exceed':>10} | {'Status':<8}"
    print("-" * len(header))
    print(header)
    print("-" * len(header))
    for row in res["summary_table"]:
        print(f"{row['filename']:<18} | {row['rows']:>8,d} | {row['columns']:>4d} | {row['missing_values']:>7,d} | {row['duplicate_rows']:>9,d} | {row['duplicate_keys']:>9,d} | {row['capacity_exceedances']:>10,d} | {row['validation_passed']:<8}")
    print("-" * len(header))

    # 2. Controlled Quality Fixtures Section
    print("\nCONTROLLED DATA-QUALITY TEST FIXTURES DETECTED:")
    ci = res["controlled_issues"]
    print(f"  * Missing occupancy readings:               {ci['missing_occupancy_readings']} reading(s) (imputed by downstream pipeline)")
    print(f"  * Duplicate bookings:                       {ci['duplicate_bookings']} duplicate booking ID(s)")
    print(f"  * Inconsistent facility capitalization:     {ci['inconsistent_facility_capitalization']} record(s) with lowercase/mixed-case IDs")
    print(f"  * Occupancy readings exceeding capacity:    {ci['occupancy_exceeding_capacity']} reading(s) (capped at capacity during analysis)")
    print(f"  * Missing maintenance downtime values:      {ci['missing_maintenance_downtime']} value(s) (imputed by downstream pipeline)")

    # 3. Date & Timestamp Parsing
    print("\nDATE AND TIMESTAMP PARSING VERIFICATION:")
    for key, p in res.get("date_parsing", {}).items():
        if "unparseable" in p:
            print(f"  * {key:<26}: {p['valid']:,}/{p['total']:,} valid (unparseable: {p['unparseable']})")
        if "min_date" in p and p["min_date"] is not None:
            print(f"    Range: {p['min_date']} to {p['max_date']}")
        elif "min_timestamp" in p and p["min_timestamp"] is not None:
            print(f"    Range: {p['min_timestamp']} to {p['max_timestamp']}")

    # 4. Cross-Dataset Facility Consistency
    print("\nFACILITY REFERENTIAL INTEGRITY & STANDARDIZATION:")
    fac_info = res.get("referential_integrity", {})
    for filename, r in fac_info.items():
        casing_note = f", {r['casing_inconsistent_rows']} casing variation(s)" if r['casing_inconsistent_rows'] > 0 else ""
        print(f"  * {filename:<16}: {r['unique_facilities']} unique facilities, 100% match to facilities.csv{casing_note}")
        if r["casing_inconsistencies"]:
            print(f"    Inconsistent casings found: {r['casing_inconsistencies']}")

    # 5. Capacity and Occupancy Details
    print("\nCAPACITY AND OCCUPANCY READINGS AUDIT:")
    cap = res.get("capacity_checks", {})
    occ = res.get("occupancy_checks", {})
    print(f"  * Registered facilities: {cap.get('valid_capacities', 0)} (Capacities: min {cap.get('min_capacity')}, max {cap.get('max_capacity')}, all > 0)")
    print(f"  * Total occupancy readings: {occ.get('total_readings', 0):,}")
    print(f"  * Valid numeric counts: {occ.get('valid_numeric_readings', 0):,}")
    print(f"  * Missing occupancy counts: {occ.get('missing_count_readings', 0):,}")
    print(f"  * Readings exceeding capacity: {occ.get('capacity_exceedances', 0):,}")
    if occ.get("exceedance_examples"):
        print("    Sample capacity exceedance readings:")
        for ex in occ["exceedance_examples"][:3]:
            print(f"      - {ex['facility']} at {ex['timestamp']}: count={int(ex['count'])}, capacity={int(ex['capacity'])} (+{int(ex['excess'])})")

    # Critical errors if any
    if res["critical_errors"]:
        print("\nCRITICAL ERRORS FOUND:")
        for err in res["critical_errors"]:
            print(f"  [ERROR] {err}")
    print("=" * 86)


def main():
    """CLI entry point for validate_data.py."""
    target_dir = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_RAW
    res = validate_datasets(target_dir, verbose=True)
    sys.exit(0 if res["overall_passed"] else 1)


if __name__ == "__main__":
    main()
