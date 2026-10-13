"""Generate clearly labeled SYNTHETIC demo data for CH30 Campus Facility Insight.
These records are for software testing and pipeline evaluation only,
not actual campus measurements.

Controlled data-quality anomalies injected for testing:
  1. Missing occupancy readings (1 reading with NaN count)
  2. Duplicate bookings (1 duplicate booking ID)
  3. Inconsistent facility-name capitalization (1 lowercase facility code: 'cl-a101')
  4. Occupancy readings exceeding registered capacity (9 readings with count > capacity)
  5. Missing maintenance downtime values (1 NaN downtime_hours)
  6. Duplicate occupancy row (1 exact duplicate reading)
"""
from pathlib import Path
import sys
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parent
RAW = ROOT / "data" / "raw"

# Ensure UTF-8 stdout on Windows where possible
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


def write_dataset_notice(output_dir: Path):
    """Write an explicit notice file confirming the synthetic status of raw files."""
    notice_text = (
        "CH30 CAMPUS FACILITY INSIGHT -- DATASET NOTICE\n"
        "=====================================================\n"
        "LABEL: SYNTHETIC DEMO DATA\n\n"
        "All files in this directory (facilities.csv, bookings.csv, occupancy.csv,\n"
        "calendar.csv, maintenance.csv) contain entirely synthetic campus facility\n"
        "usage data generated for software testing and DS Sprint evaluation.\n\n"
        "IMPORTANT RULES & TRANSPARENCY:\n"
        "1. These records are NOT actual campus measurements.\n"
        "2. Controlled data-quality anomalies have been intentionally injected\n"
        "   for pipeline testing (missing readings, duplicate bookings, inconsistent\n"
        "   casing, capacity exceedances, and missing maintenance downtime).\n"
        "3. For official submission, replace these files with official assessment data\n"
        "   maintaining the required column schemas.\n"
    )
    (output_dir / "DATASET_NOTICE.txt").write_text(notice_text, encoding="utf-8")


def generate_synthetic_datasets(seed: int = 30, num_days: int = 120,
                                start_date: str = "2026-06-16",
                                output_dir: Path = RAW) -> dict:
    """Generate the five synthetic campus datasets and save them to output_dir.

    Args:
        seed: Random seed for deterministic generation.
        num_days: Number of days to simulate (default 120).
        start_date: Start date string (YYYY-MM-DD). Defaults to 2026-06-16 (120 days).
        output_dir: Destination path for raw CSV files.

    Returns:
        dict: Mapping of filenames to generated pandas DataFrames.
    """
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    rng = np.random.default_rng(seed)

    # 1. facilities.csv
    facilities = pd.DataFrame([
        ("CL-A101", "Classroom", 60, "A"),
        ("CL-A102", "Classroom", 45, "A"),
        ("CL-B201", "Classroom", 70, "B"),
        ("LAB-C1", "Laboratory", 35, "C"),
        ("LAB-C2", "Laboratory", 28, "C"),
        ("LIB-MAIN", "Library", 120, "M"),
        ("AUD-1", "Auditorium", 200, "D"),
        ("STUDY-1", "Study Room", 20, "M")
    ], columns=["facility", "type", "capacity", "block"])

    # 2. calendar.csv (120 days)
    anchor_start = pd.Timestamp(start_date).normalize() if start_date else (pd.Timestamp.today().normalize() - pd.Timedelta(days=num_days - 1))
    dates = pd.date_range(anchor_start, periods=num_days, freq="D")
    calendar_rows = []
    for d in dates:
        is_workday = int(d.dayofweek < 5)
        exam = int((d.dayofyear % 60) in range(40, 47))
        event = "Exam" if exam else ("Campus Event" if d.dayofweek == 5 and d.day % 3 == 0 else "")
        calendar_rows.append((d.date().isoformat(), is_workday, event, exam))
    calendar = pd.DataFrame(calendar_rows, columns=["date", "working_day", "event", "exam_period"])

    # 3. bookings.csv
    bookings_rows = []
    booking_id = 1
    for d in dates:
        for _, f in facilities.iterrows():
            # Higher bookings on working days; library/study space can be used weekends too.
            probability = 0.62 if d.dayofweek < 5 else (0.32 if f["type"] in ["Library", "Study Room"] else 0.08)
            if rng.random() < probability:
                n = int(rng.integers(1, 4))
                for _ in range(n):
                    hour = int(rng.choice([8, 9, 10, 11, 12, 13, 14, 15, 16, 17],
                                          p=[.06, .12, .15, .14, .08, .12, .13, .10, .07, .03]))
                    duration = int(rng.choice([1, 2, 3], p=[.55, .35, .10]))
                    bookings_rows.append((
                        booking_id, f["facility"], d.date().isoformat(),
                        f"{hour:02d}:00", f"{min(hour + duration, 19):02d}:00",
                        rng.choice(["Class", "Lab Session", "Study", "Meeting", "Exam"])
                    ))
                    booking_id += 1
    bookings = pd.DataFrame(bookings_rows, columns=["booking_id", "facility", "date", "start", "end", "purpose"])

    # 4. occupancy.csv (hourly readings 08:00 - 18:00)
    occ_rows = []
    base_by_type = {"Classroom": 0.47, "Laboratory": 0.58, "Library": 0.52, "Auditorium": 0.18, "Study Room": 0.35}
    for d in dates:
        exam = int(calendar.loc[calendar["date"] == d.date().isoformat(), "exam_period"].iloc[0])
        for _, f in facilities.iterrows():
            for hour in range(8, 19):
                weekday_factor = 1.0 if d.dayofweek < 5 else 0.50
                peak_factor = 1.25 if hour in [10, 11, 14, 15] else (0.72 if hour in [8, 17, 18] else 1.0)
                exam_factor = 1.30 if exam and f["type"] in ["Library", "Study Room", "Classroom"] else 1.0
                ratio = base_by_type[f["type"]] * weekday_factor * peak_factor * exam_factor + rng.normal(0, 0.12)
                # Upper bound 1.08 allows realistic capacity exceedance test cases
                ratio = float(np.clip(ratio, 0.02, 1.08))
                count = int(round(f["capacity"] * ratio))
                occ_rows.append((pd.Timestamp(d) + pd.Timedelta(hours=hour), f["facility"], count))
    occupancy = pd.DataFrame(occ_rows, columns=["timestamp", "facility", "count"])

    # 5. maintenance.csv
    maintenance_rows = []
    for _, f in facilities.iterrows():
        for _ in range(int(rng.integers(2, 8))):
            d = rng.choice(dates)
            maintenance_rows.append((
                f["facility"],
                rng.choice(["Electrical", "AC", "Furniture", "Network", "Equipment"]),
                float(round(rng.choice([1, 2, 3, 4, 6, 8]) * rng.uniform(.7, 1.3), 1)),
                pd.Timestamp(d).date().isoformat()
            ))
    maintenance = pd.DataFrame(maintenance_rows, columns=["facility", "issue", "downtime_hours", "maintenance_date"])

    # Inject controlled data-quality test fixtures
    # A. Missing occupancy count (index 5)
    if len(occupancy) > 20:
        occupancy.loc[5, "count"] = np.nan
        # B. Duplicate occupancy row (index 10)
        occupancy = pd.concat([occupancy, occupancy.iloc[[10]]], ignore_index=True)

    # C. Duplicate booking record (duplicate booking_id 1)
    if len(bookings):
        bookings = pd.concat([bookings, bookings.iloc[[0]]], ignore_index=True)
        # D. Inconsistent facility capitalization ('cl-a101' lowercase in first row)
        bookings.loc[bookings.index[0], "facility"] = str(bookings.loc[bookings.index[0], "facility"]).lower()

    # E. Missing maintenance downtime value (last row)
    if len(maintenance):
        maintenance.loc[maintenance.index[-1], "downtime_hours"] = np.nan

    datasets = {
        "facilities.csv": facilities,
        "bookings.csv": bookings,
        "occupancy.csv": occupancy,
        "calendar.csv": calendar,
        "maintenance.csv": maintenance
    }

    # Save to data/raw/
    for name, df in datasets.items():
        df.to_csv(output_dir / name, index=False)

    write_dataset_notice(output_dir)
    return datasets


def main():
    """CLI entry point for generate_demo_data.py."""
    print("Generating SYNTHETIC demo datasets in data/raw/...")
    datasets = generate_synthetic_datasets(seed=30, num_days=120, output_dir=RAW)
    print("SUCCESS: 5 synthetic datasets written to data/raw/.")
    print("NOTICE: These records are SYNTHETIC DEMO DATA for testing only, NOT real campus data.\n")

    # Automatically validate the generated datasets
    try:
        from validate_data import validate_datasets
        validate_datasets(data_dir=RAW, verbose=True)
    except ImportError:
        for name, df in datasets.items():
            print(f"  {name}: {len(df):,} records, {len(df.columns)} attributes")


if __name__ == "__main__":
    main()
