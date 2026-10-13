"""Data cleaning, EDA, feature engineering, ML evaluation and Excel export for CH30."""
from pathlib import Path
import numpy as np
import pandas as pd
from sklearn.compose import ColumnTransformer
from sklearn.preprocessing import OneHotEncoder
from sklearn.pipeline import Pipeline
from sklearn.impute import SimpleImputer
from sklearn.ensemble import RandomForestClassifier
from sklearn.model_selection import train_test_split
from sklearn.metrics import accuracy_score, f1_score, classification_report
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, Reference

ROOT = Path(__file__).resolve().parent
RAW = ROOT / "data" / "raw"
PROCESSED = ROOT / "data" / "processed"
OUTPUTS = ROOT / "outputs"
REQUIRED = ["facilities.csv","bookings.csv","occupancy.csv","calendar.csv","maintenance.csv"]

def load_data(raw_dir=RAW):
    raw_dir = Path(raw_dir)
    missing = [f for f in REQUIRED if not (raw_dir / f).exists()]
    if missing:
        raise FileNotFoundError("Missing dataset files: " + ", ".join(missing))
    return {Path(f).stem: pd.read_csv(raw_dir / f) for f in REQUIRED}

def _normalize_facility(series):
    return series.astype("string").str.strip().str.upper()

def clean_and_prepare(d):
    facilities = d["facilities"].copy()
    bookings = d["bookings"].copy()
    occupancy = d["occupancy"].copy()
    calendar = d["calendar"].copy()
    maintenance = d["maintenance"].copy()

    facilities.columns = facilities.columns.str.strip().str.lower()
    bookings.columns = bookings.columns.str.strip().str.lower()
    occupancy.columns = occupancy.columns.str.strip().str.lower()
    calendar.columns = calendar.columns.str.strip().str.lower()
    maintenance.columns = maintenance.columns.str.strip().str.lower()

    for df in [facilities, bookings, occupancy, maintenance]:
        if "facility" in df:
            df["facility"] = _normalize_facility(df["facility"])
    if "facility" in facilities:
        facilities = facilities.drop_duplicates(subset=["facility"], keep="last")
    if "booking_id" in bookings:
        bookings = bookings.drop_duplicates(subset=["booking_id"], keep="first")
    else:
        bookings = bookings.drop_duplicates()
    occupancy = occupancy.drop_duplicates()
    maintenance = maintenance.drop_duplicates()

    facilities["capacity"] = pd.to_numeric(facilities["capacity"], errors="coerce")
    facilities["capacity"] = facilities["capacity"].where(facilities["capacity"] > 0)
    occupancy["timestamp"] = pd.to_datetime(occupancy["timestamp"], errors="coerce")
    occupancy["count"] = pd.to_numeric(occupancy["count"], errors="coerce")
    occupancy.loc[occupancy["count"] < 0, "count"] = np.nan
    occupancy = occupancy.dropna(subset=["timestamp", "facility"])
    occupancy["date"] = occupancy["timestamp"].dt.strftime("%Y-%m-%d")
    occupancy["hour"] = occupancy["timestamp"].dt.hour
    # Median imputation by facility, with global median as fallback.
    facility_medians = occupancy.groupby("facility")["count"].transform("median")
    occupancy["count"] = occupancy["count"].fillna(facility_medians).fillna(0)
    # Keep impossible readings for audit but cap at capacity for utilization analysis.
    occupancy = occupancy.merge(facilities[["facility","capacity","type","block"]], on="facility", how="left")
    occupancy["capacity_flag"] = occupancy["count"] > occupancy["capacity"]
    occupancy["count_for_utilization"] = np.where(
        occupancy["capacity"].notna(),
        np.minimum(occupancy["count"], occupancy["capacity"]),
        occupancy["count"]
    )
    occupancy["utilization_ratio"] = np.where(
        occupancy["capacity"].notna() & (occupancy["capacity"] > 0),
        occupancy["count_for_utilization"] / occupancy["capacity"], np.nan
    )
    # Outlier treatment: retain source count, but cap analysis count at registered capacity.
    occupancy["utilization_pct"] = occupancy["utilization_ratio"] * 100

    calendar["date"] = pd.to_datetime(calendar["date"], errors="coerce").dt.strftime("%Y-%m-%d")
    if "working_day" in calendar:
        calendar["working_day"] = pd.to_numeric(calendar["working_day"], errors="coerce").fillna(0).astype(int)
    if "exam_period" in calendar:
        calendar["exam_period"] = pd.to_numeric(calendar["exam_period"], errors="coerce").fillna(0).astype(int)
    calendar = calendar.drop_duplicates(subset=["date"], keep="last")

    bookings["date"] = pd.to_datetime(bookings["date"], errors="coerce").dt.strftime("%Y-%m-%d")
    if "start" in bookings:
        bookings["booking_hour"] = pd.to_datetime(bookings["start"], format="%H:%M", errors="coerce").dt.hour
    else:
        bookings["booking_hour"] = np.nan
    booking_counts = bookings.groupby(["facility","date","booking_hour"], dropna=False).size().reset_index(name="booking_count")
    booking_counts = booking_counts.rename(columns={"booking_hour":"hour"})

    maintenance["downtime_hours"] = pd.to_numeric(maintenance["downtime_hours"], errors="coerce")
    maintenance.loc[maintenance["downtime_hours"] < 0, "downtime_hours"] = np.nan
    maintenance["downtime_hours"] = maintenance["downtime_hours"].fillna(maintenance["downtime_hours"].median()).fillna(0)
    downtime = maintenance.groupby("facility", as_index=False)["downtime_hours"].sum().rename(columns={"downtime_hours":"total_downtime_hours"})

    occupancy = occupancy.merge(calendar, on="date", how="left", suffixes=("","_calendar"))
    occupancy["working_day"] = pd.to_numeric(occupancy.get("working_day", 0), errors="coerce").fillna(0).astype(int)
    occupancy["exam_period"] = pd.to_numeric(occupancy.get("exam_period", 0), errors="coerce").fillna(0).astype(int)
    occupancy["day_of_week"] = pd.to_datetime(occupancy["date"]).dt.dayofweek
    occupancy = occupancy.merge(booking_counts, on=["facility","date","hour"], how="left")
    occupancy["booking_count"] = occupancy["booking_count"].fillna(0)
    occupancy = occupancy.merge(downtime, on="facility", how="left")
    occupancy["total_downtime_hours"] = occupancy["total_downtime_hours"].fillna(0)
    occupancy["utilization_category"] = pd.cut(
        occupancy["utilization_ratio"], bins=[-0.001, 0.30, 0.70, 1.001],
        labels=["Low","Medium","High"], include_lowest=True
    ).astype("string").fillna("Unknown")

    PROCESSED.mkdir(parents=True, exist_ok=True)
    for name, df in [("clean_occupancy.csv",occupancy),("clean_bookings.csv",bookings),
                     ("clean_facilities.csv",facilities),("clean_calendar.csv",calendar),
                     ("clean_maintenance.csv",maintenance)]:
        df.to_csv(PROCESSED / name, index=False)
    audit = {
        "raw_occupancy_rows": int(len(d["occupancy"])),
        "clean_occupancy_rows": int(len(occupancy)),
        "raw_booking_rows": int(len(d["bookings"])),
        "clean_booking_rows": int(len(bookings)),
        "missing_occupancy_values_imputed": int(pd.to_numeric(d["occupancy"]["count"], errors="coerce").isna().sum()),
        "capacity_exceedance_readings": int(occupancy["capacity_flag"].sum()),
        "facilities": int(facilities["facility"].nunique()),
        "occupancy_missing_after_cleaning": int(occupancy["count"].isna().sum())
    }
    return {"facilities":facilities, "bookings":bookings, "occupancy":occupancy,
            "calendar":calendar, "maintenance":maintenance, "downtime":downtime, "audit":audit}

def train_model(occ):
    df = occ.loc[occ["utilization_category"].isin(["Low","Medium","High"])].copy()
    feature_cols = ["type","block","hour","day_of_week","working_day","exam_period",
                    "booking_count","total_downtime_hours"]
    target = "utilization_category"
    X = df[feature_cols]
    y = df[target]
    if len(df) < 30 or y.nunique() < 2:
        return {"available":False, "reason":"Not enough labeled rows/classes to evaluate a model.", "features":feature_cols}
    categorical = ["type","block"]
    numerical = [c for c in feature_cols if c not in categorical]
    preprocessor = ColumnTransformer([
        ("cat", Pipeline([("imputer",SimpleImputer(strategy="most_frequent")),
                          ("onehot",OneHotEncoder(handle_unknown="ignore"))]), categorical),
        ("num", Pipeline([("imputer",SimpleImputer(strategy="median"))]), numerical)
    ])
    model = Pipeline([
        ("prep",preprocessor),
        ("classifier",RandomForestClassifier(n_estimators=150, random_state=30, class_weight="balanced"))
    ])
    try:
        X_train, X_test, y_train, y_test = train_test_split(
            X, y, test_size=0.25, random_state=30, stratify=y
        )
    except ValueError:
        X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.25, random_state=30)
    model.fit(X_train, y_train)
    pred = model.predict(X_test)
    return {"available":True, "model":model, "accuracy":float(accuracy_score(y_test,pred)),
            "f1_macro":float(f1_score(y_test,pred,average="macro",zero_division=0)),
            "test_rows":int(len(y_test)), "train_rows":int(len(y_train)),
            "features":feature_cols, "target":target,
            "report":classification_report(y_test,pred,zero_division=0,output_dict=True),
            "predictions":pd.DataFrame({"actual":y_test.to_numpy(),"predicted":pred})}

def build_summaries(result):
    occ = result["occupancy"]
    facility_summary = occ.groupby(["facility","type","block","capacity"], dropna=False).agg(
        avg_occupancy=("count_for_utilization","mean"),
        avg_utilization_pct=("utilization_pct","mean"),
        peak_occupancy=("count_for_utilization","max"),
        readings=("timestamp","count"),
        capacity_exceedances=("capacity_flag","sum")
    ).reset_index().sort_values("avg_utilization_pct", ascending=False)
    peak_hours = occ.groupby("hour", as_index=False).agg(
        avg_occupancy=("count_for_utilization","mean"), avg_utilization_pct=("utilization_pct","mean")
    ).sort_values("avg_utilization_pct", ascending=False)
    underuse = facility_summary.sort_values("avg_utilization_pct").copy()
    cal = occ.groupby("exam_period", as_index=False)["utilization_pct"].mean()
    cal["period"] = cal["exam_period"].map({0:"Normal period",1:"Exam period"}).fillna("Unknown")
    exam_comparison = cal[["period","utilization_pct"]]
    downtime = result["downtime"].merge(result["facilities"][["facility","type","block"]], on="facility", how="left")
    downtime = downtime.sort_values("total_downtime_hours", ascending=False)
    return {"facility_summary":facility_summary, "peak_hours":peak_hours, "underuse":underuse,
            "exam_comparison":exam_comparison, "downtime":downtime}

def create_excel(result=None, model_result=None, summaries=None, path=None):
    """Generate the complete 12-worksheet CH30 Excel workbook."""
    from export_excel import generate_excel_workbook
    OUTPUTS.mkdir(parents=True, exist_ok=True)
    target_path = Path(path) if path else OUTPUTS / "CH30_Campus_Facility_Analysis.xlsx"
    return generate_excel_workbook(raw_dir=RAW, output_path=target_path)

def run_pipeline(raw_dir=RAW, is_demo=False):
    data = load_data(raw_dir)
    result = clean_and_prepare(data)
    result["is_demo"] = is_demo
    summaries = build_summaries(result)
    model_result = train_model(result["occupancy"])
    excel_path = create_excel(result, model_result, summaries)
    return result, summaries, model_result, excel_path

if __name__ == "__main__":
    print("Running CH30 Campus Facility Insight Pipeline...")
    res, sums, mod, excel_path = run_pipeline(is_demo=True)
    print(f"Analysis complete. Deliverable workbook: {excel_path}")
