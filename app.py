from pathlib import Path
import sys
import subprocess
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))
from analysis import load_data, clean_and_prepare, build_summaries, train_model, create_excel, REQUIRED

RAW = ROOT / "data" / "raw"
st.set_page_config(page_title="CH30 Campus Facility Insight", page_icon="🏫", layout="wide")
st.title("🏫 Campus Facility Insight")
st.caption("CH30 · Understand usage. Improve facilities. Support better planning.")
st.warning("Dataset status must be checked before using results. The generated demo data is synthetic and does not represent real campus measurements.")

with st.sidebar:
    st.header("Data & controls")
    st.write("Required files in `data/raw/`:")
    for f in REQUIRED:
        st.write("• " + f)
    if st.button("Generate synthetic demo data", use_container_width=True):
        proc = subprocess.run([sys.executable, str(ROOT / "generate_demo_data.py")], capture_output=True, text=True)
        if proc.returncode == 0:
            st.success("Synthetic demo data generated.")
            st.code(proc.stdout)
        else:
            st.error(proc.stderr)
    st.caption("To use actual assessment data, replace the five CSV files with the supplied datasets using the same column names.")

missing = [f for f in REQUIRED if not (RAW / f).exists()]
if missing:
    st.info("No complete dataset found yet. Click **Generate synthetic demo data** in the sidebar to test the system, or copy your actual CSV files into `data/raw/`.")
    st.stop()

try:
    raw = load_data(RAW)
    result = clean_and_prepare(raw)
    # Generated demo files include a small intentional missing value and duplicate; label by user choice/file marker.
    # We never claim files are real; user can switch the dataset label below.
    data_label = st.sidebar.radio("Current data label", ["Synthetic demo data", "Actual assessment data"], index=0)
    result["is_demo"] = data_label == "Synthetic demo data"
    summaries = build_summaries(result)
    model = train_model(result["occupancy"])
except Exception as e:
    st.error(f"Could not process the CSV files: {e}")
    st.write("Check filenames and required columns. See README.md for schemas.")
    st.stop()

if result["is_demo"]:
    st.warning("SYNTHETIC DEMO DATA — all figures and ML metrics below are computed from generated sample records. Do not submit them as real findings.")
else:
    st.success("Using files labeled as actual assessment data. Verify that they are the official dataset before submission.")

occ = result["occupancy"]
fac = summaries["facility_summary"]
mean_util = float(occ["utilization_pct"].mean()) if len(occ) else 0
peak = summaries["peak_hours"].iloc[0] if len(summaries["peak_hours"]) else None
top_fac = fac.iloc[0] if len(fac) else None
low_fac = summaries["underuse"].iloc[0] if len(summaries["underuse"]) else None

c1,c2,c3,c4 = st.columns(4)
c1.metric("Facilities", f'{result["facilities"]["facility"].nunique():,}')
c2.metric("Occupancy readings", f'{len(occ):,}')
c3.metric("Mean utilization", f"{mean_util:.1f}%")
c4.metric("Capacity exceedance readings", f'{int(occ["capacity_flag"].sum()):,}')

tab1,tab2,tab3,tab4,tab5 = st.tabs(["Overview","Facility usage","Peak times","Maintenance & exams","ML & data quality"])
with tab1:
    st.subheader("Facility utilization")
    st.dataframe(fac, use_container_width=True, hide_index=True)
    st.subheader("Initial observations")
    if top_fac is not None:
        st.write(f"• Highest average utilization in this dataset: **{top_fac['facility']}** ({top_fac['avg_utilization_pct']:.1f}%).")
    if low_fac is not None:
        st.write(f"• Lowest average utilization in this dataset: **{low_fac['facility']}** ({low_fac['avg_utilization_pct']:.1f}%).")
    if peak is not None:
        st.write(f"• Highest average occupancy by hour: **{int(peak['hour']):02d}:00** ({peak['avg_utilization_pct']:.1f}% average utilization).")
    st.caption("These are computed observations, not causal explanations. Interpret them in the context of the dataset label.")
with tab2:
    st.subheader("Average utilization by facility")
    fig, ax = plt.subplots()
    plot_df = fac.sort_values("avg_utilization_pct")
    ax.barh(plot_df["facility"], plot_df["avg_utilization_pct"])
    ax.set_xlabel("Average utilization (%)"); ax.set_ylabel("Facility"); ax.set_xlim(left=0)
    st.pyplot(fig); plt.close(fig)
    st.subheader("Capacity and occupancy")
    st.dataframe(fac[["facility","capacity","avg_occupancy","peak_occupancy","avg_utilization_pct","capacity_exceedances"]], use_container_width=True, hide_index=True)
with tab3:
    st.subheader("Average utilization by hour")
    fig, ax = plt.subplots()
    ph = summaries["peak_hours"].sort_values("hour")
    ax.plot(ph["hour"], ph["avg_utilization_pct"], marker="o")
    ax.set_xlabel("Hour of day"); ax.set_ylabel("Average utilization (%)"); ax.set_xticks(sorted(ph["hour"].unique()))
    ax.grid(True, alpha=.25)
    st.pyplot(fig); plt.close(fig)
    st.dataframe(ph, use_container_width=True, hide_index=True)
with tab4:
    st.subheader("Maintenance downtime")
    dw = summaries["downtime"]
    fig, ax = plt.subplots()
    if len(dw):
        ax.bar(dw["facility"], dw["total_downtime_hours"])
        ax.set_ylabel("Total recorded downtime (hours)")
        ax.tick_params(axis="x", rotation=45)
    st.pyplot(fig); plt.close(fig)
    st.subheader("Exam vs normal-period utilization")
    st.dataframe(summaries["exam_comparison"], use_container_width=True, hide_index=True)
with tab5:
    st.subheader("Model evaluation")
    if model.get("available"):
        m1,m2,m3 = st.columns(3)
        m1.metric("Accuracy", f'{model["accuracy"]*100:.2f}%')
        m2.metric("Macro F1-score", f'{model["f1_macro"]:.4f}')
        m3.metric("Test rows", f'{model["test_rows"]:,}')
        st.write("Algorithm: Random Forest Classifier")
        st.write("Target: utilization category (Low / Medium / High), derived from occupancy divided by capacity.")
        st.write("Input features:", ", ".join(model["features"]))
        st.caption("This is a classification of utilization categories from recorded patterns, not a forecast of future occupancy. Scores depend on dataset quality and should be reported with the dataset label.")
        with st.expander("Classification report"):
            st.json(model["report"])
    else:
        st.info(model.get("reason","Model unavailable"))
    st.subheader("Preprocessing audit")
    st.json(result["audit"])
    st.write("Cleaning steps: standardized facility identifiers, removed duplicate records, parsed timestamps, imputed missing occupancy/downtime values, and capped utilization counts at known capacity for analysis.")
    st.caption("Raw occupancy counts are retained in the `count` column. `count_for_utilization` is capped at capacity for utilization calculations, while `capacity_flag` records readings above capacity.")

st.divider()
st.subheader("Download outputs")
excel_path = create_excel(result, model, summaries)
with open(excel_path, "rb") as f:
    st.download_button("Download Excel analysis workbook", data=f, file_name="CH30_Campus_Facility_Analysis.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
st.download_button("Download cleaned occupancy CSV", data=occ.to_csv(index=False).encode("utf-8"),
                   file_name="clean_occupancy.csv", mime="text/csv")
st.caption("For final submission, use official assessment CSVs and rerun all analyses. Keep a note of whether the results came from synthetic demo data or official data.")
