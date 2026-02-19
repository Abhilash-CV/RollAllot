import streamlit as st
import pandas as pd

st.set_page_config(page_title="Allotment + Reports", layout="wide")
st.title("🎓 Preference Based Centre Allotment + Reports")

# =========================
# Upload files
# =========================
cand_file = st.file_uploader("Upload Candidate Excel", type=["xlsx"])
lab_file = st.file_uploader("Upload Venue / Lab Excel", type=["xlsx"])

if not cand_file or not lab_file:
    st.info("Upload both files to continue")
    st.stop()

# =========================
# Read Excel
# =========================
df_cand = pd.read_excel(cand_file, engine="openpyxl")
df_lab = pd.read_excel(lab_file, engine="openpyxl")

df_cand.columns = df_cand.columns.str.strip()
df_lab.columns = df_lab.columns.str.strip()

# =========================
# Configuration
# =========================
st.subheader("⚙️ Configuration")

appl_col = st.selectbox("ApplNo Column", df_cand.columns, index=df_cand.columns.get_loc("ApplNo"))
date_col = st.selectbox("FSubDate Column", df_cand.columns, index=df_cand.columns.get_loc("FSubDate"))

pref_cols = st.multiselect(
    "Preference Columns (Order)",
    df_cand.columns,
    default=["Center1", "Center2", "Center3"]
)

name_col = st.selectbox("Candidate Name Column (optional)", ["None"] + list(df_cand.columns))

code_col = st.selectbox("Code Column", df_lab.columns)
venue_col = st.selectbox("Venue No Column", df_lab.columns)
centre_col = st.selectbox("Centre Name Column", df_lab.columns)
district_col = st.selectbox("District Column", df_lab.columns)
lab_col = st.selectbox("Lab Name Column", df_lab.columns)
strength_col = st.selectbox("Strength Column", df_lab.columns)

roll_start = st.number_input("Roll No Start", value=7100001, step=1)

# =========================
# Generate
# =========================
if st.button("🚀 Generate Allotment + Reports"):

    # Sort candidates
    df_cand[date_col] = pd.to_datetime(df_cand[date_col], errors="coerce")
    df_cand = df_cand.sort_values([appl_col, date_col]).reset_index(drop=True)
    df_cand["RollNo"] = range(roll_start, roll_start + len(df_cand))

    # Allotment columns
    df_cand["Allot_Code"] = None
    df_cand["Allot_Venue"] = None
    df_cand["Allot_Centre"] = None
    df_cand["Allot_District"] = None
    df_cand["Allot_Lab"] = None
    df_cand["Allot_Pref"] = None

    # Build lab capacity (row order preserved)
    labs = []
    for _, r in df_lab.iterrows():
        cap = pd.to_numeric(r[strength_col], errors="coerce")
        if pd.isna(cap) or cap <= 0:
            continue
        labs.append({
            "Code": r[code_col],
            "Venue": r[venue_col],
            "Centre": r[centre_col],
            "District": r[district_col],
            "Lab": r[lab_col],
            "Remaining": int(cap)
        })

    # Core allotment
    for i, cand in df_cand.iterrows():
        for p_idx, pref in enumerate(pref_cols, start=1):
            dist = cand[pref]
            if pd.isna(dist):
                continue
            for lab in labs:
                if lab["District"] == dist and lab["Remaining"] > 0:
                    df_cand.loc[i, [
                        "Allot_Code", "Allot_Venue", "Allot_Centre",
                        "Allot_District", "Allot_Lab", "Allot_Pref"
                    ]] = [
                        lab["Code"], lab["Venue"], lab["Centre"],
                        lab["District"], lab["Lab"], f"P{p_idx}"
                    ]
                    lab["Remaining"] -= 1
                    break
            if df_cand.at[i, "Allot_Pref"] is not None:
                break

    # =========================
    # REPORT 1: Preference Satisfaction
    # =========================
    # =========================
    # REPORT 1: Preference Satisfaction
    # =========================
    pref_report = (
        df_cand["Allot_Pref"]
        .fillna("Not Allotted")
        .value_counts()
        .reset_index()
        .rename(columns={"index": "Preference", "Allot_Pref": "Count"})
    )
    
    # 🔒 FORCE NUMERIC TYPE
    pref_report["Count"] = pd.to_numeric(
        pref_report["Count"], errors="coerce"
    )
    
    pref_report["Percentage"] = (
        pref_report["Count"] / float(len(df_cand)) * 100
    ).round(2)
    
    st.subheader("📊 Preference Satisfaction Report")
    st.dataframe(pref_report, use_container_width=True)

    # =========================
    # REPORT 2: Not Allotted
    # =========================
    not_allotted = df_cand[df_cand["Allot_Pref"].isna()]

    st.subheader("❌ Not Allotted Candidates")
    st.dataframe(not_allotted[[appl_col, "RollNo"]].head(), use_container_width=True)

    # =========================
    # REPORT 3: Venue-wise Attendance
    # =========================
    attendance_sheets = {}
    for venue, g in df_cand[df_cand["Allot_Venue"].notna()].groupby("Allot_Venue"):
        cols = ["RollNo", appl_col, "Allot_Lab"]
        if name_col != "None":
            cols.insert(2, name_col)
        sheet = g[cols].copy()
        sheet["Signature"] = ""
        attendance_sheets[f"Venue_{venue}"] = sheet

    # =========================
    # EXPORT EXCEL
    # =========================
    output = "allotment_with_reports.xlsx"
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_cand.to_excel(writer, sheet_name="Allotment", index=False)
        pref_report.to_excel(writer, sheet_name="Preference_Report", index=False)
        not_allotted.to_excel(writer, sheet_name="Not_Allotted", index=False)
        for name, sheet in attendance_sheets.items():
            sheet.to_excel(writer, sheet_name=name[:31], index=False)

    st.success("✅ Allotment & Reports Generated")

    with open(output, "rb") as f:
        st.download_button(
            "⬇️ Download Excel (All Reports)",
            f,
            file_name=output,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
