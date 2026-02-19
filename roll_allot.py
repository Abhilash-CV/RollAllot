import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

# -------------------------------------------------
# PAGE SETUP
# -------------------------------------------------
st.set_page_config(page_title="Exam Allotment System", layout="wide")
st.title("🎓 Exam Centre & Lab Allotment System")

# -------------------------------------------------
# FILE UPLOAD
# -------------------------------------------------
cand_file = st.file_uploader("Upload Candidate Excel", type=["xlsx"])
lab_file = st.file_uploader("Upload Venue / Lab Excel", type=["xlsx"])

if not cand_file or not lab_file:
    st.info("📂 Upload both Candidate and Venue/Lab Excel files to continue.")
    st.stop()

# -------------------------------------------------
# READ FILES
# -------------------------------------------------
df_cand = pd.read_excel(cand_file, engine="openpyxl")
df_lab = pd.read_excel(lab_file, engine="openpyxl")

df_cand.columns = df_cand.columns.str.strip()
df_lab.columns = df_lab.columns.str.strip()

# -------------------------------------------------
# CONFIGURATION
# -------------------------------------------------
st.subheader("⚙️ Configuration")

appl_col = st.selectbox("Application No Column", df_cand.columns)
date_col = st.selectbox("Final Submission Date Column", df_cand.columns)

pref_cols = st.multiselect(
    "Preference Columns (District order)",
    df_cand.columns,
    default=["Center1", "Center2", "Center3"]
)

name_col = st.selectbox("Candidate Name Column (optional)", ["None"] + list(df_cand.columns))

st.markdown("### 🏫 Venue / Lab Master Columns")

code_col = st.selectbox("Code Column", df_lab.columns)
venue_col = st.selectbox("Venue No Column", df_lab.columns)
centre_col = st.selectbox("Centre Name Column", df_lab.columns)
district_col = st.selectbox("District Column", df_lab.columns)
lab_col = st.selectbox("Lab Name Column", df_lab.columns)
strength_col = st.selectbox("Strength Column", df_lab.columns)

roll_start = st.number_input("Roll Number Start From", value=7100001, step=1)

# -------------------------------------------------
# PROCESS
# -------------------------------------------------
if st.button("🚀 Generate Allotment + Reports"):

    # -----------------------------
    # SORT CANDIDATES
    # -----------------------------
    df_cand[date_col] = pd.to_datetime(df_cand[date_col], errors="coerce")
    df_cand = df_cand.sort_values([appl_col, date_col]).reset_index(drop=True)
    df_cand["RollNo"] = range(roll_start, roll_start + len(df_cand))

    # -----------------------------
    # OUTPUT COLUMNS
    # -----------------------------
    df_cand["Allot_Code"] = None
    df_cand["Allot_Venue"] = None
    df_cand["Allot_Centre"] = None
    df_cand["Allot_District"] = None
    df_cand["Allot_Lab"] = None
    df_cand["Allot_Pref"] = None

    # -----------------------------
    # BUILD LAB CAPACITY (ROW ORDER PRESERVED)
    # -----------------------------
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
            "Remaining": int(cap),
            "Strength": int(cap)
        })

    # -----------------------------
    # CORE ALLOTMENT
    # -----------------------------
    for i, cand in df_cand.iterrows():
        for p_idx, pref in enumerate(pref_cols, start=1):
            pref_dist = cand[pref]
            if pd.isna(pref_dist):
                continue
            for lab in labs:
                if lab["District"] == pref_dist and lab["Remaining"] > 0:
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

    # =========================================================
    # REPORT 1: PREFERENCE SATISFACTION (OVERALL)
    # =========================================================
    pref_report = (
        df_cand["Allot_Pref"]
        .fillna("Not Allotted")
        .value_counts()
        .reset_index()
    )

    # FORCE STABLE COLUMNS
    pref_report.columns = ["Preference", "Count"]
    pref_report["Count"] = pd.to_numeric(pref_report["Count"], errors="coerce")
    pref_report["Percentage"] = (
        pref_report["Count"] / float(len(df_cand)) * 100
    ).round(2)

    st.subheader("📊 Preference Satisfaction (Overall)")
    st.dataframe(pref_report, use_container_width=True)

    # =========================================================
    # REPORT 2: DISTRICT-WISE PREFERENCE
    # =========================================================
    district_pref = (
        df_cand
        .assign(Preference=df_cand["Allot_Pref"].fillna("Not Allotted"))
        .groupby(["Allot_District", "Preference"])
        .size()
        .reset_index(name="Count")
    )

    district_total = (
        district_pref.groupby("Allot_District")["Count"]
        .sum()
        .reset_index(name="Total")
    )

    district_pref = district_pref.merge(district_total, on="Allot_District")
    district_pref["Percentage"] = (
        district_pref["Count"] / district_pref["Total"] * 100
    ).round(2)

    st.subheader("📍 District-wise Preference Satisfaction")
    st.dataframe(district_pref, use_container_width=True)

    # =========================================================
    # REPORT 3: VENUE-WISE SUMMARY
    # =========================================================
    # =========================================================
    # REPORT 3: VENUE-WISE SUMMARY (WITH NAME + DISTRICT)
    # =========================================================
    venue_summary = (
        pd.DataFrame(labs)
        .groupby(
            ["Venue", "Centre", "District"],  # 👈 IMPORTANT
            as_index=False
        )
        .agg(
            Strength=("Strength", "sum"),
            Remaining=("Remaining", "sum")
        )
    )
    
    venue_summary["Allotted"] = (
        venue_summary["Strength"] - venue_summary["Remaining"]
    )
    
    # Reorder columns (clean output)
    venue_summary = venue_summary[
        ["Venue", "Centre", "District", "Strength", "Allotted", "Remaining"]
    ]
    
    st.subheader("🏫 Venue-wise Capacity Summary")
    st.dataframe(venue_summary, use_container_width=True)


    # =========================================================
    # REPORT 4: NOT ALLOTTED
    # =========================================================
    not_allotted = df_cand[df_cand["Allot_Pref"].isna()]
    st.subheader("❌ Not Allotted Candidates")
    st.dataframe(not_allotted[[appl_col, "RollNo"]], use_container_width=True)

    # =========================================================
    # REPORT 5: VENUE-WISE ATTENDANCE SHEETS
    # =========================================================
    attendance_sheets = {}
    for venue, g in df_cand[df_cand["Allot_Venue"].notna()].groupby("Allot_Venue"):
        cols = ["RollNo", appl_col, "Allot_Lab"]
        if name_col != "None":
            cols.insert(2, name_col)
        sheet = g[cols].copy()
        sheet["Signature"] = ""
        attendance_sheets[f"Venue_{venue}"] = sheet

    # =========================================================
    # EXPORT EXCEL
    # =========================================================
    excel_out = "allotment_with_reports.xlsx"
    with pd.ExcelWriter(excel_out, engine="openpyxl") as writer:
        df_cand.to_excel(writer, sheet_name="Allotment", index=False)
        pref_report.to_excel(writer, sheet_name="Preference_Overall", index=False)
        district_pref.to_excel(writer, sheet_name="Preference_District", index=False)
        venue_summary.to_excel(writer, sheet_name="Venue_Summary", index=False)
        not_allotted.to_excel(writer, sheet_name="Not_Allotted", index=False)
        for name, sheet in attendance_sheets.items():
            sheet.to_excel(writer, sheet_name=name[:31], index=False)

    # =========================================================
    # AUTO SUMMARY PDF (ERROR-PROOF)
    # =========================================================
    pdf_file = "allotment_summary.pdf"
    c = canvas.Canvas(pdf_file, pagesize=A4)
    width, height = A4
    y = height - 40

    c.setFont("Helvetica-Bold", 14)
    c.drawString(40, y, "Exam Allotment Summary")

    y -= 25
    c.setFont("Helvetica", 10)
    c.drawString(40, y, f"Total Candidates: {len(df_cand)}")
    y -= 15
    c.drawString(40, y, f"Allotted: {df_cand['Allot_Pref'].notna().sum()}")
    y -= 15
    c.drawString(40, y, f"Not Allotted: {df_cand['Allot_Pref'].isna().sum()}")

    y -= 25
    c.setFont("Helvetica-Bold", 12)
    c.drawString(40, y, "Preference Satisfaction")

    y -= 15
    c.setFont("Helvetica", 10)

    # SAFE INDEX-BASED ACCESS (NO KEYERROR EVER)
    for _, r in pref_report.iterrows():
        c.drawString(
            40, y,
            f"{r.iloc[0]}: {int(r.iloc[1])} ({r.iloc[2]}%)"
        )
        y -= 12

    y -= 20
    c.setFont("Helvetica-Bold", 12)
    c.drawString(40, y, "Venue-wise Summary")

    y -= 15
    c.setFont("Helvetica", 9)
    for _, r in venue_summary.iterrows():
        c.drawString(
            40, y,
            f"Venue {r['Venue']} | Strength: {int(r['Strength'])} | "
            f"Allotted: {int(r['Allotted'])} | Remaining: {int(r['Remaining'])}"
        )
        y -= 12
        if y < 50:
            c.showPage()
            y = height - 40

    c.save()

    # =========================================================
    # DOWNLOADS
    # =========================================================
    st.success("✅ Allotment, Reports & PDF Generated")

    with open(excel_out, "rb") as f:
        st.download_button(
            "⬇️ Download Excel (All Reports)",
            f,
            file_name=excel_out,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    with open(pdf_file, "rb") as f:
        st.download_button(
            "⬇️ Download Auto Summary PDF",
            f,
            file_name=pdf_file,
            mime="application/pdf"
        )
