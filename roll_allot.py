import streamlit as st
import pandas as pd

st.set_page_config(page_title="Preference Based Lab Allotment", layout="wide")
st.title("🎓 Preference Based Centre / Lab Allotment")

# =====================
# Upload files
# =====================
cand_file = st.file_uploader("Upload Candidate Excel", type=["xlsx"])
lab_file = st.file_uploader("Upload Lab Capacity Excel", type=["xlsx"])

if not cand_file or not lab_file:
    st.info("Upload both files to continue")
    st.stop()

# =====================
# Read data
# =====================
df_cand = pd.read_excel(cand_file, engine="openpyxl")
df_lab = pd.read_excel(lab_file, engine="openpyxl")

# Clean columns
df_cand.columns = df_cand.columns.str.strip()
df_lab.columns = df_lab.columns.str.strip()

# =====================
# Configuration
# =====================
st.subheader("⚙️ Configuration")

appl_col = st.selectbox("ApplNo Column", df_cand.columns, index=df_cand.columns.get_loc("ApplNo"))
date_col = st.selectbox("FSubDate Column", df_cand.columns, index=df_cand.columns.get_loc("FSubDate"))

pref_cols = st.multiselect(
    "Preference Columns (in order)",
    df_cand.columns,
    default=["Center1", "Center2", "Center3"]
)

district_col = st.selectbox("District Column (Lab Excel)", df_lab.columns)
lab_col = st.selectbox("Lab Name Column", df_lab.columns)
strength_col = st.selectbox("Strength Column", df_lab.columns)

roll_start = st.number_input("Roll No Start", value=7100001, step=1)

# =====================
# Process
# =====================
if st.button("🚀 Generate Preference Based Allotment"):

    # Sort candidates
    df_cand[date_col] = pd.to_datetime(df_cand[date_col], errors="coerce")
    df_cand = df_cand.sort_values([appl_col, date_col]).reset_index(drop=True)

    df_cand["RollNo"] = range(roll_start, roll_start + len(df_cand))
    df_cand["AllottedDistrict"] = None
    df_cand["AllottedLab"] = None

    # Build capacity tracker
    lab_capacity = []

    for _, row in df_lab.iterrows():
        cap = pd.to_numeric(row[strength_col], errors="coerce")
        if pd.isna(cap) or cap <= 0:
            continue

        lab_capacity.append({
            "District": row[district_col],
            "Lab": row[lab_col],
            "Remaining": int(cap)
        })

    # =====================
    # CORE ALLOTMENT LOGIC
    # =====================
    for idx, cand in df_cand.iterrows():

        for pref in pref_cols:
            pref_district = cand[pref]

            if pd.isna(pref_district):
                continue

            for lab in lab_capacity:
                if (
                    lab["District"] == pref_district
                    and lab["Remaining"] > 0
                ):
                    # Allot
                    df_cand.at[idx, "AllottedDistrict"] = lab["District"]
                    df_cand.at[idx, "AllottedLab"] = lab["Lab"]
                    lab["Remaining"] -= 1
                    break
            if df_cand.at[idx, "AllottedDistrict"] is not None:
                break

    # =====================
    # Result
    # =====================
    st.subheader("✅ Allotment Preview")
    st.dataframe(
        df_cand[
            [appl_col, "RollNo", "AllottedDistrict", "AllottedLab"]
        ].head(40),
        use_container_width=True
    )

    # =====================
    # Download
    # =====================
    output = "preference_based_allotment.xlsx"
    df_cand.to_excel(output, index=False)

    with open(output, "rb") as f:
        st.download_button(
            "⬇️ Download Allotment Result",
            data=f,
            file_name=output,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
