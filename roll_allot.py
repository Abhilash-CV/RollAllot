import streamlit as st
import pandas as pd

st.set_page_config(page_title="Preference Based Centre + Lab Allotment", layout="wide")
st.title("🎓 Preference Based Centre + Lab Allotment")

# =========================
# Upload files
# =========================
cand_file = st.file_uploader("Upload Candidate Excel", type=["xlsx"])
lab_file = st.file_uploader("Upload Venue / Lab Excel", type=["xlsx"])

if not cand_file or not lab_file:
    st.info("Please upload BOTH Candidate and Venue/Lab Excel files.")
    st.stop()

# =========================
# Read Excel
# =========================
df_cand = pd.read_excel(cand_file, engine="openpyxl")
df_lab = pd.read_excel(lab_file, engine="openpyxl")

# Clean column names
df_cand.columns = df_cand.columns.str.strip()
df_lab.columns = df_lab.columns.str.strip()

# =========================
# Configuration
# =========================
st.subheader("⚙️ Configuration")

appl_col = st.selectbox("Application No Column", df_cand.columns, index=df_cand.columns.get_loc("ApplNo"))
date_col = st.selectbox("Final Submission Date Column", df_cand.columns, index=df_cand.columns.get_loc("FSubDate"))

pref_cols = st.multiselect(
    "Preference Columns (District order)",
    df_cand.columns,
    default=["Center1", "Center2", "Center3"]
)

st.markdown("### 🏫 Venue / Lab Master Columns")

code_col = st.selectbox("Code Column", df_lab.columns)
venue_col = st.selectbox("Venue No Column", df_lab.columns)
centre_col = st.selectbox("Centre Name Column", df_lab.columns)
district_col = st.selectbox("District Column", df_lab.columns)
lab_col = st.selectbox("Lab Name Column", df_lab.columns)
strength_col = st.selectbox("Strength Column", df_lab.columns)

roll_start = st.number_input("Roll No Start", value=7100001, step=1)

# =========================
# Generate Allotment
# =========================
if st.button("🚀 Generate Preference Based Allotment"):

    # Sort candidates
    df_cand[date_col] = pd.to_datetime(df_cand[date_col], errors="coerce")
    df_cand = df_cand.sort_values([appl_col, date_col]).reset_index(drop=True)

    # Roll numbers
    df_cand["RollNo"] = range(roll_start, roll_start + len(df_cand))

    # Output columns
    df_cand["Allot_Code"] = None
    df_cand["Allot_VenueNo"] = None
    df_cand["Allot_CentreName"] = None
    df_cand["Allot_District"] = None
    df_cand["Allot_Lab"] = None

    # =========================
    # Build lab capacity tracker
    # (ROW ORDER PRESERVED)
    # =========================
    labs = []

    for _, row in df_lab.iterrows():
        cap = pd.to_numeric(row[strength_col], errors="coerce")
        if pd.isna(cap) or cap <= 0:
            continue

        labs.append({
            "Code": row[code_col],
            "VenueNo": row[venue_col],
            "CentreName": row[centre_col],
            "District": row[district_col],
            "Lab": row[lab_col],
            "Remaining": int(cap)
        })

    # =========================
    # CORE ALLOTMENT LOOP
    # =========================
    for i, cand in df_cand.iterrows():

        for pref in pref_cols:
            pref_dist = cand[pref]

            if pd.isna(pref_dist):
                continue

            for lab in labs:
                if lab["District"] == pref_dist and lab["Remaining"] > 0:

                    df_cand.at[i, "Allot_Code"] = lab["Code"]
                    df_cand.at[i, "Allot_VenueNo"] = lab["VenueNo"]
                    df_cand.at[i, "Allot_CentreName"] = lab["CentreName"]
                    df_cand.at[i, "Allot_District"] = lab["District"]
                    df_cand.at[i, "Allot_Lab"] = lab["Lab"]

                    lab["Remaining"] -= 1
                    break

            if df_cand.at[i, "Allot_District"] is not None:
                break

    # =========================
    # Preview
    # =========================
    st.subheader("✅ Allotment Preview")

    st.dataframe(
        df_cand[
            [
                appl_col,
                "RollNo",
                "Allot_Code",
                "Allot_VenueNo",
                "Allot_CentreName",
                "Allot_District",
                "Allot_Lab",
            ]
        ].head(40),
        use_container_width=True
    )

    # =========================
    # Download
    # =========================
    output_file = "preference_based_centre_lab_allotment.xlsx"
    df_cand.to_excel(output_file, index=False)

    with open(output_file, "rb") as f:
        st.download_button(
            "⬇️ Download Final Allotment Excel",
            data=f,
            file_name=output_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
