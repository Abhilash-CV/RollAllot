import streamlit as st
import pandas as pd

st.set_page_config(
    page_title="Roll & Lab Allotment",
    layout="wide"
)

st.title("🎓 Roll Number + Venue & Lab Allotment System")

# =========================
# FILE UPLOADS
# =========================
cand_file = st.file_uploader(
    "Upload Candidate Excel File",
    type=["xlsx"],
    key="cand"
)

lab_file = st.file_uploader(
    "Upload Venue / Lab Master Excel",
    type=["xlsx"],
    key="lab"
)

if cand_file and lab_file:

    # =========================
    # READ EXCEL FILES
    # =========================
    df_cand = pd.read_excel(cand_file, engine="openpyxl")
    df_lab = pd.read_excel(lab_file, engine="openpyxl")

    # =========================
    # CLEAN COLUMN NAMES
    # =========================
    df_cand.columns = df_cand.columns.str.strip()
    df_lab.columns = df_lab.columns.str.strip()

    # =========================
    # PREVIEW DATA
    # =========================
    st.subheader("📄 Candidate Data Preview")
    st.dataframe(df_cand.head(), use_container_width=True)

    st.subheader("🏫 Venue & Lab Master Preview")
    st.dataframe(df_lab, use_container_width=True)

    # =========================
    # CONFIGURATION
    # =========================
    st.subheader("⚙️ Configuration")

    col1, col2, col3 = st.columns(3)

    with col1:
        appl_col = st.selectbox(
            "Application Number Column",
            df_cand.columns,
            index=df_cand.columns.get_loc("ApplNo")
        )

    with col2:
        date_col = st.selectbox(
            "Final Submission Date Column",
            df_cand.columns,
            index=df_cand.columns.get_loc("FSubDate")
        )

    with col3:
        roll_start = st.number_input(
            "Roll Number Start From",
            min_value=1,
            value=7100001,
            step=1
        )

    st.markdown("### 🏫 Venue / Lab Mapping")

    lab_col1, lab_col2, lab_col3, lab_col4 = st.columns(4)

    with lab_col1:
        venue_col = st.selectbox("Venue No Column", df_lab.columns)

    with lab_col2:
        centre_col = st.selectbox("Centre Name Column", df_lab.columns)

    with lab_col3:
        labname_col = st.selectbox("Lab Name Column", df_lab.columns)

    with lab_col4:
        strength_col = st.selectbox("Strength / Capacity Column", df_lab.columns)

    district_col = st.selectbox(
        "District Column",
        df_lab.columns
    )

    # =========================
    # GENERATE ALLOTMENT
    # =========================
    if st.button("🚀 Generate Roll & Lab Allotment"):

        # -------------------------
        # SORT CANDIDATES
        # -------------------------
        df_cand[date_col] = pd.to_datetime(
            df_cand[date_col],
            errors="coerce"
        )

        df_cand = df_cand.sort_values(
            by=[appl_col, date_col],
            ascending=[True, True]
        ).reset_index(drop=True)

        # -------------------------
        # ASSIGN ROLL NUMBERS
        # -------------------------
        df_cand["RollNo"] = range(
            roll_start,
            roll_start + len(df_cand)
        )

        # -------------------------
        # EXPAND LAB CAPACITY
        # -------------------------
        lab_rows = []

        for _, row in df_lab.iterrows():
            cap = pd.to_numeric(
                row[strength_col],
                errors="coerce"
            )

            if pd.isna(cap) or cap <= 0:
                continue

            for _ in range(int(cap)):
                lab_rows.append({
                    "Venue No": row[venue_col],
                    "Centre Name": row[centre_col],
                    "Lab name": row[labname_col],
                    "District": row[district_col]
                })

        df_lab_expanded = pd.DataFrame(lab_rows)

        # -------------------------
        # CAPACITY VALIDATION
        # -------------------------
        if len(df_cand) > len(df_lab_expanded):
            st.error(
                f"❌ Capacity insufficient! "
                f"Candidates: {len(df_cand)}, "
                f"Available seats: {len(df_lab_expanded)}"
            )
            st.stop()

        # -------------------------
        # FINAL ALLOTMENT
        # -------------------------
        df_final = pd.concat(
            [
                df_cand.reset_index(drop=True),
                df_lab_expanded.iloc[:len(df_cand)].reset_index(drop=True)
            ],
            axis=1
        )

        # =========================
        # PREVIEW RESULT
        # =========================
        st.subheader("✅ Final Allotment Preview")

        st.dataframe(
            df_final[
                [appl_col, "RollNo", "Venue No", "Lab name"]
            ].head(30),
            use_container_width=True
        )

        # =========================
        # DOWNLOAD
        # =========================
        output_file = "roll_lab_allotted.xlsx"
        df_final.to_excel(output_file, index=False)

        with open(output_file, "rb") as f:
            st.download_button(
                "⬇️ Download Final Allotted Excel",
                data=f,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

else:
    st.info("📂 Please upload both Candidate and Venue/Lab Excel files.")
