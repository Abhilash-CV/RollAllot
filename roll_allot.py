import streamlit as st
import pandas as pd

st.set_page_config(page_title="Roll Allotment (Two Excel)", layout="wide")
st.title("🎓 Roll Number Allotment – Two Excel Inputs")

# Upload files
col1, col2 = st.columns(2)

with col1:
    cand_file = st.file_uploader(
        "Upload Candidate Excel",
        type=["xlsx"],
        key="cand"
    )

with col2:
    center_file = st.file_uploader(
        "Upload Center Details Excel",
        type=["xlsx"],
        key="center"
    )

if cand_file and center_file:
    df_cand = pd.read_excel(cand_file)
    df_center = pd.read_excel(center_file)

    st.subheader("📄 Candidate Data Preview")
    st.dataframe(df_cand.head(), use_container_width=True)

    st.subheader("🏫 Center Data Preview")
    st.dataframe(df_center.head(), use_container_width=True)

    st.subheader("⚙️ Configuration")

    # Column selections
    appl_col = st.selectbox("Application Number Column", df_cand.columns, index=df_cand.columns.get_loc("ApplNo"))
    date_col = st.selectbox("Final Submission Date Column", df_cand.columns, index=df_cand.columns.get_loc("FSubDate"))

    cand_center_col = st.selectbox(
        "Candidate Center Column",
        df_cand.columns,
        index=df_cand.columns.get_loc("Center1")
    )

    center_key_col = st.selectbox(
        "Center Master Key Column",
        df_center.columns
    )

    roll_start = st.number_input(
        "Roll Number Start From",
        min_value=1,
        value=7100001,
        step=1
    )

    if st.button("🚀 Generate Roll Numbers"):
        # Date conversion
        df_cand[date_col] = pd.to_datetime(df_cand[date_col], errors="coerce")

        # Sort
        df_cand = df_cand.sort_values(
            by=[appl_col, date_col],
            ascending=[True, True]
        ).reset_index(drop=True)

        # Roll generation
        df_cand["RollNo"] = range(roll_start, roll_start + len(df_cand))

        # Merge center details
        df_final = df_cand.merge(
            df_center,
            how="left",
            left_on=cand_center_col,
            right_on=center_key_col
        )

        st.subheader("✅ Final Roll Allotted Data")
        st.dataframe(
            df_final[[appl_col, "RollNo", cand_center_col]].head(20),
            use_container_width=True
        )

        # Download
        output_file = "roll_allotted_with_centers.xlsx"
        df_final.to_excel(output_file, index=False)

        with open(output_file, "rb") as f:
            st.download_button(
                "⬇️ Download Final Excel",
                data=f,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
