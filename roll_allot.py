import streamlit as st
import pandas as pd

st.set_page_config(page_title="Roll Number Allotment", layout="wide")

st.title("🎓 Roll Number Allotment System")

# Upload file
uploaded_file = st.file_uploader(
    "Upload Application Excel File",
    type=["xlsx"]
)

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    st.subheader("📄 Original Data Preview")
    st.dataframe(df.head(10), use_container_width=True)

    # Column selection
    st.subheader("⚙️ Configuration")

    appl_col = st.selectbox("Select Application Number Column", df.columns, index=df.columns.get_loc("ApplNo"))
    date_col = st.selectbox("Select Final Submission Date Column", df.columns, index=df.columns.get_loc("FSubDate"))

    roll_start = st.number_input(
        "Roll Number Starting From",
        min_value=1,
        value=7100001,
        step=1
    )

    if st.button("🚀 Generate Roll Numbers"):
        # Convert date safely
        df[date_col] = pd.to_datetime(df[date_col], errors="coerce")

        # Sort
        df_sorted = df.sort_values(
            by=[appl_col, date_col],
            ascending=[True, True]
        ).reset_index(drop=True)

        # Generate RollNo
        df_sorted["RollNo"] = range(roll_start, roll_start + len(df_sorted))

        st.subheader("✅ Roll Allotted Preview")
        st.dataframe(
            df_sorted[[appl_col, "RollNo", date_col]].head(20),
            use_container_width=True
        )

        # Download
        output_file = "roll_allotted.xlsx"
        df_sorted.to_excel(output_file, index=False)

        with open(output_file, "rb") as f:
            st.download_button(
                "⬇️ Download Roll Allotted File",
                data=f,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
