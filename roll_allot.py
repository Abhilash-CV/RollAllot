import streamlit as st
import pandas as pd

st.set_page_config(page_title="Roll & Lab Allotment", layout="wide")
st.title("🎓 Roll Number + Venue & Lab Allotment")

# Upload files
cand_file = st.file_uploader("Upload Candidate Excel", type=["xlsx"])
lab_file = st.file_uploader("Upload Venue / Lab Master Excel", type=["xlsx"])

if cand_file and lab_file:
    df_cand = pd.read_excel(cand_file, engine="openpyxl")
    df_lab = pd.read_excel(lab_file, engine="openpyxl")

    st.subheader("📄 Candidate Preview")
    st.dataframe(df_cand.head(), use_container_width=True)

    st.subheader("🏫 Venue & Lab Preview")
    st.dataframe(df_lab, use_container_width=True)

    st.subheader("⚙️ Configuration")

    appl_col = st.selectbox("Application No Column", df_cand.columns, index=df_cand.columns.get_loc("ApplNo"))
    date_col = st.selectbox("Submission Date Column", df_cand.columns, index=df_cand.columns.get_loc("FSubDate"))

    roll_start = st.number_input("Roll Number Start From", value=7100001, step=1)

    if st.button("🚀 Generate Roll + Lab Allotment"):

        # Sort candidates
        df_cand[date_col] = pd.to_datetime(df_cand[date_col], errors="coerce")
        df_cand = df_cand.sort_values(
            by=[appl_col, date_col],
            ascending=[True, True]
        ).reset_index(drop=True)

        # Generate Roll Numbers
        df_cand["RollNo"] = range(roll_start, roll_start + len(df_cand))

        # Prepare lab capacity list
        lab_rows = []
        for _, row in df_lab.iterrows():
            for _ in range(int(row["Strength"])):
                lab_rows.append({
                    "Venue No": row["Venue No."],
                    "Centre Name": row["Centre Name"],
                    "Lab name": row["Lab name"],
                    "District": row["District"]
                })

        df_lab_expanded = pd.DataFrame(lab_rows)

        # Capacity validation
        if len(df_cand) > len(df_lab_expanded):
            st.error("❌ Total candidates exceed available lab capacity!")
            st.stop()

        # Assign labs sequentially
        df_final = pd.concat(
            [df_cand.reset_index(drop=True),
             df_lab_expanded.iloc[:len(df_cand)].reset_index(drop=True)],
            axis=1
        )

        st.subheader("✅ Final Allotment Preview")
        st.dataframe(
            df_final[[appl_col, "RollNo", "Venue No", "Lab name"]].head(25),
            use_container_width=True
        )

        # Export
        output_file = "roll_lab_allotted.xlsx"
        df_final.to_excel(output_file, index=False)

        with open(output_file, "rb") as f:
            st.download_button(
                "⬇️ Download Final Allotted Excel",
                f,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
