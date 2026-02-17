import streamlit as st
import pandas as pd

st.set_page_config(page_title="Roll & Lab Allotment", layout="wide")
st.title("🎓 Roll + Exact Venue & Lab Allotment")

# Upload files
cand_file = st.file_uploader("Upload Candidate Excel", type=["xlsx"])
lab_file = st.file_uploader("Upload Venue/Lab Excel", type=["xlsx"])

if cand_file and lab_file:

    # Read files
    df_cand = pd.read_excel(cand_file, engine="openpyxl")
    df_lab = pd.read_excel(lab_file, engine="openpyxl")

    # Clean column names
    df_cand.columns = df_cand.columns.str.strip()
    df_lab.columns = df_lab.columns.str.strip()

    st.subheader("Candidate Preview")
    st.dataframe(df_cand.head(), use_container_width=True)

    st.subheader("Lab/Venue Preview (ORDER MATTERS)")
    st.dataframe(df_lab, use_container_width=True)

    # Candidate config
    appl_col = st.selectbox("ApplNo Column", df_cand.columns, index=df_cand.columns.get_loc("ApplNo"))
    date_col = st.selectbox("FSubDate Column", df_cand.columns, index=df_cand.columns.get_loc("FSubDate"))
    roll_start = st.number_input("Roll Start No", value=7100001, step=1)

    # Lab config
    code_col = st.selectbox("Code Column", df_lab.columns)
    venue_col = st.selectbox("Venue No Column", df_lab.columns)
    centre_col = st.selectbox("Centre Name Column", df_lab.columns)
    lab_col = st.selectbox("Lab Name Column", df_lab.columns)
    strength_col = st.selectbox("Strength Column", df_lab.columns)
    district_col = st.selectbox("District Column", df_lab.columns)

    if st.button("🚀 Generate Exact Allotment"):

        # Sort candidates
        df_cand[date_col] = pd.to_datetime(df_cand[date_col], errors="coerce")
        df_cand = df_cand.sort_values([appl_col, date_col]).reset_index(drop=True)

        # Assign roll numbers
        df_cand["RollNo"] = range(roll_start, roll_start + len(df_cand))

        # =============================
        # STRICT LAB SEAT GENERATION
        # =============================
        seat_map = []

        for _, lab in df_lab.iterrows():
            cap = int(pd.to_numeric(lab[strength_col], errors="coerce"))

            for seat in range(1, cap + 1):
                seat_map.append({
                    "Code": lab[code_col],
                    "Venue No": lab[venue_col],
                    "Centre Name": lab[centre_col],
                    "Lab name": lab[lab_col],
                    "District": lab[district_col],
                    "SeatNo": seat
                })

        df_seats = pd.DataFrame(seat_map)

        # Capacity check
        if len(df_cand) > len(df_seats):
            st.error("❌ Candidates exceed total lab capacity")
            st.stop()

        # =============================
        # FINAL ONE-TO-ONE MAPPING
        # =============================
        df_final = pd.concat(
            [
                df_cand.reset_index(drop=True),
                df_seats.iloc[:len(df_cand)].reset_index(drop=True)
            ],
            axis=1
        )

        st.subheader("✅ Correct Allotment Preview")
        st.dataframe(
            df_final[[appl_col, "RollNo", "Code", "Venue No", "Lab name", "SeatNo"]].head(40),
            use_container_width=True
        )

        # Export
        output = "roll_venue_lab_exact.xlsx"
        df_final.to_excel(output, index=False)

        with open(output, "rb") as f:
            st.download_button(
                "⬇️ Download Final Allotment",
                f,
                file_name=output
            )
