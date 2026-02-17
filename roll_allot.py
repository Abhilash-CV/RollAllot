import streamlit as st
import pandas as pd

st.set_page_config(page_title="Roll & Lab Allotment", layout="wide")
st.title("🎓 Roll Number + Exact Venue & Lab Allotment")

# =========================
# FILE UPLOAD
# =========================
cand_file = st.file_uploader("Upload Candidate Excel", type=["xlsx"])
lab_file = st.file_uploader("Upload Venue / Lab Excel", type=["xlsx"])

if not cand_file or not lab_file:
    st.info("📂 Please upload BOTH Candidate and Venue/Lab Excel files.")
    st.stop()

# =========================
# READ FILES
# =========================
df_cand = pd.read_excel(cand_file, engine="openpyxl")
df_lab = pd.read_excel(lab_file, engine="openpyxl")

# =========================
# CLEAN COLUMN NAMES
# =========================
df_cand.columns = df_cand.columns.str.strip()
df_lab.columns = df_lab.columns.str.strip()

# =========================
# PREVIEW
# =========================
st.subheader("📄 Candidate Data Preview")
st.dataframe(df_cand.head(), use_container_width=True)

st.subheader("🏫 Venue / Lab Master Preview (ROW ORDER IS FINAL)")
st.dataframe(df_lab, use_container_width=True)

# =========================
# CONFIGURATION
# =========================
st.subheader("⚙️ Configuration")

c1, c2, c3 = st.columns(3)

with c1:
    appl_col = st.selectbox(
        "Application No Column",
        df_cand.columns,
        index=df_cand.columns.get_loc("ApplNo")
    )

with c2:
    date_col = st.selectbox(
        "Final Submission Date Column",
        df_cand.columns,
        index=df_cand.columns.get_loc("FSubDate")
    )

with c3:
    roll_start = st.number_input(
        "Roll Number Start From",
        value=7100001,
        step=1
    )

st.markdown("### 🏫 Venue / Lab Columns")

l1, l2, l3 = st.columns(3)
l4, l5, l6 = st.columns(3)

with l1:
    code_col = st.selectbox("Code Column", df_lab.columns)
with l2:
    venue_col = st.selectbox("Venue No Column", df_lab.columns)
with l3:
    centre_col = st.selectbox("Centre Name Column", df_lab.columns)

with l4:
    lab_col = st.selectbox("Lab Name Column", df_lab.columns)
with l5:
    strength_col = st.selectbox("Strength / Capacity Column", df_lab.columns)
with l6:
    district_col = st.selectbox("District Column", df_lab.columns)

# =========================
# PROCESS
# =========================
if st.button("🚀 Generate Roll & Exact Lab Allotment"):

    # -------------------------
    # SORT CANDIDATES
    # -------------------------
    df_cand[date_col] = pd.to_datetime(df_cand[date_col], errors="coerce")

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
    # GENERATE SEAT MAP
    # (STRICT ROW ORDER)
    # -------------------------
    seat_rows = []

    for _, lab in df_lab.iterrows():

        cap = pd.to_numeric(lab[strength_col], errors="coerce")

        if pd.isna(cap) or cap <= 0:
            continue

        cap = int(cap)

        for seat_no in range(1, cap + 1):
            seat_rows.append({
                "Code": lab[code_col],
                "Venue No": lab[venue_col],
                "Centre Name": lab[centre_col],
                "Lab name": lab[lab_col],
                "District": lab[district_col],
                "SeatNo": seat_no
            })

    df_seats = pd.DataFrame(seat_rows)

    # -------------------------
    # CAPACITY CHECK
    # -------------------------
    if len(df_cand) > len(df_seats):
        st.error(
            f"❌ Capacity insufficient!\n\n"
            f"Candidates: {len(df_cand)}\n"
            f"Available seats: {len(df_seats)}"
        )
        st.stop()

    # -------------------------
    # FINAL ALLOTMENT
    # -------------------------
    df_final = pd.concat(
        [
            df_cand.reset_index(drop=True),
            df_seats.iloc[:len(df_cand)].reset_index(drop=True)
        ],
        axis=1
    )

    # =========================
    # PREVIEW RESULT
    # =========================
    st.subheader("✅ Final Allotment Preview")

    st.dataframe(
        df_final[
            [appl_col, "RollNo", "Code", "Venue No", "Lab name", "SeatNo"]
        ].head(40),
        use_container_width=True
    )

    # =========================
    # DOWNLOAD
    # =========================
    output_file = "roll_venue_lab_exact.xlsx"
    df_final.to_excel(output_file, index=False)

    with open(output_file, "rb") as f:
        st.download_button(
            "⬇️ Download Final Allotted Excel",
            data=f,
            file_name=output_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
