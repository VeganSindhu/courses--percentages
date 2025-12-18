import streamlit as st
import pandas as pd
from io import BytesIO

# --------------------------------------------------
# PAGE CONFIG
# --------------------------------------------------
st.set_page_config(page_title="Course Completion Report", layout="wide")
st.title("📘 Course Completion Report")

uploaded_file = st.file_uploader(
    "Upload Excel course completion file",
    type=["xlsx"]
)

if not uploaded_file:
    st.info("Upload the Excel file shown above")
    st.stop()

# --------------------------------------------------
# HELPERS
# --------------------------------------------------
def completion_color(pct):
    if pct < 10:
        return "red"
    elif 10 <= pct <= 50:
        return "orange"
    elif pct >= 90:
        return "green"
    else:
        return "black"


def office_group(office):
    group1 = [
        "SRO Nellore",
        "MBC Nellore RMS",
        "RO Chennai",
        "Gudur TMO"
    ]
    return "Office Group 1" if office in group1 else "Office Group 2"


def df_to_excel_bytes(df):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    out.seek(0)
    return out.getvalue()

# --------------------------------------------------
# READ EXCEL (AS-IS)
# --------------------------------------------------
df = pd.read_excel(uploaded_file)

# Clean column names
df.columns = df.columns.astype(str).str.strip()

# Mandatory columns
required_cols = ["Employee Name", "Office of Working"]
for col in required_cols:
    if col not in df.columns:
        st.error(f"Required column missing: {col}")
        st.stop()

# --------------------------------------------------
# IDENTIFY COURSE COLUMNS
# --------------------------------------------------
course_cols = [
    c for c in df.columns
    if c not in ["Employee Name", "Office of Working", "Total Courses"]
    and pd.api.types.is_numeric_dtype(df[c])
]

if not course_cols:
    st.error("No course columns detected.")
    st.stop()

total_courses = len(course_cols)

# --------------------------------------------------
# CALCULATE COMPLETION %
# --------------------------------------------------
df["Completed Courses"] = df[course_cols].sum(axis=1)
df["Completion %"] = round((df["Completed Courses"] / total_courses) * 100, 2)
df["Office Group"] = df["Office of Working"].apply(office_group)

# --------------------------------------------------
# DIVISION COMPLETION %
# --------------------------------------------------
division_pct = round(
    (df["Completed Courses"].sum() / (len(df) * total_courses)) * 100, 2
)

st.subheader("📊 Division Completion Status")
st.metric("Division Completion %", f"{division_pct}%")

st.divider()

# --------------------------------------------------
# OFFICE-WISE COMPLETION %
# --------------------------------------------------
office_summary = (
    df.groupby("Office Group")
    .apply(lambda x: round((x["Completed Courses"].sum() / (len(x) * total_courses)) * 100, 2))
    .reset_index(name="Completion %")
)

st.subheader("🏢 Office-wise Completion %")
st.dataframe(office_summary)

st.divider()

# --------------------------------------------------
# SEARCH EMPLOYEE
# --------------------------------------------------
st.subheader("🔍 Search Employee (min 4 characters)")

names = sorted(df["Employee Name"].dropna().unique())
query = st.text_input("Type your name")

selected_name = None
if len(query) >= 4:
    matches = [n for n in names if query.lower() in n.lower()]
    if matches:
        selected_name = st.selectbox("Select your name", matches)
    else:
        st.info("No match found")

if not selected_name:
    st.stop()

# --------------------------------------------------
# DISPLAY USER REPORT (COLORED NAME)
# --------------------------------------------------
user_row = df[df["Employee Name"] == selected_name]
pct = float(user_row["Completion %"].iloc[0])
color = completion_color(pct)

st.markdown(
    f"<h3 style='color:{color};'>👤 {selected_name} — Completion: {pct}%</h3>",
    unsafe_allow_html=True
)

st.dataframe(user_row)

# --------------------------------------------------
# DOWNLOAD USER REPORT
# --------------------------------------------------
st.download_button(
    "📥 Download My Course Report (Excel)",
    data=df_to_excel_bytes(user_row),
    file_name=f"{selected_name}_course_report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
