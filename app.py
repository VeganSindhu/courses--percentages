import streamlit as st
import pandas as pd
from io import BytesIO

# --------------------------------------------------
# PAGE CONFIG
# --------------------------------------------------
st.set_page_config(page_title="Course Completion Status", layout="wide")
st.title("📘 Course Completion Status")

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


@st.cache_data(show_spinner=False)
def process_excel(file):
    df = pd.read_excel(file)
    df.columns = df.columns.astype(str).str.strip()

    required = {"Employee Name", "Office of Working"}
    if not required.issubset(df.columns):
        raise ValueError("Required columns missing")

    course_cols = [
        c for c in df.columns
        if c not in ["Employee Name", "Office of Working", "Total Courses"]
        and pd.api.types.is_numeric_dtype(df[c])
    ]

    total_courses = len(course_cols)

    df["Completed Courses"] = df[course_cols].sum(axis=1)
    df["Completion %"] = round((df["Completed Courses"] / total_courses) * 100, 2)
    df["Office Group"] = df["Office of Working"].apply(office_group)

    division_pct = round(
        (df["Completed Courses"].sum() / (len(df) * total_courses)) * 100, 2
    )

    return df, division_pct, total_courses


def df_to_excel_bytes(df):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    out.seek(0)
    return out.getvalue()

# --------------------------------------------------
# ADMIN SECTION
# --------------------------------------------------
st.sidebar.header("🔐 Admin Panel")

admin_upload = st.sidebar.file_uploader(
    "Upload / Update Excel (Admin only)",
    type=["xlsx"]
)

if admin_upload:
    try:
        df_data, division_pct, total_courses = process_excel(admin_upload)
        st.session_state["data"] = df_data
        st.session_state["division_pct"] = division_pct
        st.session_state["total_courses"] = total_courses
        st.sidebar.success("✅ Data updated successfully")
    except Exception as e:
        st.sidebar.error(f"Upload failed: {e}")

# --------------------------------------------------
# CHECK IF DATA EXISTS
# --------------------------------------------------
if "data" not in st.session_state:
    st.info("⏳ Data not yet uploaded by Admin.")
    st.stop()

df = st.session_state["data"]
division_pct = st.session_state["division_pct"]
total_courses = st.session_state["total_courses"]

# --------------------------------------------------
# DIVISION SUMMARY
# --------------------------------------------------
st.subheader("📊 Division Completion Status")
st.metric("Division Completion %", f"{division_pct}%")

# --------------------------------------------------
# OFFICE SUMMARY
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
# END USER SEARCH
# --------------------------------------------------
st.subheader("🔍 Check Your Completion Status")

names = sorted(df["Employee Name"].dropna().unique())
query = st.text_input("Type at least 4 characters of your name")

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
# USER VIEW
# --------------------------------------------------
user_row = df[df["Employee Name"] == selected_name]
pct = float(user_row["Completion %"].iloc[0])
color = completion_color(pct)

st.markdown(
    f"<h3 style='color:{color};'>👤 {selected_name} — Completion: {pct}%</h3>",
    unsafe_allow_html=True
)

st.dataframe(user_row)

st.download_button(
    "📥 Download My Report",
    data=df_to_excel_bytes(user_row),
    file_name=f"{selected_name}_course_report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
