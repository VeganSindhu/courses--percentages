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


def unit_name(office):
    nellore = [
        "SRO Nellore",
        "MBC Nellore RMS",
        "RO Chennai",
        "Gudur TMO"
    ]
    return "Nellore Unit" if office in nellore else "Tirupati Unit"


@st.cache_data(show_spinner=False)
def process_excel(file):
    df = pd.read_excel(file)
    df.columns = df.columns.astype(str).str.strip()

    # Required columns
    if "Employee Name" not in df.columns or "Office of Working" not in df.columns:
        raise ValueError("Required columns missing")

    # Course columns = numeric, excluding known columns
    ignore = {"Employee Name", "Office of Working", "Total Courses"}
    course_cols = [
        c for c in df.columns
        if c not in ignore and pd.api.types.is_numeric_dtype(df[c])
    ]

    if not course_cols:
        raise ValueError("No course columns detected")

    total_courses = len(course_cols)

    # 🔥 LOGIC: 1 = pending, 0 / blank = completed
    df[course_cols] = df[course_cols].fillna(0)
    df["Pending Courses"] = df[course_cols].sum(axis=1)
    df["Completed Courses"] = total_courses - df["Pending Courses"]

    df["Completion %"] = round(
        (df["Completed Courses"] / total_courses) * 100, 2
    )

    df["Unit"] = df["Office of Working"].apply(unit_name)

    division_pct = round(
        (df["Completed Courses"].sum() / (len(df) * total_courses)) * 100, 2
    )

    return df, course_cols, total_courses, division_pct


def df_to_excel_bytes(df):
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    out.seek(0)
    return out.getvalue()

# --------------------------------------------------
# ADMIN PANEL
# --------------------------------------------------
st.sidebar.header("🔐 Admin Panel")

admin_file = st.sidebar.file_uploader(
    "Upload / Update Excel (Admin only)",
    type=["xlsx"]
)

if admin_file:
    try:
        df_data, course_cols, total_courses, division_pct = process_excel(admin_file)
        st.session_state["df"] = df_data
        st.session_state["courses"] = course_cols
        st.session_state["total_courses"] = total_courses
        st.session_state["division_pct"] = division_pct
        st.sidebar.success("✅ Data updated successfully")
    except Exception as e:
        st.sidebar.error(str(e))

# --------------------------------------------------
# CHECK DATA AVAILABILITY
# --------------------------------------------------
if "df" not in st.session_state:
    st.info("⏳ Data not yet uploaded by Admin.")
    st.stop()

df = st.session_state["df"]
course_cols = st.session_state["courses"]
total_courses = st.session_state["total_courses"]
division_pct = st.session_state["division_pct"]

# --------------------------------------------------
# DIVISION SUMMARY
# --------------------------------------------------
st.subheader("📊 Division Completion Status")
st.metric("Division Completion %", f"{division_pct}%")

# --------------------------------------------------
# UNIT SUMMARY
# --------------------------------------------------
unit_summary = (
    df.groupby("Unit")
    .apply(lambda x: round(
        (x["Completed Courses"].sum() / (len(x) * total_courses)) * 100, 2
    ))
    .reset_index(name="Completion %")
)

st.subheader("🏢 Unit-wise Completion %")
st.dataframe(unit_summary)

st.divider()

# --------------------------------------------------
# LIVE SEARCH (NO ENTER)
# --------------------------------------------------
st.subheader("🔍 Check Your Completion Status")

query = st.text_input("Start typing your name")

matches = []
if query:
    matches = [
        n for n in df["Employee Name"].unique()
        if query.lower() in n.lower()
    ]

selected_name = None
if matches:
    selected_name = st.selectbox("Matching names", matches)

if not selected_name:
    st.stop()

# --------------------------------------------------
# USER REPORT
# --------------------------------------------------
user_row = df[df["Employee Name"] == selected_name].copy()
pct = float(user_row["Completion %"].iloc[0])
color = completion_color(pct)

st.markdown(
    f"<h3 style='color:{color};'>👤 {selected_name} — Completion: {pct}%</h3>",
    unsafe_allow_html=True
)

# Show course-wise status
course_view = user_row[course_cols].T.reset_index()
course_view.columns = ["Course Name", "Pending (1 = Pending)"]

st.subheader("📘 Course-wise Status")
st.dataframe(course_view)

st.subheader("📄 Summary")
st.dataframe(user_row[[
    "Employee Name",
    "Office of Working",
    "Unit",
    "Completed Courses",
    "Pending Courses",
    "Completion %"
]])

# --------------------------------------------------
# DOWNLOAD
# --------------------------------------------------
st.download_button(
    "📥 Download My Report",
    data=df_to_excel_bytes(user_row),
    file_name=f"{selected_name}_course_report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
