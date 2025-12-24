import streamlit as st
import pandas as pd
import os
import re
from datetime import datetime

# --------------------------------------------------
# AUTO-DETECT DATA FILE & DATE
# --------------------------------------------------
def get_data_file_and_date():
    excel_files = []

    for f in os.listdir("."):
        if f.lower().endswith(".xlsx"):
            m = re.search(r"(\d{2}\.\d{2}\.(\d{2}|\d{4}))", f)
            if m:
                date_str = m.group(1)
                excel_files.append((f, date_str))

    if not excel_files:
        raise FileNotFoundError(
            "No Excel file with date dd.mm.yy or dd.mm.yyyy found"
        )

    def parse_date(d):
        try:
            return datetime.strptime(d, "%d.%m.%y")
        except ValueError:
            return datetime.strptime(d, "%d.%m.%Y")

    excel_files.sort(key=lambda x: parse_date(x[1]), reverse=True)

    file_name, date_part = excel_files[0]
    as_on = parse_date(date_part).strftime("%d.%m.%Y")

    return file_name, as_on


# --------------------------------------------------
# INIT
# --------------------------------------------------
DATA_FILE, AS_ON_DATE = get_data_file_and_date()

st.set_page_config(page_title="Course Completion Status", layout="wide")
st.title(f"📘 Course Completion Status as on {AS_ON_DATE}")
st.caption(f"📂 Data source: {DATA_FILE}")

# --------------------------------------------------
# HELPERS
# --------------------------------------------------
def unit_name(office):
    nellore = ["SRO Nellore", "MBC Nellore RMS", "RO Chennai", "Gudur TMO"]
    return "Nellore Unit" if office in nellore else "Tirupati Unit"


# --------------------------------------------------
# LOAD DATA
# --------------------------------------------------
@st.cache_data
def load_data(data_file):
    df = pd.read_excel(data_file)
    df.columns = df.columns.astype(str).str.strip()

    ignore = {"Employee Name", "Office of Working", "Total Courses"}
    course_cols = [
        c for c in df.columns
        if c not in ignore and pd.api.types.is_numeric_dtype(df[c])
    ]

    # Force strict 0 / 1
    df[course_cols] = df[course_cols].fillna(0).astype(int)

    df["Unit"] = df["Office of Working"].apply(unit_name)

    num_employees = df.shape[0]
    num_courses = len(course_cols)

    # Employee-level
    df["Pending Courses"] = df[course_cols].sum(axis=1)
    df["Completed Courses"] = num_courses - df["Pending Courses"]

    return df, course_cols, num_employees, num_courses


df, course_cols, num_employees, num_courses = load_data(DATA_FILE)

# --------------------------------------------------
# 📊 DIVISION SUMMARY (COURSES-BASED)
# --------------------------------------------------
st.subheader("📊 Division-level Course Status")

total_courses_div = num_employees * num_courses
pending_courses_div = df[course_cols].sum().sum()
completed_courses_div = total_courses_div - pending_courses_div
division_pct = round(
    (completed_courses_div / total_courses_div) * 100, 2
)

c1, c2, c3, c4, c5, c6 = st.columns(6)

c1.metric("Employees", num_employees)
c2.metric("Courses / Employee", num_courses)
c3.metric("Total Courses", total_courses_div)
c4.metric("Pending Courses", pending_courses_div)
c5.metric("Completed Courses", completed_courses_div)
c6.metric("Completion %", f"{division_pct}%")

st.divider()

# --------------------------------------------------
# 🏢 UNIT-WISE SUMMARY (COURSES-BASED)
# --------------------------------------------------
st.subheader("🏢 Unit-wise Course Status")

unit_rows = []

for unit, g in df.groupby("Unit"):
    emp = g.shape[0]
    total = emp * num_courses
    pending = g[course_cols].sum().sum()
    completed = total - pending
    pct = round((completed / total) * 100, 2)

    unit_rows.append({
        "Unit": unit,
        "Employees": emp,
        "Total Courses": total,
        "Pending Courses": pending,
        "Completed Courses": completed,
        "Completion %": pct
    })

unit_df = pd.DataFrame(unit_rows)
st.dataframe(unit_df, use_container_width=True)

st.divider()

# --------------------------------------------------
# 🔍 SEARCH EMPLOYEE
# --------------------------------------------------
st.subheader("🔍 Check Your Completion Status")

if "search_key" not in st.session_state:
    st.session_state["search_key"] = 0

col1, col2 = st.columns([5, 1])

with col1:
    query = st.text_input(
        "Start typing your name",
        key=f"name_query_{st.session_state['search_key']}"
    )

with col2:
    st.write("")
    if st.button("❌ Clear"):
        st.session_state["search_key"] += 1
        st.stop()

query = query.strip()
if not query:
    st.stop()

matches = df[df["Employee Name"].str.contains(query, case=False, na=False)]

if matches.empty:
    st.info("No matching names found")
    st.stop()

st.dataframe(
    matches[
        ["Employee Name", "Office of Working", "Unit",
         "Pending Courses", "Completed Courses"]
    ],
    use_container_width=True
)

st.divider()

# --------------------------------------------------
# 🚨 ZERO COMPLETION EMPLOYEES
# --------------------------------------------------
st.subheader("🚨 Employees who have NOT completed even ONE course")

zero_completed = df[df["Completed Courses"] == 0]

if zero_completed.empty:
    st.success("🎉 All employees have completed at least one course")
else:
    st.error(
        f"⚠️ {len(zero_completed)} employees have completed ZERO courses"
    )
    st.dataframe(
        zero_completed[
            ["Employee Name", "Office of Working", "Unit",
             "Pending Courses"]
        ],
        use_container_width=True
    )
