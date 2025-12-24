import streamlit as st
import pandas as pd
import os
import re
from datetime import datetime

# --------------------------------------------------
# CONFIG : SANCTIONED STRENGTH
# --------------------------------------------------
SANCTIONED_STRENGTH = {
    "Division": 183,
    "Nellore Unit": 60,
    "Tirupati Unit": 123
}

# --------------------------------------------------
# AUTO-DETECT DATA FILE & DATE
# --------------------------------------------------
def get_data_file_and_date():
    excel_files = []

    for f in os.listdir("."):
        if f.lower().endswith(".xlsx"):
            m = re.search(r"(\d{2}\.\d{2}\.(\d{2}|\d{4}))", f)
            if m:
                excel_files.append((f, m.group(1)))

    if not excel_files:
        raise FileNotFoundError("No Excel file with date found")

    def parse(d):
        try:
            return datetime.strptime(d, "%d.%m.%y")
        except ValueError:
            return datetime.strptime(d, "%d.%m.%Y")

    excel_files.sort(key=lambda x: parse(x[1]), reverse=True)
    file_name, date_part = excel_files[0]

    return file_name, parse(date_part).strftime("%d.%m.%Y")


DATA_FILE, AS_ON_DATE = get_data_file_and_date()

# --------------------------------------------------
# PAGE CONFIG
# --------------------------------------------------
st.set_page_config(page_title="Course Completion Status", layout="wide")
st.title(f"📘 Course Completion Status as on {AS_ON_DATE}")
st.caption(f"📂 Data source: {DATA_FILE}")

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

    df[course_cols] = df[course_cols].fillna(0).astype(int)
    df["Unit"] = df["Office of Working"].apply(unit_name)

    num_courses = len(course_cols)

    df["Pending Courses"] = df[course_cols].sum(axis=1)
    df["Completed Courses"] = num_courses - df["Pending Courses"]
    df["Completion %"] = round(
        (df["Completed Courses"] / num_courses) * 100, 2
    )

    return df, course_cols, num_courses


df, course_cols, num_courses = load_data(DATA_FILE)

# --------------------------------------------------
# 📊 DIVISION SUMMARY (FIXED BASE)
# --------------------------------------------------
st.subheader("📊 Division-level Course Status")

total_div = SANCTIONED_STRENGTH["Division"] * num_courses
pending_div = df[course_cols].sum().sum()
completed_div = total_div - pending_div
div_pct = round((completed_div / total_div) * 100, 2)

c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("Sanctioned Strength", SANCTIONED_STRENGTH["Division"])
c2.metric("Courses / Employee", num_courses)
c3.metric("Total Courses", total_div)
c4.metric("Pending Courses", pending_div)
c5.metric("Completed Courses", completed_div)
c6.metric("Completion %", f"{div_pct}%")

st.divider()

# --------------------------------------------------
# 🏢 UNIT-WISE SUMMARY
# --------------------------------------------------
st.subheader("🏢 Unit-wise Course Status")

unit_rows = []

for unit, g in df.groupby("Unit"):
    strength = SANCTIONED_STRENGTH[unit]
    total = strength * num_courses
    pending = g[course_cols].sum().sum()
    completed = total - pending
    pct = round((completed / total) * 100, 2)

    unit_rows.append({
        "Unit": unit,
        "Sanctioned Strength": strength,
        "Total Courses": total,
        "Pending Courses": pending,
        "Completed Courses": completed,
        "Completion %": pct
    })

st.dataframe(pd.DataFrame(unit_rows), use_container_width=True)

st.divider()

# --------------------------------------------------
# 🔍 LIVE NAME SEARCH (RESTORED)
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

display_df = matches[
    ["Employee Name", "Office of Working", "Unit", "Completion %"]
].reset_index(drop=True)

st.caption("Matching employees")
st.dataframe(display_df, use_container_width=True)

if len(display_df) == 1:
    selected_name = display_df.loc[0, "Employee Name"]
else:
    selected_name = st.selectbox(
        "Select person",
        display_df["Employee Name"].tolist()
    )

# --------------------------------------------------
# 👤 INDIVIDUAL REPORT
# --------------------------------------------------
user = df[df["Employee Name"] == selected_name]
pct = float(user["Completion %"].iloc[0])
color = completion_color(pct)

st.markdown(
    f"<h3 style='color:{color};'>👤 {selected_name} — Completion: {pct}%</h3>",
    unsafe_allow_html=True
)

pending = user[course_cols].T.reset_index()
pending.columns = ["Course Name", "Pending"]
pending = pending[pending["Pending"] == 1]

st.subheader("📘 Pending Courses")
if pending.empty:
    st.success("🎉 No pending courses")
else:
    st.dataframe(pending, use_container_width=True)

st.divider()

# --------------------------------------------------
# 🚨 ZERO COMPLETION EMPLOYEES
# --------------------------------------------------
st.subheader("🚨 Employees who have NOT completed even ONE course")

zero_completed = df[df["Completed Courses"] == 0]

if zero_completed.empty:
    st.success("🎉 All employees have completed at least one course")
else:
    st.error(f"⚠️ {len(zero_completed)} employees have completed ZERO courses")
    st.dataframe(
        zero_completed[
            ["Employee Name", "Office of Working", "Unit", "Pending Courses"]
        ],
        use_container_width=True
    )
