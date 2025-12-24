import streamlit as st
import pandas as pd
import os
import re
from datetime import datetime

# ---------------- CONFIG ----------------
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
    f, d = excel_files[0]
    return f, parse(d).strftime("%d.%m.%Y")


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
df = pd.read_excel(DATA_FILE)
df.columns = df.columns.astype(str).str.strip()

ignore = {"Employee Name", "Office of Working", "Total Courses"}
course_cols = [c for c in df.columns if c not in ignore and pd.api.types.is_numeric_dtype(df[c])]

df[course_cols] = df[course_cols].fillna(0).astype(int)
df["Unit"] = df["Office of Working"].apply(unit_name)

num_courses = len(course_cols)

# ---------------- DIVISION ----------------
pending_div = df[course_cols].sum().sum()
total_div = SANCTIONED_STRENGTH["Division"] * num_courses
completed_div = total_div - pending_div
div_pct = round((completed_div / total_div) * 100, 2)

st.subheader("📊 Division-level Course Status")

c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Sanctioned Strength", SANCTIONED_STRENGTH["Division"])
c2.metric("Courses / Employee", num_courses)
c3.metric("Total Courses", total_div)
c4.metric("Pending Courses", pending_div)
c5.metric("Completion %", f"{div_pct}%")

st.divider()

# ---------------- UNIT WISE ----------------
rows = []
for unit, g in df.groupby("Unit"):
    strength = SANCTIONED_STRENGTH[unit]
    pending = g[course_cols].sum().sum()
    total = strength * num_courses
    completed = total - pending
    pct = round((completed / total) * 100, 2)

    rows.append({
        "Unit": unit,
        "Sanctioned Strength": strength,
        "Total Courses": total,
        "Pending Courses": pending,
        "Completed Courses": completed,
        "Completion %": pct
    })

st.subheader("🏢 Unit-wise Course Status")
st.dataframe(pd.DataFrame(rows), use_container_width=True)

st.divider()

# ---------------- ZERO COMPLETION ----------------
st.subheader("🚨 Employees who have NOT completed even ONE course")
zero = df[df[course_cols].sum(axis=1) == num_courses]

if zero.empty:
    st.success("🎉 All employees have completed at least one course")
else:
    st.error(f"⚠️ {len(zero)} employees have completed ZERO courses")
    st.dataframe(
        zero[["Employee Name", "Office of Working", "Unit"]],
        use_container_width=True
    )
