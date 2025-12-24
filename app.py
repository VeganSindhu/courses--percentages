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
# INITIALIZE FILE & DATE
# --------------------------------------------------
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
# LOAD DATA (CACHE-SAFE)
# --------------------------------------------------
@st.cache_data
def load_data(data_file):
    df = pd.read_excel(data_file)
    df.columns = df.columns.astype(str).str.strip()

    # Identify course columns safely
    ignore = {"Employee Name", "Office of Working", "Total Courses"}
    course_cols = [
        c for c in df.columns
        if c not in ignore and pd.api.types.is_numeric_dtype(df[c])
    ]

    # Force strict 0 / 1 integers
    df[course_cols] = df[course_cols].fillna(0).astype(int)

    num_employees = df.shape[0]
    num_courses = len(course_cols)

    # Employee-level calculations
    df["Pending Courses"] = df[course_cols].sum(axis=1)
    df["Completed Courses"] = num_courses - df["Pending Courses"]
    df["Completion %"] = round(
        (df["Completed Courses"] / num_courses) * 100, 2
    )

    # Unit mapping
    df["Unit"] = df["Office of Working"].apply(unit_name)

    # Division-level calculation (CORRECT BASE)
    total_slots = num_employees * num_courses
    pending_slots = df[course_cols].sum().sum()
    completed_slots = total_slots - pending_slots

    division_pct = round(
        (completed_slots / total_slots) * 100, 2
    )

    return df, course_cols, num_courses, division_pct


df, course_cols, total_courses, division_pct = load_data(DATA_FILE)

# --------------------------------------------------
# DIVISION & UNIT SUMMARY
# --------------------------------------------------
st.subheader("📊 Division Completion Status")
st.metric("Division Completion %", f"{division_pct}%")

unit_rows = []

for unit, g in df.groupby("Unit"):
    n_emp = g.shape[0]
    total_slots = n_emp * total_courses

    pending_slots = g[course_cols].sum().sum()
    completed_slots = total_slots - pending_slots

    pct = round(
        (completed_slots / total_slots) * 100, 2
    )

    unit_rows.append({
        "Unit": unit,
        "Completion %": pct
    })

st.subheader("🏢 Unit-wise Completion %")
st.dataframe(pd.DataFrame(unit_rows), use_container_width=True)

st.divider()

# --------------------------------------------------
# LIVE SEARCH WITH CLEAR BUTTON
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
# USER REPORT
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
