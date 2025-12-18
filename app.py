import streamlit as st
import pandas as pd

# --------------------------------------------------
# PAGE CONFIG
# --------------------------------------------------
st.set_page_config(page_title="Course Completion Status", layout="wide")
st.title("📘 Course Completion Status")

DATA_FILE = "data.xlsx"   # Admin updates this in GitHub

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


@st.cache_data
def load_data():
    df = pd.read_excel(DATA_FILE)
    df.columns = df.columns.astype(str).str.strip()

    # Required columns
    if "Employee Name" not in df.columns or "Office of Working" not in df.columns:
        raise ValueError("Required columns missing in Excel")

    # Course columns = numeric except known fields
    ignore = {"Employee Name", "Office of Working", "Total Courses"}
    course_cols = [
        c for c in df.columns
        if c not in ignore and pd.api.types.is_numeric_dtype(df[c])
    ]

    if not course_cols:
        raise ValueError("No course columns detected")

    total_courses = len(course_cols)

    # 1 = Pending, 0 / blank = Completed
    df[course_cols] = df[course_cols].fillna(0)

    df["Pending Courses"] = df[course_cols].eq(1).sum(axis=1)
    df["Completed Courses"] = total_courses - df["Pending Courses"]

    df["Completion %"] = round(
        (df["Completed Courses"] / total_courses) * 100, 2
    )

    df["Unit"] = df["Office of Working"].apply(unit_name)

    # Division-level calculation
    total_slots = len(df) * total_courses
    pending_slots = df[course_cols].eq(1).sum().sum()
    completed_slots = total_slots - pending_slots

    division_pct = round((completed_slots / total_slots) * 100, 2)

    return df, course_cols, total_courses, division_pct


# --------------------------------------------------
# LOAD DATA (SHARED FOR ALL USERS)
# --------------------------------------------------
try:
    df, course_cols, total_courses, division_pct = load_data()
except Exception as e:
    st.error(str(e))
    st.stop()

# --------------------------------------------------
# DIVISION SUMMARY
# --------------------------------------------------
st.subheader("📊 Division Completion Status")
st.metric("Division Completion %", f"{division_pct}%")

# --------------------------------------------------
# UNIT SUMMARY
# --------------------------------------------------
unit_rows = []
for unit, g in df.groupby("Unit"):
    total_slots = len(g) * total_courses
    pending_slots = g[course_cols].eq(1).sum().sum()
    completed_slots = total_slots - pending_slots
    pct = round((completed_slots / total_slots) * 100, 2)
    unit_rows.append({"Unit": unit, "Completion %": pct})

st.subheader("🏢 Unit-wise Completion %")
st.dataframe(pd.DataFrame(unit_rows))

st.divider()

# --------------------------------------------------
# LIVE FILTERING (NO ENTER, NO DROPDOWN)
# --------------------------------------------------
st.subheader("🔍 Check Your Completion Status")

query = st.text_input("Start typing your name")

filtered_df = df
if query.strip():
    filtered_df = df[df["Employee Name"].str.contains(query, case=False, na=False)]

if filtered_df.empty:
    st.info("No matching names found")
    st.stop()

# Show filtered list
display_df = filtered_df[[
    "Employee Name",
    "Office of Working",
    "Unit",
    "Completion %"
]].reset_index(drop=True)

st.caption("Matching employees (live filtered)")

selected_index = st.radio(
    "Select your name",
    display_df.index,
    format_func=lambda i: display_df.loc[i, "Employee Name"]
)

selected_name = display_df.loc[selected_index, "Employee Name"]

# --------------------------------------------------
# USER REPORT
# --------------------------------------------------
user_row = df[df["Employee Name"] == selected_name]
pct = float(user_row["Completion %"].iloc[0])
color = completion_color(pct)

st.markdown(
    f"<h3 style='color:{color};'>👤 {selected_name} — Completion: {pct}%</h3>",
    unsafe_allow_html=True
)

# Pending courses
pending_courses = (
    user_row[course_cols]
    .T.reset_index()
)
pending_courses.columns = ["Course Name", "Pending (1 = Pending)"]
pending_courses = pending_courses[pending_courses["Pending (1 = Pending)"] == 1]

st.subheader("📘 Pending Courses")
if pending_courses.empty:
    st.success("🎉 No pending courses")
else:
    st.dataframe(pending_courses)

st.subheader("📄 Summary")
st.dataframe(user_row[[
    "Employee Name",
    "Office of Working",
    "Unit",
    "Completed Courses",
    "Pending Courses",
    "Completion %"
]])
