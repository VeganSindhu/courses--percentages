import streamlit as st
import pandas as pd
from io import BytesIO

# --------------------------------------------------
# PAGE CONFIG
# --------------------------------------------------
st.set_page_config(page_title="Course Completion Report", layout="wide")
st.title("📘 Course Completion Report")

uploaded_file = st.file_uploader(
    "Upload Excel (.xlsx multi-sheet)",
    type=["xlsx"]
)

if not uploaded_file:
    st.info("Upload the course completion Excel file")
    st.stop()

# --------------------------------------------------
# HELPERS
# --------------------------------------------------
def df_to_excel_bytes(df, sheet_name="Report"):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    buffer.seek(0)
    return buffer.getvalue()


def normalize_columns(df):
    df.columns = df.columns.map(lambda x: str(x).strip() if pd.notna(x) else "Unnamed")
    cols = pd.Series(df.columns)
    for dup in cols[cols.duplicated()].unique():
        idxs = cols[cols == dup].index.tolist()
        for i, idx in enumerate(idxs):
            if i > 0:
                cols[idx] = f"{dup}.{i}"
    df.columns = cols
    return df


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

# --------------------------------------------------
# READ & CONSOLIDATE EXCEL
# --------------------------------------------------
xls = pd.ExcelFile(uploaded_file)
combined_df = pd.DataFrame()

for sheet in xls.sheet_names:
    df_sheet = pd.read_excel(uploaded_file, sheet_name=sheet)

    df_sheet.columns = df_sheet.iloc[0]
    df_sheet = df_sheet[1:]
    df_sheet = df_sheet.dropna(axis=1, how="all")
    df_sheet = normalize_columns(df_sheet)

    # Column E = Office of Working
    office_col = df_sheet.columns[4]
    df_sheet = df_sheet.rename(columns={office_col: "Office of Working"})

    division_col = next(
        (c for c in df_sheet.columns if "division" in c.lower() or "unit" in c.lower()),
        None
    )
    if not division_col:
        continue

    df_tp = df_sheet[df_sheet[division_col].astype(str).str.contains("RMS TP", case=False, na=False)]
    if df_tp.empty:
        continue

    df_tp = normalize_columns(df_tp)
    df_tp["Course Name"] = sheet

    name_col = next((c for c in df_tp.columns if "name" in c.lower()), None)
    if name_col:
        df_tp = df_tp.rename(columns={name_col: "Employee Name"})

    combined_df = pd.concat([combined_df, df_tp], ignore_index=True)

if combined_df.empty:
    st.error("No RMS TP data found in Excel.")
    st.stop()

# --------------------------------------------------
# PIVOT + CALCULATIONS
# --------------------------------------------------
pivot_df = combined_df.pivot_table(
    index=["Employee Name", "Office of Working"],
    columns="Course Name",
    aggfunc="size",
    fill_value=0
).reset_index()

course_cols = pivot_df.columns[2:]
total_courses = len(course_cols)

pivot_df["Completed Courses"] = pivot_df[course_cols].sum(axis=1)
pivot_df["Completion %"] = round(
    (pivot_df["Completed Courses"] / total_courses) * 100, 2
)

pivot_df["Office Group"] = pivot_df["Office of Working"].apply(office_group)

# --------------------------------------------------
# DIVISION COMPLETION %
# --------------------------------------------------
division_completion_pct = round(
    (pivot_df["Completed Courses"].sum() / (len(pivot_df) * total_courses)) * 100, 2
)

st.subheader("📊 RMS TP Division Completion Status")
st.metric("Division Completion %", f"{division_completion_pct}%")

st.divider()

# --------------------------------------------------
# OFFICE-WISE COMPLETION %
# --------------------------------------------------
st.subheader("🏢 Office-wise Completion %")

office_summary = (
    pivot_df.groupby("Office Group")
    .apply(lambda x: round((x["Completed Courses"].sum() / (len(x) * total_courses)) * 100, 2))
    .reset_index(name="Completion %")
)

st.dataframe(office_summary)

st.divider()

# --------------------------------------------------
# SEARCH EMPLOYEE
# --------------------------------------------------
st.subheader("🔍 Search Employee")

names = sorted(pivot_df["Employee Name"].dropna().unique())
search_text = st.text_input("Type at least 4 characters of your name")

selected_name = None
if len(search_text) >= 4:
    matches = [n for n in names if search_text.lower() in n.lower()]
    if matches:
        selected_name = st.selectbox("Select your name", matches)
    else:
        st.info("No matching name found")

if not selected_name:
    st.stop()

# --------------------------------------------------
# DISPLAY USER REPORT WITH COLORED NAME
# --------------------------------------------------
user_row = pivot_df[pivot_df["Employee Name"] == selected_name]
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
    data=df_to_excel_bytes(user_row, "My Report"),
    file_name=f"{selected_name}_course_report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
