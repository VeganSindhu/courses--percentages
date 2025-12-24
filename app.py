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


# ✅ THIS LINE WAS MISSING
DATA_FILE, AS_ON_DATE = get_data_file_and_date()
