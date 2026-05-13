import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.title("DB Target Contribution Calculator")

# Inputs
total_db = st.number_input("Total DB Target", value=1000)
ruchi_target = st.number_input("Ruchi DB Target", value=500)
radhuni_target = st.number_input("Radhuni DB Target", value=500)

st.subheader("Market Input Table")

# Sample market input
data = {
    "Market": ["A-1", "A-2", "B-1"],
    "Type": ["Combined", "Ruchi Only", "Radhuni Only"],
    "Manual_Contribution": [0, 0, 0]
}

df = pd.DataFrame(data)
edited_df = st.data_editor(df, num_rows="dynamic")

# Logic
def calculate_contribution(row):
    if row["Type"] == "Combined":
        ruchi_contrib = (ruchi_target / total_db) * 100
        radhuni_contrib = (radhuni_target / total_db) * 100
    elif row["Type"] == "Ruchi Only":
        ruchi_contrib = 100
        radhuni_contrib = 0
    elif row["Type"] == "Radhuni Only":
        ruchi_contrib = 0
        radhuni_contrib = 100
    else:
        ruchi_contrib = 0
        radhuni_contrib = 0

    return pd.Series([ruchi_contrib, radhuni_contrib])

# Apply logic
edited_df[["Ruchi_%", "Radhuni_%"]] = edited_df.apply(calculate_contribution, axis=1)

# Add value calculation
edited_df["Ruchi_Value"] = (edited_df["Ruchi_%"] / 100) * ruchi_target
edited_df["Radhuni_Value"] = (edited_df["Radhuni_%"] / 100) * radhuni_target

st.subheader("Output Table")
st.dataframe(edited_df)

# Excel download function
def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Contribution")
    return output.getvalue()

excel_data = to_excel(edited_df)

st.download_button(
    label="Download Excel",
    data=excel_data,
    file_name="db_contribution_report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)