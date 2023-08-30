import streamlit as st
import numpy as np
import pandas as pd

from datetime import datetime, timedelta, time
from io import BytesIO

pd.set_option('display.max_columns', None)

st.set_page_config(
    page_title="Overtime Hours Calculator",
    layout="wide"
)

with st.sidebar:
    st.success("Select a page above")

reduce_header_height_style = """
    <style>
        div.block-container {
            padding-top: 1rem;
            padding-bottom: 1rem;
        }
    </style>
"""
st.markdown(reduce_header_height_style, unsafe_allow_html=True)

st.markdown(f"<h1 style='text-align: center;'>Overtime Hours Calculator</h1>", unsafe_allow_html=True)

# Note
st.markdown(f"<h3>Note</h3>", unsafe_allow_html=True)

note = """In Version 2, on weekdays, the working hours are 07:30 to 17:00. 
On weekends or designated holidays, the working hours are 08:00 to 17:00."""
st.info(note, icon="ℹ️")

# Load Data
st.markdown(f"<h3>Load Data</h3>", unsafe_allow_html=True)

file = st.file_uploader("Upload your file in Excel format", type=["xlsx"])

if file is not None:
    df_ver_2 = pd.read_excel(file)
    df_ver_2 = df_ver_2.drop("NO", axis=1)

    # Clean Data
    for col in df_ver_2.columns[1:]:
        df_ver_2[col] = df_ver_2[col].apply(
            lambda x: 
            np.nan if pd.isnull(x) or isinstance(x, time) and not ('-' in str(x)) else 
            ' - '.join([str(x).split(' - ')[0], str(x).split(' - ')[-1]]) 
            if len(str(x).split(' - ')) != 2 else str(x)
        )

    # Select Days Off
    days_off_columns = st.multiselect(
        "Select Days Off",
        df_ver_2.columns[1:],
        # [1, 2, 8, 9, 15, 16, 19, 22, 23, 29, 30]
    )

    # Calculate Overtime
    def calculate_overtime(time_str, is_day_off=False):
        # If NaN or None
        if not time_str or pd.isnull(time_str):
            return 0

        # Split start and end times
        start_str, end_str = time_str.split(' - ')
        start_time = datetime.strptime(start_str, "%H:%M:%S").time()
        end_time = datetime.strptime(end_str, "%H:%M:%S").time()

        # Regular workday rules
        if not is_day_off:
            # Adjust start time if before 07:30
            if start_time < datetime.strptime("07:30:00", "%H:%M:%S").time():
                start_time = datetime.strptime("07:30:00", "%H:%M:%S").time()

            # Calculate lateness in starting
            late_minutes = max(0, (datetime.combine(datetime.today(), start_time) - datetime.combine(datetime.today(), datetime.strptime("07:30:00", "%H:%M:%S").time())).seconds / 60)

            # Calculate adjusted end time by adding late_minutes to 17:00
            actual_end_time = (datetime.combine(datetime.today(), datetime.strptime("17:00:00", "%H:%M:%S").time()) + timedelta(minutes=late_minutes)).time()

            # Check the total time worked
            total_work_seconds = (datetime.combine(datetime.today(), end_time) - datetime.combine(datetime.today(), start_time)).total_seconds()
            if total_work_seconds <= late_minutes * 60:  # Convert late minutes to seconds
                return 0

            # Adjusted overtime calculation
            overtime_minutes = (datetime.combine(datetime.today(), min(end_time, datetime.strptime("21:00:00", "%H:%M:%S").time())) - datetime.combine(datetime.today(), actual_end_time)).seconds / 60 - late_minutes
            overtime_hours = overtime_minutes / 60
            
            # Cap at 3 hours
            return min(3, overtime_hours)

        # Day off rules
        else:
            # Adjust start time if before 08:00
            if start_time < datetime.strptime("08:00:00", "%H:%M:%S").time():
                start_time = datetime.strptime("08:00:00", "%H:%M:%S").time()

            # Cap end time at 17:00
            if end_time > datetime.strptime("17:00:00", "%H:%M:%S").time():
                end_time = datetime.strptime("17:00:00", "%H:%M:%S").time()

            # Calculate overtime
            overtime_minutes = (datetime.combine(datetime.today(), end_time) - datetime.combine(datetime.today(), start_time)).seconds / 60
            overtime_hours = max(0, overtime_minutes // 60)

            # Cap at 8 hours
            return min(8, overtime_hours)

    # Create a new DataFrame to store overtime values
    overtime_df_ver_2 = df_ver_2.copy()
    overtime_df_ver_2.iloc[:, 1:] = 0  # Initialize overtime values to zero

    for col in df_ver_2.columns[1:]:
        overtime_df_ver_2[col] = df_ver_2[col].apply(lambda x: calculate_overtime(x, int(col) in days_off_columns))

    if days_off_columns:
        st.markdown(f"<h3>Calculate Overtime</h3>", unsafe_allow_html=True)

        with st.expander("Click here to view the data overview"):
            for col in overtime_df_ver_2.columns[1:]:
                overtime_df_ver_2[col] = overtime_df_ver_2[col].astype(int)
            overtime_df_ver_2.index += 1
            st.dataframe(overtime_df_ver_2)

        def convert_df_ver_2(df_ver_2):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_ver_2.to_excel(writer, index=False, sheet_name='Sheet1')
            output.seek(0)
            return output

        excel = convert_df_ver_2(overtime_df_ver_2)

        st.download_button(
            label="Download data as Excel",
            data=excel.getvalue(),
            file_name="results_version_2.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )