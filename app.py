import streamlit as st
import pandas as pd
import os
from datetime import datetime

st.set_page_config(page_title="Pivot Report Generator", layout="wide")
st.title("📊 Automated Pivot Report Generator by Davy")

# File uploader
uploaded_file = st.file_uploader("Upload your Excel file with aggregate data", type=[".xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name='Aggregate Data')

    # Convert Time to numeric and compute Hours
    df['Time'] = pd.to_numeric(df['Time'], errors='coerce')
    df['Hours'] = df['Time'] / 60

    # Extract team members
    team_members = df['Source.Name'].dropna().unique()

    st.success(f"File loaded successfully. Found {len(team_members)} team members.")

    # Preview data
    if st.checkbox("Preview raw data"):
        st.dataframe(df.head(20))

    if st.button("Generate Pivot Report"):
        output_file = f"pivot_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"

        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # === Summary Sheet ===
            summary = pd.pivot_table(
                df,
                index=['Client', 'Activity Name'],
                values='Hours',
                aggfunc='sum',
                fill_value=0
            ).reset_index().sort_values(by=['Client', 'Activity Name'])
            summary.to_excel(writer, sheet_name='Summary', index=False)

            # === Detailed Logs and Pivots ===
            for member in team_members:
                member_df = df[df['Source.Name'] == member]

                # Pivot 1: Client x Week
                pivot1 = pd.pivot_table(member_df, index='Client', columns='Week', values='Hours', aggfunc='sum', fill_value=0)
                pivot1.to_excel(writer, sheet_name=f"{member[:15]}_ClientWeek")

                # Pivot 2: Activity x Client
                pivot2 = pd.pivot_table(member_df, index='Activity Name', columns='Client', values='Hours', aggfunc='sum', fill_value=0)
                pivot2.to_excel(writer, sheet_name=f"{member[:15]}_ActClient")

                # Pivot 3: Activity x Week
                pivot3 = pd.pivot_table(member_df, index='Activity Name', columns='Week', values='Hours', aggfunc='sum', fill_value=0)
                pivot3.to_excel(writer, sheet_name=f"{member[:15]}_ActWeek")

                # Detailed log like your screenshot
                detailed_log = member_df[['Client', 'Week', 'Activity Name', 'Comments', 'Time', 'Hours']].sort_values(by=['Client', 'Week'])
                detailed_log.to_excel(writer, sheet_name=f"{member[:15]}_Details", index=False)

        with open(output_file, "rb") as f:
            st.download_button(
                label="📥 Download Pivot Report",
                data=f,
                file_name=output_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

