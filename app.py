import streamlit as st
import pandas as pd
import os
from datetime import datetime

st.set_page_config(page_title="Pivot Report Generator", layout="wide")
st.title("📊 Automated Pivot Report Generator by Konan Davy")

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
            from openpyxl.utils.dataframe import dataframe_to_rows
            from openpyxl.styles import Font, PatternFill
            from openpyxl import Workbook

            wb = Workbook()
            ws = wb.active
            ws.title = 'Summary'

            def write_pivot_block(title, df_block):
                ws.append([title])
                ws.append(list(df_block.columns))
                for r in df_block.itertuples(index=False):
                    ws.append(list(r))
                ws.append([None])

            # Pivot 1: Client Totals
            pivot1 = df.groupby('Client')['Hours'].sum().reset_index()
            pivot1['Hours'] = pivot1['Hours'].round(2)
            pivot1.columns = ['Row Labels', 'Sum of Hours']
            write_pivot_block('Client Totals', pivot1)

            # Pivot 2: Activity Totals
            pivot2 = df.groupby('Activity Name')['Hours'].sum().reset_index()
            pivot2['Hours'] = pivot2['Hours'].round(2)
            pivot2.columns = ['Row Labels', 'Sum of Hours']
            write_pivot_block('Activity Totals', pivot2)

            # Pivot 3: Week Totals
            pivot3 = df.groupby('Week')['Hours'].sum().reset_index()
            pivot3['Hours'] = pivot3['Hours'].round(2)
            pivot3.columns = ['Row Labels', 'Sum of Hours']
            write_pivot_block('Week Totals', pivot3)

            # Grand Total
            grand_total = round(df['Hours'].sum(), 2)
            ws.append(['Grand Total', grand_total])

            # Style Headers
            bold_font = Font(bold=True)
            fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
            for row in ws.iter_rows():
                if row[0].value in ['Client Totals', 'Activity Totals', 'Week Totals', 'Grand Total']:
                    for cell in row:
                        if cell.value is not None:
                            cell.font = bold_font
                            cell.fill = fill

            wb.save(output_file)

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

