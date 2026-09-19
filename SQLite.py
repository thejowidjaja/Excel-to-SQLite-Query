import streamlit as st
import pandas as pd
import sqlite3
from io import BytesIO

st.title("Community Service Compiler")

uploaded_file = st.file_uploader(
    "Upload community service Excel file",
    type=["xlsx"]
)

if uploaded_file is not None:

    # 1. Read uploaded Excel file
    df = pd.read_excel(uploaded_file)

    #st.write("Uploaded data:")
    #st.dataframe(df.head())

    # 2. Create temporary SQLite database
    conn = sqlite3.connect(":memory:")

    # 3. Put Excel data into SQLite
    df.to_sql(
        "form_responses",
        conn,
        if_exists="replace",
        index=False
    )

    # 4. SQL query
    query = """
    SELECT
        "First Name (legal name)",
        "Last Name (legal name)",
        "Student ID",
        SUM("Total Hours") AS total_service_hours,
        GROUP_CONCAT(
            "Select an approved community service event",
            ', '
        ) AS volunteer_events
    FROM form_responses
    GROUP BY
        "Student ID",
        "First Name (legal name)",
        "Last Name (legal name)";
    """

    # 5. Run query
    result_df = pd.read_sql_query(query, conn)

    conn.close()

    st.write("Compiled results:")
    st.dataframe(result_df)

    # 6. Create output Excel file in memory
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        result_df.to_excel(
            writer,
            index=False,
            sheet_name="Compiled Service Hours"
        )

    output.seek(0)

    # 7. Download button
    st.download_button(
        label="Download Compiled Excel File",
        data=output,
        file_name="compiled_service_hours.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
