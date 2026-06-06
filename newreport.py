import streamlit as st
import pandas as pd
from io import BytesIO

st.title("TT Summary Report")

# Upload File
uploaded_file = st.file_uploader(
    "Upload Credit Report",
    type=["xlsx"]
)

if uploaded_file is not None:

    # Read workbook
    xls = pd.ExcelFile(uploaded_file)

    # Read sheets
    credit_df = pd.read_excel(xls, sheet_name=0)
    db_df = pd.read_excel(xls, sheet_name="db data")

    # Clean column names
    credit_df.columns = credit_df.columns.str.strip()
    db_df.columns = db_df.columns.str.strip()

    # Date selection
    start_date = st.date_input("From Date")
    end_date = st.date_input("To Date")

    if st.button("Generate Report"):

        # Convert dates
        credit_df["Submitted to TCM Date Time"] = pd.to_datetime(
            credit_df["Submitted to TCM Date Time"],
            errors="coerce"
        )

        credit_df["Latest Status Date"] = pd.to_datetime(
            credit_df["Latest Status Date"],
            errors="coerce"
        )

        credit_df["CCT Approved Date Time"] = pd.to_datetime(
            credit_df["CCT Approved Date Time"],
            errors="coerce"
        )

        db_df["Disbursed Date"] = pd.to_datetime(
            db_df["Disbursed Date"],
            errors="coerce"
        )

        # ------------------------
        # TCM Submitted
        # ------------------------

        tcm_sub = (
            credit_df[
                (credit_df["Submitted to TCM Date Time"] >= pd.Timestamp(start_date))
                &
                (credit_df["Submitted to TCM Date Time"] <= pd.Timestamp(end_date))
            ]
            .groupby("Territory")
            .size()
            .rename("TCM_Submitted")
        )

        # ------------------------
        # Rejected
        # ------------------------

        reject_statuses = [
            "RCM Verified and Rejected",
            "RCM Rejected",
            "TCM Rejected",
            "CCT Rejected"
        ]

        reject_count = (
            credit_df[
                (credit_df["Latest Status Date"] >= pd.Timestamp(start_date))
                &
                (credit_df["Latest Status Date"] <= pd.Timestamp(end_date))
                &
                (credit_df["Final Status"].isin(reject_statuses))
            ]
            .groupby("Territory")
            .size()
            .rename("Rejected")
        )

        # ------------------------
        # Approved
        # ------------------------

        approved = (
            credit_df[
                (credit_df["CCT Approved Date Time"] >= pd.Timestamp(start_date))
                &
                (credit_df["CCT Approved Date Time"] <= pd.Timestamp(end_date))
            ]
            .groupby("Territory")
            .size()
            .rename("Approved")
        )

        # ------------------------
        # DB Data Sheet
        # ------------------------

        db_filtered = db_df[
            (db_df["Disbursed Date"] >= pd.Timestamp(start_date))
            &
            (db_df["Disbursed Date"] <= pd.Timestamp(end_date))
        ]

        db_count = (
            db_filtered
            .groupby("Territory")
            .size()
            .rename("DB_Count")
        )

        db_amt = (
            db_filtered
            .groupby("Territory")["Disbursement amount"]
            .sum()
            .rename("DB_Amount")
        )

        # ------------------------
        # Final Report
        # ------------------------

        result = pd.concat(
            [tcm_sub, reject_count, approved, db_count, db_amt],
            axis=1
        ).fillna(0)

        result = result.reset_index()

        # Convert count columns to integers
        for col in [
            "TCM_Submitted",
            "Rejected",
            "Approved",
            "DB_Count"
        ]:
            result[col] = result[col].astype(int)

        st.subheader("TT Summary")

        st.dataframe(result)

        # Grand Total

        total_row = pd.DataFrame({
            "Territory": ["TOTAL"],
            "TCM_Submitted": [result["TCM_Submitted"].sum()],
            "Rejected": [result["Rejected"].sum()],
            "Approved": [result["Approved"].sum()],
            "DB_Count": [result["DB_Count"].sum()],
            "DB_Amount": [result["DB_Amount"].sum()]
        })

        st.subheader("Grand Total")

        st.dataframe(total_row)

        # Download Excel

        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            result.to_excel(
                writer,
                sheet_name="TT Summary",
                index=False
            )

            total_row.to_excel(
                writer,
                sheet_name="Grand Total",
                index=False
            )

        st.download_button(
            label="Download TT Summary",
            data=output.getvalue(),
            file_name="TT_Summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        #python -m streamlit run newreport.py