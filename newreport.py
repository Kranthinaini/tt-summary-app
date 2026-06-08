import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="TT Summary Report", layout="wide")

st.title("TT Summary Report")

# Upload Files
credit_file = st.file_uploader(
    "Upload Credit Report",
    type=["xlsx"],
    key="credit"
)

db_file = st.file_uploader(
    "Upload DB Report",
    type=["xlsx"],
    key="db"
)

if credit_file is not None and db_file is not None:

    # Read Credit Report
    credit_df = pd.read_excel(credit_file)
    credit_df.columns = credit_df.columns.str.strip()

    # Read DB workbook
    xls = pd.ExcelFile(db_file)

    st.subheader("Select DB Sheets")

    vfs_sheet = st.selectbox(
        "Select VFS Sheet",
        xls.sheet_names
    )

    vivifi_sheet = st.selectbox(
        "Select VIVIFI Sheet",
        xls.sheet_names
    )

    # Read selected sheets
    vfs_df = pd.read_excel(db_file, sheet_name=vfs_sheet)
    vivifi_df = pd.read_excel(db_file, sheet_name=vivifi_sheet)

    # Combine
    db_df = pd.concat(
        [vfs_df, vivifi_df],
        ignore_index=True
    )

    db_df.columns = db_df.columns.str.strip()

    #st.write("DB Columns")
    #st.write(db_df.columns.tolist())

    #st.write("Credit Columns")
    #st.write(credit_df.columns.tolist())
    # Standardize Cluster Column

    credit_df.columns = credit_df.columns.str.strip()
    db_df.columns = db_df.columns.str.strip()

    credit_df["Cluster"] = credit_df["Cluster Name"]   


    start_date = st.date_input("From Date")
    end_date = st.date_input("To Date")

    if st.button("Generate Report"):

        # -------------------------
        # Credit Dates
        # -------------------------
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

        # -------------------------
        # DB Date
        # -------------------------
        db_df["Disbursement Date"] = pd.to_datetime(
            db_df["Disbursement Date"],
            errors="coerce"
        )

        # -------------------------
        # TCM Submitted
        # -------------------------
        tcm_sub = (
            credit_df[
                (credit_df["Submitted to TCM Date Time"] >= pd.Timestamp(start_date))
                &
                (credit_df["Submitted to TCM Date Time"] <= pd.Timestamp(end_date))
            ]
            .groupby("Cluster")
            .size()
            .rename("TCM_Submitted")
        )

        # -------------------------
        # Rejected
        # -------------------------
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
            .groupby("Cluster")
            .size()
            .rename("Rejected")
        )

        # -------------------------
        # Approved
        # -------------------------
        approved = (
            credit_df[
                (credit_df["CCT Approved Date Time"] >= pd.Timestamp(start_date))
                &
                (credit_df["CCT Approved Date Time"] <= pd.Timestamp(end_date))
            ]
            .groupby("Cluster")
            .size()
            .rename("Approved")
        )
        # -------------------------
        # DB Filter
        # -------------------------
        db_filtered = db_df[
            (db_df["Disbursement Date"] >= pd.Timestamp(start_date))
            &
            (db_df["Disbursement Date"] <= pd.Timestamp(end_date))
        ]

        # -------------------------
        # DB Count
        # -------------------------
        db_count = (
            db_filtered
            .groupby("Cluster")
            .size()
            .rename("DB_Count")
        )

        # -------------------------
        # DB Amount
        # -------------------------
        db_amt = (
            db_filtered
            .groupby("Cluster")["Disbursement Amount"]
            .sum()
            .rename("DB_Amount")
        )
        # -------------------------
        # Final Report
        # -------------------------
        result = pd.concat(
            [tcm_sub, reject_count, approved, db_count, db_amt],
            axis=1
        ).fillna(0)

        result = result.reset_index()

        for col in [
            "TCM_Submitted",
            "Rejected",
            "Approved",
            "DB_Count"
        ]:
            result[col] = result[col].astype(int)

        #st.subheader("TT Summary")
        #st.dataframe(result, use_container_width=True)
        total_row = pd.DataFrame({
            "Cluster": ["TOTAL"],
            "TCM_Submitted": [result["TCM_Submitted"].sum()],
            "Rejected": [result["Rejected"].sum()],
            "Approved": [result["Approved"].sum()],
            "DB_Count": [result["DB_Count"].sum()],
            "DB_Amount": [result["DB_Amount"].sum()]
        })
        # Get Cluster-Territory mapping
        cluster_tt_credit = credit_df[
            ["Territory", "Cluster"]
        ].drop_duplicates()

        cluster_tt_db = db_df[
            ["Territory", "Cluster"]
        ].drop_duplicates()

        cluster_tt_map = pd.concat(
            [cluster_tt_credit, cluster_tt_db],
            ignore_index=True
        ).drop_duplicates(subset=["Cluster"])

        result = result.merge(
            cluster_tt_map,
            on="Cluster",
            how="left"
        )

        #st.write("Columns After Merge")
        #st.write(result.columns.tolist())
            
        result = result.sort_values(
            ["Territory", "Cluster"]
        )
        final_rows = []

        for tt, grp in result.groupby("Territory", sort=False):

            final_rows.append(grp)

            total_row_tt = pd.DataFrame({
                "Territory": [tt],
                "Cluster": [f"TOTAL-{tt}"],
                "TCM_Submitted": [grp["TCM_Submitted"].sum()],
                "Rejected": [grp["Rejected"].sum()],
                "Approved": [grp["Approved"].sum()],
                "DB_Count": [grp["DB_Count"].sum()],
                "DB_Amount": [grp["DB_Amount"].sum()]
            })

            final_rows.append(total_row_tt)

        final_report = pd.concat(
            final_rows,
            ignore_index=True
        )

        final_report = final_report[
            [
                "Territory",
                "Cluster",
                "TCM_Submitted",
                "Rejected",
                "Approved",
                "DB_Count",
                "DB_Amount"
            ]
        ]
        st.subheader("Cluster Summary with TT Totals")
        st.dataframe(final_report, use_container_width=True)

        st.subheader("Grand Total")
        st.dataframe(total_row, use_container_width=True)

        output = BytesIO()

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            final_report.to_excel(
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
        #python -m streamlit run newreport.py