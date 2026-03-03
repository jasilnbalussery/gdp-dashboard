import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="RBM Worst 15 OSG Conversion Report", layout="wide")

st.title("📊 RBM Wise Worst 15 Staff - OSG Conversion Report")

# File Uploads
product_file = st.file_uploader("Upload PRODUCT File", type=["xlsx"])
osg_file = st.file_uploader("Upload OSG File", type=["xlsx"])
lg_file = st.file_uploader("Upload LG AMC File", type=["xlsx"])

if product_file and osg_file and lg_file:

    # ==========================
    # SAFE LOAD + STANDARDIZE
    # ==========================
    def load_file(file):
        df = pd.read_excel(file, engine="openpyxl")
        df.columns = df.columns.str.strip().str.upper()
        return df

    product = load_file(product_file)
    osg = load_file(osg_file)
    lg = load_file(lg_file)

    # ==========================
    # REMOVE RETURNS FROM PRODUCT
    # ==========================
    product['INVOICE NUMBER'] = product['INVOICE NUMBER'].astype(str)
    product['QTY'] = pd.to_numeric(product['QTY'], errors='coerce').fillna(0)

    # Remove matching +1 and -1 using group logic
    product = product.groupby(
        ['INVOICE NUMBER', 'IMEI', 'RBM', 'STAFF'],
        as_index=False
    )['QTY'].sum()

    # Keep only positive net sales
    product = product[product['QTY'] > 0]

    # ==========================
    # PRODUCT COUNT PER STAFF
    # ==========================
    total_product = (
        product.groupby(['RBM', 'STAFF'])['QTY']
        .sum()
        .reset_index(name='TOTAL_PRODUCT_QTY')
    )

    # ==========================
    # OSG DATA (Using EWS QTY)
    # ==========================
    if 'EWS QTY' not in osg.columns:
        st.error("EWS QTY column not found in OSG file")
        st.stop()

    osg['EWS QTY'] = pd.to_numeric(osg['EWS QTY'], errors='coerce').fillna(0)
    osg = osg[osg['EWS QTY'] > 0]

    osg_count = (
        osg.groupby(['RBM', 'STAFF'])['EWS QTY']
        .sum()
        .reset_index(name='OSG_QTY')
    )

    # Category wise OSG
    category_wise = (
        osg.groupby(['RBM', 'STAFF', 'ITEM CATEGORY'])['EWS QTY']
        .sum()
        .reset_index()
    )

    # ==========================
    # LG AMC DATA
    # ==========================
    lg['QTY'] = pd.to_numeric(lg['QTY'], errors='coerce').fillna(0)
    lg = lg[lg['QTY'] > 0]

    lg_count = (
        lg.groupby(['RBM', 'STAFF'])['QTY']
        .sum()
        .reset_index(name='LG_AMC_QTY')
    )

    # ==========================
    # MERGE ALL
    # ==========================
    report = total_product.merge(osg_count, on=['RBM', 'STAFF'], how='left')
    report = report.merge(lg_count, on=['RBM', 'STAFF'], how='left')

    report.fillna(0, inplace=True)

    # ==========================
    # CONVERSION %
    # ==========================
    report['OSG_CONVERSION_%'] = (
        report['OSG_QTY'] / report['TOTAL_PRODUCT_QTY']
    ) * 100

    report['OSG_CONVERSION_%'] = report['OSG_CONVERSION_%'].round(2)

    # ==========================
    # DEFINE NEAR ZERO (<=2%)
    # ==========================
    report['ZERO_OR_NEAR'] = np.where(
        report['OSG_CONVERSION_%'] <= 2,
        "YES",
        "NO"
    )

    # ==========================
    # WORST 15 PER RBM
    # ==========================
    worst_list = []

    for rbm in report['RBM'].unique():
        temp = report[
            (report['RBM'] == rbm) &
            (report['ZERO_OR_NEAR'] == "YES")
        ].sort_values(
            by=['OSG_QTY', 'OSG_CONVERSION_%'],
            ascending=[False, True]
        ).head(15)

        worst_list.append(temp)

    final_report = pd.concat(worst_list)

    # ==========================
    # CATEGORY SUMMARY PIVOT
    # ==========================
    category_summary = category_wise.pivot_table(
        index=['RBM', 'STAFF'],
        columns='ITEM CATEGORY',
        values='EWS QTY',
        aggfunc='sum'
    ).reset_index()

    category_summary.fillna(0, inplace=True)

    final_report = final_report.merge(
        category_summary,
        on=['RBM', 'STAFF'],
        how='left'
    )

    final_report.fillna(0, inplace=True)

    st.success("Report Generated Successfully")
    st.dataframe(final_report)

    # ==========================
    # EXPORT TO EXCEL
    # ==========================
    def to_excel(df):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Worst_15_Staff')
        return output.getvalue()

    excel_data = to_excel(final_report)

    st.download_button(
        label="📥 Download Excel Report",
        data=excel_data,
        file_name="RBM_Worst_15_OSG_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

else:
    st.info("Please upload all three files to generate report.")
