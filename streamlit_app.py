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
    # LOAD FILES
    # ==========================
    product = pd.read_excel(product_file)
    osg = pd.read_excel(osg_file)
    lg = pd.read_excel(lg_file)

    # ==========================
    # REMOVE RETURNS FROM PRODUCT
    # ==========================
    product['Invoice Number'] = product['Invoice Number'].astype(str)
    
    # Identify return invoices (QTY -1)
    returns = product[product['QTY'] == -1]

    # Remove both -1 and corresponding +1
    for _, row in returns.iterrows():
        invoice = row['Invoice Number']
        imei = row['IMEI']
        product = product[~(
            (product['Invoice Number'] == invoice) &
            (product['IMEI'] == imei)
        )]

    # Keep only positive sales
    product = product[product['QTY'] > 0]

    # ==========================
    # PRODUCT COUNT PER STAFF
    # ==========================
    total_product = product.groupby(
        ['RBM', 'Staff']
    )['QTY'].sum().reset_index()

    total_product.rename(columns={'QTY': 'Total_Product_Qty'}, inplace=True)

    # ==========================
    # OSG DATA
    # ==========================
    osg = osg[osg['QTY'] > 0]

    osg_count = osg.groupby(
        ['RBM', 'Staff']
    )['QTY'].sum().reset_index()

    osg_count.rename(columns={'QTY': 'OSG_Qty'}, inplace=True)

    # Category wise OSG
    category_wise = osg.groupby(
        ['RBM', 'Staff', 'Item Category']
    )['QTY'].sum().reset_index()

    # ==========================
    # LG AMC DATA
    # ==========================
    lg = lg[lg['QTY'] > 0]

    lg_count = lg.groupby(
        ['RBM', 'Staff']
    )['QTY'].sum().reset_index()

    lg_count.rename(columns={'QTY': 'LG_AMC_Qty'}, inplace=True)

    # ==========================
    # MERGE ALL
    # ==========================
    report = total_product.merge(osg_count, on=['RBM', 'Staff'], how='left')
    report = report.merge(lg_count, on=['RBM', 'Staff'], how='left')

    report.fillna(0, inplace=True)

    # Conversion %
    report['OSG_Conversion_%'] = (
        report['OSG_Qty'] / report['Total_Product_Qty']
    ) * 100

    report['OSG_Conversion_%'] = report['OSG_Conversion_%'].round(2)

    # ==========================
    # DEFINE NEAR ZERO (<=2%)
    # ==========================
    report['Zero_or_Near'] = np.where(
        report['OSG_Conversion_%'] <= 2,
        "Yes",
        "No"
    )

    # ==========================
    # WORST 15 PER RBM
    # ==========================
    worst_list = []

    for rbm in report['RBM'].unique():
        temp = report[
            (report['RBM'] == rbm) &
            (report['Zero_or_Near'] == "Yes")
        ].sort_values(
            by=['OSG_Qty', 'OSG_Conversion_%'],
            ascending=[False, True]
        ).head(15)

        worst_list.append(temp)

    final_report = pd.concat(worst_list)

    # ==========================
    # CATEGORY SUMMARY PIVOT
    # ==========================
    category_summary = category_wise.pivot_table(
        index=['RBM', 'Staff'],
        columns='Item Category',
        values='QTY',
        aggfunc='sum'
    ).reset_index()

    category_summary.fillna(0, inplace=True)

    final_report = final_report.merge(
        category_summary,
        on=['RBM', 'Staff'],
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
