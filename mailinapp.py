
import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="Onsite Cases Report", layout="wide")
st.title("📋 Onsite Cases Report - GSX vs CRM vs Status")

st.markdown("Upload the following 3 Excel files:")
uploaded_gsx = st.file_uploader("Upload **GSX.xlsx**", type=["xlsx"])
uploaded_crm = st.file_uploader("Upload **CRM.xlsx**", type=["xlsx"])
uploaded_status = st.file_uploader("Upload **status.xlsx**", type=["xlsx"])

if uploaded_gsx and uploaded_crm and uploaded_status:
    try:
        gsx_df = pd.read_excel(uploaded_gsx)
        crm_df = pd.read_excel(uploaded_crm)
        status_df = pd.read_excel(uploaded_status)

        # Filter GSX for 'Onsite Service Facilitated'
        gsx_filtered = gsx_df[gsx_df['Repair Type'] == 'Onsite Service Facilitated']

        # Group GSX by 'Serial Number'
        m1 = gsx_filtered.groupby('Serial Number', as_index=False).agg({
            'Repair': 'first',
            'Repair Type': 'first',
            'Purchase Order': 'first',
            'Created Date': 'first',
            'Repair Status': 'first',
            'Part Number': 'first'
        }).rename(columns={
            'Repair': 'Repair_GSX',
            'Repair Type': 'Repair_Type_GSX',
            'Purchase Order': 'Purchase_Order_GSX',
            'Created Date': 'Created_Date_GSX',
            'Repair Status': 'Repair_Status_GSX',
            'Part Number': 'Part_Number_GSX'
        })

        # Group CRM by 'serial_number'
        m2 = crm_df.groupby('serial_number', as_index=False).agg({
            'reference_number': 'first',
            'created_at': 'first',
            'gsx_reference_number': 'first',
            'gsx_repair_type': 'first',
            'part_number': 'first'
        }).rename(columns={
            'reference_number': 'reference_number_DELVRY',
            'created_at': 'created_at_DELVRY',
            'gsx_reference_number': 'gsx_reference_number_DELVRY',
            'gsx_repair_type': 'gsx_repair_type_DELVRY',
            'part_number': 'part_number_DELVRY'
        })

        # Group status by 'serial_number'
        m3 = status_df.groupby('serial_number', as_index=False).agg({
            'reference_number': 'first',
            'created_at': 'first'
        }).rename(columns={
            'reference_number': 'reference_number_status',
            'created_at': 'created_at_status'
        })

        # Merge CRM into GSX on Serial Number (left merge)
        merged = m1.merge(m2, left_on='Serial Number', right_on='serial_number', how='left')

        # Merge status into previous merged result
        merged = merged.merge(m3, left_on='Serial Number', right_on='serial_number', how='left')

        # Drop redundant merge columns
        merged.drop(columns=['serial_number_x', 'serial_number_y'], errors='ignore', inplace=True)

        st.success("✅ Merged Successfully!")
        st.dataframe(merged)

        # Export to Excel in memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            merged.to_excel(writer, sheet_name='Onsite Cases', index=False)
        output.seek(0)

        st.download_button(
            label="📥 Download Report as Excel",
            data=output,
            file_name="onsite_cases.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ An error occurred: {e}")
else:
    st.info("⬆️ Please upload all 3 files to generate report.")
