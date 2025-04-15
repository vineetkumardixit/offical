
import pandas as pd
from pathlib import Path
import streamlit as st

st.set_page_config(page_title="📦 Mail-in Tracker", layout="wide")
st.title("📦 Mail-in Tracker - GSX Update")

# Define the status mapping
status_map = {
    'Unit Returned Replaced': 'Whole Unit replacement',
    'Unit to Be Replaced': 'Whole Unit replacement',
    'Unit Repaired': 'Part Replacement',
    'Unit Returned Repaired': 'Part Replacement',
    'Unit Abused - No Requote': 'Unauthorized Modifications',
    'Unit Returned Could not Duplicate Failure': 'NTF',
    'Requote Fixed Rate': 'Requote',
    'Requote Declined': 'Requote',
    'Requote Abuse Tier 4': 'Requote',
    'Estimation': 'Estimation',
    'Service Decline': 'Service Decline'
}

uploaded_gsx = st.file_uploader("Upload GSX Excel File", type=["xlsx"])

if uploaded_gsx:
    try:
        gsx_df = pd.read_excel(uploaded_gsx)

        # Save original Repair Status for extracting Requote descriptions
        gsx_df['Original Status'] = gsx_df['Repair Status']

        # Map to cleaned-up Repair Status for Mail_in_tracker
        gsx_df['Repair Status'] = gsx_df['Original Status'].map(lambda x: status_map.get(x, 'under process'))

        # Extract Requote description (everything after "Requote")
        def extract_requote_detail(status):
            if isinstance(status, str) and status.startswith('Requote'):
                return status.replace('Requote', '').strip()
            return ''

        gsx_df['Requote description'] = gsx_df['Original Status'].apply(extract_requote_detail)

        tracker_columns = [
            'Service Provider Ship-To',
            'Created Date',
            'Repair',
            'Purchase Order',
            'Product Name',
            'Repair Status',
            'Requote description'
        ]

        gsx_df = gsx_df[tracker_columns]

        # Filepath for saving
        tracker_path = Path.home() / "Desktop" / "Mail_in_tracker.xlsx"

        if tracker_path.exists():
            tracker_df = pd.read_excel(tracker_path)
        else:
            tracker_df = pd.DataFrame(columns=tracker_columns)

        tracker_df.set_index('Repair', inplace=True, drop=False)
        gsx_df.set_index('Repair', inplace=True, drop=False)

        for repair_id, row in gsx_df.iterrows():
            if repair_id in tracker_df.index:
                for col in tracker_columns:
                    if col != 'Repair':
                        tracker_df.at[repair_id, col] = row[col]
            else:
                tracker_df = pd.concat([tracker_df, pd.DataFrame([row], columns=tracker_columns)], ignore_index=False)

        tracker_df.reset_index(drop=True, inplace=True)

        # Export to BytesIO for download
        from io import BytesIO
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            tracker_df.to_excel(writer, index=False, sheet_name='Mail-in Tracker')
        output.seek(0)

        st.success("✅ Mail_in_tracker updated successfully!")
        st.download_button(
            label="📥 Download Updated Tracker",
            data=output,
            file_name="Mail_in_tracker.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.dataframe(tracker_df)

    except Exception as e:
        st.error(f"❌ Error: {e}")

else:
    st.info("📤 Please upload a GSX Excel file to proceed.")
