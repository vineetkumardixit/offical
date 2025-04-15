import streamlit as st
import pandas as pd
import numpy as np
import io

st.set_page_config(page_title="Mail-in Tracker", layout="wide")
st.title("📦 Mail-in Tracker Report Generator")

#!/usr/bin/env python
# coding: utf-8

# In[6]:


import pandas as pd
import os
from pathlib import Path

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

# Desktop path
desktop = Path.home() / "Desktop"
tracker_path = desktop / "Mail_in_tracker.xlsx"

# Load GSX file
gsx_path = "/Users/vineetdixit/Desktop/audit123/gsx.xlsx"  # replace with actual path
gsx_df = pd.read_excel(gsx_path)

# Rename GSX columns if needed to match
gsx_df = gsx_df.rename(columns={
    'Ship-To': 'Service Provider Ship-To',
    'Created': 'Created Date',
    'Repair ID': 'Repair',
    'PO Number': 'Purchase Order',
    'Product': 'Product Name',
    'Status': 'Repair Status'
})

# Map GSX statuses to Mail_in_tracker statuses
gsx_df['Repair Status'] = gsx_df['Repair Status'].map(lambda x: status_map.get(x, 'under process'))


# Create 'Requote Y/N' column
def extract_requote_detail(status):
    if isinstance(status, str) and status.startswith('Requote'):
        return status.replace('Requote', '').strip()
    return ''

gsx_df['Requote desciption'] = gsx_df['Repair Status'].apply(extract_requote_detail)

tracker_columns = [
    'Service Provider Ship-To',
    'Created Date',
    'Repair',
    'Purchase Order',
    'Product Name',
    'Repair Status',
    'Requote desciption'
]

# Load or create Mail_in_tracker
if tracker_path.exists():
    tracker_df = pd.read_excel(tracker_path)
else:
    tracker_df = pd.DataFrame(columns=tracker_columns)

# Index by 'Repair' for easier comparison
tracker_df.set_index('Repair', inplace=True, drop=False)
gsx_df.set_index('Repair', inplace=True, drop=False)

# Update existing or append new rows
for repair_id, row in gsx_df.iterrows():
    if repair_id in tracker_df.index:
        for col in tracker_columns:
            if col != 'Repair':  # 'Repair' is index, skip it
                tracker_df.at[repair_id, col] = row[col]
    else:
        tracker_df = pd.concat([tracker_df, pd.DataFrame([row], columns=tracker_columns)], ignore_index=False)

# Reset index and save
tracker_df.reset_index(drop=True, inplace=True)
tracker_df.to_excel(tracker_path, index=False)

st.write("Mail_in_tracker updated successfully.")


# In[24]:


import pandas as pd
from pathlib import Path

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

# Paths
desktop = Path.home() / "Desktop"
tracker_path = desktop / "Mail_in_tracker.xlsx"
gsx_path = "/Users/vineetdixit/Desktop/audit123/gsx.xlsx"  # update if needed

# Load GSX file
gsx_df = pd.read_excel(gsx_path)

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

# Rename columns for consistency
gsx_df = gsx_df.rename(columns={
    'Created Date': 'Created Date',
    'Service Provider Ship-To': 'Service Provider Ship-To',
    'Repair': 'Repair',
    'Purchase Order': 'Purchase Order',
    'Product Name': 'Product Name'
})

# Keep only necessary columns
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

# Load or create Mail_in_tracker
if tracker_path.exists():
    tracker_df = pd.read_excel(tracker_path)
else:
    tracker_df = pd.DataFrame(columns=tracker_columns)

# Index by 'Repair' for comparison
tracker_df.set_index('Repair', inplace=True, drop=False)
gsx_df.set_index('Repair', inplace=True, drop=False)

# Update or append rows
for repair_id, row in gsx_df.iterrows():
    if repair_id in tracker_df.index:
        for col in tracker_columns:
            if col != 'Repair':
                tracker_df.at[repair_id, col] = row[col]
    else:
        tracker_df = pd.concat([tracker_df, pd.DataFrame([row], columns=tracker_columns)], ignore_index=False)

# Reset index and save
tracker_df.reset_index(drop=True, inplace=True)
tracker_df.to_excel(tracker_path, index=False)

st.write("✅ Mail_in_tracker updated successfully.")


# In[ ]:





    # Add download button
    if 'df' in locals():
        output = io.BytesIO()
        df.to_excel(output, index=False, engine='xlsxwriter')
        st.download_button("📥 Download Excel Report", output.getvalue(), file_name="mail_in_tracker_output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    