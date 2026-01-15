import streamlit as st
import pandas as pd
import re
from io import BytesIO

# Page Configuration
st.set_page_config(page_title="SKU Folder Organizer", page_icon="📊")

st.title("📂 SKU Text-to-Excel Web App")
st.markdown("""
Upload all text files from your directory. The app will:
1. Identify the **Month** from the filename (e.g., `.11.` -> NOV).
2. Count **SKUs** to generate **QTY**.
3. Create an **Excel** with monthly sheets and a yearly summary.
""")

# 1. User Input for the Year
year_input = st.text_input("Enter the Year for this report:", "2025")

# 2. File Uploader (Accepts multiple files)
uploaded_files = st.file_uploader("Drop your .txt files here", type="txt", accept_multiple_files=True)

if uploaded_files:
    month_map = {
        "01": "JAN", "02": "FEB", "03": "MAR", "04": "APR", 
        "05": "MAY", "06": "JUN", "07": "JUL", "08": "AUG", 
        "09": "SEP", "10": "OCT", "11": "NOV", "12": "DEC"
    }
    
    all_data = [] # To store everything for the summary
    sheets_dict = {} # To store data grouped by month

    for file in uploaded_files:
        # Extract Month from Filename using Regex
        match = re.search(r'\.(\d{1,2})\.', file.name)
        if match:
            month_num = match.group(1).zfill(2)
            sheet_name = month_map.get(month_num, "Other")
        else:
            sheet_name = "Unsorted"

        # Read SKUs and calculate QTY
        content = file.read().decode("utf-8").splitlines()
        skus = [line.strip() for line in content if line.strip()]
        
        if skus:
            # Create a dataframe for this specific file
            df_file = pd.Series(skus).value_counts().reset_index()
            df_file.columns = ['SKU', 'QTY']
            df_file['Source File'] = file.name
            
            # Add to the month group
            if sheet_name not in sheets_dict:
                sheets_dict[sheet_name] = []
            sheets_dict[sheet_name].append(df_file)
            
            # Add to master list for summary
            all_data.append(df_file)

    if sheets_dict:
        # Prepare Excel in memory
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            
            # Create Monthly Sheets
            for month in sorted(sheets_dict.keys()):
                month_df = pd.concat(sheets_dict[month], ignore_index=True)
                month_df.to_excel(writer, sheet_name=month, index=False)
            
            # Create Annual Summary Sheet
            summary_df = pd.concat(all_data).groupby('SKU')['QTY'].sum().reset_index()
            summary_df = summary_df.sort_values(by='QTY', ascending=False)
            summary_df.to_excel(writer, sheet_name="ANNUAL SUMMARY", index=False)

        st.success(f"Processed {len(uploaded_files)} files successfully!")
        
        # Download Button
        st.download_button(
            label=f"📥 Download {year_input} Report",
            data=output.getvalue(),
            file_name=f"{year_input}_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("No valid SKU data found in the uploaded files.")