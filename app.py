import streamlit as st
import pandas as pd


# Function to extract and process relevant columns from the Excel file
def extract_summary(excel_file):
    excel_data = pd.ExcelFile(excel_file)
    summary = []

    for sheet in excel_data.sheet_names:
        data = excel_data.parse(sheet)

        # Extract only relevant rows and columns (KASIR, STRUK, SELISIH)
        relevant_data = data[['Unnamed: 1', 'Unnamed: 11', 'Unnamed: 12']].iloc[3:]
        relevant_data.columns = ['KASIR', 'SELISIH', 'STRUK']
        relevant_data['Sheet'] = sheet

        # Filter out rows where 'KASIR', 'SELISIH', or 'STRUK' are null
        filtered_data = relevant_data.dropna(subset=['KASIR', 'SELISIH', 'STRUK'])

        summary.append(filtered_data)

    # Concatenate all summaries into one dataframe
    summary_df = pd.concat(summary, ignore_index=True)

    # Convert SELISIH and STRUK to numeric for calculations
    summary_df['SELISIH'] = pd.to_numeric(summary_df['SELISIH'], errors='coerce')
    summary_df['STRUK'] = pd.to_numeric(summary_df['STRUK'], errors='coerce')

    # Group by KASIR and sum the STRUK values
    summary_grouped = summary_df.groupby('KASIR').agg(
        total_struk=pd.NamedAgg(column='STRUK', aggfunc='sum'),
        negative_selisih_count=pd.NamedAgg(column='SELISIH', aggfunc=lambda x: (x < 0).sum())
    ).reset_index()

    return summary_grouped


# Streamlit interface
st.title('Excel Data Extractor for Kasir, Selisih, and Struk')

# File uploader
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

if uploaded_file is not None:
    # Process the uploaded file
    summary_grouped_df = extract_summary(uploaded_file)

    # Display the processed summary
    st.write("Summary of Struk Sum and Negative Selisih Count:")
    st.dataframe(summary_grouped_df)

    # Option to download the processed data as CSV
    csv = summary_grouped_df.to_csv(index=False)
    st.download_button(label="Download CSV", data=csv, file_name='summary_grouped.csv', mime='text/csv')