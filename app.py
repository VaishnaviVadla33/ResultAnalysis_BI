import streamlit as st
import pandas as pd
import io
import re

def split_subjects_into_dataframes(df):
    sm_indices = [i for i, col in enumerate(df.columns) 
                  if col.startswith('SM_') and 'Module id' not in col]
    
    dataframes = {}
    
    for start_idx in sm_indices:
        for i in range(start_idx, len(df.columns)):
            if 'Attempted Credit Value' in df.columns[i]:
                end_idx = i + 1
                break
        
        subject_df = df.iloc[:, start_idx:end_idx]
        df_name = str(subject_df.iloc[:, 0].dropna().iloc[0])
        df_name = re.sub(r'[^a-zA-Z0-9_]', '_', df_name)
        df_name = re.sub(r'_{2,}', '_', df_name)
        df_name = df_name.strip('_')
        
        subject_df.columns = [
            col if not (col.startswith('SM_') and 'Module' not in col) else "Subject Name"
            for col in subject_df.columns
        ]
        
        dataframes[df_name] = subject_df
    
    return dataframes

def create_enhanced_subject_dataframes(subject_dfs, new_df):
    enhanced_dataframes = {}
    student_info = new_df[['Student Number', 'Student Name']]
    
    for subject_name, subject_df in subject_dfs.items():
        new_name = f"{subject_name}_01"
        enhanced_df = pd.concat([student_info, subject_df], axis=1)
        enhanced_dataframes[new_name] = enhanced_df
    
    return enhanced_dataframes

def clean_subject_dataframes(enhanced_subject_dfs):
    cleaned_dataframes = {}
    
    for subject_name, df in enhanced_subject_dfs.items():
        cleaned_df = df[['Student Number', 'Student Name']].copy()
        
        if 'Subject Name' in df.columns:
            cleaned_df['Subject Name'] = df['Subject Name']
            
        columns_to_find = {
            'SM_Module id': ['SM_Module id'],
            'Final Grade(N200)': ['Final Grade(N200)', 'Final Grade'],
            'Final Marks(100 )': ['Final Marks(50 )', 'Final Marks(100 )', 'Total Marks(50 )'],
            'ICA Total(50 )': ['ICA Total(50 )', 'ICA Total(100 )'],
            'Term End Examination(100 )': ['Term End Examination(50 )', 'Term End Examination(100 )']
        }
        
        for new_col_name, possible_names in columns_to_find.items():
            existing_cols = [col for col in df.columns if any(name in col for name in possible_names)]
            if existing_cols:
                matched_col = existing_cols[0]
                cleaned_df[new_col_name] = df[matched_col]
            else:
                cleaned_df[new_col_name] = None
                
        new_name = f"{subject_name.replace('_01', '')}_02"
        cleaned_dataframes[new_name] = cleaned_df
        
    return cleaned_dataframes

def remove_null_subject_rows(subject_dataframes):
    filtered_dataframes = {}
    
    for subject_name, df in subject_dataframes.items():
        filtered_df = df.copy()
        
        if 'Subject Name' in filtered_df.columns:
            filtered_df = filtered_df.dropna(subset=['Subject Name'])
            new_name = f"{subject_name.replace('_02', '')}_03"
            filtered_dataframes[new_name] = filtered_df
        else:
            filtered_dataframes[subject_name] = df
    
    return filtered_dataframes

def standardize_columns(df, columns_mapping):
    standardized_df = pd.DataFrame()
    
    standardized_df['Student Number'] = df['Student Number']
    standardized_df['Student Name'] = df['Student Name']
    
    if 'Subject Name' in df.columns:
        standardized_df['Subject Name'] = df['Subject Name']
    
    numeric_columns = [
        'Final Marks(100 )', 
        'ICA Total(50 )', 
        'Term End Examination(100 )'
    ]
    
    for new_col_name, possible_names in columns_mapping.items():
        existing_cols = [col for col in df.columns if any(name in col for name in possible_names)]
        if existing_cols:
            standardized_df[new_col_name] = df[existing_cols[0]]
            
            if new_col_name in numeric_columns:
                standardized_df[new_col_name] = pd.to_numeric(standardized_df[new_col_name], errors='coerce').fillna(0)
    
    return standardized_df

def process_excel_file(uploaded_file):
    # Read the uploaded Excel file
    df = pd.read_excel(uploaded_file)
    
    # Extract main columns
    columns_to_extract = [
        'Student Number', 'Student Name', 'Additional ID', 'Gender', 
        'Program', 'Campus', 'Total', 'Aggregate', 'SGPA', 
        'CGPA', 'Percentage', 'Result', 'Status', 'Current F'
    ]
    new_df = df[columns_to_extract]
    
    # Process the data through all functions
    subject_dfs = split_subjects_into_dataframes(df)
    enhanced_subject_dfs = create_enhanced_subject_dataframes(subject_dfs, new_df)
    cleaned_subject_dfs = clean_subject_dataframes(enhanced_subject_dfs)
    filtered_subject_dfs = remove_null_subject_rows(cleaned_subject_dfs)
    
    # Define column mapping for final standardization
    columns_to_find = {
        'SM_Module id': ['SM_Module id'],
        'Final Grade(N200)': ['Final Grade(N200)', 'Final Grade'],
        'Final Marks(100 )': ['Final Marks', 'Final Marks(50 )', 'Final Marks(100 )'],
        'ICA Total(50 )': ['ICA Total', 'ICA Total(50 )', 'ICA Total(100 )'],
        'Term End Examination(100 )': ['Term End Examination', 'Term End Examination(50 )', 'Term End Examination(100 )']
    }
    
    # Create combined sheet
    all_subjects_combined = pd.concat(
        [standardize_columns(df.copy(), columns_to_find) for df in filtered_subject_dfs.values()],
        ignore_index=True
    )
    
    # Create output Excel file in memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        new_df.to_excel(writer, sheet_name='MainTable', index=False)
        
        for sheet_name, dataframe in filtered_subject_dfs.items():
            dataframe.to_excel(writer, sheet_name=sheet_name, index=False)
        
        all_subjects_combined.to_excel(writer, sheet_name='AllInOne', index=False)
    
    output.seek(0)
    return output


def main():
    # Page configuration
    st.set_page_config(
        page_title="Excel Data Processor",
        page_icon="📊",
        layout="wide"
    )

    # Custom CSS for better styling
    st.markdown("""
        <style>
        .main {
            padding: 2rem;
        }
        .stButton>button {
            width: 100%;
            background-color: #4CAF50;
            color: white;
            padding: 0.5rem;
            margin: 0.5rem 0;
        }
        .stDownloadButton>button {
            width: 100%;
            background-color: #008CBA;
            color: white;
            padding: 0.5rem;
            margin: 0.5rem 0;
        }
        .success-message {
            padding: 1rem;
            background-color: #DFF0D8;
            border-radius: 4px;
            margin: 1rem 0;
        }
        .error-message {
            padding: 1rem;
            background-color: #F2DEDE;
            border-radius: 4px;
            margin: 1rem 0;
        }
        </style>
    """, unsafe_allow_html=True)

    # Header section with improved layout
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.title("📊 Excel Data Processor")
        st.markdown("---")

    # Main container
    with st.container():
        # Instructions in an expander
        with st.expander("📖 Instructions", expanded=True):
            st.markdown("""
            ### How to Use:
            1. **Upload Your File**: Select an Excel file (.xlsx) containing student data
            2. **Process Data**: Click the 'Process Data' button to generate:
                * MainTable - Basic student information
                * Subject Sheets - Individual subject data
                * AllInOne Sheet - Combined subject information
            3. **Download**: Get your processed Excel file
            """)

        # File upload section
        st.markdown("### 📤 Upload Your Excel File")
        uploaded_file = st.file_uploader("", type="xlsx", help="Upload an Excel file (.xlsx)")

        # Create two columns for processing and status
        if uploaded_file is not None:
            st.success("✅ File uploaded successfully!")
            
            col1, col2 = st.columns(2)
            with col1:
                process_button = st.button("🔄 Process Data", key="process")
            
            if process_button:
                try:
                    with st.spinner("🔄 Processing your data..."):
                        processed_file = process_excel_file(uploaded_file)
                    
                    st.success("✨ Processing completed successfully!")
                    
                    # Download section with better styling
                    st.markdown("### 📥 Download Processed File")
                    st.download_button(
                        label="📎 Download Processed Excel",
                        data=processed_file,
                        file_name="processed_data.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    # Add a summary of what was created
                    st.info("""
                    #### Files Created:
                    - ✅ MainTable with student information
                    - ✅ Individual subject sheets
                    - ✅ Combined AllInOne sheet
                    """)
                    
                except Exception as e:
                    st.error(f"❌ An error occurred: {str(e)}")
                    st.markdown("""
                    Please check that:
                    - Your Excel file is not corrupted
                    - The file contains all required columns
                    - The file format matches the expected structure
                    """)

    # Footer
    st.markdown("---")
    st.markdown("""
        <div style='text-align: center; color: #666;'>
        Made with ❤️ for easy data processing
        </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
