import streamlit as st
import pandas as pd
import os
from pathlib import Path
import zipfile
from io import BytesIO
import re

# Page config with custom theme
st.set_page_config(
    page_title="Data File Collator", 
    page_icon="📊", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# Modern Custom CSS
st.markdown("""
<style>
    /* Import Google Fonts */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    
    /* Global Styles */
    * {
        font-family: 'Inter', sans-serif;
    }
    
    /* Main container styling */
    .main {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
    }
    
    /* Content card */
    .block-container {
        background: white;
        border-radius: 20px;
        padding: 3rem 2rem;
        box-shadow: 0 20px 60px rgba(0,0,0,0.15);
        max-width: 1400px;
        margin: 0 auto;
    }
    
    /* Title styling */
    h1 {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        font-weight: 700;
        font-size: 3rem !important;
        margin-bottom: 0.5rem !important;
        letter-spacing: -0.5px;
    }
    
    /* Subtitle */
    .subtitle {
        color: #64748b;
        font-size: 1.1rem;
        margin-bottom: 2rem;
        font-weight: 500;
    }
    
    /* Headers */
    h2, h3 {
        color: #1e293b;
        font-weight: 600;
        margin-top: 2rem !important;
    }
    
    /* Buttons */
    .stButton > button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        border-radius: 12px;
        padding: 0.75rem 2rem;
        font-weight: 600;
        font-size: 1rem;
        transition: all 0.3s ease;
        box-shadow: 0 4px 15px rgba(102, 126, 234, 0.4);
    }
    
    .stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(102, 126, 234, 0.6);
    }
    
    /* File uploader */
    .stFileUploader {
        border: 2px dashed #cbd5e1;
        border-radius: 12px;
        padding: 2rem;
        background: #f8fafc;
        transition: all 0.3s ease;
    }
    
    .stFileUploader:hover {
        border-color: #667eea;
        background: #f1f5ff;
    }
    
    /* Multiselect */
    .stMultiSelect > div > div {
        border-radius: 10px;
        border: 2px solid #e2e8f0;
    }
    
    /* Text input */
    .stTextInput > div > div > input {
        border-radius: 10px;
        border: 2px solid #e2e8f0;
        padding: 0.75rem;
    }
    
    .stTextInput > div > div > input:focus {
        border-color: #667eea;
        box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.1);
    }
    
    /* Success/Info/Warning boxes */
    .stSuccess {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%);
        color: white;
        border-radius: 12px;
        padding: 1rem 1.5rem;
        border: none;
        box-shadow: 0 4px 15px rgba(16, 185, 129, 0.3);
    }
    
    .stInfo {
        background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%);
        color: white;
        border-radius: 12px;
        padding: 1rem 1.5rem;
        border: none;
        box-shadow: 0 4px 15px rgba(59, 130, 246, 0.3);
    }
    
    .stWarning {
        background: linear-gradient(135deg, #f59e0b 0%, #d97706 100%);
        color: white;
        border-radius: 12px;
        padding: 1rem 1.5rem;
        border: none;
        box-shadow: 0 4px 15px rgba(245, 158, 11, 0.3);
    }
    
    /* Metrics */
    [data-testid="stMetricValue"] {
        font-size: 2rem;
        font-weight: 700;
        color: #667eea;
    }
    
    [data-testid="stMetricLabel"] {
        font-weight: 600;
        color: #64748b;
        font-size: 0.9rem;
    }
    
    /* Dataframe */
    .stDataFrame {
        border-radius: 12px;
        overflow: hidden;
        box-shadow: 0 4px 15px rgba(0,0,0,0.05);
    }
    
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 1rem;
        background: #f8fafc;
        padding: 0.5rem;
        border-radius: 12px;
    }
    
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        padding: 0.75rem 1.5rem;
        font-weight: 600;
        color: #64748b;
    }
    
    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
    }
    
    /* Expander */
    .streamlit-expanderHeader {
        background: #f8fafc;
        border-radius: 12px;
        font-weight: 600;
        color: #1e293b;
        padding: 1rem;
    }
    
    .streamlit-expanderHeader:hover {
        background: #f1f5ff;
        color: #667eea;
    }
    
    /* Divider */
    hr {
        margin: 2rem 0;
        border: none;
        height: 2px;
        background: linear-gradient(90deg, transparent, #e2e8f0, transparent);
    }
    
    /* Sidebar */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #667eea 0%, #764ba2 100%);
        padding: 2rem 1rem;
    }
    
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] p,
    [data-testid="stSidebar"] li {
        color: white !important;
    }
    
    /* Radio buttons */
    .stRadio > label {
        font-weight: 600;
        color: #1e293b;
    }
    
    /* Checkbox */
    .stCheckbox > label {
        font-weight: 500;
        color: #475569;
    }
    
    /* Number input */
    .stNumberInput > div > div > input {
        border-radius: 10px;
        border: 2px solid #e2e8f0;
    }
    
    /* Download button */
    .stDownloadButton > button {
        background: linear-gradient(135deg, #10b981 0%, #059669 100%);
        color: white;
        border: none;
        border-radius: 12px;
        padding: 0.75rem 2rem;
        font-weight: 600;
        box-shadow: 0 4px 15px rgba(16, 185, 129, 0.4);
    }
    
    .stDownloadButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(16, 185, 129, 0.6);
    }
    
    /* Progress bar */
    .stProgress > div > div {
        background: #e2e8f0 !important;
        border-radius: 10px;
    }
    
    .stProgress > div > div > div {
        background: #10b981 !important;
        border-radius: 10px;
    }
    
    /* Card-like sections */
    .card {
        background: #f8fafc;
        border-radius: 16px;
        padding: 1.5rem;
        margin: 1rem 0;
        border: 1px solid #e2e8f0;
        transition: all 0.3s ease;
    }
    
    .card:hover {
        box-shadow: 0 8px 25px rgba(0,0,0,0.08);
        transform: translateY(-2px);
    }
    
    /* Feature badges */
    .badge {
        display: inline-block;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 0.25rem 0.75rem;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 600;
        margin: 0.25rem;
    }
</style>
""", unsafe_allow_html=True)

# Header with subtitle
st.title("Welcome to Data File Collator")
st.markdown('<p class="subtitle">Upload multiple data files (Excel, CSV) to merge them into a single dataset with dynamic header matching.<br>A powerful tool to merge multiple data files into a single, unified dataset with intelligent header matching.</p>', unsafe_allow_html=True)

# File type selector with modern styling
st.markdown("### 🎯 Configuration")

col1, col2 = st.columns([2, 3])

with col1:
    file_types = st.multiselect(
        "Select file types to accept:",
        options=['csv', 'xlsx', 'xls', 'xlsm', 'xlsb'],
        help="Choose which file formats you want to upload"
    )

with col2:
    if not file_types:
        st.warning("⚠️ Please select at least one file type")

# Excel sheet options (only show if non-CSV types are selected)
has_excel_types = any(ft in file_types for ft in ['xlsx', 'xls', 'xlsm', 'xlsb'])

if has_excel_types:
    st.markdown("### 📑 Excel Sheet Options")
    
    col1, col2 = st.columns(2)
    
    with col1:
        sheet_mode = st.radio(
            "How to handle Excel sheets:",
            options=['first_sheet', 'all_sheets', 'by_name'],
            format_func=lambda x: {
                'first_sheet': '📄 First Sheet Only',
                'all_sheets': '📚 All Sheets',
                'by_name': '🏷️ Specific Sheet by Name'
            }[x],
            help="Choose how to read data from Excel files"
        )
    
    with col2:
        if sheet_mode == 'by_name':
            sheet_name = st.text_input(
                "Sheet name to collate:",
                value="Sheet1",
                help="Enter the exact name of the sheet to read from all Excel files"
            )
        else:
            sheet_name = None
else:
    sheet_mode = 'first_sheet'
    sheet_name = None

st.markdown("---")

# File uploader
st.markdown("### 📤 Upload Files")

if file_types:
    uploaded_files = st.file_uploader(
        f"Choose files ({', '.join(file_types).upper()})",
        type=file_types,
        accept_multiple_files=True,
        help="Select one or more files to collate"
    )
else:
    uploaded_files = None
    st.info("👆 Select file types above to enable file upload")

if uploaded_files:
    st.success(f"✅ {len(uploaded_files)} file(s) uploaded successfully")
    
    # Check if we have processed data in session state
    if 'master_df' in st.session_state:
        master_df = st.session_state['master_df']
        log_df = st.session_state['log_df']
        files_processed = st.session_state['files_processed']
        sheets_processed = st.session_state.get('sheets_processed', 0)
        total_rows_added = st.session_state['total_rows_added']
        
        # Display summary with metrics
        st.markdown("### ✅ Processing Complete!")
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Files Processed", files_processed)
        with col2:
            st.metric("Sheets Processed", sheets_processed if has_excel_types else 'N/A')
        with col3:
            st.metric("Total Rows", f"{total_rows_added:,}")
        with col4:
            st.metric("Total Columns", len(master_df.columns))
        
        # Button to reset and start over
        if st.button("🔄 Process New Files"):
            # Clear session state
            for key in ['master_df', 'log_df', 'files_processed', 'sheets_processed', 'total_rows_added', 'column_configs', 'num_configs']:
                if key in st.session_state:
                    del st.session_state[key]
            st.rerun()
        
        # Feature: Create new columns from existing columns
        st.markdown("---")
        st.markdown("### 🔧 Create New Columns (Optional)")
        
        create_columns = st.checkbox("➕ Add new columns by combining existing columns", value=False)
        
        if create_columns:
            st.info("💡 Configure all your new columns below, then click 'Apply Column Transformations' to create them all at once")
            
            # Number of column configurations to create
            num_configs = st.number_input(
                "How many new columns do you want to create?",
                min_value=1,
                max_value=20,
                value=st.session_state.get('num_configs', 1),
                step=1,
                help="Choose the number of combined columns you want to create",
                key="num_configs_input"
            )
            
            # Store the number in session state
            st.session_state['num_configs'] = num_configs
            
            # Initialize column_configs based on num_configs
            if 'column_configs' not in st.session_state:
                st.session_state['column_configs'] = []
            
            # Adjust the list size to match num_configs
            current_count = len(st.session_state['column_configs'])
            if num_configs > current_count:
                # Add new configs
                for _ in range(num_configs - current_count):
                    st.session_state['column_configs'].append({
                        'new_col_name': '',
                        'source_cols': [],
                        'delimiter_type': 'custom',
                        'custom_delimiter': ', ',
                        'skip_empty': True
                    })
            elif num_configs < current_count:
                # Remove excess configs
                st.session_state['column_configs'] = st.session_state['column_configs'][:num_configs]
            
            st.markdown("---")
            
            # Display column configurations (no auto-processing, just collection)
            for idx in range(num_configs):
                config = st.session_state['column_configs'][idx]
                
                with st.expander(f"🔹 Column Configuration {idx + 1}", expanded=True):
                    config['new_col_name'] = st.text_input(
                        "New Column Name:",
                        value=config.get('new_col_name', ''),
                        key=f"new_col_name_{idx}",
                        placeholder="e.g., Full_Address"
                    )
                    
                    # Initialize selected columns if not present
                    if 'source_cols' not in config or config['source_cols'] is None:
                        config['source_cols'] = []
                    
                    # Show multiselect to add columns (just stores, doesn't process)
                    available_columns = list(master_df.columns)
                    
                    config['source_cols'] = st.multiselect(
                        "Select columns to combine:",
                        options=available_columns,
                        key=f"select_cols_{idx}",
                        help="Select columns in the order you want to combine them. No processing happens until you click Apply."
                    )
                    
                    # Display selected columns in order with visual feedback
                    if config['source_cols']:
                        st.markdown(f"**✓ {len(config['source_cols'])} Column(s) Selected:** " + " → ".join([f"`{col}`" for col in config['source_cols']]))
                    
                    st.markdown("---")
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        config['delimiter_type'] = st.radio(
                            "Delimiter Type:",
                            options=['newline', 'custom'],
                            index=0 if config.get('delimiter_type', 'custom') == 'newline' else 1,
                            format_func=lambda x: '↵ Newline (multi-line cell)' if x == 'newline' else '✏️ Custom Delimiter',
                            key=f"delimiter_type_{idx}",
                            horizontal=True
                        )
                    
                    with col2:
                        if config['delimiter_type'] == 'custom':
                            config['custom_delimiter'] = st.text_input(
                                "Custom Delimiter:",
                                value=config.get('custom_delimiter', ', '),
                                key=f"custom_delimiter_{idx}",
                                placeholder="e.g., ', ' or ' | ' or ' - '"
                            )
                    
                    # Add option to skip empty cells
                    config['skip_empty'] = st.checkbox(
                        "Skip empty/null values",
                        value=config.get('skip_empty', True),
                        key=f"skip_empty_{idx}",
                        help="When enabled, empty cells will be excluded from the combination (recommended for cleaner output)"
                    )
            
            st.markdown("---")
            
            # Apply button - ALL PROCESSING HAPPENS HERE
            col1, col2, col3 = st.columns([2, 1, 2])
            with col2:
                apply_button = st.button("✨ Apply Column Transformations", type="primary", use_container_width=True)
            
            if apply_button:
                if not st.session_state['column_configs']:
                    st.warning("⚠️ Please add at least one column configuration")
                else:
                    # Create a copy of master_df to modify
                    modified_df = master_df.copy()
                    warnings = []
                    success_messages = []
                    
                    with st.spinner("Processing column transformations..."):
                        for config in st.session_state['column_configs']:
                            new_col = config['new_col_name'].strip()
                            source_cols = config['source_cols']
                            delimiter = '\n' if config['delimiter_type'] == 'newline' else config.get('custom_delimiter', ', ')
                            skip_empty = config.get('skip_empty', True)
                            
                            if not new_col:
                                warnings.append("⚠️ Skipped a configuration with empty column name")
                                continue
                            
                            if not source_cols:
                                warnings.append(f"⚠️ Skipped '{new_col}': No source columns selected")
                                continue
                            
                            # Check which columns exist
                            existing_cols = [col for col in source_cols if col in modified_df.columns]
                            missing_cols = [col for col in source_cols if col not in modified_df.columns]
                            
                            if missing_cols:
                                warnings.append(f"⚠️ Column '{new_col}': The following source columns were not found: {', '.join(missing_cols)}")
                            
                            if not existing_cols:
                                warnings.append(f"❌ Column '{new_col}': None of the source columns exist in the dataset")
                                continue
                            
                            # Create the new column by combining existing columns
                            def combine_values(row):
                                values = []
                                for col in existing_cols:
                                    val = row[col]
                                    # Handle empty values based on skip_empty setting
                                    if skip_empty:
                                        # Skip null/empty values
                                        if pd.notna(val) and str(val).strip() != '':
                                            values.append(str(val))
                                    else:
                                        # Include null/empty values as empty strings
                                        if pd.isna(val):
                                            values.append('')
                                        else:
                                            values.append(str(val))
                                return delimiter.join(values)
                            
                            modified_df[new_col] = modified_df.apply(combine_values, axis=1)
                            skip_msg = " (skipping empty values)" if skip_empty else " (including empty values)"
                            success_messages.append(f"✅ Created column '{new_col}' from {len(existing_cols)} source column(s){skip_msg}")
                    
                    # Update master_df in session state
                    st.session_state['master_df'] = modified_df
                    master_df = modified_df
                    
                    # Display results
                    if success_messages:
                        for msg in success_messages:
                            st.success(msg)
                    
                    if warnings:
                        for warning in warnings:
                            st.warning(warning)
                    
                    # Force a rerun to update the display
                    st.rerun()
        
        st.markdown("---")
        
        # NEW FEATURE: Generate Apify Time-Stamp Column
        st.markdown("### ⏰ Generate Apify Time-Stamp (Optional)")
        
        generate_apify_timestamp = st.checkbox("🔖 Generate Apify Time-Stamp", value=False, 
                                                help="Creates a special column by combining specific EXPERIENCE fields")
        
        if generate_apify_timestamp:
            st.info("💡 This will create a new column called 'Apify Time-Stamp' by combining specific EXPERIENCE fields with newline delimiters")
            
            # Define required columns for Apify Time-Stamp
            # NOTE: Change the order of columns below if you need different concatenation order in future
            apify_required_columns = [
                'EXPERIENCE/0/title',
                'EXPERIENCE/0/subtitle',
                'EXPERIENCE/0/meta',
                'EXPERIENCE/0/child/0/title',
                'EXPERIENCE/0/child/0/subtitle',
                'EXPERIENCE/0/child/0/caption',
                'EXPERIENCE/0/caption'
            ]
            
            # Check which columns exist and which are missing
            existing_apify_cols = [col for col in apify_required_columns if col in master_df.columns]
            missing_apify_cols = [col for col in apify_required_columns if col not in master_df.columns]
            
            # Display status
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Available Columns", len(existing_apify_cols))
            with col2:
                st.metric("Missing Columns", len(missing_apify_cols))
            
            # Show which columns are available
            if existing_apify_cols:
                with st.expander("✅ Available Columns", expanded=False):
                    for col in existing_apify_cols:
                        st.markdown(f"- `{col}`")
            
            # Show warning if any columns are missing
            if missing_apify_cols:
                st.warning(f"⚠️ **Required columns not available:** The following {len(missing_apify_cols)} column(s) are missing from your dataset:\n\n" + 
                          "\n".join([f"- `{col}`" for col in missing_apify_cols]) +
                          f"\n\n**Would you like to proceed anyway?** The 'Apify Time-Stamp' column will be created using only the {len(existing_apify_cols)} available column(s).")
            
            # Generate button
            col1, col2, col3 = st.columns([2, 1, 2])
            with col2:
                generate_apify_button = st.button("⚡ Generate Apify Time-Stamp", type="primary", use_container_width=True)
            
            if generate_apify_button:
                if not existing_apify_cols:
                    st.error("❌ Cannot generate 'Apify Time-Stamp': None of the required columns exist in the dataset!")
                else:
                    # Create a copy of master_df to modify
                    modified_df = master_df.copy()
                    
                    with st.spinner("Generating Apify Time-Stamp column..."):
                        # Create the Apify Time-Stamp column by combining available columns
                        def create_apify_timestamp(row):
                            values = []
                            for col in existing_apify_cols:
                                val = row[col]
                                # Skip null/empty values (same as existing feature behavior)
                                if pd.notna(val) and str(val).strip() != '':
                                    values.append(str(val))
                            return '\n'.join(values)  # Join with newline delimiter
                        
                        modified_df['Apify Time-Stamp'] = modified_df.apply(create_apify_timestamp, axis=1)
                    
                    # Update master_df in session state
                    st.session_state['master_df'] = modified_df
                    master_df = modified_df
                    
                    # Display success message
                    st.success(f"✅ Successfully created 'Apify Time-Stamp' column using {len(existing_apify_cols)} available column(s) (skipping empty values)")
                    
                    if missing_apify_cols:
                        st.info(f"ℹ️ Note: {len(missing_apify_cols)} column(s) were skipped because they don't exist in the dataset")
                    
                    # Force a rerun to update the display
                    st.rerun()
        
        st.markdown("---")
        
        # Create tabs for different views
        tab1, tab2, tab3 = st.tabs(["📋 Collated Data", "📊 Processing Log", "📈 Statistics"])
        
        with tab1:
            st.markdown("#### Collated Dataset")
            st.dataframe(master_df, use_container_width=True, height=400)
            
            # Download button for collated data
            csv_buffer = BytesIO()
            master_df.to_csv(csv_buffer, index=False)
            csv_buffer.seek(0)
            
            st.download_button(
                label="⬇️ Download Collated CSV",
                data=csv_buffer,
                file_name="collated_data.csv",
                mime="text/csv"
            )
        
        with tab2:
            st.markdown("#### Processing Log")
            st.dataframe(log_df, use_container_width=True)
            
            # Download button for log
            log_buffer = BytesIO()
            log_df.to_csv(log_buffer, index=False)
            log_buffer.seek(0)
            
            st.download_button(
                label="⬇️ Download Processing Log",
                data=log_buffer,
                file_name="processing_log.csv",
                mime="text/csv"
            )
        
        with tab3:
            st.markdown("#### Dataset Statistics")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric("Total Rows", f"{len(master_df):,}")
            with col2:
                st.metric("Total Columns", len(master_df.columns))
            with col3:
                st.metric("Missing Values", f"{master_df.isna().sum().sum():,}")
            
            st.markdown("#### Column Information")
            col_info = pd.DataFrame({
                'Column Name': master_df.columns,
                'Data Type': master_df.dtypes.values,
                'Non-Null Count': master_df.count().values,
                'Null Count': master_df.isna().sum().values
            })
            st.dataframe(col_info, use_container_width=True)
            
            st.markdown("#### Rows Added by File")
            if has_excel_types and sheet_mode == 'all_sheets':
                # Group by file and show sheet-level detail
                st.bar_chart(log_df.set_index('Sheet Name')['Rows Added'])
                
                # Show file-level summary
                st.markdown("#### Rows Added by File (Summary)")
                file_summary = log_df.groupby('File Name')['Rows Added'].sum()
                st.bar_chart(file_summary)
            else:
                st.bar_chart(log_df.set_index('File Name')['Rows Added'])
        

    
    else:
        # Show the process button if files haven't been processed yet
        col1, col2, col3 = st.columns([1, 1, 1])
        with col2:
            process_button = st.button("🔄 Process and Collate Files", type="primary", use_container_width=True)
        
        if process_button:
            
            # Initialize tracking variables
            master_df = pd.DataFrame()
            log_data = []
            total_rows_added = 0
            files_processed = 0
            sheets_processed = 0
            
            # Progress bar
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            # Process each file
            for idx, uploaded_file in enumerate(uploaded_files):
                try:
                    file_extension = uploaded_file.name.split('.')[-1].lower()
                    
                    if file_extension == 'csv':
                        status_text.text(f"Processing file {idx + 1} of {len(uploaded_files)}: {uploaded_file.name}")
                        progress_bar.progress((idx) / len(uploaded_files))
                        
                        df_temp = pd.read_csv(uploaded_file)
                        rows_added = len(df_temp)
                        
                        # Merge with master dataframe
                        if master_df.empty:
                            master_df = df_temp.copy()
                        else:
                            master_df = pd.concat([master_df, df_temp], ignore_index=True, sort=False)
                        
                        log_data.append({
                            'File Name': uploaded_file.name,
                            'Sheet Name': 'N/A (CSV)',
                            'Rows Added': rows_added
                        })
                        
                        files_processed += 1
                        total_rows_added += rows_added
                        
                    elif file_extension in ['xlsx', 'xls', 'xlsm', 'xlsb']:
                        # Read Excel file
                        engine = 'openpyxl' if file_extension in ['xlsx', 'xlsm'] else ('pyxlsb' if file_extension == 'xlsb' else 'xlrd')
                        
                        if sheet_mode == 'first_sheet':
                            status_text.text(f"Processing file {idx + 1} of {len(uploaded_files)}: {uploaded_file.name} (First Sheet)")
                            progress_bar.progress((idx) / len(uploaded_files))
                            
                            df_temp = pd.read_excel(uploaded_file, sheet_name=0, engine=engine)
                            rows_added = len(df_temp)
                            
                            # Get sheet name
                            xl_file = pd.ExcelFile(uploaded_file, engine=engine)
                            first_sheet_name = xl_file.sheet_names[0]
                            
                            if master_df.empty:
                                master_df = df_temp.copy()
                            else:
                                master_df = pd.concat([master_df, df_temp], ignore_index=True, sort=False)
                            
                            log_data.append({
                                'File Name': uploaded_file.name,
                                'Sheet Name': first_sheet_name,
                                'Rows Added': rows_added
                            })
                            
                            files_processed += 1
                            sheets_processed += 1
                            total_rows_added += rows_added
                            
                        elif sheet_mode == 'all_sheets':
                            xl_file = pd.ExcelFile(uploaded_file, engine=engine)
                            
                            for sheet_idx, sheet in enumerate(xl_file.sheet_names):
                                status_text.text(f"Processing file {idx + 1} of {len(uploaded_files)}: {uploaded_file.name} - Sheet: {sheet}")
                                progress_bar.progress((idx + (sheet_idx / len(xl_file.sheet_names))) / len(uploaded_files))
                                
                                df_temp = pd.read_excel(uploaded_file, sheet_name=sheet, engine=engine)
                                rows_added = len(df_temp)
                                
                                if master_df.empty:
                                    master_df = df_temp.copy()
                                else:
                                    master_df = pd.concat([master_df, df_temp], ignore_index=True, sort=False)
                                
                                log_data.append({
                                    'File Name': uploaded_file.name,
                                    'Sheet Name': sheet,
                                    'Rows Added': rows_added
                                })
                                
                                sheets_processed += 1
                                total_rows_added += rows_added
                            
                            files_processed += 1
                            
                        elif sheet_mode == 'by_name':
                            status_text.text(f"Processing file {idx + 1} of {len(uploaded_files)}: {uploaded_file.name} - Sheet: {sheet_name}")
                            progress_bar.progress((idx) / len(uploaded_files))
                            
                            try:
                                df_temp = pd.read_excel(uploaded_file, sheet_name=sheet_name, engine=engine)
                                rows_added = len(df_temp)
                                
                                if master_df.empty:
                                    master_df = df_temp.copy()
                                else:
                                    master_df = pd.concat([master_df, df_temp], ignore_index=True, sort=False)
                                
                                log_data.append({
                                    'File Name': uploaded_file.name,
                                    'Sheet Name': sheet_name,
                                    'Rows Added': rows_added
                                })
                                
                                files_processed += 1
                                sheets_processed += 1
                                total_rows_added += rows_added
                                
                            except Exception as sheet_error:
                                st.warning(f"Sheet '{sheet_name}' not found in {uploaded_file.name}")
                                log_data.append({
                                    'File Name': uploaded_file.name,
                                    'Sheet Name': f"'{sheet_name}' (NOT FOUND)",
                                    'Rows Added': 0
                                })
                                continue
                    else:
                        st.warning(f"Unsupported file type for {uploaded_file.name}")
                        continue
                    
                except Exception as e:
                    st.error(f"Error processing {uploaded_file.name}: {str(e)}")
                    log_data.append({
                        'File Name': uploaded_file.name,
                        'Sheet Name': 'ERROR',
                        'Rows Added': 0
                    })
                    continue
            
            # Complete progress
            progress_bar.progress(1.0)
            status_text.text("✅ Processing complete!")
            
            # Create log dataframe
            log_df = pd.DataFrame(log_data)
            
            # Store in session state for column creation feature
            st.session_state['master_df'] = master_df
            st.session_state['log_df'] = log_df
            st.session_state['files_processed'] = files_processed
            st.session_state['sheets_processed'] = sheets_processed
            st.session_state['total_rows_added'] = total_rows_added
            
            # Rerun to show the results
            st.rerun()

else:
    # Feature cards
    st.markdown("### ✨ Key Features")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("""
        <div class='card'>
            <div style='font-size: 2rem; margin-bottom: 0.5rem;'>📁</div>
            <h4>Multiple Formats</h4>
            <p>Support for CSV, XLSX, XLS, XLSM, and XLSB files</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div class='card'>
            <div style='font-size: 2rem; margin-bottom: 0.5rem;'>🔄</div>
            <h4>Smart Merging</h4>
            <p>Dynamic header matching and automatic column alignment</p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown("""
        <div class='card'>
            <div style='font-size: 2rem; margin-bottom: 0.5rem;'>📊</div>
            <h4>Excel Flexibility</h4>
            <p>Process first sheet, all sheets, or specific sheets by name</p>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # How to use section
    st.markdown("### 🚀 How to Get Started")
    
    steps = [
        ("1️⃣", "**Select File Types**", "Choose which file formats you want to upload (CSV, XLSX, etc.)"),
        ("2️⃣", "**Configure Excel Options**", "Decide how to handle Excel sheets (first, all, or by name)"),
        ("3️⃣", "**Upload Files**", "Select multiple files from your computer"),
        ("4️⃣", "**Process & Collate**", "Click the process button to merge all files"),
        ("5️⃣", "**Create New Columns**", "Optionally combine existing columns with custom delimiters"),
        ("6️⃣", "**Generate Apify Time-Stamp**", "Create special timestamp column from EXPERIENCE fields"),
        ("7️⃣", "**Download Results**", "Export your collated data as CSV"),
    ]
    
    for emoji, title, description in steps:
        st.markdown(f"""
        <div style='display: flex; align-items: start; margin: 1rem 0; padding: 1rem; background: #f8fafc; border-radius: 12px; border-left: 4px solid #667eea;'>
            <div style='font-size: 1.5rem; margin-right: 1rem;'>{emoji}</div>
            <div>
                <div style='font-weight: 600; color: #1e293b; margin-bottom: 0.25rem;'>{title}</div>
                <div style='color: #64748b;'>{description}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    # Supported formats
    st.markdown("### 📋 Supported File Types")
    
    formats = [
        ("CSV", "Comma-separated values", "#10b981"),
        ("XLSX", "Excel 2007+ workbook", "#3b82f6"),
        ("XLS", "Excel 97-2003 workbook", "#f59e0b"),
        ("XLSM", "Excel macro-enabled workbook", "#8b5cf6"),
        ("XLSB", "Excel binary workbook", "#ec4899"),
    ]
    
    cols = st.columns(5)
    for idx, (format_name, description, color) in enumerate(formats):
        with cols[idx]:
            st.markdown(f"""
            <div style='text-align: center; padding: 1rem; background: {color}15; border-radius: 12px; border: 2px solid {color}40;'>
                <div style='font-weight: 700; font-size: 1.2rem; color: {color}; margin-bottom: 0.5rem;'>{format_name}</div>
                <div style='font-size: 0.85rem; color: #64748b;'>{description}</div>
            </div>
            """, unsafe_allow_html=True)

# Sidebar with modern styling
with st.sidebar:
    st.markdown("## ℹ️ About")
    
    st.markdown("""
    <div style='background: rgba(255,255,255,0.1); padding: 1rem; border-radius: 12px; margin: 1rem 0;'>
        <p style='color: white; margin: 0;'>
            <strong>Data File Collator</strong> is a professional tool designed to streamline your data merging workflow.
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("### 🎯 Capabilities")
    st.markdown("""
    <div style='color: white;'>
        <span class='badge'>Multi-Format</span>
        <span class='badge'>Smart Merging</span>
        <span class='badge'>Excel Sheets</span>
        <span class='badge'>Column Combo</span>
        <span class='badge'>Apify Stamp</span>
        <span class='badge'>Export Options</span>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    st.markdown("### 💡 Pro Tips")
    st.markdown("""
    <div style='color: white; font-size: 0.9rem;'>
        <ul style='padding-left: 1.5rem;'>
            <li style='margin-bottom: 0.5rem;'>Files with different headers are automatically merged</li>
            <li style='margin-bottom: 0.5rem;'>Missing columns are filled with empty values</li>
            <li style='margin-bottom: 0.5rem;'>Sheet-level tracking for Excel files</li>
            <li style='margin-bottom: 0.5rem;'>Column combination warns about missing sources</li>
            <li style='margin-bottom: 0.5rem;'>Apify Time-Stamp uses newline delimiters</li>
        </ul>
    </div>
    """, unsafe_allow_html=True)
    
    if uploaded_files:
        st.markdown("---")
        st.markdown("### 📂 Uploaded Files")
        for file in uploaded_files:
            file_ext = file.name.split('.')[-1].upper()
            st.markdown(f"""
            <div style='background: rgba(255,255,255,0.1); padding: 0.5rem 1rem; border-radius: 8px; margin: 0.5rem 0; color: white;'>
                <strong>{file_ext}</strong> • {file.name}
            </div>
            """, unsafe_allow_html=True)
    
    st.markdown("---")
    st.markdown("""
    <div style='text-align: center; color: rgba(255,255,255,0.7); font-size: 0.85rem; margin-top: 2rem;'>
        Made by Felix using Streamlit
    </div>
    """, unsafe_allow_html=True)