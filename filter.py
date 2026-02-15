import streamlit as st
import pandas as pd
from io import BytesIO

# Page configuration
st.set_page_config(page_title="Excel Data Viewer", layout="wide")

# Title
st.title("📊 Excel Multi-Sheet Data Viewer")

# File uploader
uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        # Determine the engine based on file extension
        file_extension = uploaded_file.name.split('.')[-1].lower()
        
        if file_extension == 'xlsx':
            engine = 'openpyxl'
        elif file_extension == 'xls':
            engine = 'xlrd'
        else:
            st.error(f"Unsupported file format: .{file_extension}")
            st.stop()
        
        # Reset the file pointer to the beginning
        uploaded_file.seek(0)
        
        # Read all sheets from the Excel file
        excel_file = pd.ExcelFile(uploaded_file, engine=engine)
        
        # Expected sheet names
        expected_sheets = [
            "Region",
            "Issuer",
            "Merchant",
            "Acquirer",
            "Issuer-Merchant",
            "Issuer-Acquirer",
            "Acquirer-Merchant",
            "Issuer-Acquirer-Merchant"
        ]
        
        # Store dataframes in session state if not already there
        if 'dataframes' not in st.session_state:
            st.session_state.dataframes = {}
            uploaded_file.seek(0)  # Reset again before reading sheets
            for sheet in expected_sheets:
                if sheet in excel_file.sheet_names:
                    st.session_state.dataframes[sheet] = pd.read_excel(uploaded_file, sheet_name=sheet, engine=engine)
                else:
                    st.warning(f"Sheet '{sheet}' not found in the uploaded file.")
        
        if st.session_state.dataframes:
            st.success(f"✅ Loaded {len(st.session_state.dataframes)} sheets successfully!")
            
            # Sheet selection with drag-and-drop style using multiselect
            st.markdown("---")
            st.subheader("📋 Select Sheets to Display")
            
            available_sheets = list(st.session_state.dataframes.keys())
            
            # Multiselect for choosing which sheets to display
            selected_sheets = st.multiselect(
                "Choose sheets to display (order matters - drag to reorder in the selection box)",
                options=available_sheets,
                default=available_sheets[:1] if available_sheets else [],
                help="Select one or more sheets. They will be displayed in the order you select them."
            )
            
            if selected_sheets:
                st.markdown("---")
                
                # Show available columns for debugging
                with st.expander("📋 View Available Columns"):
                    for sheet in selected_sheets:
                        st.write(f"**{sheet}:**")
                        st.write(list(st.session_state.dataframes[sheet].columns))
                
                # Filter columns that we're looking for
                filter_columns = [
                    "Issuer region name",
                    "Issuer country",
                    "Product Category",
                    "Issuer",
                    "Merchant",
                    "Acquirer"
                ]
                
                # Create a global filter section
                st.subheader("🔍 Global Filters")
                
                # Add filter mode selection
                filter_mode = st.radio(
                    "Filter Mode:",
                    options=["Include (show only selected)", "Exclude (hide selected)"],
                    horizontal=True,
                    help="Include: Show only the selected values. Exclude: Hide the selected values and show everything else."
                )
                is_exclude_mode = filter_mode.startswith("Exclude")
                
                st.markdown("*Filters will be applied to all displayed sheets (if the column exists)*")
                
                # Collect all available columns from selected sheets
                all_columns = set()
                for sheet in selected_sheets:
                    all_columns.update(st.session_state.dataframes[sheet].columns)
                
                # Find which filter columns actually exist
                available_filter_columns = [col for col in filter_columns if col in all_columns]
                
                if not available_filter_columns:
                    st.warning("⚠️ None of the expected filter columns were found in your data. Please check the column names in the expander above.")
                    st.info("Expected columns: " + ", ".join(filter_columns))
                
                # Collect all unique values across all selected sheets for each filter column
                filters = {}
                
                col1, col2, col3 = st.columns(3)
                columns_list = [col1, col2, col3, col1, col2, col3]  # Reuse columns for 6 filters
                
                for idx, filter_col in enumerate(filter_columns):
                    # Collect unique values from all selected sheets
                    unique_values = set()
                    for sheet in selected_sheets:
                        df = st.session_state.dataframes[sheet]
                        if filter_col in df.columns:
                            unique_values.update(df[filter_col].dropna().unique())
                    
                    if unique_values:
                        unique_values = sorted(list(unique_values))
                        with columns_list[idx]:
                            label = f"{'🚫 Exclude' if is_exclude_mode else '✓ Include'} {filter_col}"
                            filters[filter_col] = st.multiselect(
                                label,
                                options=unique_values,
                                default=[],
                                key=f"filter_{filter_col}"
                            )
                
                st.markdown("---")
                
                # Display each selected sheet
                for sheet_name in selected_sheets:
                    df = st.session_state.dataframes[sheet_name].copy()
                    
                    # Apply filters
                    filtered_df = df.copy()
                    for filter_col, selected_values in filters.items():
                        if selected_values and filter_col in filtered_df.columns:
                            if is_exclude_mode:
                                # Exclude mode: show everything EXCEPT selected values
                                filtered_df = filtered_df[~filtered_df[filter_col].isin(selected_values)]
                            else:
                                # Include mode: show ONLY selected values
                                filtered_df = filtered_df[filtered_df[filter_col].isin(selected_values)]
                    
                    # Display sheet
                    st.subheader(f"📄 {sheet_name}")
                    
                    # Show record count
                    col_a, col_b = st.columns([3, 1])
                    with col_a:
                        st.caption(f"Showing {len(filtered_df):,} of {len(df):,} records")
                    with col_b:
                        # Download button for filtered data
                        csv = filtered_df.to_csv(index=False).encode('utf-8')
                        st.download_button(
                            label=f"📥 Download {sheet_name}",
                            data=csv,
                            file_name=f"{sheet_name}_filtered.csv",
                            mime="text/csv",
                            key=f"download_{sheet_name}"
                        )
                    
                    # Display dataframe
                    st.dataframe(
                        filtered_df,
                        use_container_width=True,
                        height=400
                    )
                    
                    st.markdown("---")
            else:
                st.info("👆 Please select at least one sheet to display")
        else:
            st.error("No valid sheets found in the uploaded file.")
            
    except Exception as e:
        st.error(f"Error reading Excel file: {str(e)}")
        
        # Provide more specific guidance based on the error
        error_str = str(e).lower()
        
        if "not a zip file" in error_str or "badzipfile" in error_str:
            st.warning("""
            **Possible solutions:**
            1. The file might be in an older Excel format (.xls). Try saving it as .xlsx in Excel:
               - Open the file in Excel
               - Go to File → Save As
               - Choose "Excel Workbook (.xlsx)" format
               - Save and try uploading again
            
            2. The file might be corrupted. Try:
               - Opening the file in Excel and saving it again
               - Creating a new file and copying the data
            
            3. The file might not be a real Excel file despite having the extension
            """)
        elif "xlrd" in error_str or "engine" in error_str:
            st.warning("""
            **Excel format issue:**
            - For .xls files (Excel 97-2003), install xlrd: `pip install xlrd`
            - For .xlsx files, openpyxl should be installed: `pip install openpyxl`
            """)
        
        st.info(f"**File details:** Name: {uploaded_file.name}, Size: {uploaded_file.size} bytes")
        
        with st.expander("Show full error details"):
            st.exception(e)
else:
    st.info("👆 Please upload an Excel file to get started")
    
    # Display instructions
    st.markdown("""
    ### Instructions:
    1. **Upload** your Excel file using the file uploader above
    2. **Select** which sheets you want to display (in the order you want)
    3. **Filter** the data using the global filters
    4. **Download** filtered data for each sheet if needed
    
    ### Expected Sheet Names:
    - Region
    - Issuer
    - Merchant
    - Acquirer
    - Issuer-Merchant
    - Issuer-Acquirer
    - Acquirer-Merchant
    - Issuer-Acquirer-Merchant
    
    ### Filter Columns:
    The app will look for these columns in your data:
    - Issuer region name
    - Issuer country
    - Product Category
    - Issuer
    - Merchant
    - Acquirer
    """)

# Add a reset button in the sidebar
with st.sidebar:
    st.markdown("### 🔧 Controls")
    if st.button("🔄 Reset All", use_container_width=True):
        if 'dataframes' in st.session_state:
            del st.session_state.dataframes
        st.rerun()
    
    st.markdown("---")
    st.markdown("### 📖 About")
    st.markdown("""
    This app allows you to:
    - Load multi-sheet Excel files
    - Display multiple sheets simultaneously
    - Apply global filters across all sheets
    - Download filtered data
    """)