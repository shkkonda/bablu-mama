import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import xlsxwriter
from datetime import datetime
import re


def fix_truncated_dates(date_str):
    """Fix truncated dates by adding current year if needed"""
    if pd.isna(date_str) or date_str == '':
        return None

    # Convert to string if not already
    date_str = str(date_str).strip()

    # Pattern to match truncated dates like "07.03.202", "15.12.202", etc.
    truncated_pattern = r'^(\d{1,2})\.(\d{1,2})\.(\d{3})$'
    match = re.match(truncated_pattern, date_str)

    if match:
        day, month, year_partial = match.groups()
        # Add current year prefix (202 -> 2025)
        current_year = datetime.now().year
        year_str = str(current_year)
        # Take the first 3 digits of current year and add the partial year digit
        if len(year_partial) == 3:
            full_year = year_str[:3] + year_partial[-1]  # e.g., "202" + "2" = "2022", but we want 2025
            # Actually, let's just use current year for truncated dates
            full_year = str(current_year)
        else:
            full_year = str(current_year)

        fixed_date = f"{day}.{month}.{full_year}"
        return fixed_date

    # Handle other common truncated patterns
    patterns_to_fix = [
        (r'^(\d{1,2})/(\d{1,2})/(\d{2,3})$', r'\1/\2/2025'),  # MM/DD/YY or MM/DD/YYY
        (r'^(\d{1,2})-(\d{1,2})-(\d{2,3})$', r'\1-\2-2025'),  # MM-DD-YY or MM-DD-YYY
        (r'^(\d{4})-(\d{1,2})-(\d{1,2})$', date_str),  # Already good format
    ]

    for pattern, replacement in patterns_to_fix:
        if re.match(pattern, date_str):
            if replacement == date_str:
                return date_str
            return re.sub(pattern, replacement, date_str)

    return date_str


def safe_date_conversion(date_series):
    """Safely convert dates with truncation repair"""
    fixed_dates = []

    for date_val in date_series:
        try:
            # First try to fix truncated dates
            fixed_date = fix_truncated_dates(date_val)
            if fixed_date is None:
                fixed_dates.append(None)
                continue

            # Try to parse the fixed date
            parsed_date = pd.to_datetime(fixed_date, errors='coerce')
            if pd.isna(parsed_date):
                # If still can't parse, try other common formats
                for fmt in ['%d.%m.%Y', '%d/%m/%Y', '%d-%m-%Y', '%Y-%m-%d', '%m/%d/%Y']:
                    try:
                        parsed_date = pd.to_datetime(fixed_date, format=fmt)
                        break
                    except:
                        continue

            fixed_dates.append(parsed_date)

        except Exception as e:
            st.warning(f"Could not parse date '{date_val}': {e}")
            fixed_dates.append(None)

    return pd.Series(fixed_dates)


def create_crosstab_excel(df):
    """Create cross-tabulation table and return Excel file as bytes"""

    df['INV DATE'] = pd.to_datetime(df['INV DATE'])

    # Create pivot table
    pivot_table = df.pivot_table(
        index='CONSIGNEE NAME',
        columns='INV DATE',
        values='Qty',
        aggfunc='sum',
        fill_value=0
    )

    pivot_table.columns = pd.to_datetime(pivot_table.columns)
    pivot_table = pivot_table.sort_index(axis=1)

    # Create Excel file in memory
    output = BytesIO()

    with pd.ExcelWriter(output, engine='xlsxwriter', datetime_format='yyyy-mm-dd') as writer:
        # Write the pivot table to Excel
        pivot_table.to_excel(writer, sheet_name='Consignee Summary', index=True)

        # Get the workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['Consignee Summary']

        # Add formatting
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1,
            'align': 'center'
        })

        date_header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#D7E4BC',
            'border': 1,
            'num_format': 'yyyy-mm-dd',
            'align': 'center'
        })

        number_format = workbook.add_format({
            'num_format': '#,##0',
            'border': 1
        })

        # Write date headers
        for col_idx, date_col in enumerate(pivot_table.columns, start=1):
            try:
                if isinstance(date_col, (pd.Timestamp, datetime)):
                    worksheet.write_datetime(0, col_idx, date_col, date_header_format)
                else:
                    # Convert date object to datetime for Excel
                    excel_date = pd.to_datetime(str(date_col))
                    worksheet.write_datetime(0, col_idx, excel_date, date_header_format)
            except:
                # Fallback to string if datetime conversion fails
                worksheet.write(0, col_idx, str(date_col), header_format)

        # Format consignee names column
        for row_num, value in enumerate(pivot_table.index):
            worksheet.write(row_num + 1, 0, value, header_format)

        # Format data cells
        for row in range(1, len(pivot_table.index) + 1):
            for col in range(1, len(pivot_table.columns) + 1):
                worksheet.write(row, col, pivot_table.iloc[row - 1, col - 1], number_format)

        # Auto-adjust column widths
        worksheet.set_column('A:A', 25)  # Consignee names column
        for i in range(len(pivot_table.columns)):
            worksheet.set_column(i + 1, i + 1, 12)  # Date columns

    output.seek(0)
    return output.getvalue()


def main():
    st.set_page_config(
        page_title="Excel Cross-Tab Generator",
        page_icon="📊",
        layout="wide"
    )

    st.title("📊 Excel Cross-Tab Generator")
    st.markdown(
        "Upload your Excel file to create a cross-tabulation summary with consignees as rows and dates as columns.")

    # File upload section
    st.header("📁 Upload Your Excel File")
    uploaded_file = st.file_uploader(
        "Choose an Excel file",
        type=['xlsx', 'xls'],
        help="Upload an Excel file containing CONSIGNEE NAME, INV DATE, and Qty columns"
    )

    if uploaded_file is not None:
        try:
            # Read the uploaded file
            df = pd.read_excel(uploaded_file)

            st.success(f"✅ File uploaded successfully! Found {len(df)} rows of data.")

            # Display original data preview
            st.header("📋 Original Data Preview")
            st.dataframe(df.head(10), use_container_width=True)

            # Auto-detect required columns
            def find_column(df, target_names):
                """Find column by matching against list of possible names"""
                df_cols_lower = [col.upper().strip() for col in df.columns]
                for target in target_names:
                    if target.upper() in df_cols_lower:
                        return df.columns[df_cols_lower.index(target.upper())]
                return None

            # Try to auto-detect columns
            consignee_col = find_column(df, ['CONSIGNEE NAME', 'CONSIGNEE', 'CUSTOMER NAME', 'CUSTOMER'])
            date_col = find_column(df, ['INV DATE', 'INVOICE DATE', 'DATE', 'INV_DATE'])
            qty_col = find_column(df, ['QTY', 'QUANTITY', 'AMOUNT', 'VALUE'])

            # Show column mapping section only if columns not found
            missing_cols = []
            if not consignee_col:
                missing_cols.append("CONSIGNEE NAME")
            if not date_col:
                missing_cols.append("INV DATE")
            if not qty_col:
                missing_cols.append("Qty")

            if missing_cols:
                st.header("🔗 Column Mapping Required")
                st.warning(f"Could not auto-detect columns: {', '.join(missing_cols)}. Please select them manually.")

                col1, col2, col3 = st.columns(3)

                with col1:
                    if not consignee_col:
                        consignee_col = st.selectbox(
                            "Select CONSIGNEE NAME column:",
                            options=df.columns.tolist(),
                            help="Column containing customer/consignee names"
                        )
                    else:
                        st.success(f"✅ CONSIGNEE NAME: {consignee_col}")

                with col2:
                    if not date_col:
                        date_col = st.selectbox(
                            "Select INV DATE column:",
                            options=df.columns.tolist(),
                            help="Column containing invoice dates"
                        )
                    else:
                        st.success(f"✅ INV DATE: {date_col}")

                with col3:
                    if not qty_col:
                        qty_col = st.selectbox(
                            "Select Qty column:",
                            options=df.columns.tolist(),
                            help="Column containing quantities to sum"
                        )
                    else:
                        st.success(f"✅ Qty: {qty_col}")
            else:
                st.header("✅ Columns Auto-Detected")
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.success(f"CONSIGNEE NAME: **{consignee_col}**")
                with col2:
                    st.success(f"INV DATE: **{date_col}**")
                with col3:
                    st.success(f"Qty: **{qty_col}**")

            if st.button("🔄 Generate Cross-Tab Summary", type="primary"):
                try:
                    # Create working dataframe with selected columns
                    work_df = df[[consignee_col, date_col, qty_col]].copy()
                    work_df.columns = ['CONSIGNEE NAME', 'INV DATE', 'Qty']

                    # Clean and prepare data
                    work_df = work_df.dropna(subset=['CONSIGNEE NAME'])

                    # Show info about date fixing
                    st.info("🔧 Checking and fixing truncated dates...")

                    # Apply safe date conversion with truncation repair
                    work_df['INV DATE'] = safe_date_conversion(work_df['INV DATE'])

                    # Remove rows where date conversion failed
                    original_count = len(work_df)
                    work_df = work_df.dropna(subset=['INV DATE'])
                    failed_dates = original_count - len(work_df)

                    if failed_dates > 0:
                        st.warning(f"⚠️ {failed_dates} rows had invalid dates and were removed.")

                    # Convert dates to date format for display
                    work_df['INV DATE'] = work_df['INV DATE'].dt.date

                    # Convert Qty to numeric, replacing non-numeric with 0
                    work_df['Qty'] = pd.to_numeric(work_df['Qty'], errors='coerce').fillna(0)

                    # Remove rows where consignee name is empty or NaN
                    work_df = work_df[work_df['CONSIGNEE NAME'].notna()]
                    work_df = work_df[work_df['CONSIGNEE NAME'] != '']

                    st.success(f"✅ Data processed successfully! Working with {len(work_df)} valid records.")

                    # Show sample of fixed dates
                    if len(work_df) > 0:
                        unique_dates = sorted(work_df['INV DATE'].unique())
                        st.info(f"📅 Found dates ranging from {unique_dates[0]} to {unique_dates[-1]}")

                    # Create pivot table for preview
                    pivot_preview = work_df.pivot_table(
                        index='CONSIGNEE NAME',
                        columns='INV DATE',
                        values='Qty',
                        aggfunc='sum',
                        fill_value=0
                    )

                    # Display summary statistics
                    st.header("📊 Summary Statistics")
                    col1, col2, col3 = st.columns(3)

                    with col1:
                        st.metric("Unique Consignees", len(pivot_preview.index))

                    with col2:
                        st.metric("Unique Dates", len(pivot_preview.columns))

                    with col3:
                        st.metric("Total Quantity", f"{work_df['Qty'].sum():,.0f}")

                    # Display cross-tab preview
                    st.header("📋 Cross-Tab Preview")
                    st.dataframe(pivot_preview, use_container_width=True)

                    # For Excel generation, we need to preserve datetime objects
                    # Create a copy with datetime objects for Excel
                    excel_df = work_df.copy()
                    excel_df['INV DATE'] = safe_date_conversion(df[date_col])
                    excel_df = excel_df.dropna(subset=['INV DATE'])
                    excel_df['Qty'] = pd.to_numeric(excel_df['Qty'], errors='coerce').fillna(0)
                    excel_df = excel_df[excel_df['CONSIGNEE NAME'].notna()]
                    excel_df = excel_df[excel_df['CONSIGNEE NAME'] != '']

                    # Generate Excel file with datetime objects
                    excel_data = create_crosstab_excel(excel_df)

                    # Download button
                    st.header("💾 Download Results")
                    st.download_button(
                        label="📥 Download Excel Summary",
                        data=excel_data,
                        file_name="consignee_summary.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )

                    st.success(
                        "🎉 Cross-tab summary generated successfully! Click the download button to save your file.")

                except Exception as e:
                    st.error(f"❌ Error processing data: {str(e)}")
                    st.error("Please check your column selections and data format.")
                    # Show more detailed error info
                    st.error(f"Detailed error: {type(e).__name__}")

        except Exception as e:
            st.error(f"❌ Error reading file: {str(e)}")
            st.error("Please make sure you uploaded a valid Excel file.")

    else:
        # Instructions when no file is uploaded
        st.info("👆 Please upload an Excel file to get started.")

        st.header("📋 Instructions")
        with st.expander("📖 How to Use This App", expanded=False):
            st.markdown("""
            **How to use this app:**
    
            1. **Upload** your Excel file using the file uploader above
            2. **Map** your columns to the required fields:
               - **CONSIGNEE NAME**: Column containing customer/consignee names
               - **INV DATE**: Column containing invoice dates
               - **Qty**: Column containing quantities to sum
            3. **Generate** the cross-tab summary
            4. **Download** the formatted Excel file
    
            **Output Format:**
            - Rows: Distinct consignee names
            - Columns: Distinct invoice dates
            - Values: Sum of quantities for each consignee-date combination
    
            **Requirements:**
            - Excel file (.xlsx or .xls)
            - Columns for consignee names, dates, and quantities
            - Data should be in tabular format with headers
    
            **🔧 Date Fixing Feature:**
            - Automatically fixes truncated dates (e.g., "07.03.202" → "07.03.2025")
            - Handles various date formats
            - Uses current year (2025) for incomplete dates
            """)


if __name__ == "__main__":
    main()
