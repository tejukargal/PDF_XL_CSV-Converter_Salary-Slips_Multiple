import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
import os
from datetime import datetime

def extract_salary_details(pdf_path):
    """
    Extract salary details from PDF and return as a list of dictionaries
    """
    salary_records = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                
                # Split text into salary slip sections
                salary_slips = re.split(r'SNO:\s+\d+', text)
                
                for slip in salary_slips[1:]:  # Skip first empty split
                    record = {}
                    
                    # Basic Details
                    emp_no_match = re.search(r'EMP No\s+(\d+)', slip)
                    name_match = re.search(r'Sri / Smt:\s+([A-Z\s]+)', slip)
                    designation_match = re.search(r'Designation:\s+([A-Z\s]+)', slip)
                    pay_scale_match = re.search(r'Pay Scale\s*:\s*(\d+)-(\d+)', slip)
                    ddo_code_match = re.search(r'DDO Code\s*:\s*(\w+)', slip)
                    days_worked_match = re.search(r'Days Worked:\s*(\d+)', slip)
                    next_increment_match = re.search(r'Next Increment Date:\s*([A-Za-z]+\s+\d{4})', slip)
                    group_match = re.search(r'Group\s*:\s*([A-Z])', slip)
                    
                    # Extract Pay Details
                    basic_match = re.search(r'Basic\s*:\s*(\d+)', slip)
                    
                    # Allowances
                    da_match = re.search(r'DA\s+(\d+)', slip)
                    hra_match = re.search(r'HRA\s+(\d+)', slip)
                    ir_match = re.search(r'IR\s+(\d+)', slip)
                    sfn_match = re.search(r'SFN\s+(\d+)', slip)
                    p_match = re.search(r'P\s+(\d+)', slip)
                    spaytypist_match = re.search(r'SPAY-TYPIST\s+(\d+)', slip)
                    
                    # Deductions
                    it_match = re.search(r'IT\s+(\d+)', slip)
                    pt_match = re.search(r'PT\s+(\d+)', slip)
                    gslic_match = re.search(r'GSLIC\s+(\d+)', slip)
                    lic_match = re.search(r'(?<!GS)LIC\s+(\d+)', slip)
                    fbf_match = re.search(r'FBF\s+(\d+)', slip)
                    
                    # Summary Details
                    gross_salary_match = re.search(r'Gross Salary:\s*Rs\.\s*(\d+)', slip)
                    net_salary_match = re.search(r'Net Salary\s*:\s*Rs\.\s*(\d+)', slip)
                    deductions_match = re.search(r'sum of deductions &Recoveries\s*:\s*Rs\.\s*(\d+)', slip)
                    
                    # Bank Details
                    account_match = re.search(r'Bank A/C Number:\s*(\d+)', slip)
                    
                    # Extract month and year
                    date_match = re.search(r'Month Of\s+([A-Za-z]+)\s+(\d{4})', slip)
                    month = date_match.group(1) if date_match else ''
                    year = date_match.group(2) if date_match else ''
                    
                    # Populate record dictionary
                    record.update({
                        'Month': month,
                        'Year': year,
                        'Employee_ID': emp_no_match.group(1) if emp_no_match else '',
                        'Name': name_match.group(1).strip() if name_match else '',
                        'Designation': designation_match.group(1).strip() if designation_match else '',
                        'Pay_Scale': f"{pay_scale_match.group(1)}-{pay_scale_match.group(2)}" if pay_scale_match else '',
                        'DDO_Code': ddo_code_match.group(1) if ddo_code_match else '',
                        'Days_Worked': days_worked_match.group(1) if days_worked_match else '',
                        'Next_Increment_Date': next_increment_match.group(1) if next_increment_match else '',
                        'Group': group_match.group(1) if group_match else '',
                        'Basic_Pay': basic_match.group(1) if basic_match else '0',
                        
                        # Allowances
                        'DA': da_match.group(1) if da_match else '0',
                        'HRA': hra_match.group(1) if hra_match else '0',
                        'IR': ir_match.group(1) if ir_match else '0',
                        'SFN': sfn_match.group(1) if sfn_match else '0',
                        'P': p_match.group(1) if p_match else '0',
                        'SPAY_TYPIST': spaytypist_match.group(1) if spaytypist_match else '0',
                        
                        # Deductions
                        'IT_Deduction': it_match.group(1) if it_match else '0',
                        'PT_Deduction': pt_match.group(1) if pt_match else '0',
                        'GSLIC_Deduction': gslic_match.group(1) if gslic_match else '0',
                        'LIC_Deduction': lic_match.group(1) if lic_match else '0',
                        'FBF_Deduction': fbf_match.group(1) if fbf_match else '0',
                        
                        # Summary
                        'Gross_Salary': gross_salary_match.group(1) if gross_salary_match else '0',
                        'Net_Salary': net_salary_match.group(1) if net_salary_match else '0',
                        'Total_Deductions': deductions_match.group(1) if deductions_match else '0',
                        
                        # Bank Details
                        'Account_Number': account_match.group(1) if account_match else ''
                    })
                    
                    if record['Employee_ID']:  # Only add if we found an employee ID
                        salary_records.append(record)
    
    except Exception as e:
        st.error(f"Error processing PDF: {str(e)}")
        return None
        
    return salary_records

def create_excel_file(df):
    """
    Create Excel file with proper formatting
    """
    buffer = io.BytesIO()
    
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Salary Data', index=False)
        
        # Get workbook and worksheet objects
        workbook = writer.book
        worksheet = writer.sheets['Salary Data']
        
        # Define number format
        number_format = workbook.add_format({'num_format': '0'})
        
        # Apply number format to specific columns
        numeric_columns = {
            'Employee_ID': 'A',
            'Year': 'B',
            'Account_Number': 'C',
            # Allowances
            'DA': 'D',
            'HRA': 'E',
            'IR': 'F',
            'SFN': 'G',
            'P': 'H',
            'SPAY_TYPIST': 'I',
            # Deductions
            'IT_Deduction': 'J',
            'PT_Deduction': 'K',
            'GSLIC_Deduction': 'L',
            'LIC_Deduction': 'M',
            'FBF_Deduction': 'N'
        }
        
        # Get column indices
        for col_name, col_letter in numeric_columns.items():
            if col_name in df.columns:
                col_idx = df.columns.get_loc(col_name)
                worksheet.set_column(col_idx, col_idx, None, number_format)
    
    return buffer

def create_monthly_summary(df):
    """
    Create a detailed monthly summary of salary data with improved date handling
    """
    # Create a month number mapping
    month_map = {
        'January': 1, 'Jan': 1, 'February': 2, 'Feb': 2, 'March': 3, 'Mar': 3,
        'April': 4, 'Apr': 4, 'May': 5, 'June': 6, 'Jun': 6,
        'July': 7, 'Jul': 7, 'August': 8, 'Aug': 8, 'September': 9, 'Sep': 9,
        'October': 10, 'Oct': 10, 'November': 11, 'Nov': 11, 'December': 12, 'Dec': 12,
        'Feburary': 2  # Common misspelling
    }
    
    # Convert month names to numbers and create date field
    df['Month_Num'] = df['Month'].str.title().map(month_map)
    df['Date'] = pd.to_datetime(df['Year'].astype(str) + '-' + df['Month_Num'].astype(str) + '-01')
    
    # Sort by date
    df = df.sort_values('Date')
    
    # Group by Month and Year
    monthly_groups = df.groupby(['Year', 'Month'])
    
    summary_data = []
    
    for (year, month), group in monthly_groups:
        summary = {
            'Year': year,
            'Month': month,
            'Employee_Count': len(group),
            
            # Allowances Summary
            'Total_Basic': group['Basic_Pay'].astype(float).sum(),
            'Total_DA': group['DA'].astype(float).sum(),
            'Total_HRA': group['HRA'].astype(float).sum(),
            'Total_IR': group['IR'].astype(float).sum(),
            'Total_SFN': group['SFN'].astype(float).sum(),
            'Total_P': group['P'].astype(float).sum(),
            'Total_SPAY_TYPIST': group['SPAY_TYPIST'].astype(float).sum(),
            
            # Deductions Summary
            'Total_IT': group['IT_Deduction'].astype(float).sum(),
            'Total_PT': group['PT_Deduction'].astype(float).sum(),
            'Total_GSLIC': group['GSLIC_Deduction'].astype(float).sum(),
            'Total_LIC': group['LIC_Deduction'].astype(float).sum(),
            'Total_FBF': group['FBF_Deduction'].astype(float).sum(),
            
            # Overall Summary
            'Total_Gross': group['Gross_Salary'].astype(float).sum(),
            'Total_Deductions': group['Total_Deductions'].astype(float).sum(),
            'Total_Net': group['Net_Salary'].astype(float).sum(),
            'Avg_Gross_Salary': group['Gross_Salary'].astype(float).mean(),
            'Avg_Net_Salary': group['Net_Salary'].astype(float).mean()
        }
        summary_data.append(summary)
    
    # Create DataFrame and sort by date
    summary_df = pd.DataFrame(summary_data)
    
    # Add Month_Num for sorting
    summary_df['Month_Num'] = summary_df['Month'].str.title().map(month_map)
    summary_df = summary_df.sort_values(['Year', 'Month_Num'])
    
    # Drop the Month_Num column as it's no longer needed
    summary_df = summary_df.drop('Month_Num', axis=1)
    
    return summary_df

def format_currency(value):
    """
    Format numbers as currency
    """
    return f"₹{value:,.2f}"

def main():
    st.set_page_config(page_title="Salary Slip PDF to CSV Converter", layout="wide")
    
    # Custom CSS for watermark
    st.markdown("""
        <style>
        [data-testid="stSidebar"] .watermark {
            position: relative;
            text-align: left;
            padding: 10px 0;
            font-size: 18px;
            font-weight: bold;
            font-style: italic;
            opacity: 0.9;
            background: linear-gradient(45deg, #1e88e5, #e91e63);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            animation: fadeInOut 2s infinite alternate;
        }
        
        @keyframes fadeInOut {
            from { opacity: 0.5; }
            to { opacity: 1; }
        }
        
        @media (prefers-color-scheme: dark) {
            [data-testid="stSidebar"] .watermark {
                background: linear-gradient(45deg, #90caf9, #f48fb1);
                -webkit-background-clip: text;
                -webkit-text-fill-color: transparent;
            }
        }
        </style>
        """, unsafe_allow_html=True)
    
    # Sidebar
    with st.sidebar:
        st.markdown('<div class="watermark">by Teju SMP</div>', unsafe_allow_html=True)
        st.title("File Management")
        st.subheader("Upload Files")
        uploaded_files = st.file_uploader(
            "Choose PDF files",
            type=['pdf'],
            accept_multiple_files=True,
            help="Select one or more PDF files to process"
        )
        
        if uploaded_files:
            st.subheader("Uploaded Files")
            for file in uploaded_files:
                st.text(file.name)
    
    # Main content
    st.title("Salary Slip PDF to CSV Converter")
    
    if uploaded_files:
        all_records = []
        processed_files = []
        failed_files = []
        
        for uploaded_file in uploaded_files:
            temp_path = f"temp_{uploaded_file.name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
            
            try:
                with open(temp_path, "wb") as f:
                    f.write(uploaded_file.getvalue())
                
                with st.spinner(f"Processing {uploaded_file.name}..."):
                    salary_records = extract_salary_details(temp_path)
                    if salary_records:
                        all_records.extend(salary_records)
                        processed_files.append(uploaded_file.name)
                    else:
                        failed_files.append(uploaded_file.name)
            
            except Exception as e:
                st.error(f"Error processing {uploaded_file.name}: {str(e)}")
                failed_files.append(uploaded_file.name)
            
            finally:
                if os.path.exists(temp_path):
                    os.remove(temp_path)
        
        if all_records:
            df = pd.DataFrame(all_records)
            
            # Main data preview
            st.subheader("Preview of Extracted Data")
            st.dataframe(df)
            
            # Download buttons in sidebar
            with st.sidebar:
                st.subheader("Download Options")
                
                # CSV download
                csv_data = df.to_csv(index=False)
                st.download_button(
                    label="Download CSV",
                    data=csv_data,
                    file_name=f"salary_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv"
                )
                
                # Excel download with formatting
                excel_buffer = create_excel_file(df)
                st.download_button(
                    label="Download Excel",
                    data=excel_buffer.getvalue(),
                    file_name=f"salary_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.ms-excel"
                )
            
            # Create monthly summary
            summary_df = create_monthly_summary(df)
            
            # Display summary tables
            st.subheader("Monthly Summary")
            
            # Format currency columns
            currency_columns = [col for col in summary_df.columns if col.startswith('Total_') or col.startswith('Avg_')]
            for col in currency_columns:
                summary_df[col] = summary_df[col].apply(format_currency)
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("Allowances Summary")
                allowances_cols = ['Year', 'Month', 'Employee_Count', 'Total_Basic', 
                                 'Total_DA', 'Total_HRA', 'Total_IR', 'Total_SFN', 
                                 'Total_P', 'Total_SPAY_TYPIST']
                st.dataframe(summary_df[allowances_cols])
            
            with col2:
                st.subheader("Deductions Summary")
                deductions_cols = ['Year', 'Month', 'Employee_Count', 'Total_IT', 
                                 'Total_PT', 'Total_GSLIC', 'Total_LIC', 'Total_FBF']
                st.dataframe(summary_df[deductions_cols])
            
            # Overall summary
            st.subheader("Overall Summary")
            overall_cols = ['Year', 'Month', 'Employee_Count', 'Total_Gross', 
                          'Total_Deductions', 'Total_Net', 'Avg_Gross_Salary', 
                          'Avg_Net_Salary']
            st.dataframe(summary_df[overall_cols])
            
            # Processing summary
            with st.sidebar:
                st.subheader("Processing Summary")
                st.write(f"Successfully processed: {len(processed_files)} files")
                if failed_files:
                    st.error(f"Failed to process: {len(failed_files)} files")
                    st.write("Failed files:")
                    for file in failed_files:
                        st.write(f"- {file}")
                
                # Add total records processed
                st.write(f"Total records processed: {len(df)}")
                st.write(f"Total months covered: {len(summary_df)}")
        else:
            st.error("No salary records found in the uploaded PDFs.")
    else:
        st.info("Please upload PDF files using the sidebar to begin processing.")

if __name__ == "__main__":
    main()
