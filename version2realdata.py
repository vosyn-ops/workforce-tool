import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO

# -------------------- Helper Functions --------------------

def filter_valid_employees(df):
    """Filter employees who are not currently assigned to any project."""
    today = pd.to_datetime(datetime.today().date())
    threshold_date = today + timedelta(days=15)
    df['Please Enter Your End Date'] = pd.to_datetime(df['Please Enter Your End Date'], errors='coerce')
    
    valid_df = df[
        (df['Please Enter Your End Date'] > threshold_date) & 
        (df['Are you currently assigned to a project?'].str.strip().str.lower() == 'no')
    ].copy()
    
    valid_df['Available Hours'] = valid_df['Availability'].apply(
        lambda x: 40 if str(x).strip().lower() == 'full-time' else 20
    )
    return valid_df

def calculate_match_percentage(employee, project):
    """Calculate match percentage based on Languages, Experience, Tools, and Certifications."""
    match_count = 0
    total_count = 0
    match_fields = [
        ('Languages', 'Langauages proficiency required (e.g. Python, Java)'),
        ('Experience', 'Skills Required (e.g. Risk Management, Data Analysis, Data Visualization )'),
        ('Tools', 'Tools (e.g. Power BI, Jira)')
    ]
    
    for emp_field, proj_field in match_fields:
        emp_val = str(employee[emp_field]).strip().lower().split(';')
        proj_val = str(project[proj_field]).strip().lower().split(',')
        if proj_val and proj_val[0] != 'n/a' and proj_val[0] != '':
            total_count += len(proj_val)
            match_count += sum(1 for val in proj_val if val.strip() in [v.strip() for v in emp_val])
    
    if pd.notna(employee['Certifications']) and str(employee['Certifications']).strip().lower() != 'n/a':
        total_count += 1
        match_count += 1
    
    return (match_count / total_count) * 100 if total_count else 0

def match_employees_to_project(project, available_employees):
    """Match employees to a project based on requirements and update availability."""
    required_people = int(project['Number of Employees Needed'])
    available_employees.loc[:, 'Match %'] = available_employees.apply(
        lambda emp: calculate_match_percentage(emp, project), axis=1
    )


    available_employees = available_employees[available_employees['Match %'] > 0]
    sorted_emps = available_employees.sort_values(by='Match %', ascending=False)
    selected = sorted_emps.head(required_people)
    
    for idx in selected.index:
        available_employees.loc[idx, 'Are you currently assigned to a project?'] = 'Yes'
    
    return selected.copy(), available_employees

def assign_employees_to_projects(projects_df, employee_df):
    """For each project, show top 10 employees with highest match percentage."""
    valid_employees = filter_valid_employees(employee_df)
    project_assignments = {}

    for _, project in projects_df.iterrows():
        # Calculate match % for every employee
        valid_employees.loc[:, 'Match %'] = valid_employees.apply(
            lambda emp: calculate_match_percentage(emp, project), axis=1
        )

        # Sort and get top 10 matches with Match % > 0
        top_matches = valid_employees[valid_employees['Match %'] > 0]
        top_matches = top_matches.sort_values(by='Match %', ascending=False).head(10)

        project_assignments[project['Project Name']] = top_matches[[
            'Name', 'Availability', 'Languages', 'Experience', 'Tools', 'Certifications', 'Match %'
        ]]

    return project_assignments, valid_employees, employee_df

def sanitize_sheet_name(name):
    """Sanitize sheet names for Excel compatibility."""
    invalid_chars = ['\\', '/', '*', '?', ':', '[', ']']
    for char in invalid_chars:
        name = name.replace(char, '-')
    return name[:31]

def to_excel_download(assignments):
    """Convert project assignments to Excel for download."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in assignments.items():
            df_reset = df.reset_index(drop=True)
            safe_name = sanitize_sheet_name(sheet_name)
            df_reset.to_excel(writer, sheet_name=safe_name, index=False)
    output.seek(0)

    #(THIS IS THE CODE FOR DOWNLOADING UPDATED EMPLOYEE DATA) REMOVED FOR NOW AS IT IS NOT USED IN THE CURRENT VERSION
    #    updated_employee_io = BytesIO()  
    #    updated_employee_df.to_excel(updated_employee_io, index=False)
    #    updated_employee_io.seek(0)
    #    st.download_button(
    #        label="📥 Download Updated Employee Data (Excel)",
    #        data=updated_employee_io,
    #        file_name="UpdatedEmployeeData.xlsx",
    #        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    #        help="Download the updated employee data with assignment status."
    #    )

    return output

# -------------------- Streamlit UI --------------------

st.set_page_config(page_title="Workforce Planning Tool", layout="wide")
st.markdown("""
    <style>
    .main { background-color: #f5f5f5; padding: 20px; border-radius: 10px; }
    .stButton>button { background-color: #0066cc; color: white; border-radius: 8px; }
    .stFileUploader { background-color: #ffffff; padding: 10px; border-radius: 8px; border: 1px solid #ddd; }
    .stDataFrame { border: 1px solid #ddd; border-radius: 8px; background-color: #ffffff; }
    .sidebar .sidebar-content { background-color: #e6f0fa; }
    h1 { color: #003087; font-family: 'Arial', sans-serif; }
    h2, h3 { color: #004b87; font-family: 'Arial', sans-serif; }
    </style>
""", unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    # st.image("", use_container_width=True)
    st.header("Workforce Planning Tool")
    st.markdown("**Version 2.0** - Optimized for efficient employee-project matching")
    st.markdown("---")
    st.markdown("### Navigation")
    st.markdown("- Upload Employee Data\n- Upload project requirements\n- View matching results\n- Download assignments")
    st.markdown("---")
    st.markdown("Developed by: Data Analytics Team")

# Main content
st.title("💼 Workforce Planning Tool")
st.markdown("**Match employees to projects based on skills, tools, languages, and certifications.**")
st.markdown("---")

st.header("📁 Upload Employee Data")
emp_file = st.file_uploader("Upload Employee Data Excel File (.xlsx)", type=["xlsx"], help="Upload an Excel file containing employee data.")

# Load employee data
#try:
#except Exception as e:
#    st.error(f"❌ Failed to load employee data: {e}")
#    st.stop()

# Project file uploader
st.header("📁 Upload Project Requirements")
proj_file = st.file_uploader("Upload Project Requirements Excel File (.xlsx)", type=["xlsx"], help="Upload an Excel file containing project requirements.")

if proj_file and emp_file:
    try:
        
        #employee_df = pd.read_excel("InternRecords.xlsx")
        employee_df = pd.read_excel(emp_file)
        # Step 1: Normalize Name1 to lowercase for duplicate detection
        employee_df['Name1_clean'] = employee_df['Name1'].str.strip().str.lower()

        # Step 2: Keep the row with the highest ID for each Name1
        employee_df = employee_df.sort_values(by='Id', ascending=False)
        employee_df = employee_df.drop_duplicates(subset='Name1_clean', keep='first')

        # Step 3: (Optional) Drop the helper column after filtering
        employee_df.drop(columns='Name1_clean', inplace=True)
        employee_df.drop(columns='Name', inplace=True)
        employee_df['Name1'] = employee_df['Name1'].str.title()
        employee_df = employee_df.rename(columns={'Name1': 'Name'})

        project_df = pd.read_excel(proj_file)
        st.success("✅ Project file loaded successfully!")
        
        # Summary metrics
        st.markdown("### 📊 Assignment Summary")
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Projects", len(project_df))
        col2.metric("Not-Assigned Employees", len(employee_df[employee_df['Are you currently assigned to a project?'].str.lower() == 'no']))
        col3.metric("Matching Criteria", "Skills, Tools, Languages, Certifications")
        
        # Run matching logic
        assignments, updated_available_emps, updated_employee_df = assign_employees_to_projects(project_df, employee_df)
        
        st.markdown("### 🧠 Project Assignments")
        for project_name, df in assignments.items():
            st.subheader(f"📌 {project_name}")
            if df.empty:
                st.warning("No matching employees found.")
            else:
                df_display = df.reset_index(drop=True)
                df_display.index += 1  # 📌 Start index from 1
                df_display.index.name = "S.No."  # Optional: Name the index column
                st.dataframe(df_display, use_container_width=True)
        
        # Download results
        st.markdown("### 📥 Download Results")
        excel_data = to_excel_download(assignments)
        st.download_button(
            label="📥 Download Project Assignments (Excel)",
            data=excel_data,
            file_name="ProjectAssignments.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Download the matched employee assignments for all projects."
        )
        
    except Exception as e:
        st.error(f"❌ Error processing project file: {e}")
else:
    st.info("⬆️ Please upload a project requirements and employee data Excel file to begin.")