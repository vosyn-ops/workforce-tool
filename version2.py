import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO

# -------------------- BACKEND LOGIC --------------------

def filter_valid_employees(df):
    today = pd.to_datetime(datetime.today().date())
    threshold_date = today + timedelta(days=15)
    df['End Date'] = pd.to_datetime(df['End Date'], errors='coerce')
    df = df[df['End Date'] > threshold_date].copy()
    df['Available Hours'] = df['Availability'].apply(lambda x: 40 if str(x).strip().lower() == 'full-time' else 20)
    df['Original Availability'] = df['Available Hours']
    return df

def calculate_match_percentage(employee, project):
    match_count = 0
    total_count = 0
    for field in ['Languages', 'Skills', 'Tools']:
        emp_val = str(employee[field]).strip().lower()
        proj_val = str(project[field]).strip().lower()
        if pd.notna(proj_val) and proj_val != '':
            total_count += 1
            if proj_val in emp_val or emp_val in proj_val:
                match_count += 1
    if pd.notna(employee['Certifications']) and str(employee['Certifications']).strip() != '':
        total_count += 1
        match_count += 1
    return (match_count / total_count) * 100 if total_count else 0

def match_employees_to_project(project, available_employees):
    required_people = int(project['Number of People Required'])
    project_hours = int(project['Hours Required'])

    available_employees['Match %'] = available_employees.apply(
        lambda e: calculate_match_percentage(e, project), axis=1
    )
    sorted_emps = available_employees.sort_values(by='Match %', ascending=False)
    exact_matches = sorted_emps[sorted_emps['Match %'] == 100]
    selected = exact_matches.head(required_people) if len(exact_matches) >= required_people else sorted_emps.head(required_people)
    #print(selected)
    if not selected.empty:
        hours_per_employee = project_hours // len(selected)
        for idx in selected.index:
            available_employees.loc[idx, 'Current Projects'] += 1
            available_employees.loc[idx, 'Current Availability'] = 0
            #available_employees.loc[idx, 'Available Hours'] = 0
            #available_employees.loc[idx, 'Booked Hours'] = hours_per_employee

    return selected.copy(), available_employees

def assign_employees_to_projects(projects_df, employee_df):
    employee_df['Current Projects'] = employee_df['Current Projects'].fillna(0).astype(int)
    employee_df['Current Availability'] = employee_df['Current Availability'].fillna(0).astype(int)
    available_employees = filter_valid_employees(employee_df)

    project_assignments = {}

    for _, project in projects_df.iterrows():
        print(available_employees.shape)
        matched_emps, available_employees = match_employees_to_project(project, available_employees)
        available_employees = available_employees.drop(index=matched_emps.index)
        project_assignments[project['Project Name']] = matched_emps[[
            'Name', 'Availability', 'Languages', 'Skills', 'Tools',
            'Certifications', 'Match %'
        ]]

    return project_assignments, available_employees, employee_df

# Sanitize sheet names for Excel
def sanitize_sheet_name(name):
    invalid_chars = ['\\', '/', '*', '?', ':', '[', ']']
    for char in invalid_chars:
        name = name.replace(char, '-')
    return name[:31]  # Excel sheet names must be ≤ 31 chars

# Convert project assignment dictionary to Excel download
def to_excel_download(assignments):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in assignments.items():
            df_reset = df.reset_index(drop=True)  # Remove index
            print(df_reset)
            safe_name = sanitize_sheet_name(sheet_name)
            df_reset.to_excel(writer, sheet_name=safe_name, index=False)
    output.seek(0)
    return output

# -------------------- STREAMLIT UI --------------------

st.set_page_config(page_title="Project Skill Matcher", layout="wide")
st.title("💼 Project Skill Matcher")
st.markdown("Match employees to projects based on **skills, tools, languages, and certifications**.")

# Load employee data
try:
    employee_df = pd.read_excel("employee_datav2.xlsx")
except Exception as e:
    st.error(f"❌ Failed to load employee data: {e}")
    st.stop()

# Upload project requirements
st.header("📁 Upload Project Requirements (.xlsx)")
proj_file = st.file_uploader("Upload Project Requirements Excel File", type=["xlsx"])

if proj_file:
    try:
        project_df = pd.read_excel(proj_file)
        st.success("✅ Project file loaded successfully!")

        # Run matching logic
        assignments, updated_available_emps, updated_employee_df = assign_employees_to_projects(project_df, employee_df)

        st.markdown("### 🧠 Project Assignments")
        for project_name, df in assignments.items():
            st.subheader(f"📌 {project_name}")
            if df.empty:
                st.warning("No matching employees found.")
            else:
                df_display = df.reset_index(drop=True) 
                st.dataframe(df_display, use_container_width=True)

        # ----- Download Results -----
        st.markdown("### 📥 Download Matched Results")

        # 1. Download assignments by project
        excel_data = to_excel_download(assignments)
        st.download_button(
            label="📥 Download Project Assignments (Excel)",
            data=excel_data,
            file_name="ProjectAssignments.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # 2. Download updated employee sheet
        updated_employee_io = BytesIO()
        updated_employee_df.to_excel(updated_employee_io, index=False)
        updated_employee_io.seek(0)

        st.download_button(
            label="📥 Download Updated Employee Data (Excel)",
            data=updated_employee_io,
            file_name="updated_employee_data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        st.error(f"❌ Error processing project file: {e}")
else:
    st.info("⬆️ Please upload a project requirements Excel file to begin.")

