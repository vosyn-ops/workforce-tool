Workforce Planning Tool

Overview

The Workforce Planning Tool is a Streamlit-based application designed to match employees to projects based on their skills, tools, languages, and certifications. It processes employee and project data from Excel files, calculates match percentages, and generates optimized employee-project assignments. The tool is developed by the Data Analytics Team to streamline workforce allocation.

Features

•	Data Input: Upload employee and project requirement data in Excel (.xlsx) format.

•	Matching Algorithm: Matches employees to projects based on:

    o	Languages (e.g., Python, SQL)
    
    o	Tools (e.g., PowerBI, Jira)
    
    o	Skills/Experience (e.g., Data Cleaning, ETL)
    
    o	Certifications
  
•	Match Percentage: Calculates a match score for each employee-project pair.

•	Filtering: Filters employees who are unassigned and available beyond a 15-day threshold.

•	Output: Displays top 10 matching employees per project and provides downloadable Excel results.

•	User Interface: Built with Streamlit for a clean, interactive experience with metrics and visualizations.

Prerequisites

•	Python: Version 3.8 or higher

•	Dependencies:

    o	streamlit
  
    o	pandas
  
    o	openpyxl
  
•	Install dependencies using:

    pip install -r requirements.txt
  
Installation

1.	Clone the repository:
   
        git clone https://github.com/your-username/workforce-planning-tool.git
2.	Navigate to the project directory:
   
        cd workforce-planning-tool
3.	Create a virtual environment (optional but recommended):
   
        python -m venv venv
        source venv/bin/activate  # On Windows: venv\Scripts\activate
4.	Install the required packages:
   
        pip install streamlit pandas openpyxl
5.	Run the application:
   
        streamlit run version2realdata.py

Usage

1.	Prepare Input Files:
   
    o	Employee Data (Excel): Must include columns like Id, Name, Availability, Languages, Tools, Experience, Certifications, Are you currently assigned to a project?, and         Please Enter Your End Date.
  	
    o	Project Requirements (Excel): Must include columns like Project Name, Languages proficiency required, Tools, Skills Required, and Number of Employees Needed.
  	
    o	Example files: InternRecords.xlsx and Project_Requirements_Form_with_Dropdowns.xlsx.
  	
2.	Run the Application:
   
    o	Launch the app using the command above.
  	
    o	Open the provided URL (e.g., http://localhost:8501) in your browser.
  	
3.	Upload Files:
   
    o	Upload the employee and project Excel files via the Streamlit interface.
  	
4.	View Results:
   
    o	The tool displays:
  	
    o	Summary metrics (total projects, unassigned employees, matching criteria).
  	
    o	Top 10 employee matches per project with match percentages.
  	
    o	Download the assignments as an Excel file (ProjectAssignments.xlsx).

Code Structure

•	version2realdata.py: Main application script containing:

  o	Helper functions for filtering employees, calculating match percentages, and assigning employees to projects.
  
  o	Streamlit UI for file uploads, result display, and Excel download.
  
•	Key Functions:

  o	filter_valid_employees: Filters unassigned employees with end dates beyond 15 days.
  
  o	calculate_match_percentage: Computes match score based on skills, tools, languages, and certifications.
  
  o	match_employees_to_project: Matches employees to a single project.
  
  o	assign_employees_to_projects: Generates top 10 matches for all projects.
  
  o	to_excel_download: Exports assignments to Excel.
  
  o	sanitize_sheet_name: Ensures Excel-compatible sheet names.

Input File Requirements

•	Employee Data:

  o	Columns: Id, Name, Availability, Languages, Tools, Experience, Certifications, Are you currently assigned to a project?, Please Enter Your End Date.
  
  o	Values in Languages, Tools, and Experience should be semicolon-separated.
  
•	Project Requirements:

  o	Columns: Project Name, Languages proficiency required, Tools, Skills Required, Number of Employees Needed.
  
  o	Values in Languages, Tools, and Skills Required should be comma-separated.

