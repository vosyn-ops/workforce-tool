import pandas as pd
import random
from faker import Faker
from datetime import datetime, timedelta


# Initialize Faker
fake = Faker()

availabilities = ["Part-Time", "Full-Time"]
languages = ["Julia", "Python", "R", "SQL", "Javascript", "NoSQL"]
tools = ["Tableau", "Power BI", "Git", "VS Code", "Microsoft Excel", "Looker", "AWS", "Azure", "GCP"]
skills = ["Data Cleaning", "Data Visualization", "Data Analysis", "Data Collection", 
          "Data Manipulation and Management", "Statitical Analysis", "Database Management"]
time_zones = ["EST", "CST", "MST", "PST", "GMT", "WAT", "CET", "IST", "AEST"]
certifications = {
    "Google Data Analytics Professional Certificate": ["SQL", "R", "Tableau", "Excel", 
                                                       "Data Cleaning", "Data Visualization"],
    "IBM Data Analyst Professional Certificate": ["Python", "SQL", "Excel", "Power BI", "Data Analysis", 
                                                  "Data Manipulation"],
    "Microsoft Certified: Power BI Data Analyst Associate": ["Power BI", "SQL", "Excel", "Azure", 
                                                             "Data Visualization", "Data Analysis"],
    "SAS Statistical Business Analyst Professional Certificate": ["SAS", "Statistical Analysis", 
                                                                  "Data Manipulation", "Data Cleaning"],
    "CompTIA Data+": ["Database Management", "Data Visualization", "SQL", "Python", "Data Analysis"],
    "Meta Data Analyst Professional Certificate": ["Python", "SQL", "Data Visualization", "Excel", 
                                                   "Power BI", "Data Manipulation"],
    "Certified Analytics Professional (CAP)": ["Data Analysis", "Machine Learning Basics", "Statistical Analysis",
                                                "Business Acumen"],
    "Tableau Desktop Certified Associate": ["Tableau", "Data Visualization", "Data Cleaning", "Data Analysis"],
    "Microsoft Certified: Azure Enterprise Data Analyst Associate": ["Azure", "SQL", "Power BI", "Data Management"],
    "Cloudera Certified Associate (CCA) Data Analyst": ["SQL", "NoSQL", "Database Management", "Data Collection"]
}
certification_probability = 0.3 
full_time = [10,20,30,40]
part_time = [5,10,15,20]


employee_data = []
used_names = set()
used_emails = set()

while len(employee_data) < 100:
    first_name = fake.first_name()
    last_name = fake.last_name()
    full_name = f"{first_name} {last_name}"

    # Ensure unique names
    if full_name in used_names:
        continue
    used_names.add(full_name)

    # Generate email based on name
    email_base = f"{first_name.lower()}.{last_name.lower()}"
    email = f"{email_base}@example.com"

    # Ensure unique emails
    if email in used_emails:
        continue
    used_emails.add(email)
    availability = random.choice(availabilities)
    start_date = fake.date_between(start_date='-1y', end_date='today')

    if availability == "Full-Time":
        end_date = start_date + timedelta(days=90)  # 3 months
        current_projects = 0
        current_availability = 40
    else:
        end_date = start_date + timedelta(days=180)  # 6 months
        current_projects = 0
        current_availability = 20

    language = random.choice(languages)
    skill = random.choice(skills)
    tool = random.choice(tools)
    tz = random.choice(time_zones)
    cert = ""
    matching_certifications = [
    cert for cert, cert_attributes in certifications.items()
    if language in cert_attributes or skill in cert_attributes or tool in cert_attributes]

    if random.random() < certification_probability:
        assigned_certification = random.choice(matching_certifications) if matching_certifications else "No suitable certification found."
        cert = assigned_certification

    employee_data.append([full_name, email, availability, current_projects, current_availability, start_date, 
                          end_date, tz, language, skill, tool, cert])

# Create a DataFrame
df = pd.DataFrame(employee_data, columns=["Name", "Email", "Availability", "Current Projects", "Current Availability",
    "Start Date", "End Date", "Time Zone", "Languages", "Skills", "Tools", "Certifications" 
])

# Save to Excel
df.to_excel("employee_datav2.xlsx", index=False)

print("100 dummy employee records have been created successfully in 'employee_datav2.xlsx'.")
