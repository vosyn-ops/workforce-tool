import pandas as pd
import random
from faker import Faker

# Initialize Faker
fake = Faker()

languages = ["Julia", "Python", "R", "SQL", "Javascript", "NoSQL"]
tools = ["Tableau", "Power BI", "Git", "VS Code", "Microsoft Excel", "Looker", "AWS", "Azure", "GCP"]
skills = ["Data Cleaning", "Data Visualization", "Data Analysis", "Data Collection", 
          "Data Manipulation and Management", "Statitical Analysis", "Database Management"]
# Define possible departments and required skills
departments = {
    "Data Science": ["Python", "SQL", "Machine Learning", "Power BI", "Statistics"],
    "Web Development": ["HTML", "CSS", "JavaScript", "React", "Node.js"],
    "Cybersecurity": ["Network Security", "Linux", "Penetration Testing", "Firewall Management"],
    "Cloud Computing": ["AWS", "Azure", "Docker", "Kubernetes"],
    "Software Development": ["Java", "Spring Boot", "C#", ".NET", "Microservices"],
    "Project Management": ["Agile", "Scrum", "Risk Management", "Stakeholder Communication"]
}

Hours = [100, 150, 200, 250, 300, 350, 400]
# Generate 10 random projects
project_data = []
for i in range(10):
    project_name = f"Project {i+1}: {fake.bs().title()}"  # Generates a business-sounding project name
    language = random.choice(languages)
    skill = random.choice(skills)
    tool = random.choice(tools)
    #department = random.choice(list(departments.keys()))  # Random department
    #required_skills = ", ".join(random.sample(departments[department], k=random.randint(2, 4)))  # Pick 2-4 skills
    number_of_people = random.randint(1, 5)  # Number of people required for the project
    hours_required = random.choice(Hours)  # Total hours to complete the project
    project_data.append([project_name, language, skill, tool, number_of_people, hours_required])

# Create a DataFrame
df = pd.DataFrame(project_data, columns=["Project Name", "Languages", "Skills", "Tools", "Number of People Required", "Hours Required"])

# Save to Excel
df.to_excel("project_requirementsv2.xlsx", index=False)

print("10 random project requirements have been created successfully in 'project_requirementsv2.xlsx'.")