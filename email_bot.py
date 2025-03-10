import pandas as pd
import random
import smtplib
from email.mime.text import MIMEText
import os
from dotenv import find_dotenv, load_dotenv

dotenv_path = find_dotenv()
load_dotenv(dotenv_path, override=True)

# Step 1: Read the list of students from a CSV file and filter relevant students
def read_students(file_path):
    df = pd.read_csv(file_path)
    return df

# Step 2: Randomly generate groups of a specified number N from the relevant subgroup
def generate_group(students, group_size, use_sections=False):
    
    # Define criteria for relevant students
    students.fillna({'Number of Selections':0}, inplace=True)
    students['Number of Selections'] = students['Number of Selections'].astype(int)
    max_selections = students['Number of Selections'].max()
    min_selections = students['Number of Selections'].min()
    
    if max_selections == min_selections:
        relevant_students = students[students['Student Email'].notna()]
    else:
        relevant_students = students[((students['Number of Selections'] == 0) | (students['Number of Selections'] < max_selections)) & (students['Student Email'].notna())]

    all_students = students[students['Student Email'].notna()]['Student Email'].tolist()
    
    if use_sections:
        sections = relevant_students['Section'].unique()
        group = []
        
        while len(relevant_students) >= group_size and len(group) < group_size:
            sections = list(sections)
            random.shuffle(sections)
            for section in sections:
                section_students = relevant_students[relevant_students['Section'] == section]
                if not section_students.empty:
                    selected_student = section_students.sample(n=1)
                    group.append(selected_student['Student Email'].values[0])
                    relevant_students = relevant_students.drop(selected_student.index)
                    if len(group) == group_size:
                        break
        
        if not relevant_students.empty and len(group) < group_size:
            remaining_group = relevant_students[relevant_students['Student Email'].notna()]['Student Email'].tolist()
            needed = group_size - len(group)
            group.extend(remaining_group[:needed])
        
        return group
    else:
        student_emails = relevant_students['Student Email'].tolist()
        random.shuffle(student_emails)
        
        group = random.sample(student_emails, group_size)
        
        # If the last group is smaller than the group size, fill it with additional students
        if len(group) < group_size:
            remaining_students = [email for email in all_students if email not in student_emails]
            random.shuffle(remaining_students)
            while len(group) < group_size and remaining_students:
                group.append(remaining_students.pop())
        
        return group

# Update the selected students in the CSV file by incrementing their selection count
def update_selected_students(file_path, selected_students):
    df = pd.read_csv(file_path)
    df.fillna({'Number of Selections':0}, inplace=True)
    df.loc[df['Student Email'].isin(selected_students), 'Number of Selections'] += 1
    df.to_csv(file_path, index=False)
    return df
    
    # Ensure the group size is not exceeded
    def validate_group_size(group, group_size):
        for sub_group in group:
            if isinstance(sub_group, list) and len(sub_group) > group_size:
                raise ValueError("Group size exceeded")

    validate_group_size(group, group_size)

# Step 3: Draft and send emails to the group of recipients
def send_emails(group, subject, body_template, smtp_server, smtp_port, sender_email, sender_password):
    # server = smtplib.SMTP(smtp_server, smtp_port)
    # server.starttls()
    # server.login(sender_email, sender_password)
    
    for individual in group:
        body = body_template.format(individual=', '.join(individual))
        msg = MIMEText(body)
        msg['Subject'] = subject
        msg['From'] = sender_email
        msg['To'] = ', '.join(individual)

        # Mock sending emails
        print(f"Sending email to: {individual}")
        
    #     server.sendmail(sender_email, individual, msg.as_string())
    
    # server.quit()

# Reset the Number of Selections to zero for all students
def reset_selections(file_path):
    df = pd.read_csv(file_path)
    df['Number of Selections'] = 0
    df.to_csv(file_path, index=False)
    return df

if __name__ == "__main__":
    file_path = 'student_data.csv'
    group_size = 5
    use_sections = True
    subject = "Group Assignment"
    body_template = "Dear Student,\n\nYou have been assigned to the following group:\n\n{individual}\n\nBest regards,\nYour Instructor"
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    sender_email = 'aewurm98@gmail.com'
    sender_password = os.getenv('EMAIL_PASSWORD')

    # Step 1: Read students from CSV
    students = read_students(file_path)

    # Step 2: Generate group
    group = generate_group(students, group_size, use_sections)

    # Step 3: Send emails
    send_emails(group, subject, body_template, smtp_server, smtp_port, sender_email, sender_password)

    # Update selected students in CSV
    update_selected_students(file_path, group)

    # Reset selections
    reset_selections(file_path)