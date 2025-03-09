import pandas as pd
import random
import smtplib
from email.mime.text import MIMEText

# Step 1: Read the list of students from a spreadsheet and filter relevant students
def read_students(file_path):
    df = pd.read_excel(file_path)
    # Assuming there's a column 'Selected' to indicate if a student has already been selected for a prior meetup
    relevant_students = df[df['Selected'] == False]
    return relevant_students

# Step 2: Randomly generate groups of a specified number N from the relevant subgroup
def generate_groups(students, group_size, use_sections=False):
    relevant_students = students[students['Selected'] == False]
    
    if use_sections:
        sections = relevant_students['Section'].unique()
        groups = []
        
        while len(relevant_students) >= group_size:
            group = []
            for section in sections:
                section_students = relevant_students[relevant_students['Section'] == section]
                if not section_students.empty:
                    selected_student = section_students.sample(n=1)
                    group.append(selected_student['Email'].values[0])
                    relevant_students = relevant_students.drop(selected_student.index)
                    if len(group) == group_size:
                        break
            groups.append(group)
        
        if not relevant_students.empty:
            remaining_group = relevant_students['Email'].tolist()
            groups.append(remaining_group)
        
        return groups
    else:
        student_emails = relevant_students['Email'].tolist()
        random.shuffle(student_emails)
        return [student_emails[i:i + group_size] for i in range(0, len(student_emails), group_size)]

def update_selected_students(file_path, selected_students):
    df = pd.read_excel(file_path)
    df.loc[df['Email'].isin(selected_students), 'Selected'] = True
    df.to_excel(file_path, index=False)

# Step 3: Draft and send emails to the group of recipients
def send_emails(groups, subject, body_template, smtp_server, smtp_port, sender_email, sender_password):
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(sender_email, sender_password)
    
    for group in groups:
        body = body_template.format(group=', '.join(group))
        msg = MIMEText(body)
        msg['Subject'] = subject
        msg['From'] = sender_email
        msg['To'] = ', '.join(group)
        
        server.sendmail(sender_email, group, msg.as_string())
    
    server.quit()

# Example usage
if __name__ == "__main__":
    file_path = 'students.xlsx'
    group_size = 10
    subject = 'Dinner Invitation'
    body_template = 'Hey everyone,\n\nYou are invited to a small-group dinner with the following classmates: {group}.\n\nCheers,\nYour Class Coordinator'
    smtp_server = 'smtp.example.com'
    smtp_port = 587
    sender_email = 'your_email@example.com'
    sender_password = 'your_password'
    
    students = read_students(file_path)
    groups = generate_groups(students, group_size)
    send_emails(groups, subject, body_template, smtp_server, smtp_port, sender_email, sender_password)
    
    # Update the selected students in the spreadsheet
    selected_students = [email for group in groups for email in group]
    update_selected_students(file_path, selected_students)