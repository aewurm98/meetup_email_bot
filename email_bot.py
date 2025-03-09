import pandas as pd
import random
import smtplib
from email.mime.text import MIMEText

# Step 1: Read the list of students from a CSV file and filter relevant students
def read_students(file_path, max_selections):
    df = pd.read_csv(file_path)
    # Assuming there's a column 'Selections' to indicate the number of times a student has been selected for a prior meetup
    df['Number of Selections'].fillna(0, inplace=True)
    relevant_students = df[(df['Number of Selections'] < max_selections) & (df['Student Email'].notna())]
    return relevant_students

# Update the selected students in the CSV file by incrementing their selection count
def update_selected_students(file_path, selected_students):
    df = pd.read_csv(file_path)
    df.loc[df['Student Email'].isin(selected_students), 'Number of Selections'] += 1
    df.to_csv(file_path, index=False)

# Step 2: Randomly generate groups of a specified number N from the relevant subgroup
def generate_groups(students, group_size, use_sections=False):
    relevant_students = students[students['Selected'] == False]
    all_students = students['Student Email'].tolist()
    
    if use_sections:
        sections = relevant_students['Section'].unique()
        groups = []
        
        while len(relevant_students) >= group_size:
            group = []
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
            groups.append(group)
        
        if not relevant_students.empty():
            remaining_group = relevant_students['Student Email'].tolist()
            groups.append(remaining_group)
        
        return groups
    else:
        student_emails = relevant_students['Student Email'].tolist()
        random.shuffle(student_emails)
        
        groups = [student_emails[i:i + group_size] for i in range(0, len(student_emails), group_size)]
        
        # If the last group is smaller than the group size, fill it with additional students
        if len(groups[-1]) < group_size:
            remaining_students = [email for email in all_students if email not in student_emails]
            random.shuffle(remaining_students)
            while len(groups[-1]) < group_size and remaining_students:
                groups[-1].append(remaining_students.pop())
        
        return groups

def update_selected_students(file_path, selected_students):
    df = pd.read_excel(file_path)
    df.loc[df['Student Email'].isin(selected_students), 'Selected'] = True
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
        file_path = 'MBA_Student_Test_Email_Dataset.csv'
        group_size = 10
        subject = 'Dinner Invitation'
        body_template = 'Hey everyone,\n\nYou are invited to a small-group dinner with the following classmates: {group}.\n\nCheers,\nYour Class Coordinator'
        smtp_server = 'smtp.example.com'
        smtp_port = 587
        sender_email = 'your_email@example.com'
        sender_password = 'your_password'
        
        students = read_students(file_path, max_selections=3)
        groups = generate_groups(students, group_size)
        
        # Output the dataframe of students who will be contacted
        selected_students = [email for group in groups for email in group]
        selected_students_df = students[students['Student Email'].isin(selected_students)]
        print(selected_students_df)
        
        # Update the selected students in the spreadsheet
        update_selected_students(file_path, selected_students)

# Example usage
# if __name__ == "__main__":
#     file_path = 'MBA_Student_Test_Email_Dataset.csv'
#     group_size = 10
#     subject = 'Dinner Invitation'
#     body_template = 'Hey everyone,\n\nYou are invited to a small-group dinner with the following classmates: {group}.\n\nCheers,\nYour Class Coordinator'
#     smtp_server = 'smtp.example.com'
#     smtp_port = 587
#     sender_email = 'your_email@example.com'
#     sender_password = 'your_password'
    
#     students = read_students(file_path, max_selections=3)
#     groups = generate_groups(students, group_size)
#     send_emails(groups, subject, body_template, smtp_server, smtp_port, sender_email, sender_password)
    
#     # Update the selected students in the spreadsheet
#     selected_students = [email for group in groups for email in group]
#     update_selected_students(file_path, selected_students)

# Reset the Number of Selections to zero for all students
def reset_selections(file_path):
    df = pd.read_csv(file_path)
    df['Number of Selections'] = 0
    df.to_csv(file_path, index=False)