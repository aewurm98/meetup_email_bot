import pandas as pd
import random
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
from dotenv import find_dotenv, load_dotenv
import imaplib
import email
from imapclient import IMAPClient
import pyzmail
from datetime import datetime

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
        relevant_students = students[students['Number of Selections'] < max_selections & students['Student Email'].notna()]

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

        # Ensure the group size is not exceeded
        if len(group) > group_size:
            raise ValueError("Group size exceeded")
        
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
        
        # Ensure the group size is not exceeded
        if len(group) > group_size:
            raise ValueError("Group size exceeded")

        return group

# Update the selected students in the CSV file by incrementing their selection count
def update_selected_students(file_path, selected_students):
    df = pd.read_csv(file_path)
    df.fillna({'Number of Selections':0}, inplace=True)
    df.loc[df['Student Email'].isin(selected_students), 'Number of Selections'] += 1
    df.to_csv(file_path, index=False)
    return df

# Create a dictionary of students with their names and emails
def create_student_dict(file_path, group):
    df = pd.read_csv(file_path)
    student_dict = df[df['Student Email'].isin(group)][['Student Name', 'Student Email']].set_index('Student Email').to_dict('index')
    return student_dict

# Reset the Number of Selections to zero for all students
def reset_selections(file_path):
    df = pd.read_csv(file_path)
    df['Number of Selections'] = 0
    df.to_csv(file_path, index=False)
    return df

def request_meeting_info(sender_email, sender_password, recipient_email, subject):
    # Create the email
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject

    # Create the body of the email
    body = """
    Where and when would you like to meet for the upcoming small group dinner?

    Please reply to this email in the following format:
    Date: YYYY-MM-DD (zero-padded numbers)
    Time: HH:MM AM/PM (standard time)
    Location: [Location Name]

    Weekly Dinner Bot \U0001F916 \U0001F37D
    """
    msg.attach(MIMEText(body, 'plain'))

    # Send the email
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        text = msg.as_string()
        server.sendmail(sender_email, recipient_email, text)
        server.quit()
        print("Email requesting meeting info sent successfully!")
    except Exception as e:
        print(f"Failed to send email: {e}")

def parse_and_delete_emails_with_subject(email_address, email_password, subject_to_delete):
    parsed_details = []
    try:
        # Connect to the server
        with IMAPClient('imap.gmail.com') as client:
            client.login(email_address, email_password)
            client.select_folder('INBOX')

            # Search for emails with the specified subject
            messages = client.search(['SUBJECT', subject_to_delete])
            # print(f"Found {len(messages)} messages with subject '{subject_to_delete}'")

            if not messages:
                print("No messages found with the specified subject.")
                return parsed_details

            # Fetch the emails and sort by date in descending order
            fetched_messages = client.fetch(messages, ['RFC822', 'INTERNALDATE'])
            sorted_messages = sorted(fetched_messages.items(), key=lambda item: item[1][b'INTERNALDATE'], reverse=True)

            for msgid, data in sorted_messages:
                msg = pyzmail.PyzMessage.factory(data[b'RFC822'])
                # print(f"Processing message ID: {msgid}")
                
                if subject_to_delete in msg.get_subject():
                    # print(f"Subject matches: {msg.get_subject()}")

                    if msg.text_part:
                        body = msg.text_part.get_payload().decode(msg.text_part.charset)
                        # print("Email body: \n", body)
                        
                        # Parse the email body for Date, Time, and Location
                        date = None
                        time = None
                        location = None
                        for line in body.split('\n'):
                            if line.startswith("Date:"):
                                date = line.split("Date:")[1].strip()
                            elif line.startswith("Time:"):
                                time = line.split("Time:")[1].strip()
                            elif line.startswith("Location:"):
                                location = line.split("Location:")[1].strip()
                        
                        parsed_details.append({
                            'date': date,
                            'time': time,
                            'location': location
                        })
                        # print(f"Parsed details - Date: {date}, Time: {time}, Location: {location}")
                        break  # Only process the most recent email
                #     else:
                #         print("No text part found in the email.")
                # else:
                #     print(f"Subject does not match: {msg.get_subject()}")

            # Delete the emails
            client.delete_messages(messages)
            client.expunge()
            print(f"Connection established successfully!")

            print(parsed_details)

            return parsed_details

    except Exception as e:
        print(f"Process Failed: {e}")

def convert_parsed_details_to_strings(parsed_details):
    if not parsed_details:
        return None, None, None
    
    # Convert date and time into single timestamp
    date_str = parsed_details[0]['date']
    time_str = parsed_details[0]['time']
    location_str = parsed_details[0]['location']

    # Convert date and time into a single string
    timestamp_str = f"{date_str} {time_str}"
    timestamp = datetime.strptime(timestamp_str, '%Y-%m-%d %I:%M %p')

    date_str = timestamp.strftime('%A, %B %d')
    time_str = timestamp.strftime('%I:%M %p')
    location_str = f"{parsed_details[0]['location']}"
    
    return date_str, time_str, location_str

def structure_email(date_str, time_str, location_str, location_details, conversation_topics):
    # Create the body of the email
    body = f"""
    Hi everyone,

    I'm organizing a randomly-generated small-group dinner series within the 2027 class every week of the semester and you're invited! This email is automated but feel free to reply with any questions.

    Details below, look forward to seeing you there if you can make it!

    Date: {date_str}
    Time: {time_str}
    Location: {location_str}

    """
    
    # Add location details
    body += f"""
    Location Details:
    {location_details}

    """

    # Add relevant conversation topics
    body += f"""
    Quick ChatGPT-generated list of recent news if getting to know your classmates isn't enough of a conversation starter \U0001F609:\n
    """

    for topic in conversation_topics:
        body += f"- {topic}\n"
    
    return body

def send_meeting_invite(sender_email, sender_password, recipient_email, subject, body):
    # Create the email
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient_email
    msg['Subject'] = subject

    # Attach the body of the email
    msg.attach(MIMEText(body, 'plain'))

    # Send the email
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, sender_password)
        text = msg.as_string()
        server.sendmail(sender_email, recipient_email, text)
        server.quit()
        print("Email sent successfully!")
    except Exception as e:
        print(f"Failed to send email: {e}")

# Example usage
if __name__ == "__main__":
    sender_email = "your_email@gmail.com"
    sender_password = "your_password"
    recipient_email = "recipient_email@gmail.com"
    subject = "Meeting Invite"
    meeting_details = "Date: 2023-10-15\nTime: 10:00 AM\nLocation: Conference Room A"
    group_members = ["Alice", "Bob", "Charlie", "David"]

    send_meeting_invite(sender_email, sender_password, recipient_email, subject, meeting_details, group_members)