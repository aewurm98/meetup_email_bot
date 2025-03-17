import meeting_planner
import os

meetup_info = meeting_planner.parse_and_delete_emails_with_subject("aewurm98@gmail.com", os.getenv('EMAIL_PASSWORD'), "Weekly Dinner Bot: Details Request")

date, time, location = meeting_planner.convert_parsed_details_to_strings(meetup_info)

message_body = meeting_planner.structure_email(date, time, location, "It is a fun location!",["It's cold outside","Bigfoot was spotted","Lemons new weight loss tool?"])

meeting_planner.send_meeting_invite("aewurm98@gmail.com", os.getenv('EMAIL_PASSWORD'), "aew98@cornell.edu","Small Group Dinner Invite", message_body)
