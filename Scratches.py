# (Imports and input_date function remain the same)

def process_email(message, template_file):
    # (Everything inside process_email function remains the same)

def get_emails(subject, start_date, end_date):
    # Connect to Outlook
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Get emails with the subject
    folder = outlook.GetDefaultFolder(6)  # Inbox folder

    messages = folder.Items
    messages.Sort("[ReceivedTime]", True)  # Sort messages by received time in descending order
    restriction = f"[ReceivedTime] >= '{start_date.strftime('%m/%d/%Y %I:%M %p')}' AND [ReceivedTime] <= '{end_date.strftime('%m/%d/%Y %I:%M %p')}'"
    restricted_messages = messages.Restrict(restriction)

    return [msg for msg in restricted_messages if subject in msg.Subject]

# (Rest of the script remains the same)

emails = get_emails(subject, start_date, end_date)

# (Rest of the script remains the same)

for selected_email_index in selected_email_indices:
    message = emails[selected_email_index]
    process_email(message, template_file)

# The rest of the code in the previous version should be removed, as it duplicates the functionality already present in the `process_email` function.
