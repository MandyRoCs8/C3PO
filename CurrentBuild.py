import openpyxl
from docx import Document
from tkinter import Tk, filedialog, Toplevel, Listbox, Button, END, EXTENDED
import datetime
import win32com.client
import pymsgbox
import os
from tkcalendar import Calendar  # Add this import
selected_email_indices = ()


def input_date(prompt):
    def on_select_date():
        date_obj = calendar.selection_get()
        date_dialog.quit()

    date_dialog = Toplevel(root)
    date_dialog.title(prompt)
    date_dialog.lift()
    date_dialog.grab_set()
    calendar = Calendar(date_dialog, selectmode="day", year=datetime.datetime.now().year, month=datetime.datetime.now().month, day=datetime.datetime.now().day)
    calendar.pack(padx=10, pady=10)

    select_button = Button(date_dialog, text="Select", command=on_select_date)
    select_button.pack(pady=5)

    date_dialog.mainloop()
    date_dialog.destroy()

    return calendar.selection_get()


def process_email(message, template_file):
    # Load Word document
    doc = Document(template_file)

    # Extract data from email
    attachments = message.Attachments
    excel_file = ''
    for attachment in attachments:
        if attachment.FileName.endswith('.xlsx'):
            attachment.SaveAsFile(os.path.join(os.getcwd(), attachment.FileName))
            excel_file = attachment.FileName
            print("Attachment saved to '{}'".format(excel_file))
            break

    if not excel_file:
        print("No Excel file found in attachments.")
        return

    # Load Excel data
    wb = openpyxl.load_workbook(excel_file)
    ws = wb.active

    # Create dictionary with labels and their corresponding cell locations
    labels = {
        'Cherwell': (2, 8),
        'CM Ticket': (3, 8),
        'Anticipated GL Date': (4, 8),
        'Short Summary': (5, 8),
        'Reason': (8, 8),
        'Value': (9, 8),
        'Impact': (10, 8),
        'Additional Information': (11, 8)
    }

    # Create dictionary to store values for each label
    values = {}

    # Extract values from Excel file
    for label, cell in labels.items():
        value = ws.cell(row=cell[0], column=cell[1]).value
        values[label] = value

        # Find the paragraph with the label in the Word document and add the value
        for paragraph in doc.paragraphs:
            if label in paragraph.text:
                # Add the value from the Excel file to the end of the paragraph
                paragraph.add_run(" " + str(value))
                print(f"Updated paragraph text: '{paragraph.text}'")

        # Save the output file with a new filename
    output_file = f"{os.path.splitext(template_file)[0]}_{message.Subject}_{message.ReceivedTime.strftime('%Y%m%d_%H%M%S')}.docx"
    doc.save(output_file)

    # Comment out the following line to stop automatically opening the output file
    # os.startfile(output_file)

    print("Data has been extracted and saved to the Word template.")
    # Delete temp Excel file
    os.remove(excel_file)


def on_select_emails():
    global selected_email_indices
    selected_email_indices = email_listbox.curselection()
    email_selection_dialog.quit()


# Show message box
pymsgbox.alert("Please choose a Word template file in the upcoming dialog.", "Choose Word Template File")

# Choose Word template file
root = Tk()
root.withdraw()
template_file = filedialog.askopenfilename(title="Choose Word Template File", filetypes=[("Word files", "*.docx")])
if not template_file:
    print("No Word template file selected.")
    exit()

# Load Word document
doc = Document(template_file)

# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Get the date range from the user
print("Please enter the date range to search for emails with the subject 'C3PO'.")
start_date = input_date("Enter the start date (MM/DD/YYYY): ")
end_date = input_date("Enter the end date (MM/DD/YYYY): ")

# Get emails with the subject 'C3PO'
subject = 'C3PO'
folder = outlook.GetDefaultFolder(6)  # Inbox folder

messages = folder.Items
messages.Sort("[ReceivedTime]", True)  # Sort messages by received time in descending order
restriction = f"[ReceivedTime] >= '{start_date.strftime('%m/%d/%Y %I:%M %p')}' AND [ReceivedTime] <= '{end_date.strftime('%m/%d/%Y %I:%M %p')}'"
restricted_messages = messages.Restrict(restriction)

emails = [msg for msg in restricted_messages if subject in msg.Subject]

if len(emails) == 0:
    print("No email found with the subject '{}'".format(subject))
else:
    email_selection_dialog = Toplevel(root)
    email_selection_dialog.title("Select Emails")
    email_selection_dialog.geometry("400x300")
    email_selection_dialog.lift()
    email_selection_dialog.grab_set()

    email_listbox = Listbox(email_selection_dialog, width=50, selectmode=EXTENDED)
    email_listbox.pack(pady=10)

    for index, msg in enumerate(emails, start=1):
        received_time = msg.ReceivedTime
        sender_name = msg.SenderName
        email_subject = msg.Subject
        email_listbox.insert(END, f"{index}. {sender_name} - {email_subject} - {received_time}")

    select_button = Button(email_selection_dialog, text="Select", command=on_select_emails)
    select_button.pack(pady=5)

    email_selection_dialog.mainloop()
    email_selection_dialog.destroy()

    for selected_email_index in selected_email_indices:
        message = emails[selected_email_index]
        process_email(message, template_file)

    # Extract data from email
    attachments = message.Attachments
    excel_file = ''
    for attachment in attachments:
        if attachment.FileName.endswith('.xlsx'):
            attachment.SaveAsFile(os.path.join(os.getcwd(), attachment.FileName))
            excel_file = attachment.FileName
            print("Attachment saved to '{}'".format(excel_file))
            break

    if not excel_file:
        print("No Excel file found in attachments.")
    else:
        # Load Excel data
        wb = openpyxl.load_workbook(excel_file)
        ws = wb.active

        # Create dictionary with labels and their corresponding cell locations
        labels = {
            'Cherwell': (2, 8),
            'CM Ticket': (3, 8),
            'Anticipated GL Date': (4, 8),
            'Short Summary': (5, 8),
            'Reason': (8, 8),
            'Value': (9, 8),
            'Impact': (10, 8),
            'Additional Information': (11, 8)
        }

        # Create dictionary to store values for each label
        values = {}

        # Extract values from Excel file
        for label, cell in labels.items():
            value = ws.cell(row=cell[0], column=cell[1]).value
            values[label] = value

            # Find the paragraph with the label in the Word document and add the value
            for paragraph in doc.paragraphs:
                if label in paragraph.text:
                    # Add the value from the Excel file to the end of the paragraph
                    paragraph.add_run(" " + str(value))
                    print(f"Updated paragraph text: '{paragraph.text}'")

        # Show message box
        pymsgbox.alert("Please choose a location to save the output file in the upcoming dialog.", "Save Output File")

        # Choose output file
        output_file = filedialog.asksaveasfilename(title="Save Output File As", defaultextension=".docx",
                                                   filetypes=[("Word files", "*.docx")])
        if not output_file:
            print("No output file selected.")
            exit()

        # Save the output file
        doc.save(output_file)

        # Open the output file in Microsoft Word
        # os.startfile(output_file)

        print("Data has been extracted and saved to the Word template.")
        # Delete temp Excel file
        os.remove(excel_file)
