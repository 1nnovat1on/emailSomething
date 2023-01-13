import os
import win32com.client as win32

# Define the folder to check
folder_to_check = 'C:\\path\\to\\folder'

# Check for files in the folder
if len(os.listdir(folder_to_check)) > 0:
    # Create an instance of Outlook
    outlook = win32.Dispatch('outlook.application')

    # Create a new email message
    mail = outlook.CreateItem(0)

    # Set the recipient, subject, and body of the email
    mail.To = 'recipient@example.com'
    mail.Subject = 'Files found in folder'
    mail.Body = 'The following files were found in the folder: '

    # Attach all files in the folder
    for file in os.listdir(folder_to_check):
        file_path = os.path.join(folder_to_check, file)
        mail.Attachments.Add(file_path)

    # Send the email
    mail.Send()

else:
    print("No files found in the folder")
