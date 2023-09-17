# Email-Sender-Script
 An automated Python email sender script using SMTP. Send emails, track delivery, and receive progress notifications. Ideal for email campaigns and automation.
# Email Sender Python Script

This Python script is designed to automate the process of sending emails to a list of recipients using the SMTP protocol. It can be useful for sending bulk emails with attachments and HTML content.

## Prerequisites

Before using this script, make sure you have the following prerequisites:

- Python 3.x installed on your system.
- An internet connection for sending emails.
- SMTP server details (configured in an Excel file as described below).
- A list of email addresses (also in an Excel file).
- Email content in HTML format (stored in an `email.html` file).
- Image and PDF attachments (stored in respective directories).

## Configuration

1. SMTP Configuration: 
   - SMTP server details (username, password, server address) should be stored in an `smtp.xlsx` Excel file.
   - Make sure to fill in the "From" name and subject in the Excel file.

2. Email List: 
   - The list of email recipients should be stored in an `email_list.xlsx` Excel file.

3. Email Body:
   - The HTML content of the email should be stored in an `email.html` file.
   - Images and PDF files to be attached should be placed in their respective directories (`images` and `pdf`).

## Usage

1. Run the script: 
   - Execute the script to start sending emails.
   - The script will send emails to the recipients, updating the delivery status in the Excel sheet.

2. Progress Notification:
   - You will receive a progress notification email every 100 emails sent.

3. Blacklist:
   - Blacklisted email addresses specified in the script will be skipped.

## Notes

- The script includes error handling to manage issues that may occur during the email sending process.

## NEED FRONTEND DESIGNING
Need Frontend design for this project

Feel free to use, modify, and distribute this script as needed.

## Author

- Aj (GitHub: [ajiq360](https://github.com/ajiq360))

If you have any questions or encounter issues, please feel free to open an issue on this repository.

Happy emailing!
