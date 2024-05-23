# Email Automation Script for Google Sheets

This Google Apps Script automates the process of sending emails using data from a Google Sheets document. The script allows users to manage email sends, clear sent email timestamps, and view the status of the script. This user-friendly solution is ideal for batch email sending and provides a streamlined workflow for email campaigns.

## Features

- **Custom Menu**: Adds a "Mailing" menu to the Google Sheets interface for easy access.

- **Send Emails**: Batch sends emails based on the data in the active sheet.

- **Clear Timestamps**: Clears delivery timestamps in the sheet to allow for a new batch of emails.

- **Script Status Sidebar**: Displays a sidebar indicating if the script is currently running.

- **Automated Triggers**: Sets up an hourly trigger to continue sending emails if the batch size exceeds 100 emails.

- **Customizable Email Templates**: Fetches email drafts from Gmail to use as templates, including the handling of inline images and attachments.

- **Error Handling**: Attempts to send an email for each row and logs any errors.

## Installation

1. **Open Google Sheets**: Open the Google Sheets document where you want to run the script.

2. **Script Editor**: Go to `Extensions > Apps Script`.

3. **Copy Code**: Copy the provided script code into the Apps Script editor.

4. **Save**: Save the project with a meaningful name.

5. **Run `onOpen` Function**: Run the `onOpen` function to add the custom menu to your Google Sheets.

## How to Use

### Sending Emails

1. **Prepare Your Sheet**:

   - Ensure that your sheet has the necessary columns, including one for `Email` and another for `Delivery Timestamp`.

2. **Navigate to Mailing Menu**:

   - Click on `Mailing > Send Emails`.

3. **Enter Subject Line**:

   - Enter or copy/paste the subject line of the Gmail draft you want to use as an email template.

4. **Emails Sent**:

   - The script will start sending emails using the data from the sheet. If there are more than 100 emails to be sent, it will set up an hourly trigger to process them in batches.

### Clearing Timestamps

1. **Navigate to Mailing Menu**:

   - Click on `Mailing > Clear Timestamps`.

2. **Confirm Action**:

   - A prompt will appear asking if you want to send a new batch of emails. Confirming this will clear the `Delivery Timestamp` column, preparing the sheet for a new email send.

### View Script Status

- If the `sendEmails` trigger is currently active, a sidebar will be displayed indicating that the script is running.

## Script Functions

- **`onOpen`**: Adds the "Mailing" menu to the Google Sheets UI.

- **`showAlertDialog`**: Displays a prompt to clear timestamps for a new email batch.

- **`clearTimestamps`**: Clears the `Delivery Timestamp` column except for the header row.

- **`displaySidebar`**: Shows a sidebar indicating script status.

- **`isTriggerPresent`**: Checks if the `sendEmails` trigger is present.

- **`evaluateTrigger`**: Evaluates if there are any remaining rows to send emails. Deletes triggers if all emails are sent.

- **`deleteEmailTriggers`**: Deletes email triggers to stop scheduled email sending.

- **`createEmailTrigger`**: Creates an hourly trigger for `sendEmails` if none exists.

- **`promptForSubjectLine`**: Prompts the user to input the subject line of their email template.

- **`sendEmails`**: Main function to send batch emails from sheet data.

- **`getGmailTemplateFromDrafts_`**: Retrieves the Gmail draft template based on the provided subject line.

- **`fillInTemplateFromObject_`**: Replaces placeholders in the email template with data from the sheet.

- **`escapeData_`**: Escapes special characters in email data to prevent JSON issues.

## Notes

- Ensure your Gmail account has an email draft saved with the subject line you intend to use for this script.

- Do not forget to modify the sender email, name, and `replyTo` settings as per your requirement in the `GmailApp.sendEmail` function.

## License

This script is provided "as-is" with no warranties or guarantees. Use at your own risk.

For any issues or contributions, please feel free to create a pull request or issue on GitHub. Enjoy your automated emailing!

**Happy Emailing!** 🎉
