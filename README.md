# Gmail Email Extractor to Google Sheets

This is a Google Apps Script add-on for Gmail that allows you to extract all emails from a specified Gmail label (folder) into a new Google Sheet.

The spreadsheet will contain the following columns:
*   `datum` (date)
*   `afzender` (sender)
*   `ontvanger` (recipient)
*   `titel` (title)
*   `tekstinhoud` (body content)
*   `attachments` (list of attachment filenames)

## Files

*   `Code.gs`: Contains the main logic for the add-on, including the UI and the extraction functionality.
*   `appsscript.json`: The manifest file that configures the add-on, its permissions, and its entry points.

## Installation

1.  Go to [script.google.com](https://script.google.com) and create a new project.
2.  You will see a `Code.gs` file by default. Copy the contents of the `Code.gs` file from this repository and paste it into the editor, replacing the default content.
3.  To create the manifest file, go to **Project Settings** (the gear icon ⚙️ on the left) and check the box for **Show "appsscript.json" manifest file in editor**.
4.  A new `appsscript.json` file will appear in the file list. Click on it.
5.  Copy the contents of the `appsscript.json` file from this repository and paste it into the editor.
6.  Save the project by clicking the save icon (💾).

## How to Use

1.  After saving the project in the Apps Script editor, refresh your Gmail.
2.  The add-on icon (an attachment icon) should appear in the right-hand sidebar.
3.  Click on the icon to open the add-on. You may be asked to authorize the script the first time you use it. Please grant the necessary permissions.
4.  An input field will appear. Enter the exact name of the Gmail label you want to extract emails from.
5.  Click the **Extract Emails** button.
6.  The script will process the emails and create a new Google Sheet in the root of your Google Drive.
7.  A notification will appear at the bottom of the add-on with a link to the newly created spreadsheet. If there's an error (e.g., the label doesn't exist), an error message will be shown instead.