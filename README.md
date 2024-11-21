# Remove Duplicate Emails in Outlook

This VBA macro removes duplicate emails from the currently selected Outlook folder. It checks for duplicates based on the email's subject, sender, received time (excluding seconds), and body content.

## Features
- Identifies and removes duplicate emails from the active folder.
- Provides a progress bar to indicate processing status.
- Displays the number of emails removed upon completion.

## Prerequisites
- Microsoft Outlook
- Basic understanding of VBA

## Installation
1. Open Outlook.
2. Press `Alt + F11` to open the VBA editor.
3. Import the `RemoveDuplicateEmails.bas` and `ProgressForm.frm` files:
   - Go to `File > Import File`.
   - Select the `.bas` file and `.frm` file.
4. Close the VBA editor and return to Outlook.

**NOTE: Make sure that all 3 files 'RemoveDuplicateEmails.bas`, 'ProgressForm.frx' and 'ProgressForm.frm' exists on the same directory while improting. however you dont have to import    'ProgressForm.frx' at all, it will be imported automatically.**

## Usage
1. Navigate to the folder where you want to remove duplicate emails.
2. Press `Alt + F8` to open the macro list.
3. Select `RemoveDuplicateEmailsSafely` and click `Run`.

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
