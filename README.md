# Outlook Attachment Downloader

This script downloads attachments from the 8th most recent email in your Outlook inbox and saves them to a `file` folder on your desktop.

## Features
- Downloads attachments from the 8th most recent email.
- Saves attachments to a designated folder on your desktop.

## Usage
- Install the required package: `pip install pywin32`.
- Run the script. 

## Requirements
- Python
- `pywin32` library
- Microsoft Outlook installed and configured

## Notes
- The script targets the 8th email (`messages[7]`), but you can modify this as needed.
- The attachments are saved to the desktop by default. You can modify the save location in the script if needed.
