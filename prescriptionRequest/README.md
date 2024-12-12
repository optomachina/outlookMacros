# Outlook Macros for Edmund Optics

A collection of VBA macros for automating email tasks in Outlook, specifically for Edmund Optics Technical Support.

## Prescription Request Handler

### Overview
Automates responses to prescription file requests by:
- Processing emails marked as "Edmund Optics Prescription Request"
- Extracting customer information and requested part numbers
- Locating and attaching corresponding prescription files
- Generating context-appropriate responses
- Handling multiple file scenarios

### Response Scenarios
The macro intelligently handles four different scenarios:

1. Single File Request (Found)
   - Attaches the requested file
   - Uses singular language in response

2. Multiple Files Request (All Found)
   - Attaches all requested files
   - Uses plural language in response

3. Single/Multiple Files (None Found)
   - Sends appropriate apology message
   - Indicates files are not available

4. Multiple Files (Some Found)
   - Attaches available files
   - Lists which part numbers were not found
   - Uses appropriate plural language

## Files
- `prescriptionRequest.bas`: Handles prescription file requests
- `draftReplyToSelectedEmail.bas`: Helper functions for email drafting
- `AddinSetup.bas`: Setup and configuration code
- `InstallAddin.bat`: Installation script

## Installation
1. Clone this repository
2. Run `InstallAddin.bat` to set up the Outlook add-in
3. Configure file paths in the macro settings:
   ```vb
   FilePath = "Q:\Released Documents\Optical Prescriptions\Zemax Files\Black box\"
   ```

## Usage
1. Ensure you have access to the Technical Support USA inbox
2. Run the appropriate macro for your task
3. Review generated emails before sending

## Requirements
- Microsoft Outlook
- Access to network file locations
- Appropriate inbox permissions

## Development
Maintained by Blaine Wilson for Edmund Optics Technical Support