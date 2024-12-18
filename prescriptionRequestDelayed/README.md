# Outlook Macros for Edmund Optics (Delayed Version)

A collection of VBA macros for automating email tasks in Outlook, specifically for Edmund Optics Technical Support. This version includes a 10-minute processing delay to coordinate with multiple team members.

## Prescription Request Handler (Delayed)

### Overview
Automates responses to prescription file requests by:
- Processing emails marked as "Edmund Optics Prescription Request"
- Waiting 10 minutes after email receipt before processing
- Extracting customer information and requested part numbers
- Locating and attaching corresponding prescription files
- Generating context-appropriate responses
- Handling multiple file scenarios

### Delay Mechanism
This version includes a 10-minute delay mechanism to coordinate with team members across different time zones:
- Checks the received time of each email
- Only processes emails that are at least 10 minutes old
- Allows other team members to process urgent requests immediately
- Prevents conflicts when multiple team members are monitoring the same inbox

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
- `prescriptionRequestDelayed.bas`: Handles prescription file requests with 10-minute delay
- `InboxEvents.cls`: Event handling for inbox monitoring
- `ThisOutlookSession`: Outlook session configuration

## Installation
1. Clone this repository
2. Import the .bas and .cls files into your Outlook VBA project
3. Configure file paths in the macro settings:
   ```vb
   FilePath = "\\us-fs2\Public\Engineering\Zemax Files\Prescriptions\"
   ```

## Usage
1. Ensure you have access to the Technical Support USA inbox
2. Run the macro - it will automatically process emails that are at least 10 minutes old
3. Monitor the Immediate window in the VBA editor to see processing status and timing information

## Coordination with Other Team Members
This delayed version is designed to work alongside the immediate-processing version:
- East Coast team members can use the standard version for immediate processing
- West Coast team members should use this delayed version
- The 10-minute delay ensures no conflicts when processing the same inbox