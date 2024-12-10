# Outlook Prescription Request Automation

This VBA macro automates the process of handling prescription file requests in Outlook. It monitors a shared inbox for prescription requests, locates the requested files, and sends automated responses with the appropriate attachments.

## Features

- Monitors a specified Outlook folder for prescription request emails
- Extracts part numbers from email body using regex
- Searches a designated folder for matching prescription files
- Automatically generates response emails with:
  - Personalized greeting using the requester's first name
  - Attached prescription files
  - Professional message body with usage instructions
  - List of any unavailable parts
  - Corporate signature
- Comprehensive error handling and logging
- Support for Zemax prescription files

## Prerequisites

- Microsoft Outlook
- Access to shared folder: `Q:\Released Documents\Optical Prescriptions\Zemax Files\Black box\`
- Appropriate permissions to send emails from the Technical Support inbox

## Installation

1. Open Outlook
2. Press Alt + F11 to open the VBA editor
3. Import the `prescriptionRequest.bas` module
4. Save and close the VBA editor

## Usage

The macro will:
1. Monitor the "Technical Support USA/Inbox" folder
2. Process emails with subject containing "Edmund Optics Prescription Request"
3. Generate and display response emails for review before sending

## Functions

- `ProcessPrescriptionRequests()`: Main routine that handles the automation flow
- `ExtractPartNumbers()`: Extracts 5-digit part numbers from email body
- `ExtractRecipientEmail()`: Extracts email address from message body
- `ExtractFormField()`: Extracts specific fields from form submissions
- `GetDefaultSignature()`: Retrieves the default Outlook signature

## Error Handling

The macro includes comprehensive error handling and debugging:
- Folder existence verification
- File access checks
- Debug logging throughout the process
- Graceful handling of missing files or permissions

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

For support, please contact the Technical Support team or raise an issue in this repository.

## Acknowledgments

- Edmund Optics Technical Support Team
- Zemax file handling documentation 