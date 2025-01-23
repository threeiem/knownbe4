# KnownBe4 Gmail Phishing Training Scanner

This Google Apps Script automatically scans your Gmail inbox for KnownBe4 phishing training emails. It helps track and monitor your organization's security awareness training emails.

## Features

- Automatically scans Gmail inbox for KnownBe4 training emails
- Uses Gmail's search capabilities to identify phishing training attempts
- Runs on a scheduled trigger for continuous monitoring
- Can be customized to track specific headers or content

## Installation

1. Visit [script.google.com](https://script.google.com)
2. Click "New project"
3. Copy and paste the script code into the editor
4. Save the project by clicking "Untitled project" at the top and giving it a name

## Setup

1. Run the script for the first time:
   - Click the "Run" button
   - Authorize the script to access your Gmail
   - Click "Advanced" and then "Go to [project name]"
   - Click "Allow" to grant the necessary permissions

2. Set up automatic scanning:
   - Click the clock icon in the left sidebar (Triggers)
   - Click "+ Add Trigger" at the bottom
   - Configure your desired schedule (recommended: hourly or daily)
   - Save the trigger

## Configuration

The script looks for specific headers and content that identify KnownBe4 training emails. You can modify the search criteria in the code to match your organization's specific KnownBe4 implementation.

## Security Notes

- This script requires Gmail access permissions
- Only install this script in environments where KnownBe4 training is expected
- Review the code before granting permissions
- Keep your script project private to maintain security

## Contributing

Feel free to submit issues and enhancement requests!

## License

[MIT License](LICENSE)
