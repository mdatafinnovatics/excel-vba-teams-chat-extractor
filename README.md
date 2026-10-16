# Excel VBA Microsoft Teams Chat Extractor

Excel VBA automation that signs in to a Microsoft 365 account and extracts
Microsoft Teams chat messages directly into Excel with one click.

Works on both **Windows** and **Mac** versions of Excel.

## What it does

- Signs in securely to the user's Microsoft 365 account
- Connects to Microsoft Teams through the Microsoft Graph API
- Extracts chat messages with date, time, sender and message text
- Writes the data into a clean, structured Excel table
- Ready for further reporting or AI processing

## Sample output

| Team | Channel | Date | Time | Sender | Message |
|------|------|--------|----------------|---------|
| Express Mobile Storage | Team Group | 06-Sep-2026 | 10:22:36 | Jacob Shapero | Hello |
| Express Mobile Storage | Team Group | 06-Sep-2026 | 10:16:57 | Jacob Shapero | Are you there? |

## Screenshots

See Attachments

## How it works

1. User clicks the **SIGN IN** button to login and then **IMPORT MESSAGES** button to import the messages from Microsoft Teams
2. Click **Only for Complete Mac SIGNIN** button if someone is using the Mac
3. The tool signs in to Microsoft 365 (OAuth authentication)
4. Chat messages are requested from the Microsoft Graph API
5. The JSON response is parsed in VBA
6. Messages are written into a formatted Excel table

## Platform support

- **Windows:** Excel for Microsoft 365 / Excel 2016 or later
- **Mac:** Excel for Microsoft 365 on macOS

## Tech used

- Excel VBA
- Microsoft Graph API
- OAuth 2.0 sign-in (Microsoft Entra ID)
- REST API calls and JSON parsing in VBA
- Cross-platform support (Windows and Mac)

## Setup

1. Register an app in Microsoft Entra ID (Azure portal) with permission
   to read chats
2. Download the demo workbook and enable macros
3. Open the VBA editor (Alt + F11 on Windows, Option + F11 on Mac)
4. Replace `YOUR_CLIENT_ID` and `YOUR_TENANT_ID` with your own values
5. Click the **SIGN IN** and **IMPORT MESSAGES** buttons

## Related project

Extracted chats can be converted into structured insights using
Generative AI here: [Excel AI Chat Insights](PASTE-AI-PROJECT-ADDRESS-HERE)

## About

Built by **Pradeep Kumar**, founder of Mdata Finnovatics.
Excel VBA, Power BI and AI integration specialist.

*This is a demo version with sample data. No client data or credentials
are included.*
