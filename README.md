# Email-Attachment-Automation

A Node.js automation service that scans Outlook folders and subfolders, mirrors them into OneDrive, extracts attachments, creates anonymous view-only links, and upserts the corresponding records in a Notion database.

## Easiest workflow

If the full Microsoft Graph and Notion API setup feels too complicated, use the simple local intake instead.

This does not require Azure app registration, Microsoft Graph permissions, a Notion token, or background polling. It creates a Notion-ready CSV from files you already saved into a OneDrive-synced folder.

1. In OneDrive, create the folders you want lenders or partners to see.
2. Save or drag Outlook attachments into those folders.
3. Run:

```bash
npm run simple -- "/path/to/your/OneDrive/Magnolia Homes at Chimes" notion-import.csv
```

4. Import `notion-import.csv` into Notion or merge it into the Magnolia database.
5. In OneDrive, create view-only sharing links for the files.
6. Paste those links into the blank `OneDrive File Link` column in Notion.

The CSV includes document name, folder path, category, file size, SHA-256 hash, and a processing key so you can track what was imported without setting up the full automation.

## What it does

- Recursively reads configured Outlook folders and their subfolders
- Recreates the same folder structure in OneDrive
- Processes both existing mail and new mail by polling on a schedule
- Downloads every file attachment without opening messages manually
- Preserves sender, received date, subject, folder path, message ID, filename, size, and SHA-256 hash
- Prevents duplicate processing with a composite processing key based on message ID, attachment ID, filename, and file hash
- Avoids overwriting stored files by versioning filename collisions
- Creates 90-day anonymous view-only OneDrive links and verifies they do not redirect to Microsoft sign-in
- Updates only the target Notion columns and creates a new page when no match exists
- Writes an audit log covering processed, duplicate, restricted, password-protected, unmatched, and failed items

## Environment variables

| Variable | Required | Description |
| --- | --- | --- |
| `MICROSOFT_TENANT_ID` | yes | Azure AD tenant ID for Microsoft Graph |
| `MICROSOFT_CLIENT_ID` | yes | App registration client ID |
| `MICROSOFT_CLIENT_SECRET` | yes | App registration client secret |
| `MAILBOX_USER_ID` | yes | User ID or UPN for the Outlook mailbox and OneDrive |
| `OUTLOOK_ROOT_FOLDER_IDS` | yes | Comma-separated Outlook folder IDs to monitor recursively |
| `ONEDRIVE_ROOT_PATH` | no | Base OneDrive path, defaults to `/Email Attachments` |
| `ONEDRIVE_DRIVE_ID` | no | Drive ID when you do not want to use the mailbox user's default drive |
| `PUBLIC_LINK_EXPIRY_DAYS` | no | Link expiry window, defaults to `90` |
| `NOTION_TOKEN` | yes | Notion integration token |
| `NOTION_DATABASE_ID` | yes | Magnolia Homes at Chimes database ID |
| `NOTION_TITLE_PROPERTY` | no | Title property name, defaults to the first title column in the database |
| `POLL_INTERVAL_MS` | no | Polling interval for continuous runs, defaults to `300000` |
| `RUN_ONCE` | no | Set to `true` to process a single pass and exit |
| `STATE_FILE` | no | JSON state file path, defaults to `./data/processing-state.json` |
| `AUDIT_LOG_FILE` | no | JSONL audit log path, defaults to `./data/audit-log.jsonl` |

## Setup

1. Create a Microsoft Graph application with application permissions for:
   - `Mail.Read`
   - `Files.ReadWrite.All`
   - `Sites.ReadWrite.All` if the target drive is on SharePoint/OneDrive for Business
2. Grant admin consent.
3. Create a Notion internal integration with access to the Magnolia Homes at Chimes database.
4. Export the environment variables listed above.
5. Start the automation:

```bash
npm start
```

For a single execution:

```bash
RUN_ONCE=true npm start
```

## Testing

```bash
npm test
```
