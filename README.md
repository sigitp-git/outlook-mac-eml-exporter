# Outlook Mac EML Exporter

AppleScript-based tool to export emails from Microsoft Outlook on macOS as `.eml` files. Works with Microsoft 365 accounts where IMAP, Azure portal, and app passwords are disabled. The scripts load locally cached/stored emails on Mac. If the emails are not cached/stored locally on Mac, the scripts doesn't work.

## Scripts

| Script | Description |
|--------|-------------|
| `message_count_log.scpt` | Counts total messages and logs to `~/Desktop/message_count.log` |
| `export_all_eml_fixed.scpt` | Early version — single export method, no logging, uses loop index for filename uniqueness |
| `export_all_eml_enhanced.scpt` | Improved version — adds 3 fallback export methods, log file (`export_log.txt`), per-message error handling, failure counting, property access guarding, message `id` in filenames for uniqueness, and more thorough filename cleaning (`"` and `'` removal) |
| `export_eml_by_month.scpt` | Latest version — builds on enhanced with command-line year/month filtering, `YYYY/MM/` folder organization, scans all folders (Inbox, Sent Items, etc.) across all Exchange accounts, and prefixes folder name in filenames |

## Usage

### Count messages first
```bash
osascript ~/message_count_log.scpt
cat ~/Desktop/message_count.log
```

### Export by month/year (recommended)
```bash
# Export single month (e.g., January 2025)
osascript ~/export_eml_by_month.scpt 2025 1

# Export a range (e.g., all of 2025)
osascript ~/export_eml_by_month.scpt 2025 1 2025 12

# Export everything (no date filter)
osascript ~/export_eml_by_month.scpt
```

### Export all messages (legacy)
```bash
osascript ~/export_all_eml_fixed.scpt
osascript ~/export_all_eml_enhanced.scpt
```

## Output Structure

`export_eml_by_month.scpt` scans all mail folders across all Exchange accounts (Inbox, Sent Items, Drafts, etc.) and organizes emails by date. The folder name is prefixed in each filename so you can tell where the email came from:

```
~/Desktop/Outlook_EML_Export/
  2025/
    01/
      Inbox_Meeting_notes_12345.eml
      Sent Items_Reply_to_John_12346.eml
    02/
      Inbox_Weekly_report_12400.eml
      Sent Items_Project_update_12401.eml
```

Real-time progress is logged to the terminal and to `~/Desktop/export_log.txt`:

```
Scanning folder: account/Inbox (5000 messages)
  [1] Inbox_Meeting_notes_12345.eml
  [2] Inbox_Project_update_12346.eml
Scanning folder: account/Sent Items (3000 messages)
  [5001] Sent Items_Reply_to_John_12500.eml
```

## Export Fallback Methods

The enhanced and by-month scripts try 3 methods per message:

1. `save as "eml"` — native Outlook EML export
2. `save as "msg"` + rename to `.eml`
3. Construct a text-based `.eml` from message properties (subject, sender, date, body)

## Requirements

- macOS
- Microsoft Outlook for Mac
- Outlook must be open and running when executing the scripts
