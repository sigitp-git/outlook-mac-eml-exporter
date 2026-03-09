# Outlook Mac EML Exporter

AppleScript-based tool to export emails from Microsoft Outlook on macOS as `.eml` files. Works with corporate Microsoft 365 accounts where IMAP, Azure portal, and app passwords are disabled.

## Scripts

| Script | Description |
|--------|-------------|
| `message_count_log.scpt` | Counts total messages and logs to `~/Desktop/message_count.log` |
| `export_all_eml_fixed.scpt` | Basic export of all messages to `~/Desktop/Outlook_EML_Export/` |
| `export_all_eml_enhanced.scpt` | Robust export with 3 fallback methods and logging |
| `export_eml_by_month.scpt` | Date-filtered export organized into `YYYY/MM/` folders |

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

```
~/Desktop/Outlook_EML_Export/
  2025/
    01/
      Meeting_notes_12345.eml
      Project_update_12346.eml
    02/
      Weekly_report_12400.eml
```

A log file is written to `~/Desktop/export_log.txt` with progress and any errors.

## Export Fallback Methods

The enhanced and by-month scripts try 3 methods per message:

1. `save as "eml"` — native Outlook EML export
2. `save as "msg"` + rename to `.eml`
3. Construct a text-based `.eml` from message properties (subject, sender, date, body)

## Requirements

- macOS
- Microsoft Outlook for Mac
- Outlook must be open and running when executing the scripts
