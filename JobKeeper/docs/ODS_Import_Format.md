# ODS Import Format Guide

## Overview
JobKeeper can import job application data from OpenOffice/LibreOffice Calc (.ods) files.

## File Format

### Required Columns
The ODS file should have a **header row** with column names. The following column headers are recognized (case-insensitive):

| Column Name | Alternative Names | Required | Format |
|-------------|------------------|----------|--------|
| COMPANY | Company Name | **Yes** | Text |
| WEBSITE | URL, Site | No | Text (http:// or https://) |
| JOB TITLE | Position | No | Text |
| SUBMITTED | Applied, Date | No | Date (MM/dd/yyyy or similar) |
| RESUME | CV | No | File path |
| COVER LETTER | Cover | No | File path |
| STATUS | - | No | SUBMITTED, REJECTED, INTERVIEW CHANGED, GHOSTED |
| INTERVIEW 1 | Interview First | No | Date (MM/dd/yyyy or similar) |
| INTERVIEW 2 | Interview Second | No | Date (MM/dd/yyyy or similar) |

### Example ODS Structure

```
| COMPANY    | WEBSITE              | JOB TITLE        | SUBMITTED  | STATUS    | INTERVIEW 1 |
|------------|----------------------|------------------|------------|-----------|-------------|
| Acme Corp  | https://acme.com     | Software Dev     | 01/15/2025 | SUBMITTED | 01/20/2025  |
| Tech Inc   | https://techinc.com  | Senior Engineer  | 01/10/2025 | REJECTED  |             |
| DevCo      | https://devco.io     | Full Stack Dev   | 01/05/2025 | GHOSTED   |             |
```

## Supported Date Formats
- MM/dd/yyyy (e.g., 01/15/2025)
- dd/MM/yyyy (e.g., 15/01/2025)
- yyyy-MM-dd (e.g., 2025-01-15)
- M/d/yyyy (e.g., 1/5/2025)

## Status Values
The STATUS column accepts the following values (case-insensitive):
- **SUBMITTED** (default if empty)
- **REJECTED** (also recognizes "REJECT")
- **INTERVIEW CHANGED** (also recognizes "INTERVIEW")
- **GHOSTED** (also recognizes "GHOST")

## Import Process

1. Go to **File → Import from ODS...**
2. Select your .ods file
3. Choose import mode:
   - **Yes** - Add imported records to existing data
   - **No** - Replace all existing data with imported records
   - **Cancel** - Abort the import

## Tips

- The first row **must** be a header row with column names
- Empty rows are automatically skipped
- At minimum, the COMPANY field must be filled for a row to be imported
- Column order doesn't matter as long as headers are present
- Extra columns are ignored
- Missing optional fields will be left empty in JobKeeper

## Creating a Template in LibreOffice Calc

1. Open LibreOffice Calc
2. Create headers in the first row: COMPANY, WEBSITE, JOB TITLE, SUBMITTED, RESUME, COVER LETTER, STATUS, INTERVIEW 1, INTERVIEW 2
3. Fill in your job application data starting from row 2
4. Save as .ods format (File → Save As → ODF Spreadsheet (.ods))
5. Import into JobKeeper

## Troubleshooting

- **"Invalid ODS file"**: Make sure the file is a valid .ods format
- **"No valid applications found"**: Check that rows have at least a COMPANY value
- **Dates not importing**: Verify date format matches one of the supported formats
- **Status not recognized**: Check spelling of status values (SUBMITTED, REJECTED, INTERVIEW CHANGED, GHOSTED)
