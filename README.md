# Duplicate Report Formatter

## Overview

The Duplicate Report Formatter is a custom application designed to streamline the process of creating duplicate agreement reports for LinkSquares customers. This application eliminates the manual steps involved in formatting spreadsheets, making the reports easier to understand and interpret.

The formatting and logic used in this tool are specifically tailored for LinkSquares customers, ensuring compatibility with their unique data and reporting requirements.

## Why This Was Created

In my role at LinkSquares, customers often request duplicate reports for their repository of agreements. Manually formatting these reports involves multiple tedious steps, which can be time-consuming and prone to error. To address this, I developed this application to:

- Automate the formatting process.
- Ensure consistency and accuracy in the reports.
- Save time for both internal teams and customers.

## Key Features

- **Column Management:** Automatically removes unnecessary columns to declutter the spreadsheet.
- **Duplicate Highlighting:** Highlights duplicate agreements based on unique identifiers, such as content checksums.
- **Custom Formatting:** Applies specific formatting rules, such as conditional formatting and hiding irrelevant columns.
- **Final Touches:** Adds customer-friendly headers and annotations to guide interpretation of the report.

## How It Works

1. The user uploads a CSV or Excel file containing agreement data.
2. The application processes the file by:
   - Removing unnecessary columns.
   - Highlighting duplicate entries in the `Content Checksum` column.
   - Applying filters for easier navigation.
   - Adding a "Delete? (X)" column for marking duplicates to delete.
3. The formatted file is made available for download.

## Limitations

- The application is not generic and is tailored to the specific formatting requirements of LinkSquares customers.
- It is optimized for agreement data and may not function as expected with other types of data.

## Acknowledgments

This project was created to reduce manual repetitive tasks and cut down ticket resolution timing.

If you have any feedback or suggestions for improving the tool, feel free to reach out!
