Prompt

I have a spread sheet that contains the anual income for charity by members. The spreadsheet came from the bank statement but has been edited to contain just the income with two new colums [Purpose and Description]. I want to use the transaction description column in the income spreadsheet to find the payee name and use this payee name to find the full payee details in another  spread sheet containing the members database. Can you implment a plan to create a website that will allow people to load the input files and generate an output that can used to identify how much a member has donated in the year.

Response

## Implementation Plan: Charity Member Donation Reconciliation Web App

### Overview

The goal is to create a web application that allows users to upload two spreadsheets:
- **Income Spreadsheet:** Contains annual income transactions, with columns such as Transaction Description, Purpose, and Description.
- **Members Database Spreadsheet:** Contains member details (including payee names and other identifying information).

The app will match payee names from the income spreadsheet (using the transaction description) to the members database, aggregate the total donations per member for the year, and generate a downloadable report.

---

### Requirements

- **File Upload**
  - Allow users to upload the income spreadsheet (CSV or Excel).
  - Allow users to upload the members database spreadsheet (CSV or Excel).
- **Data Processing**
  - Parse both spreadsheets in the browser.
  - Extract payee names from the transaction description column in the income spreadsheet.
  - Match payee names to member records in the members database.
  - Aggregate total donations per member for the year.
- **Output**
  - Display a summary table of total donations per member.
  - Allow users to download the results as a CSV or Excel file.
- **User Interface**
  - Simple, clear, and responsive UI.
  - Instructions for users on required file formats and columns.
- **Privacy**
  - All processing is done client-side; no data is sent to a server.

---

### Implementation Steps

1. **Project Setup**
   - Create a new web project directory.
   - Add `index.html`, `styles.css`, and `app.js` files.
   - Add a library for spreadsheet parsing (e.g., [SheetJS](https://sheetjs.com/)).

2. **User Interface (HTML & CSS)**
   - Create a form with two file input fields (one for each spreadsheet).
   - Add a button to trigger processing.
   - Add a section to display results (summary table).
   - Add a button to download the output.
   - Provide clear instructions and sample file format information.

3. **Spreadsheet Parsing (JavaScript)**
   - Use SheetJS or similar to parse uploaded files (supporting both CSV and Excel).
   - Extract relevant columns from each file.
   - For the income spreadsheet, extract the payee name from the transaction description.
   - For the members database, index members by name (and possibly other identifiers for fuzzy matching).

4. **Matching Logic**
   - Implement logic to match payee names from the income spreadsheet to member records.
   - Handle possible variations in name formatting (consider case-insensitive and partial matches, or allow user to review/resolve ambiguous matches).
   - Aggregate donation amounts per matched member.

5. **Results Display & Export**
   - Display a table showing each member and their total donations for the year.
   - Allow the user to download the results as a CSV or Excel file.

6. **Testing**
   - Test with real and sample data for various edge cases (missing names, ambiguous matches, unmatched transactions).
   - Test on different browsers and devices.

7. **Accessibility & Usability**
   - Ensure the UI is accessible (labels, keyboard navigation).
   - Provide error messages and feedback for invalid files or unmatched records.

---

### Testing

- **Unit Testing (Manual/Automated)**
  - Test file parsing with various spreadsheet formats.
  - Test name extraction and matching logic.
  - Test aggregation and output generation.
- **User Testing**
  - Verify the workflow with real users and real data.
  - Ensure the output matches expectations and is easy to use.

---

### Optional Enhancements

- Allow users to manually resolve ambiguous or unmatched payee names.
- Support for additional file formats (Google Sheets, etc.).
- Visualization of donation data (charts, graphs).
- Save and reload previous sessions (in browser storage).

---

This plan will result in a user-friendly, privacy-preserving web tool for reconciling and reporting member donations from bank-derived income spreadsheets.
