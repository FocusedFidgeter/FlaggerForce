# Paul's Python & VBA Projects

- All the VBA subroutines will be related to `Duke`, because it was the most manual process when I started.
    - There is data cleaning that was automated, files are created, reports are emailed within the company, and Invoices are submitted via email.
- All of the Python files will refer to `Duke`, or `Verizon`, but that is just the nomenclature used inhouse.
    - The more approachable terms would be `generic`, or `exception`.

## `VBA` 📊

### `Utilities.bas` 🔧

*Helper functions for the other VBA subroutines.*

| Function             | Description                                                                                |
|----------------------|--------------------------------------------------------------------------------------------|
| ImportDataToWorksheet| Imports CSV or Excel files. (while also ignoring the report that is under the main table)  |
| CombineTimesheets    | Combines timesheets into one file.                                                         |
| SplitINVs            | Splits our concatenated Invoice Reports into individual files.                            |
| RefreshPivotTables   | Refreshes pivot tables on given a worksheet.                                               |
| RefreshPowerQueries  | Refreshes power queries on given a worksheet.                                              |
| IsWorkBookOpen       | Checks if a Workbook is open.                                                             |
| MoveWithLog          | Moves a file and logs if the expected file does not exist.                                 |
| CopyWithLog          | Copies a file and logs if the expected file does not exist.                                |
| MarkMissingFile      | Marks a cell in an excel sheet so that the user can be notified that the file is missing. |
| CreateWeeklyFolder   | Creates a subfolder for each week in the given yearly folder.                              |

#### Problem this corrected:
- Reduced code redundancy. (you'll see old query refreshes using the standard notation as I have not found sought them all out yet)

### `MondayMorning.bas`, `PdesClosed.bas`, & `DukeSubmissions.bas` ⚙️

*A few noteable functions.*

- `ClearOldData`: Clears old data from the worksheet allowing the Excel Workbook to be reused without duplicating the Macro code every time.
- `SaveWOsToDataModel`: Saves `Work_Order` numbers to the data model so we can store the cleaned data for later.
- `SaveINVsToDataModel`: Saves Invoice numbers to the data model so we can pair them to the cleaned `Work_Order`s.
- `AddressCleanup`: Cleans up the `Address` data deterministically. These issues appear every other week or so, so we just handle them before the user ever sees the "dirty" data they just imported.
- `ProcessDukeLunches`: Cross reference order details with employee hours to determine if adjustments need to be made.
    - Lunches are default for all flaggers unless otherwise specified, and adjustments must be made to reconcile the conflicting data.
- `ProcessAndSortFiles`: Processes and sorts files based on the Invoice number.
    - Enables the worker to find these files again if information is missing or incorrect.

#### Problem this corrected:

- The FinanceDepartment was spending **dozens** of Labor Hours each week performing mundane tasks that no one had thought to make less time-consuming.
    - Hand typing the correct info into PDFs or Excel files.
    - Manually creating and editing PDFs even after they were corrected in an earlier process inside Excel.

## `Python` 🐍

### `Split Invoice Reports.pyw` 📄

*Task:* Split invoice reports into multiple files based on invoice number.

- Uses parsed PDF data to determine the invoice number, and invoice page length.
- Uses the invoice number to determine the destination file.

#### Problem this corrected:

- The department was spending **tens** of Labor Hours each week splitting these by hand across multiple clients.

### `Duke- Combine Timesheets.pyw` 📊

*Task:* Combine timesheets into one file.

- Uses an Excel table to calculate the filenames, and combines them into one file.
- Build a GUI which allows the user to use the program without prior knowledge.

#### Problem this corrected:

- The Department was spending tens of Labor Hours each week doing this by hand.
    - We started with smaller and more acheiveable workloads, but became time consuming as the client started working with us more.
- The GUI allowed others to use this script without my involvement.

### `Verizon- Combine timesheets to INVs.pyw` 📊

*Task:* Combine timesheets and Invoices into one file.

- After running the `split_invoice_reports.py` script, the timesheets were combined into one file and appended to the invoices.
- The GUI allows the user to use this script without prior knowledge.

#### Problem this corrected:

- Again, the Department was spending **dozens** of Labor Hours each week doing this by hand.

### `Duke- Move_or_Delete INVs & CTRs.pyw` 📁

*Task:* Move invoices and CTRs from one folder to another.

- Delete all the files within subfolders to remake the files (correctly, this time).
- Move the invoices and CTRs from their creation folder to the target subfolder (named after the INV).

#### Problem this corrected:

- There were just too many files within one folder, and no organization.
- So, we moved the files to subfolders instead. This enabled us to find and fix the Invoices if a client were to dispute one.
    - Maybe we captured the wrong `Work_Order` number, or the wrong date.
    - Or maybe the old `Rate` was used after a contract was updated.
- This was *mainly written as a helper script* when testing the VBA that sorted the files in the first place.
- I later used this to experiment and learn `TKinter`.