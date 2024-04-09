# Paul's Python & VBA Projects

- All of these files will refer to `Duke`, or `Verizon`, but that is just the nomenclature used inhouse.
    - The more approachable terms would be `generic`, or `exception`.
- All the VBA subroutines will be related to `Duke`, because it was the most manual process when I started.
    - There is data cleaning that was automated, files are created, reports are emailed within the company, and Invoices are submitted via email.

## Python

### Split Invoice Reports 

#### Task: Split invoice reports into multiple files based on invoice number.

- Uses parsed PDF data to determine the invoice number, and invoice page length.
- Uses the invoice number to determine the destination file.

#### Problem this corrected:

- The department was spending **tens** of Labor Hours each week splitting these by hand across multiple clients.

### Duke- Move_or_Delete INVs & CTRs

#### Task: Move invoices and CTRs from one folder to another.

- Delete all the files within subfolders in order to remake the files (correctly, this time).
- Move the invoices and CTRs from their creation folder to the target subfolder (named after the INV).

#### Problem this corrected:

- There were just too many files within one folder, and no orginization.
- So, we moved the files to subfolders instead. This enabled us to find and fix the Invoices if a client were to dispute one.
    - Maybe we captured the wrong `Work_Order` number, or the wrong date.
    - Or maybe the old `Rate` was used after a contract was updated.
- This was *mainly written as a helper script* when testing the VBA that sorted the files in the first place.
- I later used this to experiment and learn `TKinter`.

### Duke- Combine Timesheets

#### Task: Combine timesheets into one file.

- Uses an Excel table to calculate the filenames, and combines them into one file.
- Build a GUI which allows the user to use the program without prior knowledge.

#### Problem this corrected:

- The Department was spending tens of Labor Hours each week doing this by hand with only a single client.
    - We started with smaller and more acheiveable workloads, but became time consuming as the client started working with us more.
- The GUI allowed others to use this script without my involment.

### Verizon- Combine timesheets to INVs

#### Task: Combine timesheets and Invoices into one file.

- After running the `split_invoice_reports.py` script, the timesheets were combined into one file and appended to the invoices.
- The GUI allows the user to use this script without prior knowledge.

#### Problem this corrected:

- Again, the Department was spending *dozens* of Labor Hours each week doing this by hand with only a single client.

## VBA

### Utilities.bas

#### Task: Helper functions for the other VBA subroutines.

- 