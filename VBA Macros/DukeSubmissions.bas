Attribute VB_Name = "DukeSubmissions"
Option Explicit

Sub ImportOETs()
' This VBA code imports data from a file and performs operations on Excel worksheets within the current workbook.
' It also turns off screen updating to improve performance and then turns it back on at the end.
' Finally, it informs the user that the import is complete.
    Dim wsOETs As Worksheet
    Dim wsCTR As Worksheet
    Dim wsINV As Worksheet
    Dim wsEmail As Worksheet
    Dim lastRowImport As Long
    Dim lastRowOETs As Long
    Dim i As Long

    ' Tell the user what will be done
    MsgBox "Select the `Order Entry Transactions` file that you have downloaded from Intaact Sage."

    ' Set the separate WSs
    Set wsOETs = ThisWorkbook.Worksheets("Order Entry Transactions")
    Set wsCTR = ThisWorkbook.Worksheets("CTR Template")
    Set wsEmail = ThisWorkbook.Worksheets("Email Template")

    ' Import Billing Details from PowerBI
    ImportDatatoWorksheet "Order Entry Transactions", "A1", "CSV"

    ' Turn off screen updating
    Application.ScreenUpdating = False

    With wsOETs
        ' Find the last row in the data
        lastRowOETs = .Cells(.Rows.Count, "A").End(xlUp).Row
        ' Copy the Formulas in row 2 and paste them down the rest of the table
        .Range("O2:R2").Copy
        .Range("O3:R" & lastRowOETs).PasteSpecial Paste:=xlPasteFormulas
    End With

    With wsEmail
        ' Copy the Formulas in row 2 and paste them down the rest of the table
        .Range("B2:P2").Copy
        .Range("B3:B" & lastRowOETs).PasteSpecial Paste:=xlPasteFormulas
        .Range("B2:C" & lastRowOETs).Copy
        .Range("B2:B" & lastRowOETs).PasteSpecial Paste:=xlPasteValues
        .Range("F2:F2" & lastRowOETs).Copy
        .Range("F3:F" & lastRowOETs).PasteSpecial Paste:=xlPasteValues
    End With

    ActiveWorkbook.Save
    ' Turn on screen updating
    Application.ScreenUpdating = True
    ' Inform the user that the import is complete
    MsgBox "Invoice Summary imported successfully!"

End Sub

Sub refresh_cities()
    ' Refresh the this workbook's connection to the
    '   master City List inside Duke/Resources/duke_data_model.xlsx
    ThisWorkbook.Connections("Query - Duke_Ops_Centers").Refresh
End Sub

Function IsWorkBookOpen(Name As String) As Boolean
    Dim xWb As Workbook
    On Error Resume Next
    Set xWb = Application.Workbooks.Item(Name)
    IsWorkBookOpen = (Not xWb Is Nothing)
End Function

Sub Slow_n_SafeFormulas()
' This function exists because Paul couldn't figure out the data model
'    needed to do this properly with Power Query.
' This will lock the user out of the notebook and prevent them from
'   interrupting, and possibly corrupting, this week's WorkBook.
' We slowly copy and paste-as-values sections of the CTR worksheet so
'   that the CPU can keep up. First are the details extracted straight
'   from sheet("Import").
' Then XLOOKUPing the INV# from the Intaact Details. This is why the
'   `Wait_a_minute` function exists. We have a janky method to trim an
'   array of all (150+) WO#s for each of the 2000+ line items.
' Finally we use a standard-ish method of indexing the individual
'   INV line-items. We do this by sequencually calculating each Index#
'   starting at row 2 and waiting until it calculates through row `lastRow`.
' BigO=N^2 total (I think...)

    On Error GoTo ErrorHandler

    ' Disable user interaction and screen updating
    Application.Interactive = False
    Application.ScreenUpdating = False

    ' Define variables
    Dim wsDraftImport As Worksheet
    Dim wsCTR As Worksheet
    Dim lastRow As Long

    ' Set the worksheet references
    Set wsDraftImport = ThisWorkbook.Sheets("Draft_Import")
    Set wsCTR = ThisWorkbook.Sheets("CTR Template")

    ' Find the last row with data in a specific column on "Draft_Import" sheet
    ' Change "A" to the column you want to check for the last row of data
    lastRow = wsDraftImport.Cells(wsDraftImport.Rows.Count, "A").End(xlUp).Row

    ' Make the "CTR Template" sheet active
    wsCTR.Activate

    ' Enter formulas into the second row
    With wsCTR
        .Range("A2").Formula = "=XLOOKUP($F2,TRIM('Order Entry Transactions'!$B:$B),'Order Entry Transactions'!$C:$C,0)"
        .Range("B2").Formula = "=IF(A2<>A1,1,B1+1)"
        .Range("C2").Formula = "=Import!G2"
        .Range("D2").Formula = "=Import!J2"
        .Range("F2").Formula = "=TRIM(Import!E2)"
        .Range("J2").Formula = "=Import!S2"
        .Range("K2").Formula = "=Import!I9"
        .Range("L2").Formula = "=K2"
        .Range("M2").Formula = _
            "=Import!F2&"", ""&Import!H2&"", ""&Import!R2"
        .Range("N2").Formula = "=Import!K2"
        .Range("O2").Formula = "=Import!P2"
        .Range("C2:O2").Copy
        .Range("C3:C" & lastRow).PasteSpecial xlPasteFormulas
        Application.CutCopyMode = False

        ' Convert formulas to values for columns C to O
        .Columns("C:O").Copy
        .Columns("C:O").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, _
                                        SkipBlanks:=False, Transpose:=False

        ' Copy the values in A2:B2 down to the last row
        .Range("A2:B2").Copy
        .Range("A3:A" & lastRow).PasteSpecial xlPasteFormulas
        Application.CutCopyMode = False

        ' Convert formulas to values for columns A to B
        .Columns("A:B").Copy
        .Columns("A:B").PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        Application.CutCopyMode = False
    End With

    ' Disable user interaction and screen updating
    Application.Interactive = True
    Application.ScreenUpdating = True
    Exit Sub

    ' Select the first cell to reset the selection
    wsCTR.Range("A1").Select

ErrorHandler:
    ' Error handling code - re-enable interaction and screen updating
    Application.Interactive = True
    Application.ScreenUpdating = True
    MsgBox "An error occurred: " & Err.Description
End Sub
Function Wait_a_minute()
    Dim waitTime As Double
    waitTime = Now + TimeValue("00:01:00") ' 1 minute
    Application.Wait waitTime
End Function
Sub CreateCtrFiles()
    Dim wsImport As Worksheet
    Dim wsCTR As Worksheet
    Dim rngToCopy As Range
    Dim lastRow As Long
    Dim wsDest As Worksheet
    Dim sInvoice As String
    Dim dte As String
    Dim sCTRFile As String
    Dim sDestFileLocation As String
    Dim iFirstInvoiceRow As Integer
    Dim i As Integer
    Dim sCTRFileName As String

    Set wsImport = ThisWorkbook.Sheets("Import")
    Set wsCTR = ThisWorkbook.Sheets("CTR Template")

    lastRow = wsImport.Cells(wsImport.Rows.Count, "A").End(xlUp).Row

    ' Set the range to copy from A1 to O and the last row in the "CTR Template" sheet
    Set rngToCopy = wsCTR.Range("A1:O" & lastRow)
    
    ' Stop updating the screen
    Application.ScreenUpdating = False
    
    ' Copy the range
    rngToCopy.Copy
    
    ' Paste the copied data back into the same place in the "CTR Template" sheet
    ' using PasteSpecial to paste values, to remove any formulas if present
    rngToCopy.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    ' Clear the clipboard
    Application.CutCopyMode = False
    ' Start updating the screen
    Application.ScreenUpdating = True
        
    ' Ask for the date
    dte = Sheets("Instructions").Range("C3").Value
    sCTRFileName = "ctr_template.xlsx" ' This is in case this name changes, you won't have to change it 15 places below
    sCTRFile = "\\hum-vmqb-01\Billing\Duke\Resources\" & sCTRFileName ' Adjust this to where the CTR template is being saved
    sDestFileLocation = ThisWorkbook.Sheets("Instructions").Range("C5").Value & "Outputs\CTR " ' Adjust this location to where you want the Invoice files saved"
    
    ' Stop updating the screen
    Application.ScreenUpdating = False
    
    ' If the CTR Template file is not open, open it so we can copy data over
    If Not IsWorkBookOpen(sCTRFileName) Then
        Workbooks.Open sCTRFile
    End If
    
    ' Set variables for copy and destination sheets
    Set wsDest = Workbooks(sCTRFileName).Worksheets("Template for Vendors")
    
    ' Get the first Invoice # in spreadsheet
    sInvoice = wsCTR.Range("A2").Value
    
    ' Set some of the Invoice template fields from what we have already
    wsDest.Cells(4, 2).Value = dte
    wsDest.Cells(4, 6).Value = sInvoice
    iFirstInvoiceRow = 2
    
    For i = 2 To wsCTR.Cells(2, 1).End(xlDown).Row + 1
        
        If wsCTR.Cells(i, 1).Value <> sInvoice Then
            ' Clear data in case there is something there
            wsDest.Range("A9:M2500").Clear
            
            ' Copy only the rows associated with a specific Invoice
            wsCTR.Range("B" & iFirstInvoiceRow & ":N" & i - 1).Copy _
            wsDest.Range("A9")
            
            iFirstInvoiceRow = i
            
            ' Construct the full path of the file to save
            Dim sFullPath As String
            sFullPath = sDestFileLocation & sInvoice & ".xlsx"
            
            ' Check if the file already exists. If it does, skip the file creation process
            If Dir(sFullPath) = "" Then
                ' Save file with name that includes Invoice #
                Application.DisplayAlerts = False
                wsDest.SaveAs sFullPath, 51
                Application.DisplayAlerts = True
            
                ' Close the new Invoice
                Workbooks.Open(sFullPath).Close
                ' Notify that the File was created and saved
                Debug.Print "created: " & sFullPath
            Else
                ' Optional: Notify that the file was skipped or log this information
                Debug.Print "skipped: " & sFullPath
            End If
            
            ' Attempt to reopen the template file
            If Not IsWorkBookOpen(sCTRFileName) Then
                Workbooks.Open sCTRFile
                Set wsDest = Workbooks(sCTRFileName).Worksheets("Template for Vendors")
            End If
            
            ' If this is a new Invoice #, store the # and place it on the new Invoice template
            sInvoice = wsCTR.Cells(i, 1).Value
            wsDest.Cells(4, 6).Value = sInvoice
            wsDest.Cells(4, 2).Value = dte
    
            ' State checking logic
            Dim cellValue As String
            cellValue = wsCTR.Cells(i, "O").Value
            
            If cellValue = "FL" Then
                wsDest.Range("A4").Value = "TD-FL"
            ElseIf cellValue = "NC" Or cellValue = "SC" Then
                wsDest.Range("A4").Value = "TD-NC-SC"
            Else
                wsDest.Range("A4").Value = "TD-NC-SC"
            End If
            
        End If
    Next i
    
    ' Close template
    If IsWorkBookOpen(sCTRFileName) Then
        Application.DisplayAlerts = False
        Workbooks(sCTRFileName).Close
        Application.DisplayAlerts = True
    End If
    
    ' Start updating the screen again
    Application.ScreenUpdating = True
End Sub

Sub ProcessAndSortFiles()
    'Set up error handling
    On Error GoTo ErrorHandler

    ' Set up variables for operations
    Dim i As Integer
    Dim last_row As Integer
    Dim emailData As Worksheet
    Dim weeklyFolder As String
    Dim attachmentFolder As String
    Dim invoiceFile As String
    Dim overTimeFile As String
    Dim ctrFile As String
    Dim timesheetFile As String
    Dim checkFile As String
    Dim fso As Object
    Dim logFilePath As String
    Dim logFile As Object

    ' Save our Intaact data into the data model
    SaveINVsToDataModel

    ' Initialize FileSystemObject and Worksheets
    Set emailData = ThisWorkbook.Sheets("Email Template")
    last_row = Application.WorksheetFunction.CountA(emailData.Range("B:B"))
    
    ' Set the folder and file locations as variables to make this easier to read
    weeklyFolder = ThisWorkbook.Sheets("Instructions").Range("C5").Value
    attachmentFolder = ThisWorkbook.Sheets("Instructions").Range("C5").Value & "Outputs\"
    overTimeFile = attachmentFolder & "\EE OT Breakdown " & ThisWorkbook.Sheets("Instructions").Range("C3").Value & ".xlsx"
    
    ' Create the file manipulation object and the Debug file
    Set fso = CreateObject("Scripting.FileSystemObject")
    logFilePath = ThisWorkbook.Path & "\Missing Files.txt"
    
    ' Setup log file for operations
    If Not fso.FileExists(overTimeFile) Then
        MsgBox "You need to make the EE Overtime file with the Pivots."
    Else
        logFile.WriteLine "GOOD: EE OT file exists."
    End If
    
    ' Setup log file for operations
    If Not fso.FileExists(logFilePath) Then
        Set logFile = fso.CreateTextFile(logFilePath, True)
    Else
        Set logFile = fso.OpenTextFile(logFilePath, 8)
    End If

    'Loop through each row to check for file existence and sort files
    For i = 2 To last_row
        ' File checks and markings as in `missingFiles`
        invoiceFile = attachmentFolder & emailData.Range("C" & i) & ".pdf"
        ctrFile = attachmentFolder & "CTR " & emailData.Range("C" & i) & ".xlsx"
        timesheetFile = attachmentFolder & emailData.Range("B" & i) & ".pdf"

        ' Check for Invoice, CTR, Timesheet files existence and mark in the worksheet
        MarkMissingFile emailData.Range("A" & i), invoiceFile, "I ", "_ "
        MarkMissingFile emailData.Range("A" & i), ctrFile, "C", "_"
        MarkMissingFile emailData.Range("A" & i), timesheetFile, " T", " _"
    
        ' Sort and move/copy files into the Invoice Subfolders
        Dim folderPath As String
        folderPath = weeklyFolder & "Invoices\" & emailData.Range("C" & i).Text

        ' Ensure the destination folder exists
        If Not fso.FolderExists(folderPath) Then
            fso.CreateFolder (folderPath)
        End If

        ' Move and log operations for each file
        MoveWithLog fso, emailData.Range("I" & i).Text, folderPath, logFile
        MoveWithLog fso, emailData.Range("J" & i).Text, folderPath, logFile
        MoveWithLog fso, emailData.Range("K" & i).Text, folderPath, logFile
        CopyWithLog fso, emailData.Range("L" & i).Text, folderPath, logFile
    Next i

    logFile.Close
    MsgBox "Files Moved Successfully. Open 'Missing Files.txt' to find any additional files that may be required."
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred. Please check the log file for more details."
    If Not logFile Is Nothing Then logFile.Close
End Sub

Sub SendEmails(startRow As Integer, endRow As Integer, action As String)
    Dim OutApp As Object
    Dim outmail As Object
    Dim i As Long
    Dim strbody As String
    Dim whoTo As String
    Dim emailData As Worksheet
    Set emailData = ThisWorkbook.Sheets("Email Template")
    
    ' Send emails
    Set OutApp = CreateObject("Outlook.Application")
    
    For i = startRow To endRow
        Set outmail = OutApp.createitem(0)

        ' Check the region and assign the appropriate email address
        Select Case emailData.Range("E" & i).Value
            Case "Piedmont"
                whoTo = "scott.smith@FlaggerForce.com"
            Case "Storm"
                whoTo = "DistStormVendorInv@duke-energy.com"
            Case "Duke"
                whoTo = "DistributionCTR@duke-energy.com"
            Case Else
                whoTo = "DistributionCTR@duke-energy.com" ' Fallback email, change as necessary
        End Select
        
        With outmail
            .Sentonbehalfofname = "billingteam@flaggerforce.com"
            .To = whoTo
            .CC = ""
            .bcc = ""
            .Subject = emailData.Range("G" & i) & " " & emailData.Range("F" & i) & " " & emailData.Range("C" & i) & " " & emailData.Range("H" & i)
            .HTMLbody = strbody & .HTMLbody
            .attachments.Add emailData.Range("M" & i).Text
            .attachments.Add emailData.Range("N" & i).Text
            .attachments.Add emailData.Range("O" & i).Text
            .attachments.Add emailData.Range("P" & i).Text
            
            If action = "Send" Then
                .send
            Else
                .display
            End If
        End With
    
        Set outmail = Nothing
    Next i
End Sub

Sub CreateCTRs()
    Slow_n_SafeFormulas
    Wait_a_minute
    CreateCtrFiles
End Sub

Sub SubmitDuke()
    Dim lastEmail As Integer
    Dim emailData As Worksheet
    Set emailData = ThisWorkbook.Sheets("Email Template")
    
    ' Find the last row of data so that you know
    lastEmail = Application.WorksheetFunction.CountA(emailData.Range("B:B"))
    MsgBox ("Last row of data: " & lastEmail)
    
    EmailForm.Show
End Sub

Function SaveINVsToDataModel()
    Dim sourceSheet As Worksheet
    Dim destinationWorkbook As Workbook
    Dim destinationSheet As Worksheet
    Dim lastRowSource As Long
    Dim lastRowDestination As Long
    Dim destinationPath As String
    Dim sourceRange As Range, destRange As Range
    
    ' Define the source sheet
    Set sourceSheet = ThisWorkbook.Sheets("Order Entry Transactions")
    
    ' Path to the destination workbook - adjust if necessary
    destinationPath = "\\hum-vmqb-01\Profitability_Reporting\Billing\Duke\Resources\duke_data_model.xlsx"
    
    ' Attempt to reference the destination workbook if it's already open
    On Error Resume Next ' To avoid error if the workbook is not already open
    Set destinationWorkbook = Workbooks("duke_data_model.xlsx")
    If destinationWorkbook Is Nothing Then
        ' If not found, try to open it
        Set destinationWorkbook = Workbooks.Open(destinationPath)
    End If
    On Error GoTo 0 ' Turn back on regular error handling
    If destinationWorkbook Is Nothing Then
        MsgBox "Failed to open the destination workbook. Path: " & destinationPath, vbCritical
        Exit Function
    End If
    
    ' Define the destination sheet
    Set destinationSheet = destinationWorkbook.Sheets("Invoices_n_WOs")
    
    ' Determine the last row of data in the source sheet
    lastRowSource = sourceSheet.Cells(sourceSheet.Rows.Count, "A").End(xlUp).Row
    
    ' Find the first empty row in the destination sheet
    lastRowDestination = destinationSheet.Cells(destinationSheet.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Copy from source to destination without using the clipboard, directly assigning values
    Set sourceRange = sourceSheet.Range("A2:O" & lastRowSource)
    Set destRange = destinationSheet.Range("A" & lastRowDestination).Resize(sourceRange.Rows.Count, sourceRange.Columns.Count)
    Debug.Print lastRowSource
    Debug.Print lastRowDestination
    ' Directly transferring values
    destRange.Value = sourceRange.Value
    
    ' Save and close the "Data Model" workbook
    With destinationWorkbook
        .Save
        .Close False
    End With
    
    ' Save this workbook
    With ThisWorkbook
        .Save
    End With
End Function

Sub CreateWeeklyFolder()
    Dim ServerFolder As FileDialog
    Dim ServerFolderPath As String
    Dim NewFolderName As String
    Dim fso As Scripting.FileSystemObject
    
    ' Set the source folder path directly
    Dim TemplateFolderPath As String
    TemplateFolderPath = "\\hum-vmqb-01\Billing\Duke\Resources\Template"
    
    ' Create FileDialog instance for selecting the target folder
    Set ServerFolder = Application.FileDialog(msoFileDialogFolderPicker)
    
    ' Prompt user to select target folder
    With ServerFolder
        .Title = "Select Target Folder"
        .AllowMultiSelect = False
        If .Show = -1 Then
            ServerFolderPath = .SelectedItems(1)
        Else
            MsgBox "No folder selected. Exiting."
            Exit Sub
        End If
    End With
    
    ' Prompt user for new folder name in mm.dd format with the current date as default input
    NewFolderName = InputBox("Enter the new folder name in mm.dd format:", "New Folder Name", Format(Date, "mm.dd"))
    If NewFolderName = "" Then
        MsgBox "No folder name provided. Exiting."
        Exit Sub
    End If
    
    ' Create FileSystemObject instance
    Set fso = New Scripting.FileSystemObject
    
    ' Copy the entire source folder into the target folder with the new folder name
    On Error Resume Next
    fso.CopyFolder TemplateFolderPath, ServerFolderPath & Application.PathSeparator & NewFolderName, True
    If Err.Number <> 0 Then
        MsgBox "Error in copying folder: " & Err.Description
    Else
        MsgBox "Folder copied successfully."
    End If
    On Error GoTo 0
    
    ' Clean up
    Set fso = Nothing
    Set ServerFolder = Nothing
End Sub


