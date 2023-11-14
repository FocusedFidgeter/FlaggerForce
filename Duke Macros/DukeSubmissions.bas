Attribute VB_Name = "DukeSubmissions"
Option Explicit

Function ImportDataToWorksheet(targetSheetName As String, targetRange As String)
    Dim sourceWorkbook As Workbook
    Dim sourceWorksheet As Worksheet
    Dim destWorkbook As Workbook
    Dim destWorksheet As Worksheet
    Dim fileName As Variant

    ' Set the destination workbook
    Set destWorkbook = ThisWorkbook

    ' Check if the destination sheet exists, if not create it
    On Error Resume Next
    Set destWorksheet = destWorkbook.Worksheets(targetSheetName)
    If destWorksheet Is Nothing Then
        Set destWorksheet = destWorkbook.Worksheets.Add
        destWorksheet.Name = targetSheetName
    Else
        destWorksheet.Visible = xlSheetVisible
    End If
    On Error GoTo 0

    ' Prompt the user to select a file
    fileName = Application.GetOpenFilename(FileFilter:="CSV Files (*.csv), *.csv, All Files (*.*), *.*, CSV Files (*.csv), *.csv, Excel Files (*.xls*), *.xls*", _
                                           Title:="Select a file to import")
    
    ' Exit if the user cancels the file selection
    If fileName = False Then
        MsgBox "No file selected. Operation canceled."
        Exit Function
    End If

    ' Open the source workbook
    Set sourceWorkbook = Workbooks.Open(fileName)

    ' Set the source worksheet (assuming data is in the first sheet)
    Set sourceWorksheet = sourceWorkbook.Worksheets(1)

    ' Copy the contiguous range starting from A1 in the source worksheet to the destination worksheet
    Dim firstCell As Range
    Dim lastCell As Range
    Dim contiguousRange As Range

    Set firstCell = sourceWorksheet.Range("A1")
    Set lastCell = sourceWorksheet.Cells(firstCell.End(xlDown).Row, firstCell.End(xlToRight).Column)
    Set contiguousRange = sourceWorksheet.Range(firstCell, lastCell)

    contiguousRange.Copy destWorksheet.Range(targetRange)

    ' Close the source workbook without saving
    sourceWorkbook.Close SaveChanges:=False

End Function
Sub ImportOATs()
    Dim wsOATs As Worksheet
    Dim wsCTR As Worksheet
    Dim wsINV As Worksheet
    Dim wsEmail As Worksheet
    Dim lastRowImport As Long
    Dim lastRowOATs As Long
    Dim i As Long
    
    ' Tell the user what will be done
    MsgBox "Select the `Order Entry Transactions` file that you have downloaded from Intaact Sage."
    
    ' Set the separate WSs
    Set wsOATs = ThisWorkbook.Worksheets("Order Entry Transactions")
    Set wsCTR = ThisWorkbook.Worksheets("CTR Template")
    Set wsINV = ThisWorkbook.Worksheets("INVOICE TEMPLATE")
    Set wsEmail = ThisWorkbook.Worksheets("Email Template")
    
    ' Import Billing Details from PowerBI
    ImportDataToWorksheet "Order Entry Transactions", "A1"
    
    ' Turn off screen updating
    Application.ScreenUpdating = False
    
    With wsOATs
        ' Find the last row in the data
        lastRowOATs = .Cells(.Rows.Count, "A").End(xlUp).Row
        
        ' Copy the Formulas in row 2 and paste them down the rest of the table
        .Range("O2:R2").Copy
        .Range("O3:R" & lastRowOATs).PasteSpecial Paste:=xlPasteFormulas
        
        ' Filter Col E for "CXL"
        .Range("R1:R" & lastRowOATs).AutoFilter Field:=18, Criteria1:="0"
    End With
    
    wsEmail.Visible = True
    ' Copy the WO#s and INV#s into the email template worksheet
    wsOATs.Range("B2:C" & lastRowOATs).Copy Destination:=wsEmail.Range("B2")

    ' Turn on screen updating
    Application.ScreenUpdating = True
    ' Inform the user that the import is complete
    MsgBox "Invoice Summary imported successfully!"

End Sub

Function IsWorkBookOpen(Name As String) As Boolean
    Dim xWb As Workbook
    On Error Resume Next
    Set xWb = Application.Workbooks.Item(Name)
    IsWorkBookOpen = (Not xWb Is Nothing)
End Function
Sub CreateCTRs()
    Dim wsImport As Worksheet
    Dim wsInvoice As Worksheet
    Dim wsCTR As Worksheet
    Dim rngToCopy As Range
    Dim lastRow As Long
    Dim wsCopy As Worksheet
    Dim wsDest As Worksheet
    Dim sInvoice As String
    Dim dte As String
    Dim sCTRFile As String
    Dim sDestFileLocation As String
    Dim iFirstInvoiceRow As Integer
    Dim i As Integer
    Dim sCTRFileName As String

    Set wsImport = ThisWorkbook.Sheets("IMPORT")
    Set wsInvoice = ThisWorkbook.Sheets("INVOICE TEMPLATE")
    Set wsCTR = ThisWorkbook.Sheets("CTR Template")

    wsCTR.Visible = True
    wsInvoice.Visible = True

    lastRow = wsImport.Cells(wsImport.Rows.Count, "A").End(xlUp).Row

    ' Set the range to copy from A1 to O and the last row in the "CTR Template" sheet
    Set rngToCopy = wsCTR.Range("A1:O" & lastRow)
    
    ' Stop updating the screen
    Application.ScreenUpdating = False
    
    ' Copy the range
    rngToCopy.Copy
    
    ' Paste the copied data into the "INVOICE TEMPLATE" sheet starting from cell A1
    wsInvoice.Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    ' Clear the clipboard
    Application.CutCopyMode = False
    
    ' Start updating the screen
    Application.ScreenUpdating = True
        
    ' Ask for the date
    dte = Sheets("DukeInstructions").Range("B3").Value
    sCTRFileName = "CTR Template.xlsx" ' This is in case this name changes, you won't have to change it 15 places below
    sCTRFile = "\\hum-vmqb-01\Billing\Duke\Resources\" & sCTRFileName ' Adjust this to where the CTR template is being saved
    sDestFileLocation = ThisWorkbook.Sheets("DukeInstructions").Range("B5").Value & "Outputs\CTR " ' Adjust this location to where you want the Invoice files saved"
    
    ' Ask for the date
    dte = Sheets("DukeInstructions").Range("B3").Value
    
    ' Stop updating the screen
    Application.ScreenUpdating = False
    
    ' If the CTR Template file is not open, open it so we can copy data over
    If Not IsWorkBookOpen(sCTRFileName) Then
        Workbooks.Open sCTRFile
    End If
    
    ' Set variables for copy and destination sheets
    Set wsCopy = ThisWorkbook.Worksheets("INVOICE TEMPLATE")
    Set wsDest = Workbooks(sCTRFileName).Worksheets("Template for Vendors")
    
    ' Get the first Invoice # in spreadsheet
    sInvoice = wsCopy.Range("A2").Value
    
    ' Set some of the Invoice template fields from what we have already
    wsDest.Cells(4, 2).Value = dte
    wsDest.Cells(4, 6).Value = sInvoice
    iFirstInvoiceRow = 2
    
    For i = 2 To wsCopy.Cells(2, 1).End(xlDown).Row + 1
    
        If wsCopy.Cells(i, 1).Value <> sInvoice Then
            ' Clear data in case there is something there
            wsDest.Range("A9:M2500").Clear
            
            ' Copy  only the rows associated with a specific Invoice
            wsCopy.Range("B" & iFirstInvoiceRow & ":N" & i - 1).Copy _
            wsDest.Range("A9")
            
            iFirstInvoiceRow = i
            
            ' Save file with name that includes Invoice #
            Application.DisplayAlerts = False
            Debug.Print sDestFileLocation & sInvoice
            wsDest.SaveAs sDestFileLocation & sInvoice, 51
            Application.DisplayAlerts = True
            
            ' Close the new Invoice and re-open the template file
            Workbooks("CTR " & sInvoice & ".xlsx").Close
            
            If Not IsWorkBookOpen(sCTRFileName) Then
                Workbooks.Open sCTRFile
                Set wsDest = Workbooks(sCTRFileName).Worksheets("Template for Vendors")
            End If
            
            ' If this is a new Invoice #, store the # and place it on the new Invoice template
            sInvoice = wsCopy.Range("A" & i).Value
            
            wsDest.Cells(4, 6).Value = sInvoice
            wsDest.Cells(4, 2).Value = dte
    
            ' State checking logic
            Dim cellValue As String
            cellValue = wsCopy.Range("O" & i).Value
            
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
    
    ' Hide the Source Sheet
    wsInvoice.Visible = False
    
    ' Start updating the screen again
    Application.ScreenUpdating = True
End Sub
Sub missingFiles()
    'Set up for the loops
    Dim i As Integer
    Dim j As Integer
    Dim last_row As Integer
    Dim attachmentFolder As String
    Dim invoiceFile As String
    Dim ctrFile As String
    Dim timesheetFile As String
    Dim checkFile As String
    Dim emailData As Worksheet
    Set emailData = ThisWorkbook.Sheets("Email Template")
    
    last_row = Application.WorksheetFunction.CountA(emailData.Range("B:B"))
    attachmentFolder = ThisWorkbook.Sheets("DukeInstructions").Range("B5").Value & "Outputs\"
    'loop through each row checking for each of the four files
    For i = 2 To last_row
    
    
        invoiceFile = emailData.Range("C" & i) & ".pdf"
        ctrFile = "CTR " & emailData.Range("C" & i) & ".xlsx"
        timesheetFile = emailData.Range("B" & i) & ".pdf"
    
        'Invoice
        'Debug.Print attachmentFolder & invoiceFile
        checkFile = Dir(attachmentFolder & invoiceFile)
        If checkFile = "" Then
            emailData.Range("A" & i).Value = "I "
        Else
            emailData.Range("A" & i).Value = "_ "
        End If
        
        'CTR
        'Debug.Print attachmentFolder & ctrFile
        checkFile = Dir(attachmentFolder & ctrFile)
        If checkFile = "" Then
            emailData.Range("A" & i).Value = emailData.Range("A" & i).Text & "C"
        Else
            emailData.Range("A" & i).Value = emailData.Range("A" & i).Text & "_"
        End If
        
        'Timesheet
        'Debug.Print attachmentFolder & timesheetFile
        checkFile = Dir(attachmentFolder & timesheetFile)
        If checkFile = "" Then
            emailData.Range("A" & i).Value = emailData.Range("A" & i).Text & " T"
        Else
            emailData.Range("A" & i).Value = emailData.Range("A" & i).Text & " _"
        End If
        
    Next i

End Sub
Sub SortFiles()
    Dim OutApp As Object
    Dim outmail As Object
    Dim strbody As String
    Dim emailData As Worksheet
    Set emailData = ThisWorkbook.Sheets("Email Template")
    
    'Set up  for the loops
    Dim i As Integer
    Dim last_row As Integer
    
    last_row = Application.WorksheetFunction.CountA(emailData.Range("B:B"))
    
    MsgBox ("We are going to group and sort the files into the /Invoice/ folder.")
    
    ' Create folder and move files
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folderPath As String
    
    For i = 2 To last_row
    
        folderPath = ThisWorkbook.Path & "\Invoices\" & emailData.Range("C" & i)
    
        If Not fso.FolderExists(folderPath) Then
            fso.CreateFolder (folderPath)
        End If
    
        fso.MoveFile Source:=emailData.Range("H" & i).Text, Destination:=folderPath & "\" & fso.GetFileName(emailData.Range("H" & i).Text)
        fso.MoveFile Source:=emailData.Range("I" & i).Text, Destination:=folderPath & "\" & fso.GetFileName(emailData.Range("I" & i).Text)
        fso.MoveFile Source:=emailData.Range("J" & i).Text, Destination:=folderPath & "\" & fso.GetFileName(emailData.Range("J" & i).Text)
        fso.CopyFile Source:=emailData.Range("K" & i).Text, Destination:=folderPath & "\" & fso.GetFileName(emailData.Range("K" & i).Text)
    
    Next i

End Sub
Sub SendEmails(startRow As Integer, endRow As Integer, action As String)
    
    Dim OutApp As Object
    Dim outmail As Object
    Dim i As Long
    Dim strbody As String
    Dim emailData As Worksheet
    Set emailData = ThisWorkbook.Sheets("Email Template")
    
    ' I honestly don't really understand this scripting part.
    ' It has been in here since Steve first wrote this macro. - Pd
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folderPath As String
    
    ' Send emails
    Set OutApp = CreateObject("Outlook.Application")
    
    For i = startRow To endRow
        Set outmail = OutApp.createitem(0)
    
        With outmail
            .To = "DistributionCTR@duke-energy.com"
            .Sentonbehalfofname = "billingteam@flaggerforce.com"
            .CC = ""
            .bcc = ""
            .Subject = emailData.Range("F" & i) & " " & emailData.Range("E" & i) & " " & emailData.Range("C" & i) & " " & emailData.Range("G" & i)
            .HTMLbody = strbody & .HTMLbody
            .attachments.Add emailData.Range("L" & i).Text
            .attachments.Add emailData.Range("M" & i).Text
            .attachments.Add emailData.Range("N" & i).Text
            .attachments.Add emailData.Range("O" & i).Text
            
            If action = "Send" Then
                .send
            Else
                .display
            End If
        End With
    
        Set outmail = Nothing
    Next i

End Sub

Sub SortAndSend()
    Dim lastEmail As Integer
    Dim emailData As Worksheet
    Set emailData = ThisWorkbook.Sheets("Email Template")
        
    lastEmail = Application.WorksheetFunction.CountA(emailData.Range("B:B"))
    MsgBox ("Jot down this number so that you remember it for the second round of email submissions.\n\nLast row of data: " & lastEmail)
    
    ' Send emails
    EmailForm.Show
    
    ' Create Next Weeks Folder
    CreateWeeklyFolder
End Sub
