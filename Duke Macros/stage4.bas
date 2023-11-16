Attribute VB_Name = "stage4"
Option Explicit

Sub ImportOATs()
    Dim wsOATs As Worksheet
    Dim wsEmail As Worksheet
    Dim lastRowOATs As Long
    
    ' Tell the user what will be done
    MsgBox "Select the `Order Entry Transactions` file that you have downloaded from Intaact Sage."

    ' Check if the "Order Entry Transactions" worksheet exists and import billing details
    If Not WorksheetExists("Order Entry Transactions", ThisWorkbook) Then
        Set wsOATs = ThisWorkbook.Worksheets.Add
        wsOATs.Name = "Order Entry Transactions"
    Else
        Set wsOATs = ThisWorkbook.Worksheets("Order Entry Transactions")
    End If
    ImportDataToWorksheet "Order Entry Transactions", "A1"
    
    ' Turn off screen updating
    Application.ScreenUpdating = False
    
    ' Find the last row in column A of wsOATs
    lastRowOATs = FindLastRow("A", wsOATs)
    
    ' Copy the formulas in range O2:R2 down the rest of the table without using the clipboard
    wsOATs.Range("O2:R" & lastRowOATs).FillDown
    
    ' Filter Col R for "0"
    AutoFilterColumn wsOATs, "R", "0", lastRowOATs
    
    ' Make Email Template worksheet visible and copy the data without using the clipboard
    Set wsEmail = AddOrGetWorksheet("Email Template", wsEmail)
    wsEmail.Visible = True
    wsEmail.Range("B2:C" & lastRowOATs).Value = wsOATs.Range("B2:C" & lastRowOATs).Value
    
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
    Dim i As Long
    Dim lastRow As Long
    Dim attachmentFolder As String
    Dim invoiceFile As String
    Dim ctrFile As String
    Dim timesheetFile As String
    Dim emailData As Worksheet
    Dim missingIndicator As String
    Dim dukeInstructions As Worksheet
    
    ' Set references to worksheets
    Set emailData = ThisWorkbook.Sheets("Email Template")
    Set dukeInstructions = ThisWorkbook.Sheets("DukeInstructions")
    
    ' Get last row with data in column B
    lastRow = FindLastRow("B", emailData)
    
    ' Location of attachment folder
    attachmentFolder = dukeInstructions.Range("B5").Value & "Outputs\"
    
    ' Loop through each row checking for each of the files
    For i = 2 To lastRow
    
        invoiceFile = emailData.Range("C" & i).Value & ".pdf"
        ctrFile = "CTR " & emailData.Range("C" & i).Value & ".xlsx"
        timesheetFile = emailData.Range("B" & i).Value & ".pdf"
        missingIndicator = ""
        
        'Invoice
        If Not FileExists(attachmentFolder & invoiceFile) Then
            missingIndicator = "I "
        Else
            missingIndicator = "_ "
        End If
        
        'CTR
        If Not FileExists(attachmentFolder & ctrFile) Then
            missingIndicator = missingIndicator & "C"
        Else
            missingIndicator = missingIndicator & "_"
        End If
        
        'Timesheet
        If Not FileExists(attachmentFolder & timesheetFile) Then
            missingIndicator = missingIndicator & " T"
        Else
            missingIndicator = missingIndicator & " _"
        End If
        
        ' Update the indicator in column A
        emailData.Range("A" & i).Value = missingIndicator
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
    
    Dim i As Long
    Dim strbody As String
    Dim emailData As Worksheet
    Dim attachmentsPaths As New Collection
    Dim subject As String
    Dim toRecipients As String
    Dim emailAttachment As String
    
    Set emailData = ThisWorkbook.Sheets("Email Template")
    
    ' Send emails
    For i = startRow To endRow
        ' Build email subject line
        subject = emailData.Range("F" & i).Value & " " & _
                  emailData.Range("E" & i).Value & " " & _
                  emailData.Range("C" & i).Value & " " & _
                  emailData.Range("G" & i).Value
        
        ' Define recipients
        toRecipients = "DistributionCTR@duke-energy.com"
        
        ' Add attachment file paths to the collection
        attachmentsPaths.Clear
        emailAttachment = emailData.Range("L" & i).Value
        If FileExists(emailAttachment) Then attachmentsPaths.Add emailAttachment
        emailAttachment = emailData.Range("M" & i).Value
        If FileExists(emailAttachment) Then attachmentsPaths.Add emailAttachment
        emailAttachment = emailData.Range("N" & i).Value
        If FileExists(emailAttachment) Then attachmentsPaths.Add emailAttachment
        emailAttachment = emailData.Range("O" & i).Value
        If FileExists(emailAttachment) Then attachmentsPaths.Add emailAttachment
        
        ' Create and send or display email
        CreateAndSendEmail toRecipients, subject, "", attachmentsPaths, action
    Next i

End Sub

Sub SortAndSend()
    Dim lastEmail As Integer
    Dim emailData As Worksheet
    Set emailData = ThisWorkbook.Sheets("Email Template")
    
    SortFiles
    lastEmail = Application.WorksheetFunction.CountA(emailData.Range("B:B"))
    MsgBox ("Jot down this number so that you remember it for the second round of email submissions.\n\nLast row of data: " & lastEmail)
    
    ' Send emails
    EmailForm.Show
    
    ' Create Next Weeks Folder
    CreateWeeklyFolder
End Sub
