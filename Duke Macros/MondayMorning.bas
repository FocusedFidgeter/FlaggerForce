Attribute VB_Name = "MondayMorning"
Option Explicit
Public thisWeeksBillingDate As String
Public thisWeeksServerFolder As String
Public yearYYYY As String
Public yearYY As String
Public dateMMDD As String

Function ImportDataToWorksheet(targetSheetName As String, targetRange As String)

    ' Declare variables
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
    fileName = Application.GetOpenFilename(FileFilter:="Excel Files (*.xls*), *.xls*, All Files (*.*), *.*, CSV Files (*.csv), *.csv", _
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

    ' Inform the user that the import is complete
    'MsgBox "Data imported successfully!"
End Function

Function ImportBillingDetails()

    ' Declare variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim deleteRange As Range

    ' Tell the user what will be done
    MsgBox "Select the `Billing Details` file that you have downloaded from PowerBI."
    ' Import Billing Details from PowerBI
    ImportDataToWorksheet "PowerBI Details", "G1"
    
    ' Turn off screen updating
    Application.ScreenUpdating = False

    ' Set ws
    Set ws = ThisWorkbook.Worksheets("PowerBI Details")

    ' 0. Find the last row in the data and format necessary columns
    With ws
        lastRow = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
        .Range("K2:K" & lastRow).NumberFormat = "m/d/yyyy"
        .Range("L2:L" & lastRow).NumberFormat = "h:mm AM/PM"
    End With

    ' 1. Set formulas
    ws.Range("A2").Formula = "=TRUNC(I2&K2, 0)"
    ws.Range("B2").Formula = "=IF(Z2="""",""yes"",IF(AB2=""Billable Cancelled"",""CXL"",""no""))"
    ws.Range("C2").Formula = "=CONCAT(AW2, "", "", AX2, "", "", AY2, "" "", AZ2)"
    ws.Range("D2").Formula = "=A2=A1"
    ws.Range("E2").Formula = "=COUNTIF('Raw Hours'!A:A,$A2)"
    ws.Range("F2").Formula = "=W2-E2"


    ' 2. Fill in the rest of the formulas
    Sheets("PowerBI Details").Select
    Range("A2:F2").Select
    Selection.AutoFill Destination:=Range("A2:F" & lastRow)
    
    
    ' 3. Process Data
    ' This is a loop that checks if col Z = "Cancelled" and col AB  = "" then delete the row
    ' If col Z = "Cancelled" and col AB is "Billable CXL" and col X (equipment ordered) is >0 set it to 0
    For i = lastRow To 2 Step -1
        If ws.Cells(i, "D").Value = True Or _
            (ws.Cells(i, "Z").Value = "Cancelled" And ws.Cells(i, "AB").Value = "") Then
    
            ' The first time DeleteRange is set, Set it; Otherwise, Union it
            If deleteRange Is Nothing Then
                Set deleteRange = ws.Rows(i)
            Else
                Set deleteRange = Union(deleteRange, ws.Rows(i))
            End If
        End If
        
        If ws.Cells(i, "Z").Value = "Cancelled" And ws.Cells(i, "AB").Value = "Billable Cancelled" And ws.Cells(i, "X").Value > 0 Then
            ws.Cells(i, "X").Value = 0
        End If
    Next i
    
    ' Delete the collected Rows
    If Not deleteRange Is Nothing Then deleteRange.Delete
    
    ' 4. Copy and paste to fix the reference errors caused in Column D "Duplicate"
    Sheets("PowerBI Details").Select
    
    ' Find the NEW last row
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Range("D2").Formula = "=A2=A1"
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D" & lastRow)
    
    ' 5. Sort Table
    Dim sortRange As Range
    Set sortRange = ws.Range("A1:BC" & lastRow)
    
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=ws.Range("Z2:Z" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add2 Key:=ws.Range("AB2:AB" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add2 Key:=ws.Range("I2:I" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add2 Key:=ws.Range("K2:K" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range("A1:BC" & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
        
    ' Turn on screen updating
    Application.ScreenUpdating = True

End Function

Function ImportEEHistory()
    
    ' Declare variables
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' Set the target worksheet and unhide it
    Set ws = ThisWorkbook.Worksheets("EE History")
    ws.Visible = True
    
    ' Tell the user what will be done
    MsgBox "Select the `EE History` file that you have downloaded from TDOC."
    ' Import EE Pre-PDES-Closed Hours from TDOC
    ImportDataToWorksheet "EE History", "A1"

    With ws
        ' 1. Find the last row in the data
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' 2. Format Columns
        Range("G2:G" & lastRow).NumberFormat = "m/d/yyyy"
        Range("U2:U" & lastRow).NumberFormat = "0"
    End With

    ' 3. Fill in the rest of the formulas
    Sheets("EE History").Select
    Range("M2:U2").Select
    Selection.AutoFill Destination:=Range("M2:U" & lastRow)
    
    ' 4. Sort Table
    Dim sortRange As Range
    Set sortRange = ws.Range("A1:U" & lastRow)
    sortRange.Sort Key1:=ws.Range("M1"), Order1:=xlAscending, Key2:=ws.Range("D1"), Order2:=xlAscending, Header:=xlYes

End Function

Function SubmitDukeLunches()
    
    Dim OutApp As Object
    Dim outmail As Object
    Dim strbody As String
    
    ' Send emails
    Set OutApp = CreateObject("Outlook.Application")
    ' Send the newly created Lunch file to Tim Yeatts (or other/future IT)
    Set outmail = OutApp.createitem(0)
    strbody = "Good Morning,\n\nHere is Duke's lunches that need adjusted. Thank you!"
    With outmail
        .To = "paul.devey@flaggerforce.com"
        .Sentonbehalfofname = "billingteam@flaggerforce.com"
        .CC = ""
        .bcc = ""
        .Subject = "Lunch Adjustments for Duke Energy"
        .HTMLbody = strbody & .HTMLbody
        .attachments.Add newWb
        
        If action = "Send" Then
            .send
        Else
            .display
        End If
    End With

    Set outmail = Nothing
End Function

Function ProcessDukeLunches()

    Dim wb As Workbook
    Dim wsTDOC As Worksheet
    Dim wsUnder As Worksheet
    Dim wsOver As Worksheet
    Dim newWb As Workbook

    Set wb = ThisWorkbook

    ' Create a new workbook
    Set newWb = Workbooks.Add

    ' Create new "Under Reported" and "Over Reported" sheets in the new workbook
    newWb.Worksheets.Add(After:=newWb.Worksheets(newWb.Worksheets.Count)).Name = "Under Reported"
    newWb.Worksheets.Add(After:=newWb.Worksheets(newWb.Worksheets.Count)).Name = "Over Reported"

    Set wsTDOC = wb.Worksheets("EE History")
    Set wsUnder = newWb.Worksheets("Under Reported")
    Set wsOver = newWb.Worksheets("Over Reported")
    
    ' Delete "Sheet1"
    Application.DisplayAlerts = False ' Suppress the confirmation prompt
    newWb.Sheets("Sheet1").Delete
    Application.DisplayAlerts = True ' Turn the confirmation prompt back on

    'Filter & copy under reported records
    wsTDOC.Rows(1).AutoFilter _
        Field:=18, _
        Criteria1:="=Under", _
        Operator:=xlFilterValues
    wsTDOC.UsedRange.Copy
    wsUnder.Range("A1").PasteSpecial Paste:=xlPasteValues

    'Filter & copy over reported records
    wsTDOC.Rows(1).AutoFilter _
        Field:=18, _
        Criteria1:="=Over", _
        Operator:=xlFilterValues
    wsTDOC.UsedRange.Copy
    wsOver.Range("A1").PasteSpecial Paste:=xlPasteValues

    ' Remove the filters
    wsTDOC.AutoFilterMode = False
    
    ' Formatting Sheets Data
    Dim ws As Worksheet
    For Each ws In newWb.Worksheets
        ws.Columns("G:G").NumberFormat = "m/d/yyyy"
        ws.Columns("K:L").NumberFormat = "[$-x-systime]h:mm:ss AM/PM"
        ws.Cells.EntireColumn.AutoFit
    Next ws

    ' Save the new workbook as "Duke Lunches mm.dd"
    newWb.SaveAs thisWeeksServerFolder & "Outputs\Duke Lunches " & dateMMDD & ".xlsx", FileFormat:=51
    
    ' Close the new workbook
    newWb.Close
    
    ' Hide the worksheet "EE History"
    wsTDOC.Visible = False
    
    ' Send the lunch file to IT
    SubmitDukeLunches
    
End Function

Function AddressCleanup()
    Dim wb As Workbook
    Dim wsPowerBI As Worksheet
    Dim wsAddressCleanup As Worksheet
    Dim wsAddressCheck As Worksheet
    Dim lastRow As Long
    Dim pt As PivotTable

    Set wb = ThisWorkbook
    Set wsPowerBI = wb.Worksheets("PowerBI Details")
    Set wsAddressCleanup = wb.Worksheets("Address Cleanup")
    Set wsAddressCheck = wb.Worksheets("Address Check")

    ' 1) Find the last row in the "PowerBI Details" sheet
    lastRow = wsPowerBI.Cells(wsPowerBI.Rows.Count, "A").End(xlUp).Row

    With wsAddressCleanup
        .Range("A2").Formula = "=TRUNC(B2&C2,0)"
        .Range("B2").Formula = "='PowerBI Details'!I2"
        .Range("C2").Formula = "='PowerBI Details'!K2"
        .Range("D2").Formula = "='PowerBI Details'!B2"
        .Range("E2").Formula = "=TRIM(IF(I2=0,IF(H2=0,IF(G2=0,I2,G2),H2),I2))"
        .Range("F2").Formula = "=PROPER(TRIM(IF(K2=0,J2,K2)))"
        .Range("G2").Formula = "='PowerBI Details'!AG2"
        .Range("H2").Formula = "='PowerBI Details'!AH2"
        .Range("I2").Formula = "='PowerBI Details'!AJ2"
        .Range("J2").Formula = "='PowerBI Details'!C2"
        .Range("K2").Formula = "='PowerBI Details'!AK2"
        .Range("L2").Formula = "='PowerBI Details'!AY2"
        .Range("M2").Formula = "='PowerBI Details'!W2"
        
        ' 2. Fill in the rest of the formulas
        .Range("A2:M2").AutoFill Destination:=.Range("A2:M" & lastRow)

        ' 3. Paste the columns as values and add a filter to all the headers
        .Range("A2:M" & lastRow).Value = .Range("A2:M" & lastRow).Value
        .Range("A1:M1").AutoFilter
    End With

    ' Select the Address Check worksheet
    wsAddressCheck.Select

    ' Reference and Refresh the "AddressCheck" PivotTable
    Set pt = wsAddressCheck.PivotTables("AddressCheck")
    pt.RefreshTable
End Function

Sub ProcessPrelim()
    ' Set the Globals
    thisWeeksBillingDate = InputBox("Enter the billing date in mm.dd.yyyy format: ", "Billing Date", Format(Now(), "mm.dd.yyyy"))
    ' Extract the year and date parts from the billing date
    yearYYYY = Right(thisWeeksBillingDate, 4)
    yearYY = Right(thisWeeksBillingDate, 2)
    dateMMDD = Left(thisWeeksBillingDate, 5)
    
    ' Save the
    Sheets("DukeInstructions").Range("B3").Value = thisWeeksBillingDate
    ' Create the server folder path
    thisWeeksServerFolder = Sheets("DukeInstructions").Range("B5").Value
    'Debug.Print thisWeeksServerFolder ' Uncomment this line if you change the folder location. Make sure it works!
    
    ' Grab the Duke Billing Details from PowerBI
    ImportBillingDetails
    ' Grab the EE History from TDOC. These are NOT finalized numbers
    ImportEEHistory
    ' Process the hours to ensure they follow Duke's policy (Everything >6hrs takes a lunch, unless a specific note is included)
    ProcessDukeLunches
    ' Copy relevent info for the WO#s and Addresses
    AddressCleanup
End Sub


