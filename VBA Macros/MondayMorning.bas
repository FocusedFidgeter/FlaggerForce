Attribute VB_Name = "MondayMorning"
Option Explicit
Public thisWeeksBillingDate As String
Public thisWeeksServerFolder As String
Public yearYYYY As String
Public yearYY As String
Public dateMMDD As String

Sub CreateWeeklyFolder()
    Dim ServerFolder As FileDialog
    Dim ServerFolderPath As String
    Dim NewFolderName As String
    Dim fso As Scripting.FileSystemObject

    ' Set the source folder path directly
    Dim TemplateFolderPath As String
    TemplateFolderPath = "\\hum-vmqb-01\Billing\Duke\Resources\Template"

    ' Create FileDialog instance for selecting the target folder
    Set ServerFolderPath = "\\hum-vmqb-01\Billing\Duke\2024"

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

Function ImportBillingDetails()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim deleteRange As Range, dataRange As Range

    ' Tell the user what will be done
    MsgBox "Select the `Billing Details` file that you have downloaded from PowerBI."
    ' Import Billing Details from PowerBI
    ImportDatatoWorksheet "PowerBI Details", "E1", "Excel"

    ' Turn off screen updating
    Application.ScreenUpdating = False

    ' Set ws
    Set ws = ThisWorkbook.Worksheets("PowerBI Details")

    ' Find the last row in the data and format necessary columns
    With ws
        lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
        .Range("I2:I" & lastRow).NumberFormat = "m/d/yyyy"
        .Range("J2:J" & lastRow).NumberFormat = "h:mm AM/PM"
    End With

    ' Set formulas
    ws.Range("A2").Formula = "=TRUNC(G2&I2,0)"
    ws.Range("B2").Formula = "=IF(X2="""",""yes"",IF(Z2=""Billable Cancelled"", IF(OR(AA2=""Attendance"", AA2=""Resources""), ""no"", ""CXL""), ""no""))"
    ws.Range("C2").Formula = "=CONCAT(AU2,"", "",AV2,"", "",AW2,"" "",AX2)"
    ws.Range("D2").Formula = "=A2=A1"

    ' Fill in the rest of the formulas
    Sheets("PowerBI Details").Activate
    Range("A2:D2").Select
    Selection.AutoFill Destination:=Range("A2:D" & lastRow)

    ' Custom Sort by "Cancelled Order", "Billable Cancelled", & "FFID&WorkDate"
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=ws.Range("X2:X" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add2 Key:=ws.Range("Z2:Z" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add2 Key:=ws.Range("A2:A" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange dataRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ' Process Data
    ' This is a loop that checks if col X = "Cancelled" and col Z  = "" then delete the row
    For i = lastRow To 2 Step -1
        If ws.Cells(i, "D").Value = True Or _
            (ws.Cells(i, "X").Value = "Cancelled" And ws.Cells(i, "Z").Value = "") Then

            ' The first time DeleteRange is set, Set it; Otherwise, Union it
            If deleteRange Is Nothing Then
                Set deleteRange = ws.Rows(i)
            Else
                Set deleteRange = Union(deleteRange, ws.Rows(i))
            End If
        End If

        ' If col X = "Cancelled" and col Z is "Billable CXL" and col V (equipment ordered) is >0 set it to 0
        If ws.Cells(i, "X").Value = "Cancelled" And ws.Cells(i, "Z").Value = "Billable Cancelled" And ws.Cells(i, "V").Value > 0 Then
            ws.Cells(i, "V").Value = 0
        End If
    Next i

    ' Delete the collected Rows
    If Not deleteRange Is Nothing Then deleteRange.Delete

    ' Find the NEW last row
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Set dataRange = ws.Range("A1:BA" & lastRow)

    ' Fix the reference errors we just created
    Range("D2").Formula = "=A2=A1"
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D" & lastRow)

    ' Apply filter on column B for 'CXL'
    dataRange.AutoFilter Field:=2, Criteria1:="CXL"

    Sheets("PowerBI Details").Activate

    ' Turn on screen updating
    Application.ScreenUpdating = True
End Function

Function ImportEEHistory()
    Dim ws As Worksheet
    Dim lastRow As Long

    ' Set the target worksheet
    Set ws = ThisWorkbook.Worksheets("EE History")

    ' Tell the user what will be done
    MsgBox "Select the `EE History` file that you have downloaded from TDOC."
    ' Import EE Pre-PDES-Closed Hours from TDOC
    ImportDatatoWorksheet "EE History", "A1", "Excel"

    With ws
        ' Find the last row in the data
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        ' Format Columns
        Range("G2:G" & lastRow).NumberFormat = "m/d/yyyy"
        Range("U2:U" & lastRow).NumberFormat = "0"
    End With

    ' Fill in the rest of the formulas
    Sheets("EE History").Activate
    Range("M2:U2").Select
    Selection.AutoFill Destination:=Range("M2:U" & lastRow)

    ' Sort Table
    Dim sortRange As Range
    Set sortRange = ws.Range("A1:U" & lastRow)
    sortRange.Sort Key1:=ws.Range("M1"), Order1:=xlAscending, Key2:=ws.Range("D1"), Order2:=xlAscending, Header:=xlYes

    ' Go back to the Instructions worksheet
    Sheets("Instructions").Activate
End Function

Function ProcessDukeLunches() As String
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
    Dim lunchesFilePath As String
    lunchesFilePath = thisWeeksServerFolder & "Outputs\Duke Lunches " & dateMMDD & ".xlsx"
    newWb.SaveAs lunchesFilePath, FileFormat:=51
    
    ' Close the new workbook
    newWb.Close

    ' Return the path of the saved file
    ProcessDukeLunches = lunchesFilePath
End Function

Function SendDukeLunchesEmail()
    Dim filePath As String
    Dim OutApp As Object
    Dim outmail As Object
    Dim strbody As String
    
    ' Get the path of the file created by ProcessDukeLunches
    filePath = ProcessDukeLunches()
    
    ' Set up the email body here
    strbody = "Morning Tim! Here's the file for Duke Lunches."
    ' Send emails
    Set OutApp = CreateObject("Outlook.Application")
    Set outmail = OutApp.createitem(0)

    With outmail
        .To = "timothy.yeatts@flaggerforce.com"
        .CC = "payrollteam@flaggerforce.com; brittani.priester@flaggerforce.com"
        .Subject = "Duke Lunches " & Format(Now, "mm.dd")
        .attachments.Add filePath
        .display
        .HTMLbody = strbody & .HTMLbody
    End With
    Set outmail = Nothing
    Set OutApp = Nothing
End Function

Function SaveWOsToDataModel()
    Dim sourceSheet As Worksheet, instructionsSheet As Worksheet
    Dim destinationWorkbook As Workbook
    Dim destinationSheet As Worksheet
    Dim lastRowSource As Long
    Dim lastRowDestination As Long
    Dim destinationPath As String, billingDate As String
    Dim sourceRange As Range, destRange As Range, billingDateRange As Range
    
    ' Define the source sheet
    Set sourceSheet = ThisWorkbook.Sheets("Address Cleanup")
    
    ' Define the instructions sheet (missing in the original code)
    Set instructionsSheet = ThisWorkbook.Sheets("Instructions")  ' Adjust the sheet name as necessary
    
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
    Set destinationSheet = destinationWorkbook.Sheets("WOs_n_Addresses")
    
    ' Determine the last row of data in the source sheet
    lastRowSource = sourceSheet.Cells(sourceSheet.Rows.Count, "A").End(xlUp).Row
    
    ' Find the first empty row in the destination sheet
    lastRowDestination = destinationSheet.Cells(destinationSheet.Rows.Count, "B").End(xlUp).Row + 1
    
    ' Copy from source to destination without using the clipboard, directly assigning values
    Set sourceRange = sourceSheet.Range("A2:F" & lastRowSource)
    Set destRange = destinationSheet.Range("B" & lastRowDestination).Resize(sourceRange.Rows.Count, sourceRange.Columns.Count)
    Debug.Print lastRowSource
    Debug.Print lastRowDestination
    ' Directly transferring values
    destRange.Value = sourceRange.Value
    
    ' Fetch billing date, convert it to the correct format, and apply to the destination sheet
    billingDate = instructionsSheet.Range("C3").Value
    billingDate = Replace(billingDate, ".", "/")
    ' Set range for billing dates in column A in the destination sheet
    Set billingDateRange = destinationSheet.Range("A" & lastRowDestination & ":A" & (lastRowDestination + sourceRange.Rows.Count - 1))
    ' Apply the billing date to each cell in the range and format them
    billingDateRange.Value = billingDate
    billingDateRange.NumberFormat = "mm/dd/yyyy" ' Format as short date
    
    ' Save and close the "Data Model" workbook
    With destinationWorkbook
        .Save
        '.Close False
    End With
    
    MsgBox (lastRowSource - 1) & " rows copied from the 'Address Cleanup' worksheet."
End Function

Sub ProcessPrelim()
    ' Set the Globals
    Dim thisWeeksBillingDate As String
    thisWeeksBillingDate = InputBox("Enter the billing date in mm.dd.yyyy format: ", "Billing Date", Format(Now(), "mm.dd.yyyy"))
    
    ' Extract the year and date parts from the billing date
    yearYYYY = Right(thisWeeksBillingDate, 4)
    yearYY = Right(thisWeeksBillingDate, 2)
    dateMMDD = Left(thisWeeksBillingDate, 5)
    
    ' Save the folder and file
    Sheets("Instructions").Range("C3").Value = thisWeeksBillingDate
    
    ' Create and copy the folder using the billing date
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    ' Define the Template folder path
    Dim TemplateFolderPath As String
    TemplateFolderPath = "\\hum-vmqb-01\Billing\Duke\Resources\Template"
    
    ' Define the Server folder path
    Dim ServerFolderPath As String
    ServerFolderPath = "\\hum-vmqb-01\Billing\Duke\" & yearYYYY
    
    ' New folder name from the billing date (mm.dd format)
    Dim NewFolderName As String
    NewFolderName = dateMMDD
    
    ' Copy the entire source folder into the target folder with the new folder name
    On Error Resume Next
    fso.CopyFolder TemplateFolderPath, ServerFolderPath & "\" & NewFolderName, True
    If Err.Number <> 0 Then
        MsgBox "Error in copying the template folder: " & Err.Description
    Else
        MsgBox "Template folder was copied successfully."
    End If
    On Error GoTo 0
    
    ' Clean up
    Set fso = Nothing
    
    ' Grab the Duke Billing Details from PowerBI
    ImportBillingDetails
    ' Grab the EE History from TDOC. These are NOT finalized numbers
    ImportEEHistory
    
    ' Process the hours to ensure they follow Duke's policy
    ' paraphrased~ "Everybody working >6hrs takes a lunch, unless a specific note is included."
    SendDukeLunchesEmail
    
    ' Go back to the Instructions worksheet
    Sheets("Instructions").Activate
    ActiveWorkbook.Save
End Sub

Sub AddressCleanup()
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
    
    ' Turn off screen updating
    Application.ScreenUpdating = False

    ' 1. Find the last row in the "PowerBI Details" sheet
    lastRow = wsPowerBI.Cells(wsPowerBI.Rows.Count, "A").End(xlUp).Row

    With wsAddressCleanup
        .Range("A2").Formula = "=TRUNC(B2&C2,0)"
        .Range("B2").Formula = "='PowerBI Details'!G2"
        .Range("C2").Formula = "='PowerBI Details'!I2"
        .Range("D2").Formula = "='PowerBI Details'!B2"
        .Range("E2").Formula = "=TRIM(IF(I2=0,IF(H2=0,IF(G2=0,I2,G2),H2),I2))"
        .Range("F2").Formula = "=PROPER(IF(K2=0,J2,K2))"
        .Range("G2").Formula = "='PowerBI Details'!AE2"
        .Range("H2").Formula = "='PowerBI Details'!AF2"
        .Range("I2").Formula = "='PowerBI Details'!AH2"
        .Range("J2").Formula = "='PowerBI Details'!C2"
        .Range("K2").Formula = "='PowerBI Details'!AI2"
        .Range("L2").Formula = "='PowerBI Details'!AW2"
        .Range("M2").Formula = "='PowerBI Details'!U2"

        ' 2. Fill in the rest of the formulas
        .Range("A2:M2").AutoFill Destination:=.Range("A2:M" & lastRow)

        ' 3. Paste the columns as values and add a filter to all the headers
        .Range("A2:L" & lastRow).Value = .Range("A2:L" & lastRow).Value
        .Range("A1:M1").AutoFilter
    End With

    ' 4. Search and replace within the range `wsAddressCleanup.Range("A2:L" & lastRow)` to correct the most common errors from our data (grammar, spelling, etc.)
    With wsAddressCleanup.Range("A2:J" & lastRow)
        ' Correct the state capitalization
        .Replace What:=" Fl", Replacement:=" FL", LookAt:=xlPart, MatchCase:=False
        .Replace What:=" Nc", Replacement:=" NC", LookAt:=xlPart, MatchCase:=False
        .Replace What:=" Sc", Replacement:=" SC", LookAt:=xlPart, MatchCase:=False
        ' Correct the streets named `School<something>`
        .Replace What:=" SCh", Replacement:=" Sch", LookAt:=xlPart, MatchCase:=False

        ' Remove blanks
        .Replace What:="(Blank); ", Replacement:="", LookAt:=xlPart, MatchCase:=False
        .Replace What:="(Blank) ", Replacement:="", LookAt:=xlPart, MatchCase:=False
        .Replace What:="(Blank)", Replacement:="", LookAt:=xlPart, MatchCase:=False

        ' Trim whitespace
        .Replace What:=" ;", Replacement:=";", LookAt:=xlPart, MatchCase:=False
        .Replace What:=" ,", Replacement:=",", LookAt:=xlPart, MatchCase:=False
        .Replace What:="    ", Replacement:=" ", LookAt:=xlPart, MatchCase:=False

        ' Misc.
        .Replace What:="WO# ", Replacement:="", LookAt:=xlPart, MatchCase:=False
        .Replace What:="WO#", Replacement:="", LookAt:=xlPart, MatchCase:=False
        .Replace What:="MAX", Replacement:="", LookAt:=xlPart, MatchCase:=False
        .Replace What:=":", Replacement:="", LookAt:=xlPart, MatchCase:=False
        .Replace What:=". ", Replacement:="", LookAt:=xlPart, MatchCase:=False
        .Replace What:=" .", Replacement:="", LookAt:=xlPart, MatchCase:=False
        .Replace What:=".", Replacement:="", LookAt:=xlPart, MatchCase:=False
        .Replace What:="�", Replacement:="-", LookAt:=xlPart, MatchCase:=False
        .Replace What:="�", Replacement:="-", LookAt:=xlPart, MatchCase:=False
        .Replace What:="�", Replacement:="-", LookAt:=xlPart, MatchCase:=False

        ' Additional replacements for ordinal numbers
        Dim i As Integer
        Dim j As Integer
        Dim suffix As String
        Dim ordinalSuffixes As Variant
        ordinalSuffixes = Array("St", "Nd", "Rd", "Th") ' Define the ordinal suffixes to replace

        For i = 1 To 9 ' Loop through the digits 1 to 9
            For j = LBound(ordinalSuffixes) To UBound(ordinalSuffixes)
                ' Replace the incorrect ordinal with the correct one
                wsAddressCleanup.Range("A2:L" & lastRow).Replace What:=CStr(i) & ordinalSuffixes(j), _
                    Replacement:=CStr(i) & LCase(ordinalSuffixes(j)), LookAt:=xlPart, MatchCase:=True
            Next j
        Next i

        ' Handle the special case for 10th separately to avoid replacing "1st" in "10th"
        wsAddressCleanup.Range("A2:L" & lastRow).Replace What:="0Th", _
                Replacement:="0th", LookAt:=xlPart, MatchCase:=True
    End With

    ' 5. Refresh the pivot tables
    RefreshPivotTablesOnSheet ("Address Check")
    
    ' 6. Go to the Address Cleanup worksheet
    Sheets("Address Cleanup").Activate
    ActiveWorkbook.Save
    ' Turn on screen updating
    Application.ScreenUpdating = True
End Sub

Sub clearOldData()
    ' FIRST! Save our data before deleting everything for next week
    'SaveWOsToDataModel
    
    ' Clear Order details
    Sheets("PowerBI Details").Select
    ' Unfilter table if there IS a filter, else do nothing.
    Range("A1").Select
    If Not Selection.AutoFilter Is Nothing Then
        Selection.AutoFilter
    End If
    
    ' Delete lower rows
    Rows("3:3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    ' Then the Raw Data in 1st row
    Range("E2:BA2").Select
    Selection.ClearContents
    Range("A1").Select
    
    ' Clear Employee Hours
    Sheets("EE History").Select
    ' Unfilter table if there IS a filter, else do nothing.
    Range("A1").Select
    If Not Selection.AutoFilter Is Nothing Then
        Selection.AutoFilter
    End If
    
    ' Delete lower rows
    Rows("3:3").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    ' Then the Raw Data in 1st row
    Range("A2:L2").Select
    Selection.ClearContents
    Range("A1").Select
    
    ' Clear WOs & Addresses
    Sheets("Address Cleanup").Select
    ' Unfilter table if there IS a filter, else do nothing.
    Range("A1").Select
    If Not Selection.AutoFilter Is Nothing Then
        Selection.AutoFilter
    End If
    
    ' Delete *ALL* data for this sheet
    ' a Macro provides this data rather than formulas
    Rows("2:2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    
    ' Go home and reset Billing Date
    Sheets("Instructions").Select
    Sheets("Instructions").Range("C3").Value = "mm.dd.yyyy"
    
    ' Save & close this workbook
    With ThisWorkbook
        .Save
        .Close False
    End With
End Sub
