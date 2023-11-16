Attribute VB_Name = "stage1"
Option Explicit
Public thisWeeksBillingDate As String
Public thisWeeksServerFolder As String
Public yearYYYY As String
Public yearYY As String
Public dateMMDD As String
Function ImportBillingDetails()

    ' Declare variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim deleteRange As Range

    ' Tell the user what will be done
    MsgBox "Select the `Billing Details` file that you have downloaded from PowerBI."
    
    ' Using the provided function to import Billing Details from PowerBI
    ImportDataToWorksheet "PowerBI Details", "G1"
    
    ' Turn off screen updating
    Application.ScreenUpdating = False

    ' Set ws
    Set ws = ThisWorkbook.Worksheets("PowerBI Details")

    ' Find the last row in the data and format necessary columns
    lastRow = FindLastRow("G", ws)
    FormatColumnAsDate ws, "K2:K" & lastRow
    FormatColumnAsDate ws, "L2:L" & lastRow
   
    ' Set formulas and fill in the rest of the columns
    AddFormulaAndCopyDown ws, "A", lastRow
    AddFormulaAndCopyDown ws, "B", lastRow
    AddFormulaAndCopyDown ws, "C", lastRow
    AddFormulaAndCopyDown ws, "D", lastRow
    AddFormulaAndCopyDown ws, "E", lastRow
    AddFormulaAndCopyDown ws, "F", lastRow
    
     ' Process Data for deletion and zero out certain cells
    deleteRange = DeleteRowsBasedOnCondition(ws, "Z", "Cancelled", lastRow)
    With ws
        For i = lastRow To 2 Step -1
            If .Cells(i, "AB").Value = "Billable Cancelled" And .Cells(i, "X").Value > 0 Then
                .Cells(i, "X").Value = 0
            End If
        Next i
    End With

    ' Delete the collected Rows
    If Not deleteRange Is Nothing Then deleteRange.Delete
    
    ' Find new lastRow after the deletion
    lastRow = FindLastRow("A", ws)

    ' Correct the formulas in column D after the deletions
    AddFormulaAndCopyDown ws, "D", lastRow
    
    ' Sort the ws table
    Dim sortRange As Range
    Set sortRange = ws.Range("A1:BC" & lastRow)
    SortTable ws, sortRange, lastRow
    
    ' Turn on screen updating
    Application.ScreenUpdating = True

    ' Indicate completion of the import
    MsgBox "Billing details have been successfully imported and processed."

End Function

Private Sub SortTable(ws As Worksheet, sortRange As Range, lastRow As Long)
    ' Use Excel's built-in sort feature to sort the inputs in the given worksheet
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=ws.Range("Z2:Z" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add2 Key:=ws.Range("AB2:AB" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add2 Key:=ws.Range("I2:I" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add2 Key:=ws.Range("K2:K" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange sortRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Sub ImportEEHistory()
    
    ' Declare variables
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' Set the target worksheet using custom function to add or get it
    AddOrGetWorksheet "EE History", ws
    
    ' Unhide the worksheet
    ws.Visible = xlSheetVisible
    
    ' Tell the user what will be done
    MsgBox "Select the `EE History` file that you have downloaded from TDOC."
    
    ' Import EE Pre-PDES-Closed Hours from TDOC
    ImportDataToWorksheet "EE History", "A1"
    
    ' Find the last row in the data using custom function
    lastRow = FindLastRow("A", ws)
    
    ' Format Columns using custom functions
    FormatColumnAsDate ws, "G2:G" & lastRow
    FormatColumn "U2:U" & lastRow, "0", "EE History"
    
    ' Fill in the rest of the formulas using custom function
    AddFormulaAndCopyDown ws, "M", lastRow
    AddFormulaAndCopyDown ws, "N", lastRow
    AddFormulaAndCopyDown ws, "O", lastRow
    AddFormulaAndCopyDown ws, "P", lastRow
    AddFormulaAndCopyDown ws, "Q", lastRow
    AddFormulaAndCopyDown ws, "R", lastRow
    AddFormulaAndCopyDown ws, "S", lastRow
    AddFormulaAndCopyDown ws, "T", lastRow
    AddFormulaAndCopyDown ws, "U", lastRow
    
    ' Sort Table using built-in Excel VBA methods
    Dim sortRange As Range
    Set sortRange = ws.Range("A1:U" & lastRow)
    sortRange.Sort Key1:=ws.Range("M1"), Order1:=xlAscending, Key2:=ws.Range("D1"), Order2:=xlAscending, Header:=xlYes
    
End Sub

Sub SubmitDukeLunches(action As String)

    Dim strbody As String
    Dim attachments As New Collection
    Dim attachmentPath As String

    ' Define the full path to the new Workbook, assuming it's saved already
    ' TODO: Set 'attachmentPath' to the actual path where the workbook is saved
    attachmentPath = "C:\Path\To\Your\Workbook.xlsx" ' Change this to the actual path

    ' Check if the file exists before trying to attach
    If FileExists(attachmentPath) Then
        attachments.Add attachmentPath
    Else
        MsgBox "Attachment file not found.", vbExclamation, "File Not Found"
        Exit Sub
    End If

    ' Prepare the email body
    strbody = "Good Morning," & "<br><br>" & _
              "Here is Duke's lunches that need adjusted. Thank you!" & "<br>"

    ' Send the email
    ' The 'action' variable should be either "Send" or any other value to display
    CreateAndSendEmail "paul.devey@flaggerforce.com", _
                        "Lunch Adjustments for Duke Energy", _
                        strbody, _
                        attachments, _
                        action

End Sub

Sub ProcessDukeLunches()

    ' Declare variable for the new workbook
    Dim newWb As Workbook
    
    ' Create a new workbook
    Set newWb = CreateNewWorkbook()

    ' Create "Under Reported" and "Over Reported" sheets in the new workbook
    AddOrGetWorksheet "Under Reported", newWb.Worksheets(1)
    AddOrGetWorksheet "Over Reported", newWb.Worksheets(2)
    
    ' Delete "Sheet1" (no longer needed since "Under Reported" and "Over Reported" sheets are created)
    Application.DisplayAlerts = False
    newWb.Sheets("Sheet1").Delete
    Application.DisplayAlerts = True

    ' Filter & copy under reported records
    AutoFilterColumn ThisWorkbook.Worksheets("EE History"), "R", "=Under", FindLastRow("A", ThisWorkbook.Worksheets("EE History"))
    CopyFormulasAndFilter newWb.Worksheets("Under Reported"), "A", 18, "=Under"

    ' Filter & copy over reported records
    AutoFilterColumn ThisWorkbook.Worksheets("EE History"), "R", "=Over", FindLastRow("A", ThisWorkbook.Worksheets("EE History"))
    CopyFormulasAndFilter newWb.Worksheets("Over Reported"), "A", 18, "=Over"

    ' Formatting Sheets Data
    Dim ws As Worksheet
    For Each ws In newWb.Worksheets
        FormatColumnAsDate ws, "G:G"
        FormatColumn "K:L", "[$-x-systime]h:mm:ss AM/PM", ws.Name
        ws.Cells.EntireColumn.AutoFit
    Next ws

    ' Save the new workbook as "Duke Lunches mm.dd"
    SaveWorkbookAs newWb, thisWeeksServerFolder & "Outputs\", "Duke Lunches " & dateMMDD & ".xlsx"

    ' Hide the worksheet "EE History"
    ThisWorkbook.Worksheets("EE History").Visible = xlSheetHidden
    
    ' Send the lunch file to IT
    SubmitDukeLunches newWb.FullName, "Send" ' Assuming SubmitDukeLunches is updated to accept parameters
    
End Sub

Sub AddressCleanup()
    
    Dim wsPowerBI As Worksheet
    Dim wsAddressCleanup As Worksheet
    Dim lastRow As Long

    ' Set worksheets
    Set wsPowerBI = ThisWorkbook.Worksheets("PowerBI Details")
    Set wsAddressCleanup = ThisWorkbook.Worksheets("Address Cleanup")
    
    ' Find the last row in the "PowerBI Details" sheet
    lastRow = FindLastRow("A", wsPowerBI)
    
    ' Add new formulas
    AddCalculationColumn "FFID&WorkDate", "=TRUNC(B2&C2,0)", lastRow
    AddCalculationColumn "FFID", "='PowerBI Details'!I2", lastRow
    AddCalculationColumn "WorkDate", "='PowerBI Details'!K2", lastRow
    AddCalculationColumn "Workable", "='PowerBI Details'!B2", lastRow
    AddCalculationColumn "Release", "=TRIM(IF(I2=0,IF(H2=0,IF(G2=0,I2,G2),H2),I2))", lastRow
    AddCalculationColumn "Address", "=PROPER(TRIM(IF(K2=0,J2,K2)))", lastRow
    AddCalculationColumn "timeOfOrderWO", "='PowerBI Details'!AG2", lastRow
    AddCalculationColumn "timeOfOrderDailyWO", "='PowerBI Details'!AH2", lastRow
    AddCalculationColumn "jobLeadWO", "='PowerBI Details'!AJ2", lastRow
    AddCalculationColumn "timeOfOrderAddress", "='PowerBI Details'!C2", lastRow
    AddCalculationColumn "jobLeadAddress", "='PowerBI Details'!AK2", lastRow
    AddCalculationColumn "State", "='PowerBI Details'!AY2", lastRow
    AddCalculationColumn "flaggersOrdered", "='PowerBI Details'!W2", lastRow

    ' Copy the columns as values using the .Value = .Value technique
    With wsAddressCleanup.Range("A2:M" & lastRow)
        .Value = .Value
    End With

    ' Apply an AutoFilter to all the headers
    wsAddressCleanup.Range("A1").CurrentRegion.AutoFilter
    
    ' Refresh the pivot table in "Address Check"
    With ThisWorkbook.Worksheets("Address Check")
        .PivotTables("AddressCheck").RefreshTable
    End With
    
End Sub

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


