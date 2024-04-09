Attribute VB_Name = "PdesClosed"
Option Explicit
Public thisWeeksBillingDate As String
Public thisWeeksServerFolder As String
Public yearYYYY As String
Public yearYY As String
Public dateMMDD As String

Function ImportBillingDetails()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dataRange As Range
    Dim deleteRange As Range

    ' Tell the user what will be done
    MsgBox "Sorry, but we must re-import the `Billing Details` file so that the Rec still works."
    ' Import Billing Details from PowerBI
    ImportDatatoWorksheet "PowerBI Details", "E1", "Excel"

    ' Set ws
    Sheets("PowerBI Details").Activate
    Set ws = ThisWorkbook.Worksheets("PowerBI Details")

    ' Turn off screen updating
    Application.ScreenUpdating = False

    ' Find the last row in the data and format necessary columns
    With ws
        lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
        .Range("I2:I" & lastRow).NumberFormat = "m/d/yyyy"
        .Range("J2:J" & lastRow).NumberFormat = "h:mm AM/PM"
    End With

    Set dataRange = ws.Range("A1:BA" & lastRow)

    ' Set Formulas
    ws.Range("A2").FormulaR1C1 = "=TRUNC(RC[8] & RC[10], 0)"
    ws.Range("B2").FormulaR1C1 = "=IF(RC[24]="""", ""yes"", IF(RC[26]=""Billable Cancelled"", IF(OR(RC[27]=""Attendance"", RC[27]=""Resources""), ""no"", ""CXL""), ""no""))"
    ws.Range("C2").FormulaR1C1 = "=CONCAT(RC[46], "" "", RC[47], "" "", RC[48], "" "", RC[49])"
    ws.Range("D2").FormulaR1C1 = "=RC[-3]=R[-1]C[-3]"

    ' Fill in the rest of the formulas
    Sheets("PowerBI Details").Activate
    Range("A2:D2").Select
    Selection.AutoFill Destination:=Range("A2:D" & lastRow)

    ' Process Data
    ' This is a loop that checks if col Z = "Cancelled" and col AB  = "" then delete the row
    ' If col Z = "Cancelled" and col AB is "Billable CXL" and col X (equipment ordered) is >0 set it to 0
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

        If ws.Cells(i, "X").Value = "Cancelled" And ws.Cells(i, "Z").Value = "Billable Cancelled" And ws.Cells(i, "V").Value > 0 Then
            ws.Cells(i, "V").Value = 0
        End If
    Next i

    ' Delete the collected Rows
    If Not deleteRange Is Nothing Then deleteRange.Delete

    ' Find the NEW last row
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Range("D2").Formula = "=A2=A1"
    Range("D2").Select
    Selection.AutoFill Destination:=Range("D2:D" & lastRow)
    
    ' Sort Table
    Dim sortRange As Range
    Set sortRange = ws.Range("A1:BA" & lastRow)

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add2 Key:=ws.Range("X2:X" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add2 Key:=ws.Range("Z2:Z" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add2 Key:=ws.Range("A2:A" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range("A1:BA" & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    ' Apply filter on column B for 'CXL'
    dataRange.AutoFilter Field:=2, Criteria1:="CXL"

    ' Turn on screen updating
    Application.ScreenUpdating = True

End Function

Sub ImportRawHours()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Tell the user what will be done
    MsgBox "Select the `Raw Data` file that you have downloaded from PDES."
    
    ' Turn off screen updating
    Application.ScreenUpdating = False
    
    ' Import Billing Details from PowerBI
    ImportDatatoWorksheet "Raw Hours", "O1", "CSV"
    
    ' Set ws
    Sheets("Raw Hours").Activate ' Go to the actual sheet to make sure it executes correctly
    Set ws = ThisWorkbook.Worksheets("Raw Hours")

    With ws
        ' Find the last row in the data
        lastRow = .Cells(.Rows.Count, "O").End(xlUp).Row
        
        ' Format Columns
        .Range("R2:R" & lastRow).NumberFormat = "m/d/yyyy"
        .Range("R2:R" & lastRow).TextToColumns Destination:=Range("R2"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
            :=Array(Array(1, 3), Array(2, 9), Array(3, 9)), TrailingMinusNumbers:=True
        
        ' Loop through the last row to row 2
        ' Process any rows that are Billable CXL
        For i = lastRow To 2 Step -1
            If .Cells(i, "E").Value = "no" Then
                .Rows(i).Delete
            ElseIf .Cells(i, "E").Value = "CXL" Then
                .Cells(i, "p").Value = 0
                .Cells(i, "T").Value = .Cells(i, "S").Value + 1 / 24 ' Equals 1 hour per Excel data structure for time keeping
                .Cells(i, "U").Value = 1
            End If
        Next i
        
        ' Find the NEW last row
        lastRow = .Cells(.Rows.Count, "O").End(xlUp).Row
        
        ' Fix the formulas that were ruined in the row deletions earlier
        .Range("J2:J" & lastRow).FormulaR1C1 = "=IF(RC[-5]=""yes"", IF(RC[6]=R[-1]C[6], (RC[11]+RC[12])+R[-1]C, RC[11]+RC[12]), 1)"
        .Range("K2:K" & lastRow).FormulaR1C1 = "=IF(RC[-6]=""yes"", MIN(IF(RC[5]=R[-1]C[5], RC[-3]+R[-1]C[-1], RC[-3]), 40), 1)"
        
        ' Filter Col E for "CXL"
        .Range("E1:E" & lastRow).AutoFilter Field:=5, Criteria1:="CXL"
    End With
    
    ' Turn on screen updating
    Application.ScreenUpdating = True
    ActiveWorkbook.Save
    ' Inform the user that the import is complete
    MsgBox "Raw Hours imported." & vbNewLine & "Last Row: " & lastRow
End Sub

Sub ImportRawEquip()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Tell the user what will be done
    MsgBox "Select the `Raw Data - Equipment` file that you have downloaded from PDES."
    
    ' Turn off screen updating
    Application.ScreenUpdating = False
    
    ' Import Billing Details from PowerBI
    ImportDatatoWorksheet "Raw Equipment", "H1", "CSV"

    ' Set ws
    Sheets("Raw Equipment").Activate
    Set ws = ThisWorkbook.Worksheets("Raw Equipment")

    With ws
        ' Find the last row in the data
        lastRow = .Cells(.Rows.Count, "H").End(xlUp).Row
        
        ' Format Columns
        .Range("L2:L" & lastRow).NumberFormat = "m/d/yyyy"
        
        ' Loop through the rows and delete any where Col D = "no"
        For i = lastRow To 2 Step -1
            If .Cells(i, "D").Value = "no" Then
                .Rows(i).Delete
            End If
        Next i
    End With
    
    ' Turn on screen updating
    Application.ScreenUpdating = True
    ActiveWorkbook.Save
    
    ' Inform the user that the import is complete
    MsgBox "Raw Equipment imported." & vbNewLine & "Last Row: " & lastRow
End Sub

Sub EEotPivots()
    ' Refresh pivot tables on the "EE OT Breakdown" sheet
    RefreshPivotTablesOnSheet ("EE OT Breakdown")
End Sub

Sub refreshDraftImport()
    Dim lastRow As Long
    Dim wsHours As Worksheet
    
    Set wsHours = ThisWorkbook.Worksheets("Raw Hours")
    ' Find the last row
    lastRow = wsHours.Cells(wsHours.Rows.Count, "O").End(xlUp).Row
    
    ' Fix the formulas in case there are reference errors from row deletion
    With wsHours
        ' Running Total & Running Daily
        .Range("J2:J" & lastRow).FormulaR1C1 = "=IF(RC[-5]=""yes"", IF(RC[6]=R[-1]C[6], (RC[11]+RC[12])+R[-1]C, RC[11]+RC[12]), 1)"
        .Range("K2:K" & lastRow).FormulaR1C1 = "=IF(RC[-6]=""yes"", MIN(IF(RC[5]=R[-1]C[5], RC[-3]+R[-1]C[-1], RC[-3]), 40), 1)"
    End With
    
    ' Refresh the query table on the "Draft_Import" sheet
    RefreshQueriesOnSheet ("Draft_Import")
    
    Sheets("Draft_Import").Range("S2:S10000").FormulaR1C1 = _
        "=IFERROR(RC[-10]*XLOOKUP(RC[-15],'RATE SHEET'!C[-16],'RATE SHEET'!C[-12],""ERR""),0)"
    
    MsgBox "Draft_Import refreshed!"
    ActiveWorkbook.Save
End Sub

Sub SplitImport()
    Dim sourceWorksheet As Worksheet
    Dim targetWorkbook As Workbook
    Dim targetWorksheet As Worksheet
    Dim currentRow As Long
    Dim startRow As Long
    Dim rowCount As Long
    Dim fileIndex As Integer
    Dim fileName As String
    Dim outputPath As String
    Dim lastFullReleaseRow As Long
    
    ' Initialize variables
    fileIndex = 1
    startRow = 2
    rowCount = 0
    lastFullReleaseRow = 1 ' Last row of a full set before a release number changes
    
    ActiveWorkbook.Save
    
    ' Set the input and output worksheets
    Set sourceWorksheet = ThisWorkbook.Sheets("Draft_Import")
    
    ' Set the output path
    outputPath = ThisWorkbook.Path & "\Outputs\"
    
    ' Determine the last row containing data in "Draft_Import"
    Dim lastRow As Long
    lastRow = sourceWorksheet.Cells(sourceWorksheet.Rows.Count, "A").End(xlUp).Row

    ' Copy headers from "Draft_Import" to "sourceWorksheet"
    Dim headerRange As Range
    Set headerRange = sourceWorksheet.Range("A1:R1")
    
    ' Initialize a variable to track the number of rows copied to the new workbook
    Dim copiedRowCount As Long
    copiedRowCount = 0

    ' Create a new workbook with the header
    Set targetWorkbook = Application.Workbooks.Add
    Set targetWorksheet = targetWorkbook.Sheets(1)
    headerRange.Copy Destination:=targetWorksheet.Rows(1)

    ' Loop through all rows
    For currentRow = startRow To lastRow
        rowCount = rowCount + 1

        ' Check if release number has changed and update the lastFullReleaseRow
        If sourceWorksheet.Cells(currentRow, "E").Value <> sourceWorksheet.Cells(currentRow + 1, "E").Value Or currentRow = lastRow Then
            lastFullReleaseRow = currentRow

            ' Define data range to copy from "Draft_Import"
            Dim dataRange As Range
            Set dataRange = sourceWorksheet.Range(sourceWorksheet.Rows(startRow), sourceWorksheet.Rows(lastFullReleaseRow)).Columns("A:R")
            dataRange.Copy Destination:=targetWorksheet.Rows(copiedRowCount + 2)

            ' Increment copiedRowCount
            copiedRowCount = copiedRowCount + rowCount

            ' If copiedRowCount is greater than or equal to 900 or currentRow is the lastRow, save and close the workbook, and create a new one
            If copiedRowCount > 750 Or currentRow = lastRow Then
                ' AutoFit the columns
                targetWorksheet.Cells.EntireColumn.AutoFit

                ' Save the new workbook
                fileName = outputPath & "Duke Import pt" & fileIndex & ".xlsx"
                targetWorkbook.SaveAs fileName:=fileName
                targetWorkbook.Close SaveChanges:=False

                ' Create a new workbook with the header
                Set targetWorkbook = Application.Workbooks.Add
                Set targetWorksheet = targetWorkbook.Sheets(1)
                headerRange.Copy Destination:=targetWorksheet.Rows(1)

                ' Increment fileIndex and reset copiedRowCount for the next set of data
                fileIndex = fileIndex + 1
                copiedRowCount = 0
            End If

            ' Reset startRow for the next set of data
            startRow = lastFullReleaseRow + 1
            rowCount = 0
        End If
    Next currentRow
    
    ' Close the last workbook without saving
    targetWorkbook.Close SaveChanges:=False
    
    MsgBox "Split data into " & (fileIndex - 1) & " import files."
End Sub

Sub CatchUpToMonday()
    Dim wb As Workbook
    Dim wsInstructions As Worksheet
    Dim wsPowerBI As Worksheet
    Dim lastRow As Long
    
    ' Set the Globals
    thisWeeksBillingDate = InputBox("Enter the billing date in mm.dd.yyyy format: ", "Billing Date", Format(Now(), "mm.dd.yyyy"))
    ' Extract the year and date parts from the billing date
    yearYYYY = Right(thisWeeksBillingDate, 4)
    yearYY = Right(thisWeeksBillingDate, 2)
    dateMMDD = Left(thisWeeksBillingDate, 5)
    
    ' Set the local variables
    Set wb = ThisWorkbook
    Set wsInstructions = wb.Worksheets("Instructions")
    Set wsPowerBI = wb.Worksheets("PowerBI Details")
    
    ' Save the folder and file
    Sheets("Instructions").Range("C3").Value = thisWeeksBillingDate
    ' Create the server folder path
    thisWeeksServerFolder = Sheets("Instructions").Range("C5").Value
    Debug.Print thisWeeksServerFolder

    ' Grab the Billing Details (again)
    ImportBillingDetails
End Sub

Function SetFormulas()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim dataRange As Range
    Dim deleteRange As Range
    
    ' Set ws
    Sheets("PowerBI Details").Activate
    Set ws = ThisWorkbook.Worksheets("PowerBI Details")
    
    ' Turn off screen updating
    Application.ScreenUpdating = False
    
    ' Find the last row in the data and format necessary columns
    With ws
        lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
        .Range("I2:I" & lastRow).NumberFormat = "m/d/yyyy"
        .Range("J2:J" & lastRow).NumberFormat = "h:mm AM/PM"
    End With
    
    Set dataRange = ws.Range("A1:BA" & lastRow)
    
    ' Set Formulas
    ws.Range("A2").FormulaR1C1 = "=TRUNC(RC[8] & RC[10], 0)"
    ws.Range("B2").FormulaR1C1 = "=IF(RC[24]="""", ""yes"", IF(RC[26]=""Billable Cancelled"", IF(OR(RC[27]=""Attendance"", RC[27]=""Resources""), ""no"", ""CXL""), ""no""))"
    ws.Range("C2").FormulaR1C1 = "=CONCAT(RC[46], "" "", RC[47], "" "", RC[48], "" "", RC[49])"
    ws.Range("D2").FormulaR1C1 = "=RC[-3]=R[-1]C[-3]"
    
    ' Turn off screen updating
    Application.ScreenUpdating = False
End Function
