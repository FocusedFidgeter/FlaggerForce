Attribute VB_Name = "stage2"
Option Explicit
Public thisWeeksBillingDate As String
Public thisWeeksServerFolder As String
Public yearYYYY As String
Public yearYY As String
Public dateMMDD As String

Function ImportRawHours()
    ' Declare variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim sortRange As Range
    Dim i As Long
    
    ' Tell the user what will be done
    MsgBox "Select the `Raw Data` file that you have downloaded from PDES."
    
    ' Import Billing Details from PowerBI
    ImportDataToWorksheet "Raw Hours", "O1"
    
    ' Turn off screen updating
    Application.ScreenUpdating = False

    ' Set ws
    Set ws = ThisWorkbook.Worksheets("Raw Hours")

    ' Format Columns using the FormatColumnAsDate helper function
    Call FormatColumnAsDate(ws, "R2:R" & FindLastRow("O", ws))
    
    ' Process any rows that are Billable CXL using the For loop and DeleteRowsBasedOnCondition helper function
    lastRow = FindLastRow("O", ws) ' Find the last row using util function
    For i = lastRow To 2 Step -1
        If ws.Cells(i, "E").Value = "no" Then
            ws.Rows(i).Delete
        ElseIf ws.Cells(i, "E").Value = "CXL" Then
            ws.Cells(i, "P").Value = 0
            ws.Cells(i, "Q").Value = "Billable CXL"
            ws.Cells(i, "T").Value = ws.Cells(i, "S").Value + 1 / 24
            ws.Cells(i, "U").Value = 1
        End If
    Next i
    
    ' Copy the Formulas in row 2 and paste them down the rest of the table using the AddFormulaAndCopyDown helper function
    Call AddFormulaAndCopyDown(ws, "J", FindLastRow("O", ws))
    Call AddFormulaAndCopyDown(ws, "K", FindLastRow("O", ws))
    
    ' Sort the table by the following Columns: P desc, and R asc
    lastRow = FindLastRow("O", ws)
    Set sortRange = ws.Range("A1:U" & lastRow)
    With sortRange
        .AutoFilter
        .Sort.SortFields.Clear
        .Sort.SortFields.Add2 Key:=ws.Range("P1"), Order:=xlDescending
        .Sort.SortFields.Add2 Key:=ws.Range("R1"), Order:=xlAscending
        With .Sort
            .SetRange sortRange
            .Header = xlYes
            .Apply
        End With
    End With
    
    ' Filter Col E for "CXL" using the AutoFilterColumn helper function
    Call AutoFilterColumn(ws, "E", "CXL", FindLastRow("E", ws))
    
    ' Turn on screen updating
    Application.ScreenUpdating = True
    
    ' Inform the user that the import is complete
    MsgBox "Raw Hours imported successfully!"
End Function

Function ImportRawEquip()
    Dim ws As Worksheet
    Dim lastRow As Long
    
    ' Tell the user what will be done
    MsgBox "Select the `Raw Data - Equipment` file that you have downloaded from PDES."
    
    ' Import Billing Details from PowerBI
    ImportDataToWorksheet "Raw Equipment", "H1"
    
    ' Turn off screen updating
    Application.ScreenUpdating = False

    ' Set ws
    Set ws = ThisWorkbook.Worksheets("Raw Equipment")

    ' Find the last row in the data
    lastRow = FindLastRow("H", ws)
    
    ' Format Column L as date
    FormatColumnAsDate ws, "L2:L" & lastRow
    
    ' Delete rows in Column D where the value is "no"
    DeleteRowsBasedOnCondition ws, "D", "no", lastRow
        
    ' Sort By Col M asc, Col L asc
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("M2:M" & lastRow), Order:=xlAscending
        .SortFields.Add Key:=ws.Range("L2:L" & lastRow), Order:=xlAscending
        .SetRange ws.Range("A1:Z" & lastRow)
        .Header = xlYes
        .Apply
    End With
        
    ' Filter Col B for "Research Required"
    AutoFilterColumn ws, "B", "Research Required", lastRow
    
    ' Turn on screen updating
    Application.ScreenUpdating = True

    ' Inform the user that the import is complete
    MsgBox "Raw Equipment imported successfully!"
End Function


Sub ProcessPdes()
    ' Grab the Payroll-Approved Employee Hours
    ImportRawHours
    ' Grab the finalized Equipment
    ImportRawEquip
End Sub

Sub AddTruckHours()
    ' Define workbook and worksheets
    Dim wb As Workbook
    Dim wsRaw As Worksheet
    Dim wsTruck As Worksheet
    Dim wsDraft As Worksheet
    Dim wsOT As Worksheet
    Dim lastRowRaw As Long
    Dim newRowDraft As Long
    Dim tbl As ListObject
    
    Set wb = ThisWorkbook
    Set wsRaw = wb.Worksheets("Raw Hours")
    Set wsTruck = wb.Worksheets("Truck Hours")
    Set wsDraft = wb.Worksheets("Draft_Import")
    Set wsOT = wb.Worksheets("EE OT Breakdown")

    ' Determine last rows of each worksheet
    lastRowRaw = wsRaw.Cells(Rows.Count, "A").End(xlUp).Row
    newRowDraft = wsDraft.Cells(Rows.Count, "A").End(xlUp).Row + 1
    
    ' Copy the formulas
    wsTruck.Range("A2:Q2").Copy wsTruck.Range("A2:Q" & lastRowRaw)
    
    ' Set the table object
    Set tbl = wsTruck.ListObjects("Truck_Hours")
    
    ' Resize the table
    tbl.Resize wsTruck.Range("A1:Q" & lastRowRaw)
    
    ' Copy and Paste-A-Values range("B2:Q" & lastRowRaw)
    wsTruck.Range("B2:Q" & lastRowRaw).Copy
    wsTruck.Range("B2:Q" & lastRowRaw).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    ' Hide the sheets "Raw Hours", "Raw Equipment", and "Truck Hours"
    wsRaw.Visible = False
    'wsTruck.Visible = False
    wb.Sheets("Raw Equipment").Visible = False

    ' Select the Auditor Sheet
    wsOT.Visible = True
    'wsOT.Range("A1").Select

End Sub

Sub SplitImport()
    Const MaxRowsPerFile As Long = 900
    Dim ws As Worksheet
    Dim newWorkbook As Workbook
    Dim newWorksheet As Worksheet
    Dim currentRow As Long
    Dim startRow As Long
    Dim lastRow As Long
    Dim fileIndex As Integer
    Dim outputPath As String
    Dim fileName As String

    ' Ensure IMPORT sheet is visible and refresh Draft_Import
    Sheets("IMPORT").Visible = True
    With Sheets("Draft_Import")
        .Range("A1").ListObject.QueryTable.Refresh BackgroundQuery:=False
        lastRow = .Cells(.Rows.Count, "A").End(xlUp).row
        .Range("A1", .Cells(lastRow, .Cells(1, Columns.Count).End(xlToLeft).Column)).Copy
    End With

    ' Paste as values into "IMPORT"
    With Sheets("IMPORT")
        .Range("E1").PasteSpecial Paste:=xlPasteValues
        .Range("A2:D" & lastRow).Value = .Range("A2:D2").Value
        .Range("S2:S" & lastRow).Value = .Range("S2").Value
    End With

    ' Save the workbook before proceeding
    ThisWorkbook.Save

    ' Refer to the IMPORT worksheet
    Set ws = ThisWorkbook.Sheets("IMPORT")
    
    ' Initialize
    startRow = 2 ' Assuming headers are in row 1
    currentRow = 2
    fileIndex = 1

    ' Create Outputs folder if it doesn't exist
    outputPath = ThisWorkbook.Path & "\Outputs\"
    CreateFolderIfNotExists outputPath

    ' Start splitting process
    While currentRow <= lastRow
        If currentRow - startRow >= MaxRowsPerFile Or ws.Cells(currentRow + 1, "E").Value <> ws.Cells(currentRow, "E").Value Then
            ' Create a new workbook
            Set newWorkbook = CreateNewWorkbook()
            Set newWorksheet = newWorkbook.Worksheets(1)
            ws.Rows(1).Copy Destination:=newWorksheet.Rows(1)
            ws.Range(ws.Rows(startRow), ws.Rows(currentRow)).Copy
            
            ' Paste as values to remove links to the original workbook
            With newWorksheet.Range(newWorksheet.Rows(2), newWorksheet.Rows(currentRow - startRow + 2))
                .PasteSpecial Paste:=xlPasteValues
                .PasteSpecial Paste:=xlPasteFormats
            End With
            Application.CutCopyMode = False

            ' Save the new workbook
            fileName = "Duke Import pt" & fileIndex & ".xlsx"
            SaveWorkbookAs newWorkbook, outputPath, fileName
            
            ' Increment fileIndex and update startRow for next batch
            fileIndex = fileIndex + 1
            startRow = currentRow + 1
        End If
        
        ' Increment row counter
        currentRow = currentRow + 1
    Wend

End Sub

