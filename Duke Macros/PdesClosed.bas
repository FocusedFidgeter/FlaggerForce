Attribute VB_Name = "PdesClosed"
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

Function ImportRawHours()
    ' Declare variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Tell the user what will be done
    MsgBox "Select the `Raw Data` file that you have downloaded from PDES."
    
    ' Import Billing Details from PowerBI
    ImportDataToWorksheet "Raw Hours", "O1"
    
    ' Turn off screen updating
    Application.ScreenUpdating = False

    ' Set ws
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
                .Cells(i, "P").Value = 0
                .Cells(i, "Q").Value = "Billable CXL"
                .Cells(i, "T").Value = .Cells(i, "S").Value + 1 / 24
                .Cells(i, "U").Value = 1
            End If
        Next i
        
        ' Find the new last row
        lastRow = .Cells(.Rows.Count, "O").End(xlUp).Row
        
        ' Copy the Formulas in row 2 and paste them down the rest of the table
        .Range("J2:K2").Copy
        .Range("J3:K" & lastRow).PasteSpecial Paste:=xlPasteFormulas
        
        ' Sort the table by the following Columns: P dsc, and R asc
        
        
        ' Filter Col E for "CXL"
        .Range("E1:E" & lastRow).AutoFilter Field:=5, Criteria1:="CXL"
    End With
    
    ' Turn on screen updating
    Application.ScreenUpdating = True
    ' Inform the user that the import is complete
    MsgBox "Raw Hours imported successfully!"
End Function

Function ImportRawEquip()
' TODO: Implement the empty comments
    ' Declare variables
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    
    ' Tell the user what will be done
    MsgBox "Select the `Raw Data - Equipment` file that you have downloaded from PDES."
    
    ' Import Billing Details from PowerBI
    ImportDataToWorksheet "Raw Equipment", "H1"
    
    ' Turn off screen updating
    Application.ScreenUpdating = False

    ' Set ws
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
    
    ' Sort By Col M asc, Col L asc
        
    ' Filter Col B for "Research Required"
    ws.Range("B1:B" & lastRow).AutoFilter Field:=2, Criteria1:="Research Required"
    
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

    Dim ws As Worksheet
    Dim newWorkbook As Workbook
    Dim newWorksheet As Worksheet
    Dim currentRow As Long
    Dim startRow As Long
    Dim totalRows As Long
    Dim fileIndex As Integer
    Dim fileName As String
    Dim outputPath As String
    Dim lastFullReleaseRow As Long
    Dim lastRow As Long

    Sheets("IMPORT").Visible = True

    ' Refresh Draft_Import 1 more time and copy the full table
    With Sheets("Draft_Import")
        .Range("A1").ListObject.QueryTable.Refresh BackgroundQuery:=False
        lastRow = .Cells(Rows.Count, "A").End(xlUp).Row
        .Range("A1", .Cells(lastRow, .Cells(1, Columns.Count).End(xlToLeft).Column)).Copy
    End With

    ' Paste as values into "IMPORT"
    With Sheets("IMPORT")
        .Range("E1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        .Range("A2:D2").AutoFill Destination:=.Range("A2:D" & lastRow)
        .Range("S2").AutoFill Destination:=.Range("S2:S" & lastRow)
    End With

    ' Save before something goes wrong
    ActiveWorkbook.Save

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("IMPORT")
    
    ' Initialize variables
    currentRow = 2 ' Assuming row 1 has headers
    startRow = 2
    totalRows = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row
    fileIndex = 1
    lastFullReleaseRow = 2
    
    ' Set the output path
    outputPath = ThisWorkbook.Path & "\Outputs\"
    
    ' Loop through all rows
    While currentRow <= totalRows
        
        ' Check if we have reached the row limit
        If currentRow - startRow + 1 >= 900 Then
            
            ' Check if the next "Release" code is the same as the current one
            If ws.Cells(currentRow + 1, "E").Value = ws.Cells(currentRow, "E").Value Then
                ' If it's the same, continue to the next row
                currentRow = currentRow + 1
            Else
                ' If it's not the same, split the file at the last full release number
                ' Create a new workbook
                Set newWorkbook = Application.Workbooks.Add
                Set newWorksheet = newWorkbook.Sheets(1)
                
                ' Copy the headers and the rows into the new workbook
                ws.Rows(1).Copy Destination:=newWorksheet.Rows(1)
                ws.Range(ws.Rows(startRow), ws.Rows(lastFullReleaseRow)).Copy Destination:=newWorksheet.Rows(2)
                
                ' Delete the import price column
                Columns("S:S").Select
                Selection.Delete Shift:=xlToLeft
                
                ' AutoFit the columns
                newWorksheet.Cells.EntireColumn.AutoFit
                
                ' Save the new workbook
                fileName = outputPath & "Duke Import pt" & fileIndex & ".xlsx"
                newWorkbook.SaveAs fileName:=fileName
                newWorkbook.Close SaveChanges:=False
                
                ' Update the start row and the file index
                startRow = lastFullReleaseRow + 1
                fileIndex = fileIndex + 1
            End If
        End If
        
        ' Check if the "Release" code has changed
        If ws.Cells(currentRow + 1, "E").Value <> ws.Cells(currentRow, "E").Value Then
            ' If it has changed, update the last full release row
            lastFullReleaseRow = currentRow
        End If
        
        ' Move to the next row
        currentRow = currentRow + 1
    Wend
    
    ' Check if there are remaining rows less than 900 and the file reaches the end
    If currentRow - startRow + 1 < 900 And currentRow > totalRows Then
        ' Create a new workbook
        Set newWorkbook = Application.Workbooks.Add
        Set newWorksheet = newWorkbook.Sheets(1)
        
        ' Copy the headers and the rows into the new workbook
        ws.Rows(1).Copy Destination:=newWorksheet.Rows(1)
        ws.Range(ws.Rows(startRow), ws.Rows(currentRow)).Copy Destination:=newWorksheet.Rows(2)
        
        ' AutoFit the columns
        newWorksheet.Cells.EntireColumn.AutoFit
        
        ' Save the new workbook
        fileName = outputPath & "Duke Import pt" & fileIndex & ".xlsx"
        newWorkbook.SaveAs fileName:=fileName
        newWorkbook.Close SaveChanges:=False
    End If

End Sub

