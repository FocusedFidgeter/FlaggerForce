Attribute VB_Name = "FileIO"
Function ImportDatatoWorksheet(targetSheetName As String, targetRange As String, fileType As String)

    ' Declare variables
    Dim sourceWorkbook As Workbook
    Dim sourceWorksheet As Worksheet
    Dim destWorkbook As Workbook
    Dim destWorksheet As Worksheet
    Dim fileName As Variant
    Dim fileFilter As String

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

    ' Set the file filter based on fileType parameter
    If fileType = "Excel" Then
        fileFilter = "Excel Files (*.xls*), *.xls*, All Files (*.*), *.*"
    ElseIf fileType = "CSV" Then
        fileFilter = "CSV Files (*.csv), *.csv, All Files (*.*), *.*"
    Else
        MsgBox "Invalid file type. Please select 'Excel' or 'CSV'."
        Exit Function
    End If

    ' Prompt the user to select a file based on the fileFilter
    fileName = Application.GetOpenFilename(fileFilter:=fileFilter, Title:="Select a file to import")
    
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

Function CombineTimesheets()
    Dim pythonExe As String
    Dim scriptPath As String
    Dim command As String
    
    ' Refresh the data we will use in the python script
    Sheets("TimesheetCombiner").Activate
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
    
    ' Go back to the instructions page
    Sheets("Instructions").Activate
    ActiveWorkbook.Save

    ' Set the path to the Python executable
    pythonExe = "python3"

    ' Set the path to the Python script
    scriptPath = "\\hum-vmqb-01\Billing\Python Scripts\timesheet_combiner_duke.py"

    ' Create the command to run the Python script
    command = pythonExe & " " & Chr(34) & scriptPath & Chr(34)

    ' Execute the command
    Shell command, vbNormalFocus
End Function

Function RefreshPivotTablesOnSheet(sheetName As String)
    Dim ws As Worksheet
    Dim pt As PivotTable
    
    ' Check if the sheet exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    
    ' If the worksheet object is nothing, the sheet does not exist
    If ws Is Nothing Then
        MsgBox "Sheet '" & sheetName & "' does not exist.", vbExclamation
        Exit Function
    End If
    
    ' Loop through all pivot tables in the sheet and refresh each
    For Each pt In ws.PivotTables
        pt.RefreshTable
    Next pt
    
    'MsgBox "All pivot tables on '" & sheetName & "' have been refreshed.", vbInformation
End Function

Function RefreshQueriesOnSheet(sheetName As String)
    Dim ws As Worksheet
    Dim qt As QueryTable
    Dim lo As ListObject
    
    ' Attempt to set the worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    
    ' Check if the worksheet exists
    If ws Is Nothing Then
        MsgBox "Sheet '" & sheetName & "' not found!", vbCritical
        Exit Function
    End If
    
    ' Refresh QueryTables
    For Each qt In ws.QueryTables
        qt.Refresh BackgroundQuery:=False
    Next qt
    
    ' Refresh ListObjects (Tables)
    For Each lo In ws.ListObjects
        ' Check if the ListObject has a query table associated with it
        If Not lo.QueryTable Is Nothing Then
            lo.QueryTable.Refresh BackgroundQuery:=False
        ' Check if the ListObject is connected to an external data source
        ElseIf lo.SourceType = xlSrcQuery Then
            lo.Refresh
        End If
    Next lo
    
    'MsgBox "Queries on sheet '" & sheetName & "' have been refreshed.", vbInformation
End Function

