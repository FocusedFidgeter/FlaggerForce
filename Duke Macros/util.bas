Function FileExists(filePath As String) As Boolean
    ' Check if a file exists at the specified file path
    FileExists = (Dir(filePath) <> "")
End Function

Sub CreateFolderIfNotExists(folderPath As String)
    ' Create a folder if it doesn't exist at the specified folder path
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
End Sub

Sub MoveFile(sourceFilePath As String, destFilePath As String)
    ' Move a file from the source file path to the destination file path
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If FileExists(sourceFilePath) Then
        fso.MoveFile Source:=sourceFilePath, Destination:=destFilePath
    End If
End Sub

Sub CopyFile(sourceFilePath As String, destFilePath As String)
    ' Copy a file from the source file path to the destination file path
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If FileExists(sourceFilePath) Then
        fso.CopyFile Source:=sourceFilePath, Destination:=destFilePath
    End If
End Sub

Function CreateNewWorkbook() As Workbook
    ' Create a new workbook and return a reference to it
    Set CreateNewWorkbook = Application.Workbooks.Add
End Function

Function IsWorkBookOpen(Name As String) As Boolean
    ' Check if a workbook with a specific name is open
    Dim xWb As Workbook
    On Error Resume Next
    Set xWb = Application.Workbooks.Item(Name)
    IsWorkBookOpen = (Not xWb Is Nothing)
End Function

Function WorksheetExists(sheetName As String, wb As Workbook) As Boolean
    ' Check if a worksheet with a specific name exists in a workbook
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = wb.Sheets(sheetName)
    WorksheetExists = Not sheet Is Nothing
    On Error GoTo 0
End Function

Sub AddOrGetWorksheet(sheetName As String, ByRef outSheet As Worksheet)
    ' Add a new worksheet with the given name or get the existing worksheet
    With ThisWorkbook
        If Not WorksheetExists(sheetName, ThisWorkbook) Then
            Set outSheet = .Sheets.Add(After:=.Sheets(.Sheets.Count))
            outSheet.Name = sheetName
        Else
            Set outSheet = .Sheets(sheetName)
        End If
    End With
    outSheet.Visible = xlSheetVisible
End Sub

Function FindLastRow(column As String, sheet As Worksheet) As Long
    ' Find the last used row in a specific column of a worksheet
    FindLastRow = sheet.Cells(sheet.Rows.Count, column).End(xlUp).Row
End Function

Function ImportDataToWorksheet(targetSheetName As String, targetRange As String)
    ' Import data from a file into a specified worksheet and range
    ' (Note: This function contains interaction with the user and opening of external files, consider using it with caution in an automated environment)
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

Sub SaveWorkbookAs(wb As Workbook, Path As String, fileName As String)
    ' Save a workbook with a given filename at the specified path
    wb.SaveAs fileName:=Path & fileName
    wb.Close SaveChanges:=False

End Sub

Function FormatColumn(column As String, format As String, Optional sheet As String)
    ' Apply formatting to the specified column in a worksheet
    ' (Note: The format parameter should be a valid Excel number format)

End Function

Function FormatColumnAsDate(ws As Worksheet, columnRange As String)
    ' Format a column range as date in the specified worksheet
    ws.Range(columnRange).NumberFormat = "m/d/yyyy"

End Function

Function AddCalculationColumn(header As String, formula As String, lastRow As Integer, _
                             Optional columnWidth As Long, Optional sheet As String, _
                             Optional lastRowSheet As String)
    ' Add a new calculation column to the specified worksheet using the provided formula
    ' (Note: This function modifies the worksheet structure and data, use with caution)

    Dim wsTarget As Worksheet
    Dim wsLastRow As Worksheet
    Dim targetColumn As Integer
    
    ' Set default sheet to the active sheet if the sheet is not provided
    If sheet = vbNullString Then
        Set wsTarget = ActiveSheet
    Else
        Set wsTarget = ThisWorkbook.Worksheets(sheet)
    End If
    
    ' Use the same sheet for lastRow if lastRowSheet is not provided
    If lastRowSheet = vbNullString Then
        Set wsLastRow = wsTarget
    Else
        Set wsLastRow = ThisWorkbook.Worksheets(lastRowSheet)
    End If
    
    ' Add a new column on the end
    targetColumn = wsTarget.Cells(1, wsTarget.Columns.Count).End(xlToLeft).Column + 1
    
    ' Set column header
    wsTarget.Cells(1, targetColumn).Value = header
    
    ' Set formula for cell in row 2
    wsTarget.Cells(2, targetColumn).Formula = formula
    
    ' Copy the formula down to lastRow
    If lastRow > 2 Then
        wsTarget.Range(wsTarget.Cells(2, targetColumn), _
                       wsTarget.Cells(lastRow, targetColumn)).FillDown
    End If
    
    ' Autofit or set specific width for the column
    If IsMissing(columnWidth) Then
        wsTarget.Columns(targetColumn).AutoFit
    Else
        wsTarget.Columns(targetColumn).ColumnWidth = columnWidth
    End If

End Function

Function AddFormulaAndCopyDown(ws As Worksheet, startCell As String, lastRow As Long)
    ' Add formula to a cell and copy it down to the last row in the specified worksheet
    ' (Note: This function modifies the worksheet data, use with caution)

End Function

Function AutoFilterColumn(ws As Worksheet, column As String, criteria As String, lastRow As Long)
    ' Apply an AutoFilter to a column based on given criteria in the specified worksheet
    ' (Note: This function modifies the worksheet data and structure, use with caution)

End Function

Function DeleteRowsBasedOnCondition(ws As Worksheet, column As String, condition As String, lastRow As Long)
    ' Delete rows in the specified worksheet based on a given condition in a specific column
    ' (Note: This function modifies the worksheet data and structure, use with caution)

End Function

Sub CopyFormulasAndFilter(sheet As Worksheet, copyRange As String, filterField As Integer, filterCriteria As String)
    ' Copy formulas and apply a filter in the specified worksheet based on a field and criteria
    ' (Note: This function modifies the worksheet data and structure, use with caution)

End Sub

Function GetLastFullReleaseRow(ws As Worksheet, startRow As Long, totalRows As Long) As Long
    ' Find the last full release code row in the specified worksheet
    ' (Note: This function is used for Excel automation and analysis tasks)
    Dim currentRow As Long
    currentRow = startRow
    While currentRow <= totalRows And ws.Cells(currentRow + 1, "E").Value = ws.Cells(currentRow, "E").Value
        currentRow = currentRow + 1
    Wend
    GetLastFullReleaseRow = currentRow

End Function

Sub CreateAndSendEmail(toRecipients As String, subject As String, htmlBody As String, attachmentsPaths As Collection, Optional action As String = "Send")
    ' Create and send an email with optional attachments using Outlook
    ' (Note: This function interacts with the user's email system and may require permission)
    Dim OutApp As Object
    Dim outmail As Object
    Dim attachmentPath As Variant
    
    Set OutApp = CreateObject("Outlook.Application")
    Set outmail = OutApp.CreateItem(0)
    
    With outmail
        .To = toRecipients
        .Subject = subject
        .HTMLBody = htmlBody & .HTMLBody
        
        For Each attachmentPath In attachmentsPaths
            If FileExists(attachmentPath) Then
                .Attachments.Add attachmentPath
            End If
        Next attachmentPath
        
        If action = "Send" Then
            .Send
        Else
            .Display
        End If
    End With
    
    Set outmail = Nothing
    Set OutApp = Nothing

End Sub