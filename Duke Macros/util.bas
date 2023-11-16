Attribute VB_Name = "util"
Option Explicit

Sub AddCalculationColumn(header As String, formula As String, lastRow As Integer, _
                         Optional columnWidth As Long, Optional sheet As String, _
                         Optional lastRowSheet As String)

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

End Sub

Sub FormatColumn(column As String, format As String, Optional sheet As String)

    Dim ws As Worksheet
    Dim rangeToFormat As Range
    
    ' Set default sheet to the active sheet if the sheet is not provided
    If sheet = vbNullString Then
        Set ws = ActiveSheet
    Else
        Set ws = ThisWorkbook.Worksheets(sheet)
    End If
    
    ' Set the range to format from row 2 to row 10000 of the specified column
    Set rangeToFormat = ws.Range(column & "2:" & column & "10000")
    
    ' Apply formatting to the range
    With rangeToFormat
        .NumberFormat = format
    End With

End Sub

Function CreateNewWorkbook() As Workbook
    ' Create a new workbook and return a reference to it
    Set CreateNewWorkbook = Application.Workbooks.Add
End Function

Function GetLastFullReleaseRow(ws As Worksheet, startRow As Long, totalRows As Long) As Long
    ' Find the last full release code row
    Dim currentRow As Long
    currentRow = startRow
    While currentRow <= totalRows And ws.Cells(currentRow + 1, "E").Value = ws.Cells(currentRow, "E").Value
        currentRow = currentRow + 1
    Wend
    GetLastFullReleaseRow = currentRow
End Function

Function SaveWorkbookAs(wb As Workbook, Path As String, fileName As String)
    ' Save a workbook with a given filename at the specified path
    wb.SaveAs fileName:=Path & fileName
    wb.Close SaveChanges:=False
End Function

Function IsWorkBookOpen(Name As String) As Boolean
    Dim xWb As Workbook
    On Error Resume Next
    Set xWb = Application.Workbooks.Item(Name)
    IsWorkBookOpen = (Not xWb Is Nothing)
End Function

Function AddFormulaAndCopyDown(ws As Worksheet, startCell As String, lastRow As Long)
    ' Add formula to a cell and copy it down to the last row
    ws.Range(startCell & "2").Copy
    ws.Range(startCell & "3:" & startCell & lastRow).PasteSpecial Paste:=xlPasteFormulas
    Application.CutCopyMode = False
End Function

Function DeleteRowsBasedOnCondition(ws As Worksheet, column As String, condition As String, lastRow As Long)
    ' Loop from the end towards the top to delete rows based on a condition
    Dim i As Long
    For i = lastRow To 2 Step -1
        If ws.Cells(i, column).Value = condition Then
            ws.Rows(i).Delete
        End If
    Next i
End Function

Function FormatColumnAsDate(ws As Worksheet, columnRange As String)
    ' Format a column range as date
    ws.Range(columnRange).NumberFormat = "m/d/yyyy"
End Function

Function AutoFilterColumn(ws As Worksheet, column As String, criteria As String, lastRow As Long)
    ' Apply an AutoFilter to a column based on given criteria
    ws.Range(column & "1:" & column & lastRow).AutoFilter Field:=1, Criteria1:=criteria
End Function

Sub CopyFormulasAndFilter(sheet As Worksheet, copyRange As String, filterField As Integer, filterCriteria As String)
    Dim lastRow As Long
    lastRow = sheet.Cells(sheet.Rows.Count, "A").End(xlUp).Row
    With sheet
        .Range(copyRange & "2").Copy
        .Range(copyRange & "3:" & copyRange & lastRow).PasteSpecial Paste:=xlPasteFormulas
        .Range(copyRange & "1:" & copyRange & lastRow).AutoFilter Field:=filterField, Criteria1:=filterCriteria
    End With
    Application.CutCopyMode = False
End Sub

Function WorksheetExists(sheetName As String, wb As Workbook) As Boolean
    Dim sheet As Worksheet
    On Error Resume Next
    Set sheet = wb.Sheets(sheetName)
    WorksheetExists = Not sheet Is Nothing
    On Error GoTo 0
End Function

Sub AddOrGetWorksheet(sheetName As String, ByRef outSheet As Worksheet)
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
    FindLastRow = sheet.Cells(sheet.Rows.Count, column).End(xlUp).Row
End Function

Function FileExists(filePath As String) As Boolean
    FileExists = (Dir(filePath) <> "")
End Function

Sub MoveFile(sourceFilePath As String, destFilePath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If FileExists(sourceFilePath) Then
        fso.MoveFile Source:=sourceFilePath, Destination:=destFilePath
    End If
End Sub

Sub CopyFile(sourceFilePath As String, destFilePath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If FileExists(sourceFilePath) Then
        fso.CopyFile Source:=sourceFilePath, Destination:=destFilePath
    End If
End Sub

Sub CreateAndSendEmail(toRecipients As String, subject As String, htmlBody As String, attachmentsPaths As Collection, Optional action As String = "Send")
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

Sub CreateFolderIfNotExists(folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
End Sub