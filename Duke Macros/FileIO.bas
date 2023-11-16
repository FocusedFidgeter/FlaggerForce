Attribute VB_Name = "FileIO"

Sub CombineTimesheets()
'
' Split Invoices Macro
' split the combined invoices that are downloaded from Intaact when printing >50 of them.
'

'
    Dim pythonExe As String
    Dim scriptPath As String
    Dim command As String

    ' Refresh the data we will use in the python script
    Sheets("TimesheetCombiner").Visible = True
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=True
    Sheets("TimesheetCombiner").Visible = False

    ' Set the path to the Python executable
    pythonExe = "python"

    ' Set the path to the Python script
    scriptPath = "\\hum-vmqb-01\Billing\Duke\Resources\Processors\timesheet_combiner_duke.py"

    ' Create the command to run the Python script
    command = pythonExe & " " & Chr(34) & scriptPath & Chr(34)

    ' Execute the command
    Shell command, vbNormalFocus
End Sub

Sub SplitInvoices()
'
' Split Invoices Macro
' split the combined invoices that are downloaded from Intaact when printing >50 of them.
'

'
    Dim pythonExe As String
    Dim scriptPath As String
    Dim command As String

    ' Set the path to the Python executable
    pythonExe = "python"

    ' Set the path to the Python script
    scriptPath = "\\hum-vmqb-01\Billing\Duke\Resources\Processors\invoice_splitter_duke.pyw"

    ' Create the command to run the Python script
    command = pythonExe & " " & Chr(34) & scriptPath & Chr(34)

    ' Execute the command
    Shell command, vbNormalFocus
End Sub

Sub CreateWeeklyFolder()
    Dim TemplateFolder As FileDialog
    Dim ServerFolder As FileDialog
    Dim TemplateFolderPath As String
    Dim ServerFolderPath As String
    Dim NewFolderName As String
    Dim fso As Scripting.FileSystemObject
    Dim TemplateFileName As String
    
    ' Create FileDialog instances for selecting folders
    Set TemplateFolder = Application.FileDialog(msoFileDialogFolderPicker)
    Set ServerFolder = Application.FileDialog(msoFileDialogFolderPicker)
    
    ' Prompt user to select source folder
    With TemplateFolder
        .Title = "Select Source Folder"
        .AllowMultiSelect = False
        If .Show = -1 Then
            TemplateFolderPath = .SelectedItems(1)
        Else
            MsgBox "No folder selected. Exiting."
            Exit Sub
        End If
    End With
    
    ' Prompt user to select target folder
    With ServerFolder
        .Title = "Select Target Folder"
        .AllowMultiSelect = False
        If .Show = -1 Then
            ServerFolderPath = .SelectedItems(1)
        Else
            MsgBox "No folder selected. Exiting."
            Exit Sub
        End If
    End With
    
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

'    ' Define the source file name
'    TemplateFileName = "Duke Book Template"
'
'    ' Rename the file
'    On Error Resume Next
'    Name ServerFolderPath & Application.PathSeparator & NewFolderName & Application.PathSeparator & TemplateFileName & ".xlsm" As ServerFolderPath & Application.PathSeparator & NewFolderName & Application.PathSeparator & "Duke Book " & NewFolderName & ".xlsm"
'    If Err.Number <> 0 Then
'        MsgBox "Error in renaming file: " & Err.Description
'    Else
'        MsgBox "File renamed successfully."
'    End If
'    On Error GoTo 0
    
    ' Clean up
    Set fso = Nothing
    Set TemplateFolder = Nothing
    Set ServerFolder = Nothing
End Sub



