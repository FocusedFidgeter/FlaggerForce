Attribute VB_Name = "stage3"
Option Explicit

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
    Dim TemplateFolderPath As String
    Dim ServerFolderPath As String
    Dim NewFolderName As String
    
    ' Prompt user to select source folder
    TemplateFolderPath = PickFolder("Select Source Folder")
    If TemplateFolderPath = "" Then
        MsgBox "No source folder selected. Exiting."
        Exit Sub
    End If
    
    ' Prompt user to select target folder
    ServerFolderPath = PickFolder("Select Target Folder")
    If ServerFolderPath = "" Then
        MsgBox "No target folder selected. Exiting."
        Exit Sub
    End If
    
    ' Prompt user for new folder name in mm.dd format with the current date as default input
    NewFolderName = InputBox("Enter the new folder name in mm.dd format:", "New Folder Name", Format(Date, "mm.dd"))
    If NewFolderName = "" Then
        MsgBox "No folder name provided. Exiting."
        Exit Sub
    End If
    
    ' Create new folder path
    Dim NewFolderPath As String
    NewFolderPath = ServerFolderPath & Application.PathSeparator & NewFolderName
    
    ' Create the folder if it does not exist or exit if it already exists
    If Not FileExists(NewFolderPath) Then
        CreateFolderIfNotExists NewFolderPath
    Else
        MsgBox "The folder already exists. Exiting."
        Exit Sub
    End If
    
    ' Copy the entire source folder into the target folder with the new folder name
    CopyFolder TemplateFolderPath, NewFolderPath
    
    ' Define the source and destination file paths for renaming
    Dim SourceFilePath As String
    Dim DestFilePath As String
    Dim TemplateFileName As String
    
    TemplateFileName = "Duke Book Template.xlsm" ' Name of the template file inside the source folder
    SourceFilePath = NewFolderPath & Application.PathSeparator & TemplateFileName
    DestFilePath = NewFolderPath & Application.PathSeparator & "Duke Book " & dateMMDD & ".xlsm"
    
    ' Rename the file if it exists
    If FileExists(SourceFilePath) Then
        Name SourceFilePath As DestFilePath
        MsgBox "File renamed successfully."
    Else
        MsgBox "Template file not found. Cannot rename."
    End If
    
    MsgBox "Folder copied successfully."
End Sub

Function PickFolder(prompt As String) As String
    ' Function to display a FolderPicker and return the selected path
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fd
        .Title = prompt
        .AllowMultiSelect = False
        If .Show = -1 Then
            PickFolder = .SelectedItems(1)
        End If
    End With
    
    Set fd = Nothing
End Function

Sub CopyFolder(SourceFolderPath As String, TargetFolderPath As String)
    ' Copy the entire source folder into the target folder
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    fso.CopyFolder Source:=SourceFolderPath, Destination:=TargetFolderPath, OverwriteFiles:=True
    If Err.Number <> 0 Then
        MsgBox "Error in copying folder: " & Err.Description
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    
    ' If needed, renaming of files inside the folder can be handled here
    
    Set fso = Nothing
End Sub