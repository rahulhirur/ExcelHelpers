# ExcelHelpers
Help in Excel usage and VBA

jaon to excel


Sub ImportJSONFolderToExcel()
    Dim folderPath As String
    Dim fileName As String
    Dim fileContent As String
    Dim jsonData As Object
    Dim jsonObj As Object
    Dim rowIndex As Integer
    Dim colIndex As Integer
    Dim fso As Object
    
    ' Create a FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Prompt the user to select a folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder Containing JSON Files"
        .AllowMultiSelect = False
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
        Else
            MsgBox "No folder selected. Exiting macro.", vbExclamation
            Exit Sub
        End If
    End With
    
    ' Clear existing data in the active worksheet
    ActiveSheet.Cells.Clear
    
    ' Get a list of all JSON files in the selected folder
    fileName = Dir(folderPath & "\*.json")
    
    ' Loop through each JSON file in the folder
    Do While fileName <> ""
        ' Read the contents of the JSON file
        fileContent = fso.OpenTextFile(folderPath & "\" & fileName).ReadAll
        
        ' Parse the JSON content
        Set jsonData = JsonConverter.ParseJson(fileContent)
        
        ' Loop through each JSON object and write data to Excel
        rowIndex = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row + 1 ' Find the next available row
        For Each jsonObj In jsonData
            colIndex = 1
            For Each key In jsonObj.keys
                Cells(rowIndex, colIndex).Value = key
                Cells(rowIndex + 1, colIndex).Value = jsonObj(key)
                colIndex = colIndex + 1
            Next key
            rowIndex = rowIndex + 2
        Next jsonObj
        
        ' Get the next JSON file in the folder
        fileName = Dir
    Loop
    
    MsgBox "JSON files imported successfully!", vbInformation
End Sub

