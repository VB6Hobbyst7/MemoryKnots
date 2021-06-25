Attribute VB_Name = "UpdateAddin"


Sub ExportNotes()
    Dim fso As Object
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim strPath As String

    strPath = PathGet                            'PickFolder
    If strPath = "" Then Exit Sub
    
    On Error Resume Next
    '    Kill strPath & "MemoryKnots.xlam"
    fso.CopyFile Workbooks("MemoryKnots.xlam").FullName, strPath
    
    Workbooks("MemoryKnots.xlam").Sheets.Copy

    'Workbooks("MemoryKnots.xlam").Sheets(Array("SETTINGS", "> NOTES", "> RESOLVED")).Copy   'or .delete
    
    ActiveWorkbook.Sheets("SETTINGS").Delete
    ActiveWorkbook.SaveAs Filename:=strPath & "MemoryKnots.xlsx", _
                          FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWorkbook.Saved = True
    ActiveWorkbook.Close
    CreateObject("WScript.Shell").PopUp "Successfully exported to " & Chr(10) & strPath, 1
    'MsgBox "Successfully exported to " & Chr(10) & strPath
    ActiveWindow.Close
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub

Function PickFolder()
    With Application.FileDialog(msoFileDialogFolderPicker)
        '.Title=
        '.ButtonName=
        .InitialFileName = Environ("USERprofile") & "\Desktop\" 'ThisWorkbook.Path 'Left(ThisWorkbook.Path, InStrRev(ThisWorkbook.Path, "\"))
        If .Show = -1 Then                       ' if OK is pressed
            PickFolder = .SelectedItems(1) & "\"
        Else
            Exit Function
        End If
    End With
End Function

Sub ImportNotes()
    Dim strADMIN As String
    Dim answer As Integer
    answer = MsgBox("ATTENTION!" & Chr(10) & Chr(10) & _
                    "Present Notebooks will be DELETED and REPLACED from IMPORT file" & Chr(10) & Chr(10) & _
                    "Proceed? (YES) or Cancel import? (NO)", _
                    vbYesNo)
    If answer = vbYes Then
        '        KeepLoading = True
    Else
      
        '        'stop update to clear list(0)= update
        '        strADMIN = InputBox("Non ADMIN ignore this")
        '        If UCase(strADMIN) = "FREEZE" Then
        '            KeepLoading = True
        '            '        Workbooks("MemoryKnots").Sheets("> NOTES").Cells(2, 2).EntireRow.Delete
        '            Exit Sub
        '        End If
      
        'boolean to force unload userform
        '        KeepLoading = False
        Exit Sub
    End If

    Dim AddinWorkbook As Workbook
    Set AddinWorkbook = Workbooks("MemoryKnots.xlam")


    Dim ImportWorkbook As Workbook
    If IsWorkBookOpen("MemoryKnots.xlsx") Then
        Set ImportWorkbook = Workbooks("MemoryKnots.xlsx")
    Else
        If Dir(PathGet & "MemoryKnots.xlsx") > Len(PathGet) Then
            Set ImportWorkbook = Workbooks.Open(PathGet & "MemoryKnots.xlsx")
        Else
            CreateObject("WScript.Shell").PopUp "MemoryKnots.xlsx not found. Use EXPORT first.", 1
            '            MsgBox "MemoryKnots.xlsx not found. Use EXPORT first."
            GoTo cleanup
        End If
    End If

    ImportWorkbook.Save

    'import procedure

    Application.ScreenUpdating = False
    
    AddinWorkbook.IsAddin = False
    Dim ws As Worksheet
    
    '    Set ws = AddinWorkbook.Sheets.Add(before:=Sheets(1))
    '    ws.Name = "tmp"
    '    Set ws = Nothing

    Application.DisplayAlerts = False

    For Each ws In AddinWorkbook.Worksheets
        If ws.Name <> "SETTINGS" Then ws.Delete
    Next ws

    'format sheets to be imported
    Dim cell As Range
    For Each ws In ImportWorkbook.Worksheets
        If Left(ws.Name, 1) = ">" Then
            For Each cell In ws.Columns(2)
                If cell.Value = "" Then Exit For
                If cell.Offset(0, -1) = "" Then
                    cell.Offset(0, -1) = Now()
                End If
            Next cell
        End If
    Next ws


    'import sheets
    For Each ws In ImportWorkbook.Worksheets
        If ws.Name <> "SETTINGS" Then
            ws.Copy After:=AddinWorkbook.Sheets(AddinWorkbook.Sheets.count)
        End If
    Next ws

    '    AddinWorkbook.Worksheets("tmp").Delete

cleanup:
    Application.DisplayAlerts = True

    AddinWorkbook.IsAddin = True
    Set ws = Nothing
    Set ImportWorkbook = Nothing
    Set AddinWorkbook = Nothing

    Application.ScreenUpdating = True
End Sub

Sub testgetfile()
    Debug.Print GetFileToImport("xlsx", False)
End Sub

Function GetFileToImport(Optional fileType As String, Optional multiSelect As Boolean) As String

    Dim blArray As Boolean
    Dim strErrMsg As String, strTitle As String
    strTitle = "Import Notebooks"
       
    'check whether the file type parameter was passed
    If IsMissing(fileType) Then
        Exit Function
    End If
    'proceed
    If strErrMsg = vbNullString Then
        ' set title of dialog box
        With Application.FileDialog(msoFileDialogFilePicker)
            .InitialFileName = "MemoryKnots.xlsx"     'Environ("USERprofile") & "\Desktop\"  'ThisWorkbook.Path 'Left(ThisWorkbook.Path, InStrRev(ThisWorkbook.Path, "\"))
            .AllowMultiSelect = multiSelect
            .Filters.Clear
            If blArray Then .Filters.Add "File type", "*." & fileType
            .Title = strTitle
            
            ' show the file picker dialog box
            If .Show <> 0 Then
                GetFilePath = .SelectedItems(1)
            End If
        End With
        ' error message
    Else
        MsgBox strErrMsg, vbCritical, "Error!"
    End If
End Function

Function IsWorkBookOpen(Name As String) As Boolean
    Dim xWb As Workbook
    On Error Resume Next
    Set xWb = Application.Workbooks.Item(Name)
    IsWorkBookOpen = (Not xWb Is Nothing)
End Function


