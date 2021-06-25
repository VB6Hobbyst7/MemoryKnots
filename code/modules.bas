Attribute VB_Name = "modules"

Function OutlookCheck() As Boolean
    Dim xOLApp As Object
    '    On Error GoTo L1
    Set xOLApp = CreateObject("Outlook.Application")
    If Not xOLApp Is Nothing Then
        OutlookCheck = True
        '        MsgBox "Outlook " & xOLApp.Version & " installed", vbExclamation
        Set xOLApp = Nothing
        Exit Function
    End If
    OutlookCheck = False
    'L1: MsgBox "Outlook not installed", vbExclamation, "Kutools for Outlook"
End Function

Function Clipboard(Optional StoreText As String) As String
    'PURPOSE: Read/Write to Clipboard
    'Source: ExcelHero.com (Daniel Ferry)

    Dim X As Variant

    'Store as variant for 64-bit VBA support
    X = StoreText

    'Create HTMLFile Object
    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            Select Case True
            Case Len(StoreText)
                'Write to the clipboard
                .SetData "text", X
            Case Else
                'Read from the clipboard (no variable passed through)
                Clipboard = .GetData("text")
            End Select
        End With
    End With

End Function




Sub AddCommandbar()
    On Error Resume Next                         'Just in case
    'Delete any existing menu item that may have been left.
    Dim bar As CommandBarControl
    For Each bar In Application.CommandBars("Worksheet Menu Bar").Controls
        If bar.Caption = "MemoryKnots" Then bar.Delete
        'Debug.Print bar.Caption
    Next

    '    Application.CommandBars("Worksheet Menu Bar").Controls("MemoryKnots").Delete
    
    'Add the new menu item and set a CommandBarButton variable to it
    Set cControl = Application.CommandBars("Worksheet Menu Bar").Controls.Add
    With cControl
        .Caption = "MemoryKnots"
        .Style = msoButtonIconAndCaption
        .FaceId = 1838
        .OnAction = "Simple_Notes"               'Macro stored in a Standard Module
    End With
    On Error GoTo 0
End Sub

Sub Simple_Notes()
    If Not IsLoaded("SimpleNotes") Then
        Call OpenUserForm("SimpleNotes")
    End If
End Sub

Sub OpenUserForm(formName As String)
    Dim oUserForm As Object
    On Error GoTo err
    Set oUserForm = UserForms.Add(formName)
    oUserForm.Show (vbModeless)
    Exit Sub
err:
    Select Case err.Number
    Case 424:
        MsgBox "The Userform with the name " & formName & " was not found.", vbExclamation, "Load userforn by name"
    Case Else:
        MsgBox err.Number & ": " & err.Description, vbCritical, "Load userforn by name"
    End Select
End Sub

Function IsLoaded(formName As String) As Boolean
    Dim frm As Object
    For Each frm In VBA.UserForms
        If frm.Name = formName Then
            IsLoaded = True
            Exit Function
        End If
    Next frm
    IsLoaded = False
End Function

Sub ListboxSortAZ(myListBox As MSForms.ListBox, Optional resetMacro As String)
    'Create variables
    Dim j As Long
    Dim i As Long
    Dim TEMP As Variant
    'Reset the listBox into standard order
    If resetMacro <> "" Then
        Run resetMacro, myListBox
    End If
    'Use Bubble sort method to put listBox in A-Z order
    With myListBox
        For j = 0 To .ListCount - 2
            For i = 0 To .ListCount - 2
                If LCase(.List(i)) > LCase(.List(i + 1)) Then
                    TEMP = .List(i)
                    .List(i) = .List(i + 1)
                    .List(i + 1) = TEMP
                End If
            Next i
        Next j
    End With
End Sub

Sub ListboxSortZA(myListBox As MSForms.ListBox, Optional resetMacro As String)
    'Create variables
    Dim j As Long
    Dim i As Long
    Dim TEMP As Variant

    'Reset the listBox into standard order
    If resetMacro <> "" Then
        Run resetMacro, myListBox
    End If

    'Use Bubble sort method to put listBox in Z-A order
    With myListBox
        For j = 0 To .ListCount - 2
            For i = 0 To .ListCount - 2
                If LCase(.List(i)) < LCase(.List(i + 1)) Then
                    TEMP = .List(i)
                    .List(i) = .List(i + 1)
                    .List(i + 1) = TEMP
                End If
            Next i
        Next j
    End With
End Sub

Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function

Function SelectionValues(link As String)
    Dim c As Range
    If TypeName(Selection) = "Range" _
                             And Selection.Cells.count = 1 Then
        SelectionValues = Selection.Value
    ElseIf TypeName(Selection) = "Range" Then
        For Each c In Selection.SpecialCells(xlCellTypeVisible)
            If Len(c.Value) <> 0 Then
                If SelectionValues = "" Then
                    SelectionValues = c.Value
                Else
                    SelectionValues = SelectionValues & link & c.Value
                End If
            End If
        Next c
    End If
End Function

Function Listbox_Selected(LBox As MSForms.ListBox, Count_Indexes_Values As Integer)

    Dim SelectedIndexes As String
    Dim SelectedValues As String
    Dim SelectedCount As Integer
    Dim i As Long
    With LBox
        For i = 0 To .ListCount - 1
            If .Selected(i) Then
                'items count
                SelectedCount = SelectedCount + 1
                'items' indexes
                SelectedIndexes = SelectedIndexes & i & ","
                'items' values
                SelectedValues = SelectedValues & .List(i) & ","
            End If
        Next i
    End With
       
    If SelectedCount = 0 Then
        Listbox_Selected = 0
        Exit Function
    End If
    SelectedIndexes = Left(SelectedIndexes, Len(SelectedIndexes) - 1)
    
    SelectedValues = Left(SelectedValues, Len(SelectedValues) - 1)
    
    Select Case Count_Indexes_Values
    Case Is = 1
        Listbox_Selected = SelectedCount
    Case Is = 2
        Listbox_Selected = SelectedIndexes
    Case Is = 3
        Listbox_Selected = SelectedValues
    End Select
   
End Function

Sub PathCreate(strFolder As String)
    Dim WshShell As Object
    Dim fso As Object
    Dim SpecialPath As String

    Set WshShell = CreateObject("WScript.Shell")
    Set fso = CreateObject("scripting.filesystemobject")

    SpecialPath = WshShell.SpecialFolders("MyDocuments")

    If Right(SpecialPath, 1) <> "\" Then
        SpecialPath = SpecialPath & "\"
    End If
    
    If fso.FolderExists(SpecialPath & strFolder) = False Then
        On Error Resume Next
        MkDir SpecialPath & strFolder
        On Error GoTo 0
    End If
cleanup:
    Set WshShell = Nothing
    Set fso = Nothing
End Sub

Function PathGet()
    Dim WshShell As Object
    Dim fso As Object
    Dim SpecialPath As String

    Set WshShell = CreateObject("WScript.Shell")
    Set fso = CreateObject("scripting.filesystemobject")

    SpecialPath = WshShell.SpecialFolders("MyDocuments")

    If Right(SpecialPath, 1) <> "\" Then
        SpecialPath = SpecialPath & "\"
    End If
    
    PathGet = SpecialPath & "MemoryKnots\"
    '    Debug.Print PathGet
End Function

Sub ListboxClearSelection(LBox As MSForms.ListBox)
    On Error Resume Next
    For i = 0 To LBox.ListCount
        LBox.Selected(i) = False
    Next i
End Sub

Sub ListboxSelectValue(LBox As MSForms.ListBox, str As String, Optional clr As Boolean = True)
    If Not clr Is Nothing Then
        Call ListboxClearSelection(LBox)
    End If

    For i = 0 To LBox.ListCount - 1
        If LBox.List(i) = str Then
            LBox.Selected(i) = True
            Exit Sub
        End If
    Next i
End Sub


