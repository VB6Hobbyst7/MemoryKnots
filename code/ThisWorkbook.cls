VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Public cControl As CommandBarButton

Private Sub Workbook_AddinInstall()
    Call AddCommandbar
    'if MyDocuments\MemoryKnotes doesn't exist create it
    Call PathCreate("MemoryKnotes")
End Sub

Private Sub Workbook_AddinUninstall()
    On Error Resume Next
    Dim bar As CommandBarControl
    For Each bar In Application.CommandBars("Worksheet Menu Bar").Controls
        If bar.Caption = "MemoryKnotes" Then bar.Delete
    Next
End Sub

