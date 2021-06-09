Attribute VB_Name = "JPG"
' Procedure : ExportRangeAsImage
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Capture a picture of a worksheet range and save it to disk
'               Returns True if the operation is successful
' Note      : *** Overwrites files, if already exists, without any warning! ***
' Copyright : The following is release as Attribution-ShareAlike 4.0 International
'             (CC BY-SA 4.0) - https://creativecommons.org/licenses/by-sa/4.0/
' Req'd Refs: Uses Late Binding, so none required
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' ws            : Worksheet to capture the image of the range from
' rng           : Range to capture an image of
' sPath         : Fully qualified path where to export the image to
' sFileName     : filename to save the image to WITHOUT the extension, just the name
' sImgExtension : The image file extension, commonly: JPG, GIF, PNG, BMP
'                   If omitted will be JPG format
'
' Usage:
' ~~~~~~
' ? ExportRangeAsImage(Sheets("Sheet1"), Range("A1"), "C:\Temp\Charts\", "test01". "JPG")
' ? ExportRangeAsImage(Sheets("Products"), Range("D5:F23"), "C:\Temp\Charts", "test02")
' ? ExportRangeAsImage(Sheets("Sheet1"), Range("A1"), "C:\Temp\Charts\", "test01", "PNG")
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2020-04-06              Initial Release
'---------------------------------------------------------------------------------------
Function ExportRangeAsImage(ws As Worksheet, _
                            rng As Range, _
                            sPath As String, _
                            sFileName As String, _
                            Optional sImgExtension As String = "JPG") As Boolean
    Dim oChart                As ChartObject
 
    On Error GoTo Error_Handler
 
    If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
 
    Application.ScreenUpdating = False
    ws.Activate
    rng.CopyPicture xlScreen, xlPicture          'Copy Range Content
    Set oChart = ws.ChartObjects.Add(0, 0, rng.Width, rng.Height) 'Add chart
    oChart.Activate
    With oChart.Chart
        .Paste                                   'Paste our Range
        .Export sPath & sFileName & "." & LCase(sImgExtension), sImgExtension 'Export the chart as an image
    End With
    oChart.Delete                                'Delete the chart
    ExportRangeAsImage = True
 
Error_Handler_Exit:
    On Error Resume Next
    Application.ScreenUpdating = True
    If Not oChart Is Nothing Then Set oChart = Nothing
    Exit Function
 
Error_Handler:
    '76 - Path not found
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & err.Number & vbCrLf & _
           "Error Source: ExportRangeAsImage" & vbCrLf & _
           "Error Description: " & err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function


