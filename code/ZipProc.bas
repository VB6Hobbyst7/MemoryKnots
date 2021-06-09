Attribute VB_Name = "ZipProc"
' Zip a file or a folder to a zip file/folder using Windows Explorer.
' Default behaviour is similar to right-clicking a file/folder and selecting:
'   Send to zip file.
'
' Parameters:
'   Path:
'       Valid (UNC) path to the file or folder to zip.
'   Destination:
'       (Optional) Valid (UNC) path to file with zip extension or other extension.
'   Overwrite:
'       (Optional) Leave (default) or overwrite an existing zip file.
'       If False, the created zip file will be versioned: Example.zip, Example (2).zip, etc.
'       If True, an existing zip file will first be deleted, then recreated.
'
'   Path and Destination can be relative paths. If so, the current path is used.
'
'   If success, 0 is returned, and Destination holds the full path of the created zip file.
'   If error, error code is returned, and Destination will be zero length string.
'
' Early binding requires references to:
'
'   Shell:
'       Microsoft Shell Controls And Automation
'
'   Scripting.FileSystemObject:
'       Microsoft Scripting Runtime
'
' 2017-10-22. Gustav Brock. Cactus Data ApS, CPH.
'
Public Function Zip( _
       ByVal Path As String, _
       Optional ByRef Destination As String, _
       Optional ByVal Overwrite As Boolean) _
        As Long
   
    #If EarlyBinding Then
        ' Microsoft Scripting Runtime.
        Dim FileSystemObject    As Scripting.FileSystemObject
        ' Microsoft Shell Controls And Automation.
        Dim ShellApplication    As Shell
   
        Set FileSystemObject = New Scripting.FileSystemObject
        Set ShellApplication = New Shell
    #Else
        Dim FileSystemObject    As Object
        Dim ShellApplication    As Object
        Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
        Set ShellApplication = CreateObject("Shell.Application")
    #End If
   
    ' Mandatory extension of zip file.
    Const ZipExtensionName  As String = "zip"
    Const ZipExtension      As String = "." & ZipExtensionName
    ' Custom error values.
    Const ErrorPathFile     As Long = 75
    Const ErrorOther        As Long = -1
    Const ErrorNone         As Long = 0
    ' Maximum (arbitrary) allowed count of created zip versions.
    Const MaxZipVersion     As Integer = 1000
   
    Dim ZipHeader           As String
    Dim ZipPath             As String
    Dim ZipName             As String
    Dim ZipFile             As String
    Dim ZipBase             As String
    Dim ZipTemp             As String
    Dim Version             As Integer
    Dim result              As Long
   
    If FileSystemObject.FileExists(Path) Then
        ' The source is an existing file.
        ZipName = FileSystemObject.GetBaseName(Path) & ZipExtension
        ZipPath = FileSystemObject.GetFile(Path).ParentFolder
    ElseIf FileSystemObject.FolderExists(Path) Then
        ' The source is an existing folder.
        ZipName = FileSystemObject.GetBaseName(Path) & ZipExtension
        ZipPath = FileSystemObject.GetFolder(Path).ParentFolder
    Else
        ' The source does not exist.
    End If
       
    If ZipName = "" Then
        ' Nothing to zip. Exit.
        Destination = ""
    Else
        If Destination <> "" Then
            If FileSystemObject.GetExtensionName(Destination) = "" Then
                ' Destination is a folder.
                ZipPath = Destination
            Else
                ' Destination is a file.
                ZipName = FileSystemObject.GetFileName(Destination)
                ZipPath = FileSystemObject.GetParentFolderName(Destination)
            End If
        Else
            ' Use the already found folder of the source.
        End If
        ZipFile = FileSystemObject.BuildPath(ZipPath, ZipName)
        If FileSystemObject.FileExists(ZipFile) Then
            If Overwrite = True Then
                ' Delete an existing file.
                FileSystemObject.DeleteFile ZipFile, True
                ' At this point either the file is deleted or an error is raised.
            Else
                ZipBase = FileSystemObject.GetBaseName(ZipFile)
                ' Modify name of the zip file to be created to preserve an existing file:
                '   "Example.zip" -> "Example (2).zip", etc.
                Version = Version + 1
                Do
                    Version = Version + 1
                    ZipFile = FileSystemObject.BuildPath(ZipPath, ZipBase & Format(Version, " \(0\)") & ZipExtension)
                Loop Until FileSystemObject.FileExists(ZipFile) = False Or Version > MaxZipVersion
                If Version > MaxZipVersion Then
                    ' Give up.
                    err.Raise ErrorPathFile, "Zip Create", "File could not be created."
                End If
            End If
        End If
   
        ' Create a temporary zip name to allow for a final destination file with another extension than zip.
        ZipTemp = FileSystemObject.BuildPath(ZipPath, FileSystemObject.GetBaseName(FileSystemObject.GetTempName()) & ZipExtension)
        ' Create empty zip folder.
        ' Header string provided by Stuart McLachlan <stuart@lexacorp.com.pg>.
        ZipHeader = Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, vbNullChar)
        With FileSystemObject.OpenTextFile(ZipTemp, ForWriting, True)
            .Write ZipHeader
            .Close
        End With
       
        ' Resolve relative paths.
        ZipTemp = FileSystemObject.GetAbsolutePathName(ZipTemp)
        Path = FileSystemObject.GetAbsolutePathName(Path)
        ' Copy the source file/folder into the zip file.
        With ShellApplication
            Debug.Print Timer, "Zipping started . ";
            .Namespace(CVar(ZipTemp)).CopyHere CVar(Path)
            ' Ignore error while looking up the zipped file before is has been added.
            On Error Resume Next
            ' Wait for the file to created.
            Do Until .Namespace(CVar(ZipTemp)).Items.count = 1
                ' Wait a little ...
                Application.Wait (Now + TimeValue("0:00:01"))
                Debug.Print ".";
            Loop
            Debug.Print
            ' Resume normal error handling.
            On Error GoTo 0
            Debug.Print Timer, "Zipping finished."
        End With
        ' Rename the temporary zip to its final name.
        FileSystemObject.MoveFile ZipTemp, ZipFile
    End If
   
    Set ShellApplication = Nothing
    Set FileSystemObject = Nothing
   
    If err.Number <> ErrorNone Then
        Destination = ""
        result = err.Number
    ElseIf Destination = "" Then
        result = ErrorOther
    End If
   
    Zip = result
End Function


