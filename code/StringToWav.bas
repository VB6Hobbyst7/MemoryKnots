Attribute VB_Name = "StringToWav"
'needs reference Microsoft Speech Object Library
Option Explicit

Sub TestStringToWavFile()
    'run this to make a wav file from a text input

    Dim sP As String, sFN As String, sStr As String, sFP As String

    'set parameter values - insert your own profile name first
    'paths
    sP = Environ("USERPROFILE") & "\Desktop\"    'for example
    sFN = "Mytest.wav"                           'overwrites if file name same
    sFP = sP & sFN
    
    'string to use for the recording
    sStr = "This is a short test string to be spoken in a user's wave file."
    
    'make voice wav file from string
    StringToWavFile sStr, sFP

End Sub

Function StringToWavFile(sIn As String, sPath As String) As Boolean
    'makes a spoken wav file from parameter text string
    'sPath parameter needs full path and file name to new wav file
    'If wave file does not initially exist it will be made
    'If wave file does initially exist it will be overwritten
    'Needs reference set to Microsoft Speech Object Library
    
    Dim fs As New SpFileStream
    Dim Voice As New SpVoice

    'set the audio format
    fs.Format.Type = SAFT22kHz16BitMono

    'create wav file for writing without events
    fs.Open sPath, SSFMCreateForWrite, False
 
    'Set wav file stream as output for Voice object
    Set Voice.AudioOutputStream = fs

    'send output to default wav file "SimpTTS.wav" and wait till done
    Voice.Speak sIn, SVSFDefault

    'Close file
    fs.Close

    'wait
    Voice.WaitUntilDone (6000)

    'release object variables
    Set fs = Nothing
    Set Voice.AudioOutputStream = Nothing

    'transfers
    StringToWavFile = True

End Function

