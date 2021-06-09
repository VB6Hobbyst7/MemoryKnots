Attribute VB_Name = "VoiceToWav"
'Option Compare Database
Option Explicit

Rem found at http://www.vbarchiv.net

Public Declare PtrSafe Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
(ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_FILENAME = &H20000

Public Declare PtrSafe Function mciSendString Lib "winmm.dll" _
Alias "mciSendStringA" ( _
ByVal lpstrCommand As String, _
ByVal lpstrReturnString As String, _
ByVal uReturnLength As Long, _
ByVal hwndCallback As Long) As Long

Public Enum BitsPerSec
    Bits16 = 16
    Bits8 = 8
End Enum

Public Enum SampelsPerSec
    Sampels8000 = 8000
    Sampels11025 = 11025
    Sampels12000 = 12000
    Sampels16000 = 16000
    Sampels22050 = 22050
    Sampels24000 = 24000
    Sampels32000 = 32000
    Sampels44100 = 44100
    Sampels48000 = 48000
End Enum

Public Enum Channels
    Mono = 1
    Stereo = 2
End Enum

Public Sub play(file As String)
    Dim wavefile
    wavefile = file
    Call sndPlaySound(wavefile, SND_ASYNC Or SND_FILENAME)
End Sub

Public Sub StartRecord(ByVal BPS As BitsPerSec, _
                       ByVal SPS As SampelsPerSec, ByVal Mode As Channels)

    Dim retStr As String
    Dim cBack As Long
    Dim BytesPerSec As Long

    retStr = Space$(128)
    BytesPerSec = (Mode * BPS * SPS) / 8
    mciSendString "open new type waveaudio alias capture", retStr, 128, cBack
    mciSendString "set capture time format milliseconds" & _
                  " bitspersample " & CStr(BPS) & _
                  " samplespersec " & CStr(SPS) & _
                  " channels " & CStr(Mode) & _
                  " bytespersec " & CStr(BytesPerSec) & _
                  " alignment 4", retStr, 128, cBack
    mciSendString "record capture", retStr, 128, cBack
End Sub

Public Sub SaveRecord(strFile)
    Dim retStr As String
    Dim TempName As String
    Dim cBack As Long
    Dim fs, F

    TempName = strFile      'Left$(strFile, 3) & "Temp.wav"
    retStr = Space$(128)
    mciSendString "stop capture", retStr, 128, cBack
    mciSendString "save capture " & TempName, retStr, 128, cBack
    mciSendString "close capture", retStr, 128, cBack

End Sub

Public Sub StartRecord_Click()
    VoiceToWav.StartRecord Bits16, Sampels32000, Mono
End Sub

Public Sub EndRecord_Click()
    VoiceToWav.SaveRecord Environ("USERPROFILE") & "\Desktop\test.wav"
End Sub

Public Sub Play_Click()
    VoiceToWav.play Environ("USERPROFILE") & "\Desktop\test.wav"
End Sub


