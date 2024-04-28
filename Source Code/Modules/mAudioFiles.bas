Attribute VB_Name = "mAudioFiles"
Option Explicit
' *** This module stores all data related to loading, reading and manipulating the audio files in the project ***

' *** Libraries for importing an audio file to Excel. ***
#If VBA7 Then ' Excel 2010 or later
    Private Declare PtrSafe Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As LongPtr) As Long
#Else ' Excel 2007 or earlier
    Private Declare Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
#End If

' *** Library for reading the .wav file ***
#If VBA7 Then ' Excel 2010 or later
    Private Declare PtrSafe Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" _
        (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
         ByVal uReturnLength As LongPtr, ByVal hwndCallback As LongPtr) As Long
#Else ' Excel 2007 or earlier
    Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" _
        (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, _
         ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
#End If

Public Function GetWAVDurationInSeconds(ByVal strPath As String) As Double
    ' ***
    ' Reads the duration of the audio file at the specified path.
    ' Written by ChatGPT, might be unstable.
    ' ***

    ' Accepts:
    '   - strPath [string] - the path where the file is located.
    ' Returns:
    '   - double - the duration of the .wav file in seconds.

    Dim dblReturnValue As Double
    Dim strDurationCode As String * 255 ' Changed to a string variable for the return value

    ' Check if the file exists
    If Not mUtils.IsFileExists(strPath) Then
        Debug.Print "WARNING! The file at path " & strPath & " does not exist!"
        Exit Function
    End If

    ' Open the .wav file
    '  alias audiofile
    dblReturnValue = mciSendString("open """ & strPath & """ type waveaudio alias audiofile", 0&, 0, 0)
    If dblReturnValue <> 0 Then
        Debug.Print "Failed to open audio file. Error code: " & dblReturnValue
        Exit Function
    End If

    ' Send command to get the duration
    dblReturnValue = mciSendString("status audiofile length", strDurationCode, Len(strDurationCode), 0)
    If dblReturnValue <> 0 Then
        Debug.Print "Failed to get audio duration. Error code: " & dblReturnValue
        Exit Function
    End If

    ' Close the .wav file
    dblReturnValue = mciSendString("close audiofile", 0&, 0, 0)
    If dblReturnValue <> 0 Then
        Debug.Print "Failed to close audio file. Error code: " & dblReturnValue
        Exit Function
    End If

    ' Convert duration string to seconds
    GetWAVDurationInSeconds = Val(Left(strDurationCode, InStr(strDurationCode, Chr(0)) - 1)) / 1000
End Function

Public Sub PlaySoundAsynchronymusly(ByVal strPath As String)
    ' *** Plays the .wav file, but without waiting for code to be executed. ***
    
    ' Accepts:
    '   - strPath [string] - the place in a drive where a .wav file is located
    ' Returns:
    '   - None
    
    If Not mUtils.IsFileExists(strPath) Then
        Debug.Print "WARNING! The file at path " & strPath & " does not exist!"
        Exit Sub
    End If
    
    ' This means that sound will be played asynchronymously.
    Const SND_ASYNC As Long = &H1
    Call sndPlaySound32(strPath, SND_ASYNC)
End Sub
