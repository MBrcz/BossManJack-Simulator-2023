Attribute VB_Name = "mBSJResponse"
Option Explicit
' ***
' This module handles all operations related to the BossmanJack and it's potential responses.
' ***

Public boolDoDebug As Boolean
' Handle animation calling.
Private boolIsBossManResponding As Boolean
Private strSourceSheetName As String
' See more: InitializeModule function.
Private Const strPortrairSignature = "img_PortrairBSJ"
Private strResponseName As String
Private intResponseItem As Integer
Private intTotalResponseCount As Integer
Private strAudioPath As String
Private strFramesPath As String
Private rngPortrairRange As Range

' - ENUM -
Public Enum eBSJResponse
    win = 1
    loss = 2
    rage = 3
    beg = 4
    big_win = 5
End Enum
' - END ENUM -

Public Property Get IsBossmanResponding() As Boolean
    IsBossmanResponding = boolIsBossManResponding
End Property

Public Property Get SourceSheetName() As String
    SourceSheetName = strSourceSheetName
End Property

Public Property Get PortrairSignature() As String
    PortrairSignature = strPortrairSignature
End Property

Private Function GetResponseName(ByVal enumResponse As eBSJResponse) As String
    ' *** Reads the response name of the bossman. ***
    
    ' Accepts:
    '   - enumResponse [enum] - the number of response written
    ' Returns:
    '   - string - the translated name of the response
    
    Dim strResponse As String
    
    Select Case enumResponse
        Case Is = win
            strResponse = "Win"
        Case Is = loss
            strResponse = "Loss"
        Case Is = rage
            strResponse = "Rage"
        Case Is = beg
            strResponse = "Beg"
        Case Is = big_win
            strResponse = "BigWin"
    End Select
    
    GetResponseName = strResponse
End Function

Private Function GetResponseCount(ByVal enumResponse As eBSJResponse) As Integer
    ' *** Returns the count of the responses that Bossman can give back. ***
    
    ' Accepts:
    '   - enumResponse [enum] - the response that Bossman gives
    
    ' Returns:
    '   - integer - the quantity of items that each response can have
    
    Dim intItemCount As Integer
    
    Select Case enumResponse
        Case Is = win
            intItemCount = 7
        Case Is = loss
            intItemCount = 5
        Case Is = rage
            intItemCount = 6
        Case Is = beg
            intItemCount = 4
        Case Is = big_win
            intItemCount = 1
    End Select
    
    GetResponseCount = intItemCount
End Function

' -----------------------------------------------------------------
' ------------------------ FUNCTIONS ------------------------------
' -----------------------------------------------------------------

Private Sub ClearBossManPortrairs()
    ' *** This procedure is responsible for clearing all the Bossman Portrairs in page. ***
    
    If strPortrairSignature = "" Then
        Call InitializeModule(1, 1)
    End If
    
    mUtils.DeleteShapeWithSignature (strPortrairSignature)
End Sub

Private Sub InitializeModule(ByVal enumResponse As eBSJResponse, ByVal intResponseItem As Integer)
    ' *** This function is responsible for initializing all constans related to module operations ***
    
    Dim strBasePath As String
    
    strResponseName = GetResponseName(enumResponse)
    intResponseItem = intResponseItem
    intTotalResponseCount = GetResponseCount(enumResponse)
    strBasePath = mUtils.GetPathToRootDir()
    strAudioPath = strBasePath & "Audio\" & strResponseName & intResponseItem & ".wav"
    strFramesPath = strBasePath & "Frames\" & strResponseName & intResponseItem & "\frame_"
End Sub

Private Sub PlayBossmanResponse(ByVal enumResponse As eBSJResponse, ByVal intResponseItem As Integer)
    ' ***
    ' This procedure is responsible for playing whole the bossmanjack response himself.
    ' It requires the module to be initialized! So do not play on it's own cuz it won't work
    ' ***
    
    ' Accepts:
    '   - enumResponse [enum] - the chosen response to play
    '   - intResponseItem [integer] - the number of matching response that will be played.
    ' Returns:
    '   - None
    
    Dim dblAudioDurationInSeconds As Double
    Dim dblDurationFrameInSeconds As Double
    Dim startTime As Double
    Dim nextFrameStartTime As Double
    Dim intFrames As Integer
    Dim i As Integer
    
    dblAudioDurationInSeconds = GetWAVDurationInSeconds(strAudioPath)
    
    ' Calculate ammount of frames in dir
    intFrames = 0
    Do While mUtils.IsFileExists(strFramesPath & intFrames & ".jpg")
        intFrames = intFrames + 1
    Loop
    
    dblDurationFrameInSeconds = dblAudioDurationInSeconds / intFrames
    startTime = Timer
    
    On Error GoTo stopAnimation
    Call mAudioFiles.PlaySoundAsynchronymusly(strAudioPath)
    For i = 0 To intFrames - 1
        nextFrameStartTime = startTime + (i * dblDurationFrameInSeconds)
        
        ' Halt execution
        Do While Timer <= nextFrameStartTime
            DoEvents
        Loop
        
        ' Place frame
        Call mUtils.PlaceImage(rngPortrairRange, strFramesPath & i & ".jpg", strPortrairSignature & i)
        ' Remove previous frame
        If i >= 1 Then
             ActiveSheet.Shapes(strPortrairSignature & i - 1).Delete
        End If
    Next i
    
    Call DebugRun(i, dblDurationFrameInSeconds)
stopAnimation:
End Sub

Private Sub DebugRun(ByVal intTotalFrames As Integer, ByVal dblAnimDurationTime As Double)
    ' *** Debugs the current response.
    ' For Development purpouses only!
    ' Agree, it is provisoric as hell
    ' ***
    
    If boolDoDebug Then
        Debug.Print "---BossManResponse ---"
        Debug.Print "Name: " & strResponseName
        Debug.Print "Num: " & intResponseItem
        Debug.Print "Total Frames: " & intTotalFrames
        Debug.Print "AnimationTime (s): " & dblAnimDurationTime
        Debug.Print "Count Type Responses: " & intTotalResponseCount
        Debug.Print "---END BossManResponse----"
    End If
End Sub

' -----------------------------------------------------------------
' ------------------------ FUNCTIONS TO CALL ----------------------
' -----------------------------------------------------------------

Public Sub CallBossmanResponse(ByVal strFunctionName As String, strPortrairRange As String)
    ' ***
    ' Use this function whenever calling BSJ responses from the callbacks.
    ' ***
    
    ' Accepts:
    '   - strFunctionName [string] - the name of the function that will be called
    '   - strPortrairRange [string] - the range of the portrair where bossman will be shown.
    ' Returns:
    '   - None
    
    boolIsBossManResponding = True
    strSourceSheetName = ActiveSheet.Name
    
    Set rngPortrairRange = ActiveSheet.Range(strPortrairRange)
    Call Application.Run(strFunctionName)
    
    boolIsBossManResponding = False
    strSourceSheetName = ""
End Sub

Public Sub ShowBossmanResponse(ByVal enumResponse As eBSJResponse, Optional ByVal intResponseNum As Integer = 0)
    ' *** Entry point, call this function from games. ***
    
    ' Accepts:
    '   - enumResponse [enum] - the response which Bossman shall give
    '   - intResponseNum [integer, optional] - the number of response that will be played. If 0, then it would be random.
    ' Returns:
    '   - None
    
    Dim intItemsCount As Integer
    
    If rngPortrairRange Is Nothing Then
        Set rngPortrairRange = ActiveSheet.Range("P1:R11")
    End If
    
    If intResponseNum = 0 Then
        intItemsCount = GetResponseCount(enumResponse)
        intResponseNum = WorksheetFunction.RandBetween(1, intItemsCount)
    End If
    
    Call InitializeModule(enumResponse, intResponseNum)
    Call ClearBossManPortrairs
    Call PlayBossmanResponse(enumResponse, intResponseNum)
End Sub

Public Sub UnstuckBossmanResponse()
    ' ***
    ' Unstucks the response of the bossman.
    ' Might potentially be bugging out.
    ' ***
    
    If boolIsBossManResponding Then
        boolIsBossManResponding = False
    End If
End Sub

