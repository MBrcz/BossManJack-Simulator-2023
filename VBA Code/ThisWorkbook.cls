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
' *** Libraries for importing an excel file ***
#If VBA7 Then ' Excel 2010 or later
    Private Declare PtrSafe Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As LongPtr) As Long
#Else ' Excel 2007 or earlier
    Private Declare Function sndPlaySound32 Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
#End If

Public Enum eBossManResponseTypes
    win = 1
    loss = 2
    beg = 3
    rage = 4
End Enum

Private Function GetResponseName(ByVal enumBossmanResponse As eBossManResponseTypes) As String
    
    Select Case enumBossmanResponse
        Case Is = win
            GetResponseName = "Win"
        Case Is = loss
            GetResponseName = "Loss"
        Case Is = beg
            GetResponseName = "Beg"
        Case Is = rage
            GetResponseName = "Rage"
    End Select

End Function

Private Function GetResponseTypeQuantity(ByVal enumBossmanResponse As eBossManResponseTypes)
    
    Select Case enumBossmanResponse
        Case Is = win
            GetResponseTypeQuantity = 2
        Case Is = loss
            GetResponseTypeQuantity = 3
        Case Is = beg
            GetResponseTypeQuantity = 2
        Case Is = rage
            GetResponseTypeQuantity = 3
    End Select
    
End Function

Private Function GetResponseRepetition(ByVal enumBossmanResponse As eBossManResponseTypes, ByVal intRepetition As Integer) As Integer
    ' *** Gets the repetition assigned to each respose. ***
    
    Select Case enumBossmanResponse
        Case Is = win
        
            Select Case intRepetition
                Case Is = 1
                    GetResponseRepetition = 3
                Case Is = 2
                    GetResponseRepetition = 4
                Case Else
                    Debug.Print "WARNING! NO SUCH REPETITION FOR WIN!"
                    GetResponseRepetition = 1
            End Select
            
        Case Is = loss
            Select Case intRepetition
                Case Is = 1
                    GetResponseRepetition = 4
                Case Is = 2
                    GetResponseRepetition = 3
                Case Is = 3
                    GetResponseRepetition = 4
                Case Else
                    Debug.Print "WARNING! NO SUCH REPETITION FOR LOSS!"
                    GetResponseRepetition = 1
            End Select
            
        Case Is = beg
            Select Case intRepetition
                Case Is = 1
                    GetResponseRepetition = 3
                Case Is = 2
                    GetResponseRepetition = 4
                Case Else
                    Debug.Print "WARNING! NO SUCH REPETITION FOR BEG!"
                    GetResponseRepetition = 1
            End Select
        
        Case Is = rage
            Select Case intRepetition
                Case Is = 1
                    GetResponseRepetition = 3
                Case Is = 2
                    GetResponseRepetition = 4
                Case Is = 3
                    GetResponseRepetition = 5
                Case Else
                    Debug.Print "WARNING! NO SUCH REPETITION FOR RAGE!"
            End Select
                
    End Select
    
End Function

Private Function GetBSJPortrairSignature() As String
    GetBSJPortrairSignature = "BSJ"
End Function

Private Function GetProjectRootPath() As String
    GetProjectRootPath = Replace(ThisWorkbook.FullName, "\BossManJack Simulator 2023.xlsm", "")
End Function

Private Function GetBasePathToAudio() As String
    GetBasePathToAudio = GetProjectRootPath() & "\Audio"
End Function

Private Function GetBasePathToImages() As String
    GetBasePathToImages = GetProjectRootPath() & "\Frames"
End Function

' ----------------------------------------------------------------
' ------------------------ DEALING WITH PATHS --------------------
' ----------------------------------------------------------------

Private Function DrawBossmanResponse(ByVal enumBossmanResponse As eBossManResponseTypes) As Integer
    
    Dim intResponsesCount As Integer
    Dim intResponse As Integer
    
    intResponsesCount = GetResponseTypeQuantity(enumBossmanResponse)
    intResponse = WorksheetFunction.RandBetween(1, intResponsesCount)
    DrawBossmanResponse = intResponse
End Function

Private Function GetAudioFile(ByVal enumBossmanResponse As eBossManResponseTypes, ByVal intResponseNum As Integer)
    GetAudioFile = GetBasePathToAudio & "\" & GetResponseName(enumBossmanResponse) & intResponseNum & ".wav"
End Function

Private Function GetFramesFiles(ByVal enumBossmanResponse As eBossManResponseTypes, ByVal intResponseNum As Integer) As Variant()
    
    Dim strPathToFile As String
    Dim arrFramesFiles() As Variant
    Dim i As Integer
    
    strPathToFile = GetBasePathToImages & "\" & GetResponseName(enumBossmanResponse) & intResponseNum
    arrFramesFiles = GetAllFilesFromDirectoryToArray(strPathToFile)
    
    i = 0
    Do While i < UBound(arrFramesFiles)
        arrFramesFiles(i) = strPathToFile & "\frame_" & i & ".jpg"
        i = i + 1
    Loop
    
    GetFramesFiles = arrFramesFiles
End Function

' ----------------------------------------------------------------
' ------------------------ PORTRAIR METHODS ----------------------
' ----------------------------------------------------------------

Private Sub ClearPortrairsFromPage()
    
    Dim strBsjSignature As String
    Dim shpShape As Shape
    
    strBsjSignature = GetBSJPortrairSignature()
    
    For Each shpShape In ActiveSheet.Shapes
        If InStr(1, shpShape.Name, strBsjSignature, vbTextCompare) Then
            shpShape.Delete
        End If
    Next shpShape
End Sub

Private Sub RepeatProcessingImages(ByVal strPathToImage As String, ByVal intRepetition As Integer)
    ' *** intRepetition is kinda hacky way of synhronizing images with audio. Really hacky. ***
    
    Dim rngPortrairRange As Range
    Dim intCount As Integer
    
    Set rngPortrairRange = ActiveSheet.Range("P1:R11") ' Will solve it in the future, if necessary.
    
    intCount = 0
    Do While intCount < intRepetition
        Call PlaceImage(rngPortrairRange, strPathToImage, "BSJ" & intCount)
        intCount = intCount + 1
    Loop
    
End Sub

' ----------------------------------------------------------------------------------------
' ---------------------------- RESPONSE HANDLER ------------------------------------------
' ----------------------------------------------------------------------------------------

Public Sub PerformBossManResponse(ByVal enumBossmanResponse As eBossManResponseTypes, Optional ByVal intResponseNum As Integer = 0)
    
    Call ClearPortrairsFromPage
    Call ShowBossmanResponse(enumBossmanResponse, intResponseNum)
End Sub

Private Sub ShowBossmanResponse(ByVal enumBossmanResponse As eBossManResponseTypes, Optional ByVal intResponseNum As Integer = 0)
    
    Dim strAudioFilePath As String
    Dim arrPortrairImages() As Variant
    Dim intRepetitions As Integer
       
    If intResponseNum = 0 Then
        intResponseNum = DrawBossmanResponse(enumBossmanResponse)
    End If
    
    Debug.Print "Chosen response: " & GetResponseName(enumBossmanResponse) & ": " & intResponseNum
    intRepetitions = GetResponseRepetition(enumBossmanResponse, intResponseNum)
    arrPortrairImages = GetFramesFiles(enumBossmanResponse, intResponseNum)
    strAudioFilePath = GetAudioFile(enumBossmanResponse, intResponseNum)
    
    Call RunResponse(strAudioFilePath, arrPortrairImages, intRepetitions)
    
End Sub

Private Sub RunResponse(ByVal strPathToAudio As String, ByRef arrPortrairImages() As Variant, ByVal intRepetitions As Integer)
    
    Dim i As Integer
    Dim boolIsAudioRun As Boolean
    
    i = 0
    Do While i < UBound(arrPortrairImages)
        Call RepeatProcessingImages(arrPortrairImages(i), intRepetitions)
        DoEvents
        If Not boolIsAudioRun Then
            boolIsAudioRun = True
            Call sndPlaySound32(strPathToAudio, 3)
        ElseIf boolIsAudioRun Then
            Application.Wait Now() + TimeValue("00:00:01") / 2
        End If
        
        i = i + 1
    Loop
    
End Sub

Public Sub PlaceImage(ByVal rngImageRange As Range, ByVal strPath As String, _
                      ByVal newImageName As String)
    ' ***
    '   Places Images In The Current Worksheet in the chosen range and sets the size according to range in question.
    '   Works only for images that have set arbitrary path.
    ' ***
    
    ' Accepts:
    '   - rngImageRange [range] - the range where image shall be placed
    '   - strPath [string] - the name of image that will be loaded.
    '   - newImageName [string] - the name of the new image that will be placed.
    ' Returns:
    '   - None
    Dim shpImg As Shape
    
    On Error Resume Next
    Set shpImg = ActiveSheet.Shapes.AddPicture(strPath, msoFalse, msoCTrue, 1, 1, 1, 1)
    Call PlaceImageInTheRange(shpImg, rngImageRange, newImageName)
    
End Sub

Private Sub PlaceImageInTheRange(ByVal shpImage As Shape, ByVal rngTarget As Range, ByVal strNewName As String)
    ' *** Places the image on the chosen range ***
    
    ' Accepts:
    '   - shpImage [Shape] - the image that shall be moved,
    '   - rngTarget [range] - the place where image will be copied to
    '   - strNewName [string] - the new name of the copied image.
    
    With shpImage
        .LockAspectRatio = msoFalse
        .Top = rngTarget.Top
        .Left = rngTarget.Left
        .Width = rngTarget.Width
        .Height = rngTarget.Height
        .Name = strNewName
    End With

End Sub

Public Function GetAllFilesFromDirectoryToArray(ByVal strPath As String, Optional boolIsFullPath As Boolean = False) As Variant()
    ' *** Retrieves all files from the specified directory matching the given extensions ***
    
    ' Accepts:
    '   - strPath [String] - The path of the directory to search.
    '   - boolIsFullPath [string] - True means that all files will be got with prefix of
    '                               current path, otherwise only names of files.
    
    ' Returns:
    '   - Array - An array containing the filenames matching the specified extensions.

    Dim oFso As Object
    Dim oFolder As Object
    Dim oFile As Object
    
    Dim arrFileNames() As Variant
    Dim varExtension As Variant
    Dim i As Integer
    
    Set oFso = CreateObject("Scripting.FileSystemObject")
    
    ' If folder does not exist there, exit function
    If Not oFso.folderexists(strPath) Then
        Exit Function
    End If
    
    Set oFolder = oFso.GetFolder(strPath)
    
    i = 0
    For Each oFile In oFolder.Files
        ReDim Preserve arrFileNames(i)
        If boolIsFullPath And Right(strPath, 1) = "\" Then
            arrFileNames(i) = strPath & oFile.Name
        
        ElseIf boolIsFullPath And Right(strPath, 1) <> "\" Then
            arrFileNames(i) = strPath & "\" & oFile.Name
            
        Else
            arrFileNames(i) = oFile.Name
            
        End If
        i = i + 1
    Next oFile
    
    GetAllFilesFromDirectoryToArray = arrFileNames

End Function

' DEVELOPMENT ONLY!
Public Sub DevShowBossResponse(ByVal enumeResponse As eBossManResponseTypes, intNum As Integer)
    Call ThisWorkbook.PerformBossManResponse(enumeResponse, intNum)
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call ClearPortrairsFromPage
    Range("F5").Value = ""
End Sub
