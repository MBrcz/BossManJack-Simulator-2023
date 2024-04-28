VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' -------------------------------------------------------------------------------
' This module holds data related to the inner workings of the game known as kino.
' -------------------------------------------------------------------------------

Public boolDoDebug As Boolean

' Board Const
Private Const intBoardRows = 5
Private Const intBoardColumns = 8
Private Const intBoardCellWidth = 1
Private Const intBoardCellHeight = 6 ' Yep, it's bs but it works. DO NOT TOUCH!
Private Const intYBoardStart = 5
Private Const intXBoardStart = 5

' LPanel Const
Private Const intCellsCount = 10
Private Const intCellsXStart = 35
Private Const intCellsYStart = 5
Private Const intCellsOffset = 3

' Game const
Private Const intTargetBoardCellsCount = 10
Private Const strGemImageName = "img_Gem"
Private Const strGemPathJPG = "GameObjects\Kino\Gem.jpg"
Private Const strPortrairRange = "N1:P11"

' Logic related CONST
Private Const intGemBigWinCount = 6
Private Const intRedrewChance = 20
Private Const intGemNotLossCount = 4
Private Const intDontationAcceptanceThreshold = 5
Private Const strWinTable = "0, 0, 0, 0, 3.5, 8, 13, 63, 500, 800, 1000"

' Account related constans
Private Const strBalanceAccountRange = "H1"
Private Const strWagedMoneyRange = "B9"

' Variables
Private arrBoardCells() As tBoardCell
Private arrPanelLowerCells() As tLowerPanelCell

Private Enum eColors
    ' Primitive, ad hoc implementation of colors just for the purpouses of this game.
    eBlack = 1
    eViolet = 2
    eGreen = 3
    eGrey = 4
    eLPanelGrey = 5
    ' eLPanelGreen = 6
End Enum

Private Enum eCellState
    ' Flag for tracking target and non-target cells.
    eTarget = 1    ' This means that this cell should be hitted in order to get multiplier
    eNormal = 2    ' This means that this cell should not be hitted in order to get multiplier
End Enum

Private Type tBoardCell
    intNum As Integer            ' The number that cell represents
    strCellAdress As String      ' The address of a cell in question
    eState As eCellState         ' State of a cell [Target] or [Normal]
    eInteriorColor As eColors    ' Color currently applied to the cell
End Type

Private Type tLowerPanelCell
    strMultAddress As String     ' The address of the multiplier cell.
    strCountAddress As String    ' The address of the green counts cell.
    dblMult As Double            ' The multiplier value assigned to the cell.
    intCount As Integer          ' The integer that indices what element is it in cells..
    eCurrentColor As eColors     ' The current colors that both cells outght to have.
End Type

Private Function GetRBGArrayColors(ByVal enumColors As eColors) As Variant()
    ' ***
    ' Simple implementation of getting the colors by RGB array.
    ' Only for the purpouse of this mini-game.
    ' ***
    
    ' Accepts:
    '   - enumColors [enum] - the color name.
    ' Returns:
    '   - array - RGB pallete of defined colors.
    
    Dim arrRBG() As Variant
    
    Select Case enumColors
        Case Is = eBlack
            arrRBG = Array(0, 0, 0)
        Case Is = eViolet
            arrRBG = Array(139, 0, 139)
        Case Is = eGreen
            arrRBG = Array(0, 128, 0)
        Case Is = eGrey
            arrRBG = Array(100, 100, 100)
        Case Is = eLPanelGrey
            arrRBG = Array(64, 64, 64)
       '  Case Is = eLPanelGreen
       '      arrRBG = Array(0, 128, 128)
    End Select
    
    GetRBGArrayColors = arrRBG
    
End Function

' ---------------------------------------------------------------------------------------------
' -------------------------------------------- GAME -------------------------------------------
' ---------------------------------------------------------------------------------------------

Private Sub InitializeBoard()
    ' *** This function is responsible for initializing FLAT the board for the game. ***
    
    Dim x As Byte
    Dim y As Byte
    Dim bytCount As Byte
    
    ReDim arrBoardCells(1 To intBoardRows, 1 To intBoardColumns)
    
    bytCount = 1
    For x = 1 To intBoardRows
        For y = 1 To intBoardColumns
            Call SetDefaultBoardCell(x, y, bytCount)
            bytCount = bytCount + 1
        Next y
    Next x

End Sub

Private Sub SetDefaultBoardCell(ByVal bytXPosition As Byte, bytYPosition As Byte, ByVal bytBoardCellNum As Byte)
    ' *** This function is responsible for setting the cell to a default parameters. ***
    
    ' Accepts:
    '   - bytXPosition [byte] - the position in x (row) axis of a cell.
    '   - bytYPosition [byte] - the position in y (column) axis of a cell.
    '   - bytBoardCellNum [byte] - the number assigned to the cell.
    ' Returns:
    '   - None
    
    Dim bytBoardCellXPos As Byte
    Dim bytBoardCellYPos As Byte
    Dim typCell As tBoardCell
   
    bytBoardCellXPos = intXBoardStart + (bytXPosition - 1) * intBoardCellHeight
    bytBoardCellYPos = bytYPosition * intBoardCellWidth + intYBoardStart
    
    With typCell
        .intNum = bytBoardCellNum
        .strCellAdress = Cells(bytBoardCellXPos, bytBoardCellYPos).MergeArea.Address
        .eInteriorColor = eBlack
        .eState = eNormal
    End With
    
    arrBoardCells(bytXPosition, bytYPosition) = typCell
End Sub

Private Sub UpdateBoardCells()
    ' *** This procedure is responsible for updating the cells in the page. ***
    
    Dim x As Byte
    Dim y As Byte
    
    For x = 1 To intBoardRows
        For y = 1 To intBoardColumns
            Call UpdateTheCell(x, y)
        Next y
    Next x
End Sub

Private Sub UpdateTheCell(ByVal bytXPosition As Byte, bytYPosition As Byte)
    ' *** Updates the shown cell in the project to according to the cell object. ***
    
    Dim typCell As tBoardCell
    Dim arrChosenColor() As Variant
    Dim strCellRange As String
    
    typCell = arrBoardCells(bytXPosition, bytYPosition)
    strCellRange = typCell.strCellAdress
    arrChosenColor = GetRBGArrayColors(typCell.eInteriorColor)
    
    With ActiveSheet.Range(strCellRange)
        .Value = typCell.intNum
        .Interior.Color = RGB(arrChosenColor(0), arrChosenColor(1), arrChosenColor(2))
    End With
End Sub

Private Sub DrawTargetBoardCells()
    ' *** Draws and marks the cells that will be needed to catch in order to get money. ***
    
    Dim bytRndXPOS As Byte
    Dim bytRndYPOS As Byte
    Dim i As Byte
    
    i = 1
    Do While i <= intTargetBoardCellsCount
        bytRndXPOS = Application.WorksheetFunction.RandBetween(1, intBoardRows)
        bytRndYPOS = Application.WorksheetFunction.RandBetween(1, intBoardColumns)
        
        If arrBoardCells(bytRndXPOS, bytRndYPOS).eInteriorColor = eBlack Then
            arrBoardCells(bytRndXPOS, bytRndYPOS).eInteriorColor = eViolet
            arrBoardCells(bytRndXPOS, bytRndYPOS).eState = eTarget
            i = i + 1
        End If
    Loop
End Sub

Private Sub DrawRandomCells()
    ' *** Draws the random cells that are selected by the game. ***
    
    Dim bytRndXPOS As Byte
    Dim bytRndYPOS As Byte
    Dim i As Byte
        
    For i = 1 To intTargetBoardCellsCount
        ' Find not occupied pieces.
        Do While True
            bytRndXPOS = Application.WorksheetFunction.RandBetween(1, intBoardRows)
            bytRndYPOS = Application.WorksheetFunction.RandBetween(1, intBoardColumns)
            
            If Not (arrBoardCells(bytRndXPOS, bytRndYPOS).eInteriorColor = eGreen Or _
                    arrBoardCells(bytRndXPOS, bytRndYPOS).eInteriorColor = eGrey) Then
                Exit Do
            End If
        Loop
        
        ' Place the found stuff.
        If arrBoardCells(bytRndXPOS, bytRndYPOS).eState = eTarget Then
            ' Skew game a little
            If Rnd() * 100 >= intRedrewChance Then
                Call PlaceImageOfGemOnTopOfCell(bytRndXPOS, bytRndYPOS)
                arrBoardCells(bytRndXPOS, bytRndYPOS).eInteriorColor = eGreen
                Call UpdateLowerPanelDurningIteration
            Else
                i = i - 1
            End If
        
        ElseIf arrBoardCells(bytRndXPOS, bytRndYPOS).eState = eNormal Then
            arrBoardCells(bytRndXPOS, bytRndYPOS).eInteriorColor = eGrey
        End If
    Next i
        
    Call UpdateLowerPanelDurningIteration ' Makes sure that whenever none green has been found, something is selected.
End Sub

Private Sub PlaceImageOfGemOnTopOfCell(ByVal bytXPOS As Byte, bytYPOS As Byte)
    ' *** Places an image of the gem on top of the cell. ***
    
    Dim strPathToImage As String
    Dim rngImagePlace As Range
    
    Set rngImagePlace = Application.Range(arrBoardCells(bytXPOS, bytYPOS).strCellAdress)
    strPathToImage = mUtils.GetPathToRootDir() & strGemPathJPG
    
    Call mUtils.PlaceImage( _
                rngImageRange:=rngImagePlace, _
                strPath:=strPathToImage, _
                newImageName:=strGemImageName _
    )
End Sub

Private Sub RemoveGemImages()
    ' *** Remove all gems images in the page. ***
    
    mUtils.DeleteShapeWithSignature (strGemImageName)
End Sub

Private Sub RestoreTheBoardCellsColors()
    ' *** Restores cells colors to either violet or black. ***
    
    Dim x As Byte
    Dim y As Byte
    
    For x = 1 To intBoardRows
        For y = 1 To intBoardColumns
            Select Case arrBoardCells(x, y).eInteriorColor
                Case Is = eGreen
                    arrBoardCells(x, y).eInteriorColor = eViolet
                Case Is = eGrey
                    arrBoardCells(x, y).eInteriorColor = eBlack
            End Select
        Next y
    Next x
End Sub

Private Function CalculateBoardCellsWithColor(ByRef enumColor As eColors) As Integer
    ' *** Counts how many there are cells with a specified color. ***
    
    Dim x As Byte
    Dim y As Byte
    Dim intCount As Integer
    
    intCount = 0
    For x = 1 To intBoardRows
        For y = 1 To intBoardColumns
            If arrBoardCells(x, y).eInteriorColor = enumColor Then
                intCount = intCount + 1
            End If
        Next y
    Next x
    
    CalculateBoardCellsWithColor = intCount
End Function

Private Function IsBoardInitialized() As Boolean
    ' ***
    ' Checks if the board is initialized or not.
    ' The measure of if the board is initialized is if there are any violet cells inside board.
    ' ***
    
    Dim intVioletCells As Integer
    
    On Error Resume Next
    intVioletCells = CalculateBoardCellsWithColor(eViolet)
    
    If Err.Number <> 0 Or intVioletCells < 1 Then
        IsBoardInitialized = False
    Else
        IsBoardInitialized = True
    End If
End Function

' ---------------------------------------------------------------------------------
' ----------------------------- PANEL AT THE BOTTOM -------------------------------
' ---------------------------------------------------------------------------------

Private Sub InitializeLowerPanel()
    ' *** Handles the initialization of the Lower Panel and it's shenenigans. ***
    
    Dim typLPanelCell As tLowerPanelCell
    Dim arrWinTable As Variant
    Dim x As Byte
    
    ReDim arrPanelLowerCells(0 To intCellsCount)
    arrWinTable = Split(strWinTable, ", ")
    
    For x = 0 To intCellsCount
        With typLPanelCell
            .eCurrentColor = eLPanelGrey
            .strCountAddress = Cells(intCellsXStart + intCellsOffset, x + intCellsYStart).MergeArea.Address
            .strMultAddress = Cells(intCellsXStart, x + intCellsYStart).MergeArea.Address
            .dblMult = arrWinTable(x)
            .intCount = x
        End With
        
        arrPanelLowerCells(x) = typLPanelCell
    Next x
End Sub

Private Sub UpdateLowerPanelDurningIteration()
    ' *** Updates the colors that are shown in the panel durning the iteration of board. ***
    
    Dim intHittedGems As Integer
    
    On Error Resume Next
    intHittedGems = CalculateBoardCellsWithColor(eGreen)
    arrPanelLowerCells(intHittedGems).eCurrentColor = eGreen
    arrPanelLowerCells(intHittedGems - 1).eCurrentColor = eLPanelGrey
    
    ActiveSheet.Shapes("Gem" & intHittedGems).Line.Visible = True
    ActiveSheet.Shapes("Gem" & intHittedGems - 1).Line.Visible = False
    
    Call UpdateColorsOfLowerPanel
End Sub

Private Sub RestoreTheLowerPanelColors()
    ' *** Restores the lower panels colors to the default. ***

    Dim x As Byte

    For x = 0 To intCellsCount
        arrPanelLowerCells(x).eCurrentColor = eLPanelGrey
    Next x
End Sub

Private Sub HideTheLastGemLine()
    ' *** Hides the line from the last selected shape. ***
    
    On Error Resume Next
    ActiveSheet.Shapes("Gem" & CalculateBoardCellsWithColor(eGreen)).Line.Visible = False
End Sub

Private Sub UpdateColorsOfLowerPanel()
    ' *** Updates all the colors assigned to the lower panels. ***
    
    Dim arrColor() As Variant
    Dim strMultRange As String
    Dim strCountRange As String
    Dim x As Byte
    
    For x = 0 To intCellsCount
        strMultRange = arrPanelLowerCells(x).strCountAddress
        strCountRange = arrPanelLowerCells(x).strMultAddress
        arrColor = GetRBGArrayColors(arrPanelLowerCells(x).eCurrentColor)
        
        Range(strMultRange).Interior.Color = RGB(arrColor(0), arrColor(1), arrColor(2))
        Range(strCountRange).Interior.Color = RGB(arrColor(0), arrColor(1), arrColor(2))
    Next x
End Sub

' ---------------------------------------------------------------------------------
' --------------------------------- OTHER -----------------------------------------
' ---------------------------------------------------------------------------------

Private Sub CashOut()
    ' *** Handles the operation of cashing out from game. ***
    
    Dim arrWinTable As Variant
    Dim dblAmmountWon As Double
    Dim intGemsCount As Integer
    
    arrWinTable = Split(strWinTable, ", ")
    intGemsCount = CalculateBoardCellsWithColor(eGreen)
    dblAmmountWon = arrWinTable(intGemsCount) * mBalanceManager.WagedMoney
    
    mBalanceManager.UpdateAccountAfterGambling (dblAmmountWon)
    If intGemsCount >= intGemNotLossCount Then
        mMessageBoxes.msgInfoKino_Win (dblAmmountWon)
    End If
End Sub

Private Sub UpdateAccount()
    ' *** Updates the account state **
    
    Dim rngBalanceAccount As Range
    Set rngBalanceAccount = ActiveSheet.Range(strBalanceAccountRange)
    
    rngBalanceAccount.Value = mBalanceManager.AccountBalance
End Sub

Private Sub UpdateWagedMoney()
    ' *** Updates the waged money state ***
    
    Dim rngWagedMoney As Range
    Set rngWagedMoney = ActiveSheet.Range(strWagedMoneyRange)
    
    rngWagedMoney.Value = mBalanceManager.WagedMoney
End Sub

Private Sub SetWagedMoney()
    ' *** Sets the waged money ***
    
    Dim rngWagedMoney As Range
    
    Set rngWagedMoney = ActiveSheet.Range(strWagedMoneyRange)
    mBalanceManager.WagedMoney = rngWagedMoney.Value
End Sub

Private Function IsPlayingPossible() As Boolean
    ' *** Checks if playing Kino is possible right now. ***
    
    If mBSJResponse.IsBossmanResponding Then
        mMessageBoxes.msgCritical_BSJSpeaks
        Exit Function
    ElseIf Not IsBoardInitialized() Then
        mMessageBoxes.msgCritical_KinoBoardNotInstanciated
        Exit Function
    ElseIf mBalanceManager.AccountBalance < 0.01 Then
        mMessageBoxes.msgCritical_NoneFunds
        Exit Function
    End If
    
    IsPlayingPossible = True
End Function

Public Sub HandleBossManReactions(Optional ByVal varProxy As Variant = 0)
    ' *** This procedure is responsible for handling the reactions of the BSJ to the game. ***
    
    Dim intShootedGems As Integer
    
    intShootedGems = CalculateBoardCellsWithColor(eGreen)
    
    If intGemBigWinCount <= intShootedGems Then
        Call mBSJResponse.ShowBossmanResponse(win)
        Debug.Print "NOT IMPLEMENTED"
    ElseIf intShootedGems >= intGemNotLossCount Then
        Call mBSJResponse.ShowBossmanResponse(win)
    ElseIf intShootedGems < intGemNotLossCount And mBalanceManager.AccountBalance <= intDontationAcceptanceThreshold Then
        Call mBSJResponse.ShowBossmanResponse(rage)
    ElseIf intShootedGems < intGemNotLossCount Then
        Call mBSJResponse.ShowBossmanResponse(loss)
    End If
End Sub

Public Sub UpdateWallet(Optional ByVal varProxy As Variant = 0)
    ' *** Handles operation of updating a wallet of the BossMan Jack. ***
    
    If mBalanceManager.AccountBalance < intDontationAcceptanceThreshold Then
        Call mBalanceManager.AcceptJuicerFromRat
        Call UpdateAccount
        Call UpdateWagedMoney
        Call InitializeBoard
        Call UpdateBoardCells
        Call mBSJResponse.ShowBossmanResponse(beg)
    Else
        MsgBox "You have enough money, Bossman. No one will buy your bs till you're empty!", vbInformation
    End If
End Sub

Private Sub DebugGame():
    ' *** Debugs the game. ***
    
    Dim arrWinTable As Variant
    
    arrWinTable = Split(strWinTable, ", ")
    If boolDoDebug Then
        Debug.Print "--- KINO GAME ---"
        Debug.Print "BOARD: count of basic (black) cells:   " & CalculateBoardCellsWithColor(eBlack)
        Debug.Print "BOARD: count of target (violet) cells: " & CalculateBoardCellsWithColor(eViolet) + CalculateBoardCellsWithColor(eGreen)
        Debug.Print "BOARD: count of hitted (green) cells:  " & CalculateBoardCellsWithColor(eGreen)
        Debug.Print "BOARD: count of missed (grey) cells:   " & CalculateBoardCellsWithColor(eGrey)
        Debug.Print "LOGIC: waged money:                    " & mBalanceManager.WagedMoney
        Debug.Print "LOGIC: gained money:                   " & arrWinTable(CalculateBoardCellsWithColor(eGreen)) * mBalanceManager.WagedMoney
        Debug.Print "LOGIC: account state at the end        " & mBalanceManager.AccountBalance + arrWinTable(CalculateBoardCellsWithColor(eGreen)) * mBalanceManager.WagedMoney
        Debug.Print "LOGIC: used multiplier:                " & arrWinTable(CalculateBoardCellsWithColor(eGreen))
        Debug.Print "--- END KINO ---"
    End If
End Sub

' ---------------------------------------------------------------------------------
' -------------------------------- CALLBACKS --------------------------------------
' ---------------------------------------------------------------------------------

Public Sub DoubleWagedMoney_click()
    ' *** Doubles the money that will be waged. ***
    
    If mBSJResponse.IsBossmanResponding Then
        Exit Sub
    End If
    
    Call SetWagedMoney
    Call mBalanceManager.DoubleTheGamblingMoney
    Call UpdateWagedMoney
End Sub

Public Sub HalfWagedMoney_click()
    ' *** Halves the money that will be waged. ***
    
    If mBSJResponse.IsBossmanResponding Then
        Exit Sub
    End If

    Call SetWagedMoney
    Call mBalanceManager.HalfTheGamblingMoney
    Call UpdateWagedMoney
End Sub

Public Sub Wallet_click()
    ' *** Handles operations on the wallet. ***
    
    If mBSJResponse.IsBossmanResponding Then
        Exit Sub
    End If
    
    Call mBSJResponse.CallBossmanResponse("Sheet3.UpdateWallet", strPortrairRange)
End Sub

Public Sub SetupBoard_click()
    ' *** Initializes the board for KINO Game. ***
    
    If mBSJResponse.IsBossmanResponding Then
        Exit Sub
    End If
    
    Call InitializeBoard
    Call InitializeLowerPanel
    Call DrawTargetBoardCells
    Call UpdateBoardCells
End Sub

Public Sub DonateToSlake_click()
    ' *** Handles donating money to the slake.com ***
    
    Call SetWagedMoney
    If Not IsPlayingPossible Then
        Exit Sub
    End If
    
    mBalanceManager.UpdateAccountBeforeGambling (-mBalanceManager.WagedMoney)
    Call DrawRandomCells
    Call UpdateBoardCells
    Call DebugGame
    Call CashOut
    Call UpdateAccount
    Call UpdateWagedMoney
    
    Call mBSJResponse.CallBossmanResponse("Sheet3.HandleBossmanReactions", strPortrairRange)
    Call HideTheLastGemLine
    Call RestoreTheBoardCellsColors
    Call RestoreTheLowerPanelColors
    Call UpdateColorsOfLowerPanel
    Call UpdateBoardCells
    Call RemoveGemImages
End Sub

Private Sub Worksheet_Activate()
    
    If mBSJResponse.IsBossmanResponding Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Call UpdateAccount
    Call UpdateWagedMoney
    Call InitializeBoard
    Call UpdateBoardCells
    Application.ScreenUpdating = True
End Sub
