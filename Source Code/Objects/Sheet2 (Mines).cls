VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' ***
' This module contains logic related to the so called "Mines".
' It works just like minesweeper, but it is f***cking rigged ;x.
' ***

Public boolDoDebug As Boolean

' BOARD CONST
Private Const intBoardRow = 5
Private Const intBoardColumn = 5
Private Const intCellHeight = 2
Private Const intCellWidth = 7
Private Const intXBoardStart = 5
Private Const intYBoardStart = 6
Private Const strMineImageName = "img_Rat"
Private Const strGemImageName = "img_Gem"
Private Const strMineJPGName = "Rat.jpg"
Private Const strGemJPGName = "Gem.jpg"
Private Const strJPGSPath = "GameObjects\Mines\"

' GAME CONST
Private Const intMineChanceConst = 30
Private Const strMinesCountRange = "B17"
Private Const strPortrairRange = "P1:R11"

' BALANCE CONST
Private Const strBalanceAccountRange = "H1"
Private Const bytDonationAcceptThreshold = 5
Private Const strWagedMoneyRange = "B9"
Private Const dblWagedMoneyMultiplier = 1.05
Private Const shpDonateButtonName = "Rectangle 9"
Private Const strCashOutRange = "B13"

' VARIABLES
Private intXSelectedCell As Integer
Private intYSelectedCell As Integer
Private isGameBeingPlayed As Boolean
Private intGemsFound As Integer
Private intMinesFound As Integer
Private arrCells(1 To intBoardRow, 1 To intBoardColumn) As tCell ' Stores all cells in board as array.

' ------------------------------------------------------------------------------
' --------------------------------- GAME BASIC ---------------------------------
' ------------------------------------------------------------------------------

Private Enum eCellImageState
    ' Stores the states of the images.
    Mine = 1
    Gem = 2
    Clear = 3
End Enum

Private Type tCell
    strAddress As String            ' Adress of a cell in question (single).
    strMergedAddress As String      ' Adress of a cell in question (multiple).
    boolIsLocked As Boolean         ' State of a cell - Locked means that it is not possible to change it.
    enumImage As eCellImageState    ' The name of the image assigned to the cell; by default: None.
End Type

Private Sub InitializeBoard()
    ' *** This procedure is responsible for initializing the board for the project ***
    
    Dim i As Byte
    Dim y As Byte
      
    Call RefreshCurrentCellPosition
    For i = 1 To intBoardRow
        For y = 1 To intBoardColumn
            Call SetDefaultCell(i, y)
        Next y
    Next i
  
End Sub

Private Sub RefreshCurrentCellPosition()
    ' *** This function refreshes the current positions of the cell in game. ***
    
    intXSelectedCell = 0
    intYSelectedCell = 0
End Sub

Private Function IsSelectedCellInBoard(ByVal strSelectedCellAddress) As Boolean
    ' ***
    ' This functions checks if selected by user cell is part of the board.
    ' If True, it assigns the xPos and yPos to the said cell in array.
    ' ***
    
    ' Accepts:
    '   - strSelectedCellAddress [string] - the cell adress of the selected by user cell.
    
    ' Returns:
    '   - boolean - True means that selected cell is in board, otherwise False.
    
    Dim strCurrentCellAddress As String
    Dim y As Byte
    Dim i As Byte
    
    For i = 1 To intBoardRow
        For y = 1 To intBoardColumn
            strCurrentCellAddress = arrCells(i, y).strMergedAddress
            If strCurrentCellAddress = strSelectedCellAddress Then
                intXSelectedCell = i
                intYSelectedCell = y
                IsSelectedCellInBoard = True
                Exit Function
            End If
        Next y
    Next i
    
    IsSelectedCellInBoard = False

End Function

Private Sub SetDefaultCell(ByVal intXPos As Integer, ByVal intYPos As Integer)
    ' *** Sets value of a cell to a default - use durning initialization of the board. ***
    
    Dim intCellXPos As Integer
    Dim intCellYPos As Integer
    Dim typCell As tCell
    
    intCellXPos = (intXPos * intCellWidth) - 2
    intCellYPos = (intYPos * intCellHeight) + intYBoardStart - 2
    With typCell
        .strAddress = Cells(intCellXPos, intCellYPos).Address
        .strMergedAddress = Range(typCell.strAddress).MergeArea.Address
        .boolIsLocked = False
        .enumImage = Clear
    End With
    
    arrCells(intXPos, intYPos) = typCell
End Sub

Private Sub SelectACell()
    ' *** This function handles the special cell selecting operation. ***
    
    Dim boolIsCellAMine As Boolean
    
    boolIsCellAMine = IsSelectedCellAMine()
    
    ' If cell is locked, gtfo
    If arrCells(intXSelectedCell, intYSelectedCell).boolIsLocked = True Then
        Exit Sub
    End If
    
    arrCells(intXSelectedCell, intYSelectedCell).boolIsLocked = True
    
    If boolIsCellAMine Then
        arrCells(intXSelectedCell, intYSelectedCell).enumImage = Mine
        Call HandleMineFound
    ElseIf Not boolIsCellAMine Then
        arrCells(intXSelectedCell, intYSelectedCell).enumImage = Gem
        Call HandleGemFound
    End If
End Sub

Private Sub ChooseRandomCell()
    ' *** Selects a random cell from a pool of all available cells. ***
    
    Do While True
        intXSelectedCell = Application.WorksheetFunction.RandBetween(1, intBoardRow)
        intYSelectedCell = Application.WorksheetFunction.RandBetween(1, intBoardColumn)
        
        If arrCells(intXSelectedCell, intYSelectedCell).enumImage = Clear Then
            Application.EnableEvents = False
            Range(arrCells(intXSelectedCell, intYSelectedCell).strMergedAddress).Select
            Application.EnableEvents = True
            Exit Sub
        End If
    Loop
End Sub

Private Function IsSelectedCellAMine() As Boolean
    ' *** Checks if selected by user cell is a mine. ***
    
    Dim dblThreshold As Double
    Dim rndRandom As Double
    
    rndRandom = Rnd() * 100
    
    dblThreshold = intMineChanceConst + (intGemsFound * 10)
    
    If rndRandom >= dblThreshold Then
        IsSelectedCellAMine = False
    ElseIf rndRandom < dblThreshold Then
        IsSelectedCellAMine = True
    End If
End Function

Private Sub PlaceImagesInCells()
    ' *** Removes all images that are placed in the board. ***
    
    Dim objCell As tCell
    Dim i As Byte
    Dim j As Byte
    
    For i = 1 To intBoardRow
        For j = 1 To intBoardColumn
            objCell = arrCells(i, j)
            Call HandleInjectingImageProcedure(objCell)
        Next j
    Next i
    
End Sub

Private Sub HandleInjectingImageProcedure(ByRef objCell As tCell)
    ' *** Manages to insert the image into the cell based on the cell property. ***
    
    Dim rngImageCell As Range
    Dim strBasePath As String
    
    Set rngImageCell = ActiveSheet.Range(objCell.strMergedAddress)
    strBasePath = mUtils.GetPathToRootDir() & strJPGSPath
    
    Select Case objCell.enumImage
        Case Is = Mine
            Call mUtils.PlaceImage(rngImageCell, strBasePath & "\" & strMineJPGName, strMineImageName)
        Case Is = Gem
            Call mUtils.PlaceImage(rngImageCell, strBasePath & "\" & strGemJPGName, strGemImageName)
        Case Else
    End Select
    
End Sub

Private Sub RemoveAllImagesFromBoard()
    ' *** Removes all images from the board object. ***
    
    Call mUtils.DeleteShapeWithSignature(strGemImageName)
    Call mUtils.DeleteShapeWithSignature(strMineImageName)
End Sub

Private Sub FillsTheBoardWithMines()
    ' ***
    ' Fills the board with mines at the random positions.
    ' Yep, it's lazy as f**k, but it works.
    ' ***
    
    Dim bytMinesToFill As Byte
    Dim bytRndXPOS As Byte
    Dim bytRndYPOS As Byte
    Dim i As Byte
    
    i = 0
    Do While i < GetMinesCount() - intMinesFound
        bytRndXPOS = WorksheetFunction.RandBetween(1, intBoardRow)
        bytRndYPOS = WorksheetFunction.RandBetween(1, intBoardColumn)
        
        If arrCells(bytRndXPOS, bytRndYPOS).enumImage = Clear Then
            arrCells(bytRndXPOS, bytRndYPOS).enumImage = Mine
            arrCells(bytRndXPOS, bytRndYPOS).boolIsLocked = True
            i = i + 1
        End If
    Loop
End Sub

Private Sub FillsTheBoardWithGems()
    ' *** Fills the rest of the board with the gems. ***
    
    Dim i As Byte
    Dim y As Byte
    
    For i = 1 To intBoardRow
        For y = 1 To intBoardColumn
            If arrCells(i, y).enumImage = Clear Then
                arrCells(i, y).enumImage = Gem
                arrCells(i, y).boolIsLocked = True
            End If
        Next y
    Next i
End Sub

Private Sub SimulateImagesInBoard()
    ' *** Shows the images in the board at the end of the game. ***
    
    Call FillsTheBoardWithMines
    Call FillsTheBoardWithGems
    Call RemoveAllImagesFromBoard
    Call PlaceImagesInCells
End Sub

' -----------------------------------------------------------------------------
' ---------------------------- GAME OUTCOMES ----------------------------------
' -----------------------------------------------------------------------------

Private Function CalculateCashOutGain() As Double
    ' *** This function is responsible for calculating how much money would BSJ get if he cash out. ***
    
    CalculateCashOutGain = ((dblWagedMoneyMultiplier + (0.05 * intGemsFound)) * intGemsFound) * mBalanceManager.WagedMoney
End Function

Private Sub DebugGame()
    ' *** Debugs the logic of the mines Game. ***
    If boolDoDebug Then
        Debug.Print "--- MINES ---"
        Debug.Print "Is Game Active:        " & isGameBeingPlayed
        Debug.Print "Is Bossman Responding: " & mBSJResponse.IsBossmanResponding
        Debug.Print "Found Gems:            " & intGemsFound
        Debug.Print "--- END MINES ---"
    End If
End Sub

Private Sub HandleMineFound()
    ' *** Handles game, whenever the mine has been found. ***
    
    ' Creates the ilusion that game is not rigged, by showing the user current mine x gems combinations
    intMinesFound = intMinesFound + 1
    Call SimulateImagesInBoard
    
    ' Play the BSJ Reaction
    If mBalanceManager.AccountBalance < bytDonationAcceptThreshold Then
        Call mBSJResponse.ShowBossmanResponse(rage)
    Else
        Call mBSJResponse.ShowBossmanResponse(loss)
    End If
    
    isGameBeingPlayed = False
    Call HandleRestart
    Call mSheetBlocker.AllowWorksheetEdition
    Call SetPlayButtonToInitMode
    Call UpdateAllAccounts
End Sub

Private Sub HandleGemFound()
    ' *** Handles the game whenever the gem has been found by bsj. ***
    
    Dim dblResponseParam As Double
    
    Const BSJResponseChance = 95
    dblResponseParam = Rnd() * 100
    
    ' Updates images
    Call RemoveAllImagesFromBoard
    Call PlaceImagesInCells
    
    intGemsFound = intGemsFound + 1
    Call UpdateCashOutGain
    
    If dblResponseParam > BSJResponseChance Then
        Call mBSJResponse.ShowBossmanResponse(win)
    End If
    Call mSheetBlocker.PreventWorksheetFromEditing
End Sub

Private Sub HandleCashOut()
    ' *** Handles cashing out operation. ***
    
    Call SimulateImagesInBoard
    
    mBalanceManager.UpdateAccountAfterGambling (CalculateCashOutGain())
    Call UpdateAllAccounts
    
    If intGemsFound > 5 Then
        Call mBSJResponse.ShowBossmanResponse(win)
        'Call mBSJResponse.ShowBossmanResponse(big_win)
    Else
        Call mBSJResponse.ShowBossmanResponse(win)
    End If
    
    isGameBeingPlayed = False
    Call HandleRestart
    Call UpdateAllAccounts
    ' Call mSheetBlocker.AllowWorksheetEdition
End Sub

Private Sub HandleRestart()
    ' *** Handles the operation of restarting the game. ***
    
    intGemsFound = 0
    intMinesFound = 0

    Call InitializeBoard
    Call RemoveAllImagesFromBoard
    Call RefreshCurrentCellPosition
End Sub

' -----------------------------------------------------------------------------------
' ---------------------------- OTHER SHEET OBJECTS ----------------------------------
' -----------------------------------------------------------------------------------

Private Function GetMinesCount() As Integer
    ' *** Reads the ammount of mines from the sheet. ***
    
    Dim rngMinesCount As Range
    
    Set rngMinesCount = Range(strMinesCountRange)
    GetMinesCount = rngMinesCount.Value
End Function

Private Sub UpdateAllAccounts()
    ' *** Updates all accounts in the project ***
    
    Call mBalanceManager.UpdateAccountAfterGambling(0)
    Call UpdateAccount
    Call UpdateWagedMoney
    Call UpdateCashOutGain
End Sub
Private Sub SetPlayButtonToInitMode()
    ' *** Sets the button responsible for playing the game to the init mode. ***

    Dim shpDonateButton As Shape
    
    Set shpDonateButton = ActiveSheet.Shapes(shpDonateButtonName)
    
    With shpDonateButton
        .TextFrame.Characters.Text = "DONATE TO SLAKE"
    End With
End Sub

Private Sub SetPlayButtonToGameMode()
    ' *** Sets the button responsible for playing the game to the game mode. ***
    
    Dim shpDonateButton As Shape
    
    Set shpDonateButton = ActiveSheet.Shapes(shpDonateButtonName)
    
    With shpDonateButton
        .TextFrame.Characters.Text = "CASH OUT"
    End With
End Sub

Private Sub UpdateCashOutGain()
    ' *** Updates the shown gain on the cash out durning play. ***
    
    Dim rngCashOutRange As Range
    
    Set rngCashOutRange = Range(strCashOutRange)
    rngCashOutRange.Value = CalculateCashOutGain
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

' -----------------------------------------------------------------------------
' ------------------------- CALLBACKS SUPPORT ---------------------------------
' -----------------------------------------------------------------------------

Private Sub InitializeGameViaButton()
    ' *** Handles operation of initializg a game via button. ***
    
    If mBalanceManager.AccountBalance < 0.01 Then
        mMessageBoxes.msgCritical_NoneFunds
        Exit Sub
    End If
    
    isGameBeingPlayed = True
    
    Call SetWagedMoney
    Call mBalanceManager.UpdateAccountBeforeGambling(-mBalanceManager.WagedMoney)
    
    Call SetPlayButtonToGameMode
    Call HandleRestart
    Call UpdateAccount
    Call mSheetBlocker.PreventWorksheetFromEditing
End Sub

Private Sub ManageEnteringTheSheet()
    ' *** This function is executed, whenever the user enters the worksheet. ***
    
    isGameBeingPlayed = False
        
    Call mSheetBlocker.AllowWorksheetEdition
    Call HandleRestart
    Call SetPlayButtonToInitMode
    Call UpdateAccount
    Call UpdateWagedMoney
End Sub

Public Sub HandleSelectingCells(Optional ByVal varProxy As Variant = 0)
    ' *** Proxy function to be called in order to track the IsBSJResponding boolean. ***
    
    Call SelectACell
    Call RefreshCurrentCellPosition
End Sub

Private Sub ManageCashOutViaButton()
    ' *** Handles managing the cashout operation by button ***
    
    If intGemsFound < 1 Then
        Exit Sub
    End If
    
    isGameBeingPlayed = False
    
    Call mSheetBlocker.AllowWorksheetEdition
    Call HandleCashOut
    Call SetPlayButtonToInitMode
End Sub

Public Sub UpdateWallet(Optional ByVal varProxy As Variant = 0)
    ' *** Proxy function that handles managing of updating the wallet in the project. ***
    
    
    If mBalanceManager.AccountBalance < bytDonationAcceptThreshold Then
        mBalanceManager.AcceptJuicerFromRat
        Call UpdateAllAccounts
        Call mBSJResponse.ShowBossmanResponse(beg)
    Else
        mMessageBoxes.msgInfo_TooMuchFunds
    End If
End Sub

' -----------------------------------------------------------------------------
' ---------------------------- CALLBACKS --------------------------------------
' -----------------------------------------------------------------------------

Public Sub DoubleWagedMoney_click()
    ' *** Handles operation of doubling the waged money ***
    
    If mBSJResponse.IsBossmanResponding Or isGameBeingPlayed Then
        Exit Sub
    End If
    
    Call SetWagedMoney
    Call mBalanceManager.DoubleTheGamblingMoney
    Call UpdateWagedMoney
End Sub

Public Sub HalveTheWagedMoney_click()
    ' *** Handles operation of waging the half money. ***
    
    If mBSJResponse.IsBossmanResponding Or isGameBeingPlayed Then
        Exit Sub
    End If
    
    Call SetWagedMoney
    Call mBalanceManager.HalfTheGamblingMoney
    Call UpdateWagedMoney
End Sub

Public Sub Wallet_click()
    ' *** Handles the operations related to the gathering more money. ***
    
    If mBSJResponse.IsBossmanResponding Or isGameBeingPlayed Then
        Exit Sub
    End If
    
    Call mBSJResponse.CallBossmanResponse("Sheet2.UpdateWallet", strPortrairRange)
End Sub

Public Sub SelectRandom_click()
    ' *** Handles operations related to selecting random cells in the game. ***
    
    On Error Resume Next
    If mBSJResponse.IsBossmanResponding Or Not isGameBeingPlayed Then
        Exit Sub
    End If
    
    Call mSheetBlocker.AllowWorksheetEdition
    Call ChooseRandomCell
    Call mBSJResponse.CallBossmanResponse("Sheet2.HandleSelectingCells", strPortrairRange)
    Call DebugGame
End Sub

Public Sub DonateToSlake_click()
    ' *** Handles operations of donate to Slake button. ***
    
    If mBSJResponse.IsBossmanResponding Then
        Exit Sub
    End If
    
    If Not isGameBeingPlayed Then
        Call InitializeGameViaButton
    Else
        Call ManageCashOutViaButton
    End If
End Sub

' -------------------------------------------------------------
' ------------------- WORKSHEET EVENTS ------------------------
' -------------------------------------------------------------

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    On Error Resume Next
    If Not isGameBeingPlayed Or mBSJResponse.IsBossmanResponding Then
        Exit Sub
    End If
    
    If IsSelectedCellInBoard(Target.Address) Then
        Call mSheetBlocker.AllowWorksheetEdition
        Call mBSJResponse.CallBossmanResponse("Sheet2.HandleSelectingCells", strPortrairRange)
    End If
    Call DebugGame
End Sub

Private Sub Worksheet_Activate()
    ' *** Updates the worksheet, whenever user selects it. ***
    
    If Not mBSJResponse.IsBossmanResponding Then
       ManageEnteringTheSheet
    End If
End Sub
