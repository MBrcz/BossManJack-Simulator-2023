VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
' ***
' This module contains logic related to the so called "YesNoGame".
' Thanks for this user for inspiration:

' https://www.reddit.com/r/bossmanjack/comments/1bqfwqo/new_game_made_just_for_the_boss/
' ***

Private boolIsWin As Boolean
Private Const strBalanceAccountRange = "H1"
Private Const strWagedMoneyRange = "B9"
Private Const strResultWindowAdress = "F5"
Private Const bytDonationAcceptThreshold = 5
Private Const intWinThreshold = 75
Private Const strPortrairRange = "P1:R11"

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

Private Sub DecideTheGambleResult()
    ' *** Calculates the result of Special Gambling Operation. ***
    
    Dim dblDraw As Double
        
    dblDraw = Rnd() * 100
    
    If dblDraw > intWinThreshold Then
        boolIsWin = True
    Else
        boolIsWin = False
    End If
    
End Sub

Private Sub ShowGambleResult()
    ' *** Shows the result of the special in Game Called Yes / No Gambling Operation. ***
    
    Dim rngResultWindow As Range
    
    Set rngResultWindow = ActiveSheet.Range(strResultWindowAdress)
    
    With rngResultWindow
        With .Font
            .Size = 140
            .Bold = True
        End With
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    If boolIsWin Then
        rngResultWindow.Font.Color = vbGreen
        rngResultWindow.Value = "YES"
    ElseIf Not boolIsWin Then
        rngResultWindow.Value = "NO"
        rngResultWindow.Font.Color = vbRed
    End If
        
End Sub

Private Sub ClearResultWindow()
    ' *** Clears the window that shows the result of a match. ***
    
    Dim rngResultWindow As Range
    
    Set rngResultWindow = ActiveSheet.Range(strResultWindowAdress)
    rngResultWindow.Value = ""
End Sub

Private Sub UpdateAccountAfterGambling()
    ' *** Updates the state of bossman account after gambling. ***
    
    Dim dblWagedMoney As Double
    
    dblWagedMoney = ActiveSheet.Range(strWagedMoneyRange).Value
    
    If boolIsWin Then
        mBalanceManager.UpdateAccountAfterGambling (dblWagedMoney)
    ElseIf Not boolIsWin Then
        mBalanceManager.UpdateAccountAfterGambling (-dblWagedMoney)
    End If
    
End Sub

Private Sub ShowGamblingReaction()
    ' *** Handles showing Bossman reaction to the gambling result ***
    
    On Error Resume Next ' Prevents from triggering the some C++ random error
    
    If mBalanceManager.AccountBalance < bytDonationAcceptThreshold And Not boolIsWin Then
        Call mBSJResponse.ShowBossmanResponse(rage)
    
    ElseIf Not boolIsWin Then
        Call mBSJResponse.ShowBossmanResponse(loss)
    
    ElseIf boolIsWin Then
        Call mBSJResponse.ShowBossmanResponse(win)
        
    ElseIf boolIsWin And mBalanceManager.AccountBalance > 210 * bytDonationAcceptThreshold Then
        Call mBSJResponse.ShowBossmanResponse(big_win)
    End If
    
    On Error GoTo 0
End Sub

' ------------------------------------------------
' ---------------- PUBLIC NON CALLABLE -----------
' ------------------------------------------------

Public Sub UpdateWallet(Optional ByVal varProxy As Variant = 0)

    If mBalanceManager.AccountBalance < bytDonationAcceptThreshold Then
        Call ClearResultWindow
        Call mBalanceManager.AcceptJuicerFromRat
        Call UpdateAccount
        Call UpdateWagedMoney
        Call mBSJResponse.ShowBossmanResponse(beg)
    Else
        MsgBox "You have enough money, Bossman. No one will buy your bs till you're empty!", vbInformation
    End If
End Sub

Public Sub DonateToSlake(Optional ByVal varProxy As Variant = 0)
    
    Call DecideTheGambleResult
    Call ShowGambleResult
    Call UpdateAccountAfterGambling
    Call UpdateAccount
    Call UpdateWagedMoney
    Call ShowGamblingReaction
End Sub

Private Function IsPlayingPossible() As Boolean
    ' *** Checks if it is possible to play this game. ***
    
    If mBalanceManager.AccountBalance < 0.01 Then
        mMessageBoxes.msgCritical_NoneFunds
        Exit Function
        
    ElseIf mBSJResponse.IsBossmanResponding Then
        mMessageBoxes.msgCritical_BSJSpeaks
        Exit Function
    End If
    
    IsPlayingPossible = True
End Function

' ------------------------------------------------
' ---------------- CALLBACKS ---------------------
' ------------------------------------------------

Public Sub Wallet_click()
    ' *** Handles clicking the wallet, ergo getting more funds. ***
    
    If Not mBSJResponse.IsBossmanResponding Then
        Call mBSJResponse.CallBossmanResponse("Sheet1.UpdateWallet", strPortrairRange)
    End If
End Sub

Public Sub HalveMoney_click()
    ' *** Handles operation of halving the waged money. ***
    
    Call SetWagedMoney
    Call mBalanceManager.HalfTheGamblingMoney
    Call UpdateWagedMoney
End Sub

Public Sub DoubleMoney_click()
    ' *** Handles operation of doubling the waged money. ***
    
    Call SetWagedMoney
    Call mBalanceManager.DoubleTheGamblingMoney
    Call UpdateWagedMoney
End Sub

Public Sub DonateToSlake_click()
    ' *** Handles operation of donating the money straight to Slake.com. ***
    
    If Not IsPlayingPossible Then
        Exit Sub
    End If
    
    Call SetWagedMoney
    Call mBSJResponse.CallBossmanResponse("Sheet1.DonateToSlake", strPortrairRange)
End Sub

Private Sub Worksheet_Activate()

    If Not mBSJResponse.IsBossmanResponding Then
        Call ClearResultWindow
        Call UpdateAccount
        Call UpdateWagedMoney
    End If
End Sub

