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

Private Function GetAccountBalance() As Double
    GetAccountBalance = Worksheets(1).Range("A100").Value
End Function

Private Sub SetAccountBalance(ByVal dblAmmount As Double)
    Worksheets(1).Range("A100").Value = dblAmmount
End Sub

Private Sub SetWagedAmmount(ByVal dblAmmount As Double)
    Worksheets(1).Range("B9").Value = dblAmmount
End Sub

Private Function GetWagedAmmount() As Double
    GetWagedAmmount = Worksheets(1).Range("B9").Value
End Function

Private Function IsGambleSuccesfull() As Boolean
    
     Dim dblDraw As Integer
    Dim boolIsWin As Boolean

    Const dblWinTreshold = 75
    dblDraw = Rnd * 100

    If dblDraw >= dblWinTreshold Then
        IsGambleSuccesfull = True
    Else
        IsGambleSuccesfull = False
    End If
    
End Function

Private Function IsAccountBalanceValid() As Boolean
    
    If GetAccountBalance() >= GetWagedAmmount And GetAccountBalance() > 0 Then
        IsAccountBalanceValid = True
    Else
        IsAccountBalanceValid = False
    End If
    
End Function

Private Sub ShowGambleResult(ByVal boolIsWin As Boolean)
    
    If boolIsWin Then
        Range("F5").Value = "YES"
        Range("F5").Font.Color = vbGreen
    ElseIf Not boolIsWin Then
        Range("f5").Value = "NO"
        Range("f5").Font.Color = vbRed
    End If
    
End Sub

Private Sub UpdateWagedAmmount(ByVal dblValue As Double)
        
    If dblValue > GetAccountBalance() Then
        Call SetWagedAmmount(GetAccountBalance)
        Exit Sub
    End If
    
    Call SetWagedAmmount(dblValue)
    
End Sub

Private Sub UpdateAccountState(ByVal boolIsWin As Boolean)
    
    If boolIsWin Then
        Call SetAccountBalance(GetAccountBalance() + GetWagedAmmount())
    
    Else
        Call SetAccountBalance(GetAccountBalance() - GetWagedAmmount())
    End If
End Sub

Private Sub IncreaseTheWalletAmmount()
    
    Dim dblDonatedAmmount As Double
    
    If GetAccountBalance() < 5 Then
        
        dblDonatedAmmount = Round(Rnd() * 500, 0)
        Call SetAccountBalance(dblDonatedAmmount)
        Call SetWagedAmmount(dblDonatedAmmount / 2)
        Call ThisWorkbook.PerformBossManResponse(beg)
        Call MsgBox("Worry not BossMan, some good willed RAT donated some juicers to you!", vbInformation + vbOKOnly)
        Range("F5").Value = ""
    End If
    
End Sub

Private Function HandlePlay()
    
    Dim boolIsWin As Boolean
    
    boolIsWin = IsGambleSuccesfull
    
    If IsAccountBalanceValid Then
        
        Call ShowGambleResult(boolIsWin)
        Call UpdateAccountState(boolIsWin)
    
        If boolIsWin Then
            Call ThisWorkbook.PerformBossManResponse(win)
        Else
            If GetAccountBalance < 5 Then
                Call ThisWorkbook.PerformBossManResponse(rage)
            Else
                Call ThisWorkbook.PerformBossManResponse(loss)
            End If
        End If
            
    Else
        MsgBox "Not sufficient funds for this operation!", vbCritical
        Call SetWagedAmmount(GetAccountBalance())
        
    End If
        
End Function

Public Sub DoubleGambledMoney_click()
    
    Call UpdateWagedAmmount(2 * GetWagedAmmount)
End Sub

Public Sub HalfGambledMoney_click()

    Call UpdateWagedAmmount(1 / 2 * GetWagedAmmount)
End Sub

Public Sub WalletButton_click()

    Call IncreaseTheWalletAmmount
End Sub

Public Sub Play_click()
    
    HandlePlay
End Sub



