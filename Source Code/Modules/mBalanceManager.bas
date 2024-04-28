Attribute VB_Name = "mBalanceManager"
Option Explicit
' ***
' This module is responsible for all operations relating to managing bossman's current account.
' ***
Public boolDebugBalance As Boolean

Private dblAccountBalance As Double
Private dblWagedMoney As Double
Private Const dblDonatedJuicerMult = 1000

' -------------------------------------------------------
' ------------------ GETTERS && SETTERS -----------------
' -------------------------------------------------------

Private Property Let AccountBalance(ByVal dblNewBalance As Double)
    dblAccountBalance = dblNewBalance
    If dblNewBalance <= 0.01 Then
        dblNewBalance = 0
    End If
End Property

Public Property Get AccountBalance() As Double
    AccountBalance = dblAccountBalance
End Property

Public Property Let WagedMoney(ByVal dblNewWagedMoney As Double)
    dblWagedMoney = dblNewWagedMoney
End Property

Public Property Get WagedMoney() As Double
    WagedMoney = dblWagedMoney
End Property

' ------------------------------------------------------
' ----------------- PRIVATE FUNCTIONS ------------------
' ------------------------------------------------------

Private Function IsGamblingPossible() As Boolean
    ' *** Checks whether is it possible to gamble with current state of account. ***
    
    If AccountBalance >= WagedMoney Or WagedMoney <> 0 Then
        IsGamblingPossible = True
    Else
        IsGamblingPossible = False
    End If
End Function

Private Sub DebugBalanceManager()
    ' ***
    ' Debugs the balance manager - use for development purpouses only.
    ' ***
    
    If boolDebugBalance Then
        Debug.Print ("--- BALANCE MANAGER ---")
        Debug.Print ("Current Account State: " & AccountBalance)
        Debug.Print ("Waged Money: " & WagedMoney)
        Debug.Print ("Is gambling possible: " & IsGamblingPossible)
        Debug.Print ("--- END BALANCE MANAGER ---")
    End If
End Sub

' ------------------------------------------------------
' ---------------------- FUNCTIONS ---------------------
' ------------------------------------------------------

'Public Sub SetAccountParameters(ByVal dblCurrentBalance As Double, ByVal dblWaged As Double)
'    ' *** Sets the account parameters for the current try of the gambling. ***
'
'    AccountBalance = dblCurrentBalance
'    WagedMoney = dblWaged
'End Sub

Public Sub DoubleTheGamblingMoney(Optional ByVal varProxy As Variant = 0)
    ' *** Doubles the money for gambling. ***
    
    Dim dblTempWaged As Double
    
    dblTempWaged = WagedMoney * 2
    
    If dblTempWaged > AccountBalance Then
        dblTempWaged = AccountBalance
    End If
    
    WagedMoney = Round(dblTempWaged, 2)
End Sub

Public Sub HalfTheGamblingMoney(Optional ByVal varProxy As Variant = 0)
    ' *** Halves the money for gambling. ***
    
    ' Yea, it's stupid but just in case.
    Dim dblTempWaged As Double
    
    dblTempWaged = WagedMoney * 1 / 2
    
    If dblTempWaged > AccountBalance Then
        dblTempWaged = AccountBalance
    End If
    
    WagedMoney = Round(dblTempWaged, 2)
End Sub

Public Sub AcceptJuicerFromRat(Optional ByVal varProxy As Variant = 0)
    ' *** It is responsible for genetating more money in the wallet ***
    
    Dim rndRandomAmmount As Double
    
    rndRandomAmmount = Rnd() * dblDonatedJuicerMult
    AccountBalance = Round(rndRandomAmmount, 2)
    WagedMoney = Round(rndRandomAmmount / 2, 2)
    Call DebugBalanceManager
End Sub

Public Sub UpdateAccountAfterGambling(Optional ByVal dblAccountChange As Double)
    ' *** Updates the user account parameters after special gambling operation. ***
    
    AccountBalance = Round(AccountBalance + dblAccountChange, 2)
    
    If WagedMoney > AccountBalance Then
        WagedMoney = AccountBalance
    End If
    Call DebugBalanceManager
End Sub

Public Sub UpdateAccountBeforeGambling(Optional ByVal dblAccountChange As Double)
    ' *** It does what function above, but does not changes waged money ammount. ***
    
    AccountBalance = Round(AccountBalance + dblAccountChange, 2)
    Call DebugBalanceManager
End Sub
