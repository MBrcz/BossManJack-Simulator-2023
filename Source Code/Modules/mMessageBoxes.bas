Attribute VB_Name = "mMessageBoxes"
Option Explicit
' *** This module holds the string text used in boxes etc ***

Public Sub msgCritical_KinoBoardNotInstanciated(Optional ByVal varProxy As Variant = 0)
    MsgBox err_KinoBoardNotInstanciated, vbCritical + vbOKOnly, "Error"
End Sub

Public Sub msgCritical_NoneFunds(Optional ByVal varProxy As Variant = 0)
    MsgBox err_NoneFunds, vbCritical + vbOKOnly, "Error"
End Sub

Public Sub msgCritical_BSJSpeaks(Optional ByVal varProxy As Variant = 0)
    MsgBox err_BSJSpeaks, vbCritical + vbOKOnly, "Error"
End Sub

Public Sub msgInfo_TooMuchFunds(Optional ByVal varProxy As Variant = 0)
    MsgBox err_TooMuchFunds, vbCritical + vbOKOnly, "Information"
End Sub

Public Sub msgInfoKino_Win(ByVal dblAmmount As Double)
    MsgBox info_KinoWin(dblAmmount), vbInformation + vbOKOnly, "Information"
End Sub

Private Function info_KinoWin(ByVal dblAmmount As Double)
    info_KinoWin = "Congratulations! You have won: " & Round(dblAmmount, 2) & " of USD!" & vbCrLf & "Just do not GAMBLE everything at once!"
End Function

Private Function err_NoneFunds() As String
    err_NoneFunds = "You have no money dude! Start begging from rats or Eddie, they MIGHT provide something!"
End Function

Private Function err_TooMuchFunds() As String
    err_TooMuchFunds = "You have too much money dude, no one will buy this BS till you're empty!"
End Function

Private Function err_KinoBoardNotInstanciated() As String
    err_KinoBoardNotInstanciated = "Board is not initialized! Press Setup Board button!"
End Function

Private Function err_BSJSpeaks() As String
    err_BSJSpeaks = "According to the flags, BSJ should be speaking right now. If it's not the case - press Slake.com logo."
End Function
