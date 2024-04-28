Attribute VB_Name = "mSheetBlocker"
Option Explicit
' *** This module should handle special operations related to the blocking the module edibitily. ***
' *** Note: This won't work for preventing changing cells content durning animations. ***

Private Const strPassword = "YOUWONTUSEITTOCHEATWOULDYOUQUESTIONMARK"

Public Sub PreventWorksheetFromEditing()
    ' *** Makes sure, that all cells of the worksheet would not be editable. ***
    
    ActiveSheet.Protect strPassword
End Sub

Public Sub AllowWorksheetEdition()
    ' *** Allows the worksheet to be edited. ***
    
    ActiveSheet.Unprotect strPassword
End Sub
