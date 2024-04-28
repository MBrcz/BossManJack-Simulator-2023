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
Public Sub DebugModules()
    
    ' mBalanceManager.boolDebugBalance = True
    ' mBSJResponse.boolDoDebug = True
    ' Sheet2.boolDoDebug = True
    Sheet3.boolDoDebug = True
End Sub
    
Private Sub Workbook_Open()
    ' *** Adds some money for user durning the opening. ***

    Dim wsSheet As Worksheet
    Dim strBSJPortrairSignature As String
    
    mBalanceManager.AcceptJuicerFromRat
    
    Application.ScreenUpdating = False
    strBSJPortrairSignature = mBSJResponse.PortrairSignature
    For Each wsSheet In ActiveWorkbook.Sheets
        wsSheet.Select
        Call mUtils.DeleteShapeWithSignature(strBSJPortrairSignature)
    Next wsSheet
    Worksheets(1).Select
    Application.ScreenUpdating = True
    Call mUtils.ZoomToLastVisibleColumnAndRow
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    ' *** Prevents leaving the sheet while BSJ animation plays. ***
    
    Dim strBSJPortrairSignature As String
    
    Application.ScreenUpdating = False
            
    strBSJPortrairSignature = mBSJResponse.PortrairSignature
    If mBSJResponse.IsBossmanResponding Then
        Worksheets(mBSJResponse.SourceSheetName).Select
    Else
        Call mUtils.ZoomToLastVisibleColumnAndRow
        Call mUtils.DeleteShapeWithSignature(strBSJPortrairSignature)
    End If
    
    Application.ScreenUpdating = True
End Sub


