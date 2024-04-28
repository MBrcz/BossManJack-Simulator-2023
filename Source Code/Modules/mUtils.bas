Attribute VB_Name = "mUtils"
Option Explicit
' *** This module shall store all utilitary functions related to the project. ***

Public Function GetPathToRootDir() As String
    ' *** Gets the path to the root directory of the project.

    GetPathToRootDir = Replace(ThisWorkbook.FullName, Split(ThisWorkbook.FullName, "\")(UBound(Split(ThisWorkbook.FullName, "\"))), "")
End Function

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
    ' Returns:
    '   - None

    With shpImage
        .LockAspectRatio = msoFalse
        .Top = rngTarget.Top
        .Left = rngTarget.Left
        .Width = rngTarget.Width
        .Height = rngTarget.Height
        .Name = strNewName
    End With

End Sub

Public Function IsFileExists(strPath As String) As Boolean
    ' *** Checks if file [NOT DIRECTORY] at the passed path exists or not ***
    
    ' Accepts:
    '   - strPath [string] - the path where the file shhould be located
    
    ' Returns:
    '   - boolean - True means that file exists otherwise False.
    
    Dim oFso As Object
    
    Set oFso = CreateObject("Scripting.FileSystemObject")
    IsFileExists = oFso.fileexists(strPath)
    Set oFso = Nothing
End Function

Public Sub DeleteShapeWithSignature(ByVal strSignature As String)
    ' *** Removes all shapes in the current sheet, that has a known signature ***
    
    ' Accepts:
    '   - strSignature [string] - the text, that shape must contain in order to be removed
    ' Returns:
    '   - None
    
    Dim shpShape As Shape
    
    For Each shpShape In ActiveSheet.Shapes
        If InStr(1, shpShape.Name, strSignature, vbTextCompare) Then
            shpShape.Delete
        End If
    Next
End Sub

Public Sub ZoomToLastVisibleColumnAndRow(Optional ByVal varProxy As Variant = 0)
    ' ***
    ' It sets the zoom of the page as high as possible regarding
    ' to user screen settings
    ' ***
        
    Dim visibleRange As Range
    
    Set visibleRange = ActiveSheet.Cells.SpecialCells(xlCellTypeVisible)
    visibleRange.Select
    
    ActiveWindow.Zoom = True
    ActiveSheet.Cells(1, 1).Select
End Sub

