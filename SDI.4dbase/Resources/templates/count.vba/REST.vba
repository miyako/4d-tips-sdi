Private Sub Worksheet_Change(ByVal rTarget As Range)
  countRows = ActiveSheet.Rows.Count

  Dim colA As Range
  Set colA = Range("A2:A" & countRows)
  Dim colC As Range
  Set colC = Range("C2:C" & countRows)
  Dim cTarget As Range

  If Not Application.Intersect(colA, rTarget) Is Nothing Then
    For Each cTarget In rTarget.Cells
      With cTarget
        Call call4D.getProductNameForCode(.Value, _
        cTarget.Offset(0, 1), _
        cTarget.Offset(0, 2), _
        cTarget.Offset(0, 3), _
        cTarget.Offset(0, 4), _
        cTarget.Offset(0, 5), _
        cTarget.Offset(0, 6))
      End With
    Next cTarget
  ElseIf Not Application.Intersect(colC, rTarget) Is Nothing Then
    For Each cTarget In rTarget.Cells
      With cTarget
If Not IsEmpty(cTarget) Then
        Call call4D.updateProductCountForCode(.Value, _
        cTarget.Offset(0, -2).Value, _
        cTarget.Offset(0, 1), _
        cTarget.Offset(0, 2), _
        cTarget.Offset(0, 3), _
        cTarget.Offset(0, 4))
End If
      End With
    Next cTarget
  End If
End Sub
