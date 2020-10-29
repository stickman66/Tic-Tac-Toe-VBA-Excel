Private Sub RESTART_Click()
Dim rng As Range, i As Integer, k As Integer
    Set rng = Range("A1:c3")
    With rng.Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThin
    End With
    Range("A1:C3").Value = ""
    rng.HorizontalAlignment = xlCenter
    rng.VerticalAlignment = xlCenter
    rng.Font.Size = 30
    rng.ColumnWidth = 15
    rng.RowHeight = 60
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
If Not Intersect(Target, Range("A1:c3")) Is Nothing Then
Dim r(1 To 3, 1 To 3) As String, p(1 To 3, 1 To 3) As String, i As Integer, j As Integer
    'create matrice
For i = 1 To 3
For j = 1 To 3
r(i, j) = Cells(i, j).Value
p(i, j) = Cells(i, j).Value
Next j
Next i
'verification first diagonal'
If r(1, 1) = r(2, 2) And r(2, 2) = r(3, 3) Then
If r(1, 1) <> "" Then
MsgBox r(1, 1) & " win!1"
Exit Sub
End If
End If
'verification second diagonal'
If r(1, 3) = r(2, 2) And r(2, 2) = r(3, 1) Then
If r(1, 3) <> "" Then
MsgBox r(1, 3) & " win!2"
Exit Sub
End If
End If
'verification row'
For i = 1 To 3
If r(i, 1) = p(i, 2) And r(i, 2) = p(i, 3) And r(i, 1) = p(i, 3) And r(i, 1) <> "" Then
MsgBox r(i, 1) & " win3"
Exit For
Exit Sub
End If
Next i
'verification colom'
For i = 1 To 3
If r(1, i) = p(2, i) And r(2, i) = p(3, i) And r(1, i) = p(3, i) And r(1, i) <> "" Then
MsgBox r(1, i) & " win4"
Exit For
End If
Next i
End If
End Sub
