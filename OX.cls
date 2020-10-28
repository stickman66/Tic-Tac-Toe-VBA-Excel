VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
Dim r(1 To 3, 1 To 3) As Integer, p(1 To 3, 1 To 3) As Integer, i As Integer, j As Integer
'matricea'
For i = 1 To 3
For j = 1 To 3
r(i, j) = Cells(i, j).Value
p(i, j) = Cells(i, j).Value
Next j
Next i
'verificare diagonale'
If r(1, 1) = r(2, 2) And r(2, 2) = r(3, 3) Then
MsgBox r(1, 1) & " win!"
End If
If r(1, 3) = r(2, 2) And r(2, 2) = r(3, 1) Then
MsgBox r(1, 3) & " win!"
End If
'verificare randuri'
For i = 1 To 3
If r(i, 1) = p(i, 2) And r(i, 2) = p(i, 3) And r(i, 1) = p(i, 3) Then
MsgBox r(i, 1) & " win"
Exit For
End If
Next i
'verificare coloane'
For i = 1 To 3
If r(1, i) = p(2, i) And r(2, i) = p(3, i) And r(1, i) = p(3, i) Then
MsgBox r(1, i) & " win"
Exit For
End If
Next i




End Sub

Private Sub CommandButton2_Click()
Dim rng As Range

     Set rng = Range("A1:c3")

    With rng.Borders
        .LineStyle = xlContinuous
        .Color = vbBlack
        .Weight = xlThin
    End With
    rng.ClearContents
    rng.HorizontalAlignment = xlCenter
    rng.VerticalAlignment = xlCenter
    rng.Font.Size = 30
    rng.ColumnWidth = 15
    rng.RowHeight = 60
End Sub
