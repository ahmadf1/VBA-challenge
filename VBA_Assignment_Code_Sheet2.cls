VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub yearly_open()

Dim a As Integer
Dim i As Long
Dim j As Integer
Dim LR As Long


j = 1
For i = 1 To 753001
    a = StrComp(Sheet2.Cells(i, 1).Value, Sheet2.Cells(i + 1, 1).Value, vbTextCompare)
        If a <> 0 Then
            j = j + 1
            Sheet2.Cells(i + 1, 3).Copy Sheet2.Cells(j, 13)
            
        ElseIf Sheet2.Cells(i + 1, 13).Value = " " Then
            Sheet2.Cells(i + 1, 13).Delete
        End If
Next i

End Sub


Sub yearly_close()


Dim a As Integer
Dim i As Long
Dim j As Integer
Dim k As Long

j = 0
k = 753001
For i = 1 To k
    a = StrComp(Sheet2.Cells(i, 1).Value, Sheet2.Cells(i + 1, 1).Value, vbTextCompare)
        If a <> 0 Then
            j = j + 1
            Sheet2.Cells(i, 6).Copy Sheet2.Cells(j, 14)
        End If
Next i

End Sub

Sub yearly_difference()


Dim a As Double
Dim i As Long
Dim b As Double
Dim c As Double



For i = 2 To 753001
    a = Sheet2.Cells(i, 13).Value
    b = Sheet2.Cells(i, 14).Value
    c = b - a
    Sheet2.Cells(i, 15).Value = c


Next i

End Sub

Sub yearly_percent()


Dim a As Double
Dim i As Long
Dim b As Double
Dim c As Double


For i = 2 To 753001
    a = Sheet2.Cells(i, 13).Value
    b = Sheet2.Cells(i, 15).Value
        If a <> 0 Then
            c = b / a
            Sheet2.Cells(i, 16).Value = c

        End If
Next i
End Sub



Sub maximum_minimum()


Dim Max As Double
Dim Rng As Integer
Dim Min As Double


Max = Application.WorksheetFunction.Max(Sheet2.Range("P:P").Value)
Sheet2.Range("R3") = Max

Min = Application.WorksheetFunction.Min(Sheet2.Range("P:P").Value)
Sheet2.Range("R4") = Min

End Sub

Sub highlights()

Dim Rng As Range

Set Rng = Sheet2.Range("O:O")

For Each cell In Rng
        If cell.Value < 0 Then
            cell.Interior.ColorIndex = 3
        Else
            cell.Interior.ColorIndex = 4
        End If
Next
End Sub



