Attribute VB_Name = "Module1"
Sub GetUnique()

Dim wb As Workbook
Dim ws As Worksheet
Dim Rng As Range

Set wb = ThisWorkbook
    
For x = 2018 To 2020
    Set ws = wb.Worksheets(CStr(x))
    Set Rng = ws.Range("K1")
    ws.Range("A:A").AdvancedFilter Action:=xlFilterCopy, CopyToRange:=Rng, Unique:=True

Next x

End Sub


Sub yearly_open()

Dim a As Integer
Dim i As Long
Dim j As Integer
Dim LR As Long


j = 1
For i = 1 To 753001
    a = StrComp(Sheet1.Cells(i, 1).Value, Sheet1.Cells(i + 1, 1).Value, vbTextCompare)
        If a <> 0 Then
            j = j + 1
            Sheet1.Cells(i + 1, 3).Copy Sheet1.Cells(j, 13)
            
        ElseIf Sheet1.Cells(i + 1, 13).Value = " " Then
            Sheet1.Cells(i + 1, 13).Delete
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
    a = StrComp(Sheet1.Cells(i, 1).Value, Sheet1.Cells(i + 1, 1).Value, vbTextCompare)
        If a <> 0 Then
            j = j + 1
            Sheet1.Cells(i, 6).Copy Sheet1.Cells(j, 14)
        End If
Next i

End Sub

Sub yearly_difference()


Dim a As Double
Dim i As Long
Dim b As Double
Dim c As Double



For i = 2 To 753001
    a = Sheet1.Cells(i, 13).Value
    b = Sheet1.Cells(i, 14).Value
    c = b - a
    Sheet1.Cells(i, 15).Value = c


Next i

End Sub

Sub yearly_percent()


Dim a As Double
Dim i As Long
Dim b As Double
Dim c As Double


For i = 2 To 753001
    a = Sheet1.Cells(i, 13).Value
    b = Sheet1.Cells(i, 15).Value
        If a <> 0 Then
            c = b / a
            Sheet1.Cells(i, 16).Value = c

        End If
Next i
End Sub



Sub maximum_minimum()


Dim Max As Double
Dim Rng As Integer
Dim Min As Double


Max = Application.WorksheetFunction.Max(Sheet1.Range("P:P").Value)
Sheet1.Range("R3") = Max

Min = Application.WorksheetFunction.Min(Sheet1.Range("P:P").Value)
Sheet1.Range("R4") = Min

End Sub

Sub highlights()

Dim Rng As Range

Set Rng = Sheet1.Range("O:O")

For Each cell In Rng
        If cell.Value < 0 Then
            cell.Interior.ColorIndex = 3
        Else
            cell.Interior.ColorIndex = 4
        End If
Next
End Sub


