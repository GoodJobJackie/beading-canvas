Attribute VB_Name = "Canvas"
Option Explicit

Sub Canvas()
Attribute Canvas.VB_ProcData.VB_Invoke_Func = "g\n14"

    Cells(1, 1).Value = 1
    
    With CanvasBox
        .Top = 200
        .Left = 1000
    End With

    CanvasBox.Show

End Sub

Sub HighlightCells()

    Dim col As Integer
    Dim row As Integer
    Dim rngEveryOther As Range
    
    row = Cells(1, 1).Value

    If Cells(1, 1).Value = 1 Then
        'first row
        Range("C2, E2, G2, I2, K2, M2, O2, Q2, S2, U2, W2, Y2, AA2, AC2, AE2, AG2, AI2, AK2, AM2, AO2, AQ2, AS2, AU2, AW2, AY2, BA2").Select
    ElseIf (Cells(1, 1).Value > 1) And (Cells(1, 1).Value Mod 2 = 1) Then
        'odd
        For col = 3 To 53 Step 2 'loop from C to BA skipping 1 column in between
            If rngEveryOther Is Nothing Then
                Set rngEveryOther = Range(Cells(row, col), Cells(row + 1, col))
            Else
                Set rngEveryOther = Union(rngEveryOther, Range(Cells(row, col), Cells(row + 1, col)))
            End If
        Next col

    rngEveryOther.Select
        
    ElseIf (Cells(1, 1).Value > 1) And (Cells(1, 1).Value Mod 2 = 0) Then
        'even
        For col = 2 To 52 Step 2 'loop from C to BA skipping 1 column in between
            If rngEveryOther Is Nothing Then
                Set rngEveryOther = Range(Cells(row, col), Cells(row + 1, col))
            Else
                Set rngEveryOther = Union(rngEveryOther, Range(Cells(row, col), Cells(row + 1, col)))
            End If

        Next col

    rngEveryOther.Select
    End If

End Sub
