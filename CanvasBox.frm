VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CanvasBox 
   Caption         =   "UserForm2"
   ClientHeight    =   3375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2055
   OleObjectBlob   =   "CanvasBox.frx":0000
End
Attribute VB_Name = "CanvasBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnFormatCanvas_Click()

    Dim row As Integer
    Dim col As Integer
    
    Columns("A:BB").ColumnWidth = 2.5
    Rows("1:44").RowHeight = 10
    With Range("B2:BA43")
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Weight = xlMedium
    End With
        
    For row = 2 To 43 Step 2
        For col = 2 To 53 Step 2
            Range(Cells(row, col), Cells(row + 1, col)).Merge
        Next col
    Next row

    For row = 3 To 42 Step 2
        For col = 3 To 53 Step 2
            Range(Cells(row, col), Cells(row + 1, col)).Merge
        Next col
    Next row
    
    'Cells(1, 1).Font.Color = RGB(255, 255, 255)
    Cells(1, 1) = 1

End Sub

Private Sub btnMoveDown_Click()

    If Cells(1, 1).Value = 43 Then
        Cells(1, 1).Value = 1
    Else
        Cells(1, 1).Value = Cells(1, 1).Value + 1
    End If
    
    HighlightCells

End Sub

Private Sub btnMoveUp_Click()

        If Cells(1, 1).Value = 1 Then
        Cells(1, 1).Value = 43
    Else
        Cells(1, 1).Value = Cells(1, 1).Value - 1
    End If

    HighlightCells

End Sub

Private Sub CommandButton1_Click()

Unload Me

End Sub
