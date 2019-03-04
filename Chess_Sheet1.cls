VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal target As Range)

Dim Turn_W As Boolean 'Trueが白のターン'
Dim Piece_Select As Boolean 'Trueが駒を選択した状態'

'Turn_W = True'
'Piece_Select = False'

Dim Turn As Range 'どっちのターンか表示するセル
Dim Piece As Range 'どの駒を選択中か表示するセル
Dim Board As Range 'チェス盤の起点セル（左上）
Set Turn = Range("K2")
Set Piece = Range("K3")
Set Board = Range("B2")

Dim Piece_Color_W As Long
Dim Piece_Color_B As Long
Piece_Color_W = RGB(255, 255, 255) '白
Piece_Color_B = RGB(0, 0, 0) '黒

Dim SelectInterior_Color As Long '移動可能範囲の色
SelectInterior_Color = RGB(155, 194, 230) '水色

Dim i, j As Long
Dim blueCells() As Long

'ターンの判別'
If Turn.Value = "白" Then
 Turn_W = True
End If

'駒かどうか判別
If Turn_W And Not target.Value = "" And target.Font.Color = Piece_Color_W Then
 Piece_Select = True
ElseIf Not Turn_W And Not target.Value = "" And target.Font.Color = Piece_Color_B Then
 Piece_Select = True
Else
 Piece_Select = False
End If

If Piece_Select Then '駒の場合
    If Not Piece.Value = "" Then
        Call 背景色初期化
    End If
    If Turn_W Then '白のターン
        If target.Value = "歩" And target.Font.Color = Piece_Color_W Then
                blueCells() = 白のポーン(target, Piece_Color_B)
        ElseIf target.Value = "騎" And target.Font.Color = Piece_Color_W Then
                blueCells() = ナイト(target, Piece_Color_B)
        ElseIf target.Value = "僧" And target.Font.Color = Piece_Color_W Then
                blueCells() = ビショップ(target, Piece_Color_B)
        ElseIf target.Value = "城" And target.Font.Color = Piece_Color_W Then
                blueCells() = ルーク(target, Piece_Color_B)
        ElseIf target.Value = "女" And target.Font.Color = Piece_Color_W Then
                blueCells() = クイーン(target, Piece_Color_B)
        ElseIf target.Value = "王" And target.Font.Color = Piece_Color_W Then
                blueCells() = キング(target, Piece_Color_B)
        End If
        Piece.Value = target.Address
                
        If Not UBound(blueCells, 1) = 0 Then
            For i = 0 To UBound(blueCells, 2)
                If Not Application.Intersect(Cells(blueCells(0, i), blueCells(1, i)), Range(Board, Board.Offset(7, 7))) Is Nothing Then
                    Cells(blueCells(0, i), blueCells(1, i)).Interior.Color = SelectInterior_Color
                End If
            Next
        End If

    Else '黒のターン
        If target.Value = "歩" And target.Font.Color = Piece_Color_B Then
                blueCells() = 黒のポーン(target, Piece_Color_W)
        ElseIf target.Value = "騎" And target.Font.Color = Piece_Color_B Then
                blueCells() = ナイト(target, Piece_Color_W)
        ElseIf target.Value = "僧" And target.Font.Color = Piece_Color_B Then
                blueCells() = ビショップ(target, Piece_Color_W)
        ElseIf target.Value = "城" And target.Font.Color = Piece_Color_B Then
                blueCells() = ルーク(target, Piece_Color_W)
        ElseIf target.Value = "女" And target.Font.Color = Piece_Color_B Then
                blueCells() = クイーン(target, Piece_Color_W)
        ElseIf target.Value = "王" And target.Font.Color = Piece_Color_B Then
                blueCells() = キング(target, Piece_Color_W)
        End If
        Piece.Value = target.Address
                
        If Not UBound(blueCells, 1) = 0 Then
            For i = 0 To UBound(blueCells, 2)
                If Not Application.Intersect(Cells(blueCells(0, i), blueCells(1, i)), Range(Board, Board.Offset(7, 7))) Is Nothing Then
                    Cells(blueCells(0, i), blueCells(1, i)).Interior.Color = SelectInterior_Color
                End If
            Next
        End If

    End If
    Piece_Select = False
Else '駒ではない場合
    If Not Piece.Value = "" Then
        If target.Interior.Color = SelectInterior_Color Then
            target.Value = Range(Piece.Value)
            target.Font.Color = Range(Piece.Value).Font.Color
            Range(Piece.Value).Value = ""
            Piece.Value = ""
            If Turn_W Then
                Turn.Value = "黒"
            Else
                Turn.Value = "白"
            End If
        End If
        Call 背景色初期化
        Piece_Select = True
        Piece.Value = ""
    End If
End If
End Sub


