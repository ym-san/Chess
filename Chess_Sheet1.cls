VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_SelectionChange(ByVal Target As Range)

Dim Turn_W As Boolean 'Trueが白のターン
Dim Piece_Select As Boolean 'Trueが駒を選択した状態

Turn_W = True
'Piece_Select = False
'---------------↑後ほどPublic化---------------------

Dim Turn As Range 'どっちのターンか表示するセル
Dim Piece As Range 'どの駒を選択中か表示するセル
Set Turn = Range("K2")
Set Piece = Range("K3")

Dim Piece_Color_W As Long
Dim Piece_Color_B As Long
Piece_Color_W = RGB(255, 255, 255) '白
Piece_Color_B = RGB(0, 0, 0) '黒
Dim Interior_Color_P As Long
Dim Interior_Color_B As Long
Interior_Color_P = RGB(244, 176, 132) '背景色_ペールオレンジ
Interior_Color_B = RGB(198, 89, 17) '背景色_茶

Dim SelectInterior_Color As Long '移動可能範囲の色
SelectInterior_Color = RGB(155, 194, 230) '水色

Dim i, j As Integer

'駒かどうか判別
If Turn_W And Not Target.Value = "" And Target.Font.Color = Piece_Color_W Then
 Piece_Select = True
ElseIf Turn_B And Not Target.Value = "" And Target.Font.Color = Piece_Color_B Then
 Piece_Select = True
Else
 Piece_Select = False
End If

If Piece_Select Then '駒の場合
    If Turn_W Then '白のターン
        If Not Piece.Value = "" Then
            '--背景色初期化--
            Range("B2:I9").Interior.Color = Interior_Color_P
            For i = 3 To 9 Step 2
                For j = 2 To 8 Step 2
                    Cells(i, j).Interior.Color = Interior_Color_B
                Next j
            Next i
            For i = 2 To 8 Step 2
                For j = 3 To 9 Step 2
                    Cells(i, j).Interior.Color = Interior_Color_B
                Next j
            Next i
            '--背景色初期化おわり--
        End If
        If Target.Value = "歩" And Target.Font.Color = Piece_Color_W Then
            Target.Offset(-1, 0).Interior.Color = SelectInterior_Color
            Target.Offset(-2, 0).Interior.Color = SelectInterior_Color
            Piece.Value = Target.Address
        End If
    Else '黒のターン(のちほど)

    End If
    Piece_Select = False
Else '駒ではない場合
    If Not Piece.Value = "" Then
        If Target.Interior.Color = SelectInterior_Color Then
            Target.Value = Range(Piece.Value)
            Target.Font.Color = Range(Piece.Value).Font.Color
            Range(Piece.Value).Value = ""
            Piece.Value = ""
        End If
            '--背景色初期化--
            Range("B2:I9").Interior.Color = Interior_Color_P
            For i = 3 To 9 Step 2
                For j = 2 To 8 Step 2
                    Cells(i, j).Interior.Color = Interior_Color_B
                Next j
            Next i
            For i = 2 To 8 Step 2
                For j = 3 To 9 Step 2
                    Cells(i, j).Interior.Color = Interior_Color_B
                Next j
            Next i
            '--背景色初期化おわり--
        Piece_Select = True
        Piece.Value = ""
    End If
End If
End Sub
