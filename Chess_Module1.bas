Attribute VB_Name = "Module1"
Public Sub 背景色初期化()

Dim Interior_Color_P As Long
Dim Interior_Color_B As Long
Interior_Color_P = RGB(244, 176, 132) '背景色_ペールオレンジ
Interior_Color_B = RGB(198, 89, 17) '背景色_茶
Dim i, j As Integer
            
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

End Sub

Sub はじめから()

Dim Piece_Color_W As Long
Dim Piece_Color_B As Long
Piece_Color_W = RGB(255, 255, 255) '白'
Piece_Color_B = RGB(0, 0, 0) '黒'

Dim Turn As Range 'どっちのターンか表示するセル'
Set Turn = Range("K2")

'最初のターンは白'
Turn.Value = "白"

'背景色'
Call 背景色初期化

'文字色'
Range("B2:I3").Font.Color = Piece_Color_B
Range("B8:I9").Font.Color = Piece_Color_W

'駒'
'クリア'
Range("B2:I9").Value = ""
'ポーン'
Range("B3:I3").Value = "歩"
Range("B8:I8").Value = "歩"
'ナイト'
Range("C2").Value = "騎"
Range("H2").Value = "騎"
Range("C9").Value = "騎"
Range("H9").Value = "騎"
'ビショップ'
Range("D2").Value = "僧"
Range("G2").Value = "僧"
Range("D9").Value = "僧"
Range("G9").Value = "僧"
'ルーク'
Range("B2").Value = "城"
Range("I2").Value = "城"
Range("B9").Value = "城"
Range("I9").Value = "城"
'クイーン'
Range("E2").Value = "女"
Range("E9").Value = "女"
'キング'
Range("F2").Value = "王"
Range("F9").Value = "王"

End Sub


Function 白のポーン(ByVal target As Range, Piece_Color_W As Long, Piece_Color_B As Long) As Long()
    
    Dim i As Long
    Dim result() As Long
    Dim arr As Variant
    i = 0
        
    m = target.Column
    arr = Array(-1, -2)
    For j = 0 To 1
        n = target.row + arr(j)
        If n > 0 And m > 0 Then
            If Cells(n, m).Value = "" Then
                If j = 1 And Cells(n + 1, m).Value <> "" Then
                Else
                    ReDim Preserve result(1, i)
                    result(0, i) = n
                    result(1, i) = m
                    i = i + 1
                End If
            End If
        End If
    Next
    
    n = target.row - 1
    arr = Array(-1, 1)
    For j = 0 To 1
        m = target.Column + arr(j)
        If n > 0 And m > 0 Then
            If Cells(n, m).Value <> "" And Cells(n, m).Font.Color = Piece_Color_B Then
                ReDim Preserve result(1, i)
                result(0, i) = n
                result(1, i) = m
                i = i + 1
            End If
        End If
    Next
        
    If i = 0 Then
        ReDim Preserve result(0, 0)
    Else
    End If
    白のポーン = result()
End Function
 
