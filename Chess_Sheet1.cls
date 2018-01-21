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

Dim Turn_W As Boolean 'True�����̃^�[��
Dim Piece_Select As Boolean 'True�����I���������

Turn_W = True
'Piece_Select = False
'---------------����ق�Public��---------------------

Dim Turn As Range '�ǂ����̃^�[�����\������Z��
Dim Piece As Range '�ǂ̋��I�𒆂��\������Z��
Set Turn = Range("K2")
Set Piece = Range("K3")

Dim Piece_Color_W As Long
Dim Piece_Color_B As Long
Piece_Color_W = RGB(255, 255, 255) '��
Piece_Color_B = RGB(0, 0, 0) '��
Dim Interior_Color_P As Long
Dim Interior_Color_B As Long
Interior_Color_P = RGB(244, 176, 132) '�w�i�F_�y�[���I�����W
Interior_Color_B = RGB(198, 89, 17) '�w�i�F_��

Dim SelectInterior_Color As Long '�ړ��\�͈͂̐F
SelectInterior_Color = RGB(155, 194, 230) '���F

Dim i, j As Integer

'��ǂ�������
If Turn_W And Not Target.Value = "" And Target.Font.Color = Piece_Color_W Then
 Piece_Select = True
ElseIf Turn_B And Not Target.Value = "" And Target.Font.Color = Piece_Color_B Then
 Piece_Select = True
Else
 Piece_Select = False
End If

If Piece_Select Then '��̏ꍇ
    If Turn_W Then '���̃^�[��
        If Not Piece.Value = "" Then
            '--�w�i�F������--
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
            '--�w�i�F�����������--
        End If
        If Target.Value = "��" And Target.Font.Color = Piece_Color_W Then
            Target.Offset(-1, 0).Interior.Color = SelectInterior_Color
            Target.Offset(-2, 0).Interior.Color = SelectInterior_Color
            Piece.Value = Target.Address
        End If
    Else '���̃^�[��(�̂��ق�)

    End If
    Piece_Select = False
Else '��ł͂Ȃ��ꍇ
    If Not Piece.Value = "" Then
        If Target.Interior.Color = SelectInterior_Color Then
            Target.Value = Range(Piece.Value)
            Target.Font.Color = Range(Piece.Value).Font.Color
            Range(Piece.Value).Value = ""
            Piece.Value = ""
        End If
            '--�w�i�F������--
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
            '--�w�i�F�����������--
        Piece_Select = True
        Piece.Value = ""
    End If
End If
End Sub