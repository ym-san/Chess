Attribute VB_Name = "Module1"
Public Sub �w�i�F������()

Dim Interior_Color_P As Long
Dim Interior_Color_B As Long
Interior_Color_P = RGB(244, 176, 132) '�w�i�F_�y�[���I�����W
Interior_Color_B = RGB(198, 89, 17) '�w�i�F_��
Dim i, j As Integer
            
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

End Sub


