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

Sub �͂��߂���()

Dim Piece_Color_W As Long
Dim Piece_Color_B As Long
Piece_Color_W = RGB(255, 255, 255) '��'
Piece_Color_B = RGB(0, 0, 0) '��'

Dim Turn As Range '�ǂ����̃^�[�����\������Z��'
Set Turn = Range("K2")

'�ŏ��̃^�[���͔�'
Turn.Value = "��"

'�w�i�F'
Call �w�i�F������

'�����F'
Range("B2:I3").Font.Color = Piece_Color_B
Range("B8:I9").Font.Color = Piece_Color_W

'��'
'�N���A'
Range("B2:I9").Value = ""
'�|�[��'
Range("B3:I3").Value = "��"
Range("B8:I8").Value = "��"
'�i�C�g'
Range("C2").Value = "�R"
Range("H2").Value = "�R"
Range("C9").Value = "�R"
Range("H9").Value = "�R"
'�r�V���b�v'
Range("D2").Value = "�m"
Range("G2").Value = "�m"
Range("D9").Value = "�m"
Range("G9").Value = "�m"
'���[�N'
Range("B2").Value = "��"
Range("I2").Value = "��"
Range("B9").Value = "��"
Range("I9").Value = "��"
'�N�C�[��'
Range("E2").Value = "��"
Range("E9").Value = "��"
'�L���O'
Range("F2").Value = "��"
Range("F9").Value = "��"

End Sub


Function ���̃|�[��(ByVal target As Range, Piece_Color_W As Long, Piece_Color_B As Long) As Long()
    
    Dim i, j As Long
    Dim result() As Long
    Dim arr As Variant
    i = 0
        
    m = target.Column
    arr = Array(-1, -2)
    For j = 0 To 1
        n = target.Row + arr(j)
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
    
    n = target.Row - 1
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
    ���̃|�[�� = result()
End Function

Function ���̃i�C�g(ByVal target As Range, Piece_Color_W As Long, Piece_Color_B As Long) As Long()
    
    Dim i, j As Long
    Dim result() As Long
    Dim arr_row() As Variant
    Dim arr_col() As Variant
    
    arr_row() = Array(1, 2, 2, 1, -1, -2, -2, -1)
    arr_col() = Array(2, 1, -1, -2, -2, -1, 1, 2)
        
    i = 0
    For j = 0 To 7
    n = target.Row + arr_row(j)
    m = target.Column + arr_col(j)
        If n > 0 And m > 0 Then
            If Cells(n, m).Value = "" Or Cells(n, m).Font.Color = Piece_Color_B Then
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
    ���̃i�C�g = result()
End Function
 
Function ���̃r�V���b�v(ByVal target As Range, Piece_Color_W As Long, Piece_Color_B As Long) As Long()
    
    Dim i, j As Long
    Dim result() As Long
            
    Dim UpperRight, UpperLeft, BottomRight, BottomLeft As Boolean
    UpperRight = True
    UpperLeft = True
    BottomRight = True
    BottomLeft = True
    
    i = 0
    
    For j = 1 To 7
        If UpperRight Then
            n = target.Row + (-1 * j)
            m = target.Column + (1 * j)
            If n > 0 And m > 0 Then
                If Cells(n, m).Value = "" Or Cells(n, m).Font.Color = Piece_Color_B Then
                    ReDim Preserve result(1, i)
                    result(0, i) = n
                    result(1, i) = m
                    i = i + 1
                End If
                If Cells(n, m).Value <> "" Then
                    UpperRight = False
                End If
            End If
        End If
            
        If UpperLeft Then
            n = target.Row + (-1 * j)
            m = target.Column + (-1 * j)
            If n > 0 And m > 0 Then
                If Cells(n, m).Value = "" Or Cells(n, m).Font.Color = Piece_Color_B Then
                    ReDim Preserve result(1, i)
                    result(0, i) = n
                    result(1, i) = m
                    i = i + 1
                End If
                If Cells(n, m).Value <> "" Then
                    UpperLeft = False
                End If
            End If
        End If
        
        If BottomRight Then
            n = target.Row + (1 * j)
            m = target.Column + (1 * j)
            If n > 0 And m > 0 Then
                If Cells(n, m).Value = "" Or Cells(n, m).Font.Color = Piece_Color_B Then
                    ReDim Preserve result(1, i)
                    result(0, i) = n
                    result(1, i) = m
                    i = i + 1
                End If
                If Cells(n, m).Value <> "" Then
                    BottomRight = False
                End If
            End If
        End If
    
        If BottomLeft Then
            n = target.Row + (1 * j)
            m = target.Column + (-1 * j)
            If n > 0 And m > 0 Then
                If Cells(n, m).Value = "" Or Cells(n, m).Font.Color = Piece_Color_B Then
                    ReDim Preserve result(1, i)
                    result(0, i) = n
                    result(1, i) = m
                    i = i + 1
                End If
                If Cells(n, m).Value <> "" Then
                    BottomLeft = False
                End If
            End If
        End If
    
    Next
                
    
    If i = 0 Then
        ReDim Preserve result(0, 0)
    Else
    End If
    ���̃r�V���b�v = result()
End Function

Function ���̃��[�N(ByVal target As Range, Piece_Color_W As Long, Piece_Color_B As Long) As Long()
    
    Dim i, j As Long
    Dim result() As Long
            
    Dim Up, Right, Bottom, Left As Boolean
    Up = True
    Right = True
    Bottom = True
    Left = True
    
    i = 0
    
    For j = 1 To 7
        If Up Then
            n = target.Row + (-1 * j)
            m = target.Column
            If n > 0 And m > 0 Then
                If Cells(n, m).Value = "" Or Cells(n, m).Font.Color = Piece_Color_B Then
                    ReDim Preserve result(1, i)
                    result(0, i) = n
                    result(1, i) = m
                    i = i + 1
                End If
                If Cells(n, m).Value <> "" Then
                    Up = False
                End If
            End If
        End If
            
        If Right Then
            n = target.Row
            m = target.Column + (1 * j)
            If n > 0 And m > 0 Then
                If Cells(n, m).Value = "" Or Cells(n, m).Font.Color = Piece_Color_B Then
                    ReDim Preserve result(1, i)
                    result(0, i) = n
                    result(1, i) = m
                    i = i + 1
                End If
                If Cells(n, m).Value <> "" Then
                    Right = False
                End If
            End If
        End If
        
        If Bottom Then
            n = target.Row + (1 * j)
            m = target.Column
            If n > 0 And m > 0 Then
                If Cells(n, m).Value = "" Or Cells(n, m).Font.Color = Piece_Color_B Then
                    ReDim Preserve result(1, i)
                    result(0, i) = n
                    result(1, i) = m
                    i = i + 1
                End If
                If Cells(n, m).Value <> "" Then
                    Bottom = False
                End If
            End If
        End If
    
        If Left Then
            n = target.Row
            m = target.Column + (-1 * j)
            If n > 0 And m > 0 Then
                If Cells(n, m).Value = "" Or Cells(n, m).Font.Color = Piece_Color_B Then
                    ReDim Preserve result(1, i)
                    result(0, i) = n
                    result(1, i) = m
                    i = i + 1
                End If
                If Cells(n, m).Value <> "" Then
                    Left = False
                End If
            End If
        End If
    
    Next
                
    
    If i = 0 Then
        ReDim Preserve result(0, 0)
    Else
    End If
    ���̃��[�N = result()
End Function


Function ���̃N�C�[��(ByVal target As Range, Piece_Color_W As Long, Piece_Color_B As Long) As Long()
    
    Dim i, j As Long
    Dim result() As Long
            
    Dim Up, Right, Bottom, Left, UpperRight, UpperLeft, BottomRight, BottomLeft As Boolean
    Up = True
    Right = True
    Bottom = True
    Left = True
    UpperRight = True
    UpperLeft = True
    BottomRight = True
    BottomLeft = True
    
    i = 0
    
    For j = 1 To 7
        If Up Then
            n = target.Row + (-1 * j)
            m = target.Column
            If n > 0 And m > 0 Then
                If Cells(n, m).Value = "" Or Cells(n, m).Font.Color = Piece_Color_B Then
                    ReDim Preserve result(1, i)
                    result(0, i) = n
                    result(1, i) = m
                    i = i + 1
                End If
                If Cells(n, m).Value <> "" Then
                    Up = False
                End If
            End If
        End If
            
        If Right Then
            n = target.Row
            m = target.Column + (1 * j)
            If n > 0 And m > 0 Then
                If Cells(n, m).Value = "" Or Cells(n, m).Font.Color = Piece_Color_B Then
                    ReDim Preserve result(1, i)
                    result(0, i) = n
                    result(1, i) = m
                    i = i + 1
                End If
                If Cells(n, m).Value <> "" Then
                    Right = False
                End If
            End If
        End If
        
        If Bottom Then
            n = target.Row + (1 * j)
            m = target.Column
            If n > 0 And m > 0 Then
                If Cells(n, m).Value = "" Or Cells(n, m).Font.Color = Piece_Color_B Then
                    ReDim Preserve result(1, i)
                    result(0, i) = n
                    result(1, i) = m
                    i = i + 1
                End If
                If Cells(n, m).Value <> "" Then
                    Bottom = False
                End If
            End If
        End If
    
        If Left Then
            n = target.Row
            m = target.Column + (-1 * j)
            If n > 0 And m > 0 Then
                If Cells(n, m).Value = "" Or Cells(n, m).Font.Color = Piece_Color_B Then
                    ReDim Preserve result(1, i)
                    result(0, i) = n
                    result(1, i) = m
                    i = i + 1
                End If
                If Cells(n, m).Value <> "" Then
                    Left = False
                End If
            End If
        End If
    
        If UpperRight Then
            n = target.Row + (-1 * j)
            m = target.Column + (1 * j)
            If n > 0 And m > 0 Then
                If Cells(n, m).Value = "" Or Cells(n, m).Font.Color = Piece_Color_B Then
                    ReDim Preserve result(1, i)
                    result(0, i) = n
                    result(1, i) = m
                    i = i + 1
                End If
                If Cells(n, m).Value <> "" Then
                    UpperRight = False
                End If
            End If
        End If
            
        If UpperLeft Then
            n = target.Row + (-1 * j)
            m = target.Column + (-1 * j)
            If n > 0 And m > 0 Then
                If Cells(n, m).Value = "" Or Cells(n, m).Font.Color = Piece_Color_B Then
                    ReDim Preserve result(1, i)
                    result(0, i) = n
                    result(1, i) = m
                    i = i + 1
                End If
                If Cells(n, m).Value <> "" Then
                    UpperLeft = False
                End If
            End If
        End If
        
        If BottomRight Then
            n = target.Row + (1 * j)
            m = target.Column + (1 * j)
            If n > 0 And m > 0 Then
                If Cells(n, m).Value = "" Or Cells(n, m).Font.Color = Piece_Color_B Then
                    ReDim Preserve result(1, i)
                    result(0, i) = n
                    result(1, i) = m
                    i = i + 1
                End If
                If Cells(n, m).Value <> "" Then
                    BottomRight = False
                End If
            End If
        End If
    
        If BottomLeft Then
            n = target.Row + (1 * j)
            m = target.Column + (-1 * j)
            If n > 0 And m > 0 Then
                If Cells(n, m).Value = "" Or Cells(n, m).Font.Color = Piece_Color_B Then
                    ReDim Preserve result(1, i)
                    result(0, i) = n
                    result(1, i) = m
                    i = i + 1
                End If
                If Cells(n, m).Value <> "" Then
                    BottomLeft = False
                End If
            End If
        End If
    
    
    Next
                
    
    If i = 0 Then
        ReDim Preserve result(0, 0)
    Else
    End If
    ���̃N�C�[�� = result()
End Function

Function ���̃L���O(ByVal target As Range, Piece_Color_W As Long, Piece_Color_B As Long) As Long()
    
    Dim i, j As Long
    Dim result() As Long
    Dim arr_row() As Variant
    Dim arr_col() As Variant
    
    arr_row() = Array(-1, -1, -1, 0, 1, 1, 1, 0)
    arr_col() = Array(-1, 0, 1, 1, 1, 0, -1, -1)
    i = 0
        
    i = 0
    For j = 0 To 7
    n = target.Row + arr_row(j)
    m = target.Column + arr_col(j)
        If n > 0 And m > 0 Then
            If Cells(n, m).Value = "" Or Cells(n, m).Font.Color = Piece_Color_B Then
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
    ���̃L���O = result()
End Function
