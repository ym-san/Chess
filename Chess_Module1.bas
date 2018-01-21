Attribute VB_Name = "Module1"
Public Sub ”wŒiF‰Šú‰»()

Dim Interior_Color_P As Long
Dim Interior_Color_B As Long
Interior_Color_P = RGB(244, 176, 132) '”wŒiF_ƒy[ƒ‹ƒIƒŒƒ“ƒW
Interior_Color_B = RGB(198, 89, 17) '”wŒiF_’ƒ
Dim i, j As Integer
            
            '--”wŒiF‰Šú‰»--
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
            '--”wŒiF‰Šú‰»‚¨‚í‚è--

End Sub


