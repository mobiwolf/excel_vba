Sub copyrandomrows()
    Const initrows& = 564
    Const randrows& = 100
    Dim b(initrows) As Boolean
    Dim c&, x&
    Do
        x = Application.RandBetween(1, initrows)
        If Not b(x) Then
            c = c + 1
            b(x) = True
            Sheets("sheet1").Rows(x).Copy Sheets("sheet2").Cells(c, 1)
        End If
    Loop Until c = randrows
End Sub