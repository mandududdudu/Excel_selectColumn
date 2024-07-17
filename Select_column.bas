Attribute VB_Name = "Module1"
Sub selecct_column()
    input_c = Application.InputBox("please input columns", Type:=2)
    input_len = Len(input_c)
    If input_len > 3 Then
        MsgBox "not support"
    ElseIf 0 < input_len < 3 Then
        If IsNumeric(input_c) = True Then
            If input_c < 200 Then
                MsgBox " cloumn name: " & input_c
                Columns(Int(input_c)).Select
            Else
                MsgBox "not support"
            End If
        ElseIf Asc(UCase(input_c)) >= 65 And Asc(UCase(input_c)) <= 90 Then
            new_input = UCase(input_c)
            If Len(new_input) = 1 Then
                MsgBox "your column is " & new_input
                Columns(newinput).Select
            Else
                first = Left(new_input, 1)
                last = Right(new_input, 1)
                MsgBox "your column is " & first & last
                Columns(new_input).Select
            End If
        Else
            MsgBox "idiot"
        End If
    Else
        MsgBox "invalid"
    End If
End Sub
