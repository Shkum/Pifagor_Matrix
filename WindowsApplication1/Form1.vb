Public Class Form1

    Dim cText1, cText2, cText3, cText4, cText5, cText6, cText7, cText8, cText9 As String

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        Try
            If Asc(e.KeyChar) <> 8 And (Asc(e.KeyChar) > 57 Or Asc(e.KeyChar) < 48) Then
                e.KeyChar = ""
                Exit Sub
            End If

        Catch

        End Try
    End Sub

    Private Sub TextBox1_KeyUp(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyUp
        Dim cYear As Integer

        If CheckDate(TextBox1.Text) = True And TextBox1.Text.Length = 8 Then
            ToolStripStatusLabel1.Text = "Расчёт Матрицы: " & TextBox1.Text
            Pif_Calc(TextBox1.Text)
            Opisanie()
            cYear = Mid(TextBox1.Text, 5, 4)
            If cYear > 1999 Then
                MsgBox(My.Resources.Res1.god2000, vbInformation, "Информация о дате рождения...")
            End If
            TextBox2.Text = "Щелкните на цифры для получения описания"
        Else
            ClearMatrix()
        End If
    End Sub

    Private Sub ClearMatrix()
        ToolStripStatusLabel1.Text = "Расчёт Матрицы: -"
        TextBox2.Text = ""
        Button1.Text = "-"
        Button2.Text = "-"
        Button3.Text = "-"
        Button4.Text = "-"
        Button5.Text = "-"
        Button6.Text = "-"
        Button7.Text = "-"
        Button8.Text = "-"
        Button9.Text = "-"
        Label3.Text = "Первое рабочее число: -"
        Label4.Text = "Второе рабочее число: -"
        Label5.Text = "Третье рабочее число: -"
        Label6.Text = "Четвертое рабочее число: -"
        Label7.Text = "Самооценка: -"
        Label8.Text = "Материальное благополучие: -"
        Label9.Text = "Талант: -"
        Label10.Text = "Целеустремлённость: -"
        Label11.Text = "Семейность: -"
        Label12.Text = "Привычки и привязанность: -"
        Label13.Text = "Плотские интересы: -"
        Label14.Text = "Духовные интересы: -"
        Label15.Text = "Воплощение №-"
        Me.Text = "Расчёт Матрицы Пифагора"
        cText1 = ""
        cText2 = ""
        cText3 = ""
        cText4 = ""
        cText5 = ""
        cText6 = ""
        cText7 = ""
        cText8 = ""
        cText9 = ""

    End Sub

    Private Sub Opisanie()

        If RadioButton1.Checked = True Then
            Dim StrArr() As String

            cText1 = My.Resources.Res1.f1
            StrArr = Split(cText1, "@@@@@")
            Select Case Button1.Text.Length
                Case 1
                    cText1 = StrArr(1)
                Case 2
                    cText1 = StrArr(2)
                Case 3
                    cText1 = StrArr(3)
                Case 4
                    cText1 = StrArr(4)
                Case 5
                    cText1 = StrArr(5)
                Case Is > 5
                    cText1 = StrArr(6)
            End Select
            

            cText2 = My.Resources.Res1.f2
            StrArr = Split(cText2, "@@@@@")
            Select Case Button2.Text.Length
                Case 0
                    cText2 = StrArr(1)
                Case 1
                    cText2 = StrArr(2)
                Case 2
                    cText2 = StrArr(3)
                Case 3
                    cText2 = StrArr(4)
                Case Is > 3
                    cText2 = StrArr(5)
            End Select


            cText3 = My.Resources.Res1.f3
            StrArr = Split(cText3, "@@@@@")
            Select Case Button3.Text.Length
                Case 0
                    cText3 = StrArr(1)
                Case 1
                    cText3 = StrArr(2)
                Case 2
                    cText3 = StrArr(3)
                Case 3
                    cText3 = StrArr(4)
                Case Is > 3
                    cText3 = StrArr(5)
            End Select

        
            cText4 = My.Resources.Res1.f4
            StrArr = Split(cText4, "@@@@@")
            Select Case Button4.Text.Length
                Case 0
                    cText4 = StrArr(1)
                Case 1
                    cText4 = StrArr(2)
                Case Is > 1
                    cText4 = StrArr(3)
            End Select
            
            
            cText5 = My.Resources.Res1.f5
            StrArr = Split(cText5, "@@@@@")
            Select Case Button5.Text.Length
                Case 0
                    cText5 = StrArr(1)
                Case 1
                    cText5 = StrArr(2)
                Case 2
                    cText5 = StrArr(3)
                Case Is > 2
                    cText5 = StrArr(4)
            End Select


            cText6 = My.Resources.Res1.f6
            StrArr = Split(cText6, "@@@@@")
            Select Case Button6.Text.Length
                Case 0
                    cText6 = StrArr(1)
                Case 1
                    cText6 = StrArr(2)
                Case 2
                    cText6 = StrArr(3)
                Case Is > 2
                    cText6 = StrArr(4)
            End Select


            cText7 = My.Resources.Res1.f7
            StrArr = Split(cText7, "@@@@@")
            Select Case Button7.Text.Length
                Case 0
                    cText7 = StrArr(1)
                Case 1
                    cText7 = StrArr(2)
                Case 2
                    cText7 = StrArr(3)
                Case Is > 2
                    cText7 = StrArr(4)
            End Select


            cText8 = My.Resources.Res1.f8
            StrArr = Split(cText8, "@@@@@")
            Select Case Button8.Text.Length
                Case 0
                    cText8 = StrArr(1)
                Case 1
                    cText8 = StrArr(2)
                Case Is > 1
                    cText8 = StrArr(3)
            End Select


            cText9 = My.Resources.Res1.f9
            StrArr = Split(cText9, "@@@@@")
            Select Case Button9.Text.Length
                Case 0
                    cText9 = StrArr(1)
                Case 1
                    cText9 = StrArr(2)
                Case 2
                    cText9 = StrArr(3)
                Case Is > 2
                    cText9 = StrArr(4)
            End Select
        End If


        If RadioButton2.Checked = True Then
            cText1 = My.Resources.Res1.g1
            cText2 = My.Resources.Res1.g2
            cText3 = My.Resources.Res1.g3
            cText4 = My.Resources.Res1.g4
            cText5 = My.Resources.Res1.g5
            cText6 = My.Resources.Res1.g6
            cText7 = My.Resources.Res1.g7
            cText8 = My.Resources.Res1.g8
            cText9 = My.Resources.Res1.g9
        End If



    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Form2.Show()
        Form2.TextBox1.Text = cText1 & vbCrLf & vbCrLf & vbCrLf & cText2 & vbCrLf & vbCrLf & vbCrLf & cText3 & vbCrLf & vbCrLf & vbCrLf & cText4 & vbCrLf & vbCrLf & vbCrLf & cText5 & vbCrLf & vbCrLf & vbCrLf & cText6 & vbCrLf & vbCrLf & vbCrLf & cText7 & vbCrLf & vbCrLf & vbCrLf & cText8 & vbCrLf & vbCrLf & vbCrLf & cText9
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox2.Text = ctext1
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox2.Text = ctext2
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox2.Text = ctext3
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TextBox2.Text = ctext4
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        TextBox2.Text = ctext5
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        TextBox2.Text = ctext6
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        TextBox2.Text = ctext7
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        TextBox2.Text = ctext8
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        TextBox2.Text = ctext9
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If CheckDate(TextBox1.Text) = True And TextBox1.Text.Length = 8 Then
            ToolStripStatusLabel1.Text = "Расчёт Матрицы: " & TextBox1.Text
            Pif_Calc(TextBox1.Text)
            Opisanie()
        Else
            ClearMatrix()
        End If
        TextBox2.Text = ""
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If CheckDate(TextBox1.Text) = True And TextBox1.Text.Length = 8 Then
            ToolStripStatusLabel1.Text = "Расчёт Матрицы: " & TextBox1.Text
            Pif_Calc(TextBox1.Text)
            Opisanie()
        Else
            ClearMatrix()
        End If
        TextBox2.Text = ""
    End Sub


    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Close()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        MsgBox("Программу написал Шкумат Сергей (pepelnici@meta.ua)" & _
               vbCrLf & vbCrLf & "По книге Александра Александрова" & _
               vbCrLf & "'Дата рождения - ключ к пониманию'" & vbCrLf & vbCrLf & _
               My.Resources.Res1.Abt, vbInformation, "О программе")
    End Sub

    Private Sub Label3_Click(sender As Object, e As EventArgs) Handles Label3.Click
        TextBox2.Text = My.Resources.Res1.PRC

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click
        TextBox2.Text = My.Resources.Res1.VRC
    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click
        TextBox2.Text = My.Resources.Res1.TRC
    End Sub

    Private Sub Label6_Click(sender As Object, e As EventArgs) Handles Label6.Click
        TextBox2.Text = My.Resources.Res1.TRC
    End Sub

    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles Label7.Click
        TextBox2.Text = My.Resources.Res1.SA
    End Sub

    Private Sub Label8_Click(sender As Object, e As EventArgs) Handles Label8.Click
        TextBox2.Text = My.Resources.Res1.MA
    End Sub

    Private Sub Label9_Click(sender As Object, e As EventArgs) Handles Label9.Click
        TextBox2.Text = My.Resources.Res1.TA
    End Sub


    Private Sub Label10_Click(sender As Object, e As EventArgs) Handles Label10.Click
        TextBox2.Text = My.Resources.Res1.CE
    End Sub

    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click
        TextBox2.Text = My.Resources.Res1.SE
    End Sub

    Private Sub Label12_Click(sender As Object, e As EventArgs) Handles Label12.Click
        TextBox2.Text = My.Resources.Res1.PR
    End Sub

    Private Sub Label13_Click(sender As Object, e As EventArgs) Handles Label13.Click
        TextBox2.Text = My.Resources.Res1.PL
    End Sub

    Private Sub Label14_Click(sender As Object, e As EventArgs) Handles Label14.Click
        TextBox2.Text = My.Resources.Res1.DU
    End Sub

    Private Sub Label15_Click(sender As Object, e As EventArgs) Handles Label15.Click
        TextBox2.Text = My.Resources.Res1.VO
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Form2.Hide()
        Form4.Hide()
        Form3.Show(Me)
        Form3.TextBox1.Text = TextBox1.Text
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Form2.Hide()
        Form3.Hide()
        Form4.Show(Me)

    End Sub

End Class
