Public Class Form2

    Private cText1, cText2, cText3, cText4, cText5, cText6, cText7, cText8, cText9 As String

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        Opisanie()
    End Sub


    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        Opisanie()
    End Sub

    Private Sub Opisanie()
        
        If RadioButton1.Checked = True Then
            cText1 = My.Resources.Res1.f1
            cText2 = My.Resources.Res1.f2
            cText3 = My.Resources.Res1.f3
            cText4 = My.Resources.Res1.f4
            cText5 = My.Resources.Res1.f5
            cText6 = My.Resources.Res1.f6
            cText7 = My.Resources.Res1.f7
            cText8 = My.Resources.Res1.f8
            cText9 = My.Resources.Res1.f9
            TextBox1.Text = Replace(cText1 & cText2 & cText3 & cText4 & cText5 & cText6 & cText7 & cText8 & cText9, "@@@@@", "")
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
            TextBox1.Text = cText1 & cText2 & cText3 & cText4 & cText5 & cText6 & cText7 & cText8 & cText9
        End If


    End Sub


    Private Sub Form2_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        RadioButton2.Checked = True
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        TextBox1.Text = Replace(cText1, "@@@@@", "")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox1.Text = Replace(cText2, "@@@@@", "")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        TextBox1.Text = Replace(cText3, "@@@@@", "")
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        TextBox1.Text = Replace(cText4, "@@@@@", "")
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        TextBox1.Text = Replace(cText5, "@@@@@", "")
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        TextBox1.Text = Replace(cText6, "@@@@@", "")
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        TextBox1.Text = Replace(cText7, "@@@@@", "")
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        TextBox1.Text = Replace(cText8, "@@@@@", "")
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        TextBox1.Text = Replace(cText9, "@@@@@", "")
    End Sub

    Private Sub RadioButton1_Click(sender As Object, e As EventArgs) Handles RadioButton1.Click
        Opisanie()
    End Sub

    Private Sub RadioButton2_Click(sender As Object, e As EventArgs) Handles RadioButton2.Click
        Opisanie()
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Icon = Form1.Icon
    End Sub
End Class