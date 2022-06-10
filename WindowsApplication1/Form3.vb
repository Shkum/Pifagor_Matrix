Imports System.Math
Public Class Form3
    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        Try
            If Asc(e.KeyChar) <> 8 And (Asc(e.KeyChar) > 57 Or Asc(e.KeyChar) < 48) Then
                e.KeyChar = ""
                Exit Sub
            End If
        Catch
        End Try
    End Sub
    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        Try
            If Asc(e.KeyChar) <> 8 And (Asc(e.KeyChar) > 57 Or Asc(e.KeyChar) < 48) Then
                e.KeyChar = ""
                Exit Sub
            End If
        Catch
        End Try
    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Icon = Form1.Icon
    End Sub

    Private Sub ClearMatrix1()
        Label7.Text = "-"
        Label9.Text = "-"
        Label14.Text = "-"
        Label4.Text = "-"
        Label16.Text = "-"
        Button2.Text = "-"
        Button3.Text = "-"
        Button4.Text = "-"
        Button5.Text = "-"
        Button6.Text = "-"
        Button7.Text = "-"
        Button8.Text = "-"
        Button9.Text = "-"
        Button10.Text = "-"
       End Sub
    Private Sub ClearMatrix2()
        Label8.Text = "-"
        Label11.Text = "-"
        Label14.Text = "-"
        Label4.Text = "-"
        Label16.Text = "-"
       Button11.Text = "-"
        Button12.Text = "-"
        Button13.Text = "-"
        Button14.Text = "-"
        Button15.Text = "-"
        Button16.Text = "-"
        Button17.Text = "-"
        Button18.Text = "-"
        Button19.Text = "-"
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Dim cYear As Integer
        If CheckDate(TextBox1.Text) = True And TextBox1.Text.Length = 8 Then
            Pif_Calc_Muz(TextBox1.Text)
            If CheckDate(TextBox1.Text) = True And CheckDate(TextBox2.Text) = True And TextBox1.Text.Length = 8 And TextBox2.Text.Length = 8 Then Raschet()
            cYear = Mid(TextBox1.Text, 5, 4)
            If cYear > 1999 Then
                MsgBox(My.Resources.Res1.god2000, vbInformation, "Информация о дате рождения...")
            End If
        Else
            ClearMatrix1()
        End If
    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Dim cYear As Integer
        If CheckDate(TextBox2.Text) = True And TextBox2.Text.Length = 8 Then
            cYear = Mid(TextBox1.Text, 5, 4)
            If cYear > 1999 Then
                MsgBox(My.Resources.Res1.god2000, vbInformation, "Информация о дате рождения...")
            End If
            Pif_Calc_Zena(TextBox2.Text)
            If CheckDate(TextBox1.Text) = True And CheckDate(TextBox2.Text) = True And TextBox1.Text.Length = 8 And TextBox2.Text.Length = 8 Then Raschet()
        Else
            ClearMatrix2()
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        MsgBox(My.Resources.Res1.About, vbInformation, "О расчёте ...")
    End Sub
    Private Sub Raschet()
        Dim pdm, ssm, psm, BSM As Integer
        Dim ddm, sm, crm, DSM As Integer
        Dim pdz, ssz, psz, BSZ As Integer
        Dim ddz, sz, crz, DSZ As Integer
        pdm = Button3.Text.Length + Button5.Text.Length + Button7.Text.Length
        ssm = Button2.Text.Length + Button5.Text.Length + Button8.Text.Length
        psm = Button3.Text.Length + Button6.Text.Length + Button9.Text.Length
        BSM = pdm * ssm * psm
        Label9.Text = BSM
        ddm = Button10.Text.Length + Button5.Text.Length + Button9.Text.Length
        sm = Button10.Text.Length + Button2.Text.Length + Button3.Text.Length
        crm = Button10.Text.Length + Button4.Text.Length + Button7.Text.Length
        DSM = ddm * sm * crm
        Label7.Text = DSM

        pdz = Button17.Text.Length + Button15.Text.Length + Button13.Text.Length
        ssz = Button18.Text.Length + Button15.Text.Length + Button12.Text.Length
        psz = Button17.Text.Length + Button14.Text.Length + Button11.Text.Length
        BSZ = pdz * ssz * psz
        Label11.Text = BSZ
        ddz = Button19.Text.Length + Button15.Text.Length + Button11.Text.Length
        sz = Button19.Text.Length + Button18.Text.Length + Button17.Text.Length
        crz = Button19.Text.Length + Button16.Text.Length + Button13.Text.Length
        DSZ = ddz * sz * crz
        Label8.Text = DSZ
        Label14.Text = Round(DSM * DSZ / 365, 1)
        Label4.Text = Round(BSM * BSZ / 365, 1)
        Label16.Text = Round(DSM * DSZ / 365 + BSM * BSZ / 365, 1)
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        MsgBox("Не имеет значения с какой стороны будут введены даты рождения мужа и жены. Приписки 'муж' и 'жена' сделаны для удобства отображения информации.", vbInformation, "О расчёте...")
    End Sub
End Class