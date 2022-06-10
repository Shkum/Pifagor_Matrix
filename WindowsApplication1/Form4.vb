Imports System.Math

Public Class Form4

    Dim cStop As Boolean

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


    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        If Asc(e.KeyChar) = 8 Then Exit Sub

        Try
            If Asc(e.KeyChar) <> 8 And Asc(e.KeyChar) <> 49 Then
                e.KeyChar = ""
                Exit Sub
            End If
            If CheckLength() = True Then e.KeyChar = ""
        Catch

        End Try
    End Sub

    Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        If Asc(e.KeyChar) = 8 Then Exit Sub
        Try
            If Asc(e.KeyChar) <> 8 And Asc(e.KeyChar) <> 50 Then
                e.KeyChar = ""
                Exit Sub
            End If
            If CheckLength() = True Then e.KeyChar = ""
        Catch

        End Try
    End Sub

    Private Sub TextBox5_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        If Asc(e.KeyChar) = 8 Then Exit Sub
        Try
            If Asc(e.KeyChar) <> 8 And Asc(e.KeyChar) <> 51 Then
                e.KeyChar = ""
                Exit Sub
            End If
            If CheckLength() = True Then e.KeyChar = ""
        Catch

        End Try
    End Sub

    Private Sub TextBox6_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox6.KeyPress
        If Asc(e.KeyChar) = 8 Then Exit Sub
        Try
            If Asc(e.KeyChar) <> 8 And Asc(e.KeyChar) <> 52 Then
                e.KeyChar = ""
                Exit Sub
            End If
            If CheckLength() = True Then e.KeyChar = ""
        Catch

        End Try
    End Sub

    Private Sub TextBox7_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox7.KeyPress
        If Asc(e.KeyChar) = 8 Then Exit Sub
        Try
            If Asc(e.KeyChar) <> 8 And Asc(e.KeyChar) <> 53 Then
                e.KeyChar = ""
                Exit Sub
            End If
            If CheckLength() = True Then e.KeyChar = ""
        Catch

        End Try
    End Sub

    Private Sub TextBox8_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox8.KeyPress
        If Asc(e.KeyChar) = 8 Then Exit Sub
        Try
            If Asc(e.KeyChar) <> 8 And Asc(e.KeyChar) <> 54 Then
                e.KeyChar = ""
                Exit Sub
            End If
            If CheckLength() = True Then e.KeyChar = ""
        Catch

        End Try
    End Sub

    Private Sub TextBox9_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox9.KeyPress
        If Asc(e.KeyChar) = 8 Then Exit Sub
        Try
            If Asc(e.KeyChar) <> 8 And Asc(e.KeyChar) <> 55 Then
                e.KeyChar = ""
                Exit Sub
            End If
            If CheckLength() = True Then e.KeyChar = ""
        Catch

        End Try
    End Sub

    Private Sub TextBox10_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox10.KeyPress
        If Asc(e.KeyChar) = 8 Then Exit Sub
        Try
            If Asc(e.KeyChar) <> 8 And Asc(e.KeyChar) <> 56 Then
                e.KeyChar = ""
                Exit Sub
            End If
            If CheckLength() = True Then e.KeyChar = ""
        Catch

        End Try
    End Sub

    Private Sub TextBox11_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox11.KeyPress
        If Asc(e.KeyChar) = 8 Then Exit Sub
        Try
            If Asc(e.KeyChar) <> 8 And Asc(e.KeyChar) <> 57 Then
                e.KeyChar = ""
                Exit Sub
            End If
            If CheckLength() = True Then e.KeyChar = ""
        Catch

        End Try
    End Sub

    Private Sub ClearMatrix()
        Button1.Text = ""
        Button2.Text = ""
        Button3.Text = ""
        Button4.Text = ""
        Button5.Text = ""
        Button6.Text = ""
        Button7.Text = ""
        Button8.Text = ""
        Button9.Text = ""
    End Sub

    Private Function CheckLength() As Boolean
        Try
            Dim s As Integer
            s = TextBox3.Text.Length + TextBox4.Text.Length + TextBox5.Text.Length + TextBox6.Text.Length + _
                TextBox7.Text.Length + TextBox8.Text.Length + TextBox9.Text.Length + TextBox10.Text.Length + _
                TextBox11.Text.Length
            If s = 15 Then
                MsgBox("В данный момент у вас 15 цифр" & vbCrLf & "Максимальное количество цифр не может быть больше 15", vbInformation, "Поиск дат")
                Return True
            End If
            If s > 15 Then Return True
        Catch ex As Exception
            Return False
        End Try

    End Function


    Public Sub Find_Calc(data As String)
        Dim cDte, cTot, cName As String
        Dim c1, c2, c3, c4, cCount As Integer
        cDte = Replace(data, "0", "")
        For cCount = 1 To cDte.Length
            c1 = c1 + Mid(cDte, cCount, 1)
            c1 = Abs(c1)
        Next
        For cCount = 1 To Len(c1.ToString)
            c2 = c2 + Mid(c1, cCount, 1)
            c2 = Abs(c2)
        Next
        c3 = c1 - Mid(cDte, 1, 1) * 2
        c3 = Abs(c3)
        For cCount = 1 To Len(c3.ToString)
            c4 = c4 + Mid(c3, cCount, 1)
        Next
        cTot = cDte & c1 & c2 & c3 & c4

        For cCount = 0 To cTot.Length - Replace(cTot, "1", "").Length - 1
            cName = cName & "1"
        Next
        Button1.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "2", "").Length - 1
            cName = cName & "2"
        Next
        Button2.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "3", "").Length - 1
            cName = cName & "3"
        Next
        Button3.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "4", "").Length - 1
            cName = cName & "4"
        Next
        Button4.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "5", "").Length - 1
            cName = cName & "5"
        Next
        Button5.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "6", "").Length - 1
            cName = cName & "6"
        Next
        Button6.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "7", "").Length - 1
            cName = cName & "7"
        Next
        Button7.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "8", "").Length - 1
            cName = cName & "8"
        Next
        Button8.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "9", "").Length - 1
            cName = cName & "9"
        Next
        Button9.Text = cName

    End Sub


    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        If CheckDate(TextBox1.Text) = True And CheckDate(TextBox2.Text) = True Then
            Dim cDiff As Long, Date1 As Date, Date2 As Date, Date3 As String
            Dim cDay, cMonth, cYear As Integer, cCount As Long

            Button11.Enabled = True
            cStop = False
            GroupBox1.Enabled = False
            GroupBox2.Enabled = False
            TextBox1.Enabled = False
            TextBox2.Enabled = False
            Button10.Enabled = False
            Me.Focus()

            ListBox1.Items.Clear()
            ClearMatrix()

            cDay = Mid(TextBox1.Text, 1, 2)
            cMonth = Mid(TextBox1.Text, 3, 2)
            cYear = Mid(TextBox1.Text, 5, 4)
            Date1 = CDate(cDay & "/" & cMonth & "/" & cYear)

            cDay = Mid(TextBox2.Text, 1, 2)
            cMonth = Mid(TextBox2.Text, 3, 2)
            cYear = Mid(TextBox2.Text, 5, 4)
            Date2 = CDate(cDay & "/" & cMonth & "/" & cYear)

            cDiff = DateDiff(DateInterval.Day, Date1, Date2)
            For cCount = 0 To cDiff
                Application.DoEvents()
                Date3 = DateAdd("d", cCount, Date1)
                Date3 = Format(CDate(Date3), "ddMMyyyy")
                Find_Calc(Date3)

                ProverkaSovpadeniy(Date3)

                Me.Text = "Поиск дат " & Date3 & " - " & Round(100 * cCount / cDiff, 1) & "%"

                If cStop = True Then
                    ClearMatrix()
                    Exit For
                End If

            Next
            Me.Text = "Поиск дат"
        Else
            MsgBox("Проверте правильность начальной и конечной дат", vbInformation, "Что то не так...")
        End If
        Button11.Enabled = False
        GroupBox1.Enabled = True
        GroupBox2.Enabled = True
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        Button10.Enabled = True
        ClearMatrix()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Button11.Enabled = False
        cStop = True
    End Sub

    Private Sub Button11_MouseUp(sender As Object, e As MouseEventArgs) Handles Button11.MouseUp
        cStop = True
        Button11.Enabled = False
    End Sub

    Private Sub ProverkaSovpadeniy(data As String)

        Dim c1, c2, c3, c4, c5, c6, c7, c8, c9 As Boolean
        If CheckBox1.Checked = True And TextBox3.Text.Length = Button1.Text.Length Then c1 = True
        If CheckBox2.Checked = True And TextBox4.Text.Length = Button2.Text.Length Then c2 = True
        If CheckBox3.Checked = True And TextBox5.Text.Length = Button3.Text.Length Then c3 = True
        If CheckBox4.Checked = True And TextBox6.Text.Length = Button4.Text.Length Then c4 = True
        If CheckBox5.Checked = True And TextBox7.Text.Length = Button5.Text.Length Then c5 = True
        If CheckBox6.Checked = True And TextBox8.Text.Length = Button6.Text.Length Then c6 = True
        If CheckBox7.Checked = True And TextBox9.Text.Length = Button7.Text.Length Then c7 = True
        If CheckBox8.Checked = True And TextBox10.Text.Length = Button8.Text.Length Then c8 = True
        If CheckBox9.Checked = True And TextBox11.Text.Length = Button9.Text.Length Then c9 = True
        If CheckBox1.Checked = False Then c1 = True
        If CheckBox2.Checked = False Then c2 = True
        If CheckBox3.Checked = False Then c3 = True
        If CheckBox4.Checked = False Then c4 = True
        If CheckBox5.Checked = False Then c5 = True
        If CheckBox6.Checked = False Then c6 = True
        If CheckBox7.Checked = False Then c7 = True
        If CheckBox8.Checked = False Then c8 = True
        If CheckBox9.Checked = False Then c9 = True
        If c1 = True And c2 = True And c3 = True And c4 = True And c5 = True And c6 = True And c7 = True And c8 = True And c9 = True Then
            ListBox1.Items.Add(data)
        End If
        
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        Find_Calc(ListBox1.SelectedItem)
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Hide()
    End Sub

    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Icon = Form1.Icon
    End Sub
End Class