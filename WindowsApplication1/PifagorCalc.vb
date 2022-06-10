Imports System.Math 
Module Module1

    Public Sub Pif_Calc(data As String)
        Dim cDte, cTot, cName As String
        Dim c1, c2, c3, c4, cVopl, cCount As Integer
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
        cVopl = cTot.Length

        For cCount = 0 To cTot.Length - Replace(cTot, "1", "").Length - 1
            cName = cName & "1"
        Next
        Form1.Button1.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "2", "").Length - 1
            cName = cName & "2"
        Next
        Form1.Button2.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "3", "").Length - 1
            cName = cName & "3"
        Next
        Form1.Button3.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "4", "").Length - 1
            cName = cName & "4"
        Next
        Form1.Button4.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "5", "").Length - 1
            cName = cName & "5"
        Next
        Form1.Button5.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "6", "").Length - 1
            cName = cName & "6"
        Next
        Form1.Button6.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "7", "").Length - 1
            cName = cName & "7"
        Next
        Form1.Button7.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "8", "").Length - 1
            cName = cName & "8"
        Next
        Form1.Button8.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "9", "").Length - 1
            cName = cName & "9"
        Next
        Form1.Button9.Text = cName

        Form1.Label3.Text = "Первое рабочее число: " & c1
        Form1.Label4.Text = "Второе рабочее число: " & c2
        Form1.Label5.Text = "Третье рабочее число: " & c3
        Form1.Label6.Text = "Четвертое рабочее число: " & c4
        Form1.Label7.Text = "Самооценка: " & Form1.Button1.Text.Length + Form1.Button2.Text.Length + Form1.Button3.Text.Length
        Form1.Label8.Text = "Материальное благополучие: " & Form1.Button4.Text.Length + Form1.Button5.Text.Length + Form1.Button6.Text.Length
        Form1.Label9.Text = "Талант: " & Form1.Button7.Text.Length + Form1.Button8.Text.Length + Form1.Button9.Text.Length
        Form1.Label10.Text = "Целеустремлённость: " & Form1.Button1.Text.Length + Form1.Button4.Text.Length + Form1.Button7.Text.Length
        Form1.Label11.Text = "Семейность: " & Form1.Button2.Text.Length + Form1.Button5.Text.Length + Form1.Button8.Text.Length
        Form1.Label12.Text = "Привычки и привязанность: " & Form1.Button3.Text.Length + Form1.Button6.Text.Length + Form1.Button9.Text.Length
        Form1.Label13.Text = "Плотские интересы: " & Form1.Button3.Text.Length + Form1.Button5.Text.Length + Form1.Button7.Text.Length
        Form1.Label14.Text = "Духовные интересы: " & Form1.Button1.Text.Length + Form1.Button5.Text.Length + Form1.Button9.Text.Length
        Form1.Label15.Text = "Воплощение №" & cVopl
        Form1.Text = "Расчёт Матрицы Пифагора: " & data
    End Sub

End Module
