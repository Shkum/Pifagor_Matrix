Imports System.Math
Module Module3
    Public Sub Pif_Calc_Muz(data As String)
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
        Form3.Button10.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "2", "").Length - 1
            cName = cName & "2"
        Next
        Form3.Button2.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "3", "").Length - 1
            cName = cName & "3"
        Next
        Form3.Button3.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "4", "").Length - 1
            cName = cName & "4"
        Next
        Form3.Button4.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "5", "").Length - 1
            cName = cName & "5"
        Next
        Form3.Button5.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "6", "").Length - 1
            cName = cName & "6"
        Next
        Form3.Button6.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "7", "").Length - 1
            cName = cName & "7"
        Next
        Form3.Button7.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "8", "").Length - 1
            cName = cName & "8"
        Next
        Form3.Button8.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "9", "").Length - 1
            cName = cName & "9"
        Next
        Form3.Button9.Text = cName
    End Sub

    Public Sub Pif_Calc_Zena(data As String)
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
        Form3.Button19.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "2", "").Length - 1
            cName = cName & "2"
        Next
        Form3.Button18.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "3", "").Length - 1
            cName = cName & "3"
        Next
        Form3.Button17.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "4", "").Length - 1
            cName = cName & "4"
        Next
        Form3.Button16.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "5", "").Length - 1
            cName = cName & "5"
        Next
        Form3.Button15.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "6", "").Length - 1
            cName = cName & "6"
        Next
        Form3.Button14.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "7", "").Length - 1
            cName = cName & "7"
        Next
        Form3.Button13.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "8", "").Length - 1
            cName = cName & "8"
        Next
        Form3.Button12.Text = cName

        cName = ""
        For cCount = 0 To cTot.Length - Replace(cTot, "9", "").Length - 1
            cName = cName & "9"
        Next
        Form3.Button11.Text = cName
    End Sub

End Module
