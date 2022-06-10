Module Module2
    Public Function CheckDate(str As String) As Boolean
        Try
            Dim cDay, cMonth, cYear As Integer
            cDay = Mid(str, 1, 2)
            cMonth = Mid(str, 3, 2)
            cYear = Mid(str, 5, 4)
            If IsDate(cDay & "/" & cMonth & "/" & cYear) And cYear < 3000 And cYear > 1000 Then

                Return True
            End If
        Catch
            Return False
        End Try
    End Function
End Module
