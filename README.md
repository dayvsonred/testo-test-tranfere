# testo-test-tranfere
teste

Public Function Fix(ByVal Number As Double) As Double
    If Number >= 0 Then
        Return Math.Floor(Number)
    Else
        Return -Math.Floor(-Number)
    End If
End Function

Public Function Trunca(ByVal Num_Dec As Integer, ByVal Numero As Double) As Double
    Dim multiplicador As Double = Math.Pow(10, Num_Dec)
    Return Fix(Numero * multiplicador) / multiplicador
End Function
