Module Module1

    ' *************** input variables
    Private sheetName As String = ""
    ' *************** input variables

    ' *************** output variables
    Private isRudOrRuePresent As Boolean
    ' *************** output variables

    Sub Main()

        If sheetName.Contains("RUD") Then
            isRudOrRuePresent = True
        Else
            isRudOrRuePresent = False
        End If

    End Sub

End Module
