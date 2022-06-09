Module Module1

    Private nameRun As String = "DB - EPAP"

    Private resultNameRun As String

    Sub Main()

        If nameRun.Contains("RUD") Then
            nameRun = nameRun.Replace("DB -", "Citibank -")
        End If

        resultNameRun = nameRun

    End Sub

End Module
