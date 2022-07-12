Module Module1

    ' *************** input variables
    Private fileName As String = "ero .Общий план_ХХ-ХХ_недели_RU17.xlsb"
    Private currentTime As DateTime = DateTime.Now
    ' *************** input variables

    ' *************** output variables
    Private resultPlanFileName As String
    ' *************** output variables

    Sub Main()
        Dim indexPoint As Integer = 0
        Dim currentTimeToString As String
        Console.WriteLine(fileName.Length)
        For i As Integer = fileName.Length To 1 Step -1
            If CChar(Mid(fileName, i, 1)) = Chr(46) Then
                indexPoint = i
                Exit For
            End If
        Next

        'Console.WriteLine(RemoveUnnecessaryChar(Replace(Mid(currentTime.ToString(), 11), ":", "")))
        currentTimeToString = RemoveUnnecessaryChar(Left(currentTime.ToString(), 10)) & "_" & RemoveUnnecessaryChar(Replace(Mid(currentTime.ToString(), 11), ":", ""))
        resultPlanFileName = Left(fileName, indexPoint - 1) & "_" & currentTimeToString & Mid(fileName, indexPoint)
        Console.WriteLine(resultPlanFileName)
    End Sub

    Private Function RemoveUnnecessaryChar(str As String) As String
        'Chr(8)  Backspace character
        'Chr(32) Space
        'Chr(34) Quotation Mark
        'Chr(160)    Non-breaking space
        Return Replace(Replace(Replace(Replace(Replace(Replace(Replace(str, Chr(13), ""), Chr(7), ""), Chr(9), ""), Chr(11), ""), Chr(160), ""), Chr(32), ""), Chr(46), "")

    End Function

End Module
