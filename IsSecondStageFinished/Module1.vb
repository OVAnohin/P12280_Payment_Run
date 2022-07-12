Module Module1

    ' *************** input variables
    Private localFolder As String = "C:\Temp\WorkDir"
    Private runControlTableInXML As String = "RunControlTableRu01.XML"
    ' *************** input variables

    ' *************** output variables
    Private isSecondStageFinished As Boolean
    ' *************** output variables

    Sub Main()

        Dim runControlTable As Data.DataTable = New Data.DataTable()
        isSecondStageFinished = True
        runControlTable = GetTableFromFile(localFolder, runControlTableInXML)
        For i As Integer = 0 To runControlTable.Rows.Count - 1
            If runControlTable.Rows(i)("IsComplete") = False Then
                isSecondStageFinished = False
                Exit For
            End If
        Next

        Console.WriteLine("IsSecondStageFinished {0}", isSecondStageFinished)

    End Sub

    Private Function GetTableFromFile(localFolder As String, tableInXML As String) As System.Data.DataTable
        Dim table As System.Data.DataTable = New System.Data.DataTable
        table.ReadXmlSchema(localFolder & "\" & tableInXML)
        table.ReadXml(localFolder & "\" & tableInXML)
        Return table
    End Function

End Module
