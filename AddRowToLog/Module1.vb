Imports System.IO
Imports System.Xml.Serialization

Module Module1

    Private localFolder As String = "C:\Temp\WorkDir"
    Private remoteFolder As String = "\\rus.efesmoscow\DFS\MOSC\Projects.MOSC\Robotic\P12280 Payment Run\WorkDir\Log"
    Private dateAndTime As DateTime = Convert.ToDateTime("18.06.2022")
    Private processName As String
    Private isRu01 As Boolean
    Private isRu14 As Boolean
    Private isRu17 As Boolean
    Private filePlanRu01 As String
    Private filePlanRu14 As String
    Private filePlanRu17 As String
    Private paymentDateRu01 As Date
    Private paymentDateRu14 As Date
    Private paymentDateRu17 As Date
    Private isCompleteRu01 As Boolean
    Private isCompleteRu14 As Boolean
    Private isCompleteRu17 As Boolean

    Sub Main()

        Dim logTableFileName As String = "Log.xml"
        DeleteFile(localFolder, logTableFileName)
        CopyFile(remoteFolder, logTableFileName, localFolder, logTableFileName)

        Dim logTable As System.Data.DataTable = GetTableFromFile(localFolder, logTableFileName)
        Dim row As DataRow = logTable.NewRow()
        row("DateAndTime") = dateAndTime
        row("ProcessName") = processName
        row("Ru01") = isRu01
        row("Ru14") = isRu14
        row("Ru17") = isRu17
        row("FilePlanRu01") = filePlanRu01
        row("FilePlanRu14") = filePlanRu14
        row("FilePlanRu17") = filePlanRu17
        row("PaymentDateRu01") = paymentDateRu01
        row("PaymentDateRu14") = paymentDateRu14
        row("PaymentDateRu17") = paymentDateRu17
        row("IsCompleteRu01") = isCompleteRu01
        row("IsCompleteRu14") = isCompleteRu14
        row("IsCompleteRu17") = isCompleteRu17

        logTable.Rows.Add(row)
        logTable.AcceptChanges()

        SaveDataTableToFile(localFolder & "\" & logTableFileName, logTable)
        DeleteFile(remoteFolder, logTableFileName)
        CopyFile(localFolder, logTableFileName, remoteFolder, logTableFileName)

    End Sub

    Private Function GetTableFromFile(localFolder As String, tableInXML As String) As System.Data.DataTable
        Dim table As System.Data.DataTable = New System.Data.DataTable
        table.ReadXmlSchema(localFolder & "\" & tableInXML)
        table.ReadXml(localFolder & "\" & tableInXML)
        Return table
    End Function

    Private Sub DeleteFile(localFolder As String, fileToRemove As String)
        Try
            File.Delete(localFolder & "\" & fileToRemove)
        Catch ex As Exception
            Throw New Exception("Не могу удалить файл " & fileToRemove)
        End Try
    End Sub

    Private Sub CopyFile(localFolder As String, fileToCopy As String, destinationFolder As String, newCopy As String)
        Try
            File.Copy(localFolder & "\" & fileToCopy, destinationFolder & "\" & newCopy, True)
        Catch ex As Exception
            Throw New Exception("Не могу скопировать файл " & fileToCopy & " в " & newCopy & " " & ex.Message)
        End Try
    End Sub

    Private Sub SaveDataTableToFile(ByVal fileName As String, ByVal table As System.Data.DataTable)
        Dim Stream As FileStream = New FileStream(fileName, FileMode.Create)
        Dim serializer As XmlSerializer = New XmlSerializer(table.GetType())
        serializer.Serialize(Stream, table)
        Stream.Close()
    End Sub

End Module
