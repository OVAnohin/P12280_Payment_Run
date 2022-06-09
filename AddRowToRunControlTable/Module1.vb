Imports System.IO
Imports System.Threading
Imports System.Xml.Serialization

Module Module1

    Dim localFolder As String = "C:\Temp\WorkDir"
    Dim sheetName As String = "DB - USD, EUR 0" 'это у нас имя листа
    Dim nameRun As String = "DB - USD, EUR"
    Dim identifier As String = ""
    Dim runControlTableInXML As String = "RunControlTable.XML"
    Dim tableToRunInXML As String = "TableToRun.XML"
    Dim be As String = "Ru17"

    Sub Main()
        Console.WriteLine("Первичный поток: Id {0}", Thread.CurrentThread.ManagedThreadId)

        '*********************** Begin
        be = be.ToUpper()
        Dim runControlTable As System.Data.DataTable = GetTableFromFile(localFolder, runControlTableInXML)
        Dim resultTable As System.Data.DataTable = runControlTable
        Dim tableToRun As System.Data.DataTable = New System.Data.DataTable()
        tableToRun = GetTableFromFile(localFolder, tableToRunInXML)

        Dim view As System.Data.DataView
        view = New System.Data.DataView(tableToRun)
        Dim paymentsAccountsTable As System.Data.DataTable = view.ToTable(True, "F4")
        paymentsAccountsTable = RemoveNullValue(paymentsAccountsTable, "F4")

        If (resultTable IsNot Nothing) Then
            Dim row As System.Data.DataRow = resultTable.NewRow()
            If be = "RU17" Then
                If nameRun.Contains("DB -") Then
                    row("NameRun") = nameRun.Replace("DB -", "Citibank -")
                Else
                    row("NameRun") = nameRun
                End If
            Else
                row("NameRun") = nameRun
            End If
            row("SheetName") = sheetName
            row("Identifier") = "Zero"
            row("PaymentAccounts") = paymentsAccountsTable
            row("SheetData") = tableToRun
            row("IsComplete") = False
            row("IsRunCreated") = False
            resultTable.Rows.Add(row)
            resultTable.AcceptChanges()
        End If

        SaveDataTableToFile(localFolder & "\" & runControlTableInXML, resultTable)

    End Sub

    Private Function GetTableFromFile(localFolder As String, tableInXML As String) As System.Data.DataTable
        Dim table As System.Data.DataTable = New System.Data.DataTable
        table.ReadXmlSchema(localFolder & "\" & tableInXML)
        table.ReadXml(localFolder & "\" & tableInXML)
        Return table
    End Function

    Private Function RemoveNullValue(table As System.Data.DataTable, columnName As String) As System.Data.DataTable
        For i As Integer = table.Rows.Count - 1 To 0 Step -1
            If DBNull.Value.Equals(table.Rows(i)(columnName)) Then
                table.Rows.Remove(table.Rows(i))
            End If
        Next

        Return table
    End Function

    Private Sub SaveDataTableToFile(ByVal fileName As String, ByVal table As System.Data.DataTable)
        Dim Stream As FileStream = New FileStream(fileName, FileMode.Create)
        Dim serializer As XmlSerializer = New XmlSerializer(table.GetType())
        serializer.Serialize(Stream, table)
        Stream.Close()
    End Sub


End Module
