Imports System.IO
Imports System.Xml.Serialization

Module Module1

    ' *************** input variables
    Dim localFolder As String = "C:\Temp\WorkDir"
    Dim sheetName As String = "Citibank - RUB debtors I 0" 'это у нас имя листа
    Dim runControlTableInXML As String = "RunControlTableRu01.XML"
    ' *************** input variables

    Sub Main()

        Dim runControlTable As Data.DataTable = New Data.DataTable()
        ' обновить данные в глобальной таблице
        runControlTable = GetTableFromFile(localFolder, runControlTableInXML)
        For i As Integer = 0 To runControlTable.Rows.Count - 1
            If runControlTable.Rows(i)("SheetName") = sheetName Then
                runControlTable.Rows(i)("IsJournalCreated") = True
                Exit For
            End If
        Next
        runControlTable.AcceptChanges()
        SaveDataTableToFile(localFolder & "\" & runControlTableInXML, runControlTable)
        runControlTable = New Data.DataTable()

    End Sub

    Private Sub SaveDataTableToFile(fileName As String, table As System.Data.DataTable)
        Dim Stream As FileStream = New FileStream(fileName, FileMode.Create)
        Dim serializer As XmlSerializer = New XmlSerializer(table.GetType())
        serializer.Serialize(Stream, table)
        Stream.Close()
    End Sub

    Private Function GetTableFromFile(localFolder As String, tableInXML As String) As System.Data.DataTable
        Dim table As System.Data.DataTable = New System.Data.DataTable
        table.ReadXmlSchema(localFolder & "\" & tableInXML)
        table.ReadXml(localFolder & "\" & tableInXML)
        Return table
    End Function

End Module
