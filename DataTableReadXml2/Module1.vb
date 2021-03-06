Module Module1

    Sub Main()

        'Dim table As DataTable = CreateTestTable("XmlDemo")
        'PrintValues(table, "Original table")

        '' Write the schema and data to XML in a memory stream.
        'Dim xmlStream As New System.IO.MemoryStream()
        'table.WriteXml(xmlStream, XmlWriteMode.WriteSchema)

        '' Rewind the memory stream.
        'xmlStream.Position = 0

        Dim table As DataTable = CreateTestTable("XmlDemo")
        PrintValues(table, "Original table")

        Dim fileName As String = "c:\Temp\WorkDir\test.xml"
        'table.WriteXml(fileName, XmlWriteMode.WriteSchema)

        Dim newTable As New DataTable
        'newTable.Columns.Add(New DataColumn("id", GetType(System.Int32)))
        'newTable.Columns.Add(New DataColumn("item", GetType(System.String)))
        newTable.ReadXmlSchema(fileName)
        newTable.ReadXml(fileName)

        ' Print out values in the table.
        PrintValues(newTable, "New Table")

    End Sub

    Private Function CreateTestTable(ByVal tableName As String) As DataTable
        ' Create a test DataTable with two columns and a few rows.
        Dim table As New DataTable(tableName)
        Dim column As New DataColumn("id", GetType(System.Int32))
        column.AutoIncrement = True
        table.Columns.Add(column)

        column = New DataColumn("item", GetType(System.String))
        table.Columns.Add(column)

        ' Add ten rows.
        Dim row As DataRow
        For i As Integer = 0 To 9
            row = table.NewRow()
            row("item") = "item " & i
            table.Rows.Add(row)
        Next i

        table.AcceptChanges()
        Return table
    End Function

    Private Sub PrintValues(ByVal table As DataTable, ByVal label As String)
        ' Display the contents of the supplied DataTable:
        Console.WriteLine(label)
        For Each row As DataRow In table.Rows
            For Each column As DataColumn In table.Columns
                Console.Write("{0}{1}", ControlChars.Tab, row(column))
            Next column
            Console.WriteLine()
        Next row
    End Sub

    Private Sub DemonstrateReadWriteXMLDocumentWithString()
        Dim table As DataTable = CreateTestTable("XmlDemo")
        PrintValues(table, "Original table")

        ' Write the schema and data to XML in a file.
        Dim fileName As String = "C:\TestData.xml"
        table.WriteXml(fileName, XmlWriteMode.WriteSchema)

        Dim newTable As New DataTable
        newTable.ReadXml(fileName)

        ' Print out values in the table.
        PrintValues(newTable, "New Table")
    End Sub


End Module
