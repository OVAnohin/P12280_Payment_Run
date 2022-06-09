Imports System.IO
Imports System.Data
Imports System.Xml.Serialization

Module Module1

    Sub Main()

        Dim table As System.Data.DataTable = New System.Data.DataTable
        'table.Columns.Add(New DataColumn("NameRun", GetType(System.String)))
        'table.Columns.Add(New DataColumn("SheetName", GetType(System.String)))
        'table.Columns.Add(New DataColumn("Identifier", GetType(System.String)))
        'table.Columns.Add(New System.Data.DataColumn("PaymentAccounts", table.GetType()))
        'table.Columns.Add(New System.Data.DataColumn("SheetData", table.GetType()))
        'table.Columns.Add(New DataColumn("IsComplete", GetType(System.Boolean)))
        'table.Columns.Add(New DataColumn("IsRunCreated", GetType(System.Boolean)))

        table.Columns.Add(New DataColumn("F1", GetType(System.String)))
        table.Columns.Add(New DataColumn("F2", GetType(System.String)))
        table.Columns.Add(New DataColumn("F3", GetType(System.String)))
        table.Columns.Add(New DataColumn("F4", GetType(System.Decimal)))
        table.Columns.Add(New DataColumn("F5", GetType(System.String)))
        table.Columns.Add(New DataColumn("F6", GetType(System.String)))
        table.Columns.Add(New DataColumn("F7", GetType(System.Decimal)))
        table.Columns.Add(New DataColumn("F8", GetType(System.Decimal)))
        table.Columns.Add(New DataColumn("F9", GetType(System.Decimal)))
        table.Columns.Add(New DataColumn("F10", GetType(System.String)))
        table.Columns.Add(New DataColumn("F11", GetType(System.Decimal)))
        table.Columns.Add(New DataColumn("F12", GetType(System.DateTime)))
        table.Columns.Add(New DataColumn("F13", GetType(System.DateTime)))
        table.Columns.Add(New DataColumn("F14", GetType(System.DateTime)))
        table.Columns.Add(New DataColumn("F15", GetType(System.String)))
        table.Columns.Add(New DataColumn("F16", GetType(System.DateTime)))
        table.Columns.Add(New DataColumn("F17", GetType(System.DateTime)))
        table.Columns.Add(New DataColumn("F18", GetType(System.String)))
        table.Columns.Add(New DataColumn("F19", GetType(System.Decimal)))
        table.Columns.Add(New DataColumn("F20", GetType(System.String)))
        table.Columns.Add(New DataColumn("F21", GetType(System.String)))
        table.Columns.Add(New DataColumn("F22", GetType(System.Decimal)))
        table.Columns.Add(New DataColumn("F23", GetType(System.String)))
        table.Columns.Add(New DataColumn("F24", GetType(System.String)))
        table.Columns.Add(New DataColumn("F25", GetType(System.String)))
        table.Columns.Add(New DataColumn("F26", GetType(System.String)))
        table.Columns.Add(New DataColumn("F27", GetType(System.Decimal)))
        table.Columns.Add(New DataColumn("F28", GetType(System.Decimal)))
        table.Columns.Add(New DataColumn("F29", GetType(System.String)))
        table.Columns.Add(New DataColumn("F30", GetType(System.String)))
        table.Columns.Add(New DataColumn("F31", GetType(System.String)))
        table.Columns.Add(New DataColumn("F32", GetType(System.String)))
        table.Columns.Add(New DataColumn("F33", GetType(System.String)))
        table.Columns.Add(New DataColumn("F34", GetType(System.String)))
        table.Columns.Add(New DataColumn("F35", GetType(System.String)))
        table.Columns.Add(New DataColumn("F36", GetType(System.String)))
        table.Columns.Add(New DataColumn("F37", GetType(System.String)))
        table.Columns.Add(New DataColumn("F38", GetType(System.String)))
        table.Columns.Add(New DataColumn("F39", GetType(System.String)))
        table.Columns.Add(New DataColumn("F40", GetType(System.String)))
        table.Columns.Add(New DataColumn("F41", GetType(System.String)))
        table.Columns.Add(New DataColumn("F42", GetType(System.String)))
        table.Columns.Add(New DataColumn("F43", GetType(System.String)))
        table.Columns.Add(New DataColumn("F44", GetType(System.Decimal)))
        table.Columns.Add(New DataColumn("F45", GetType(System.String)))
        table.Columns.Add(New DataColumn("F46", GetType(System.String)))
        table.Columns.Add(New DataColumn("F47", GetType(System.String)))
        table.Columns.Add(New DataColumn("F48", GetType(System.String)))
        table.Columns.Add(New DataColumn("F49", GetType(System.String)))
        table.Columns.Add(New DataColumn("F50", GetType(System.String)))
        table.Columns.Add(New DataColumn("F51", GetType(System.String)))
        table.Columns.Add(New DataColumn("F52", GetType(System.String)))
        table.Columns.Add(New DataColumn("F53", GetType(System.String)))
        table.Columns.Add(New DataColumn("F54", GetType(System.String)))
        table.Columns.Add(New DataColumn("F55", GetType(System.Decimal)))
        table.Columns.Add(New DataColumn("F56", GetType(System.String)))
        table.Columns.Add(New DataColumn("F57", GetType(System.String)))
        table.Columns.Add(New DataColumn("F58", GetType(System.Decimal)))
        table.Columns.Add(New DataColumn("F66", GetType(System.String)))
        Dim fileReader As System.IO.StreamReader
        fileReader = My.Computer.FileSystem.OpenTextFileReader("c:\Temp\WorkDir\TableToRun.xml")
        'fileReader = My.Computer.FileSystem.OpenTextFileReader("c:\Temp\WorkDir\test.txt")
        table.ReadXml(fileReader)
        fileReader.Close()

        'Dim stream As FileStream = New FileStream("c:\Temp\WorkDir\RunControlTable.XML", FileMode.Open, FileAccess.Read)
        'Dim deSerializer As XmlSerializer = New XmlSerializer(table.GetType())
        'table = deSerializer.Deserialize(stream)
        'stream.Close()

        Dim tableToRun As DataTable = New Data.DataTable()
        tableToRun = TryCast(table.Rows(0)("SheetData"), DataTable)

    End Sub

    Private Sub DemonstrateReadWriteXMLDocumentWithStream()
        Dim table As DataTable = CreateTestTable("XmlDemo")
        PrintValues(table, "Original table")

        ' Write the schema and data to XML in a memory stream.
        Dim xmlStream As New System.IO.MemoryStream()
        table.WriteXml(xmlStream, XmlWriteMode.WriteSchema)

        ' Rewind the memory stream.
        xmlStream.Position = 0

        Dim newTable As New DataTable
        newTable.ReadXml(xmlStream)

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

End Module
