Module Module1

    Sub Main()
        Dim ds As New DataSet
        Dim customerTable As DataTable = GetCustomers()
        Dim orderTable As DataTable = GetOrders()

        ds.Tables.Add(customerTable)
        ds.Tables.Add(orderTable)
        ds.Relations.Add("CustomerOrder", New DataColumn() {customerTable.Columns(0)}, New DataColumn() {orderTable.Columns(1)}, True)

        Dim writer As New System.IO.StringWriter
        customerTable.WriteXml(writer, XmlWriteMode.WriteSchema, False)
        PrintOutput(writer, "Customer table, without hierarchy")

        Dim file As System.IO.StreamWriter
        file = My.Computer.FileSystem.OpenTextFileWriter("c:\Temp\WorkDir\test.txt", True)
        customerTable.WriteXml(file, XmlWriteMode.WriteSchema, False)
        file.Close()

        Console.WriteLine("Press any key to continue.")
        Console.ReadKey()
    End Sub

    Private Function GetOrders() As DataTable
        ' Create sample Customers table, in order
        ' to demonstrate the behavior of the DataTableReader.
        Dim table As New DataTable

        ' Create three columns, OrderID, CustomerID, and OrderDate.
        table.Columns.Add(New DataColumn("OrderID", GetType(System.Int32)))
        table.Columns.Add(New DataColumn("CustomerID", GetType(System.Int32)))
        table.Columns.Add(New DataColumn("OrderDate", GetType(System.DateTime)))

        ' Set the OrderID column as the primary key column.
        table.PrimaryKey = New DataColumn() {table.Columns(0)}

        table.Rows.Add(New Object() {1, 1, #12/2/2003#})
        table.Rows.Add(New Object() {2, 1, #1/3/2004#})
        table.Rows.Add(New Object() {3, 2, #11/13/2004#})
        table.Rows.Add(New Object() {4, 3, #5/16/2004#})
        table.Rows.Add(New Object() {5, 3, #5/22/2004#})
        table.Rows.Add(New Object() {6, 4, #6/15/2004#})
        table.AcceptChanges()
        Return table
    End Function

    Private Function GetCustomers() As DataTable
        ' Create sample Customers table, in order
        ' to demonstrate the behavior of the DataTableReader.
        Dim table As New DataTable

        ' Create two columns, ID and Name.
        Dim idColumn As DataColumn = table.Columns.Add("ID",
      GetType(System.Int32))
        table.Columns.Add("Name", GetType(System.String))

        ' Set the ID column as the primary key column.
        table.PrimaryKey = New DataColumn() {idColumn}

        table.Rows.Add(New Object() {1, "Mary"})
        table.Rows.Add(New Object() {2, "Andy"})
        table.Rows.Add(New Object() {3, "Peter"})
        table.Rows.Add(New Object() {4, "Russ"})
        table.AcceptChanges()
        Return table
    End Function

    Private Sub PrintOutput(
   ByVal writer As System.IO.TextWriter, ByVal caption As String)

        Console.WriteLine("==============================")
        Console.WriteLine(caption)
        Console.WriteLine("==============================")
        Console.WriteLine(writer.ToString())
    End Sub

End Module
