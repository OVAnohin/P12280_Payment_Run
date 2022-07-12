Imports System.Data
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports System.Diagnostics
Imports System.IO
Imports System.Xml.Serialization
Imports System.Data.DataSetExtensions

Module Module1

    Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Integer, ByRef lpdwProcessId As Integer) As Integer

    Sub Main()

		Dim localFolder As String = "C:\Temp\WorkDir"
		Dim excelFileName As String = "Общий план_26-33_недели_RU01 (version 1).xlsb"
		'Dim excelFileName As String = "Украина-План_платежей_03-10 week_18.01.2022-09.03.2022.xlsb"
		Dim paymentDate As Date = Convert.ToDateTime("28.06.2022")
		Dim xmlFileName As String = "Table_PlanRU01.xml"

		'*********************** Result
		'Dim dataFromExcelOut As DataTable

		'*********************** Begin
		'*********************** Begin
		Dim fullFileName As String = localFolder & "\" & excelFileName
		'Dim sheetName As String = GetNameSheet(fullFileName, sheetNumber)
		Dim sheetName As String = "ПЛАН"

		Dim dataFromExcel As DataTable = New DataTable()
		Dim connetionString As String
		Dim sql As String

		connetionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fullFileName & ";" + "Extended Properties='Excel 12.0 Xml;HDR=No;'"
		sql = "Select * from[" & sheetName & "$] Where [F17] like '" & paymentDate & "'"
		Using oledbCnn = New OleDbConnection(connetionString)
			Using oledbCmd = New OleDbCommand(sql, oledbCnn)
				Using oledbAdaper As OleDbDataAdapter = New OleDbDataAdapter(oledbCmd)
					'oledbAdaper.FillSchema(dataFromExcel, SchemaType.Source)
					'dataFromExcel.Columns(15).DataType = Type.GetType("System.DateTime")
					'dataFromExcel.Columns(16).DataType = Type.GetType("System.DateTime")
					oledbAdaper.Fill(dataFromExcel)
				End Using
			End Using
			oledbCnn.Close()
		End Using

		'dataFromExcel.AsEnumerable().Where(Function(row As DataRow) row.ItemArray.All(Function(field) field Is Nothing Or field Is DBNull.Value Or field.Equals(""))).ToList().ForEach(Sub(row) row.Delete())
		For i As Integer = dataFromExcel.Rows.Count - 1 To 0 Step -1
			Dim row As DataRow = dataFromExcel.Rows(i)
			If row.Item(0) Is Nothing Then
				dataFromExcel.Rows.Remove(row)
			ElseIf row.Item(0).ToString = "" Then
				dataFromExcel.Rows.Remove(row)
			End If
		Next
		dataFromExcel.AcceptChanges()

		Dim view As DataView
		Dim filter As String
		Dim tempTable As DataTable
		Dim dataFromExcelUTC As DataTable = New DataTable(sheetName)

		If (dataFromExcel IsNot Nothing) Then
			If (dataFromExcel.Rows.Count > 0) Then

				dataFromExcel.Columns.Remove("F59")
				dataFromExcel.Columns.Remove("F60")
				dataFromExcel.Columns.Remove("F61")
				dataFromExcel.Columns.Remove("F62")
				dataFromExcel.Columns.Remove("F63")
				dataFromExcel.Columns.Remove("F64")
				dataFromExcel.Columns.Remove("F65")
				dataFromExcel.AcceptChanges()

				Dim nameOfQColumn As String
				dataFromExcel.CaseSensitive = False

				'Платить Да/Нет
				nameOfQColumn = dataFromExcel.Columns(0).ColumnName
				view = New DataView(dataFromExcel)
				filter = "[" & nameOfQColumn & "] = 'да' Or [" & nameOfQColumn & "] = 'да-другой план' Or [" & nameOfQColumn & "] = 'да-ТАП'"
				view.RowFilter = filter
				tempTable = view.ToTable()

				Dim dtCloned As DataTable = tempTable.Clone()
				dtCloned.Columns(15).DataType = Type.GetType("System.DateTime")
				dtCloned.Columns(16).DataType = Type.GetType("System.DateTime")

				For i As Integer = 0 To tempTable.Rows.Count - 1
					dtCloned.ImportRow(tempTable.Rows(i))
				Next

				tempTable.Reset()

				Try
					nameOfQColumn = dataFromExcel.Columns(16).ColumnName
					Console.WriteLine(nameOfQColumn)
					Dim planedDate As Date
					For i As Integer = 0 To dtCloned.Rows.Count - 1
						If DBNull.Value.Equals(dtCloned.Rows(i)(nameOfQColumn)) Then
							dtCloned.Rows(i)(nameOfQColumn) = Convert.ToDateTime("01.01.0001")
						End If
					Next

					'столбец Q = дата платежа
					'nameOfQColumn = dataFromExcel.Columns(16).ColumnName
					dataFromExcel = (From row In dtCloned.AsEnumerable()
									 Where Date.TryParse(row.Field(Of Date)(nameOfQColumn), planedDate) AndAlso planedDate = paymentDate).CopyToDataTable()

					dtCloned.Reset()

				Catch ex As Exception
					dataFromExcel = tempTable.Clone()
				End Try

				view = New DataView(dataFromExcel)
				'столбец D сортируем.
				nameOfQColumn = dataFromExcel.Columns(3).ColumnName
				view.Sort = "[" & nameOfQColumn & "]"
				dataFromExcel = view.ToTable()

				dataFromExcelUTC = dataFromExcel.Clone()
				dataFromExcelUTC.TableName = sheetName

				For i As Integer = 0 To dataFromExcelUTC.Columns.Count - 1
					If Type.GetType(dataFromExcelUTC.Columns(i).DataType.ToString()).ToString() = "System.DateTime" Then
						dataFromExcelUTC.Columns(i).DateTimeMode = DataSetDateTime.Utc
					End If
				Next
				dataFromExcelUTC.AcceptChanges()

				For i As Integer = 0 To dataFromExcel.Rows.Count - 1
					Dim row As DataRow = dataFromExcel.Rows(i)
					dataFromExcelUTC.ImportRow(row)
				Next
			End If
		End If

		SaveDataTableToFile(localFolder & "\" & xmlFileName, dataFromExcelUTC)

		Console.WriteLine("End")

	End Sub

    Private Function GetNameSheet(ByVal fullFileName As String, ByVal sheetNumber As Integer) As String
        Dim oMissing As Object = System.Reflection.Missing.Value
        Dim excelApp As Excel.Application = New Excel.Application()
        Dim excelAppProcess As Process = GetExcelProcess(excelApp)
        excelApp.DisplayAlerts = False
        excelApp.FileValidationPivot = Excel.XlFileValidationPivotMode.xlFileValidationPivotRun
        Dim excelWb As Excel.Workbook = excelApp.Workbooks.Open(fullFileName)
        Dim excelWs As Excel.Worksheet = TryCast(excelWb.Worksheets(sheetNumber), Excel.Worksheet)

        Dim sheetName As String = excelWs.Name
        excelWb.Close(oMissing, oMissing, oMissing)
        excelApp.Quit()
        excelApp = Nothing
        excelAppProcess.Kill()

        ReleaseObject(excelApp)
        ReleaseObject(excelWb)
        ReleaseObject(excelWs)
        Return sheetName
    End Function

    Private Function GetExcelProcess(ByVal excelApp As Excel.Application) As Process
        Dim id As Integer
        GetWindowThreadProcessId(excelApp.Hwnd, id)
        Return Process.GetProcessById(id)
    End Function

    Private Sub SaveDataTableToFile(ByVal fileName As String, ByVal table As DataTable)
        Dim Stream As FileStream = New FileStream(fileName, FileMode.Create)
        Dim serializer As XmlSerializer = New XmlSerializer(table.GetType())
        serializer.Serialize(Stream, table)
        Stream.Close()
    End Sub

    Private Sub ReleaseObject(ByVal comOj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(comOj)
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(comOj)
            comOj = Nothing
        Catch ex As Exception
            comOj = Nothing
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
        End Try
    End Sub

End Module
