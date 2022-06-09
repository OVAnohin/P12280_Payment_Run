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
        Dim excelFileName As String = "Общий план_03-12_недели_RU01.xlsb"
        'Dim excelFileName As String = "Украина-План_платежей_03-10 week_18.01.2022-09.03.2022.xlsb"
        Dim sheetNumber As Integer = 1

        '*********************** Result
        Dim dataFromExcelOut As DataTable

        '*********************** Begin
        Dim fullFileName As String = localFolder & "\" & excelFileName
        Dim sheetName As String = GetNameSheet(fullFileName, sheetNumber)

        Dim dataFromExcel As DataTable = New DataTable()
        Dim connetionString As String
        Dim sql As String

        connetionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fullFileName & ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'"
        sql = "Select * from[" & sheetName & "$]"
        Using oledbCnn = New OleDbConnection(connetionString)
            Using oledbCmd = New OleDbCommand(sql, oledbCnn)
                Using oledbAdaper As OleDbDataAdapter = New OleDbDataAdapter(oledbCmd)
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

        If (dataFromExcel IsNot Nothing) Then
            If (dataFromExcel.Rows.Count > 0) Then
                Dim dataFromExcelUTC As DataTable = dataFromExcel.Clone()

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

                dataFromExcelOut = dataFromExcelUTC
            End If
        End If

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

    ' VB .net function ReleaseObject
    Private Sub ReleaseObject(ByVal comOj As Object)
        Try
            If comOj IsNot Nothing AndAlso System.Runtime.InteropServices.Marshal.IsComObject(comOj) Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(comOj)
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(comOj)
            End If
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
