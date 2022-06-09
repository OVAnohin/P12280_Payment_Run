Imports System.IO
Imports System.Xml.Serialization
Imports System.Security.Cryptography
Imports System.Text
Imports System.Threading
Imports System.Data
Imports System.Threading.Tasks
Imports System.Runtime.InteropServices
Imports System.Windows.Automation
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System.Diagnostics
Imports System.Data.OleDb
Imports System.Data.DataSetExtensions

Module Module1

    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function SendMessageHM Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Int32, ByVal wMsg As Int32, ByVal wParam As Int32, ByVal lParam As String) As Int32
    Private Declare Function SendMessageW Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As System.IntPtr, ByVal lpEnumFunc As EnumWindowProc, ByVal lParam As Integer) As Boolean
    Private Delegate Function EnumWindowProc(ByVal hWnd As IntPtr, ByVal lParam As IntPtr) As Boolean
    Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Boolean
    Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Integer, ByVal nCmdShow As Integer) As Integer
    Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Integer, ByRef lpdwProcessId As Integer) As Integer

    Private Const WM_COMMAND = &H111
    Private Const BM_CLICK As Integer = &HF5

    Sub Main()

        Dim localFolder As String = "C:\Temp\WorkDir"
        Dim sheetName As String = "DB - USD, EUR 0" 'это у нас имя листа
        Dim xlsbFileName As String = sheetName & ".xlsb"

        Dim dataFromExcel As System.Data.DataTable = GetDatatableFromExcel(localFolder, xlsbFileName, 1)

    End Sub

    Private Function GetDatatableFromExcel(ByVal localFolder As String, ByVal excelFileName As String, ByVal sheetNumber As Integer) As System.Data.DataTable

        Dim fullFileName As String = localFolder & "\" & excelFileName
        Dim sheetName As String = GetNameSheet(fullFileName, sheetNumber)

        Dim dataFromExcel As System.Data.DataTable = New System.Data.DataTable()
        Dim dataFromExcelOut As System.Data.DataTable = New System.Data.DataTable()
        Dim connetionString As String
        Dim sql As String


        connetionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fullFileName & ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'"
        sql = "Select * from[" & sheetName & "$]"
        Using oledbCnn = New System.Data.OleDb.OleDbConnection(connetionString)
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
                Dim dataFromExcelUTC As System.Data.DataTable = dataFromExcel.Clone()

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

        Return dataFromExcelOut
    End Function

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
