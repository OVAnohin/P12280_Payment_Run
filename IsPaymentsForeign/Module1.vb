Imports System.Threading

Module Module1

    ' *************** input variables
    Private localFolder As String = "C:\Temp\WorkDir"
    Private identifier As String = ""
    Private sheetName As String = "DB - USD, EUR 1" 'это у нас имя листа
    ' *************** input variables

    ' *************** output variables
    Private isPaymentsForeign As Boolean
    Private ownBank As String
    ' *************** output variables

    Sub Main()

        Console.WriteLine("Первичный поток: Id {0}", Thread.CurrentThread.ManagedThreadId)
        '*********************** Begin
        ' текущая таблица для запуска или текущий лист
        'Dim tablesToRunFileNameXml As String = "TablesToRun.xml"
        'Dim tableCurrentRun As System.Data.DataTable = New Data.DataTable()
        'Dim currentPaymentAccountsTable As System.Data.DataTable = New Data.DataTable()
        ''получаем данные текущего прогона из таблицы "TablesToRun.xml", она обновляется тут в коде
        'Dim tablesToRun As Data.DataTable = GetTableFromFile(localFolder, tablesToRunFileNameXml)
        'For i As Integer = tablesToRun.Rows.Count - 1 To 0 Step -1
        '    If tablesToRun.Rows(i)("SheetName") = sheetName Then
        '        tableCurrentRun = tablesToRun.Rows(i)("SheetData")
        '        Exit For
        '    End If
        'Next

        'isPaymentsForeign = False
        'Dim view As DataView
        'Dim tempTable As System.Data.DataTable
        'view = New DataView(tableCurrentRun)
        'tempTable = view.ToTable(True, "F10") 'это таблица валют
        'For Each row As DataRow In tempTable.Rows
        '    If row("F10") = "EUR" OrElse row("F10") = "USD" OrElse row("F10") = "JBP" Then
        '        isPaymentsForeign = True
        '        Exit For
        '    End If
        'Next

        If identifier.Contains("RUL") Then
            isPaymentsForeign = False
        Else
            isPaymentsForeign = True
        End If
        ownBank = GetOwnBank(sheetName)

        Console.WriteLine("Первичный поток: Id {0} Is Ended", Thread.CurrentThread.ManagedThreadId)
        Console.WriteLine("isPaymentsForeign = {0}", isPaymentsForeign)
        Console.WriteLine("ownBank = {0}", ownBank)
        Console.ReadKey()

    End Sub

    Private Function GetTableFromFile(localFolder As String, tableInXML As String) As System.Data.DataTable
        Dim table As System.Data.DataTable = New System.Data.DataTable
        table.ReadXmlSchema(localFolder & "\" & tableInXML)
        table.ReadXml(localFolder & "\" & tableInXML)
        Return table
    End Function

    Private Function GetOwnBank(ByVal sheetName As String) As String
        If Left(sheetName, 2) = "DB" Then
            Return "DBSBK"
        End If

        Return "CITBK"
    End Function

End Module
