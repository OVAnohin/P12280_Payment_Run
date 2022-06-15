Imports System.Threading
Imports System.Data

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

    ' *************** Thread
    Private _thread As Thread
    ' *************** Thread

    ' *************** input variables
    Dim localFolder As String = "C:\Temp\WorkDir"
    Dim paymentDate As Date = Convert.ToDateTime("14.06.2022")
    Dim sheetName As String = "DB - USD, EUR 0" 'это у нас имя листа
    Dim be As String = "RU01" ' текущая BE
    Dim runControlTableInXML As String = "RunControlTableRu01.XML"
    ' *************** input variables

    ' *************** output variables
    Dim identifier As String
    Dim ownBank As String
    Dim paymentMethod As String
    Dim isRunCreated As Boolean
    Dim isComplete As Boolean
    ' *************** output variables

    Sub Main()
        Console.WriteLine("Первичный поток: Id {0}", Thread.CurrentThread.ManagedThreadId)
        '*********************** Begin
        Dim tablesToRunFileNameXml As String = "TablesToRun.xml"
        Dim isExit As Boolean = False
        Dim runControlTable As Data.DataTable = New Data.DataTable()
        Dim columns As DataColumnCollection = runControlTable.Columns
        Dim isPaymentsForeign As Boolean


        ' текущая таблица для запуска или текущий лист
        Dim tableCurrentRun As System.Data.DataTable = New Data.DataTable()
        Dim currentPaymentAccountsTable As System.Data.DataTable = New Data.DataTable()
        ' startSearchRow это номер строки предыдущего прогона (-1 значит это нету предыдущих прогонов)
        Dim startSearchRow As Integer = 0
        identifier = "Zero" ' Для первой таблицы
        Dim isCurrentTableComplete As Boolean
        'получаем данные текущего прогона из таблицы "TablesToRun.xml", она обновляется тут в коде
        Dim tablesToRun As Data.DataTable = GetTableFromFile(localFolder, tablesToRunFileNameXml)
        For i As Integer = tablesToRun.Rows.Count - 1 To 0 Step -1
            If tablesToRun.Rows(i)("SheetName") = sheetName Then
                tableCurrentRun = tablesToRun.Rows(i)("SheetData")
                identifier = tablesToRun.Rows(i)("Identifier")
                currentPaymentAccountsTable = tablesToRun.Rows(i)("PaymentAccounts")
                isCurrentTableComplete = tablesToRun.Rows(i)("IsComplete")
                isRunCreated = tablesToRun.Rows(i)("IsRunCreated")
                startSearchRow = i - 1
                Exit For
            End If
        Next

        Dim view As DataView
        Dim tempTable As System.Data.DataTable

        isPaymentsForeign = False
        view = New DataView(tableCurrentRun)
        tempTable = view.ToTable(True, "F10") 'это таблица валют
        For Each row As DataRow In tempTable.Rows
            If row("F10") = "RUE" OrElse row("F10") = "RUD" OrElse row("F10") = "EUR" OrElse row("F10") = "USD" Then
                isPaymentsForeign = True
                Exit For
            End If
        Next

        'Currency
        Dim currency As String = GetCurrency(tableCurrentRun)

        paymentMethod = GetPaymentMethod(tableCurrentRun)

        ' Идентификатор (номер прогона по порядку в течение дня)
        ' Нужно проверить, были ли выполненые предыдущие листы с текущими 
        ownBank = GetOwnBank(sheetName)

        'если isCurrentTableComplete, это может быть второй или третий запуск
        'ничего не надо, выходим
        If isCurrentTableComplete AndAlso isRunCreated Then
            isComplete = True
            Exit Sub
        End If

        ' По идее если есть подобный счет в предыдущем запуске, то мы запускать новый не можем
        ' поэтому не нужно проверять больше одного листа.
        Dim oldRunIdentifier As String = "Zero"
        Dim isCompleteOldRun As Boolean = True ' считаем что предыдущий прогон исполнен
        Dim oldSheet As String = ""

        Dim isAccountsPresentInOldIdentifier As Boolean = False 'старые счета есть в прогоне?
        If startSearchRow >= 0 Then
            oldRunIdentifier = tablesToRun.Rows(startSearchRow)("Identifier")
            Dim oldPaymentAccountsTable As System.Data.DataTable = tablesToRun.Rows(startSearchRow)("PaymentAccounts")
            isCompleteOldRun = tablesToRun.Rows(startSearchRow)("IsComplete")
            oldSheet = tablesToRun.Rows(startSearchRow)("SheetName")
            'Проверка на совпадение по таблицам счетов и законченности прошлого прогона
            'если счета есть, то нужно проверить прошлый прогон на то, что он закончен.
            For i As Integer = 0 To currentPaymentAccountsTable.Rows.Count - 1
                Dim row As Data.DataRow = currentPaymentAccountsTable.Rows(i)
                Dim _filter As String = "[F4] = '" & row("F4") & "'"
                view = New DataView(oldPaymentAccountsTable)
                view.RowFilter = _filter
                tempTable = view.ToTable()
                If tempTable.Rows.Count > 0 Then
                    isAccountsPresentInOldIdentifier = True
                    Exit For
                End If
            Next
        End If

        Console.WriteLine(New String("*", 20))
        Console.WriteLine("isCompleteOldRun = {0}", isCompleteOldRun)
        Console.WriteLine("isRunCreated = {0}", isRunCreated)
        Console.WriteLine("identifier = {0}", identifier)
        Console.WriteLine("oldRunIdentifier = {0}", oldRunIdentifier)
        Console.WriteLine("isCurrentTableComplete = {0}", isCurrentTableComplete)

        If isCompleteOldRun AndAlso isRunCreated = False AndAlso identifier = "Zero" AndAlso oldRunIdentifier = "Zero" AndAlso isCurrentTableComplete = False Then
            ' Первый лист
            identifier = GetIdentifier(ownBank, sheetName, currency, paymentMethod)
        ElseIf isCompleteOldRun AndAlso isRunCreated AndAlso identifier <> "Zero" AndAlso oldRunIdentifier = "Zero" AndAlso isCurrentTableComplete = False Then
            ' Первый лист, но прогон не завершен
            ' тут для проверки
            Console.WriteLine("identifier менять не надо, это лист проверяем на прогон")
        ElseIf isCompleteOldRun = False AndAlso isRunCreated = False AndAlso identifier = "Zero" AndAlso oldRunIdentifier <> "Zero" AndAlso isCurrentTableComplete = False Then
            identifier = GetIdentifier(ownBank, sheetName, currency, paymentMethod)
            If Left(identifier, 3) = Left(oldRunIdentifier, 3) Then
                identifier = ChangeIdentifier(oldRunIdentifier)
            End If
        ElseIf isCompleteOldRun = False AndAlso isRunCreated = False AndAlso identifier = "Zero" AndAlso oldRunIdentifier = "Zero" AndAlso isCurrentTableComplete = False Then
            ' это ноый счет в общем прогоне (пример после 58 идет 363392)
            identifier = GetIdentifier(ownBank, sheetName, currency, paymentMethod)
            If Left(identifier, 3) = Left(oldRunIdentifier, 3) Then
                identifier = ChangeIdentifier(oldRunIdentifier)
            End If
        ElseIf isCompleteOldRun AndAlso isRunCreated = False AndAlso identifier = "Zero" AndAlso oldRunIdentifier <> "Zero" AndAlso isCurrentTableComplete = False Then
            'это когда предыдущий 58 счет (например) прогон выполнен, и у нас ещё есть лист 58 счетов
            identifier = GetIdentifier(ownBank, sheetName, currency, paymentMethod)
            If Left(identifier, 3) = Left(oldRunIdentifier, 3) Then
                identifier = ChangeIdentifier(oldRunIdentifier)
            End If
        End If

        isComplete = True

        Console.WriteLine(New String("*", 20))
        Console.WriteLine("Choosen Identifier = {0}", identifier)

        Console.WriteLine("Первичный поток: Id {0} Is Ended", Thread.CurrentThread.ManagedThreadId)
        Console.WriteLine("isRunCreated = {0}", isRunCreated)
        Console.WriteLine("isComplete = {0}", isComplete)
        Console.ReadKey()

    End Sub

    Private Function ChangeIdentifier(ByVal identifier As String) As String
        Dim str As String = Mid(identifier, 4)
        Dim currentNumber As Integer
        If Int32.TryParse(str, currentNumber) Then
            If currentNumber < 9 Then
                Return Left(identifier, 3) & "0" & (currentNumber + 1)
            End If
        Else
            Throw New Exception("Не могу преобразовать в число Identifier")
        End If

        Return Left(identifier, 3) & (currentNumber + 1)
    End Function

    Private Function GetPaymentMethod(ByVal tableCurrentRun As Data.DataTable) As String
        Dim view As DataView
        Dim tempTable As System.Data.DataTable
        Dim paymentMethod As String = ""
        ' [F21]
        view = New DataView(tableCurrentRun)
        tempTable = view.ToTable(True, "F21")
        tempTable = RemoveNullValue(tempTable, "F21")
        For i As Integer = 0 To tempTable.Rows.Count - 1
            If paymentMethod <> "" Then
                paymentMethod = paymentMethod & ", " & tempTable.Rows(i)("F21")
            Else
                paymentMethod = tempTable.Rows(i)("F21")
            End If
        Next

        Return paymentMethod
    End Function

    Private Function GetCurrency(ByVal tableCurrentRun As Data.DataTable) As String
        Dim view As DataView
        Dim tempTable As System.Data.DataTable
        Dim currency As String = ""
        ' [F10]
        view = New DataView(tableCurrentRun)
        tempTable = view.ToTable(True, "F10")
        tempTable = RemoveNullValue(tempTable, "F10")
        For i As Integer = 0 To tempTable.Rows.Count - 1
            If currency <> "" Then
                currency = currency & "and" & tempTable.Rows(i)("F10")
            Else
                currency = tempTable.Rows(i)("F10")
            End If
        Next

        Return currency
    End Function

    Private Function RemoveNullValue(ByVal table As Data.DataTable, ByVal columnName As String) As Data.DataTable
        For i As Integer = table.Rows.Count - 1 To 0 Step -1
            If DBNull.Value.Equals(table.Rows(i)(columnName)) Then
                table.Rows.Remove(table.Rows(i))
            End If
        Next

        Return table
    End Function

    Private Function GetIdentifier(ownBank As String, sheetName As String, currency As String, paymentMethod As String) As String
        Dim identifier As String = ""
        ' Идентификатор (номер прогона по порядку в течение дня)
        If sheetName.Contains("debtors") Then
            ' Для дебитора: 
            ' RUY (Порядковый номер прогона за текущий день в валюте EUR, USD)
            If currency.Contains("EUR") OrElse currency.Contains("USD") Then
                Return "RUY01"
            End If
            ' RUI (Порядковый номер прогона за текущий день в валюте RUB)
            If currency.Contains("RUB") Then
                Return "RUI01"
            End If
            ' RUL (Порядковый номер прогона за текущий день в валюте RUB -нерезиденты)
            If sheetName.Contains("non-residents") Then
                Return "RUL01"
            End If
        Else
            ' RUL (Порядковый номер прогона за текущий день в валюте RUB -нерезиденты)
            ' по другому быть не может
            If sheetName.Contains("non-residents") Then
                Return "RUL01"
            End If

            ' Для кредитора / дебитора: 
            ' RUL(Порядковый номер прогона за текущий день в валютах RUB, RUE или RUD)-Сити банк, Дойче банк
            ' Может ли быть способ отплаты P?
            If currency.Contains("RUB") OrElse currency.Contains("RUE") OrElse currency.Contains("RUD") OrElse currency.Contains("RUDandRUE") OrElse currency.Contains("RUEandRUD") Then
                Return "RUL01"
            End If

            ' Для кредитора: 
            ' RUP(Порядковый номер прогона за текущий день в валюте (столбец J) EUR, USD для проводок со способом платежа P);
            If (currency.Contains("EUR") OrElse currency.Contains("USD") OrElse currency.Contains("EURandUSD") OrElse currency.Contains("USDandEUR")) AndAlso paymentMethod = "P" Then
                Return "RUP01"
            End If
            ' RUO(Порядковый номер прогона за текущий день в валюте(столбец J) GBP для проводок со способом платежа O)
            If currency.Contains("GBP") AndAlso paymentMethod = "O" Then
                Return "RUO01"
            End If
        End If

        Throw New Exception("Не смог выбрать Identifier.")
    End Function

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

    Private Function RemoveUnnecessaryChar(str As String) As String
        'Chr(8)  Backspace character
        'Chr(32) Space
        'Chr(34) Quotation Mark
        'Chr(160)    Non-breaking space
        RemoveUnnecessaryChar = Replace(Replace(Replace(Replace(Replace(Replace(Replace(str, Chr(13), ""), Chr(7), ""), Chr(9), ""), Chr(11), ""), Chr(160), ""), Chr(32), ""), Chr(46), "")

    End Function

End Module
