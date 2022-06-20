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
Imports System.Drawing

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
    Dim paymentDate As Date = Convert.ToDateTime("22.06.2022")
    Dim sheetName As String = "DB - USD, EUR 1" 'это у нас имя листа
    Dim be As String = "RU01" ' текущая BE
    Dim pathToMailAttachments As String = "\\rus.efesmoscow\DFS\MOSC\Projects.MOSC\Robotic\P12280 Payment Run\WorkDir\MailAttachments"
    Dim runControlTableInXML As String = "RunControlTableRu01.XML"
    Dim xmlNameNotIncludedTable As String = "NotIncludedTableRu01.xml"
    Dim xmlNameErrors_F110 As String = "Errors_F110Ru01.xml"
    Dim xmlNameJournals As String = "JournalsRu01.xml"
    ' *************** input variables

    ' *************** in variables
    Dim _oLock As Object = New Object()
    'Dim _table As System.Data.DataTable = New System.Data.DataTable()
    ' *************** in variables

    ' *************** output variables
    Dim exceptionMessage As String
    Dim isComplete As Boolean
    Dim isPaymentsForeign As Boolean
    Dim identifier As String
    Dim ownBank As String
    Dim paymentMethod As String
    Dim numberOfPayments As Integer
    Dim totalRunAmount As Double
    Dim fileName As String
    Dim screenShotFileName As String
    Dim macrosForRunFileName As String
    Dim currentCurrency As String
    Dim isRunCreated As Boolean
    ' *************** output variables

    Sub Main()
        Console.WriteLine("Первичный поток: Id {0}", Thread.CurrentThread.ManagedThreadId)
        '*********************** Begin
        Dim errorsLogFileNameXml As String = "errorsLog.xml"
        Dim errorsLogFileNameXlsb As String = "errorsLog.xlsb"
        Dim errorsLogFileNameTxt As String = "errorsLog.txt"
        Dim tablesToRunFileNameXml As String = "TablesToRun.xml"
        Dim xlsbNameNotIncludedTable As String = Replace(xmlNameNotIncludedTable, ".xml", ".xlsb")
        Dim isExit As Boolean = False
        Dim runControlTable As Data.DataTable = New Data.DataTable()
        Dim columns As DataColumnCollection = runControlTable.Columns

        totalRunAmount = 0.0
        numberOfPayments = 0

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

        'таблица листа ошибок errorLogTable это Journals
        Dim tableJournals As Data.DataTable = New Data.DataTable()
        Try
            tableJournals = GetTableFromFile(localFolder, xmlNameJournals)
        Catch ex As Exception
        End Try

        Dim view As DataView
        Dim tempTable As System.Data.DataTable
        view = New DataView(tableCurrentRun)
        tempTable = view.ToTable(True, "F4") 'это таблица LenderAaccount distinct
        Dim lenderAccountFile As String = sheetName & "LA.txt"
        CreateTXTFileFromCollection(tempTable, "F4", localFolder, lenderAccountFile)

        isPaymentsForeign = False
        view = New DataView(tableCurrentRun)
        tempTable = view.ToTable(True, "F10") 'это таблица валют
        For Each row As DataRow In tempTable.Rows
            If row("F10") = "RUE" OrElse row("F10") = "RUD" OrElse row("F10") = "EUR" OrElse row("F10") = "USD" Then
                isPaymentsForeign = True
                Exit For
            End If
        Next

        ' Если таблица EPAP нужна самая старшая дата платежа [F16]
        If sheetName.Contains("EPAP") Then
            view = New DataView(tableCurrentRun)
            view.Sort = "[F16] DESC"
            tableCurrentRun = view.ToTable()
            paymentDate = tableCurrentRun.Rows(0)("F16")
        End If

        'Currency
        Dim currency As String = GetCurrency(tableCurrentRun)

        paymentMethod = GetPaymentMethod(tableCurrentRun)

        ' Идентификатор (номер прогона по порядку в течение дня)
        ' Нужно проверить, были ли выполненые предыдущие листы с текущими 
        ownBank = GetOwnBank(sheetName)

        'если isCurrentTableComplete, это может быть второй или третий запуск
        'ничего не надо, выходим
        If isCurrentTableComplete AndAlso isRunCreated Then
            ' ставим isRunCreated в False, что бы не отправлять письмо.
            isRunCreated = False
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

        Console.WriteLine(New String("*", 20))
        Console.WriteLine("Choosen Identifier = {0}", identifier)

        'Dim connection = GetObject("SAPGUI").GetScriptingEngine.Children(0)
        Dim session = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0)

        Try
            SyncLock _oLock
                session.findById("wnd[0]").maximize
                session.findById("wnd[0]/tbar[0]/okcd").Text = "/nF110"
                session.findById("wnd[0]/tbar[0]/btn[0]").press

                Thread.Sleep(500)

                'Проверка на совпадение по таблицам счетов и законченности прошлого прогона
                'Если прогон не закончен и счета присутсвуют
                'isAccountsPresentInOldIdentifier = true
                If isCompleteOldRun = False AndAlso isAccountsPresentInOldIdentifier Then
                    session.findById("wnd[0]/usr/ctxtF110V-LAUFD").text = paymentDate
                    session.findById("wnd[0]/usr/ctxtF110V-LAUFI").text = oldRunIdentifier
                    session.findById("wnd[0]/tbar[1]/btn[14]").press
                    Thread.Sleep(300)

                    Console.WriteLine(session.findbyid("wnd[0]/usr/tabsF110_TABSTRIP/tabpSTA/ssubSUBSCREEN_BODY:SAPF110V:0201/sub:SAPF110V:0201/txtF110V-STATU[1,0]").Text)
                    If session.findbyid("wnd[0]/usr/tabsF110_TABSTRIP/tabpSTA/ssubSUBSCREEN_BODY:SAPF110V:0201/sub:SAPF110V:0201/txtF110V-STATU[1,0]").Text = "Предложение по платежам создано" Then
                        If Not session.findbyid("wnd[0]/usr/tabsF110_TABSTRIP/tabpSTA/ssubSUBSCREEN_BODY:SAPF110V:0201/sub:SAPF110V:0201/txtF110V-STATU[2,0]", False) Is Nothing Then
                            Console.WriteLine(session.findbyid("wnd[0]/usr/tabsF110_TABSTRIP/tabpSTA/ssubSUBSCREEN_BODY:SAPF110V:0201/sub:SAPF110V:0201/txtF110V-STATU[2,0]").Text)
                            If session.findbyid("wnd[0]/usr/tabsF110_TABSTRIP/tabpSTA/ssubSUBSCREEN_BODY:SAPF110V:0201/sub:SAPF110V:0201/txtF110V-STATU[2,0]").Text = "Прогон ПрогПлатежей выполнен" Then
                                Console.WriteLine(session.findbyid("wnd[0]/usr/tabsF110_TABSTRIP/tabpSTA/ssubSUBSCREEN_BODY:SAPF110V:0201/sub:SAPF110V:0201/txtF110V-STATU[2,0]").Text)
                                'предыдуший прогон успешен 
                                tablesToRun.Rows(startSearchRow)("IsComplete") = True
                                tablesToRun.AcceptChanges()
                                SaveDataTableToFile(localFolder & "\" & tablesToRunFileNameXml, tablesToRun)
                                ' обновить данные в глобальной таблице
                                runControlTable = GetTableFromFile(localFolder, runControlTableInXML)
                                For i As Integer = 0 To runControlTable.Rows.Count - 1
                                    If runControlTable.Rows(i)("SheetName") = oldSheet Then
                                        runControlTable.Rows(i)("IsComplete") = True
                                        Exit For
                                    End If
                                Next
                                runControlTable.AcceptChanges()
                                SaveDataTableToFile(localFolder & "\" & runControlTableInXML, runControlTable)
                                runControlTable = New Data.DataTable()
                                'можно идти дальше
                                isCompleteOldRun = True
                            End If
                        End If
                    End If
                End If

                ' Если текущий прогон у нас уже был, надо его проверить, был ли он выпонен
                If isCompleteOldRun AndAlso isRunCreated Then
                    session.findById("wnd[0]/usr/ctxtF110V-LAUFD").text = paymentDate
                    session.findById("wnd[0]/usr/ctxtF110V-LAUFI").text = identifier
                    session.findById("wnd[0]/tbar[1]/btn[14]").press
                    Thread.Sleep(300)

                    Console.WriteLine(session.findbyid("wnd[0]/usr/tabsF110_TABSTRIP/tabpSTA/ssubSUBSCREEN_BODY:SAPF110V:0201/sub:SAPF110V:0201/txtF110V-STATU[1,0]").Text)
                    If session.findbyid("wnd[0]/usr/tabsF110_TABSTRIP/tabpSTA/ssubSUBSCREEN_BODY:SAPF110V:0201/sub:SAPF110V:0201/txtF110V-STATU[1,0]").Text = "Предложение по платежам создано" Then
                        If Not session.findbyid("wnd[0]/usr/tabsF110_TABSTRIP/tabpSTA/ssubSUBSCREEN_BODY:SAPF110V:0201/sub:SAPF110V:0201/txtF110V-STATU[2,0]", False) Is Nothing Then
                            Console.WriteLine(session.findbyid("wnd[0]/usr/tabsF110_TABSTRIP/tabpSTA/ssubSUBSCREEN_BODY:SAPF110V:0201/sub:SAPF110V:0201/txtF110V-STATU[2,0]").Text)
                            If session.findbyid("wnd[0]/usr/tabsF110_TABSTRIP/tabpSTA/ssubSUBSCREEN_BODY:SAPF110V:0201/sub:SAPF110V:0201/txtF110V-STATU[2,0]").Text = "Прогон ПрогрПлатежей выполнен" Then
                                'предыдуший прогон успешен 
                                tablesToRun.Rows(startSearchRow + 1)("IsComplete") = True
                                tablesToRun.AcceptChanges()
                                SaveDataTableToFile(localFolder & "\" & tablesToRunFileNameXml, tablesToRun)
                                ' обновить данные в глобальной таблице
                                runControlTable = GetTableFromFile(localFolder, runControlTableInXML)
                                For i As Integer = 0 To runControlTable.Rows.Count - 1
                                    If runControlTable.Rows(i)("SheetName") = sheetName Then
                                        runControlTable.Rows(i)("IsComplete") = True
                                        Exit For
                                    End If
                                Next
                                runControlTable.AcceptChanges()
                                SaveDataTableToFile(localFolder & "\" & runControlTableInXML, runControlTable)
                                runControlTable = New Data.DataTable()
                            End If
                        End If
                    End If

                    'Выходим все хорошо, это повторный запуск, проверки выполнения листа
                    isComplete = True
                    isRunCreated = False
                    ReturnToMainWindow(session)
                    Exit Sub
                End If

                'нашли старый прогон? он исполнен? мы текущий прогон исполнить не можем
                If isCompleteOldRun = False AndAlso isAccountsPresentInOldIdentifier Then
                    isComplete = True
                    isRunCreated = False
                    ReturnToMainWindow(session)
                    Exit Sub
                End If

                session.findById("wnd[0]/usr/ctxtF110V-LAUFD").text = paymentDate
                session.findById("wnd[0]/usr/ctxtF110V-LAUFI").text = identifier
                session.findById("wnd[0]/tbar[1]/btn[14]").press

                Dim controlNumber As Integer = 200
                While Not session.findbyid("wnd[0]/usr/tabsF110_TABSTRIP/tabpSTA/ssubSUBSCREEN_BODY:SAPF110V:0201/sub:SAPF110V:0201/txtF110V-STATU[0,0]").Text = "Параметры еще не введены."
                    controlNumber = controlNumber - 1
                    identifier = ChangeIdentifier(identifier)
                    session.findById("wnd[0]/usr/ctxtF110V-LAUFI").text = identifier
                    session.findById("wnd[0]/tbar[1]/btn[14]").press
                    If controlNumber = 0 Then
                        Throw New Exception("identifier уже больше ста, так не должно быть")
                    End If
                End While

                session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR").select

                'Перешли или нет на вкладку Параметры
                If Not session.findbyid("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssubSUBSCREEN_BODY:SAPF110V:0202/boxTXT01", False) Is Nothing Then
                    Console.WriteLine(session.findbyid("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssubSUBSCREEN_BODY:SAPF110V:0202/boxTXT01").Text)
                Else
                    Throw New Exception("Не перешли на вкладку Параметры")
                End If

                session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssubSUBSCREEN_BODY:SAPF110V:0202/tblSAPF110VCTRL_FKTTAB/txtF110V-BUKLS[0,0]").text = be
                session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssubSUBSCREEN_BODY:SAPF110V:0202/tblSAPF110VCTRL_FKTTAB/ctxtF110V-ZWELS[1,0]").text = paymentMethod
                session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssubSUBSCREEN_BODY:SAPF110V:0202/tblSAPF110VCTRL_FKTTAB/ctxtF110V-NEDAT[2,0]").text = paymentDate.AddDays(1)
                session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssubSUBSCREEN_BODY:SAPF110V:0202/subSUBSCR_SEL:SAPF110V:7004/btn%_R_LIFNR_%_APP_%-VALU_PUSH").press

                UploadFromFileInMultipleSelectionWindow(session, lenderAccountFile, localFolder)

                If isPaymentsForeign Then
                    session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpPAR/ssubSUBSCREEN_BODY:SAPF110V:0202/ctxtF110V-KURST").text = "M"
                End If

                'Переходим на вкладку Произвольный выбор
                session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL").select

                'Перешли или нет на вкладку Произвольный выбор
                If Not session.findbyid("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssubSUBSCREEN_BODY:SAPF110V:0203/box%#AUTOTEXT001", False) Is Nothing Then
                    Console.WriteLine(session.findbyid("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssubSUBSCREEN_BODY:SAPF110V:0203/box%#AUTOTEXT001").Text)
                Else
                    Throw New Exception("Не перешли на вкладку Произвольный выбор")
                End If

                session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssubSUBSCREEN_BODY:SAPF110V:0203/sub:SAPF110V:0203/ctxtF110V-TEXT1[0,11]").setFocus
                session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssubSUBSCREEN_BODY:SAPF110V:0203/sub:SAPF110V:0203/ctxtF110V-TEXT1[0,11]").caretPosition = 0
                session.findById("wnd[0]").sendVKey(4)

                Dim numberOfRowsInTable As Integer = 2
                isExit = False
                While (isExit = False AndAlso numberOfRowsInTable < 50)
                    numberOfRowsInTable += 1
                    Try
                        If session.findById("wnd[1]/usr/lbl[1," & numberOfRowsInTable.ToString() & "]").Text = "Блокировка платежа" Then
                            session.findById("wnd[1]/usr/lbl[1," & numberOfRowsInTable.ToString() & "]").setFocus
                            session.findById("wnd[1]/usr/lbl[1," & numberOfRowsInTable.ToString() & "]").caretPosition = 35
                            session.findById("wnd[1]/tbar[0]/btn[0]").press
                            isExit = True
                        End If
                    Catch ex As Exception
                        Throw New Exception("Закладку «Произвольный выбор». не могу найти 'Блокировка платежа'.")
                    End Try
                End While

                ' Исключить значения первое поле = true
                If session.findbyid("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssubSUBSCREEN_BODY:SAPF110V:0203/sub:SAPF110V:0203/chkF110V-XIGNO[0,44]").Selected = False Then
                    session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssubSUBSCREEN_BODY:SAPF110V:0203/sub:SAPF110V:0203/chkF110V-XIGNO[0,44]").selected = True
                End If

                session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssubSUBSCREEN_BODY:SAPF110V:0203/sub:SAPF110V:0203/txtF110V-LIST1[1,11]").text = "A,X,W,Y,B,Q,N,S,G,F,Z,E,D"
                session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssubSUBSCREEN_BODY:SAPF110V:0203/sub:SAPF110V:0203/chkF110V-XIGNO[0,44]").setFocus

                session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssubSUBSCREEN_BODY:SAPF110V:0203/sub:SAPF110V:0203/ctxtF110V-TEXT1[3,11]").setFocus
                session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssubSUBSCREEN_BODY:SAPF110V:0203/sub:SAPF110V:0203/ctxtF110V-TEXT1[3,11]").caretPosition = 0
                session.findById("wnd[0]").sendVKey(4)

                numberOfRowsInTable = 2
                isExit = False
                While (isExit = False AndAlso numberOfRowsInTable < 50)
                    numberOfRowsInTable += 1
                    Try
                        If session.findById("wnd[1]/usr/lbl[1," & numberOfRowsInTable.ToString() & "]").Text = "Собственный банк" Then
                            session.findById("wnd[1]/usr/lbl[1," & numberOfRowsInTable.ToString() & "]").setFocus
                            session.findById("wnd[1]/usr/lbl[1," & numberOfRowsInTable.ToString() & "]").caretPosition = 35
                            session.findById("wnd[1]/tbar[0]/btn[0]").press
                            isExit = True
                        End If
                    Catch ex As Exception
                        Throw New Exception("Закладку «Произвольный выбор». не могу найти 'Собственный банк'.")
                    End Try
                End While

                ' Исключить значения второе поле = false
                If session.findbyid("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssubSUBSCREEN_BODY:SAPF110V:0203/sub:SAPF110V:0203/chkF110V-XIGNO[3,44]").Selected = True Then
                    session.findbyid("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssubSUBSCREEN_BODY:SAPF110V:0203/sub:SAPF110V:0203/chkF110V-XIGNO[3,44]").Selected = False
                End If

                session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssubSUBSCREEN_BODY:SAPF110V:0203/sub:SAPF110V:0203/txtF110V-LIST1[4,11]").text = ownBank
                'Просто переход на следующее поле для закрепления ввода
                session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssubSUBSCREEN_BODY:SAPF110V:0203/sub:SAPF110V:0203/txtF110V-LIST2[5,11]").setFocus
                session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSEL/ssubSUBSCREEN_BODY:SAPF110V:0203/sub:SAPF110V:0203/txtF110V-LIST2[5,11]").caretPosition = 0

                'Переходим на вкладку ДополнитЖурнал
                session.findbyid("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG").Select

                'Перешли или нет на вкладку ДополнитЖурнал выбор
                If Not session.findbyid("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG/ssubSUBSCREEN_BODY:SAPF110V:0204/boxTXT00", False) Is Nothing Then
                    Console.WriteLine(session.findbyid("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG/ssubSUBSCREEN_BODY:SAPF110V:0204/boxTXT00").Text)
                Else
                    Throw New Exception("Не перешли на вкладку ДополнитЖурнал")
                End If

                session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG/ssubSUBSCREEN_BODY:SAPF110V:0204/chkF110V-XTRFA").selected = True
                session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG/ssubSUBSCREEN_BODY:SAPF110V:0204/chkF110V-XTRZW").selected = True
                session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG/ssubSUBSCREEN_BODY:SAPF110V:0204/chkF110V-XTRBL").selected = True
                session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG/ssubSUBSCREEN_BODY:SAPF110V:0204/chkF110V-XTRBL").setFocus

                'Переходим на вкладку Статус
                'При этом сохраняем введенное 
                session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpSTA").select

                Thread.Sleep(300)
                ' Появилось окно Закончить обработку?
                If Not session.findbyid("wnd[1]", False) Is Nothing Then
                    session.findbyid("wnd[1]/usr/btnSPOP-OPTION1").Press
                Else
                    Throw New Exception("Не Появилось окно Закончить обработку")
                End If

                Thread.Sleep(300)
                'ищем окно Статус
                If Not session.findbyid("wnd[0]/usr/tabsF110_TABSTRIP/tabpSTA/ssubSUBSCREEN_BODY:SAPF110V:0201/boxTXT02", False) Is Nothing Then
                    If Not session.findbyid("wnd[0]/usr/tabsF110_TABSTRIP/tabpSTA/ssubSUBSCREEN_BODY:SAPF110V:0201/sub:SAPF110V:0201/txtF110V-STATU[0,0]").Text = "Параметры введены." Then
                        Throw New Exception("После ввода всех параметров не сохранилось 'Параметры введены'")
                    End If
                Else
                    Throw New Exception("Не Появилось окно Статус")
                End If

                'Предложение
                ' Для формирования предложения на оплату необходимо вверху нажать кнопку «Предложение»
                session.findbyid("wnd[0]/tbar[1]/btn[13]").Press
                Thread.Sleep(300)
                'ищем окно Запланировать предложение
                If Not session.findbyid("wnd[1]", False) Is Nothing Then
                    If Not session.findbyid("wnd[1]").Text = "Запланировать предложение" Then
                        Throw New Exception("Не открылось окно 'Запланировать предложение'")
                    End If
                Else
                    Throw New Exception("Не Появилось окно 'Запланировать предложение")
                End If

                Try
                    'Немедленный запуск
                    session.findbyid("wnd[1]/usr/chkF110V-XSTRF").Selected = True
                Catch ex As Exception
                End Try

                Try
                    'Создание носителей плат.
                    session.findbyid("wnd[1]/usr/chkF110V-XMITD").Selected = True
                Catch ex As Exception
                End Try

                'Запланировать   (Enter)
                session.findbyid("wnd[1]/tbar[0]/btn[0]").Press

                Thread.Sleep(1000)
                'Жмем статус
                session.findbyid("wnd[0]/tbar[1]/btn[14]").Press

                'Еще раз нажимаем кнопку «Статус» или «Enter» до появления сообщения: «Предложение по платежам создано»
                ' wait window
                isExit = False
                Dim _timeout As DateTime = DateTime.Now.AddSeconds(1200)
                While (isExit = False)
                    If session.findbyid("wnd[0]/usr/tabsF110_TABSTRIP/tabpSTA/ssubSUBSCREEN_BODY:SAPF110V:0201/sub:SAPF110V:0201/txtF110V-STATU[1,0]").Text = "Предложение по платежам создано" Then
                        isExit = True
                    Else
                        If (DateTime.Now > _timeout) Then
                            Throw New Exception("Не появилось сообщение: 'Предложение по платежам создано'")
                        End If
                        session.findbyid("wnd[0]/tbar[1]/btn[14]").Press
                        Thread.Sleep(300)
                    End If
                End While
                ' wait window

                'Нажимаем кнопку "Предложение"  для проверки созданного предложения прогона.
                session.findbyid("wnd[0]/tbar[1]/btn[21]").Press
                Thread.Sleep(300)

                ' если вдруг вылезла такая обшика, по формированию листа
                ' нужно её откатить 
                If Not session.findbyid("wnd[0]/usr/lblREGUH-LAUFI", False) Is Nothing Then
                    Dim paneText As String = session.findbyid("wnd[0]/sbar/pane[0]").Text
                    If paneText.Contains("не фигурирует, исправьте") Then
                        session.findById("wnd[0]/tbar[0]/btn[15]").press
                        session.findById("wnd[0]/mbar/menu[1]/menu[1]/menu[3]").select
                        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
                        Thread.Sleep(300)
                        ReturnToMainWindow(session)

                        Throw New Exception("Нужно сделать запись что не получилось создать прогон")
                        'Нужно сделать запись что не получилось создать прогон
                    End If
                End If

                'Проверяем что все нашли
                'Отправляем почту на получателя, это работа макроса.
                screenShotFileName = RemoveUnnecessaryChar(sheetName) & "_ScrSht"
                screenShotFileName = ChangeScreenShotFileName(screenShotFileName)
                TakeScreenShot(localFolder, screenShotFileName)
                screenShotFileName = screenShotFileName & ".png"
                Thread.Sleep(1000)
                CopyFile(localFolder, screenShotFileName, pathToMailAttachments, screenShotFileName)
                Dim tableErrors_F110 As Data.DataTable = GetTableFromFile(localFolder, xmlNameErrors_F110)
                columns = tableErrors_F110.Columns
                If Not columns.Contains("Identifier") Then
                    tableErrors_F110.Columns.Add("Identifier", Type.GetType("System.String"))
                End If
                If Not session.findbyid("wnd[0]/usr/boxTXT01", False) Is Nothing Then
                    ' ищем ошибку, красный цвет
                    ' @5B\QПлатеж@ - все ровно
                    ' @5C\QОсобая ситуация@ - наш вариант
                    Dim tableShell As Object
                    Dim rowShellCount As Integer
                    tableShell = session.findbyid("wnd[0]/usr/cntlCONTAINER_0111/shellcont/shell")
                    rowShellCount = tableShell.rowCount
                    Dim color As String
                    Dim isExceptError As Boolean = False
                    For i As Integer = 0 To rowShellCount - 1
                        session.findById("wnd[0]/usr/cntlCONTAINER_0111/shellcont/shell").setCurrentCell(i, "LIFNR")
                        session.findById("wnd[0]/usr/cntlCONTAINER_0111/shellcont/shell").selectedRows = i
                        color = session.findbyid("wnd[0]/usr/cntlCONTAINER_0111/shellcont/shell").GetCellValue(i, "EXCEPT")
                        If color <> "@5B\QПлатеж@" Then
                            'If color = "@5B\QПлатеж@" Then
                            isExceptError = True
                            ' заполнить коллекцию строк ошибок
                            ' потом это нужно сохранить в файл
                            Dim row As Data.DataRow = tableErrors_F110.NewRow()
                            row("Кредитор") = session.findbyid("wnd[0]/usr/cntlCONTAINER_0111/shellcont/shell").GetCellValue(i, "LIFNR")
                            row("СуммПлт_ВВ") = TranslateStringSumToNumber(session.findbyid("wnd[0]/usr/cntlCONTAINER_0111/shellcont/shell").GetCellValue(i, "RBETR"))
                            row("Сумма") = TranslateStringSumToNumber(session.findbyid("wnd[0]/usr/cntlCONTAINER_0111/shellcont/shell").GetCellValue(i, "RWBTR"))
                            row("Identifier") = identifier
                            tableErrors_F110.Rows.Add(row)
                        End If
                        numberOfPayments = numberOfPayments + 1
                        totalRunAmount = totalRunAmount + TranslateStringSumToNumber(session.findbyid("wnd[0]/usr/cntlCONTAINER_0111/shellcont/shell").GetCellValue(i, "RWBTR"))
                        currentCurrency = session.findbyid("wnd[0]/usr/cntlCONTAINER_0111/shellcont/shell").GetCellValue(i, "WAERS")
                    Next

                    ' сохраним если были ошибки
                    If isExceptError Then
                        tableErrors_F110.AcceptChanges()
                        SaveDataTableToFile(localFolder & "\" & xmlNameErrors_F110, tableErrors_F110)
                        If SaveTableToExcel(localFolder, xlsbNameNotIncludedTable, tableErrors_F110, 2, exceptionMessage) = False Then
                            Throw New Exception("Не могу сохранить данные в '" & xlsbNameNotIncludedTable & "'")
                        End If

                        ' перейти на страницу для выгрузки журнала
                        session.findById("wnd[0]/tbar[0]/btn[15]").press
                        session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG").select
                        session.findById("wnd[0]/usr/tabsF110_TABSTRIP/tabpLOG/ssubSUBSCREEN_BODY:SAPF110V:0204/btnPUSH_VOPR").press
                        session.findById("wnd[0]/mbar/menu[1]/menu[4]/menu[1]").select

                        'выгружаем
                        session.findById("wnd[1]/usr/radRB_OTHERS").setFocus
                        session.findById("wnd[1]/usr/radRB_OTHERS").select
                        session.findById("wnd[1]/usr/cmbG_LISTBOX").setFocus
                        session.findById("wnd[1]/usr/cmbG_LISTBOX").Key = "04"

                        Thread.Sleep(500)
                        TaskPressButtonOk(session)
                        If (IsGuiModalWindow(session, "wnd[1]")) Then
                            session.findById("wnd[1]/usr/ctxtDY_PATH").text = localFolder
                            session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = errorsLogFileNameTxt
                            session.findById("wnd[1]/tbar[0]/btn[11]").press
                        Else
                            'Real Mode
                            TryToFoundWindowAndSetForeground("Открыть")
                            SaveAsWindow(errorsLogFileNameTxt, localFolder, "Сохранение", "Сохранить")
                        End If

                        ' Ищем наш файл лога, сохраняем его и добавляем на лист 2
                        'DeleteFile(localFolder, "errorsLog.xlsb")
                        DeleteFile(localFolder, errorsLogFileNameXlsb)
                        'FoundExcelAndSaveIt(localFolder, "errorsLog.xlsb")
                        If TryFindNotepadAndCloseIt() = False Then
                            Throw New Exception("Нет выгрузки ")
                        End If

                        If CheckFileExists(localFolder, errorsLogFileNameTxt) = False Then
                            Throw New Exception("Нет выгрузки ")
                        End If

                        RenameFile(localFolder, errorsLogFileNameTxt, localFolder, errorsLogFileNameXml)

                        OpenXmlFileAndSaveIt(localFolder, errorsLogFileNameXml, errorsLogFileNameXlsb)

                        Dim dataFromExcel As System.Data.DataTable = GetDatatableFromExcel(localFolder, errorsLogFileNameXlsb, 1)
                        columns = dataFromExcel.Columns
                        If Not columns.Contains("Identifier") Then
                            dataFromExcel.Columns.Add("Identifier", Type.GetType("System.String"))
                        End If
                        For i As Integer = 0 To dataFromExcel.Rows.Count - 1
                            dataFromExcel.Rows(i)("Identifier") = identifier
                        Next
                        columns = tableJournals.Columns
                        If Not columns.Contains("Identifier") Then
                            tableJournals.Columns.Add("Identifier", Type.GetType("System.String"))
                        End If
                        tableJournals = MergeTwoTables(tableJournals, dataFromExcel)
                        tableJournals.TableName = "JournalsLog"
                        'тут надо dataFromExcel типа сохранить на лист, это у нас то что не вошло в выбор документов
                        tableJournals.AcceptChanges()
                        SaveDataTableToFile(localFolder & "\" & xmlNameJournals, tableJournals)
                        If SaveTableToExcel(localFolder, xlsbNameNotIncludedTable, tableJournals, 3, exceptionMessage) = False Then
                            Throw New Exception("Не могу сохранить данные в '" & xlsbNameNotIncludedTable & "'")
                        End If
                    End If

                    ' Сохраняем в MacrosForRun.xlsb данные перед отправкой почты
                    macrosForRunFileName = "MacrosForRun_" & sheetName & "_.xlsb"
                    macrosForRunFileName = ChangeScreenShotFileName(macrosForRunFileName)
                    DeleteFile(localFolder, macrosForRunFileName)
                    CopyFile(localFolder, "MacrosForRun.xlsb", localFolder, macrosForRunFileName)
                    UpdateMacrosForRun(localFolder, isPaymentsForeign, be, identifier, paymentDate, ownBank, paymentMethod, numberOfPayments, totalRunAmount, macrosForRunFileName, currentCurrency)
                    Thread.Sleep(1000)
                    CopyFile(localFolder, macrosForRunFileName, pathToMailAttachments, macrosForRunFileName)

                    isRunCreated = True
                Else
                    Throw New Exception("Не смог провалится в окно 'Платежи / исключения'")
                    'Нужно сделать запись что не получилось создать прогон
                End If

                ReturnToMainWindow(session)

                'нужно обновить у таблицы tablesToRun identifier and IsRunCreated
                UpdateIdentifierInXML(localFolder, identifier, tablesToRun, tablesToRunFileNameXml, sheetName, isRunCreated)
                'нужно обновить у глобальной таблицы identifier
                'Dim runControlTableInXML As String = "RunControlTable.XML"
                'Dim runControlTable As Data.DataTable = GetTableFromFile(localFolder, runControlTableInXML)
                runControlTable = GetTableFromFile(localFolder, runControlTableInXML)
                UpdateIdentifierInXML(localFolder, identifier, runControlTable, runControlTableInXML, sheetName, isRunCreated)
                runControlTable = New Data.DataTable()
                isComplete = True

            End SyncLock

        Catch ex As Exception
            exceptionMessage = ex.Message
            If (Not ex.InnerException Is Nothing) Then
                exceptionMessage = exceptionMessage & " Inner Exception  " & ex.InnerException.Message
            End If
            isComplete = False

            Console.WriteLine("Exception        " & ex.GetType().ToString())
            Console.WriteLine("Message          " & ex.Message)
            If (Not ex.InnerException Is Nothing) Then
                Console.WriteLine("Inner Exception  " & ex.InnerException.Message)
            End If
        Finally
            session = Nothing
            Console.WriteLine("exceptionMessage " & exceptionMessage)
            Console.WriteLine("isComplete " & isComplete)
            Console.WriteLine("Первичный поток Id {0} Is Ended", Thread.CurrentThread.ManagedThreadId)
            Console.WriteLine("isRunCreated = {0}", isRunCreated)
            Console.ReadKey()
        End Try

        Console.WriteLine("Первичный поток: Id {0} Is Ended", Thread.CurrentThread.ManagedThreadId)
        Console.WriteLine("isRunCreated = {0}", isRunCreated)
        Console.ReadKey()

    End Sub

    Private Sub OpenXmlFileAndSaveIt(folder As String, xmlFileName As String, xlsbFileName As String)
        Dim xlApp As Excel.Application = New Excel.Application()
        Dim xlWorkBook As Excel.Workbook

        xlWorkBook = xlApp.Workbooks.Open(folder & "\" & xmlFileName)
        xlWorkBook.SaveAs(folder & "\" & xlsbFileName, 50)
        xlWorkBook.Close(True)
        xlApp.Quit()
        ReleaseObject(xlApp)
        ReleaseObject(xlWorkBook)
        xlWorkBook = Nothing
        xlApp = Nothing
    End Sub

    Private Sub RenameFile(localFolder As String, localFileName As String, remoteFolder As String, remoteFileName As String)
        Try
            File.Delete(remoteFolder & "\" & remoteFileName)
            File.Move(localFolder & "\" & localFileName, remoteFolder & "\" & remoteFileName)
        Catch ex As Exception
        End Try
    End Sub

    Private Function CheckFileExists(folder As String, fileName As String) As Boolean
        Dim curFile = folder + "\\" + fileName
        Return If(File.Exists(curFile), True, False)
    End Function

    Private Function TryFindNotepadAndCloseIt() As Boolean
        Dim isExit As Boolean = False
        Dim isComplete As Boolean = False
        Dim timeout As DateTime = DateTime.Now.AddSeconds(5)
        While (isExit = False)
            Dim localByName As Process() = Process.GetProcessesByName("notepad")
            For Each proc As Process In localByName
                proc.Kill()
                isExit = True
                isComplete = True
            Next
            If (isExit = False AndAlso DateTime.Now > timeout) Then
                isComplete = False
                isExit = True
            End If
        End While

        Return isComplete
    End Function

    Private Function ChangeScreenShotFileName(screenShotFileName As String) As String
        Return Replace(screenShotFileName, " ", "_")
    End Function

    Private Sub UpdateMacrosForRun(localFolder As String, isPaymentsForeign As Boolean, be As String, identifier As String, paymentDate As Date, ownBank As String, paymentMethod As String, numberOfPayments As Integer, totalRunAmount As Double, fileName As String, currentCurrency As String)
        Dim xlApp As Excel.Application = New Excel.Application()
        Dim xlWorkBook As Object = New Object()
        Dim xlWorkSheet As Object = New Object()
        Dim misValue As Object = Reflection.Missing.Value
        Dim fullFileName As String = localFolder & "\" & fileName

        Try
            xlWorkBook = xlApp.Workbooks.Open(fullFileName)
            xlWorkSheet = CType(xlWorkBook.Sheets(1), Excel.Worksheet)

            'edit the cell with new value
            If isPaymentsForeign Then
                xlWorkSheet.Cells(2, 1) = "Да"
            Else
                xlWorkSheet.Cells(2, 1) = ""
            End If
            xlWorkSheet.Cells(2, 2) = be
            xlWorkSheet.Cells(2, 3) = identifier
            xlWorkSheet.Cells(2, 4) = paymentDate
            xlWorkSheet.Cells(2, 5) = ownBank
            xlWorkSheet.Cells(2, 6) = paymentMethod
            xlWorkSheet.Cells(2, 7) = numberOfPayments
            xlWorkSheet.Cells(2, 8) = totalRunAmount & ", " & currentCurrency

            xlWorkBook.Save()
            xlWorkBook.Close()
            xlApp.Quit()

            ReleaseObject(xlApp)
            ReleaseObject(xlWorkBook)
            ReleaseObject(xlWorkSheet)

        Catch e As Exception
        Finally
            xlApp = Nothing
            xlWorkBook = Nothing
            xlWorkSheet = Nothing
        End Try
    End Sub

    Private Function TranslateStringSumToNumber(ByVal str As String) As Double
        ': symmPlatVV : "85.451.402,10-" : String
        Dim result As Double
        str = RemoveUnnecessaryChar(str)
        Dim strLenght As Integer = str.Length - 1
        For i As Integer = strLenght To 0 Step -1
            If Not Char.IsDigit(Convert.ToChar(str(i))) Then
                str = Left(str, i)
            Else
                Exit For
            End If
        Next
        result = Convert.ToDouble(str)
        If result < 0 Then
            result = -1 * result
        End If

        Return result
    End Function

    Private Sub UpdateIdentifierInXML(localFolder As String, identifier As String, inputTable As Data.DataTable, xmlFileName As String, sheetName As String, isRunCreated As Boolean)
        For i As Integer = inputTable.Rows.Count - 1 To 0 Step -1
            If inputTable.Rows(i)("SheetName") = sheetName Then
                inputTable.Rows(i)("Identifier") = identifier
                inputTable.Rows(i)("isRunCreated") = isRunCreated
                inputTable.AcceptChanges()
                Exit For
            End If
        Next
        SaveDataTableToFile(localFolder & "\" & xmlFileName, inputTable)
    End Sub

    Private Function SaveTableToExcel(ByVal localFolder As String, ByVal excelFile As String, ByVal tempTable As Data.DataTable, ByVal sheetNumber As Integer, ByRef exceptionMessage As String) As Boolean
        Dim xlApp As Excel.Application = New Excel.Application()
        Dim xlWorkBook As Object = New Object()
        Dim xlWorkSheet As Object = New Object()
        Dim misValue As Object = Reflection.Missing.Value
        Dim isSaved As Boolean = False

        Dim fullFileName As String = localFolder & "\" & excelFile

        Try
            xlWorkBook = xlApp.Workbooks.Open(fullFileName)
            xlWorkSheet = CType(xlWorkBook.Sheets(sheetNumber), Excel.Worksheet)

            Dim timeArray(tempTable.Rows.Count, tempTable.Columns.Count) As Object
            Dim row As Integer, col As Integer

            For row = 0 To tempTable.Rows.Count - 1
                For col = 0 To tempTable.Columns.Count - 1
                    timeArray(row, col) = tempTable.Rows(row).Item(col)
                Next
            Next

            'col = 0
            'For Each column As DataColumn In tempTable.Columns
            '    xlWorkSheet.Cells(1, col + 1) = column.ColumnName
            '    col += 1
            'Next

            xlWorkSheet.Range(xlWorkSheet.Cells(2, 1), xlWorkSheet.Cells(tempTable.Rows.Count + 1, tempTable.Columns.Count)).Value = timeArray

            xlWorkBook.Save()
            xlWorkBook.Close()
            xlApp.Quit()

            isSaved = True
        Catch ex As Exception
            exceptionMessage = ex.Message
            If (Not ex.InnerException Is Nothing) Then
                exceptionMessage = exceptionMessage & " Inner Exception : " & ex.InnerException.Message
            End If
        Finally
            ReleaseObject(xlWorkSheet)
            ReleaseObject(xlWorkBook)
            ReleaseObject(xlApp)

            KillExcell()
        End Try

        Return isSaved
    End Function

    Private Sub KillExcell()
        Dim proc As Process
        For Each proc In Process.GetProcessesByName("EXCEL")
            proc.Kill()
        Next
    End Sub

    Private Function MergeTwoTables(ByVal collection1 As Data.DataTable, ByVal collection2 As Data.DataTable) As Data.DataTable

        If (collection1 IsNot Nothing) Then
            If (collection1.Rows.Count > 0) Then
                For i As Integer = 0 To collection2.Rows.Count - 1
                    collection1.ImportRow(collection2.Rows(i))
                Next
                Return collection1
            End If
        End If

        If (collection2 IsNot Nothing) Then
            If (collection2.Rows.Count > 0) Then
                For i As Integer = 0 To collection1.Rows.Count - 1
                    collection2.ImportRow(collection1.Rows(i))
                Next
                Return collection2
            End If
        End If

        Throw New Exception("MergeTwoTables")
    End Function

    Private Sub DeleteFile(localFolder As String, fileToRemove As String)
        Try
            File.Delete(localFolder & "\" & fileToRemove)
        Catch ex As Exception
            Throw New Exception("Не могу удалить файл " & fileToRemove)
        End Try
    End Sub

    Private Sub CopyFile(localFolder As String, fileToCopy As String, destinationFolder As String, newCopy As String)
        Try
            File.Copy(localFolder & "\" & fileToCopy, destinationFolder & "\" & newCopy, True)
        Catch ex As Exception
            Throw New Exception("Не могу скопировать файл " & fileToCopy & " в " & newCopy & " " & ex.Message)
        End Try
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
            Using oledbCmd = New System.Data.OleDb.OleDbCommand(sql, oledbCnn)
                Using oledbAdaper As System.Data.OleDb.OleDbDataAdapter = New OleDbDataAdapter(oledbCmd)
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

    Private Sub SaveDataTableToFile(fileName As String, table As System.Data.DataTable)
        Dim Stream As FileStream = New FileStream(fileName, FileMode.Create)
        Dim serializer As XmlSerializer = New XmlSerializer(table.GetType())
        serializer.Serialize(Stream, table)
        Stream.Close()
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

    Private Sub CreateTXTFileFromCollection(ByVal tempTable As Data.DataTable, ByVal columnName As String, ByVal localFolder As String, ByVal fileName As String)
        Dim fullFileName = localFolder + "\" + fileName
        Dim _file As New FileInfo(fullFileName)
        Dim _streamWriter As StreamWriter = _file.CreateText()

        For i As Integer = 0 To tempTable.Rows.Count - 1
            _streamWriter.WriteLine(tempTable.Rows(i)(columnName).ToString())
        Next

        _streamWriter.Close()
    End Sub

    Private Function GetOwnBank(ByVal sheetName As String) As String
        If Left(sheetName, 2) = "DB" Then
            Return "DBSBK"
        End If

        Return "CITBK"
    End Function

    Private Sub ReturnToMainWindow(ByVal session As Object)
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
        session.findbyid("wnd[0]/tbar[0]/btn[0]").Press
    End Sub

    Private Sub CheckTimeout(ByVal timeout As Date)
        If (DateTime.Now > timeout) Then
            Throw New Exception("Не открылась транзакция")
        End If
    End Sub

    Private Sub UploadFromFileInMultipleSelectionWindow(ByVal session As Object, ByVal fileToUpload As String, ByVal localFolder As String)
        Dim timeout As DateTime
        Dim isExit As Boolean
        ' wait window
        isExit = False
        timeout = DateTime.Now.AddSeconds(1)
        While (isExit = False)
            If Not session.findbyid("wnd[1]/tbar[0]/btn[23]", False) Is Nothing Then
                isExit = True
            Else
                If (DateTime.Now > timeout) Then
                    Throw New Exception("Не открылось окно Многократный выбор")
                End If
            End If
        End While
        ' wait window

        SetForegroundWindow(session.findById("wnd[1]").Handle)
        session.findById("wnd[1]/tbar[0]/btn[16]").press

        TaskPressIpmportButton(session)

        If (IsGuiModalWindow(session, "wnd[2]")) Then
            session.findById("wnd[2]/usr/ctxtDY_PATH").text = localFolder
            session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = fileToUpload
            session.findById("wnd[2]/tbar[0]/btn[0]").press
        Else
            'Real Mode
            TryToFoundWindowAndSetForeground("Открыть")
            SaveAsWindow(fileToUpload, localFolder, "Открыть", "Открыть")
        End If
        ' жмем Скопировать (F8)
        ' wait window
        isExit = False
        timeout = DateTime.Now.AddSeconds(2)
        While (isExit = False)
            If Not session.findbyid("wnd[1]/tbar[0]/btn[8]", False) Is Nothing Then
                isExit = True
            Else
                If (DateTime.Now > timeout) Then
                    Throw New Exception("Не открылось окно после загрузки из файла")
                End If
            End If
        End While
        ' wait window
        SetForegroundWindow(session.findById("wnd[1]").Handle)
        session.findById("wnd[1]/tbar[0]/btn[8]").press
    End Sub

    Private Sub TaskPressButtonOk(ByVal session As Object)
        Dim taskPressButton As Task = New Task(AddressOf PressButtonOk, session)
        Try
            taskPressButton.Start()
            taskPressButton.Wait(100)
            _thread.Abort()
            Thread.Sleep(300)
            taskPressButton.Dispose()
        Catch ex As Exception
            exceptionMessage = ex.Message
        End Try
    End Sub

    Private Sub TaskPressIpmportButton(ByVal session As Object)
        Dim taskPressButton As Task = New Task(AddressOf PressIpmportButton, session)
        Try
            taskPressButton.Start()
            taskPressButton.Wait(300)
            _thread.Abort()
            Thread.Sleep(300)
            taskPressButton.Dispose()
        Catch ex As Exception
        End Try
    End Sub

    Private Sub PressButtonOk(ByVal session As Object)

        _thread = Thread.CurrentThread
        session.findbyid("wnd[1]/tbar[0]/btn[0]").Press

    End Sub

    Private Sub PressIpmportButton(ByVal session As Object)

        _thread = Thread.CurrentThread
        session.findById("wnd[1]/tbar[0]/btn[23]").press

    End Sub

    Private Function InvokeButtonOk(elementCollectionAll As AutomationElementCollection, fieldName As String) As Boolean

        For Each autoElement As AutomationElement In elementCollectionAll
            If (autoElement.Current.ClassName.Equals("Button")) Then
                If (autoElement.Current.Name = fieldName) Then
                    Dim btnPattern As InvokePattern = TryCast(autoElement.GetCurrentPattern(InvokePattern.Pattern), InvokePattern)
                    autoElement.SetFocus()
                    SetForegroundWindow(autoElement.Current.NativeWindowHandle)
                    SendMessageW(autoElement.Current.NativeWindowHandle, BM_CLICK, IntPtr.Zero, IntPtr.Zero)
                    Try
                        SendMessageW(autoElement.Current.NativeWindowHandle, BM_CLICK, IntPtr.Zero, IntPtr.Zero)
                    Catch ex As Exception
                    End Try
                    Return True
                End If
            End If
        Next

        Return False
    End Function

    Private Function SendFileNameToDialogBox(elementCollectionAll As AutomationElementCollection, fullfileName As String) As Boolean
        For Each autoElement As AutomationElement In elementCollectionAll
            Dim WM_SETTEXT As Long = &HC
            If autoElement.Current.Name.Equals("Имя файла:  ") And autoElement.Current.ClassName.Contains("Edit") Then
                autoElement.SetFocus()
                SendMessageHM(autoElement.Current.NativeWindowHandle, WM_SETTEXT, 0, fullfileName)
                Return True
            End If
        Next

        Return False
    End Function

    Private Function GetHWNDWiondow(ByVal windowHeader As String, ByVal timeout As DateTime) As IntPtr

        Dim hWindow As IntPtr = New IntPtr()
        Dim isExit As Boolean = False

        While (isExit = False)

            hWindow = FindWindow("#32770", windowHeader)
            If (Not IsValidHandle(hWindow)) Then
                If (DateTime.Now > timeout) Then
                    Throw New ArgumentNullException("Cannot found launched window \"" + windowHeader + " \ "")
                End If
            Else
                isExit = True
            End If
        End While

        Return hWindow
    End Function

    Private Function IsValidHandle(ByVal hWindow As IntPtr) As Boolean
        Return hWindow <> IntPtr.Zero
    End Function

    Private Function GetChildWindows(ByVal parent As IntPtr) As List(Of IntPtr)

        Dim result As List(Of IntPtr) = New List(Of IntPtr)()
        Dim listHandle As GCHandle = GCHandle.Alloc(result)
        Try
            Dim childProc As EnumWindowProc = New EnumWindowProc(AddressOf EnumWindow)
            EnumChildWindows(parent, childProc, GCHandle.ToIntPtr(listHandle))
        Finally
            If (listHandle.IsAllocated) Then
                listHandle.Free()
            End If
        End Try

        Return result

    End Function

    Private Function EnumWindow(ByVal handle As IntPtr, ByVal pointer As IntPtr) As Boolean

        Dim gch As GCHandle = GCHandle.FromIntPtr(pointer)
        Dim list As List(Of IntPtr) = TryCast(gch.Target, List(Of IntPtr))

        If (list Is Nothing) Then
            Throw New InvalidCastException("GCHandle Targer could Not be cast as list")
        End If

        list.Add(handle)

        Return True
    End Function

    Private Function InvokeButtonWithSendkeys(ByVal elementCollectionAll As AutomationElementCollection, ByVal fieldName As String, ByVal command As String) As Boolean

        For Each autoElement As AutomationElement In elementCollectionAll
            If (autoElement.Current.ClassName.Equals("Button")) Then
                If (autoElement.Current.Name = fieldName) Then
                    Dim btnPattern As InvokePattern = TryCast(autoElement.GetCurrentPattern(InvokePattern.Pattern), InvokePattern)
                    autoElement.SetFocus()
                    SendKeys.SendWait(command)
                    Thread.Sleep(300)
                    Return True
                End If
            End If
        Next

        Return False
    End Function

    Private Sub PressNamedButtonWithSendkeys(ByVal windowName As String, ByVal command As String, ByVal buttonName As String)
        Dim timeout As DateTime = DateTime.Now
        timeout = timeout.AddSeconds(180)
        Dim hWindow As IntPtr = GetHWNDWiondow(windowName, timeout)
        Thread.Sleep(500)
        SetForegroundWindow(hWindow)

        Dim windowsList As List(Of IntPtr) = GetChildWindows(hWindow)
        For i As Integer = 0 To windowsList.Count - 1
            Dim saveAsWindow As AutomationElement = AutomationElement.FromHandle(windowsList(i))
            Dim elementCollectionAll As AutomationElementCollection = saveAsWindow.FindAll(TreeScope.Subtree, Condition.TrueCondition)
            'If (InvokeButtonWithSendkeys(elementCollectionAll, buttonName, command)) Then
            If (InvokeButtonOk(elementCollectionAll, buttonName)) Then
                Exit For
            End If
        Next
        Thread.Sleep(300)
    End Sub

    Private Sub TryToFoundWindowAndSetForeground(windowName As String)
        Dim timeout As DateTime = DateTime.Now
        timeout = timeout.AddSeconds(60)
        Dim nullString As String = Nothing

        Dim hWindow As IntPtr = GetHWNDWiondow(nullString, windowName, timeout)

        SetForegroundWindow(hWindow)
    End Sub

    Private Function GetHWNDWiondow(ByRef nameWindow As String, ByRef windowHeader As String, ByVal timeout As DateTime) As IntPtr

        Dim hWindow As IntPtr = New IntPtr()
        Dim isExit As Boolean = False

        While (isExit = False)
            hWindow = FindWindow(nameWindow, windowHeader)
            If (Not IsValidHandle(hWindow)) Then
                If (DateTime.Now > timeout) Then
                    Throw New ArgumentNullException("Cannot found launched window \"" + windowHeader + " \ "")
                End If
            Else
                isExit = True
            End If
        End While

        Return hWindow
    End Function

    Private Sub SaveAsWindow(ByVal fileName As String, ByVal localFolder As String, ByVal windowName As String, ByVal buttonName As String)
        Dim timeout As DateTime = DateTime.Now
        timeout = timeout.AddSeconds(20)
        Dim fullfileName As String = localFolder + "\" + fileName
        Dim hWindow As IntPtr = GetHWNDWiondow(windowName, timeout)
        Thread.Sleep(500)
        SetForegroundWindow(hWindow)

        Dim windowsList As List(Of IntPtr) = GetChildWindows(hWindow)

        For i As Integer = 0 To windowsList.Count - 1
            Dim saveAsWindow As AutomationElement = AutomationElement.FromHandle(windowsList(i))
            Dim elementCollectionAll As AutomationElementCollection = saveAsWindow.FindAll(TreeScope.Subtree, Condition.TrueCondition)
            If (SendFileNameToDialogBox(elementCollectionAll, fullfileName)) Then
                Exit For
            End If
        Next
        Thread.Sleep(500)

        For i As Integer = 0 To windowsList.Count - 1
            Dim saveAsWindow As AutomationElement = AutomationElement.FromHandle(windowsList(i))
            Dim elementCollectionAll As AutomationElementCollection = saveAsWindow.FindAll(TreeScope.Subtree, Condition.TrueCondition)
            If (InvokeButtonOk(elementCollectionAll, buttonName)) Then
                Exit For
            End If
        Next
        Thread.Sleep(500)
    End Sub

    Private Function RemoveUnnecessaryChar(str As String) As String
        'Chr(8)  Backspace character
        'Chr(32) Space
        'Chr(34) Quotation Mark
        'Chr(160)    Non-breaking space
        RemoveUnnecessaryChar = Replace(Replace(Replace(Replace(Replace(Replace(Replace(str, Chr(13), ""), Chr(7), ""), Chr(9), ""), Chr(11), ""), Chr(160), ""), Chr(32), ""), Chr(46), "")

    End Function

    Private Sub StartSap(ByVal login As String, ByVal password As String, ByVal connectionString As String)

        TaskKill("saplogon")

        Dim pidSap As Integer = Shell("C\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", 1)

        Dim timeout As DateTime
        Dim hwnd As Long
        Dim isExit As Boolean

        timeout = DateTime.Now.AddSeconds(10)
        isExit = False
        While (isExit = False)
            hwnd = FindWindow("#32770", "SAP Logon 750")
            If hwnd Then
                isExit = True
            Else
                If (DateTime.Now > timeout) Then
                    Throw New Exception("Не открылось окно 'SAP Logon 750'")
                End If
            End If
        End While

        Dim session = GetObject("SAPGUI").GetScriptingEngine.OpenConnection(connectionString, True).Children(0).Children(0)
        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/usr/txtRSYST-BNAME").Text = login
        session.findById("wnd[0]/usr/pwdRSYST-BCODE").Text = password
        session.findById("wnd[0]/tbar[0]/btn[0]").press
        session = Nothing

    End Sub

    Private Sub TaskKill(ByVal taskName As String)
        For Each oProcess As System.Diagnostics.Process In System.Diagnostics.Process.GetProcessesByName(taskName)
            oProcess.Kill()
        Next
    End Sub

    Private Sub ReleaseObject(comOj As Object)
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

    Private Function CheckSAPMainWindow() As Boolean
        Dim timeout As DateTime = DateTime.Now.AddSeconds(1)
        Dim nullString As String = Nothing

        Try
            Dim hWindow As IntPtr = GetHWNDWiondow(nullString, "SAP Easy Access", timeout)
            SetForegroundWindow(hWindow)
            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    Private Sub CloseUnnecessarySessions()
        Dim sessionCount = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children.Length
        Dim session = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0)
        Dim id As Integer

        If sessionCount > 1 Then
            For id = 1 To sessionCount - 1
                Try
                    session = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0 + id)
                    session.findbyid("wnd[0]").Close
                    session = Nothing
                Catch ex As Exception
                    Throw New Exception("Не могу закрыть соседние сессии.")
                End Try
            Next
        End If

        session = Nothing
    End Sub

    Private Sub ExitSap()

        Dim session = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0)
        session.findById("wnd[0]").maximize
        For i As Integer = 0 To 10
            Dim name As String
            Try
                name = session.findById("wnd[0]/mbar/menu[" & (i + 0) & "]").Name
                If name = "Система" Then
                    session.findById("wnd[0]/mbar/menu[" & (i + 0) & "]/menu[12]").select
                    Exit For
                End If
            Catch ex As Exception
                Exit For
            End Try
        Next

        ' wait window
        Dim timeout As DateTime = DateTime.Now.AddSeconds(2)
        Dim isExit As Boolean
        isExit = False

        While (isExit = False)
            If Not session.findbyid("wnd[1]/usr/btnSPOP-OPTION1", False) Is Nothing Then
                session.findbyid("wnd[1]/usr/btnSPOP-OPTION1").Press
                isExit = True
            Else
                If (DateTime.Now > timeout) Then
                    TaskKill("saplogon")
                    Throw New Exception("Не открылось окно Выхода")
                End If
            End If
        End While
        ' wait window

        Thread.Sleep(300)
        TaskKill("saplogon")
        session = Nothing
    End Sub

    Private Function IsGuiModalWindow(ByVal session As Object, ByVal windowName As String) As Boolean

        Dim timeout As DateTime

        Try
            ' wait window
            timeout = DateTime.Now.AddSeconds(3)
            While (True)
                Try
                    If Not session.findbyid(windowName, False) Is Nothing Then
                        Return True
                    Else
                        If (DateTime.Now > timeout) Then
                            Return False
                        End If
                    End If
                Catch ex As Exception
                    If (DateTime.Now > timeout) Then
                        Return False
                    End If
                End Try
            End While
            ' wait window

        Catch ex As Exception
        End Try

        Throw New Exception("Ошибка в функции IsGuiModalWindow.")

    End Function

    Private Function FoundExcelAndSaveIt(ByVal localFolder As String, ByVal fileName As String) As Boolean
        Dim isExit As Boolean = False
        Dim timeout As DateTime = DateTime.Now.AddSeconds(10)
        Dim excelProcesses As Process()
        Dim misValue As Object = Reflection.Missing.Value
        Dim exceptionMessage As String = ""

        While (isExit = False)
            excelProcesses = Process.GetProcessesByName("EXCEL")
            If excelProcesses.Length = 0 Then
                If (DateTime.Now > timeout) Then
                    Throw New Exception("Не могу найти процесс Excel.")
                End If
            Else
                ReleaseObject(excelProcesses)
                excelProcesses = Nothing
                isExit = True
            End If
        End While

        Dim xlApplication As Excel.Application
        isExit = False
        timeout = DateTime.Now.AddSeconds(30)
        While (isExit = False)
            Try
                Thread.Sleep(500)
                For Each app As Process In Process.GetProcessesByName("EXCEL")
                    Dim ptrWindow As IntPtr = FindWindow(Nothing, app.MainWindowTitle)
                    If ptrWindow <> IntPtr.Zero Then
                        ShowHideWindow(ptrWindow)
                        ''BringWindowToTop(hWnd)
                    End If
                Next
                Console.WriteLine("Before xlApp")
                'xlApp = System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application")
                xlApplication = TryCast(Marshal.GetActiveObject("Excel.Application"), Excel.Application)
                Console.WriteLine("After xlApp")
                If xlApplication Is Nothing Then
                    Continue While
                End If
                For Each xlWorkBook As Workbook In xlApplication.Workbooks
                    xlWorkBook.SaveAs(localFolder & "\" & fileName, 50)
                    xlWorkBook.Close(True)
                    ReleaseObject(xlWorkBook)
                    xlWorkBook = Nothing
                Next
                xlApplication.Quit()
                ReleaseObject(xlApplication)
                isExit = True
            Catch ex As Exception
                TryLaunchIE()
                If (DateTime.Now > timeout) Then
                    Throw New Exception("Не могу найти Excel.Application. " & exceptionMessage)
                End If
            Finally
                xlApplication = Nothing
            End Try
        End While

        Dim proc As Process
        For Each proc In Process.GetProcessesByName("EXCEL")
            proc.Kill()
        Next

        Return True

    End Function

    Private Sub ShowHideWindow(hWindow As IntPtr)
        Dim autoElement As AutomationElement = AutomationElement.FromHandle(hWindow)
        Dim elementCollectionAll As AutomationElementCollection = autoElement.FindAll(TreeScope.Subtree, Condition.TrueCondition)
        SetFocusOnWindow(elementCollectionAll)
        Dim ptrWindow As Integer = CType(hWindow, Integer)
        ShowWindow(ptrWindow, 0)
        Thread.Sleep(300)
        ShowWindow(ptrWindow, 9)
        Thread.Sleep(300)
        ShowWindow(ptrWindow, 3)
        Thread.Sleep(300)
        SendMessageW(ptrWindow, BM_CLICK, IntPtr.Zero, IntPtr.Zero)
        Thread.Sleep(300)
    End Sub

    Private Sub TryLaunchIE()
        Dim ie As Process = Process.Start("iexplore.exe", "localhost")
        'close the website
        Thread.Sleep(2000)
        Try
            Dim ieMainWindow As AutomationElement = AutomationElement.FromHandle(ie.MainWindowHandle)
            Dim elementCollectionAll As AutomationElementCollection = ieMainWindow.FindAll(TreeScope.Subtree, Condition.TrueCondition)
            SetFocusOnWindow(elementCollectionAll)

            Thread.Sleep(200)
            Dim ieProc As Process
            For Each ieProc In Process.GetProcessesByName("iexplore")
                ieProc.Kill()
            Next
        Catch ex As Exception
        End Try
    End Sub

    Private Function SetFocusOnWindow(elementCollectionAll As AutomationElementCollection) As Boolean

        For Each autoElement As AutomationElement In elementCollectionAll
            autoElement.SetFocus()
            Return True
        Next

        Return False
    End Function

    Private Sub TakeScreenShot(ByVal folder As String, ByVal fileName As String)

        Dim screenSize As Size = New Size(My.Computer.Screen.Bounds.Width, My.Computer.Screen.Bounds.Height)
        Dim screenGrab As New Bitmap(My.Computer.Screen.Bounds.Width, My.Computer.Screen.Bounds.Height)

        Dim graphic As Graphics = Graphics.FromImage(screenGrab)

        graphic.CopyFromScreen(New System.Drawing.Point(0, 0), New System.Drawing.Point(0, 0), screenSize)
        screenGrab.Save(folder & "\" & fileName & ".png")

    End Sub

End Module
