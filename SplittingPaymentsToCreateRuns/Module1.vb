Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Xml.Serialization
Imports Microsoft.Office.Interop
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Data
Imports System.Data.OleDb
Imports System.Diagnostics

Module Module1

    Sub Main()
        ' *************** input variables
        Dim localFolder As String = "C:\Temp\WorkDir"
        Dim xmlTableName As String = "Table_PlanRU17.xml"
        Dim excelFile As String = "Table_PlanRU17.xlsb"
        Dim excelShablonFileName As String = "ShablonRu.xlsb"
        Dim isCreateExcelFile As Boolean = True
        Dim nameRun As String = ""
        Dim be As String = "Ru17"
        ' *************** input variables

        ' *************** output variables
        Dim exceptionMessage As String
        Dim isComplete As Boolean
        Dim isRowPresent As Boolean
        Dim rowCount As Integer
        Dim tablesToRun As DataTable
        ' *************** output variables

        ' *************** Begin
        Dim outputTable As DataTable = New DataTable()
        exceptionMessage = ""
        Try
            isComplete = False
            be = be.ToUpper()

            If isCreateExcelFile Then
                File.Copy(localFolder & "\" & excelShablonFileName, localFolder & "\" & excelFile, True)
            End If

            Dim mainTable As DataTable = GetDataTableFromFile(localFolder & "\" & xmlTableName)
            Dim sheetName As String
            Dim filter As String
            Dim view As DataView

            Dim numberOfRows As Integer = 0
            Dim part As Integer = 0
            Dim maxRowsInOnePart As Integer = 0
            Dim partsCollection As DataTable = New DataTable()
            Dim cutedTable As DataTable = New DataTable()
            Dim tempTable As DataTable = New DataTable()

            outputTable = ResetOutputTable()

            If (mainTable IsNot Nothing) Then
                If (mainTable.Rows.Count > 0) Then
                    mainTable.CaseSensitive = False

                    '•Cити банк\Дойче банк – содержит: «Epap» (на английском) (фильтр по столбцу С(F3)) банк определяется способом платежа (СП) 
                    ' (столбец U(F21)- для СИТИ- СП- q, для Дойче – B\L)
                    sheetName = "Citibank - EPAP "
                    filter = "[F3] Like '%epap%' And [F21] = 'q'"
                    If isCreateExcelFile Then
                        CreateSheetEPAP(localFolder, excelFile, exceptionMessage, mainTable, filter, sheetName)
                    ElseIf nameRun = "Citibank - EPAP" Then
                        outputTable = GetTableEPAP(exceptionMessage, mainTable, filter, sheetName, outputTable)
                        isComplete = True
                        Exit Sub
                    End If

                    sheetName = "DB - EPAP "
                    filter = "[F3] Like '%epap%' And ([F21] = 'b' Or [F21] = 'l')"
                    If isCreateExcelFile Then
                        CreateSheetEPAP(localFolder, excelFile, exceptionMessage, mainTable, filter, sheetName)
                    ElseIf nameRun = "DB - EPAP" Then
                        outputTable = GetTableEPAP(exceptionMessage, mainTable, filter, sheetName, outputTable)
                        isComplete = True
                        Exit Sub
                    End If

                    ' убираем из таблицы все %epap%
                    filter = "[F3] not like '%epap%' or [F3] is null"
                    mainTable = DeleteByFilterFromTable(mainTable, filter)
                    'view = New DataView(mainTable)
                    'view.RowFilter = filter
                    'mainTable = view.ToTable()
                    ' End убираем из таблицы все %epap%

                    '• DeutscheBank Main
                    Dim tempTableBeforeMerge As DataTable = New DataTable()
                    sheetName = "DB - USD, EUR "
                    filter = "[F21] = 'P' And ([F10] = 'USD' Or [F10] = 'EUR')"
                    view = New DataView(mainTable)
                    view.RowFilter = filter
                    Dim tableDeutscheBank As DataTable = view.ToTable()
                    'убираем из DeutscheBank 3000486 (это отдельный лист)
                    filter = "[F4] <> '3000486'"
                    view = New DataView(tableDeutscheBank)
                    view.RowFilter = filter
                    tableDeutscheBank = view.ToTable()
                    'End убираем из DeutscheBank 3000486 (это отдельный лист)
                    ' далее условие по 58 контракту отдельный прогон
                    If tableDeutscheBank.Rows.Count > 0 Then
                        view = New DataView(tableDeutscheBank)
                        filter = "[F4] = '58'"
                        view.RowFilter = filter
                        Dim table58Contract As DataTable = view.ToTable() ' получаем все 58 контракты
                        If table58Contract.Rows.Count > 0 Then
                            'делим по "Ссылка на платеж"
                            view = New DataView(table58Contract)
                            Dim paymentsReferencesTable As DataTable = view.ToTable(True, "F22") 'это таблица ссылок на платеж
                            paymentsReferencesTable = RemoveNullValue(paymentsReferencesTable, "F22")

                            For Each paymentReference As DataRow In paymentsReferencesTable.Rows
                                ' накладываем на неё фильтр по ссылке на платеж
                                view = New DataView(table58Contract)
                                filter = "[F22] = '" & paymentReference("F22") & "'"
                                view.RowFilter = filter
                                If isCreateExcelFile Then
                                    PrepareToSaveTableToExcel(view.ToTable(), localFolder, excelFile, exceptionMessage, sheetName, 450, part)
                                Else
                                    outputTable = PrepareTable(view.ToTable(), exceptionMessage, sheetName, 450, part, outputTable)
                                    tempTableBeforeMerge = MergeTwoTables(tempTableBeforeMerge, outputTable)
                                    outputTable = ResetOutputTable()
                                End If
                            Next
                        End If

                        ' далее условие по 363392 контракту отдельный прогон
                        filter = "[F4] = '363392'"
                        view = New DataView(tableDeutscheBank)
                        view.RowFilter = filter
                        Dim table363392Contract As DataTable = view.ToTable() ' tempTable по 363392 фильтр
                        If table363392Contract.Rows.Count > 0 Then
                            'делим по ссылочному ключу
                            'получаем все ссылочные ключи
                            view = New DataView(table363392Contract)
                            Dim referencesKeysTable As DataTable = view.ToTable(True, "F29")
                            'referencesKeysTable = RemoveNullValue(referencesKeysTable, "F29")

                            For Each referenceKey As DataRow In referencesKeysTable.Rows
                                view = New DataView(table363392Contract)
                                filter = "[F29] = '" & referenceKey("F29") & "'"
                                view.RowFilter = filter
                                Dim referenceKeyTable As DataTable = view.ToTable() ' 'это таблица по одному ссылочному ключу

                                'делим по "Ссылка на платеж"
                                view = New DataView(referenceKeyTable)
                                Dim paymentsReferencesTable As DataTable = view.ToTable(True, "F22") 'это таблица ссылок на платеж
                                'paymentsReferencesTable = RemoveNullValue(paymentsReferencesTable, "F22")

                                For Each paymentReference As DataRow In paymentsReferencesTable.Rows
                                    ' tempTable накладываем на неё фильтр по ссылочному ключу и по ссылке на платеж
                                    view = New DataView(table363392Contract)
                                    filter = "[F22] = '" & paymentReference("F22") & "' And " & "[F29] = '" & referenceKey("F29") & "'"
                                    view.RowFilter = filter
                                    If isCreateExcelFile Then
                                        PrepareToSaveTableToExcel(view.ToTable(), localFolder, excelFile, exceptionMessage, sheetName, 800, part)
                                    Else
                                        outputTable = PrepareTable(view.ToTable(), exceptionMessage, sheetName, 800, part, outputTable)
                                        tempTableBeforeMerge = MergeTwoTables(tempTableBeforeMerge, outputTable)
                                        outputTable = ResetOutputTable()
                                    End If
                                Next
                            Next
                        End If

                        ' далее условие по остальным контрактам отдельный прогон
                        filter = "[F4] <> '363392' And [F4] <> '58'"
                        If isCreateExcelFile Then
                            CreateTableFromFilter(localFolder, excelFile, exceptionMessage, tableDeutscheBank, sheetName, filter, part, 450)
                        Else
                            outputTable = CreateTableFromFilter(localFolder, exceptionMessage, tableDeutscheBank, sheetName, filter, part, 450, outputTable)
                            tempTableBeforeMerge = MergeTwoTables(tempTableBeforeMerge, outputTable)
                            outputTable = ResetOutputTable()
                        End If
                    End If
                    If nameRun = "DB - USD, EUR" Then
                        outputTable = tempTableBeforeMerge
                        isComplete = True
                        Exit Sub
                    End If
                    outputTable = ResetOutputTable()
                    tempTableBeforeMerge = New DataTable()
                    'End DeutscheBank Main

                    ' DeutscheBank 3000486
                    part = 0
                    sheetName = "DB - 3000486 "
                    filter = "[F4] = '3000486'"
                    view = New DataView(mainTable)
                    view.RowFilter = filter
                    view.Sort = "[F4]"
                    Dim tableDeutscheBank3000486 As DataTable = view.ToTable()
                    If tableDeutscheBank3000486.Rows.Count > 0 Then
                        ' отдельные прогоны по каждому типу Банка(столбец АА(F27))
                        view = New DataView(tableDeutscheBank3000486)
                        Dim tableTypeOfPartners As DataTable = view.ToTable(True, "F27") 'это таблица Тип банка-партнера
                        'tableTypeOfPartners = RemoveNullValue(tableTypeOfPartners, "F27")

                        For Each typeOfPartner As DataRow In tableTypeOfPartners.Rows
                            view = New DataView(tableDeutscheBank3000486)
                            filter = "[F27] = '" & typeOfPartner("F27") & "'"
                            view.RowFilter = filter
                            If isCreateExcelFile Then
                                PrepareToSaveTableToExcel(view.ToTable(), localFolder, excelFile, exceptionMessage, sheetName, 450, part)
                            Else
                                outputTable = PrepareTable(view.ToTable(), exceptionMessage, sheetName, 450, part, outputTable)
                            End If
                        Next
                    End If
                    If nameRun = "DB - 3000486" Then
                        isComplete = True
                        Exit Sub
                    End If
                    outputTable = ResetOutputTable()
                    ' End DeutscheBank 3000486
                    filter = "Not ([F21] = 'P' And ([F10] = 'USD' Or [F10] = 'EUR'))"
                    mainTable = DeleteByFilterFromTable(mainTable, filter)
                    filter = "Not ([F4] = '3000486')"
                    mainTable = DeleteByFilterFromTable(mainTable, filter)

                    '• Дойче банк – GBP (столбец J(F10)) – способ платежа O (столбец U(F21))
                    part = 0
                    sheetName = "DB - GBP "
                    filter = "[F10] = 'GBP' And [F21] = 'O'"
                    If isCreateExcelFile Then
                        CreateTableFromFilter(localFolder, excelFile, exceptionMessage, mainTable, sheetName, filter, part, 450)
                    ElseIf nameRun = "DB - GBP" Then
                        outputTable = CreateTableFromFilter(localFolder, exceptionMessage, mainTable, sheetName, filter, part, 450, outputTable)
                        isComplete = True
                        Exit Sub
                    End If
                    outputTable = ResetOutputTable()
                    'End • Дойче банк – GBP (столбец J(F10)) – способ платежа O (столбец U(F21))
                    filter = "Not ([F10] = 'GBP' And [F21] = 'O')"
                    mainTable = DeleteByFilterFromTable(mainTable, filter)

                    '!!! этот пункт уехал в Дойче банк - RUB
                    '''• Дойче банк - RUB – нерезиденты (столбец BF(F58)) не равно пусто (убрать пусто и #Н/Д)– способ платежа B(столбец U(F21)) 
                    ''   и выбираем валюту RUB столбец J(F10)
                    'part = 0
                    'sheetName = "DB - RUB (non-residents) "
                    'filter = "[F58] Is NOT NULL And [F21] = 'B' And [F10] = 'RUB'"
                    'If isCreateExcelFile Then
                    '    CreateTableFromFilter(localFolder, excelFile, exceptionMessage, mainTable, sheetName, filter, part, 450)
                    'ElseIf nameRun = "DB - RUB (non-residents)" Then
                    '    outputTable = CreateTableFromFilter(localFolder, exceptionMessage, mainTable, sheetName, filter, part, 450, outputTable)
                    '    isComplete = True
                    '    Exit Sub
                    'End If
                    'outputTable = ResetOutputTable()
                    '' End Дойче банк - RUB – нерезиденты (столбец BF) – способ платежа B(столбец U)

                    '*********** Citibank часть меняем на DB
                    '*********** Citibank 
                    ' Дойче банк – по кредитору 358053 (RUB) (столбец D(F4)) по каждому контракту (по ссылке на платеж столбец V “Ссылка на платеж”(F22), 
                    ' если не указано, то по столбцу № договора (столбец Z(F26))) отдельный прогон (особенность формирования пп) 
                    ' nameRun07
                    part = 0
                    sheetName = "DB - 358053 "
                    filter = "[F10] = 'RUB' And [F4] = '358053'"
                    If isCreateExcelFile Then
                        CreateDBTableByLender(localFolder, excelFile, exceptionMessage, mainTable, sheetName, filter, part)
                    ElseIf nameRun = "DB - 358053" Then
                        outputTable = CreateDBTableByLender(localFolder, exceptionMessage, mainTable, sheetName, filter, part, outputTable)
                        isComplete = True
                        Exit Sub
                    End If
                    outputTable = ResetOutputTable()
                    ' End "DB - 358053"
                    filter = "Not ([F10] = 'RUB' And [F4] = '358053')"
                    mainTable = DeleteByFilterFromTable(mainTable, filter)

                    'Стал DB
                    'Сити банк (DB) – по кредитору 311505 (RUB) (столбец D(F4)) по каждому контракту (по ссылке на платеж столбец V “Ссылка на платеж”(F22), 
                    ' если не указано, то по столбцу № договора (столбец Z(F26))) отдельный прогон (особенность формирования пп)
                    ' nameRun08
                    part = 0
                    sheetName = "DB - 311505 "
                    filter = "[F10] = 'RUB' And [F4] = '311505'"
                    If isCreateExcelFile Then
                        CreateDBTableByLender(localFolder, excelFile, exceptionMessage, mainTable, sheetName, filter, part)
                    ElseIf nameRun = "DB - 311505" Then
                        outputTable = CreateDBTableByLender(localFolder, exceptionMessage, mainTable, sheetName, filter, part, outputTable)
                        isComplete = True
                        Exit Sub
                    End If
                    outputTable = ResetOutputTable()
                    ' End "DB - 311505"
                    filter = "Not ([F10] = 'RUB' And [F4] = '311505')"
                    mainTable = DeleteByFilterFromTable(mainTable, filter)

                    '• Сити банк (DB) – по кредитору 3011479 по каждому контракту (по ссылке на платеж-(столбец V),
                    ' если не указано, то по столбцу № договора(столбец Z) отдельный прогон 
                    ' (особенность формирования пп)- RUB (столбец J) 
                    ' сразу разбиваем для прогона, RUD, RUE(столбец J) переносим на отдельный лист
                    ' – прогон будет только когда есть курс валют на заданную дату
                    ' nameRun09
                    part = 0
                    sheetName = "DB - 3011479 RUB "
                    filter = "[F10] = 'RUB' And [F4] = '3011479'"
                    If isCreateExcelFile Then
                        CreateDBTableByLender(localFolder, excelFile, exceptionMessage, mainTable, sheetName, filter, part)
                    ElseIf nameRun = "DB - 3011479 RUB" Then
                        outputTable = CreateDBTableByLender(localFolder, exceptionMessage, mainTable, sheetName, filter, part, outputTable)
                        isComplete = True
                        Exit Sub
                    End If
                    outputTable = ResetOutputTable()
                    ' End "DB - 3011479 RUB"
                    filter = "Not ([F10] = 'RUB' And [F4] = '3011479')"
                    mainTable = DeleteByFilterFromTable(mainTable, filter)

                    part = 0
                    ' nameRun10
                    sheetName = "DB - 3011479 RUD, RUE "
                    filter = "([F10] = 'RUD' Or [F10] = 'RUE') And [F4] = '3011479'"
                    If isCreateExcelFile Then
                        CreateDBTableByLender(localFolder, excelFile, exceptionMessage, mainTable, sheetName, filter, part)
                    ElseIf nameRun = "DB - 3011479 RUD, RUE" Then
                        outputTable = CreateDBTableByLender(localFolder, exceptionMessage, mainTable, sheetName, filter, part, outputTable)
                        isComplete = True
                        Exit Sub
                    End If
                    outputTable = ResetOutputTable()
                    ' End "DB - 3011479 RUD, RUE"
                    filter = "Not (([F10] = 'RUD' Or [F10] = 'RUE') And [F4] = '3011479')"
                    mainTable = DeleteByFilterFromTable(mainTable, filter)

                    '• DB - TAP RUB B (столбец D(F4)) (по столбцу № документа (столбец G(F7)) фильтр по 17* проводкам)
                    '• Citibank - TAP RUB Q  (по столбцу № документа (столбец G(F7)) фильтр по 17* проводкам)
                    ' И разделить на разные листы по способу платежа(столбец U(F21)) (по j и q)
                    ' nameRun11
                    ' !!! Если это Ru17
                    ' (nameRun111) Сити банк - ТАП - RUB (по столбцу № документа фильтр по 17* проводкам) способ платежа J (столбец U) 
                    filter = "[F10] = 'RUB' And CONVERT([F7], System.String) Like '17%'"
                    view = New DataView(mainTable)
                    view.RowFilter = filter
                    Dim tableDBTAPRUB As DataTable = view.ToTable()
                    If tableDBTAPRUB.Rows.Count > 0 Then
                        ' И разделить на разные листы по способу платежа(столбец U) (по b и q)
                        part = 0
                        sheetName = "DB - TAP RUB B "
                        filter = "[F21] = 'B'"
                        If isCreateExcelFile Then
                            CreateTableFromFilter(localFolder, excelFile, exceptionMessage, tableDBTAPRUB, sheetName, filter, part, 450)
                        ElseIf nameRun = "DB - TAP RUB B" Then
                            isComplete = True
                            outputTable = CreateTableFromFilter(localFolder, exceptionMessage, tableDBTAPRUB, sheetName, filter, part, 450, outputTable)
                            Exit Sub
                        End If
                        outputTable = ResetOutputTable()
                        ' End "DB - TAP RUB B"

                        'nameRun12
                        part = 0
                        sheetName = "Citibank - TAP RUB Q "
                        filter = "[F21] = 'Q'"
                        If isCreateExcelFile Then
                            CreateTableFromFilter(localFolder, excelFile, exceptionMessage, tableDBTAPRUB, sheetName, filter, part, 450)
                        ElseIf nameRun = "Citibank - TAP RUB Q" Then
                            outputTable = CreateTableFromFilter(localFolder, exceptionMessage, tableDBTAPRUB, sheetName, filter, part, 450, outputTable)
                            isComplete = True
                            Exit Sub
                        End If
                        outputTable = ResetOutputTable()
                        ' End "Citibank - TAP RUB Q"

                        If be = "RU17" Then
                            'nameRun111
                            part = 0
                            sheetName = "Citibank - TAP RUB J "
                            filter = "[F21] = 'J'"
                            If isCreateExcelFile Then
                                CreateTableFromFilter(localFolder, excelFile, exceptionMessage, tableDBTAPRUB, sheetName, filter, part, 450)
                            ElseIf nameRun = "Citibank - TAP RUB J" Then
                                outputTable = CreateTableFromFilter(localFolder, exceptionMessage, tableDBTAPRUB, sheetName, filter, part, 450, outputTable)
                                isComplete = True
                                Exit Sub
                            End If
                            outputTable = ResetOutputTable()
                            ' End Ru17 "Citibank - TAP RUB J"
                        End If
                    End If
                    ' End "DB - TAP RUB B" Or End "Citibank - TAP RUB Q"
                    filter = "Not ([F10] = 'RUB' And CONVERT([F7], System.String) Like '17%')"
                    mainTable = DeleteByFilterFromTable(mainTable, filter)

                    'Далее то что осталось из вышевыбранного
                    '• DB - RUB – не более 450 строк – способ платежа B (столбец U(F21))– для обычных рублевых платежей
                    ' nameRun13
                    sheetName = "DB - RUB PM B "
                    filter = "[F10] = 'RUB' And [F4] <> '358053' And [F4] <> '311505' And [F4] <> '3011479' And [F21] = 'B' And CONVERT([F7], System.String) Not Like '17%' And [F7] Is NOT NULL"
                    If isCreateExcelFile Then
                        CreateTableFromFilter(localFolder, excelFile, exceptionMessage, mainTable, sheetName, filter, part, 450)
                    ElseIf nameRun = "DB - RUB PM B" Then
                        outputTable = CreateTableFromFilter(localFolder, exceptionMessage, mainTable, sheetName, filter, part, 450, outputTable)
                        isComplete = True
                        Exit Sub
                    End If
                    outputTable = ResetOutputTable()
                    ' End "DB - RUB PM B"
                    filter = "Not ([F10] = 'RUB' And [F4] <> '358053' And [F4] <> '311505' And [F4] <> '3011479' And [F21] = 'B' And CONVERT([F7], System.String) Not Like '17%' And [F7] Is NOT NULL)"
                    mainTable = DeleteByFilterFromTable(mainTable, filter)

                    '• Citibank - RUB – не более 800 строк – способ платежа Q (столбец U)– для налоговых платежей (платежи УФК с доп. полями)
                    ' nameRun14
                    sheetName = "Citibank - RUB PM Q "
                    filter = "[F10] = 'RUB' And [F4] <> '358053' And [F4] <> '311505' And [F4] <> '3011479' And [F21] = 'Q' And CONVERT([F7], System.String) Not Like '17%' And [F7] Is NOT NULL"
                    If isCreateExcelFile Then
                        CreateTableFromFilter(localFolder, excelFile, exceptionMessage, mainTable, sheetName, filter, part, 800)
                    ElseIf nameRun = "Citibank - RUB PM Q" Then
                        outputTable = CreateTableFromFilter(localFolder, exceptionMessage, mainTable, sheetName, filter, part, 800, outputTable)
                        isComplete = True
                        Exit Sub
                    End If
                    outputTable = ResetOutputTable()
                    ' End "Citibank - RUB PM Q"
                    filter = "Not ([F10] = 'RUB' And [F4] <> '358053' And [F4] <> '311505' And [F4] <> '3011479' And [F21] = 'Q' And CONVERT([F7], System.String) Not Like '17%' And [F7] Is NOT NULL)"
                    mainTable = DeleteByFilterFromTable(mainTable, filter)

                    '!! Ru17 (nameRun151) будет то же СитиБанк и способ платежа J
                    If be = "RU17" Then
                        sheetName = "Citibank - RUB PM J "
                        filter = "[F10] = 'RUB' And [F4] <> '358053' And [F4] <> '311505' And [F4] <> '3011479' And [F21] = 'J' And CONVERT([F7], System.String) Not Like '17%' And [F7] Is NOT NULL"
                        If isCreateExcelFile Then
                            CreateTableFromFilter(localFolder, excelFile, exceptionMessage, mainTable, sheetName, filter, part, 800)
                        ElseIf nameRun = "Citibank - RUB PM J" Then
                            outputTable = CreateTableFromFilter(localFolder, exceptionMessage, mainTable, sheetName, filter, part, 800, outputTable)
                            isComplete = True
                            Exit Sub
                        End If
                        outputTable = ResetOutputTable()
                        ' End "Citibank - RUB PM J"
                        filter = "Not ([F10] = 'RUB' And [F4] <> '358053' And [F4] <> '311505' And [F4] <> '3011479' And [F21] = 'J' And CONVERT([F7], System.String) Not Like '17%' And [F7] Is NOT NULL)"
                        mainTable = DeleteByFilterFromTable(mainTable, filter)
                    End If

                    '• DB - ТАП (RUD, RUE)  (по столбцу № документа (столбец G(F7)) фильтр по 17* проводкам) 
                    'И разделить на разные листы по способу платежа(столбец U) (по j и q)
                    filter = "([F10] = 'RUD' Or [F10] = 'RUE') And CONVERT([F7], System.String) Like '17%'"
                    view = New DataView(mainTable)
                    view.RowFilter = filter
                    Dim tableCitibankTAPRUDandRUE As DataTable = view.ToTable()
                    If tableCitibankTAPRUDandRUE.Rows.Count > 0 Then
                        ' И разделить на разные листы по способу платежа(столбец U) (по j и q)
                        ' nameRun15
                        part = 0
                        sheetName = "DB - TAP RUDandRUE J "
                        filter = "[F21] = 'J'"
                        If isCreateExcelFile Then
                            CreateTableFromFilter(localFolder, excelFile, exceptionMessage, tableCitibankTAPRUDandRUE, sheetName, filter, part, 450)
                        ElseIf nameRun = "DB - TAP RUDandRUE J" Then
                            outputTable = CreateTableFromFilter(localFolder, exceptionMessage, tableCitibankTAPRUDandRUE, sheetName, filter, part, 450, outputTable)
                            isComplete = True
                            Exit Sub
                        End If
                        outputTable = ResetOutputTable()
                        ' End "DB - TAP RUDandRUE J"

                        ' nameRun16
                        part = 0
                        sheetName = "DB - TAP RUDandRUE Q "
                        filter = "[F21] = 'Q'"
                        If isCreateExcelFile Then
                            CreateTableFromFilter(localFolder, excelFile, exceptionMessage, tableCitibankTAPRUDandRUE, sheetName, filter, part, 450)
                        ElseIf nameRun = "DB - TAP RUDandRUE Q" Then
                            outputTable = CreateTableFromFilter(localFolder, exceptionMessage, tableCitibankTAPRUDandRUE, sheetName, filter, part, 450, outputTable)
                            isComplete = True
                            Exit Sub
                        End If
                        outputTable = ResetOutputTable()
                        ' End "DB - TAP RUDandRUE Q"
                    End If
                    filter = "Not (([F10] = 'RUD' Or [F10] = 'RUE') And CONVERT([F7], System.String) Like '17%')"
                    mainTable = DeleteByFilterFromTable(mainTable, filter)

                    '• Сити банк - RUD, RUE (прогоняем только когда есть курс на платежный день, стандартно, после 14-00 накануне ПД)
                    ' nameRun17
                    part = 0
                    If be = "RU17" Then
                        sheetName = "Citibank - RUD, RUE "
                    Else
                        sheetName = "DB - RUD, RUE "
                    End If
                    filter = "([F10] = 'RUD' Or [F10] = 'RUE') And [F4] <> '358053' And [F4] <> '311505' And [F4] <> '3011479' And ([F7] Is NULL Or CONVERT([F7], System.String) NOT Like '17%')"
                    If isCreateExcelFile Then
                        CreateTableFromFilter(localFolder, excelFile, exceptionMessage, mainTable, sheetName, filter, part, 450)
                    ElseIf nameRun = "DB - RUD, RUE" Then
                        outputTable = CreateTableFromFilter(localFolder, exceptionMessage, mainTable, sheetName, filter, part, 450, outputTable)
                        isComplete = True
                        Exit Sub
                    End If
                    outputTable = ResetOutputTable()
                    ' End "DB - RUD, RUE"
                    filter = "Not (([F10] = 'RUD' Or [F10] = 'RUE') And [F4] <> '358053' And [F4] <> '311505' And [F4] <> '3011479' And ([F7] Is NULL Or CONVERT([F7], System.String) NOT Like '17%'))"
                    mainTable = DeleteByFilterFromTable(mainTable, filter)

                    ' nameRun18
                    part = 0
                    sheetName = "Citibank - 313479 "
                    filter = "[F4] = '313479'"
                    If isCreateExcelFile Then
                        CreateTableFromFilter(localFolder, excelFile, exceptionMessage, mainTable, sheetName, filter, part, 450)
                    ElseIf nameRun = "Citibank - 313479" Then
                        outputTable = CreateTableFromFilter(localFolder, exceptionMessage, mainTable, sheetName, filter, part, 450, outputTable)
                        isComplete = True
                        Exit Sub
                    End If
                    outputTable = ResetOutputTable()
                    ' End "Citibank - 313479"
                    filter = "Not ([F4] = '313479')"
                    mainTable = DeleteByFilterFromTable(mainTable, filter)

                    ' nameRun19
                    part = 0
                    sheetName = "Citibank - 313652 "
                    filter = "[F4] = '313652'"
                    If isCreateExcelFile Then
                        CreateTableFromFilter(localFolder, excelFile, exceptionMessage, mainTable, sheetName, filter, part, 450)
                    ElseIf nameRun = "Citibank - 313652" Then
                        outputTable = CreateTableFromFilter(localFolder, exceptionMessage, mainTable, sheetName, filter, part, 450, outputTable)
                        isComplete = True
                        Exit Sub
                    End If
                    outputTable = ResetOutputTable()
                    ' End "Citibank - 313652"
                    filter = "Not ([F4] = '313652')"
                    mainTable = DeleteByFilterFromTable(mainTable, filter)

                    '• Сити банк – Дебиторы –способ платежа I (фильтр по столбцу U), RUB
                    part = 0
                    sheetName = "Citibank - RUB debtors I "
                    filter = "[F10] = 'RUB' And [F21] = 'I' And [F4] <> '358053' And [F4] <> '311505' And [F4] <> '3011479' And ([F7] Is NULL Or CONVERT([F7], System.String) NOT Like '17%')"
                    If isCreateExcelFile Then
                        CreateTableFromFilter(localFolder, excelFile, exceptionMessage, mainTable, sheetName, filter, part, 450)
                    ElseIf nameRun = "Citibank - RUB debtors I" Then
                        outputTable = CreateTableFromFilter(localFolder, exceptionMessage, mainTable, sheetName, filter, part, 450, outputTable)
                        isComplete = True
                        Exit Sub
                    End If
                    outputTable = ResetOutputTable()
                    ' End "Citibank - RUB debtors I"
                    filter = "Not ([F10] = 'RUB' And [F21] = 'I' And [F4] <> '358053' And [F4] <> '311505' And [F4] <> '3011479' And ([F7] Is NULL Or CONVERT([F7], System.String) NOT Like '17%'))"
                    mainTable = DeleteByFilterFromTable(mainTable, filter)

                    '• Сити банк - Дебиторы – способ платежа Y (фильтр по столбцу U), EUR
                    part = 0
                    sheetName = "Citibank - EUR debtors Y "
                    filter = "[F10] = 'EUR'  And [F21] = 'Y' And [F4] <> '358053' And [F4] <> '311505' And [F4] <> '3011479' And ([F7] Is NULL Or CONVERT([F7], System.String) NOT Like '17%')"
                    If isCreateExcelFile Then
                        CreateTableFromFilter(localFolder, excelFile, exceptionMessage, mainTable, sheetName, filter, part, 450)
                    ElseIf nameRun = "Citibank - EUR debtors Y" Then
                        outputTable = CreateTableFromFilter(localFolder, exceptionMessage, mainTable, sheetName, filter, part, 450, outputTable)
                        isComplete = True
                        Exit Sub
                    End If
                    outputTable = ResetOutputTable()
                    ' End "Citibank - EUR debtors Y"
                    filter = "Not ([F10] = 'EUR'  And [F21] = 'Y' And [F4] <> '358053' And [F4] <> '311505' And [F4] <> '3011479' And ([F7] Is NULL Or CONVERT([F7], System.String) NOT Like '17%'))"
                    mainTable = DeleteByFilterFromTable(mainTable, filter)

                    '• Дойче банк - Дебиторы – способ платежа L (фильтр по столбцу U), RUB – нерезиденты, по критерию Ссылочный ключ 2 (столбец AС)
                    part = 0
                    sheetName = "DB - RUB debtors L "
                    filter = "[F10] = 'RUB'  And [F21] = 'L'"
                    view = New DataView(mainTable)
                    view.RowFilter = filter
                    Dim tableDBdebtorsL As DataTable = view.ToTable()
                    If tableDBdebtorsL.Rows.Count > 0 Then
                        ' по критерию Ссылочный ключ 2 (столбец AС)
                        view = New DataView(tableDBdebtorsL)
                        Dim tableReferencesKeys2 As DataTable = view.ToTable(True, "F29") 'Ссылочный ключ 2 (столбец AС)
                        'tableReferencesKeys2 = RemoveNullValue(tableReferencesKeys2, "F29")

                        For Each referenceKey As DataRow In tableReferencesKeys2.Rows
                            view = New DataView(tableDBdebtorsL)
                            filter = "[F29] = '" & referenceKey("F29") & "'"
                            view.RowFilter = filter
                            If isCreateExcelFile Then
                                PrepareToSaveTableToExcel(view.ToTable(), localFolder, excelFile, exceptionMessage, sheetName, 450, part)
                            ElseIf nameRun = "DB - RUB debtors L" Then
                                outputTable = PrepareTable(view.ToTable(), exceptionMessage, sheetName, 450, part, outputTable)
                                tempTableBeforeMerge = MergeTwoTables(tempTableBeforeMerge, outputTable)
                                outputTable = ResetOutputTable()
                            End If
                        Next
                    End If
                    If nameRun = "DB - RUB debtors L" Then
                        outputTable = tempTableBeforeMerge
                        isComplete = True
                        Exit Sub
                    End If
                    outputTable = ResetOutputTable()
                    tempTableBeforeMerge = New DataTable()
                    ' End "DB - RUB debtors L"
                    filter = "Not ([F10] = 'RUB'  And [F21] = 'L')"
                    mainTable = DeleteByFilterFromTable(mainTable, filter)

                    '*********** End Citibank 
                End If
            End If

            isComplete = True

        Catch ex As Exception
            exceptionMessage = ex.Message
            If (Not ex.InnerException Is Nothing) Then
                exceptionMessage = exceptionMessage & " Inner Exception : " & ex.InnerException.Message
            End If
        Finally
            tablesToRun = outputTable
            If (tablesToRun IsNot Nothing) Then
                If (tablesToRun.Rows.Count > 0) Then
                    isRowPresent = True
                    rowCount = tablesToRun.Rows.Count
                End If
            End If
            Console.WriteLine("ExceptionMessage : {0}", exceptionMessage)
            Console.WriteLine("Первичный поток: Id {0} is Ended", Thread.CurrentThread.ManagedThreadId)
            Console.ReadKey()
        End Try
        ' *************** End

    End Sub

    Private Function DeleteByFilterFromTable(mainTable As DataTable, filter As String) As DataTable
        filter = filter
        Dim view As DataView = New DataView(mainTable)
        view.RowFilter = filter
        Return view.ToTable()
    End Function

    Private Function RemoveNullValue(ByVal table As DataTable, ByVal columnName As String) As DataTable
        For i As Integer = table.Rows.Count - 1 To 0 Step -1
            If DBNull.Value.Equals(table.Rows(i)(columnName)) Then
                table.Rows.Remove(table.Rows(i))
            End If
        Next

        Return table
    End Function

    Private Function MergeTwoTables(ByVal collection1 As DataTable, ByVal collection2 As DataTable) As DataTable

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

    Private Function ResetOutputTable() As DataTable
        Dim outputTable As DataTable = New DataTable("OutTable")
        outputTable.Columns.Add("SheetName", Type.GetType("System.String"))
        outputTable.Columns.Add(New System.Data.DataColumn("TBL", outputTable.GetType()))
        Return outputTable
    End Function

    Private Function GetTableEPAP(ByRef exceptionMessage As String, ByVal mainTable As DataTable, ByVal filter As String, ByVal sheetName As String, ByVal outTable As DataTable) As DataTable

        Dim view As DataView
        Dim numberOfRows As Integer = 0
        Dim part As Integer = 0

        part = 0
        view = New DataView(mainTable)
        view.RowFilter = filter
        Dim tableEPAP As DataTable = view.ToTable()

        If tableEPAP.Rows.Count > 0 Then
            outTable = PrepareTable(tableEPAP, exceptionMessage, sheetName, 450, part, outTable)
        End If

        Return outTable
    End Function

    Private Function PrepareTable(ByVal tableForSave As DataTable, ByRef exceptionMessage As String, sheetName As String, ByVal maxRowsInOnePart As Integer, ByRef part As Integer, ByVal outTable As DataTable) As DataTable
        Dim numberOfRows As Integer = tableForSave.Rows.Count
        Dim partsCollection As DataTable = GetPartsCollection(numberOfRows, maxRowsInOnePart)

        For i As Integer = partsCollection.Rows.Count - 1 To 0 Step -1
            Dim cutedTable As DataTable = CutRowsFromTable(tableForSave, partsCollection.Rows(i)("startRow"), partsCollection.Rows(i)("endRow"))
            Dim rowOutTable As DataRow = outTable.NewRow
            rowOutTable("SheetName") = sheetName & part
            rowOutTable("TBL") = cutedTable
            outTable.Rows.Add(rowOutTable)
            part = part + 1
        Next

        Return outTable
    End Function

    Private Sub CreateTableFromFilter(localFolder As String, excelFile As String, ByRef exceptionMessage As String, mainTable As DataTable, sheetName As String, filter As String, part As Integer, maxRowsInOnePart As Integer)
        Dim view As DataView = New DataView(mainTable)
        view.RowFilter = filter
        Dim tableFromView As DataTable = view.ToTable()
        If tableFromView.Rows.Count > 0 Then
            PrepareToSaveTableToExcel(tableFromView, localFolder, excelFile, exceptionMessage, sheetName, maxRowsInOnePart, part)
        End If
    End Sub

    Private Function CreateTableFromFilter(localFolder As String, ByRef exceptionMessage As String, mainTable As DataTable, sheetName As String, filter As String, part As Integer, maxRowsInOnePart As Integer, outTable As DataTable) As DataTable
        Dim view As DataView = New DataView(mainTable)
        view.RowFilter = filter
        view.Sort = "[F4]"
        Dim tableFromView As DataTable = view.ToTable()
        If tableFromView.Rows.Count > 0 Then
            outTable = PrepareTable(tableFromView, exceptionMessage, sheetName, maxRowsInOnePart, part, outTable)
        End If

        Return outTable
    End Function

    Private Sub CreateDBTableByLender(localFolder As String, excelFile As String, ByRef exceptionMessage As String, mainTable As DataTable, sheetName As String, filter As String, part As Integer)

        Dim view As DataView = New DataView(mainTable)
        view.RowFilter = filter
        Dim tableCitibankByLender As DataTable = view.ToTable()
        If tableCitibankByLender.Rows.Count > 0 Then
            For i As Integer = 0 To tableCitibankByLender.Rows.Count - 1
                Dim row As DataRow = tableCitibankByLender.Rows(i)
                If DBNull.Value.Equals(row("F22")) OrElse row("F22") = Nothing Then
                    If DBNull.Value.Equals(row("F26")) OrElse row("F26") = Nothing OrElse row("F26") = "" Then
                        Continue For
                    Else
                        row("F22") = row("F26")
                    End If
                End If
            Next
            filter = "[F22] Is NOT NULL"
            view = New DataView(tableCitibankByLender)
            view.RowFilter = filter
            tableCitibankByLender = view.ToTable()
            If tableCitibankByLender.Rows.Count > 0 Then
                view = New DataView(tableCitibankByLender)
                Dim paymentsReferencesTable As DataTable = view.ToTable(True, "F22") 'это таблица ссылок на платеж
                'paymentsReferencesTable = RemoveNullValue(paymentsReferencesTable, "F22")

                For Each paymentReference As DataRow In paymentsReferencesTable.Rows
                    ' накладываем на неё фильтр по ссылке на платеж
                    view = New DataView(tableCitibankByLender)
                    filter = "[F22] = '" & paymentReference("F22") & "'"
                    view.RowFilter = filter
                    PrepareToSaveTableToExcel(view.ToTable(), localFolder, excelFile, exceptionMessage, sheetName, 450, part)
                Next
            End If
        End If
    End Sub

    Private Function CreateDBTableByLender(ByVal localFolder As String, ByRef exceptionMessage As String, ByVal mainTable As DataTable, ByVal sheetName As String, ByVal filter As String, ByVal part As Integer, ByVal outTable As DataTable) As DataTable

        Dim view As DataView = New DataView(mainTable)
        view.RowFilter = filter
        view.Sort = "[F4]"
        Dim tableCitibankByLender As DataTable = view.ToTable()
        If tableCitibankByLender.Rows.Count > 0 Then
            For i As Integer = 0 To tableCitibankByLender.Rows.Count - 1
                Dim row As DataRow = tableCitibankByLender.Rows(i)
                If DBNull.Value.Equals(row("F22")) OrElse row("F22") = Nothing Then
                    If DBNull.Value.Equals(row("F26")) OrElse row("F26") = Nothing OrElse row("F26") = "" Then
                        Continue For
                    Else
                        row("F22") = row("F26")
                    End If
                End If
            Next
            filter = "[F22] Is NOT NULL"
            view = New DataView(tableCitibankByLender)
            view.RowFilter = filter
            view.Sort = "[F4]"
            tableCitibankByLender = view.ToTable()
            If tableCitibankByLender.Rows.Count > 0 Then
                view = New DataView(tableCitibankByLender)
                Dim paymentsReferencesTable As DataTable = view.ToTable(True, "F22") 'это таблица ссылок на платеж
                'paymentsReferencesTable = RemoveNullValue(paymentsReferencesTable, "F22")

                For Each paymentReference As DataRow In paymentsReferencesTable.Rows
                    ' накладываем на неё фильтр по ссылке на платеж
                    view = New DataView(tableCitibankByLender)
                    filter = "[F22] = '" & paymentReference("F22") & "'"
                    view.RowFilter = filter
                    view.Sort = "[F4]"
                    outTable = PrepareTable(view.ToTable(), exceptionMessage, sheetName, 800, part, outTable)
                Next
            End If
        End If

        Return outTable
    End Function

    Private Sub PrepareToSaveTableToExcel(tableForSave As DataTable, localFolder As String, excelFile As String, ByRef exceptionMessage As String, sheetName As String, maxRowsInOnePart As Integer, ByRef part As Integer)
        Dim numberOfRows As Integer = tableForSave.Rows.Count
        Dim partsCollection As DataTable = GetPartsCollection(numberOfRows, maxRowsInOnePart)

        For i As Integer = partsCollection.Rows.Count - 1 To 0 Step -1
            Dim cutedTable As DataTable = CutRowsFromTable(tableForSave, partsCollection.Rows(i)("startRow"), partsCollection.Rows(i)("endRow"))
            If SaveTableToExcel(localFolder, excelFile, cutedTable, sheetName & part, exceptionMessage) = False Then
                Throw New Exception("I can't add a sheet " & sheetName & " to a file " & excelFile & ". Exception : " & exceptionMessage)
            End If
            part = part + 1
        Next
    End Sub

    Private Sub SaveDataTableToFile(fileName As String, table As System.Data.DataTable)
        Dim Stream As FileStream = New FileStream(fileName, FileMode.Create)
        Dim serializer As XmlSerializer = New XmlSerializer(table.GetType())
        serializer.Serialize(Stream, table)
        Stream.Close()
    End Sub

    Private Function GetTableFromFile(localFolder As String, tableInXML As String) As Data.DataTable
        Dim table As System.Data.DataTable = New System.Data.DataTable
        Dim stream As FileStream = New FileStream(localFolder & "\" & tableInXML, FileMode.Open, FileAccess.Read)
        Dim deSerializer As XmlSerializer = New XmlSerializer(table.GetType())
        table = deSerializer.Deserialize(stream)
        stream.Close()
        Return table
    End Function

    Private Sub CreateSheetEPAP(localFolder As String, excelFile As String, ByRef exceptionMessage As String, mainTable As DataTable, filter As String, sheetName As String)

        Dim view As DataView
        Dim numberOfRows As Integer = 0
        Dim part As Integer = 0
        Dim partsCollection As DataTable = New DataTable()
        Dim cutedTable As DataTable = New DataTable()

        part = 0
        view = New DataView(mainTable)
        view.RowFilter = filter
        Dim tableEPAP As DataTable = view.ToTable()

        If tableEPAP.Rows.Count > 0 Then
            PrepareToSaveTableToExcel(tableEPAP, localFolder, excelFile, exceptionMessage, sheetName, 450, part)
        End If
    End Sub

    Private Function GetPartsCollection(ByVal numberOfRows As Integer, ByVal maxRowsInOnePart As Integer) As DataTable
        Dim currentRow As Integer = numberOfRows
        Dim countRow = 0
        Dim partsCollection As DataTable = New DataTable("PartsCollection")

        partsCollection.Columns.Add("startRow", Type.GetType("System.Int32"))
        partsCollection.Columns.Add("endRow", Type.GetType("System.Int32"))

        While currentRow > 0
            Dim rowNumber
            If countRow = 0 Then
                rowNumber = currentRow Mod maxRowsInOnePart
                countRow = countRow + rowNumber
                currentRow = currentRow - rowNumber - 1
            Else
                rowNumber = (currentRow Mod maxRowsInOnePart) + 1
                countRow = countRow + rowNumber
                currentRow = currentRow - rowNumber
            End If
            Dim tableRow As DataRow = partsCollection.NewRow
            If countRow - maxRowsInOnePart < 0 Then
                tableRow("startRow") = 0
                tableRow("endRow") = countRow
            Else
                tableRow("startRow") = countRow - (maxRowsInOnePart - 1)
                tableRow("endRow") = countRow
            End If
            partsCollection.Rows.Add(tableRow)
        End While

        Return partsCollection
    End Function

    Private Function CutRowsFromTable(ByVal inputTable As DataTable, ByVal startRow As Integer, ByVal endRow As Integer) As DataTable
        Dim resultTable As DataTable
        resultTable = inputTable.Clone()

        If (inputTable IsNot Nothing) Then
            If (inputTable.Rows.Count > 0) Then
                For i As Integer = startRow To inputTable.Rows.Count - 1
                    resultTable.ImportRow(inputTable.Rows(i))
                    If i = endRow Then
                        Return resultTable
                    End If
                Next
            End If
        End If
        Return resultTable
    End Function

    Private Function SaveTableToExcel(ByVal localFolder As String, ByVal excelFile As String, ByVal tempTable As DataTable, ByVal sheetName As String, ByRef exceptionMessage As String) As Boolean
        Dim xlApp As Excel.Application = New Excel.Application()
        'Dim xlWorkBook As Excel.Workbook
        'Dim xlWorkSheet As Excel.Worksheet
        Dim xlWorkBook As Object = New Object()
        Dim xlWorkSheet As Object = New Object()
        Dim misValue As Object = Reflection.Missing.Value
        Dim isSaved As Boolean = False

        Dim fullFileName As String = localFolder & "\" & excelFile

        Try
            xlWorkBook = xlApp.Workbooks.Open(fullFileName)
            Dim worksheets As Excel.Sheets = xlWorkBook.Worksheets
            Dim worksheet1 As Excel.Worksheet = CType(worksheets(1), Excel.Worksheet)
            'xlWorkSheet = DirectCast(worksheets.Add(worksheets(1), Type.Missing, Type.Missing, Type.Missing), Excel.Worksheet)
            'xlWorkSheet.Name = sheetName
            worksheet1.Copy(After:=worksheet1)
            xlWorkSheet = CType(worksheets(2), Excel.Worksheet)
            xlWorkSheet.Name = sheetName

            Dim timeArray(tempTable.Rows.Count, tempTable.Columns.Count) As Object
            Dim row As Integer, col As Integer

            For row = 0 To tempTable.Rows.Count - 1
                For col = 0 To tempTable.Columns.Count - 1
                    timeArray(row, col) = tempTable.Rows(row).Item(col)
                Next
            Next

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
            releaseObject(xlWorkSheet)
            releaseObject(xlWorkBook)
            releaseObject(xlApp)

            KillExcell()
        End Try

        Return isSaved
    End Function

    Private Function CreateNewExcel(ByVal localFolder As String, ByVal excelFile As String, ByRef exceptionMessage As String) As Boolean
        ' Create Excel
        Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        'Dim xlWorkBook As Excel.Workbook
        'Dim xlWorkSheet As Excel.Worksheet
        Dim xlWorkBook As Object = New Object()
        Dim xlWorkSheet As Object = New Object()
        Dim isCreate As Boolean = False
        Try
            Dim misValue As Object = System.Reflection.Missing.Value

            xlWorkBook = xlApp.Workbooks.Add(misValue)
            xlWorkSheet = CType(xlWorkBook.Sheets(1), Excel.Worksheet)

            xlWorkBook.SaveAs(localFolder & "\" & excelFile, 50)
            xlWorkBook.Close(True, misValue, misValue)
            xlApp.Quit()
            isCreate = True
        Catch ex As Exception
            exceptionMessage = ex.Message
            If (Not ex.InnerException Is Nothing) Then
                exceptionMessage = exceptionMessage & " Inner Exception : " & ex.InnerException.Message
            End If
        Finally
            releaseObject(xlWorkSheet)
            releaseObject(xlWorkBook)
            releaseObject(xlApp)

            KillExcell()
        End Try
        ' /Create Excel
        Return isCreate
    End Function

    Private Sub KillExcell()
        Dim proc As Process
        For Each proc In Process.GetProcessesByName("EXCEL")
            proc.Kill()
        Next
    End Sub

    Private Function GetDataTableFromFile(ByVal fileName As String) As DataTable
        Dim table As DataTable = New DataTable()
        Dim stream As FileStream = New FileStream(fileName, FileMode.Open, FileAccess.Read)
        Dim deSerializer As XmlSerializer = New XmlSerializer(table.GetType())

        table = deSerializer.Deserialize(stream)
        stream.Close()

        Return table
    End Function

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

End Module
