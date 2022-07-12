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
    'Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hWnd As IntPtr) As Boolean
    Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Integer, ByVal nCmdShow As Integer) As Integer
    Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Integer, ByRef lpdwProcessId As Integer) As Integer
    Private Declare Sub keybd_event Lib "user32.dll" (bVk As Byte, bScan As Byte, dwFlags As UInteger, dwExtraInfo As Integer)

    Private Const WM_COMMAND = &H111
    Private Const BM_CLICK As Integer = &HF5
    Private Const ALT As Integer = &HA4
    Private Const EXTENDEDKEY As Integer = &H1
    Private Const KEYUP As Integer = &H2
    Private Const Restore As UInteger = 9

    ' *************** Thread
    Private _thread As Thread
    ' *************** Thread

    ' *************** input variables
    Dim localFolder As String = "C:\Temp\WorkDir"
    Dim paymentDate As Date = Convert.ToDateTime("11.07.2022")
    Dim sheetName As String = "Citibank - RUB debtors I 0" 'это у нас имя листа
    Dim be As String = "RU01"
    Dim xmlNameNotIncludedTable As String = "NotIncludedTableRu01.xml"
    Dim xmlNameErrors_FBL1N_FBL5N As String = "Errors_FBL1N_FBL5N.xml"
    ' *************** input variables

    ' *************** in variables
    Dim _oLock As Object = New Object()
    ' *************** in variables

    ' *************** output variables
    Dim exceptionMessage As String
    Dim isComplete As Boolean
    Dim viewPosition As Integer ' количество документов для обработки
    ' *************** output variables

    Sub Main()

        Console.WriteLine("Первичный поток: Id {0}", Thread.CurrentThread.ManagedThreadId)

        viewPosition = 0
        'isNumberDocumentsEqualNumberSheet = False
        'isDocumentsChanged = False

        Dim xmlFileName As String = sheetName & ".xml"
        Dim txtFileName As String = sheetName & ".txt"
        Dim xlsbFileName As String = sheetName & ".xlsb"
        Dim tablesToRunFileNameXml As String = "TablesToRun.xml" ' выгружается из призмы
        Dim xlsbNameNotIncludedTable As String = Replace(xmlNameNotIncludedTable, ".xml", ".xlsb")
        Dim timeout As DateTime
        Dim connectionList As List(Of String) = New List(Of String)
        Dim isExit As Boolean = False
        Dim ownBank As String = GetOwnBank(sheetName)
        Dim numberOfDocuments As Integer ' 'это у нас количество строк на листе оно должно быть равно количеству viewPosition
        Dim transactionName As String = GetTransactionName(sheetName)
        Dim outputFormat As String = GetOutputFormat(transactionName, be)
        Dim isDaTapPresent As Boolean

        'удаляем старые файлы для выгрузки xmlFileName and xlsbFileName
        DeleteFile(localFolder, xmlFileName)
        DeleteFile(localFolder, xlsbFileName)
        DeleteFile(localFolder, txtFileName)

        ' текущая таблица для запуска или текущий лист
        Dim tableCurrentRun As System.Data.DataTable = New Data.DataTable()
        'получаем данные текущего прогона из таблицы "TablesToRun.xml", она обновляется тут в коде
        Dim tablesToRun As Data.DataTable = GetTableFromFile(localFolder, tablesToRunFileNameXml)
        For i As Integer = tablesToRun.Rows.Count - 1 To 0 Step -1
            If tablesToRun.Rows(i)("SheetName") = sheetName Then
                tableCurrentRun = tablesToRun.Rows(i)("SheetData")
                Exit For
            End If
        Next
        numberOfDocuments = tableCurrentRun.Rows.Count

        For i As Integer = 0 To tableCurrentRun.Rows.Count - 1
            If CType(tableCurrentRun.Rows(i)("F1"), String).Contains("да-ТАП") AndAlso CType(tableCurrentRun.Rows(i)("F3"), String).Contains("EPAP") Then
                isDaTapPresent = True
            End If
        Next

        'таблица не включенных проводок в прогон
        Dim tableErrors_FBL1N_FBL5N As Data.DataTable = New Data.DataTable()
        Try
            tableErrors_FBL1N_FBL5N = GetTableFromFile(localFolder, xmlNameErrors_FBL1N_FBL5N)
        Catch ex As Exception
        End Try

        Dim session = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0)

        Try
            SyncLock _oLock
                session.findById("wnd[0]").maximize
                session.findById("wnd[0]/tbar[0]/okcd").Text = transactionName
                session.findById("wnd[0]/tbar[0]/btn[0]").press

                ' грузим LenderAaccount
                If sheetName.Contains("debtors") Then
                    session.findbyid("wnd[0]/usr/btn%_DD_KUNNR_%_APP_%-VALU_PUSH").Press
                Else
                    session.findbyid("wnd[0]/usr/btn%_KD_LIFNR_%_APP_%-VALU_PUSH").Press
                End If
                ''это у нас эмуляция нажатия alt
                'keybd_event(CType(ALT, Byte), &H45, EXTENDEDKEY Or 0, 0)
                'keybd_event(CType(ALT, Byte), &H45, EXTENDEDKEY Or KEYUP, 0)
                ''/это у нас эмуляция нажатия alt
                UploadFromFileInMultipleSelectionWindow(session, "LenderAaccount.txt", localFolder)

                ' Грузим БЕ
                If sheetName.Contains("debtors") Then
                    session.findbyid("wnd[0]/usr/btn%_DD_BUKRS_%_APP_%-VALU_PUSH").Press
                Else
                    session.findbyid("wnd[0]/usr/btn%_KD_BUKRS_%_APP_%-VALU_PUSH").Press
                End If
                UploadFromFileInMultipleSelectionWindow(session, "forUpLoadBE.txt", localFolder)

                ' Номер документа – номер проводки, через специальную вставку
                ' Динамические ограничения выбора   (Shift+F4)
                session.findbyid("wnd[0]/tbar[1]/btn[16]").Press
                session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN011_%_APP_%-VALU_PUSH").press
                UploadFromFileInMultipleSelectionWindow(session, "DocumentNumber.txt", localFolder)

                session.findbyid("wnd[0]/usr/chkX_NORM").Selected = True
                session.findById("wnd[0]/usr/chkX_SHBV").selected = True
                session.findById("wnd[0]/usr/chkX_MERK").selected = True
                session.findById("wnd[0]/usr/ctxtPA_STIDA").text = paymentDate
                session.findById("wnd[0]/usr/ctxtPA_VARI").text = outputFormat

                'Выполнить
                session.findbyid("wnd[0]/tbar[1]/btn[8]").Press

                Thread.Sleep(500)
                ' wait window
                isExit = False
                timeout = DateTime.Now.AddSeconds(5)
                While (isExit = False)
                    If Left(session.findbyid("wnd[0]/sbar/pane[0]").Text, 22) = "Выведено для просмотра" Then
                        isExit = True
                    Else
                        If session.findbyid("wnd[0]/sbar/pane[0]").Text = "Позиции не выбраны (см. подробный текст)" Then
                            Console.WriteLine("Позиции не выбраны (см. подробный текст)")
                            ReturnToMainWindow(session)
                            isComplete = True
                            Exit Sub
                        Else
                            CheckTimeout(timeout)
                        End If
                    End If
                End While
                ' wait window

                Dim viewPositionText As String = session.findbyid("wnd[0]/sbar/pane[0]").Text
                viewPositionText = Mid(viewPositionText, 23)
                Dim positionInText As Integer = viewPositionText.IndexOf("позиц", 0)
                viewPositionText = RemoveUnnecessaryChar(Left(viewPositionText, positionInText))

                'Проверка количества документов в обработке, если их меньше то нужно сделать лист в который записать все что не попало.
                If Not Int32.TryParse(Left(viewPositionText, positionInText), viewPosition) Then
                    ReturnToMainWindow(session)
                    isComplete = True
                    Console.WriteLine("Выход по причине не возможности понять, сколько документво для обработки.")
                    Exit Sub
                End If

                'Проверка количества документов в обработке, если их меньше то нужно сделать лист (новый excel) в который записать все что не попало.
                If viewPosition <> numberOfDocuments Then
                    SetForegroundWindow(session.findById("wnd[0]").Handle)
                    session.findById("wnd[0]").sendVKey(5)
                    session.findById("wnd[0]").sendVKey(16)
                    session.findById("wnd[1]/usr/radRB_OTHERS").setFocus
                    session.findById("wnd[1]/usr/radRB_OTHERS").select
                    session.findById("wnd[1]/usr/cmbG_LISTBOX").setFocus
                    session.findById("wnd[1]/usr/cmbG_LISTBOX").key = "04"

                    Thread.Sleep(500)
                    TaskPressButtonOk(session)
                    If (IsGuiModalWindow(session, "wnd[1]")) Then
                        session.findById("wnd[1]/usr/ctxtDY_PATH").text = localFolder
                        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = txtFileName
                        session.findById("wnd[1]/tbar[0]/btn[11]").press
                    Else
                        'Real Mode
                        TryToFoundWindowAndSetForeground("Открыть")
                        SaveAsWindow(txtFileName, localFolder, "Сохранение", "Сохранить")
                    End If

                    'FoundExcelAndSaveIt(localFolder, xlsbFileName)

                    If TryFindNotepadAndCloseIt() = False Then
                        Throw New Exception("Нет выгрузки ")
                    End If

                    If CheckFileExists(localFolder, txtFileName) = False Then
                        Throw New Exception("Нет выгрузки ")
                    End If

                    RenameFile(localFolder, txtFileName, localFolder, xmlFileName)

                    OpenXmlFileAndSaveIt(localFolder, xmlFileName, xlsbFileName)

                    Thread.Sleep(500)
                    Dim dataFromExcel As System.Data.DataTable = GetDatatableFromExcel(localFolder, xlsbFileName, 1)

                    For i As Integer = 0 To dataFromExcel.Rows.Count - 1
                        If Not DBNull.Value.Equals(dataFromExcel.Rows(i)("№ документа")) Then
                            For j As Integer = tableCurrentRun.Rows.Count - 1 To 0 Step -1
                                If tableCurrentRun.Rows(j)("F7") = dataFromExcel.Rows(i)("№ документа") Then
                                    tableCurrentRun.Rows.Remove(tableCurrentRun.Rows(j))
                                End If
                            Next
                        End If
                    Next

                    tableErrors_FBL1N_FBL5N = MergeTwoTables(tableErrors_FBL1N_FBL5N, tableCurrentRun)
                    tableErrors_FBL1N_FBL5N.TableName = "notIncludedTable"
                    Dim row As Data.DataRow = tableErrors_FBL1N_FBL5N.NewRow()
                    row(0) = sheetName
                    row(1) = "* * *"
                    tableErrors_FBL1N_FBL5N.Rows.Add(row)
                    'тут надо dataFromExcel типа сохранить на лист, это у нас то что не вошло в выбор документов
                    SaveDataTableToFile(localFolder & "\" & xmlNameErrors_FBL1N_FBL5N, tableErrors_FBL1N_FBL5N)
                    ResetSmartTableInExcel(localFolder, xlsbNameNotIncludedTable, "Errors_FBL1N_FBL5N")
                    If SaveTableToExcel(localFolder, xlsbNameNotIncludedTable, tableErrors_FBL1N_FBL5N, 1, exceptionMessage) = False Then
                        Throw New Exception("Не могу сохранить данные в '" & xlsbNameNotIncludedTable & "'")
                    End If

                End If

                session.findById("wnd[0]").maximize
                session.findById("wnd[0]").sendVKey(5)
                session.findById("wnd[0]").sendVKey(45)
                session.findById("wnd[1]/usr/ctxt*BSEG-HBKID").text = ownBank
                If sheetName.Contains("TAP") OrElse sheetName.Contains("debtors") Then
                    session.findById("wnd[1]/usr/ctxt*BSEG-ZLSPR").text = " "
                ElseIf isDaTapPresent Then
                    session.findById("wnd[1]/usr/ctxt*BSEG-ZLSPR").text = " "
                End If

                session.findById("wnd[1]/usr/ctxt*BSEG-HBKID").setFocus
                session.findById("wnd[1]/usr/ctxt*BSEG-HBKID").caretPosition = 5
                'Временно отключаем изменение документов.
                session.findById("wnd[1]/tbar[0]/btn[0]").press

                Thread.Sleep(1000)

                ' wait window
                isExit = False
                timeout = DateTime.Now.AddSeconds(10)
                While (isExit = False)
                    If Not session.findbyid("wnd[1]", False) Is Nothing Then
                        If session.findbyid("wnd[1]/usr/txtMESSTXT1").Text = "Не удалось изменить все документы." Then
                            session.findbyid("wnd[1]/tbar[0]/btn[0]").Press
                            Console.WriteLine("Не удалось изменить все документы.")
                            ReturnToMainWindow(session)
                            isComplete = True
                            Exit Sub
                        Else
                            Throw New Exception("Не удалось изменить все документы.")
                        End If
                        isExit = True
                    Else
                        Try
                            Dim resultText As String = session.findbyid("wnd[0]/sbar/pane[0]").Text
                            If resultText = "Изменения сохранены." Then
                                Console.WriteLine("Изменения сохранены.")
                                ReturnToMainWindow(session)
                                isComplete = True
                                Exit Sub
                            End If
                            CheckTimeout(timeout)
                        Catch ex As Exception
                        End Try
                    End If
                End While
                ' wait window

                isComplete = True

            End SyncLock

        Catch ex As Exception
            exceptionMessage = ex.Message
            If (Not ex.InnerException Is Nothing) Then
                exceptionMessage = exceptionMessage & " Inner Exception : " & ex.InnerException.Message
            End If
            isComplete = False

            Console.WriteLine("Exception       : " & ex.GetType().ToString())
            Console.WriteLine("Message         : " & ex.Message)
            If (Not ex.InnerException Is Nothing) Then
                Console.WriteLine("Inner Exception : " & ex.InnerException.Message)
            End If
        Finally
            session = Nothing
            Console.WriteLine("exceptionMessage " & exceptionMessage)
            Console.WriteLine("isComplete " & isComplete)
            Console.WriteLine("viewPosition " & viewPosition)
            Console.WriteLine("Первичный поток: Id {0} is Ended", Thread.CurrentThread.ManagedThreadId)
        End Try

        Console.WriteLine("Первичный поток: Id {0} is Ended", Thread.CurrentThread.ManagedThreadId)
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

    Private Sub ResetSmartTableInExcel(localFolder As String, fileName As String, sheetName As String)
        Dim xlApp As Excel.Application
        Dim xlWorkBook As Excel.Workbook
        Dim xlWorkSheet As Excel.Worksheet
        Dim misValue As Object = Reflection.Missing.Value
        Dim fullFileName As String = localFolder & "\" & fileName

        Try
            xlApp = New Microsoft.Office.Interop.Excel.Application()
            xlWorkBook = xlApp.Workbooks.Open(fullFileName)
            xlWorkSheet = CType(xlWorkBook.Sheets(sheetName), Excel.Worksheet)

            Dim selectedCell As Object
            Dim tableName As String

            selectedCell = xlWorkSheet.Range("A1")
            tableName = selectedCell.ListObject.Name
            Dim tbl As Object = xlWorkSheet.ListObjects(tableName)

            'Delete all table rows except first row
            If Not tbl.DataBodyRange Is Nothing Then
                If tbl.DataBodyRange.Rows.Count > 1 Then
                    tbl.AutoFilter.ShowAllData
                    tbl.DataBodyRange.Offset(1, 0).Resize(tbl.DataBodyRange.Rows.Count - 1, tbl.DataBodyRange.Columns.Count).Rows.Delete
                End If
                'Clear out data from first table row
                tbl.DataBodyRange.Rows(1).ClearContents
            End If

            xlWorkBook.Save()
            xlWorkBook.Close()
            xlApp.Quit()

            ReleaseObject(xlApp)
            ReleaseObject(xlWorkBook)
            ReleaseObject(xlWorkSheet)

        Catch e As Exception
            Throw New Exception("Не могу очистить таблицу в файле " & fullFileName & " на листе " & sheetName & " " & e.Message)
        Finally
            xlApp = Nothing
            xlWorkBook = Nothing
            xlWorkSheet = Nothing
        End Try
    End Sub

    Private Function GetOutputFormat(ByVal transactionName As String, ByVal be As String) As String
        If be = "UAH01" Then
            Return "/ПЛАН UA"
        End If

        If transactionName = "/nFBL1N" Then
            Return "/PAYM.PLAN"
        End If
        Return "/PAYMENTPLAN"
    End Function

    Private Function GetTransactionName(ByVal sheetName As String) As String
        If sheetName.Contains("debtors") Then
            Return "/nFBL5N"
        End If
        Return "/nFBL1N"
    End Function

    Private Sub DeleteFile(ByVal localFolder As String, ByVal fileToRemove As String)
        Try
            File.Delete(localFolder & "\" & fileToRemove)
        Catch ex As Exception
            Throw New Exception("Не могу удалить файл " & fileToRemove)
        End Try
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

        Return "Citbk"
    End Function

    Private Sub ReturnToMainWindow(session As Object)
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
        session.findbyid("wnd[0]/tbar[0]/btn[0]").Press
    End Sub

    Private Sub CheckTimeout(ByVal timeout As Date)
        If (DateTime.Now > timeout) Then
            Throw New Exception("Не открылась транзакция")
        End If
    End Sub

    Private Sub UploadFromFileInMultipleSelectionWindow(session As Object, fileToUpload As String, localFolder As String)
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

        'keybd_event(CType(ALT, Byte), &H45, EXTENDEDKEY Or 0, 0)
        'keybd_event(CType(ALT, Byte), &H45, EXTENDEDKEY Or KEYUP, 0)
        'Console.WriteLine(session.findById("wnd[1]").Handle)
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

    Private Sub TaskPressButtonOk(session As Object)
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

    Private Sub TaskPressIpmportButton(session As Object)
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

    Private Sub PressButtonOk(session As Object)

        _thread = Thread.CurrentThread
        session.findbyid("wnd[1]/tbar[0]/btn[0]").Press

    End Sub

    Private Sub PressIpmportButton(session As Object)

        _thread = Thread.CurrentThread
        If Not session.findbyid("wnd[1]/tbar[0]/btn[23]", False) Is Nothing Then
            session.findById("wnd[1]/tbar[0]/btn[23]").press
            'session.findbyid("wnd[1]").sendVKey(23)
        Else
            Throw New Exception("Exception from PressIpmportButton")
        End If

    End Sub

    Private Sub SaveDataTableToFile(ByVal fileName As String, ByVal table As System.Data.DataTable)
        Dim Stream As FileStream = New FileStream(fileName, FileMode.Create)
        Dim serializer As XmlSerializer = New XmlSerializer(table.GetType())
        serializer.Serialize(Stream, table)
        Stream.Close()
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
            If autoElement.Current.Name.Equals("Имя файла:") And autoElement.Current.ClassName.Contains("Edit") Then
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
                    Throw New ArgumentNullException("Cannot found launched window " & windowHeader)
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
            Throw New InvalidCastException("GCHandle Targer could not be cast as list")
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

        Dim pidSap As Integer = Shell("C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe", 1)

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

    Private Function IsGuiModalWindow(session As Object, ByVal windowName As String) As Boolean

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

    Private Function SaveTableToExcel(localFolder As String, excelFile As String, tempTable As Data.DataTable, sheetNumber As Integer, ByRef exceptionMessage As String) As Boolean
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

End Module
