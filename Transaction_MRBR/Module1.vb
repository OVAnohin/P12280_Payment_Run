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

Module Module1

    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function SendMessageHM Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Int32, ByVal wMsg As Int32, ByVal wParam As Int32, ByVal lParam As String) As Int32
    Private Declare Function SendMessageW Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Integer
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As Long) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    Private Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As System.IntPtr, ByVal lpEnumFunc As EnumWindowProc, ByVal lParam As Integer) As Boolean
    Private Delegate Function EnumWindowProc(ByVal hWnd As IntPtr, ByVal lParam As IntPtr) As Boolean
    Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Boolean

    Private Const WM_COMMAND = &H111
    Private Const BM_CLICK As Integer = &HF5

    ' *************** Private
    Private _thread As Thread
    Private _oLock As Object = New Object()
    Private _table As DataTable = New DataTable()
    ' *************** Private

    Sub Main()

        ' *************** input variables
        Dim localFolder As String = "C:\Temp\WorkDir"
        ' *************** input variables

        ' *************** output variables
        Dim exceptionMessage As String
        Dim isComplete As Boolean
        ' *************** output variables

        Console.WriteLine("Первичный поток: Id {0}", Thread.CurrentThread.ManagedThreadId)

        ' *************** Begin
        Dim timeout As DateTime
        Dim isExit As Boolean = False
        Dim session = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0)
        Dim transactionName As String = "MRBR"
        Dim forUpLoadBE As String = "forUpLoadBE.txt"
        Dim forUpLoad As String = "forUpLoad.txt"

        isComplete = False

        Try
            session.findById("wnd[0]").maximize
            session.findById("wnd[0]/tbar[0]/okcd").text = transactionName
            session.findById("wnd[0]/tbar[0]/btn[0]").press

            ' wait window
            isExit = False
            timeout = DateTime.Now.AddSeconds(5)
            While (isExit = False)
                If Not session.findbyid("wnd[0]/titl", False) Is Nothing Then
                    If session.findbyid("wnd[0]/titl").Text = "Деблокировать блокированные счета" Then
                        isExit = True
                    Else
                        CheckTimeout(timeout)
                    End If
                Else
                    CheckTimeout(timeout)
                End If
            End While
            ' wait window

            'грузим БЕ
            session.findbyid("wnd[0]/usr/btn%_SO_BUKRS_%_APP_%-VALU_PUSH").Press
            'clear rows
            SetForegroundWindow(session.findById("wnd[1]").Handle)
            session.findById("wnd[1]/tbar[0]/btn[16]").press

            TaskPressIpmportButton(session)
            'Console.WriteLine(IsGuiModalWindow(session))
            UploadFileInGuiModalWiondow(session, forUpLoadBE, localFolder)

            SetForegroundWindow(session.findById("wnd[1]").Handle)
            session.findbyid("wnd[1]/tbar[0]/btn[8]").Press

            'грузим счет документа
            session.findbyid("wnd[0]/usr/btn%_SO_BELNR_%_APP_%-VALU_PUSH").Press
            'clear rows
            SetForegroundWindow(session.findById("wnd[1]").Handle)
            session.findById("wnd[1]/tbar[0]/btn[16]").press

            TaskPressIpmportButton(session)
            'Console.WriteLine(IsGuiModalWindow(session))
            UploadFileInGuiModalWiondow(session, forUpLoad, localFolder)

            SetForegroundWindow(session.findById("wnd[1]").Handle)
            session.findbyid("wnd[1]/tbar[0]/btn[8]").Press

            'start transaction
            session.findbyid("wnd[0]/tbar[1]/btn[8]").Press

            Thread.Sleep(1000)

            ' wait window
            isExit = False
            timeout = DateTime.Now.AddSeconds(5)
            While (isExit = False)
                If Not session.findbyid("wnd[0]/usr/cntlGRID1/shellcont/shell", False) Is Nothing Then
                    isExit = True
                Else
                    If session.findbyid("wnd[0]/sbar/pane[0]").Text = "К Вашим критериям выбора блокированных счетов нет" Then
                        isComplete = True
                        ' *************** End
                        Console.WriteLine("Первичный поток: Id {0} is Ended", Thread.CurrentThread.ManagedThreadId)
                        Exit Sub
                    Else
                        CheckTimeout(timeout)
                    End If
                End If
            End While
            ' wait window

            session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell(-1, "")
            session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").selectAll
            session.findById("wnd[0]/tbar[1]/btn[9]").press

            'button save
            session.findbyid("wnd[0]/tbar[0]/btn[11]").Press

            Thread.Sleep(1000)
            isComplete = True

        Catch ex As Exception
            exceptionMessage = ex.Message
            If (Not ex.InnerException Is Nothing) Then
                exceptionMessage = exceptionMessage & " Inner Exception : " & ex.InnerException.Message
            End If

            Console.WriteLine("Exception       : " & ex.GetType().ToString())
            Console.WriteLine("Message         : " & ex.Message)
            If (Not ex.InnerException Is Nothing) Then
                Console.WriteLine("Inner Exception : " & ex.InnerException.Message)
            End If

        Finally
            session = Nothing
        End Try

        ' *************** End
        Console.WriteLine("Первичный поток: Id {0} is Ended", Thread.CurrentThread.ManagedThreadId)
        Console.ReadKey()

    End Sub

    Private Sub UploadFileInGuiModalWiondow(ByVal session As Object, ByVal fileForUpLoad As String, ByVal path As String)
        If (IsGuiModalWindow(session, "wnd[2]")) Then
            session.findById("wnd[2]/usr/ctxtDY_PATH").text = path
            session.findById("wnd[2]/usr/ctxtDY_FILENAME").text = fileForUpLoad
            session.findById("wnd[2]/tbar[0]/btn[0]").press
        Else
            'Real Mode
            TryToFoundWindowAndSetForeground("Открыть")
            SaveAsWindow(fileForUpLoad, path, "Открыть", "Открыть")
        End If
    End Sub

    Private Sub CheckTimeout(ByVal timeout As Date)
        If (DateTime.Now > timeout) Then
            Throw New Exception("Не открылась транзакция")
        End If
    End Sub

    Private Sub TaskPressButtonOk(ByVal session As Object)
        Dim taskPressButton As Task = New Task(AddressOf PressButtonOk, session)
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

    Private Sub PressIpmportButton(ByVal session As Object)

        _thread = Thread.CurrentThread
        session.findById("wnd[1]/tbar[0]/btn[23]").press

    End Sub

    Private Sub GetDataTableFromFile(ByVal fileName As String)
        Dim stream As FileStream = New FileStream(fileName, FileMode.Open, FileAccess.Read)
        Dim deSerializer As XmlSerializer = New XmlSerializer(_table.GetType())

        _table = deSerializer.Deserialize(stream)
        stream.Close()
    End Sub

    Private Sub SaveDataTableToFile(ByVal fileName As String)
        Dim Stream As FileStream = New FileStream(fileName, FileMode.Create)
        Dim serializer As XmlSerializer = New XmlSerializer(_table.GetType())
        serializer.Serialize(Stream, _table)
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
        ReleaseObject(session)
        session = Nothing

    End Sub

    Private Sub TaskKill(ByVal taskName As String)
        For Each oProcess As System.Diagnostics.Process In System.Diagnostics.Process.GetProcessesByName(taskName)
            oProcess.Kill()
        Next
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
                    ReleaseObject(session)
                    session = Nothing
                Catch ex As Exception
                    Throw New Exception("Не могу закрыть соседние сессии.")
                End Try
            Next
        End If

        ReleaseObject(session)
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
        ReleaseObject(session)
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

End Module
