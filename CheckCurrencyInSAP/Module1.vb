Imports System.Threading

Module Module1

    ' *************** input variables
    Private paymentDate As Date = Convert.ToDateTime("10.06.2020")
    ' *************** input variables

    ' *************** output variables
    Private exceptionMessage As String
    Private isComplete As Boolean
    Private isCurrencyExchangePresent As Boolean
    Private isCurrencyChecked As Boolean
    ' *************** output variables

    Sub Main()

        '*********************** Begin
        exceptionMessage = ""
        isComplete = False
        isCurrencyExchangePresent = False
        isCurrencyChecked = False

        Dim session = GetObject("SAPGUI").GetScriptingEngine.Children(0).Children(0)

        Try
            session.findById("wnd[0]").maximize
            session.findById("wnd[0]/tbar[0]/okcd").Text = "OB08"
            session.findById("wnd[0]/tbar[0]/btn[0]").press
            Dim Index
            Dim isEnd
            Dim isExit As Boolean = False
            Dim timeout As DateTime

            ' wait window
            isExit = False
            timeout = DateTime.Now.AddSeconds(3)
            While (isExit = False)
                If Not session.findbyid("wnd[1]", False) Is Nothing Then
                    isExit = True
                    session.findbyid("wnd[1]/tbar[0]/btn[0]").Press
                Else
                    If (DateTime.Now > timeout) Then
                        isExit = True
                    End If
                End If
            End While

            session.findById("wnd[0]/mbar/menu[3]/menu[0]").Select
            Dim element As Object


            Index = 4
            isEnd = False
            Do
                Index = Index - 1
                element = session.findById("wnd[1]/usr/tblSAPLSVIXTCTRL_SEL_FLDS")
                If element.Children(0).Text <> "Действит. с" Then
                    session.findById("wnd[1]/usr/tblSAPLSVIXTCTRL_SEL_FLDS").verticalScrollbar.Position = 0
                Else
                    session.findById("wnd[1]/usr/tblSAPLSVIXTCTRL_SEL_FLDS").getAbsoluteRow(0).selected = True
                    session.findById("wnd[1]/usr/tblSAPLSVIXTCTRL_SEL_FLDS").getAbsoluteRow(3).selected = True
                    session.findById("wnd[1]/usr/tblSAPLSVIXTCTRL_SEL_FLDS").verticalScrollbar.position = 1
                    session.findById("wnd[1]/usr/tblSAPLSVIXTCTRL_SEL_FLDS").getAbsoluteRow(4).selected = True
                    isEnd = True
                End If
            Loop Until Index < 0 Or isEnd = True
            If isEnd = False Then
                Throw New Exception("Мое Исключение")
            End If

            session.findById("wnd[1]/tbar[0]/btn[0]").press
            session.findById("wnd[1]/usr/tblSAPLSVIXTCTRL_QUERY/txtQUERY_TAB-BUFFER[3,0]").text = paymentDate
            session.findById("wnd[1]/usr/tblSAPLSVIXTCTRL_QUERY/txtQUERY_TAB-BUFFER[3,1]").text = "M"
            session.findById("wnd[1]/usr/tblSAPLSVIXTCTRL_QUERY/txtQUERY_TAB-BUFFER[3,2]").text = "RUB"
            session.findById("wnd[1]/tbar[0]/btn[0]").press
            session.findById("wnd[1]/tbar[0]/btn[8]").press
            Thread.Sleep(300)

            ' wait window
            isExit = False
            timeout = DateTime.Now.AddSeconds(3)
            While (isExit = False)
                Try
                    Dim resultText As String = session.findbyid("wnd[0]/sbar/pane[0]").Text
                    Console.WriteLine(resultText)
                    If resultText = "По заданным критериям записи не найдены." Then
                        isCurrencyExchangePresent = False
                        ReturnToMainWindow(session)
                        isExit = True
                    End If
                    If resultText.Contains("Число выбранных записей") Then
                        isCurrencyExchangePresent = True
                        ReturnToMainWindow(session)
                        isExit = True
                    End If
                Catch ex As Exception
                End Try
                CheckTimeout(timeout)
            End While
            ' wait window

            Thread.Sleep(300)
            isComplete = True
            isCurrencyChecked = True

        Catch ex As Exception
            exceptionMessage = exceptionMessage & " " & ex.Message
        Finally
            session = Nothing
            Console.WriteLine("isCurrencyExchangePresent = {0}", isCurrencyExchangePresent)
            Console.WriteLine("isComplete = {0}", isComplete)
            Console.WriteLine("exceptionMessage = {0}", exceptionMessage)
            Console.WriteLine("isCurrencyChecked = {0}", isCurrencyChecked)
        End Try

        Console.WriteLine("Первичный поток: Id {0} Is Ended", Thread.CurrentThread.ManagedThreadId)
        Console.ReadKey()

    End Sub

    Private Sub ReturnToMainWindow(session As Object)
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
        session.findbyid("wnd[0]/tbar[0]/btn[0]").Press
    End Sub

    Private Sub CheckTimeout(ByVal timeout As Date)
        If (DateTime.Now > timeout) Then
            Throw New Exception("Не открылась транзакция")
        End If
    End Sub

End Module
