Imports System.IO
Imports System.Xml.Serialization

Module Module1

    Private localFolder As String = "C:\Temp\WorkDir"
    Private remoteFolder As String = "\\rus.efesmoscow\DFS\MOSC\Projects.MOSC\Robotic\P12280 Payment Run\WorkDir\Log"
    'Private dateAndTime As DateTime = Convert.ToDateTime("18.06.2022")
    Private be As String = "RU17" ' текущая BE
    Private filePlanName As String = "ero .Общий план_ХХ-ХХ_недели_RU17.xlsb"
    Private paymentDate As Date = Convert.ToDateTime("18.06.2022")

    ' *************** output variables
    Private isStageOneCreated As Boolean
    ' *************** output variables

    Sub Main()

        isStageOneCreated = False
        be = be.ToUpper()
        Dim logTableFileName As String = "Log.xml"
        DeleteFile(localFolder, logTableFileName)
        CopyFile(remoteFolder, logTableFileName, localFolder, logTableFileName)

        Dim logTable As System.Data.DataTable = GetTableFromFile(localFolder, logTableFileName)
        Dim view As DataView
        Dim _filter As String = ""
        Dim tempTable As DataTable

        If be = "RU01" Then
            _filter = "[Ru01] = True And [FilePlanRu01] = '" & filePlanName & "' And [PaymentDateRu01] = '" & paymentDate & "'"
        End If
        If be = "RU14" Then
            _filter = "[Ru14] = True And [FilePlanRu14] = '" & filePlanName & "' And [PaymentDateRu14] = '" & paymentDate & "'"
        End If
        If be = "RU17" Then
            _filter = "[Ru17] = True And [FilePlanRu17] = '" & filePlanName & "' And [PaymentDateRu17] = '" & paymentDate & "'"
        End If

        view = New DataView(logTable)
        view.RowFilter = _filter
        tempTable = view.ToTable()
        If tempTable.Rows.Count > 0 Then
            isStageOneCreated = True
        End If

    End Sub

    Private Function GetTableFromFile(localFolder As String, tableInXML As String) As System.Data.DataTable
        Dim table As System.Data.DataTable = New System.Data.DataTable
        table.ReadXmlSchema(localFolder & "\" & tableInXML)
        table.ReadXml(localFolder & "\" & tableInXML)
        Return table
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

End Module
