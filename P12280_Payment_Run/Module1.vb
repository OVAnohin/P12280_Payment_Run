Imports System.IO
Imports System.Xml.Serialization

Module Module1

    Sub Main()

        Dim tablePlan As DataTable = New DataTable

        Dim stream As FileStream = New FileStream("c:\Temp\WorkDir\Table_PlanRu.xml", FileMode.Open, FileAccess.Read)
        Dim deSerializer As XmlSerializer = New XmlSerializer(tablePlan.GetType())

        tablePlan = deSerializer.Deserialize(stream)
        stream.Close()

        '*********************** Begin
        Dim view As DataView
        Dim filter As String
        Dim tempTable As DataTable
        Dim searchString As String
        Dim paymentDate As Date = Date.Today

        If (tablePlan IsNot Nothing) Then
            If (tablePlan.Rows.Count > 0) Then
                'fix date
                tablePlan = FixBadDateInTableByColumnName(tablePlan, "Дата документа")
                tablePlan = FixBadDateInTableByColumnName(tablePlan, "Дата ввода")
                tablePlan = FixBadDateInTableByColumnName(tablePlan, "БазДата платежа")
                tablePlan = FixBadDateInTableByColumnName(tablePlan, "Дата платежа")
                tablePlan = FixBadDateInTableByColumnName(tablePlan, "Планируемая Дата платежа")

                view = New DataView(tablePlan)
                filter = "[Платить Да/Нет] Like 'да%'"
                view.RowFilter = filter
                tempTable = view.ToTable()

                Dim planedDate As Date
                Dim filtered = From row In tempTable.AsEnumerable()
                               Where Date.TryParse(row.Field(Of Date)("Планируемая Дата платежа"), planedDate) AndAlso planedDate = paymentDate
                tempTable = filtered.CopyToDataTable()

            End If
        End If

        Console.WriteLine("End")

    End Sub

    Private Function FixBadDateInTableByColumnName(tablePlan As DataTable, ByVal columnName As String) As DataTable
        For i As Integer = 0 To tablePlan.Rows.Count - 1
            If tablePlan.Rows(i)(columnName).GetType().ToString() = "System.DateTime" AndAlso tablePlan.Rows(i)(columnName) = "01.01.0001 0:00:00" Then
                Continue For
            End If
            tablePlan.Rows(i)(columnName) = FixBadDate(tablePlan.Rows(i)(columnName))
        Next

        Return tablePlan
    End Function

    Private Function FixBadDate(currentDateStr As String) As String
        Dim fixDate As DateTime
        Dim resultDate As String

        resultDate = currentDateStr.Substring(0, 10)
        If currentDateStr.Length > 10 Then
            Dim tmpTime As String = currentDateStr.Substring(currentDateStr.Length - 8, 8)
            Dim tmpHour As Integer = Convert.ToInt32(tmpTime.Substring(0, 2))

            Dim tmpDay As Integer = Convert.ToInt32(currentDateStr.Substring(0, 2))
            Dim tmpMounth As Integer = Convert.ToInt32(currentDateStr.Substring(3, 2))
            Dim tmpYear As Integer = Convert.ToInt32(currentDateStr.Substring(6, 4))

            If (tmpHour >= 12 AndAlso tmpHour <= 24) Then
                fixDate = New DateTime(tmpYear, tmpMounth, tmpDay)
                fixDate = fixDate.AddDays(1)
                resultDate = fixDate.ToString().Substring(0, 10)
            End If
        End If
        Return resultDate

    End Function

End Module
