Imports System.Data.OleDb.OleDbConnection
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop.Excel
Imports System.Configuration

Class clsSafewayReconciliationData
    Dim Connection As New SqlConnection
    Dim Command As New SqlCommand
    Dim MailResults As New DirectBillImports.clsMail
    Public Sub CloseConnections()
        If ConnectionState.Open Then
            Connection.Close()
            Connection.Dispose()
        End If
    End Sub

    Public Sub ClearDataFromSafewayReconciliationTable()
        Try
            With Connection
                .ConnectionString = ConfigurationSettings.AppSettings("ELCID").ToString()
                .Open()
            End With
            With Command
                .Connection = Connection
                .CommandTimeout = 0
                .CommandText = "SIU_Delete_Data_SafewayReconciliation_Table"
                .CommandType = CommandType.StoredProcedure
                .ExecuteNonQuery()
            End With
            Debug.WriteLine("SafewayReconciliation table has been cleared.")
        Catch ex As Exception
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
        Finally
            CloseConnections()
        End Try
    End Sub

    Function ImportFileToSafewayReconciliationTable() As String
        Debug.WriteLine("starting Import process to ELCID database.")
        Try
            Dim workSheet As New Microsoft.Office.Interop.Excel.Worksheet
            Dim workBook As Microsoft.Office.Interop.Excel.Workbook
            Dim myExcel As Microsoft.Office.Interop.Excel.Application
            Dim oCon As OleDb.OleDbConnection
            Dim sCon As String = ""
            Dim tableName As String = ConfigurationSettings.AppSettings("SafewayReconciliationPath")
            sCon = "Data Source=" & tableName & ";Provider=Microsoft.ACE.OLEDB.12.0; Extended Properties=""Excel 8.0; HDR=Yes;IMEX=1"""
            Try
                Dim dSet As DataSet
                Dim tTable As New Data.DataTable

                oCon = New System.Data.OleDb.OleDbConnection(sCon)
                myExcel = CreateObject("Excel.Application")
                workBook = myExcel.Workbooks.Open(tableName)

                For Each workSheet In workBook.Worksheets
                    tableName = workSheet.Name.ToString + "$"
                Next
                Debug.WriteLine("Opening Excel file for import.")
                myExcel.Workbooks.Close()
                oCon.Open()

                Debug.WriteLine("Selecting data from excel spreadsheet")

                Dim eCommand As New OleDb.OleDbDataAdapter("Select * from [" & tableName & "];", oCon)
                dSet = New Data.DataSet
                eCommand.Fill(dSet)
                oCon.Close()
                eCommand.Dispose()

                Dim row1 As DataRow
                Dim sConnection As New SqlConnection(ConfigurationSettings.AppSettings("ELCID").ToString())
                Dim sCommand As New SqlCommand
                Try
                    Debug.WriteLine("Opening connection to EL-Cid")
                    Debug.WriteLine("Inserting Policy Number: ")
                    With sCommand
                        .Connection = sConnection
                        sConnection.Open()
                        For Each row1 In dSet.Tables(0).Rows
                            .CommandText = "SIU_Insert_SafewayReconciliation"
                            .CommandType = CommandType.StoredProcedure
                            .Parameters.AddWithValue("@PolicyNumber", row1(0).ToString().Replace(" ", ""))
                            .Parameters.AddWithValue("@PolicyRecordSeq", row1(1))
                            .Parameters.AddWithValue("@TransactionType", row1(2))
                            .Parameters.AddWithValue("@BillType", row1(3))
                            .Parameters.AddWithValue("@AgentCode", GetAgent(row1(4).ToString()))
                            .Parameters.AddWithValue("@TransactionEffectiveDate", Convert.ToDateTime(row1(5)))
                            .Parameters.AddWithValue("@Premium", row1(6))
                            .Parameters.AddWithValue("@FirstName", row1(7))
                            .Parameters.AddWithValue("@LastName", row1(8))
                            .Parameters.AddWithValue("@Address", row1(9))
                            .Parameters.AddWithValue("@City", row1(10))
                            .Parameters.AddWithValue("@State", row1(11))
                            .Parameters.AddWithValue("@Zip", row1(12))
                            .Parameters.AddWithValue("@AgntComm", row1(14))
                            .Parameters.AddWithValue("@PaidToSafeway", row1(15))
                            .Parameters.AddWithValue("@GrossComm", row1(16))
                            .Parameters.AddWithValue("@GrossCommRate", row1(17))
                            .Parameters.AddWithValue("@AgentCommRate", row1(13))
                            .ExecuteNonQuery()
                            .Parameters.Clear()
                        Next
                        myExcel.Application.Workbooks.Close()
                    End With
                Catch ex As Exception
                    Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
                    Return "Error: " & ex.Message & vbCrLf & ex.StackTrace
                End Try
                Debug.WriteLine("Closing connection and disposing of command.")
                sConnection.Close()
                sCommand.Dispose()
                myExcel.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(myExcel)
            Catch ex As Exception
                Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
                Return "Error: " & ex.Message & vbCrLf & ex.StackTrace
            End Try
            Return "Finished"
        Catch ex As Exception
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
            Return "Error: " & ex.Message & vbCrLf & ex.StackTrace
        End Try
    End Function

    'Add leading zeroes to make the size always 6 digit for agent id
    'If it starts with 003, change it to 023 - for safeway only
    Private Function GetAgent(ByVal agentId As String) As String
        agentId = agentId.PadLeft(6, "0")
        If agentId.StartsWith("003") Then
            agentId = "02" + agentId.Substring(2, 4)
        End If
        Return agentId
    End Function

    Function StageSafewayReconciliation() As String
        StageSafewayReconciliation = ""
        Dim rTrav As SqlDataReader = Nothing
        Try
            With Connection
                .ConnectionString = ConfigurationSettings.AppSettings("CIS").ToString()
                .Open()
            End With
            With Command
                .Connection = Connection
                .CommandTimeout = 0
                .CommandText = "ProcessSafewayReconciliation"
                .CommandType = CommandType.StoredProcedure
                rTrav = .ExecuteReader
            End With
            While rTrav.Read
                StageSafewayReconciliation = rTrav.Item(0)
            End While
            CloseConnections()
        Catch ex As Exception
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
            Return "Error: " & ex.Message & vbCrLf & ex.StackTrace
        End Try
    End Function

End Class