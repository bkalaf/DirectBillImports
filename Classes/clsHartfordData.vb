Imports System.Data.OleDb.OleDbConnection
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop.Excel
Imports System.Configuration

Class clsHartfordData
    Dim Connection As New SqlConnection
    Dim Command As New SqlCommand
    Dim MailResults As New DirectBillImports.clsMail
    Public Sub CloseConnections()
        If ConnectionState.Open Then
            Connection.Close()
            Connection.Dispose()
        End If
    End Sub

    Public Sub ClearDataFromHartfordTable()
        Try
            With Connection
                .ConnectionString = ConfigurationSettings.AppSettings("SunSubmit").ToString()
                .Open()
            End With
            With Command
                .Connection = Connection
                .CommandTimeout = 0
                .CommandText = "DirectBill_Delete_Hartford"
                .CommandType = CommandType.StoredProcedure
                .ExecuteNonQuery()
            End With
            Debug.WriteLine("Hartford table has been cleared.")
        Catch ex As Exception
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
        Finally
            CloseConnections()
        End Try
    End Sub

    Function ImportFileToHartfordTable() As String
        Debug.WriteLine("starting Import process to SunSubmit database.")
        Try
            Dim workSheet As New Microsoft.Office.Interop.Excel.Worksheet
            Dim workBook As Microsoft.Office.Interop.Excel.Workbook
            Dim myExcel As Microsoft.Office.Interop.Excel.Application
            Dim oCon As OleDb.OleDbConnection
            Dim sCon As String = ""
            Dim tableName As String = ConfigurationSettings.AppSettings("HartfordPath")
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
                Dim sConnection As New SqlConnection(ConfigurationSettings.AppSettings("SunSubmit").ToString())
                Dim sCommand As New SqlCommand
                Try
                    Debug.WriteLine("Opening connection to EL-Cid")
                    Debug.WriteLine("Inserting Policy Number: ")
                    With sCommand
                        .Connection = sConnection
                        sConnection.Open()
                        For Each row1 In dSet.Tables(0).Rows
                            .CommandText = "DirectBill_Insert_Hartford"
                            .CommandType = CommandType.StoredProcedure
                            .Parameters.AddWithValue("@AGENTCODE", row1(0).ToString().PadLeft(6, "0"))
                            .Parameters.AddWithValue("@POLICYNUMBER", row1(1).ToString().Replace(" ", ""))
                            .Parameters.AddWithValue("@LOB", row1(2))
                            .Parameters.AddWithValue("@NAMEDINSURED", row1(3))
                            .Parameters.AddWithValue("@INSUREDADDRESS", row1(4))
                            .Parameters.AddWithValue("@INSUREDCITY", row1(5))
                            .Parameters.AddWithValue("@INSUREDSTATE", row1(6))
                            .Parameters.AddWithValue("@INSUREDZIP", row1(7))
                            .Parameters.AddWithValue("@TRANSACTIONTYPE", row1(8))
                            .Parameters.AddWithValue("@POLICYEFFDATE", Convert.ToDateTime(row1(9)))
                            .Parameters.AddWithValue("@POLICYEXPDATE", Convert.ToDateTime(row1(10)))
                            .Parameters.AddWithValue("@TRANSACTIONEFFDATE", Convert.ToDateTime(row1(11)))
                            .Parameters.AddWithValue("@PREMIUM", row1(12))
                            .Parameters.AddWithValue("@GROSSCOMMRATE", row1(13))
                            .Parameters.AddWithValue("@GROSSCOMMAMOUNT", row1(14))
                            .Parameters.AddWithValue("@AGENTCOMMRATE", row1(15))
                            .Parameters.AddWithValue("@DBA", row1(16))
                            .Parameters.AddWithValue("@UNDERWRITER", row1(17))
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

    Function StageHartford() As String
        StageHartford = ""
        Dim rTrav As SqlDataReader = Nothing
        Try
            With Connection
                .ConnectionString = ConfigurationSettings.AppSettings("CIS").ToString()
                .Open()
            End With
            With Command
                .Connection = Connection
                .CommandTimeout = 0
                .CommandText = "DirectBillHartford"
                .CommandType = CommandType.StoredProcedure
                rTrav = .ExecuteReader
            End With
            While rTrav.Read
                StageHartford = rTrav.Item(0)
            End While
            CloseConnections()
        Catch ex As Exception
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
            Return "Error: " & ex.Message & vbCrLf & ex.StackTrace
        End Try
    End Function

End Class