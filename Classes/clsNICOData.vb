Imports System.Data.OleDb.OleDbConnection
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop.Excel
Imports System.Configuration

Class clsNICOData
    Dim Connection As New SqlConnection
    Dim Command As New SqlCommand
    Dim MailResults As New DirectBillImports.clsMail
    Public Sub CloseConnections()
        If ConnectionState.Open Then
            Connection.Close()
            Connection.Dispose()
        End If
    End Sub

    Public Sub ClearDataFromNICOTable()
        Try
            With Connection
                .ConnectionString = ConfigurationSettings.AppSettings("ELCID").ToString()
                .Open()
            End With
            With Command
                .Connection = Connection
                .CommandTimeout = 0
                .CommandText = "DirectBill_Delete_NICO"
                .CommandType = CommandType.StoredProcedure
                .ExecuteNonQuery()
            End With
            Debug.WriteLine("NICO table has been cleared.")
        Catch ex As Exception
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
        Finally
            CloseConnections()
        End Try
    End Sub

    Function ImportFileToNICOTable() As String
        Debug.WriteLine("starting Import process to SunSubmit database.")
        Try
            Dim workSheet As New Microsoft.Office.Interop.Excel.Worksheet
            Dim workBook As Microsoft.Office.Interop.Excel.Workbook
            Dim myExcel As Microsoft.Office.Interop.Excel.Application
            Dim oCon As OleDb.OleDbConnection
            Dim sCon As String = ""
            Dim tableName As String = ConfigurationSettings.AppSettings("NICOPath")
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
                            .CommandText = "DirectBill_Insert_NICO"
                            .CommandType = CommandType.StoredProcedure
                            .Parameters.AddWithValue("@PolicyId", row1(0).ToString().Replace(" ", ""))
                            .Parameters.AddWithValue("@NamedInsured", row1(1))
                            .Parameters.AddWithValue("@EffectiveDate", Convert.ToDateTime(row1(2)))
                            .Parameters.AddWithValue("@Premium", row1(3))
                            .Parameters.AddWithValue("@GrossCommPercentage", row1(4))
                            .Parameters.AddWithValue("@GrossCommAmount", row1(5))
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

    Function StageNICO() As String
        StageNICO = ""
        Dim rTrav As SqlDataReader = Nothing
        Try
            With Connection
                .ConnectionString = ConfigurationSettings.AppSettings("CIS").ToString()
                .Open()
            End With
            With Command
                .Connection = Connection
                .CommandTimeout = 0
                .CommandText = "DirectBillNICO"
                .CommandType = CommandType.StoredProcedure
                rTrav = .ExecuteReader
            End With
            While rTrav.Read
                StageNICO = rTrav.Item(0)
            End While
            CloseConnections()
        Catch ex As Exception
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
            Return "Error: " & ex.Message & vbCrLf & ex.StackTrace
        End Try
    End Function

    Public Sub UpdateCompanyIdsForNICO(ByVal sBatNbr As String)
        Try
            With Connection
                .ConnectionString = ConfigurationSettings.AppSettings("CIS").ToString()
                .Open()
            End With
            With Command
                .Connection = Connection
                .CommandTimeout = 0
                .CommandText = "NICO_UpdateCoveragInfo"
                .CommandType = CommandType.StoredProcedure
                .Parameters.AddWithValue("@BatNbr", sBatNbr)
                .ExecuteNonQuery()
            End With
            Debug.WriteLine("NICO Coverages have been updated.")
        Catch ex As Exception
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
        Finally
            CloseConnections()
        End Try
    End Sub
End Class