Imports System.Data.OleDb.OleDbConnection
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop.Excel
Imports System.Configuration

Public Class ClsOccidentalData
    Dim Connection As New SqlConnection
    Dim Command As New SqlCommand
    Dim MailResults As New DirectBillImports.clsMail
    Dim sErrorTo As String = ConfigurationSettings.AppSettings("ErrorTo")


    Public Sub ClearDataFromOccTable()
        Try
            With Connection
                .ConnectionString = ConfigurationSettings.AppSettings("ELCID")
                .Open()
            End With
            With Command
                .Connection = Connection
                .CommandTimeout = 0
                .CommandText = "SIU_Delete_Data_Occ_Table"
                .CommandType = CommandType.StoredProcedure
                .ExecuteNonQuery()
            End With
            Debug.WriteLine("Occ table has been cleared.")
        Catch ex As Exception
            MailResults.EmailResults("Error: ", ex.Message & vbCrLf & ex.StackTrace, sErrorTo)
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
        Finally
            Connection.Close()
            Connection.Dispose()
        End Try
    End Sub

    Public Sub ImportFileToOccRCPTable()
        Debug.WriteLine("Importing text file to AutoInvoicing database")
        Try
            Dim workSheet As New Microsoft.Office.Interop.Excel.Worksheet
            Dim workBook As Microsoft.Office.Interop.Excel.Workbook
            Dim myExcel As Microsoft.Office.Interop.Excel.Application
            Dim oCon As OleDb.OleDbConnection
            Dim sCon As String = ""
            Dim tableName As String = ConfigurationSettings.AppSettings("OccidentalPath")
            sCon = "Data Source=" & tableName & ";Provider=Microsoft.ACE.OLEDB.12.0; Extended Properties=""Excel 8.0; HDR=Yes;IMEX=1"""
            Try
                Dim dSet As DataSet
                Dim tTable As New Data.DataTable

                oCon = New System.Data.OleDb.OleDbConnection(sCon)
                myExcel = CreateObject("Excel.Application")
                workBook = myExcel.Workbooks.Open(tableName)
                Debug.WriteLine("Opening Excel file for import.")
                myExcel.Workbooks.Close()
                oCon.Open()

                Dim eCommand As New OleDb.OleDbDataAdapter("Select * from [" & "RPC$" & "];", oCon)
                dSet = New Data.DataSet
                eCommand.Fill(dSet)
                oCon.Close()
                eCommand.Dispose()

                Dim row1 As DataRow
                Dim sConnection As New SqlConnection(ConfigurationSettings.AppSettings("ELCID"))
                Dim sCommand As New SqlCommand
                Try
                    Debug.WriteLine("Opening connection to EL-Cid")
                    Debug.WriteLine("Inserting Policy Number: ")
                    With sCommand
                        .Connection = sConnection
                        sConnection.Open()
                        For Each row1 In dSet.Tables(0).Rows
                            .Parameters.Clear()
                            .CommandText = "SIU_Insert_Occ_RP"
                            .CommandType = CommandType.StoredProcedure
                            .Parameters.AddWithValue("@Agency", row1(0))
                            .Parameters.AddWithValue("@CoverageState", row1(1))
                            .Parameters.AddWithValue("@AcctDate", row1(2))
                            .Parameters.AddWithValue("@EntryDate", row1(3))
                            .Parameters.AddWithValue("@PolNbr", row1(4))
                            .Parameters.AddWithValue("@Module", row1(5))
                            .Parameters.AddWithValue("@MajorPeril", row1(6))
                            .Parameters.AddWithValue("@InsuredName", row1(7))
                            .Parameters.AddWithValue("@MasterCompany", row1(8))
                            .Parameters.AddWithValue("@IncpDate", row1(9))
                            .Parameters.AddWithValue("@EffDate", row1(10))
                            .Parameters.AddWithValue("@ExpDate", row1(11))
                            .Parameters.AddWithValue("@WrittenPrem", row1(12))
                            .ExecuteNonQuery()
                            .Parameters.Clear()
                            Debug.WriteLine(vbTab & vbTab & vbTab & row1(4))
                        Next
                        myExcel.Application.Workbooks.Close()
                    End With
                Catch ex As Exception
                    Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
                End Try
                Debug.WriteLine("Closing connection and disposing of command.")
                sConnection.Close()
                sCommand.Dispose()
                myExcel.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(myExcel)
            Catch ex As Exception
                MailResults.EmailResults("Error: ", ex.Message & vbCrLf & ex.StackTrace, sErrorTo)
                Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
            End Try
        Catch ex As Exception
            MailResults.EmailResults("Error: ", ex.Message & vbCrLf & ex.StackTrace, sErrorTo)
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub

    Public Sub ImportFileToOccPremiumTable()
        Debug.WriteLine("Importing text file to AutoInvoicing database")
        Try
            Dim workSheet As New Microsoft.Office.Interop.Excel.Worksheet
            Dim workBook As Microsoft.Office.Interop.Excel.Workbook
            Dim myExcel As Microsoft.Office.Interop.Excel.Application
            Dim oCon As OleDb.OleDbConnection
            Dim sCon As String = ""
            Dim tableName As String = ConfigurationSettings.AppSettings("OccidentalPath")
            sCon = "Data Source=" & tableName & ";Provider=Microsoft.ACE.OLEDB.12.0; Extended Properties=""Excel 8.0; HDR=Yes;IMEX=1"""
            Try
                Dim dSet As DataSet
                Dim tTable As New Data.DataTable

                oCon = New System.Data.OleDb.OleDbConnection(sCon)
                myExcel = CreateObject("Excel.Application")
                workBook = myExcel.Workbooks.Open(tableName)

                Debug.WriteLine("Opening Excel file for import.")
                myExcel.Workbooks.Close()
                oCon.Open()

                Dim eCommand As New OleDb.OleDbDataAdapter("Select * from [" & "Premium$" & "];", oCon)
                dSet = New Data.DataSet
                eCommand.Fill(dSet)
                oCon.Close()
                eCommand.Dispose()

                Dim row1 As DataRow
                Dim sConnection As New SqlConnection(ConfigurationSettings.AppSettings("ELCID"))
                Dim sCommand As New SqlCommand
                Try
                    Debug.WriteLine("Opening connection to EL-Cid")
                    Debug.WriteLine("Inserting Policy Number: ")
                    With sCommand
                        .Connection = sConnection
                        sConnection.Open()
                        For Each row1 In dSet.Tables(0).Rows
                            .CommandText = "SIU_Insert_Occ"
                            .CommandType = CommandType.StoredProcedure
                            .Parameters.AddWithValue("@AccountingDate", row1(0))
                            .Parameters.AddWithValue("@BusinessUnit", row1(1))
                            .Parameters.AddWithValue("@EntryDate", row1(2))
                            .Parameters.AddWithValue("@PolicyNumber", row1(3))
                            .Parameters.AddWithValue("@Module", row1(4))
                            .Parameters.AddWithValue("@InsuredName", row1(5))
                            .Parameters.AddWithValue("@MasterCompany", row1(6))
                            .Parameters.AddWithValue("@CoverageState", RTrim(row1(7)))
                            .Parameters.AddWithValue("@Agency", RTrim(row1(8)))
                            .Parameters.AddWithValue("@OriginalInceptionDate", row1(9))
                            .Parameters.AddWithValue("@PolicyEffectiveDate", row1(10))
                            .Parameters.AddWithValue("@PolicyExpirationDate", row1(11))
                            .Parameters.AddWithValue("@TypeActivity", row1(12))
                            .Parameters.AddWithValue("@WrittenPremium", row1(13))
                            .ExecuteNonQuery()
                            .Parameters.Clear()
                        Next
                        myExcel.Application.Workbooks.Close()
                    End With
                Catch ex As Exception
                    Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
                End Try
                Debug.WriteLine("Closing connection and disposing of command.")
                sConnection.Close()
                sCommand.Dispose()
                myExcel.Quit()
                System.Runtime.InteropServices.Marshal.ReleaseComObject(myExcel)
            Catch ex As Exception
                MailResults.EmailResults("Error: ", ex.Message & vbCrLf & ex.StackTrace, sErrorTo)
                Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
            End Try
        Catch ex As Exception
            MailResults.EmailResults("Error: ", ex.Message & vbCrLf & ex.StackTrace, sErrorTo)
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub

    Public Function ImportToStaging() As String
        ImportToStaging = ""
        Dim BatNbr As SqlDataReader = Nothing
        Try
            With Connection
                .ConnectionString = ConfigurationSettings.AppSettings("CIS")
                .Open()
            End With
            With Command
                .Connection = Connection
                .CommandTimeout = 0
                .CommandText = "DirectBillOccidental"
                .CommandType = CommandType.StoredProcedure
                BatNbr = .ExecuteReader()
            End With
            While BatNbr.Read
                ImportToStaging = BatNbr.Item(0)
            End While
            Connection.Close()
            Connection.Dispose()
            Debug.WriteLine(ImportToStaging)
            Return ImportToStaging
        Catch ex As SqlException
            MailResults.EmailResults("Error: ", ex.Message & vbCrLf & ex.StackTrace, sErrorTo)
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Function

    Public Sub PreProcess()

        Try
            With Connection
                .ConnectionString = ConfigurationSettings.AppSettings("ELCID")
                .Open()
            End With
            With Command
                .Connection = Connection
                .CommandTimeout = 0
                .CommandText = "DirectBillOccPreProc1"
                .CommandType = CommandType.StoredProcedure
                .ExecuteNonQuery()
            End With
            Connection.Close()
            Connection.Dispose()
        Catch ex As SqlException
            MailResults.EmailResults("Error: ", ex.Message & vbCrLf & ex.StackTrace, sErrorTo)
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
        End Try

        Try
            With Connection
                .ConnectionString = ConfigurationSettings.AppSettings("ELCID")
                .Open()
            End With
            With Command
                .Connection = Connection
                .CommandTimeout = 0
                .CommandText = "DirectBillOccPreProc2"
                .CommandType = CommandType.StoredProcedure
                .ExecuteNonQuery()
            End With
            Connection.Close()
            Connection.Dispose()
        Catch ex As SqlException
            MailResults.EmailResults("Error: ", ex.Message & vbCrLf & ex.StackTrace, sErrorTo)
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
        End Try

    End Sub

End Class