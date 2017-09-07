Imports System
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop.Excel
Imports System.Configuration

Public Class clsGeoveraData
    Dim Connection As New SqlConnection
    Dim Command As New SqlCommand
    Dim MailResults As New DirectBillImports.clsMail
    Dim sErrorTo As String = ConfigurationSettings.AppSettings("ErrorTo")

    Public Sub CloseConn()
        If ConnectionState.Open Then
            Connection.Close()
            Connection.Dispose()
        End If
    End Sub

    Sub ClearUSFG()
        Try
            With Connection
                .ConnectionString = ConfigurationSettings.AppSettings("ELCID")
                .Open()
            End With
            With Command
                .Connection = Connection
                .CommandTimeout = 0
                .CommandText = "P_ClearUSFG"
                .CommandType = CommandType.StoredProcedure
                .ExecuteNonQuery()
            End With
            Debug.WriteLine("Geovera table has been cleared.")
        Catch ex As Exception
            MailResults.EmailResults("Error: ", ex.Message & vbCrLf & ex.StackTrace, sErrorTo)
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
        Finally
            CloseConn()
        End Try
    End Sub

    Sub ImportFileToGeoveraFFB()
        Debug.WriteLine("Importing text file to AutoInvoicing database")
        Try
            Dim workSheet As New Microsoft.Office.Interop.Excel.Worksheet
            Dim workBook As Microsoft.Office.Interop.Excel.Workbook
            Dim myExcel As Microsoft.Office.Interop.Excel.Application
            Dim oCon As OleDb.OleDbConnection
            Dim sCon As String = ""
            Dim tableName As String = ConfigurationSettings.AppSettings("GeoveraFFBPath")
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

                Dim eCommand As New OleDb.OleDbDataAdapter("Select * from [" & tableName & "];", oCon)
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
                            .CommandText = "SIU_Insert_GeoveraFFB"
                            .CommandType = CommandType.StoredProcedure
                            .Parameters.AddWithValue("@DatePosted", row1.Item(0))
                            .Parameters.AddWithValue("@MarketingRep", row1.Item(1))
                            .Parameters.AddWithValue("@AgentId", row1.Item(2))
                            .Parameters.AddWithValue("@PolicyNbr", row1.Item(3))
                            .Parameters.AddWithValue("@InsuredName", row1.Item(4))
                            .Parameters.AddWithValue("@TTCode", row1.Item(5))
                            .Parameters.AddWithValue("@TTExtended", row1.Item(6))
                            .Parameters.AddWithValue("@EffDate", row1.Item(7))
                            .Parameters.AddWithValue("@ExpDate", row1.Item(8))
                            .Parameters.AddWithValue("@WrittenPremium", row1.Item(9))
                            .Parameters.AddWithValue("@NetPremium", row1.Item(10))
                            .Parameters.AddWithValue("@StateTax", row1.Item(11))
                            .Parameters.AddWithValue("@StampFee", row1.Item(12))
                            .Parameters.AddWithValue("@CPICFee", row1.Item(13))
                            .Parameters.AddWithValue("@UWFee", row1.Item(14))
                            .Parameters.AddWithValue("@PolicyFee", row1.Item(15))
                            .Parameters.AddWithValue("@FHCFFee", row1.Item(16))
                            .Parameters.AddWithValue("@SrvcCharge", row1.Item(17))
                            .Parameters.AddWithValue("@Refund", row1.Item(18))
                            .Parameters.AddWithValue("@WriteOff", row1.Item(19))
                            .Parameters.AddWithValue("@AgentRetained", row1.Item(20))
                            .Parameters.AddWithValue("@CashEntered", row1.Item(21))
                            .Parameters.AddWithValue("@PendingRefund", row1.Item(22))
                            .Parameters.AddWithValue("@BeginningBalance", row1.Item(23))
                            .Parameters.AddWithValue("@EndingBalance", row1.Item(24))
                            .Parameters.AddWithValue("@Address", row1.Item(25))
                            .Parameters.AddWithValue("@City", row1.Item(26))
                            .Parameters.AddWithValue("@State", row1.Item(27))
                            .Parameters.AddWithValue("@Zip5", row1.Item(28))
                            .Parameters.AddWithValue("@Zip4", row1.Item(29))
                            .Parameters.AddWithValue("@CovA", row1.Item(30))
                            .Parameters.AddWithValue("@APDed", row1.Item(31))
                            .Parameters.AddWithValue("@WindDed", row1.Item(32))
                            .Parameters.AddWithValue("@InspectionFee", row1.Item(33))
                            .Parameters.AddWithValue("@Coverage", row1.Item(34))
                            .ExecuteNonQuery()
                            .Parameters.Clear()
                            Debug.WriteLine(vbTab & vbTab & vbTab & Replace(row1(7), "-GA-PP-", "-"))
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

    Sub ImportFileToGeovera()
        Debug.WriteLine("Importing text file to AutoInvoicing database")
        Try
            Dim workSheet As New Microsoft.Office.Interop.Excel.Worksheet
            Dim workBook As Microsoft.Office.Interop.Excel.Workbook
            Dim myExcel As Microsoft.Office.Interop.Excel.Application
            Dim oCon As OleDb.OleDbConnection
            Dim sCon As String = ""
            Dim tableName As String = ConfigurationSettings.AppSettings("GeoveraNormalPath")
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
                Dim eCommand As New OleDb.OleDbDataAdapter("Select * from [" & tableName & "];", oCon)
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
                            .CommandText = "SIU_Insert_GeoveraNormal"
                            .CommandType = CommandType.StoredProcedure
                            .Parameters.AddWithValue("@DatePosted", row1.Item(0))
                            .Parameters.AddWithValue("@MarketingRep", row1.Item(1))
                            .Parameters.AddWithValue("@AgentId", row1.Item(2))
                            .Parameters.AddWithValue("@PolicyNbr", row1.Item(3))
                            .Parameters.AddWithValue("@InsuredName", row1.Item(4))
                            .Parameters.AddWithValue("@TTCode", row1.Item(5))
                            .Parameters.AddWithValue("@TTExtended", row1.Item(6))
                            .Parameters.AddWithValue("@EffDate", row1.Item(7))
                            .Parameters.AddWithValue("@ExpDate", row1.Item(8))
                            .Parameters.AddWithValue("@WrittenPremium", row1.Item(9))
                            .Parameters.AddWithValue("@NetPremium", row1.Item(10))
                            .Parameters.AddWithValue("@StateTax", row1.Item(11))
                            .Parameters.AddWithValue("@StampFee", row1.Item(12))
                            .Parameters.AddWithValue("@CPICFee", row1.Item(13))
                            .Parameters.AddWithValue("@UWFee", row1.Item(14))
                            .Parameters.AddWithValue("@PolicyFee", row1.Item(15))
                            .Parameters.AddWithValue("@FHCFFee", row1.Item(16))
                            .Parameters.AddWithValue("@SrvcCharge", row1.Item(17))
                            .Parameters.AddWithValue("@Refund", row1.Item(18))
                            .Parameters.AddWithValue("@WriteOff", row1.Item(19))
                            .Parameters.AddWithValue("@AgentRetained", row1.Item(20))
                            .Parameters.AddWithValue("@CashEntered", row1.Item(21))
                            .Parameters.AddWithValue("@PendingRefund", row1.Item(22))
                            .Parameters.AddWithValue("@BeginningBalance", row1.Item(23))
                            .Parameters.AddWithValue("@EndingBalance", row1.Item(24))
                            .Parameters.AddWithValue("@Address", row1.Item(25))
                            .Parameters.AddWithValue("@City", row1.Item(26))
                            .Parameters.AddWithValue("@State", row1.Item(27))
                            .Parameters.AddWithValue("@Zip5", row1.Item(28))
                            .Parameters.AddWithValue("@Zip4", row1.Item(29))
                            .Parameters.AddWithValue("@CovA", row1.Item(30))
                            .Parameters.AddWithValue("@APDed", row1.Item(31))
                            .Parameters.AddWithValue("@WindDed", row1.Item(32))
                            .Parameters.AddWithValue("@InspectionFee", row1.Item(33))
                            .Parameters.AddWithValue("@Coverage", row1.Item(34))
                            .ExecuteNonQuery()
                            .Parameters.Clear()
                            Debug.WriteLine(vbTab & vbTab & vbTab & row1(3))
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

    Function StageGeoveraNormal() As String
        StageGeoveraNormal = ""
        Dim BatNbr As SqlDataReader = Nothing
        Try
            With Connection
                .ConnectionString = ConfigurationSettings.AppSettings("CIS")
                .Open()
            End With
            With Command
                .Connection = Connection
                .CommandTimeout = 0
                .CommandText = "DirectBillUSFGNormal"
                .CommandType = CommandType.StoredProcedure
                BatNbr = .ExecuteReader()
            End With
            While BatNbr.Read
                StageGeoveraNormal = BatNbr.Item(0)
            End While
            Debug.WriteLine(StageGeoveraNormal)
            Return StageGeoveraNormal
            CloseConn()
        Catch ex As SqlException
            MailResults.EmailResults("Error: ", ex.Message & vbCrLf & ex.StackTrace, sErrorTo)
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Function

    Function StageGeoveraFFB() As String
        StageGeoveraFFB = ""
        Dim BatNbr As SqlDataReader = Nothing
        Try
            With Connection
                .ConnectionString = ConfigurationSettings.AppSettings("CIS")
                .Open()
            End With
            With Command
                .Connection = Connection
                .CommandTimeout = 0
                .CommandText = "DirectBillUSFGFBB"
                .CommandType = CommandType.StoredProcedure
                BatNbr = .ExecuteReader()
            End With
            While BatNbr.Read
                StageGeoveraFFB = BatNbr.Item(0)
            End While
            Debug.WriteLine(StageGeoveraFFB)
            CloseConn()
            Return StageGeoveraFFB
        Catch ex As SqlException
            MailResults.EmailResults("Error: ", ex.Message & vbCrLf & ex.StackTrace, sErrorTo)
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Function
End Class