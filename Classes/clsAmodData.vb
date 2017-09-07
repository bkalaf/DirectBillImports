Imports System.Data.OleDb.OleDbConnection
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop.Excel
Imports System.Configuration

Public Class clsAmodData
    Dim Connection As New SqlConnection
    Dim Command As New SqlCommand
    Dim MailResults As New DirectBillImports.clsMail
    Dim sCon As String = ConfigurationSettings.AppSettings("CIS")
    Dim sConnection As String = ConfigurationSettings.AppSettings("")
    Public Sub CloseConnections()
        If ConnectionState.Open = True Then
            Connection.Close()
            Connection.Dispose()
        End If
    End Sub

    Public Sub ClearDataFromSFWYTable()
        Try
            With Connection
                .ConnectionString = sCon
                .Open()
            End With
            With Command
                .Connection = Connection
                .CommandTimeout = 0
                .CommandText = "SIU_Delete_Data_AMod_Table"
                .CommandType = CommandType.StoredProcedure
                .ExecuteNonQuery()
            End With
            Debug.WriteLine("AMod table has been cleared.")
        Catch ex As Exception
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
        Finally
            CloseConnections()
        End Try
    End Sub

    Function ImportFileToAMODTable() As String
        Debug.WriteLine("starting Import process to Auto Invoice database.")
        Try
            Dim workSheet As New Microsoft.Office.Interop.Excel.Worksheet
            Dim workBook As Microsoft.Office.Interop.Excel.Workbook
            Dim myExcel As Microsoft.Office.Interop.Excel.Application
            Dim oCon As OleDb.OleDbConnection
            Dim sCon As String = ""
            Dim tableName As String = "\\siuins.com\siu\Process\AutoInvoicing\Amod\Import.xls"
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
                Dim sConnection As New SqlConnection(ConfigurationSettings.AppSettings("ELCID"))
                Dim sCommand As New SqlCommand
                Try
                    Debug.WriteLine("Opening connection to EL-Cid")
                    Debug.WriteLine("Inserting Policy Number: ")
                    With sCommand
                        .Connection = sConnection
                        sConnection.Open()
                        For Each row1 In dSet.Tables(0).Rows
                            .CommandText = "SIU_Insert_Text_File_Test"
                            .CommandType = CommandType.StoredProcedure
                            .Parameters.AddWithValue("@FileId", row1(0))
                            .Parameters.AddWithValue("@CompanyID", row1(1))
                            .Parameters.AddWithValue("@PolicyNumber", row1(2))
                            .Parameters.AddWithValue("@AgentID", row1(3))
                            .Parameters.AddWithValue("@LOB", row1(4))
                            .Parameters.AddWithValue("@PaymentType", row1(5))
                            .Parameters.AddWithValue("@NewRenewlnd", row1(6))
                            .Parameters.AddWithValue("@EffDate", row1(7))
                            .Parameters.AddWithValue("@ExpDate", row1(8))
                            .Parameters.AddWithValue("@InsuredLast", row1(9))
                            .Parameters.AddWithValue("@InsuredFirst", row1(10))
                            .Parameters.AddWithValue("@AgentPremiumAmount", row1(11))
                            .Parameters.AddWithValue("@AgentCommRate", row1(12))
                            .Parameters.AddWithValue("@AgentGrossComm", row1(13))
                            .Parameters.AddWithValue("@AgentNetComm", row1(14))
                            .Parameters.AddWithValue("@AgentCommUnreleased", row1(15))
                            .Parameters.AddWithValue("@SubNumber", row1(16))
                            .Parameters.AddWithValue("@SubProdName", row1(17))
                            .Parameters.AddWithValue("@SubProdPremiumAmount", row1(18))
                            .Parameters.AddWithValue("@SubProdCommRate", row1(19))
                            .Parameters.AddWithValue("@SubProdCommAmount", row1(20))
                            .ExecuteNonQuery()
                            .Parameters.Clear()
                            Debug.WriteLine(vbTab & vbTab & vbTab & Replace(row1(7), "-GA-PP-", "-"))
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

    Function StageAMOD() As String
        Try
            With Connection
                .ConnectionString = ConfigurationSettings.AppSettings("CIS")
                .Open()
            End With
            With Command
                .Connection = Connection
                .CommandTimeout = 0
                .CommandText = "DirectBillAMOD"
                .CommandType = CommandType.StoredProcedure
                StageAMOD = .ExecuteNonQuery()
            End With
            CloseConnections()
        Catch ex As Exception
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
            Return "Error: " & ex.Message & vbCrLf & ex.StackTrace
        End Try
    End Function
End Class
