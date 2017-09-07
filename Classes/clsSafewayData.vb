Imports System.Data.OleDb.OleDbConnection
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop.Excel
Imports System.Configuration

Public Class clsSafewayData
    Dim Connection As New SqlConnection
    Dim Command As New SqlCommand
    Dim MailResults As New DirectBillImports.clsMail
    Dim sErrorTo As String = ConfigurationSettings.AppSettings("ErrorTo")

    Public Sub CloseConnections()
        If ConnectionState.Open = True Then
            Connection.Close()
            Connection.Dispose()
        End If
    End Sub

    Public Sub ClearDataFromSFWYTable()
        Try
            With Connection
                .ConnectionString = ConfigurationSettings.AppSettings("ELCID")
                .Open()
            End With
            With Command
                .Connection = Connection
                .CommandTimeout = 0
                .CommandText = "SIU_Delete_Data_SFWY_Table"
                .CommandType = CommandType.StoredProcedure
                .ExecuteNonQuery()
            End With
            Debug.WriteLine("SFWY table has been cleared.")
        Catch ex As Exception
            MailResults.EmailResults("Error: ", ex.Message & vbCrLf & ex.StackTrace, sErrorTo)
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
        Finally
            Connection.Close()
            Connection.Dispose()
        End Try
    End Sub

    Public Sub ImportFileToSFWYTable()
        Debug.WriteLine("Importing text file to AutoInvoicing database")
        Try
            Dim workSheet As New Microsoft.Office.Interop.Excel.Worksheet
            Dim workBook As Microsoft.Office.Interop.Excel.Workbook
            Dim myExcel As Microsoft.Office.Interop.Excel.Application
            Dim oCon As OleDb.OleDbConnection
            Dim sCon As String = ""
            Dim tableName As String = ConfigurationSettings.AppSettings("SafewayPath")
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
                'myExcel.Workbooks.Close()
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
                            .CommandText = "SIU_Insert_Text_File_Test"
                            .CommandType = CommandType.StoredProcedure
                            .Parameters.AddWithValue("@TrackNo", row1(0))
                            .Parameters.AddWithValue("@AccTransType", row1(1))
                            .Parameters.AddWithValue("@AgentCode", row1(2))
                            .Parameters.AddWithValue("@DateSystem", row1(3))
                            .Parameters.AddWithValue("@TransEffective", row1(4))
                            .Parameters.AddWithValue("@EffectiveDate", row1(5))
                            .Parameters.AddWithValue("@EntryType", row1(6))
                            .Parameters.AddWithValue("@PolicyNumber", Replace(row1(7), "-GA-PP-", "-"))
                            .Parameters.AddWithValue("@CLName", RTrim(row1(8)))
                            .Parameters.AddWithValue("@TransDesc", RTrim(row1(9)))
                            .Parameters.AddWithValue("@Rate", row1(10))
                            .Parameters.AddWithValue("@GrossAmount", row1(11))
                            .Parameters.AddWithValue("@Commission", row1(12))
                            .Parameters.AddWithValue("@SIUAmount", row1(13))
                            .Parameters.AddWithValue("@F15", row1(13) - row1(13))
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
                .CommandText = "DirectBillSafeway"
                .CommandType = CommandType.StoredProcedure
                BatNbr = .ExecuteReader()
            End With
            While BatNbr.Read
                ImportToStaging = BatNbr.Item(0)
            End While
            Debug.WriteLine(ImportToStaging)
            Connection.Close()
            Connection.Dispose()
            Return ImportToStaging
        Catch ex As SqlException
            MailResults.EmailResults("Error: ", ex.Message & vbCrLf & ex.StackTrace, sErrorTo)
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Function

    Public Sub UpdateAgentComm(ByVal sResult As String)
        Try
            With Connection
                .ConnectionString = ConfigurationSettings.AppSettings("CIS")
                .Open()
            End With
            With Command
                .Connection = Connection
                .CommandTimeout = 0
                .CommandText = "DirectBillUpdateSafewayAgents"
                .Parameters.AddWithValue("@BatNbr", sResult)
                .CommandType = CommandType.StoredProcedure
                .ExecuteNonQuery()
            End With
            Connection.Close()
            Connection.Dispose()
            Debug.WriteLine("")
        Catch ex As SqlException
            MailResults.EmailResults("Error: ", ex.Message & vbCrLf & ex.StackTrace, sErrorTo)
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub
End Class