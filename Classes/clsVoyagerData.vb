Imports System.Data.SqlClient
Imports System.Configuration
Imports Microsoft.Office.Interop.Excel
Imports System.Data.OleDb.OleDbConnection

Public Class clsVoyagerData

    Dim Connection As New SqlConnection
    Dim Command As New SqlCommand
    Dim MailResults As New DirectBillImports.clsMail
    Dim ELCid As String = ConfigurationSettings.AppSettings("ELCID")
    Dim sCIS As String = ConfigurationSettings.AppSettings("CIS")
    Dim sVoyagerPath As String = ConfigurationSettings.AppSettings("VoyagerPath")
    Dim sTo As String = ConfigurationSettings.AppSettings("DBTo")
    Dim sErrorTo As String = ConfigurationSettings.AppSettings("ErrorTo")

    Public Sub CloseConnection()
        If ConnectionState.Open = True Then
            Connection.Close()
            Connection.Dispose()
        End If
    End Sub

    Public Sub ClearData()
        Try
            With Connection
                .ConnectionString = ELCid
                .Open()
            End With
            With Command
                .Connection = Connection
                .CommandTimeout = 0
                .CommandText = "SIU_Delete_Data_Voyager_Table"
                .CommandType = CommandType.StoredProcedure
                .ExecuteNonQuery()
            End With
            Debug.WriteLine("VOY table has been cleared.")
        Catch ex As Exception
            MailResults.EmailResults("Error: ", ex.Message & vbCrLf & ex.StackTrace, sErrorTo)
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
        Finally
            CloseConnection()
        End Try
    End Sub

    Public Function ImportFileToVoyagerTable() As String
        ClearData()
        Debug.WriteLine("starting Import process to Auto Invoice database.")
        Try
            Dim workSheet As New Microsoft.Office.Interop.Excel.Worksheet
            Dim workBook As Microsoft.Office.Interop.Excel.Workbook
            Dim myExcel As Microsoft.Office.Interop.Excel.Application
            Dim oCon As OleDb.OleDbConnection
            Dim sCon As String = ""
            Dim tableName As String = sVoyagerPath
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
                        .Connection = sConnection 'sConnection
                        sConnection.Open()
                        For Each row1 In dSet.Tables(0).Rows
                            .CommandText = "SIU_Insert_Voyager"
                            .CommandType = CommandType.StoredProcedure
                            .Parameters.AddWithValue("@PolicyNbr", row1(0))
                            .Parameters.AddWithValue("@InsuredName", row1(1))
                            .Parameters.AddWithValue("@SUB", row1(2))
                            .Parameters.AddWithValue("@TRANSACTION", row1(3))
                            .Parameters.AddWithValue("@ST", row1(4))
                            .Parameters.AddWithValue("@LOB", row1(5))
                            .Parameters.AddWithValue("@EFFDATE", row1(6))
                            .Parameters.AddWithValue("@RATE", row1(7))
                            .Parameters.AddWithValue("@PREMIUM", row1(8))
                            .Parameters.AddWithValue("@COMM$", row1(9))
                            .ExecuteNonQuery()
                            .Parameters.Clear()
                            Debug.WriteLine(vbTab & vbTab & vbTab & Replace(row1(7), "-GA-PP-", "-"))
                        Next
                        myExcel.Application.Workbooks.Close()
                        sConnection.Close()
                        sCommand.Dispose()
                    End With
                Catch ex As Exception
                    sConnection.Close()
                    sCommand.Dispose()
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
                Return "Error: " & ex.Message & vbCrLf & ex.StackTrace
            End Try
            Return "Finished"
        Catch ex As Exception
            MailResults.EmailResults("Error: ", ex.Message & vbCrLf & ex.StackTrace, sErrorTo)
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
            Return "Error: " & ex.Message & vbCrLf & ex.StackTrace
        End Try
    End Function

    Public Function PreProcess() As String
        If ConnectionState.Open Then
            Connection.Close()
        End If
        Try
            With Connection
                .ConnectionString = ELCid
                .Open()
            End With
            With Command
                .Connection = Connection
                .CommandTimeout = 0
                .CommandText = "SIU_Voyager_PreProcess"
                .CommandType = CommandType.StoredProcedure
                .ExecuteNonQuery()
            End With
            Return "Finished"
            CloseConnection()
        Catch ex As Exception
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
            MailResults.EmailResults("Error: ", ex.Message & vbCrLf & ex.StackTrace, sErrorTo)
            Return "Error: " & ex.Message & vbCrLf & ex.StackTrace
        End Try
    End Function

    Public Function StageVoyager() As String
        If ConnectionState.Open Then
            Connection.Close()
        End If
        Dim rStageVoyager As SqlDataReader = Nothing
        StageVoyager = ""
        Try
            With Connection
                .ConnectionString = sCIS 'ConfigurationSettings.AppSettings("CIS")
                .Open()
            End With
            With Command
                .Connection = Connection
                .CommandTimeout = 0
                .CommandText = "DirectbillVoyager"
                .CommandType = CommandType.StoredProcedure
                rStageVoyager = .ExecuteReader
            End With
            While rStageVoyager.Read
                StageVoyager = rStageVoyager.Item(0)
            End While
            Return StageVoyager
            CloseConnection()
        Catch ex As Exception
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
            MailResults.EmailResults("Error: ", ex.Message & vbCrLf & ex.StackTrace, sErrorTo)
            StageVoyager = "Error: " & ex.Message & ex.StackTrace
        End Try
    End Function

End Class
