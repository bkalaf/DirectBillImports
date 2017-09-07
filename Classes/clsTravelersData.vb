Imports System.Data.OleDb.OleDbConnection
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop.Excel
Imports System.Configuration

Public Class clsTravelersData
    Dim Connection As New SqlConnection
    Dim Command As New SqlCommand
    Dim MailResults As New DirectBillImports.clsMail

    Public Sub CloseConnections()
        If ConnectionState.Open Then
            Connection.Close()
            Connection.Dispose()
        End If
    End Sub

    Public Sub ClearDataFromTravelersTable()
        Try
            With Connection
                .ConnectionString = ConfigurationSettings.AppSettings("ELCID")
                .Open()
            End With
            With Command
                .Connection = Connection
                .CommandTimeout = 0
                .CommandText = "SIU_Delete_Data_Travelers_Table"
                .CommandType = CommandType.StoredProcedure
                .ExecuteNonQuery()
            End With
            Debug.WriteLine("Travelers table has been cleared.")
        Catch ex As Exception
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
        Finally
            CloseConnections()
        End Try
    End Sub

    Function ImportFileToTravelersTable() As String
        Debug.WriteLine("starting Import process to Auto Invoice database.")
        Try
            Dim workSheet As New Microsoft.Office.Interop.Excel.Worksheet
            Dim workBook As Microsoft.Office.Interop.Excel.Workbook
            Dim myExcel As Microsoft.Office.Interop.Excel.Application
            Dim oCon As OleDb.OleDbConnection
            Dim sCon As String = ""
            Dim tableName As String = ConfigurationSettings.AppSettings("TravelersPath")
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
                            .CommandText = "SIU_Insert_Travelers"
                            .CommandType = CommandType.StoredProcedure
                            .Parameters.AddWithValue("@STATEMENT", row1(0))
                            .Parameters.AddWithValue("@OFF", row1(1))
                            .Parameters.AddWithValue("@AGENT", row1(2))
                            .Parameters.AddWithValue("@SEC", row1(3))
                            .Parameters.AddWithValue("@SUB", row1(4))
                            .Parameters.AddWithValue("@NAMEOFINSURED", row1(5))
                            .Parameters.AddWithValue("@POLICYNUMBER", row1(6))
                            .Parameters.AddWithValue("@PAYMENT", row1(7))
                            .Parameters.AddWithValue("@PAID", row1(8))
                            .Parameters.AddWithValue("@COMM", row1(9))
                            .Parameters.AddWithValue("@POLICYTERM", row1(10))
                            .Parameters.AddWithValue("@PREMIUMFOR", row1(11))
                            .Parameters.AddWithValue("@POLEFFDT", row1(12))
                            .Parameters.AddWithValue("@PMT", row1(13))
                            .Parameters.AddWithValue("@POLHOLDERS", row1(14))
                            .Parameters.AddWithValue("@TRANSACTION", row1(15))
                            .Parameters.AddWithValue("@TOTALCOMMISSION", row1(16))
                            .Parameters.AddWithValue("@PSO", row1(17))
                            .Parameters.AddWithValue("@SECONDARYAGENT", row1(18))
                            .Parameters.AddWithValue("@SUBAGENT", row1(19))
                            .Parameters.AddWithValue("@TOTALCOMMISSION1", row1(20))
                            .Parameters.AddWithValue("@CONTROLLINGAGENT", row1(21))
                            .Parameters.AddWithValue("@COMBINEDCURRENT", row1(22))
                            .Parameters.AddWithValue("@SHAREDREPORT", row1(23))
                            .Parameters.AddWithValue("@RSANB", row1(24))
                            .Parameters.AddWithValue("@TOTALCOMMISSION2", row1(25))
                            .Parameters.AddWithValue("@TOTALCOMMISSION3", row1(26))
                            .Parameters.AddWithValue("@PRIORMONTHTOTAL", row1(27))
                            .Parameters.AddWithValue("@PRIORMONTHTOTAL1", row1(28))
                            .Parameters.AddWithValue("@PRIORMONTHTOTAL2", row1(29))
                            .ExecuteNonQuery()
                            .Parameters.Clear()
                            Debug.WriteLine(vbTab & vbTab & vbTab & row1(6))
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

    Function StageTravelers() As String
        StageTravelers = ""
        Dim rTrav As SqlDataReader = Nothing
        Try
            With Connection
                .ConnectionString = ConfigurationSettings.AppSettings("CIS")
                .Open()
            End With
            With Command
                .Connection = Connection
                .CommandTimeout = 0
                .CommandText = "DirectBillTravelers"
                .CommandType = CommandType.StoredProcedure
                rTrav = .ExecuteReader
            End With
            While rTrav.Read
                StageTravelers = rTrav.Item(0)
            End While
            CloseConnections()
        Catch ex As Exception
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
            Return "Error: " & ex.Message & vbCrLf & ex.StackTrace
        End Try
    End Function
End Class
