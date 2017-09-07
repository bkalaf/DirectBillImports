Imports System
Imports System.Data.SqlClient
Imports System.Configuration

Public Class clsDataCommon

    Dim sConn As String = ConfigurationSettings.AppSettings("CIS")
    Dim Conn As New SqlConnection
    Dim com As New SqlCommand
    Dim MailResponse As New clsMail
    Dim sErrorTo As String = ConfigurationSettings.AppSettings("ErrorTo")

    Public Sub ConnClose()
        If ConnectionState.Open Then
            Conn.Close()
            Conn.Dispose()
        End If
    End Sub

    Public Sub DirectBillHistoryInsert(ByVal sCarrier As String)
        Try
            With Conn
                .ConnectionString = sConn
                .Open()
            End With
            With com
                .Connection = Conn
                .CommandTimeout = 0
                .CommandText = "DirectBillHistoryInsert"
                .CommandType = CommandType.StoredProcedure
                .Parameters.AddWithValue("@Carrier", sCarrier)
                .Parameters.AddWithValue("@NbrOfPolicies", "")
                .Parameters.AddWithValue("@StartTime", Date.Now.ToString("MM/dd/yyyy hh:mm.ss"))
                .Parameters.AddWithValue("@FinishTime", "")
                .Parameters.AddWithValue("@NetPremImport", "")
                .Parameters.AddWithValue("@Month", Date.Now.Month)
                .Parameters.AddWithValue("@Year", Date.Now.Year)
                .ExecuteNonQuery()
            End With
            ConnClose()
        Catch ex As Exception
            MailResponse.EmailResults("Error", ex.Message & vbCrLf & ex.StackTrace, sErrorTo)
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Sub

    Public Function GetNumbers(ByVal sBatNbr As String) As SqlDataReader
        GetNumbers = Nothing
        Try
            With Conn
                .ConnectionString = sConn
                .Open()
            End With
            With com
                .Connection = Conn
                .CommandTimeout = 0
                .CommandText = "SIU_GetImportNumbers"
                .CommandType = CommandType.StoredProcedure
                .Parameters.AddWithValue("@BatchId", sBatNbr)
                GetNumbers = .ExecuteReader
            End With
        Catch ex As Exception
            MailResponse.EmailResults("Error", ex.Message & vbCrLf & ex.StackTrace, sErrorTo)
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
        End Try
        Return GetNumbers
    End Function

    Public Function GetPolicyCount(ByVal sBatNbr As String) As String
        Dim sqlRead As SqlDataReader = Nothing
        GetPolicyCount = ""
        Try
            With Conn
                .ConnectionString = sConn
                .Open()
            End With
            With com
                .Parameters.Clear()
                .Connection = Conn
                .CommandTimeout = 0
                .CommandText = "SIU_GetPolcyCount"
                .CommandType = CommandType.StoredProcedure
                .Parameters.AddWithValue("@BatId", sBatNbr)
                sqlRead = .ExecuteReader()
            End With
            While sqlRead.Read
                GetPolicyCount = sqlRead.Item("PolicyCount")
            End While
            ConnClose()
            Return GetPolicyCount
        Catch ex As Exception
            MailResponse.EmailResults("Error", ex.Message & vbCrLf & ex.StackTrace, sErrorTo)
            Debug.WriteLine(ex.Message & vbCrLf & ex.StackTrace)
        End Try
    End Function

End Class