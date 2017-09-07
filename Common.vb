Imports System.Data.SqlClient
Imports System.Configuration
Public Class Common
    Public Shared ReadOnly Property AIMConnectionString(ByVal pPorT As String) As String
        Get
            AIMConnectionString = ""
            Select Case UCase(pPorT)
                Case "P"
                    AIMConnectionString = ConfigurationSettings.AppSettings("CIS")
                Case "T"
                    AIMConnectionString = "Data Source=testcisaim;Initial Catalog=CIS;User Id=sa;Password=testcisaim;"
            End Select
            Return AIMConnectionString
        End Get
    End Property
    Public Shared Function Exists(ByRef conn As SqlConnection, ByVal sSQL As String) As Boolean
        Dim comm As SqlCommand
        Dim rs As SqlDataReader

        comm = New SqlCommand(sSQL, conn)
        With comm
            .CommandTimeout = 0
            .CommandType = CommandType.Text
            rs = .ExecuteReader
            .Dispose()
        End With
        Exists = rs.Read
        rs.Close()
    End Function
    Public Shared Function GetKeyField(ByRef conn As SqlConnection, ByVal sFieldName As String) As Integer
        Dim comm As New SqlCommand("SIU_p_GetKeyField", conn)
        Dim param As New SqlParameter

        With param

            .ParameterName = "@KeyValue"
            .Direction = ParameterDirection.Output
            .SqlDbType = SqlDbType.Int
        End With

        With comm
            .CommandTimeout = 0
            .CommandType = CommandType.StoredProcedure
            .Parameters.AddWithValue("@FieldName", sFieldName)
            .Parameters.Add(param)
            .ExecuteNonQuery()
            GetKeyField = param.Value
            .Dispose()
        End With

    End Function
    Public Shared Function GetKeyStrField(ByRef conn As SqlConnection, ByVal sFieldName As String) As String
        Dim comm As New SqlCommand("GetKeyStrField", conn)
        Dim param As New SqlParameter

        With param
            .ParameterName = "@KeyValue"
            .Direction = ParameterDirection.Output
            .SqlDbType = SqlDbType.VarChar
            .Size = 36
        End With

        With comm
            .CommandTimeout = 0
            .CommandType = CommandType.StoredProcedure
            .Parameters.AddWithValue("@FieldName", sFieldName)
            .Parameters.Add(param)
            .ExecuteNonQuery()
            GetKeyStrField = param.Value
            .Dispose()
        End With
        Dim dObject As Type = GetType(Aim)
    End Function
    Public Shared Function GetPaidByStatement(ByRef conn As SqlConnection, ByVal pProducerID As String) As String
        Dim comm As New SqlCommand("SIU_p_GetPaidByStatementByProducerID", conn)
        Dim rs As SqlDataReader
        With comm
            .CommandTimeout = 0
            .CommandType = CommandType.StoredProcedure
            .Parameters.AddWithValue("@ProducerID", pProducerID)
            rs = .ExecuteReader
            .Dispose()
        End With
        If rs.Read Then
            GetPaidByStatement = rs(0)
        Else
            GetPaidByStatement = "N"
        End If
        rs.Close()
    End Function
    Public Shared Function GetNewInsuredID(ByRef conn As SqlConnection) As String
        Dim comm As New SqlCommand("SIU_P_GETNEWINSUREDID", conn)
        Dim rs As SqlDataReader
        With comm
            .CommandType = CommandType.StoredProcedure
            .CommandTimeout = 0
            rs = .ExecuteReader
            .Dispose()
        End With
        rs.Read()
        GetNewInsuredID = rs("InsuredID")
        rs.Close()
    End Function
End Class