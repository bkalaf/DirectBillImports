Imports System.data.sqlclient
Imports DirectBillImports.Common
Public Class clsDetail
#Region "Local Variables"

    Private _LineTypeID As String
    Private _TransCD As String
    Private _Amount As Double
    Private _GrossComm As Double
    Private _AgentComm As Double
    Private _CollectedBy As String
    Private _Description As String
    Private _PayID As String
#End Region
#Region "Properties"
    Friend ReadOnly Property Revenue_Amt() As Double
        Get
            Return Math.Round((_GrossComm * 0.01 * _Amount), 2)
        End Get
    End Property
    Friend ReadOnly Property Expense_Amt() As Double
        Get
            Return Math.Round((_AgentComm * 0.01 * _Amount), 2)
        End Get
    End Property
    Public Property PayID() As String
        Get
            Return _PayID
        End Get
        Set(ByVal Value As String)
            _PayID = Value
        End Set
    End Property
    Public Property CollectedBy() As String
        Get
            Return _CollectedBy
        End Get
        Set(ByVal Value As String)
            _CollectedBy = Value
        End Set
    End Property
    Public Property Description() As String
        Get
            Return _Description
        End Get
        Set(ByVal Value As String)
            _Description = Value
        End Set
    End Property
    Public Property LineTypeID() As String
        Get
            Return _LineTypeID
        End Get
        Set(ByVal Value As String)
            _LineTypeID = Value
        End Set
    End Property
    Public Property TransCD() As String
        Get
            Return _TransCD
        End Get
        Set(ByVal Value As String)
            _TransCD = Value
        End Set
    End Property
    Public Property Amount() As Double
        Get
            Return _Amount
        End Get
        Set(ByVal Value As Double)
            _Amount = Value
        End Set
    End Property
    Public Property GrossComm() As Double
        Get
            Return _GrossComm
        End Get
        Set(ByVal Value As Double)
            _GrossComm = Value
        End Set
    End Property
    Public Property AgentComm() As Double
        Get
            Return _AgentComm
        End Get
        Set(ByVal Value As Double)
            _AgentComm = Value
        End Set
    End Property

#End Region
#Region "Methods"
    Public Function HTML(ByVal pCounter As Integer) As String
        Dim sAnswer As String = ""

        sAnswer += "<table>" & vbCrLf
        sAnswer += "<tr><td>" & CStr(pCounter) & "</td><td>Pay ID:</td><td>" & _PayID & "</td></tr>" & vbCrLf
        sAnswer += "<tr><td></td><td>Line Type ID:</td><td>" & _LineTypeID & "</td></tr>" & vbCrLf
        sAnswer += "<tr><td></td><td>Trans CD:</td><td>" & _TransCD & "</td></tr>" & vbCrLf
        sAnswer += "<tr><td></td><td>Amount:</td><td>" & CStr(_Amount) & "</td></tr>" & vbCrLf
        sAnswer += "<tr><td></td><td>Gross Comm:</td><td>" & CStr(_GrossComm) & "</td></tr>" & vbCrLf
        sAnswer += "<tr><td></td><td>Agent Comm:</td><td>" & CStr(_AgentComm) & "</td></tr>" & vbCrLf
        sAnswer += "<tr><td></td><td>Collected By:</td><td>" & _CollectedBy & "</td></tr>" & vbCrLf
        sAnswer += "<tr><td></td><td>Description:</td><td>" & _Description & "</td></tr>" & vbCrLf

        sAnswer += "</table>" & vbCrLf
        Return sAnswer
    End Function
    Public Function Save(ByRef conn As SqlConnection, ByVal pHeaderID As Integer, ByVal pDetailID As Integer) As String
        Dim comm As New SqlCommand("siu_p_insertstaging_detail", conn)
        Try
            With comm
                .CommandType = CommandType.StoredProcedure
                .Parameters.AddWithValue("@HeaderID", pHeaderID)
                .Parameters.AddWithValue("@DetailID", pDetailID)
                .Parameters.AddWithValue("@LineTypeID", _LineTypeID)
                .Parameters.AddWithValue("@TransCD", _TransCD)
                .Parameters.AddWithValue("@Amount", _Amount)
                .Parameters.AddWithValue("@GrossComm", _GrossComm)
                .Parameters.AddWithValue("@AgentComm", _AgentComm)
                .Parameters.AddWithValue("@CollectedBy", _CollectedBy)
                .Parameters.AddWithValue("@Description", _Description)
                .Parameters.AddWithValue("@PayID", _PayID)
                '.ExecuteNonQuery()
            End With
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
#End Region
End Class
