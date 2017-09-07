Imports System.Data.SqlClient
Imports DirectBillImports.Common


Public Class clsHeader
#Region "Local Variables"
    Private _HeaderID As Integer
    Private _SubmissionExists As Boolean = False
    Private _PolicyNumber As String
    Private _TranType As String
    Private _EffDate As Date
    Private _ExpDate As Date
    Private _AgencyID As String
    Private _CompanyID As String
    Private _CoverageID As String
    Private _ProductID As String
    Private _TeamID As String
    Private _InsuredID As String
    Private _InsuredName As String
    Private _InsuredAddress1 As String
    Private _InsuredAddress2 As String
    Private _InsuredCity As String
    Private _InsuredState As String
    Private _InsuredZip As String
    Private _Errors As New Collection
    Private _Details As New colDetail
    Private _Insured As New clsInsured
    Private _Quote As New clsQuote
    Private _Version As New clsVersion
    Private _Policy As New clsPolicy
    Private _InvoiceID As String
    Private _InvoiceKey_PK As Integer

    Private _CoverageA As Integer
    Private _APDeductible As Integer
    Private _WindDeductible As Integer
    Private _Coverage As String = ""
    Private _InceptionDate As Date
#End Region
#Region "Properties"
    Public Property CoverageA() As Integer
        Get
            Return _CoverageA
        End Get
        Set(ByVal Value As Integer)
            _CoverageA = Value
        End Set
    End Property
    Public Property APDeductible() As Integer
        Get
            Return _APDeductible
        End Get
        Set(ByVal Value As Integer)
            _APDeductible = Value
        End Set
    End Property
    Public Property WindDeductible() As Integer
        Get
            Return _WindDeductible
        End Get
        Set(ByVal Value As Integer)
            _WindDeductible = Value
        End Set
    End Property
    Public Property Coverage() As String
        Get
            Return _Coverage
        End Get
        Set(ByVal Value As String)
            _Coverage = Value
        End Set
    End Property
    Public Property HeaderID() As Integer
        Get
            Return _HeaderID
        End Get
        Set(ByVal Value As Integer)
            _HeaderID = Value
        End Set
    End Property
    Public Property InsuredID() As String
        Get
            Return _InsuredID
        End Get
        Set(ByVal Value As String)
            _InsuredID = Value
        End Set
    End Property
    Public Property InvoiceKey_PK() As Integer
        Get
            Return _InvoiceKey_PK
        End Get
        Set(ByVal Value As Integer)
            _InvoiceKey_PK = Value
        End Set
    End Property
    Public Property InvoiceID() As String
        Get
            Return _InvoiceID
        End Get
        Set(ByVal Value As String)
            _InvoiceID = Value
        End Set
    End Property
    Public Property SubmissionExists() As Boolean
        Get
            Return _SubmissionExists
        End Get
        Set(ByVal Value As Boolean)
            _SubmissionExists = Value
        End Set
    End Property
    Public Property Details() As colDetail
        Get
            Return _Details
        End Get
        Set(ByVal Value As colDetail)
            _Details = Value
        End Set
    End Property
    Public Property PolicyNumber() As String
        Get
            Return _PolicyNumber
        End Get
        Set(ByVal Value As String)
            _PolicyNumber = Value
        End Set
    End Property
    Public Property TranType() As String
        Get
            Return _TranType
        End Get
        Set(ByVal Value As String)
            _TranType = Value
        End Set
    End Property
    Public Property EffDate() As Date
        Get
            Return _EffDate
        End Get
        Set(ByVal Value As Date)
            _EffDate = Value
        End Set
    End Property
    Public Property InceptionDate() As Date
        Get
            Return _InceptionDate
        End Get
        Set(ByVal Value As Date)
            _InceptionDate = Value
        End Set
    End Property
    Public Property ExpDate() As Date
        Get
            Return _ExpDate
        End Get
        Set(ByVal Value As Date)
            _ExpDate = Value
        End Set
    End Property
    Public Property AgencyID() As String
        Get
            Return _AgencyID
        End Get
        Set(ByVal Value As String)
            _AgencyID = Value
        End Set
    End Property
    Public Property CompanyID() As String
        Get
            Return _CompanyID
        End Get
        Set(ByVal Value As String)
            _CompanyID = Value
        End Set
    End Property
    Public Property CoverageID() As String
        Get
            Return _CoverageID
        End Get
        Set(ByVal Value As String)
            _CoverageID = Value
        End Set
    End Property
    Public Property ProductID() As String
        Get
            Return _ProductID
        End Get
        Set(ByVal Value As String)
            _ProductID = Value
        End Set
    End Property
    Public Property TeamID() As String
        Get
            Return _TeamID
        End Get
        Set(ByVal Value As String)
            _TeamID = Value
        End Set
    End Property
    Public Property InsuredName() As String
        Get
            Return _InsuredName
        End Get
        Set(ByVal Value As String)
            _InsuredName = Value
        End Set
    End Property
    Public Property InsuredAddress1() As String
        Get
            Return _InsuredAddress1
        End Get
        Set(ByVal Value As String)
            _InsuredAddress1 = Value
        End Set
    End Property
    Public Property InsuredAddress2() As String
        Get
            Return _InsuredAddress2
        End Get
        Set(ByVal Value As String)
            _InsuredAddress2 = Value
        End Set
    End Property
    Public Property InsuredCity() As String
        Get
            Return _InsuredCity
        End Get
        Set(ByVal Value As String)
            _InsuredCity = Value
        End Set
    End Property
    Public Property InsuredState() As String
        Get
            Return _InsuredState
        End Get
        Set(ByVal Value As String)
            _InsuredState = Value
        End Set
    End Property
    Public Property InsuredZip() As String
        Get
            Return _InsuredZip
        End Get
        Set(ByVal Value As String)
            _InsuredZip = Value
        End Set
    End Property
    Public ReadOnly Property Errors() As Collection
        Get
            Return _Errors
        End Get
    End Property
    Friend Property Policy() As clsPolicy
        Get
            Return _Policy
        End Get
        Set(ByVal Value As clsPolicy)
            _Policy = Value
        End Set
    End Property
    Friend Property Quote() As clsQuote
        Get
            Return _Quote
        End Get
        Set(ByVal Value As clsQuote)
            _Quote = Value
        End Set
    End Property
    Friend Property Insured() As clsInsured
        Get
            Return _Insured
        End Get
        Set(ByVal Value As clsInsured)
            _Insured = Value
        End Set
    End Property
    Friend Property Version() As clsVersion
        Get
            Return _Version
        End Get
        Set(ByVal Value As clsVersion)
            _Version = Value
        End Set
    End Property
#End Region
#Region "Methods"
    Public Sub New(ByVal pPolicyNumber As String, ByVal pTranType As String, ByVal pEffDate As Date, ByVal pExpDate As Date, ByVal pAgencyID As String, ByVal pCompanyID As String, ByVal pCoverageID As String, ByVal pProductID As String, ByVal pTeamID As String, ByVal pInsuredName As String, ByVal pInsuredAddress1 As String, ByVal pInsuredAddress2 As String, ByVal pInsuredCity As String, ByVal pInsuredState As String, ByVal pInsuredZip As String, ByVal pHeaderID As Integer, ByVal pCoverageA As Integer, ByVal pAPDeductible As Integer, ByVal pWindDeductible As Integer, ByVal pCoverage As String, ByVal pInceptionDate As Date)
        Try
            _PolicyNumber = pPolicyNumber
            _TranType = pTranType
            _EffDate = pEffDate
            _ExpDate = pExpDate
            _AgencyID = pAgencyID
            _CompanyID = pCompanyID
            _CoverageID = pCoverageID
            _ProductID = pProductID
            _TeamID = pTeamID
            _InsuredName = pInsuredName
            _InsuredAddress1 = pInsuredAddress1
            _InsuredAddress2 = pInsuredAddress2
            _InsuredCity = pInsuredCity
            _InsuredState = pInsuredState
            _InsuredZip = pInsuredZip
            _HeaderID = pHeaderID
            _CoverageA = pCoverageA
            _APDeductible = pAPDeductible
            _WindDeductible = pWindDeductible
            _Coverage = pCoverage
            _InceptionDate = pInceptionDate

            If ExpDate <= EffDate Then
                Errors.Add("Expiration Date Must be greater than Effective Date")
            End If
            Dim conn As New SqlConnection(AIMConnectionString("P"))

            conn.Open()
            If Not Exists(conn, "select producerid from producer where producerid = '" & _AgencyID & "'") Then
                _Errors.Add("Invalid Agency ID")
            End If

            If Not Exists(conn, "select companyid from company where companyid = '" & _CompanyID & "'") Then
                _Errors.Add("Invalid Company ID")
            End If

            If Not Exists(conn, "select Coverageid from COVERAGE where Coverageid = '" & _CoverageID & "'") Then
                _Errors.Add("Invalid Coverage ID")
            End If

            If Not Exists(conn, "select Productid from PRODUCT where Productid = '" & _ProductID & "'") Then
                _Errors.Add("Invalid Product ID")
            End If
            'If pCoverageA <> 0 Then Stop


            conn.Close()

        Catch ex As Exception
            _Errors.Add(ex.Message)
        End Try

    End Sub
    Public Function SummaryHTML() As String
        Dim shtml As String = "<table>"
        shtml += "<TR><TD>PolicyNumber:</TD><TD>" & CStr(_PolicyNumber) & "</TD></TR>"
        shtml += "<TR><TD>TranType:</TD><TD>" & CStr(_TranType) & "</TD></TR>"
        shtml += "<TR><TD>EffDate:</TD><TD>" & CStr(_EffDate) & "</TD></TR>"
        shtml += "<TR><TD>ExpDate:</TD><TD>" & CStr(_ExpDate) & "</TD></TR>"
        shtml += "<TR><TD>AgencyID:</TD><TD>" & CStr(_AgencyID) & "</TD></TR>"
        shtml += "<TR><TD>CompanyID:</TD><TD>" & CStr(_CompanyID) & "</TD></TR>"
        shtml += "<TR><TD>CoverageID:</TD><TD>" & CStr(_CoverageID) & "</TD></TR>"
        shtml += "<TR><TD>ProductID:</TD><TD>" & CStr(_ProductID) & "</TD></TR>"
        shtml += "<TR><TD>TeamID:</TD><TD>" & CStr(_TeamID) & "</TD></TR>"
        shtml += "<TR><TD>InsuredName:</TD><TD>" & CStr(_InsuredName) & "</TD></TR>"
        shtml += "<TR><TD>InsuredAddress1:</TD><TD>" & CStr(_InsuredAddress1) & "</TD></TR>"
        shtml += "<TR><TD>InsuredAddress2:</TD><TD>" & CStr(_InsuredAddress2) & "</TD></TR>"
        shtml += "<TR><TD>InsuredCity:</TD><TD>" & CStr(_InsuredCity) & "</TD></TR>"
        shtml += "<TR><TD>InsuredState:</TD><TD>" & CStr(_InsuredState) & "</TD></TR>"
        shtml += "<TR><TD>InsuredZip:</TD><TD>" & CStr(_InsuredZip) & "</TD></TR>"
        shtml += "</table>"

        Return shtml
    End Function
    Public Function LoadDetails(ByRef conn As SqlConnection, ByVal pHeaderID As Integer) As String
        Try
            Dim comm As New SqlCommand("select * from SIU_v_ListStagingDetails where headerid = " & _HeaderID, conn)
            Dim rs As SqlDataReader
            With comm
                .CommandTimeout = 0
                rs = .ExecuteReader
                .Dispose()
            End With
            Do While rs.Read
                _Details.Add(rs("linetypeid"), rs("transcd"), rs("amount"), rs("grosscomm"), rs("agentcomm"), rs("collectedby"), rs("description"), rs("payid"))
            Loop
            rs.Close()
            Return ""
        Catch EX As Exception
            Return "Load Details: " & EX.Message
        Finally

        End Try

    End Function
    Public Function Save(ByRef conn As SqlConnection) As String
        Dim comm As New SqlCommand("siu_p_insertstaging_header", conn)
        Try
            With comm
                .CommandTimeout = 0
                .CommandType = CommandType.StoredProcedure
                Dim myP As New SqlParameter
                With myP
                    .Direction = ParameterDirection.Output
                    .ParameterName = "@HeaderID"
                    .SqlDbType = SqlDbType.Int
                End With
                .Parameters.Add(myP)
                .Parameters.Add(New SqlParameter("@PolicyNumber", _PolicyNumber))
                .Parameters.Add(New SqlParameter("@TranType", _TranType))
                .Parameters.Add(New SqlParameter("@EffDate", _EffDate))
                .Parameters.Add(New SqlParameter("@ExpDate", _ExpDate))
                .Parameters.Add(New SqlParameter("@AgencyID", _AgencyID))
                .Parameters.Add(New SqlParameter("@CompanyID", _CompanyID))
                .Parameters.Add(New SqlParameter("@CoverageID", _CoverageID))
                .Parameters.Add(New SqlParameter("@ProductID", _ProductID))
                .Parameters.Add(New SqlParameter("@TeamID", _TeamID))
                .Parameters.Add(New SqlParameter("@InsuredName", _InsuredName))
                .Parameters.Add(New SqlParameter("@InsuredAddress1", _InsuredAddress1))
                .Parameters.Add(New SqlParameter("@InsuredAddress2", _InsuredAddress2))
                .Parameters.Add(New SqlParameter("@InsuredCity", _InsuredCity))
                .Parameters.Add(New SqlParameter("@InsuredState", _InsuredState))
                .Parameters.Add(New SqlParameter("@InsuredZip", _InsuredZip))
                .Parameters.Add(New SqlParameter("@InceptionDate", _InceptionDate))
                '.ExecuteNonQuery()
                _HeaderID = myP.Value
                .Dispose()
            End With
            Return _Details.Save(conn, _HeaderID)
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Function OutputXML() As String
        Dim Answer As String = "<Output>"

        Answer = Answer & "<InvoiceID>" & InvoiceID & "</InvoiceID>"

        Answer = Answer & "<QuoteID>" & Me.Quote.QuoteID & "</QuoteID>"

        Answer = Answer & "<InsuredID>" & Me.InsuredID & "</InsuredID>"

        Answer = Answer & "</Output>"

        Return Answer
    End Function
#End Region
End Class
