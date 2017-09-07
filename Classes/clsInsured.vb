Imports System.Data.SqlClient
Imports DirectBillImports.Common
Public Class clsInsured
#Region "Local Variables"

    Private _InsuredID As String
    Private _NamedInsured As String
    Private _NameType As String
    Private _DBAName As String
    Private _Prefix As String
    Private _First_Name As String
    Private _Last_Name As String
    Private _Middle_Name As String
    Private _Suffix As String
    Private _CombinedName As String
    Private _Address1 As String
    Private _Address2 As String
    Private _City As String
    Private _State As String
    Private _Zip As String
    Private _AddressID As Long
    Private _ProducerID As String
    Private _Reference As Long
    Private _AcctExec As String
    Private _AcctAsst As String
    Private _CSR As String
    Private _Entity As String
    Private _FormMakerName As String
    Private _DirectBillFlag As String
    Private _MailAddress1 As String
    Private _MailAddress2 As String
    Private _MailCity As String
    Private _MailState As String
    Private _MailZip As String
    Private _ContactName As String
    Private _Phone As String
    Private _Fax As String
    Private _EMail As String
    Private _DateOfBirth As Date
    Private _SSN As String
    Private _PhoneExt As String
    Private _WorkPhone As String
    Private _AcctExecID As String
    Private _AcuityKey As Long
    Private _DateAdded As Date
    Private _VehicleCount As Long
    Private _BusinessStructureID As String
    Private _NCCI As String
    Private _Employees As Long
    Private _Payroll As Double
    Private _SicID As String
    Private _Attention As String
    Private _ContactID As Long
    Private _ClaimCount As Long
    Private _PolicyCount As Long
    Private _TeamID As String
    Private _InsuredKey_PK As Long
    Private _GroupKey_FK As Long
    Private _FlagProspect As String
    Private _FlagAssigned As String
    Private _MembershipTypeID As String
    Private _ParentKey_FK As Long
    Private _License As String
    Private _CareOfKey_FK As Long
    Private _Website As String
    Private _SLA As String
    Private _Exempt As String
    Private _RackleyClientKey_FK As Long
    Private _MapToID As String
    Private _Notes As String
    Private _Country As String
    Private _FileNo As String
    Private _DateConverted As Date
    Private _UserDefinedStr1 As String
    Private _UserDefinedStr2 As String
    Private _UserDefinedStr3 As String
    Private _UserDefinedStr4 As String
    Private _UserDefinedDate1 As Date
    Private _UserDefinedValue1 As Double
    Private _CountryID As String
    Private _AcctgInsuredID As String
    Private _ParentInsuredName As String
    Private _FlagParentInsured As String

#End Region
#Region "Properties"

    Public Property InsuredID() As String
        Get
            Return _InsuredID
        End Get
        Set(ByVal Value As String)
            _InsuredID = Value
        End Set
    End Property
    Public Property NamedInsured() As String
        Get
            Return _NamedInsured
        End Get
        Set(ByVal Value As String)
            _NamedInsured = Value
        End Set
    End Property
    Public Property NameType() As String
        Get
            Return _NameType
        End Get
        Set(ByVal Value As String)
            _NameType = Value
        End Set
    End Property
    Public Property DBAName() As String
        Get
            Return _DBAName
        End Get
        Set(ByVal Value As String)
            _DBAName = Value
        End Set
    End Property
    Public Property Prefix() As String
        Get
            Return _Prefix
        End Get
        Set(ByVal Value As String)
            _Prefix = Value
        End Set
    End Property
    Public Property First_Name() As String
        Get
            Return _First_Name
        End Get
        Set(ByVal Value As String)
            _First_Name = Value
        End Set
    End Property
    Public Property Last_Name() As String
        Get
            Return _Last_Name
        End Get
        Set(ByVal Value As String)
            _Last_Name = Value
        End Set
    End Property
    Public Property Middle_Name() As String
        Get
            Return _Middle_Name
        End Get
        Set(ByVal Value As String)
            _Middle_Name = Value
        End Set
    End Property
    Public Property Suffix() As String
        Get
            Return _Suffix
        End Get
        Set(ByVal Value As String)
            _Suffix = Value
        End Set
    End Property
    Public Property CombinedName() As String
        Get
            Return _CombinedName
        End Get
        Set(ByVal Value As String)
            _CombinedName = Value
        End Set
    End Property
    Public Property Address1() As String
        Get
            Return _Address1
        End Get
        Set(ByVal Value As String)
            _Address1 = Value
        End Set
    End Property
    Public Property Address2() As String
        Get
            Return _Address2
        End Get
        Set(ByVal Value As String)
            _Address2 = Value
        End Set
    End Property
    Public Property City() As String
        Get
            Return _City
        End Get
        Set(ByVal Value As String)
            _City = Value
        End Set
    End Property
    Public Property State() As String
        Get
            Return _State
        End Get
        Set(ByVal Value As String)
            _State = Value
        End Set
    End Property
    Public Property Zip() As String
        Get
            Return _Zip
        End Get
        Set(ByVal Value As String)
            _Zip = Value
        End Set
    End Property
    Public Property AddressID() As Long
        Get
            Return _AddressID
        End Get
        Set(ByVal Value As Long)
            _AddressID = Value
        End Set
    End Property
    Public Property ProducerID() As String
        Get
            Return _ProducerID
        End Get
        Set(ByVal Value As String)
            _ProducerID = Value
        End Set
    End Property
    Public Property Reference() As Long
        Get
            Return _Reference
        End Get
        Set(ByVal Value As Long)
            _Reference = Value
        End Set
    End Property
    Public Property AcctExec() As String
        Get
            Return _AcctExec
        End Get
        Set(ByVal Value As String)
            _AcctExec = Value
        End Set
    End Property
    Public Property AcctAsst() As String
        Get
            Return _AcctAsst
        End Get
        Set(ByVal Value As String)
            _AcctAsst = Value
        End Set
    End Property
    Public Property CSR() As String
        Get
            Return _CSR
        End Get
        Set(ByVal Value As String)
            _CSR = Value
        End Set
    End Property
    Public Property Entity() As String
        Get
            Return _Entity
        End Get
        Set(ByVal Value As String)
            _Entity = Value
        End Set
    End Property
    Public Property FormMakerName() As String
        Get
            Return _FormMakerName
        End Get
        Set(ByVal Value As String)
            _FormMakerName = Value
        End Set
    End Property
    Public Property DirectBillFlag() As String
        Get
            Return _DirectBillFlag
        End Get
        Set(ByVal Value As String)
            _DirectBillFlag = Value
        End Set
    End Property
    Public Property MailAddress1() As String
        Get
            Return _MailAddress1
        End Get
        Set(ByVal Value As String)
            _MailAddress1 = Value
        End Set
    End Property
    Public Property MailAddress2() As String
        Get
            Return _MailAddress2
        End Get
        Set(ByVal Value As String)
            _MailAddress2 = Value
        End Set
    End Property
    Public Property MailCity() As String
        Get
            Return _MailCity
        End Get
        Set(ByVal Value As String)
            _MailCity = Value
        End Set
    End Property
    Public Property MailState() As String
        Get
            Return _MailState
        End Get
        Set(ByVal Value As String)
            _MailState = Value
        End Set
    End Property
    Public Property MailZip() As String
        Get
            Return _MailZip
        End Get
        Set(ByVal Value As String)
            _MailZip = Value
        End Set
    End Property
    Public Property ContactName() As String
        Get
            Return _ContactName
        End Get
        Set(ByVal Value As String)
            _ContactName = Value
        End Set
    End Property
    Public Property Phone() As String
        Get
            Return _Phone
        End Get
        Set(ByVal Value As String)
            _Phone = Value
        End Set
    End Property
    Public Property Fax() As String
        Get
            Return _Fax
        End Get
        Set(ByVal Value As String)
            _Fax = Value
        End Set
    End Property
    Public Property EMail() As String
        Get
            Return _EMail
        End Get
        Set(ByVal Value As String)
            _EMail = Value
        End Set
    End Property
    Public Property DateOfBirth() As Date
        Get
            Return _DateOfBirth
        End Get
        Set(ByVal Value As Date)
            _DateOfBirth = Value
        End Set
    End Property
    Public Property SSN() As String
        Get
            Return _SSN
        End Get
        Set(ByVal Value As String)
            _SSN = Value
        End Set
    End Property
    Public Property PhoneExt() As String
        Get
            Return _PhoneExt
        End Get
        Set(ByVal Value As String)
            _PhoneExt = Value
        End Set
    End Property
    Public Property WorkPhone() As String
        Get
            Return _WorkPhone
        End Get
        Set(ByVal Value As String)
            _WorkPhone = Value
        End Set
    End Property
    Public Property AcctExecID() As String
        Get
            Return _AcctExecID
        End Get
        Set(ByVal Value As String)
            _AcctExecID = Value
        End Set
    End Property
    Public Property AcuityKey() As Long
        Get
            Return _AcuityKey
        End Get
        Set(ByVal Value As Long)
            _AcuityKey = Value
        End Set
    End Property
    Public Property DateAdded() As Date
        Get
            Return _DateAdded
        End Get
        Set(ByVal Value As Date)
            _DateAdded = Value
        End Set
    End Property
    Public Property VehicleCount() As Long
        Get
            Return _VehicleCount
        End Get
        Set(ByVal Value As Long)
            _VehicleCount = Value
        End Set
    End Property
    Public Property BusinessStructureID() As String
        Get
            Return _BusinessStructureID
        End Get
        Set(ByVal Value As String)
            _BusinessStructureID = Value
        End Set
    End Property
    Public Property NCCI() As String
        Get
            Return _NCCI
        End Get
        Set(ByVal Value As String)
            _NCCI = Value
        End Set
    End Property
    Public Property Employees() As Long
        Get
            Return _Employees
        End Get
        Set(ByVal Value As Long)
            _Employees = Value
        End Set
    End Property
    Public Property Payroll() As Double
        Get
            Return _Payroll
        End Get
        Set(ByVal Value As Double)
            _Payroll = Value
        End Set
    End Property
    Public Property SicID() As String
        Get
            Return _SicID
        End Get
        Set(ByVal Value As String)
            _SicID = Value
        End Set
    End Property
    Public Property Attention() As String
        Get
            Return _Attention
        End Get
        Set(ByVal Value As String)
            _Attention = Value
        End Set
    End Property
    Public Property ContactID() As Long
        Get
            Return _ContactID
        End Get
        Set(ByVal Value As Long)
            _ContactID = Value
        End Set
    End Property
    Public Property ClaimCount() As Long
        Get
            Return _ClaimCount
        End Get
        Set(ByVal Value As Long)
            _ClaimCount = Value
        End Set
    End Property
    Public Property PolicyCount() As Long
        Get
            Return _PolicyCount
        End Get
        Set(ByVal Value As Long)
            _PolicyCount = Value
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
    Public Property InsuredKey_PK() As Long
        Get
            Return _InsuredKey_PK
        End Get
        Set(ByVal Value As Long)
            _InsuredKey_PK = Value
        End Set
    End Property
    Public Property GroupKey_FK() As Long
        Get
            Return _GroupKey_FK
        End Get
        Set(ByVal Value As Long)
            _GroupKey_FK = Value
        End Set
    End Property
    Public Property FlagProspect() As String
        Get
            Return _FlagProspect
        End Get
        Set(ByVal Value As String)
            _FlagProspect = Value
        End Set
    End Property
    Public Property FlagAssigned() As String
        Get
            Return _FlagAssigned
        End Get
        Set(ByVal Value As String)
            _FlagAssigned = Value
        End Set
    End Property
    Public Property MembershipTypeID() As String
        Get
            Return _MembershipTypeID
        End Get
        Set(ByVal Value As String)
            _MembershipTypeID = Value
        End Set
    End Property
    Public Property ParentKey_FK() As Long
        Get
            Return _ParentKey_FK
        End Get
        Set(ByVal Value As Long)
            _ParentKey_FK = Value
        End Set
    End Property
    Public Property License() As String
        Get
            Return _License
        End Get
        Set(ByVal Value As String)
            _License = Value
        End Set
    End Property
    Public Property CareOfKey_FK() As Long
        Get
            Return _CareOfKey_FK
        End Get
        Set(ByVal Value As Long)
            _CareOfKey_FK = Value
        End Set
    End Property
    Public Property Website() As String
        Get
            Return _Website
        End Get
        Set(ByVal Value As String)
            _Website = Value
        End Set
    End Property
    Public Property SLA() As String
        Get
            Return _SLA
        End Get
        Set(ByVal Value As String)
            _SLA = Value
        End Set
    End Property
    Public Property Exempt() As String
        Get
            Return _Exempt
        End Get
        Set(ByVal Value As String)
            _Exempt = Value
        End Set
    End Property
    Public Property RackleyClientKey_FK() As Long
        Get
            Return _RackleyClientKey_FK
        End Get
        Set(ByVal Value As Long)
            _RackleyClientKey_FK = Value
        End Set
    End Property
    Public Property MapToID() As String
        Get
            Return _MapToID
        End Get
        Set(ByVal Value As String)
            _MapToID = Value
        End Set
    End Property
    Public Property Notes() As String
        Get
            Return _Notes
        End Get
        Set(ByVal Value As String)
            _Notes = Value
        End Set
    End Property
    Public Property Country() As String
        Get
            Return _Country
        End Get
        Set(ByVal Value As String)
            _Country = Value
        End Set
    End Property
    Public Property FileNo() As String
        Get
            Return _FileNo
        End Get
        Set(ByVal Value As String)
            _FileNo = Value
        End Set
    End Property
    Public Property DateConverted() As Date
        Get
            Return _DateConverted
        End Get
        Set(ByVal Value As Date)
            _DateConverted = Value
        End Set
    End Property
    Public Property UserDefinedStr1() As String
        Get
            Return _UserDefinedStr1
        End Get
        Set(ByVal Value As String)
            _UserDefinedStr1 = Value
        End Set
    End Property
    Public Property UserDefinedStr2() As String
        Get
            Return _UserDefinedStr2
        End Get
        Set(ByVal Value As String)
            _UserDefinedStr2 = Value
        End Set
    End Property
    Public Property UserDefinedStr3() As String
        Get
            Return _UserDefinedStr3
        End Get
        Set(ByVal Value As String)
            _UserDefinedStr3 = Value
        End Set
    End Property
    Public Property UserDefinedStr4() As String
        Get
            Return _UserDefinedStr4
        End Get
        Set(ByVal Value As String)
            _UserDefinedStr4 = Value
        End Set
    End Property
    Public Property UserDefinedDate1() As Date
        Get
            Return _UserDefinedDate1
        End Get
        Set(ByVal Value As Date)
            _UserDefinedDate1 = Value
        End Set
    End Property
    Public Property UserDefinedValue1() As Double
        Get
            Return _UserDefinedValue1
        End Get
        Set(ByVal Value As Double)
            _UserDefinedValue1 = Value
        End Set
    End Property
    Public Property CountryID() As String
        Get
            Return _CountryID
        End Get
        Set(ByVal Value As String)
            _CountryID = Value
        End Set
    End Property
    Public Property AcctgInsuredID() As String
        Get
            Return _AcctgInsuredID
        End Get
        Set(ByVal Value As String)
            _AcctgInsuredID = Value
        End Set
    End Property
    Public Property ParentInsuredName() As String
        Get
            Return _ParentInsuredName
        End Get
        Set(ByVal Value As String)
            _ParentInsuredName = Value
        End Set
    End Property
    Public Property FlagParentInsured() As String
        Get
            Return _FlagParentInsured
        End Get
        Set(ByVal Value As String)
            _FlagParentInsured = Value
        End Set
    End Property

#End Region
#Region "Methods"
    Public Function Load(ByRef conn As SqlConnection, ByVal pInsuredID As String) As String
        Try
            Dim comm As New SqlCommand("SIU_p_GetInsured", conn)
            Dim rs As SqlDataReader

            With comm
                .CommandTimeout = 0
                .CommandType = CommandType.StoredProcedure
                .Parameters.AddWithValue("@InsuredID", pInsuredID)
                rs = .ExecuteReader
            End With
            If rs.Read Then
                _InsuredID = rs("InsuredID")
                _NamedInsured = rs("NamedInsured")
                _NameType = rs("NameType")
                _DBAName = rs("DBAName")
                _Prefix = rs("Prefix")
                _First_Name = rs("First_Name")
                _Last_Name = rs("Last_Name")
                _Middle_Name = rs("Middle_Name")
                _Suffix = rs("Suffix")
                _CombinedName = rs("CombinedName")
                _Address1 = rs("Address1")
                _Address2 = rs("Address2")
                _City = rs("City")
                _State = rs("State")
                _Zip = rs("Zip")
                _AddressID = rs("AddressID")
                _ProducerID = rs("ProducerID")
                _Reference = rs("Reference")
                _AcctExec = rs("AcctExec")
                _AcctAsst = rs("AcctAsst")
                _CSR = rs("CSR")
                _Entity = rs("Entity")
                _FormMakerName = rs("FormMakerName")
                _DirectBillFlag = rs("DirectBillFlag")
                _MailAddress1 = rs("MailAddress1")
                _MailAddress2 = rs("MailAddress2")
                _MailCity = rs("MailCity")
                _MailState = rs("MailState")
                _MailZip = rs("MailZip")
                _ContactName = rs("ContactName")
                _Phone = rs("Phone")
                _Fax = rs("Fax")
                _EMail = rs("EMail")
                _DateOfBirth = rs("DateOfBirth")
                _SSN = rs("SSN")
                _PhoneExt = rs("PhoneExt")
                _WorkPhone = rs("WorkPhone")
                _AcctExecID = rs("AcctExecID")
                _AcuityKey = rs("AcuityKey")
                _DateAdded = rs("DateAdded")
                _VehicleCount = rs("VehicleCount")
                _BusinessStructureID = rs("BusinessStructureID")
                _NCCI = rs("NCCI")
                _Employees = rs("Employees")
                _Payroll = rs("Payroll")
                _SicID = rs("SicID")
                _Attention = rs("Attention")
                _ContactID = rs("ContactID")
                _ClaimCount = rs("ClaimCount")
                _PolicyCount = rs("PolicyCount")
                _TeamID = rs("TeamID")
                _InsuredKey_PK = rs("InsuredKey_PK")
                _GroupKey_FK = rs("GroupKey_FK")
                _FlagProspect = rs("FlagProspect")
                _FlagAssigned = rs("FlagAssigned")
                _MembershipTypeID = rs("MembershipTypeID")
                _ParentKey_FK = rs("ParentKey_FK")
                _License = rs("License")
                _CareOfKey_FK = rs("CareOfKey_FK")
                _Website = rs("Website")
                _SLA = rs("SLA")
                _Exempt = rs("Exempt")
                _RackleyClientKey_FK = rs("RackleyClientKey_FK")
                _MapToID = rs("MapToID")
                _Notes = rs("Notes")
                _Country = rs("Country")
                _FileNo = rs("FileNo")
                _DateConverted = rs("DateConverted")
                _UserDefinedStr1 = rs("UserDefinedStr1")
                _UserDefinedStr2 = rs("UserDefinedStr2")
                _UserDefinedStr3 = rs("UserDefinedStr3")
                _UserDefinedStr4 = rs("UserDefinedStr4")
                _UserDefinedDate1 = rs("UserDefinedDate1")
                _UserDefinedValue1 = rs("UserDefinedValue1")
                _CountryID = rs("CountryID")
                _AcctgInsuredID = rs("AcctgInsuredID")
                _ParentInsuredName = rs("ParentInsuredName")
                _FlagParentInsured = rs("FlagParentInsured")
            End If
            rs.Close()
            conn.Close()
            Return ""
        Catch ex As Exception
            Return "Insured: " & ex.Message
            conn.Close()
        End Try
    End Function

    Public Function Exists(ByRef conn As SqlConnection, ByVal pInsuredID As String) As Boolean
        Dim comm As New SqlCommand("SIU_p_InsuredCount", conn)
        Dim rs As SqlDataReader
        Dim Answer As Boolean = True
        With comm
            .CommandTimeout = 0
            .CommandType = CommandType.StoredProcedure
            .Parameters.AddWithValue("@InsuredID", pInsuredID)
            rs = .ExecuteReader
            .Dispose()
        End With
        If rs.Read Then
            Answer = rs("HowMany") > 0
        Else
            Answer = True
        End If
        rs.Close()
        Return Answer
    End Function

#End Region

End Class