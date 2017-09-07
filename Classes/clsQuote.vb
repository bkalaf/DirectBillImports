Imports System.Data.SqlClient
Imports DirectBillImports.Common
Public Class clsQuote
#Region "Local Variables"

    Private _QuoteID As String
    Private _VersionBound As String
    Private _ProducerID As String
    Private _NamedInsured As String
    Private _TypeID As String
    Private _UserID As String
    Private _Attention As String
    Private _Received As Date
    Private _Acknowledged As Date
    Private _Quoted As Date
    Private _TeamID As String
    Private _DivisionID As String
    Private _StatusID As String
    Private _CreatedID As String
    Private _Renewal As String
    Private _OldPolicyID As String
    Private _OldVersion As String
    Private _OldExpiration As Date
    Private _OpenItem As String
    Private _Notes As String
    Private _PolicyID As String
    Private _VersionCounter As String
    Private _InsuredID As String
    Private _Description As String
    Private _FileLocation As String
    Private _Address1 As String
    Private _Address2 As String
    Private _City As String
    Private _State As String
    Private _Zip As String
    Private _Bound As Date
    Private _Submitted As Date
    Private _SubmitType As String
    Private _NoteAttached As String
    Private _AcctExec As String
    Private _InsuredInterest As String
    Private _RiskInformation As String
    Private _EC As String
    Private _BndPremium As Double
    Private _BndFee As Double
    Private _CompanyID As String
    Private _ProductID As String
    Private _Effective As Date
    Private _Expiration As Date
    Private _Setup As Date
    Private _PolicyMailOut As Date
    Private _BinderRev As Integer
    Private _PriorCarrier As String
    Private _TargetPremium As Double
    Private _CsrID As String
    Private _PolicyVer As String
    Private _OldQuoteID As String
    Private _PolicyGrpID As String
    Private _PendingSuspenseID As Long
    Private _ReferenceID As Long
    Private _MapToID As String
    Private _SubmitGrpID As String
    Private _AcctAsst As String
    Private _TaxState As String
    Private _SicID As String
    Private _CoverageID As String
    Private _OldPremium As Double
    Private _AddressID As Long
    Private _OldEffective As Date
    Private _TaxBasis As Double
    Private _QuoteRequiredBy As Date
    Private _RequiredLimits As Double
    Private _RequiredDeduct As Double
    Private _Retroactive As Date
    Private _PrevCancelFlag As String
    Private _PrevNonRenew As String
    Private _PriorPremium As Double
    Private _PriorLimits As String
    Private _UWCheckList As String
    Private _FileSetup As Date
    Private _ContactID As Long
    Private _SuspenseFlag As String
    Private _PriorDeductible As String
    Private _CategoryID As String
    Private _StructureID As String
    Private _RenewalStatusID As String
    Private _ClaimsFlag As String
    Private _ActivePolicyFlag As String
    Private _Assets As Double
    Private _PublicEntity As String
    Private _VentureID As String
    Private _IncorporatedState As String
    Private _ReInsuranceFlag As String
    Private _TaxedPaidBy As String
    Private _LayeredCoverage As String
    Private _Employees As Long
    Private _Stock_52wk As String
    Private _NetIncome As Double
    Private _LossHistory As String
    Private _PriorLimitsNew As String
    Private _LargeLossHistory As String
    Private _DateOfApp As Date
    Private _Stock_High As String
    Private _Stock_Low As String
    Private _Stock_Current As String
    Private _MarketCap As String
    Private _Exposures As String
    Private _AIM_TransDate As Date
    Private _LostBusinessFlag As String
    Private _YearEst As Long
    Private _LostBusiness_Carrier As String
    Private _LostBusiness_Premium As Double
    Private _AccountKey_FK As Long
    Private _FlagRewrite As String
    Private _flagWIP As String
    Private _RenewalQuoteID As String
    Private _QuoteDueDate As Date
    Private _QuoteStatus As String
    Private _BinderExpires As Date
    Private _TIV As Double
    Private _InvoicedPremium As Double
    Private _InvoicedFee As Double
    Private _InvoicedCommRev As Double
    Private _SplitAccount As String
    Private _FileCloseReason As String
    Private _FileCloseReasonID As String
    Private _SourceOfLeadID As String
    Private _ServiceUWID As String
    Private _SubmitTypeID As String
    Private _SubProducerID As String
    Private _AgtAccountNumber As String
    Private _BndMarketID As String
    Private _RefQuoteID As String
    Private _FlagHeldFile As String
    Private _HeldFileMessage As String
    Private _TermPremium As Double
    Private _ProcessBatchKey_FK As Long
    Private _PolicyInception As Date
    Private _QuoteClassID As String
    Private _ScheduleIRM As Double
    Private _ClaimExpRM As Double
    Private _DateAppRecvd As Date
    Private _DateLossRunRecvd As Date
    Private _CoverageEffective As Date
    Private _CoverageExpired As Date
    Private _SLA As String
    Private _QuoteClass As String
    Private _IRFileNum As String
    Private _IRDrawer As String
    Private _FlagOverRideBy As String
    Private _RackleyQuoteID As Long
    Private _FlagCourtesyFiling As String
    Private _FlagRPG As String
    Private _CurrencyType As String
    Private _CurrencySymbol As String
    Private _FileNo As String
    Private _UserDefinedStr1 As String
    Private _UserDefinedStr2 As String
    Private _UserDefinedStr3 As String
    Private _UserDefinedStr4 As String
    Private _UserDefinedDate1 As Date
    Private _UserDefinedValue1 As Double
    Private _ReservedContractID As String
    Private _CountryID As String
    Private _RatingKey_FK As Long
    Private _eAttached As String
    Private _NewField As String
    Private _TotalCoinsuranceLimit As Double
    Private _TotalCoinsurancePremium As Double
    Private _CurrencyExchRate As Double
    Private _Invoiced As String
    Private _OtherLead As String
    Private _LeadCarrierID As String
    Private _RenewTypeID As String
    Private _IsoCode As String
    Private _CedingPolicyID As String
    Private _CedingPolicyDate As Date
    Private _ConversionStatusID As String
    Private _FlagTaxExempt As String
    Private _Units As Long
    Private _SubUnits As Long
    Private _LicenseAgtKey_FK As Long
    Private _ContractPlanKey_FK As Long
    Private _AltStatusID As String
    Private _FlagNonResidentAgt As String
    Private _FirewallTeamID As String
    Private _CedingPolicyEndDate As Date
    Private _Risk_CommProperty As New clsRisk_CommProperty

#End Region
#Region "Properties"
    Public Property Risk_CommProperty() As clsRisk_CommProperty
        Get
            Return _Risk_CommProperty
        End Get
        Set(ByVal Value As clsRisk_CommProperty)
            _Risk_CommProperty = Value
        End Set
    End Property
    Public Property QuoteID() As String
        Get
            Return _QuoteID
        End Get
        Set(ByVal Value As String)
            _QuoteID = Value
        End Set
    End Property
    Public Property VersionBound() As String
        Get
            Return _VersionBound
        End Get
        Set(ByVal Value As String)
            _VersionBound = Value
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
    Public Property NamedInsured() As String
        Get
            Return _NamedInsured
        End Get
        Set(ByVal Value As String)
            _NamedInsured = Value
        End Set
    End Property
    Public Property TypeID() As String
        Get
            Return _TypeID
        End Get
        Set(ByVal Value As String)
            _TypeID = Value
        End Set
    End Property
    Public Property UserID() As String
        Get
            Return _UserID
        End Get
        Set(ByVal Value As String)
            _UserID = Value
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
    Public Property Received() As Date
        Get
            Return _Received
        End Get
        Set(ByVal Value As Date)
            _Received = Value
        End Set
    End Property
    Public Property Acknowledged() As Date
        Get
            Return _Acknowledged
        End Get
        Set(ByVal Value As Date)
            _Acknowledged = Value
        End Set
    End Property
    Public Property Quoted() As Date
        Get
            Return _Quoted
        End Get
        Set(ByVal Value As Date)
            _Quoted = Value
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
    Public Property DivisionID() As String
        Get
            Return IIf(Len(Trim(_DivisionID)) > 0, _DivisionID, "")
        End Get
        Set(ByVal Value As String)
            _DivisionID = Value
        End Set
    End Property
    Public Property StatusID() As String
        Get
            Return _StatusID
        End Get
        Set(ByVal Value As String)
            _StatusID = Value
        End Set
    End Property
    Public Property CreatedID() As String
        Get
            Return _CreatedID
        End Get
        Set(ByVal Value As String)
            _CreatedID = Value
        End Set
    End Property
    Public Property Renewal() As String
        Get
            Return _Renewal
        End Get
        Set(ByVal Value As String)
            _Renewal = Value
        End Set
    End Property
    Public Property OldPolicyID() As String
        Get
            Return _OldPolicyID
        End Get
        Set(ByVal Value As String)
            _OldPolicyID = Value
        End Set
    End Property
    Public Property OldVersion() As String
        Get
            Return _OldVersion
        End Get
        Set(ByVal Value As String)
            _OldVersion = Value
        End Set
    End Property
    Public Property OldExpiration() As Date
        Get
            Return _OldExpiration
        End Get
        Set(ByVal Value As Date)
            _OldExpiration = Value
        End Set
    End Property
    Public Property OpenItem() As String
        Get
            Return _OpenItem
        End Get
        Set(ByVal Value As String)
            _OpenItem = Value
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
    Public Property PolicyID() As String
        Get
            Return _PolicyID
        End Get
        Set(ByVal Value As String)
            _PolicyID = Value
        End Set
    End Property
    Public Property VersionCounter() As String
        Get
            Return _VersionCounter
        End Get
        Set(ByVal Value As String)
            _VersionCounter = Value
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
    Public Property Description() As String
        Get
            Return _Description
        End Get
        Set(ByVal Value As String)
            _Description = Value
        End Set
    End Property
    Public Property FileLocation() As String
        Get
            Return _FileLocation
        End Get
        Set(ByVal Value As String)
            _FileLocation = Value
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
    Public Property Bound() As Date
        Get
            Return _Bound
        End Get
        Set(ByVal Value As Date)
            _Bound = Value
        End Set
    End Property
    Public Property Submitted() As Date
        Get
            Return _Submitted
        End Get
        Set(ByVal Value As Date)
            _Submitted = Value
        End Set
    End Property
    Public Property SubmitType() As String
        Get
            Return _SubmitType
        End Get
        Set(ByVal Value As String)
            _SubmitType = Value
        End Set
    End Property
    Public Property NoteAttached() As String
        Get
            Return _NoteAttached
        End Get
        Set(ByVal Value As String)
            _NoteAttached = Value
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
    Public Property InsuredInterest() As String
        Get
            Return _InsuredInterest
        End Get
        Set(ByVal Value As String)
            _InsuredInterest = Value
        End Set
    End Property
    Public Property RiskInformation() As String
        Get
            Return _RiskInformation
        End Get
        Set(ByVal Value As String)
            _RiskInformation = Value
        End Set
    End Property
    Public Property EC() As String
        Get
            Return _EC
        End Get
        Set(ByVal Value As String)
            _EC = Value
        End Set
    End Property
    Public Property BndPremium() As Double
        Get
            Return _BndPremium
        End Get
        Set(ByVal Value As Double)
            _BndPremium = Value
        End Set
    End Property
    Public Property BndFee() As Double
        Get
            Return _BndFee
        End Get
        Set(ByVal Value As Double)
            _BndFee = Value
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
    Public Property ProductID() As String
        Get
            Return _ProductID
        End Get
        Set(ByVal Value As String)
            _ProductID = Value
        End Set
    End Property
    Public Property Effective() As Date
        Get
            Return _Effective
        End Get
        Set(ByVal Value As Date)
            _Effective = Value
        End Set
    End Property
    Public Property Expiration() As Date
        Get
            Return _Expiration
        End Get
        Set(ByVal Value As Date)
            _Expiration = Value
        End Set
    End Property
    Public Property Setup() As Date
        Get
            Return _Setup
        End Get
        Set(ByVal Value As Date)
            _Setup = Value
        End Set
    End Property
    Public Property PolicyMailOut() As Date
        Get
            Return _PolicyMailOut
        End Get
        Set(ByVal Value As Date)
            _PolicyMailOut = Value
        End Set
    End Property
    Public Property BinderRev() As Integer
        Get
            Return _BinderRev
        End Get
        Set(ByVal Value As Integer)
            _BinderRev = Value
        End Set
    End Property
    Public Property PriorCarrier() As String
        Get
            Return _PriorCarrier
        End Get
        Set(ByVal Value As String)
            _PriorCarrier = Value
        End Set
    End Property
    Public Property TargetPremium() As Double
        Get
            Return _TargetPremium
        End Get
        Set(ByVal Value As Double)
            _TargetPremium = Value
        End Set
    End Property
    Public Property CsrID() As String
        Get
            Return _CsrID
        End Get
        Set(ByVal Value As String)
            _CsrID = Value
        End Set
    End Property
    Public Property PolicyVer() As String
        Get
            Return _PolicyVer
        End Get
        Set(ByVal Value As String)
            _PolicyVer = Value
        End Set
    End Property
    Public Property OldQuoteID() As String
        Get
            Return _OldQuoteID
        End Get
        Set(ByVal Value As String)
            _OldQuoteID = Value
        End Set
    End Property
    Public Property PolicyGrpID() As String
        Get
            Return _PolicyGrpID
        End Get
        Set(ByVal Value As String)
            _PolicyGrpID = Value
        End Set
    End Property
    Public Property PendingSuspenseID() As Long
        Get
            Return _PendingSuspenseID
        End Get
        Set(ByVal Value As Long)
            _PendingSuspenseID = Value
        End Set
    End Property
    Public Property ReferenceID() As Long
        Get
            Return _ReferenceID
        End Get
        Set(ByVal Value As Long)
            _ReferenceID = Value
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
    Public Property SubmitGrpID() As String
        Get
            Return _SubmitGrpID
        End Get
        Set(ByVal Value As String)
            _SubmitGrpID = Value
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
    Public Property TaxState() As String
        Get
            Return _TaxState
        End Get
        Set(ByVal Value As String)
            _TaxState = Value
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
    Public Property CoverageID() As String
        Get
            Return _CoverageID
        End Get
        Set(ByVal Value As String)
            _CoverageID = Value
        End Set
    End Property
    Public Property OldPremium() As Double
        Get
            Return _OldPremium
        End Get
        Set(ByVal Value As Double)
            _OldPremium = Value
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
    Public Property OldEffective() As Date
        Get
            Return _OldEffective
        End Get
        Set(ByVal Value As Date)
            _OldEffective = Value
        End Set
    End Property
    Public Property TaxBasis() As Double
        Get
            Return _TaxBasis
        End Get
        Set(ByVal Value As Double)
            _TaxBasis = Value
        End Set
    End Property
    Public Property QuoteRequiredBy() As Date
        Get
            Return _QuoteRequiredBy
        End Get
        Set(ByVal Value As Date)
            _QuoteRequiredBy = Value
        End Set
    End Property
    Public Property RequiredLimits() As Double
        Get
            Return _RequiredLimits
        End Get
        Set(ByVal Value As Double)
            _RequiredLimits = Value
        End Set
    End Property
    Public Property RequiredDeduct() As Double
        Get
            Return _RequiredDeduct
        End Get
        Set(ByVal Value As Double)
            _RequiredDeduct = Value
        End Set
    End Property
    Public Property Retroactive() As Date
        Get
            Return _Retroactive
        End Get
        Set(ByVal Value As Date)
            _Retroactive = Value
        End Set
    End Property
    Public Property PrevCancelFlag() As String
        Get
            Return _PrevCancelFlag
        End Get
        Set(ByVal Value As String)
            _PrevCancelFlag = Value
        End Set
    End Property
    Public Property PrevNonRenew() As String
        Get
            Return _PrevNonRenew
        End Get
        Set(ByVal Value As String)
            _PrevNonRenew = Value
        End Set
    End Property
    Public Property PriorPremium() As Double
        Get
            Return _PriorPremium
        End Get
        Set(ByVal Value As Double)
            _PriorPremium = Value
        End Set
    End Property
    Public Property PriorLimits() As String
        Get
            Return _PriorLimits
        End Get
        Set(ByVal Value As String)
            _PriorLimits = Value
        End Set
    End Property
    Public Property UWCheckList() As String
        Get
            Return _UWCheckList
        End Get
        Set(ByVal Value As String)
            _UWCheckList = Value
        End Set
    End Property
    Public Property FileSetup() As Date
        Get
            Return _FileSetup
        End Get
        Set(ByVal Value As Date)
            _FileSetup = Value
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
    Public Property SuspenseFlag() As String
        Get
            Return _SuspenseFlag
        End Get
        Set(ByVal Value As String)
            _SuspenseFlag = Value
        End Set
    End Property
    Public Property PriorDeductible() As String
        Get
            Return _PriorDeductible
        End Get
        Set(ByVal Value As String)
            _PriorDeductible = Value
        End Set
    End Property
    Public Property CategoryID() As String
        Get
            Return _CategoryID
        End Get
        Set(ByVal Value As String)
            _CategoryID = Value
        End Set
    End Property
    Public Property StructureID() As String
        Get
            Return _StructureID
        End Get
        Set(ByVal Value As String)
            _StructureID = Value
        End Set
    End Property
    Public Property RenewalStatusID() As String
        Get
            Return _RenewalStatusID
        End Get
        Set(ByVal Value As String)
            _RenewalStatusID = Value
        End Set
    End Property
    Public Property ClaimsFlag() As String
        Get
            Return _ClaimsFlag
        End Get
        Set(ByVal Value As String)
            _ClaimsFlag = Value
        End Set
    End Property
    Public Property ActivePolicyFlag() As String
        Get
            Return _ActivePolicyFlag
        End Get
        Set(ByVal Value As String)
            _ActivePolicyFlag = Value
        End Set
    End Property
    Public Property Assets() As Double
        Get
            Return _Assets
        End Get
        Set(ByVal Value As Double)
            _Assets = Value
        End Set
    End Property
    Public Property PublicEntity() As String
        Get
            Return _PublicEntity
        End Get
        Set(ByVal Value As String)
            _PublicEntity = Value
        End Set
    End Property
    Public Property VentureID() As String
        Get
            Return _VentureID
        End Get
        Set(ByVal Value As String)
            _VentureID = Value
        End Set
    End Property
    Public Property IncorporatedState() As String
        Get
            Return _IncorporatedState
        End Get
        Set(ByVal Value As String)
            _IncorporatedState = Value
        End Set
    End Property
    Public Property ReInsuranceFlag() As String
        Get
            Return _ReInsuranceFlag
        End Get
        Set(ByVal Value As String)
            _ReInsuranceFlag = Value
        End Set
    End Property
    Public Property TaxedPaidBy() As String
        Get
            Return _TaxedPaidBy
        End Get
        Set(ByVal Value As String)
            _TaxedPaidBy = Value
        End Set
    End Property
    Public Property LayeredCoverage() As String
        Get
            Return _LayeredCoverage
        End Get
        Set(ByVal Value As String)
            _LayeredCoverage = Value
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
    Public Property Stock_52wk() As String
        Get
            Return _Stock_52wk
        End Get
        Set(ByVal Value As String)
            _Stock_52wk = Value
        End Set
    End Property
    Public Property NetIncome() As Double
        Get
            Return _NetIncome
        End Get
        Set(ByVal Value As Double)
            _NetIncome = Value
        End Set
    End Property
    Public Property LossHistory() As String
        Get
            Return _LossHistory
        End Get
        Set(ByVal Value As String)
            _LossHistory = Value
        End Set
    End Property
    Public Property PriorLimitsNew() As String
        Get
            Return _PriorLimitsNew
        End Get
        Set(ByVal Value As String)
            _PriorLimitsNew = Value
        End Set
    End Property
    Public Property LargeLossHistory() As String
        Get
            Return _LargeLossHistory
        End Get
        Set(ByVal Value As String)
            _LargeLossHistory = Value
        End Set
    End Property
    Public Property DateOfApp() As Date
        Get
            Return _DateOfApp
        End Get
        Set(ByVal Value As Date)
            _DateOfApp = Value
        End Set
    End Property
    Public Property Stock_High() As String
        Get
            Return _Stock_High
        End Get
        Set(ByVal Value As String)
            _Stock_High = Value
        End Set
    End Property
    Public Property Stock_Low() As String
        Get
            Return _Stock_Low
        End Get
        Set(ByVal Value As String)
            _Stock_Low = Value
        End Set
    End Property
    Public Property Stock_Current() As String
        Get
            Return _Stock_Current
        End Get
        Set(ByVal Value As String)
            _Stock_Current = Value
        End Set
    End Property
    Public Property MarketCap() As String
        Get
            Return _MarketCap
        End Get
        Set(ByVal Value As String)
            _MarketCap = Value
        End Set
    End Property
    Public Property Exposures() As String
        Get
            Return _Exposures
        End Get
        Set(ByVal Value As String)
            _Exposures = Value
        End Set
    End Property
    Public Property AIM_TransDate() As Date
        Get
            Return _AIM_TransDate
        End Get
        Set(ByVal Value As Date)
            _AIM_TransDate = Value
        End Set
    End Property
    Public Property LostBusinessFlag() As String
        Get
            Return _LostBusinessFlag
        End Get
        Set(ByVal Value As String)
            _LostBusinessFlag = Value
        End Set
    End Property
    Public Property YearEst() As Long
        Get
            Return _YearEst
        End Get
        Set(ByVal Value As Long)
            _YearEst = Value
        End Set
    End Property
    Public Property LostBusiness_Carrier() As String
        Get
            Return _LostBusiness_Carrier
        End Get
        Set(ByVal Value As String)
            _LostBusiness_Carrier = Value
        End Set
    End Property
    Public Property LostBusiness_Premium() As Double
        Get
            Return _LostBusiness_Premium
        End Get
        Set(ByVal Value As Double)
            _LostBusiness_Premium = Value
        End Set
    End Property
    Public Property AccountKey_FK() As Long
        Get
            Return _AccountKey_FK
        End Get
        Set(ByVal Value As Long)
            _AccountKey_FK = Value
        End Set
    End Property
    Public Property FlagRewrite() As String
        Get
            Return _FlagRewrite
        End Get
        Set(ByVal Value As String)
            _FlagRewrite = Value
        End Set
    End Property
    Public Property flagWIP() As String
        Get
            Return _flagWIP
        End Get
        Set(ByVal Value As String)
            _flagWIP = Value
        End Set
    End Property
    Public Property RenewalQuoteID() As String
        Get
            Return _RenewalQuoteID
        End Get
        Set(ByVal Value As String)
            _RenewalQuoteID = Value
        End Set
    End Property
    Public Property QuoteDueDate() As Date
        Get
            Return _QuoteDueDate
        End Get
        Set(ByVal Value As Date)
            _QuoteDueDate = Value
        End Set
    End Property
    Public Property QuoteStatus() As String
        Get
            Return _QuoteStatus
        End Get
        Set(ByVal Value As String)
            _QuoteStatus = Value
        End Set
    End Property
    Public Property BinderExpires() As Date
        Get
            Return _BinderExpires
        End Get
        Set(ByVal Value As Date)
            _BinderExpires = Value
        End Set
    End Property
    Public Property TIV() As Double
        Get
            Return _TIV
        End Get
        Set(ByVal Value As Double)
            _TIV = Value
        End Set
    End Property
    Public Property InvoicedPremium() As Double
        Get
            Return _InvoicedPremium
        End Get
        Set(ByVal Value As Double)
            _InvoicedPremium = Value
        End Set
    End Property
    Public Property InvoicedFee() As Double
        Get
            Return _InvoicedFee
        End Get
        Set(ByVal Value As Double)
            _InvoicedFee = Value
        End Set
    End Property
    Public Property InvoicedCommRev() As Double
        Get
            Return _InvoicedCommRev
        End Get
        Set(ByVal Value As Double)
            _InvoicedCommRev = Value
        End Set
    End Property
    Public Property SplitAccount() As String
        Get
            Return _SplitAccount
        End Get
        Set(ByVal Value As String)
            _SplitAccount = Value
        End Set
    End Property
    Public Property FileCloseReason() As String
        Get
            Return _FileCloseReason
        End Get
        Set(ByVal Value As String)
            _FileCloseReason = Value
        End Set
    End Property
    Public Property FileCloseReasonID() As String
        Get
            Return _FileCloseReasonID
        End Get
        Set(ByVal Value As String)
            _FileCloseReasonID = Value
        End Set
    End Property
    Public Property SourceOfLeadID() As String
        Get
            Return _SourceOfLeadID
        End Get
        Set(ByVal Value As String)
            _SourceOfLeadID = Value
        End Set
    End Property
    Public Property ServiceUWID() As String
        Get
            Return _ServiceUWID
        End Get
        Set(ByVal Value As String)
            _ServiceUWID = Value
        End Set
    End Property
    Public Property SubmitTypeID() As String
        Get
            Return _SubmitTypeID
        End Get
        Set(ByVal Value As String)
            _SubmitTypeID = Value
        End Set
    End Property
    Public Property SubProducerID() As String
        Get
            Return _SubProducerID
        End Get
        Set(ByVal Value As String)
            _SubProducerID = Value
        End Set
    End Property
    Public Property AgtAccountNumber() As String
        Get
            Return _AgtAccountNumber
        End Get
        Set(ByVal Value As String)
            _AgtAccountNumber = Value
        End Set
    End Property
    Public Property BndMarketID() As String
        Get
            Return _BndMarketID
        End Get
        Set(ByVal Value As String)
            _BndMarketID = Value
        End Set
    End Property
    Public Property RefQuoteID() As String
        Get
            Return _RefQuoteID
        End Get
        Set(ByVal Value As String)
            _RefQuoteID = Value
        End Set
    End Property
    Public Property FlagHeldFile() As String
        Get
            Return _FlagHeldFile
        End Get
        Set(ByVal Value As String)
            _FlagHeldFile = Value
        End Set
    End Property
    Public Property HeldFileMessage() As String
        Get
            Return _HeldFileMessage
        End Get
        Set(ByVal Value As String)
            _HeldFileMessage = Value
        End Set
    End Property
    Public Property TermPremium() As Double
        Get
            Return _TermPremium
        End Get
        Set(ByVal Value As Double)
            _TermPremium = Value
        End Set
    End Property
    Public Property ProcessBatchKey_FK() As Long
        Get
            Return _ProcessBatchKey_FK
        End Get
        Set(ByVal Value As Long)
            _ProcessBatchKey_FK = Value
        End Set
    End Property
    Public Property PolicyInception() As Date
        Get
            Return _PolicyInception
        End Get
        Set(ByVal Value As Date)
            _PolicyInception = Value
        End Set
    End Property
    Public Property ClassID() As String
        Get
            Return _QuoteClassID
        End Get
        Set(ByVal Value As String)
            _QuoteClassID = Value
        End Set
    End Property
    Public Property ScheduleIRM() As Double
        Get
            Return _ScheduleIRM
        End Get
        Set(ByVal Value As Double)
            _ScheduleIRM = Value
        End Set
    End Property
    Public Property ClaimExpRM() As Double
        Get
            Return _ClaimExpRM
        End Get
        Set(ByVal Value As Double)
            _ClaimExpRM = Value
        End Set
    End Property
    Public Property DateAppRecvd() As Date
        Get
            Return _DateAppRecvd
        End Get
        Set(ByVal Value As Date)
            _DateAppRecvd = Value
        End Set
    End Property
    Public Property DateLossRunRecvd() As Date
        Get
            Return _DateLossRunRecvd
        End Get
        Set(ByVal Value As Date)
            _DateLossRunRecvd = Value
        End Set
    End Property
    Public Property CoverageEffective() As Date
        Get
            Return _CoverageEffective
        End Get
        Set(ByVal Value As Date)
            _CoverageEffective = Value
        End Set
    End Property
    Public Property CoverageExpired() As Date
        Get
            Return _CoverageExpired
        End Get
        Set(ByVal Value As Date)
            _CoverageExpired = Value
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
    Public Property QuoteClass() As String
        Get
            Return _QuoteClass
        End Get
        Set(ByVal Value As String)
            _QuoteClass = Value
        End Set
    End Property
    Public Property IRFileNum() As String
        Get
            Return _IRFileNum
        End Get
        Set(ByVal Value As String)
            _IRFileNum = Value
        End Set
    End Property
    Public Property IRDrawer() As String
        Get
            Return _IRDrawer
        End Get
        Set(ByVal Value As String)
            _IRDrawer = Value
        End Set
    End Property
    Public Property FlagOverRideBy() As String
        Get
            Return _FlagOverRideBy
        End Get
        Set(ByVal Value As String)
            _FlagOverRideBy = Value
        End Set
    End Property
    Public Property RackleyQuoteID() As Long
        Get
            Return _RackleyQuoteID
        End Get
        Set(ByVal Value As Long)
            _RackleyQuoteID = Value
        End Set
    End Property
    Public Property FlagCourtesyFiling() As String
        Get
            Return _FlagCourtesyFiling
        End Get
        Set(ByVal Value As String)
            _FlagCourtesyFiling = Value
        End Set
    End Property
    Public Property FlagRPG() As String
        Get
            Return _FlagRPG
        End Get
        Set(ByVal Value As String)
            _FlagRPG = Value
        End Set
    End Property
    Public Property CurrencyType() As String
        Get
            Return _CurrencyType
        End Get
        Set(ByVal Value As String)
            _CurrencyType = Value
        End Set
    End Property
    Public Property CurrencySymbol() As String
        Get
            Return _CurrencySymbol
        End Get
        Set(ByVal Value As String)
            _CurrencySymbol = Value
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
    Public Property ReservedContractID() As String
        Get
            Return _ReservedContractID
        End Get
        Set(ByVal Value As String)
            _ReservedContractID = Value
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
    Public Property RatingKey_FK() As Long
        Get
            Return _RatingKey_FK
        End Get
        Set(ByVal Value As Long)
            _RatingKey_FK = Value
        End Set
    End Property
    Public Property eAttached() As String
        Get
            Return _eAttached
        End Get
        Set(ByVal Value As String)
            _eAttached = Value
        End Set
    End Property
    Public Property NewField() As String
        Get
            Return _NewField
        End Get
        Set(ByVal Value As String)
            _NewField = Value
        End Set
    End Property
    Public Property TotalCoinsuranceLimit() As Double
        Get
            Return _TotalCoinsuranceLimit
        End Get
        Set(ByVal Value As Double)
            _TotalCoinsuranceLimit = Value
        End Set
    End Property
    Public Property TotalCoinsurancePremium() As Double
        Get
            Return _TotalCoinsurancePremium
        End Get
        Set(ByVal Value As Double)
            _TotalCoinsurancePremium = Value
        End Set
    End Property
    Public Property CurrencyExchRate() As Double
        Get
            Return _CurrencyExchRate
        End Get
        Set(ByVal Value As Double)
            _CurrencyExchRate = Value
        End Set
    End Property
    Public Property Invoiced() As String
        Get
            Return _Invoiced
        End Get
        Set(ByVal Value As String)
            _Invoiced = Value
        End Set
    End Property
    Public Property OtherLead() As String
        Get
            Return _OtherLead
        End Get
        Set(ByVal Value As String)
            _OtherLead = Value
        End Set
    End Property
    Public Property LeadCarrierID() As String
        Get
            Return _LeadCarrierID
        End Get
        Set(ByVal Value As String)
            _LeadCarrierID = Value
        End Set
    End Property
    Public Property RenewTypeID() As String
        Get
            Return _RenewTypeID
        End Get
        Set(ByVal Value As String)
            _RenewTypeID = Value
        End Set
    End Property
    Public Property IsoCode() As String
        Get
            Return _IsoCode
        End Get
        Set(ByVal Value As String)
            _IsoCode = Value
        End Set
    End Property
    Public Property CedingPolicyID() As String
        Get
            Return _CedingPolicyID
        End Get
        Set(ByVal Value As String)
            _CedingPolicyID = Value
        End Set
    End Property
    Public Property CedingPolicyDate() As Date
        Get
            Return _CedingPolicyDate
        End Get
        Set(ByVal Value As Date)
            _CedingPolicyDate = Value
        End Set
    End Property
    Public Property ConversionStatusID() As String
        Get
            Return _ConversionStatusID
        End Get
        Set(ByVal Value As String)
            _ConversionStatusID = Value
        End Set
    End Property
    Public Property FlagTaxExempt() As String
        Get
            Return _FlagTaxExempt
        End Get
        Set(ByVal Value As String)
            _FlagTaxExempt = Value
        End Set
    End Property
    Public Property Units() As Long
        Get
            Return _Units
        End Get
        Set(ByVal Value As Long)
            _Units = Value
        End Set
    End Property
    Public Property SubUnits() As Long
        Get
            Return _SubUnits
        End Get
        Set(ByVal Value As Long)
            _SubUnits = Value
        End Set
    End Property
    Public Property LicenseAgtKey_FK() As Long
        Get
            Return _LicenseAgtKey_FK
        End Get
        Set(ByVal Value As Long)
            _LicenseAgtKey_FK = Value
        End Set
    End Property
    Public Property ContractPlanKey_FK() As Long
        Get
            Return _ContractPlanKey_FK
        End Get
        Set(ByVal Value As Long)
            _ContractPlanKey_FK = Value
        End Set
    End Property
    Public Property AltStatusID() As String
        Get
            Return _AltStatusID
        End Get
        Set(ByVal Value As String)
            _AltStatusID = Value
        End Set
    End Property
    Public Property FlagNonResidentAgt() As String
        Get
            Return _FlagNonResidentAgt
        End Get
        Set(ByVal Value As String)
            _FlagNonResidentAgt = Value
        End Set
    End Property
    Public Property FirewallTeamID() As String
        Get
            Return _FirewallTeamID
        End Get
        Set(ByVal Value As String)
            _FirewallTeamID = Value
        End Set
    End Property
    Public Property CedingPolicyEndDate() As Date
        Get
            Return _CedingPolicyEndDate
        End Get
        Set(ByVal Value As Date)
            _CedingPolicyEndDate = Value
        End Set
    End Property

#End Region
#Region "Methods"
    Public Function Save(ByRef conn As SqlConnection, ByVal pCoverageA As Integer, ByVal pAPDeductible As Integer, ByVal pWindDeductible As Integer, ByVal pCoverage As String) As String
        Dim comm As New SqlCommand("SIU_p_InsertQuote", conn)
        Try
            With comm
                .CommandTimeout = 0
                .CommandType = CommandType.StoredProcedure
                .Parameters.AddWithValue("@QuoteID", _QuoteID)
                .Parameters.AddWithValue("@StatusID", _StatusID)
                .Parameters.AddWithValue("@InsuredID", _InsuredID)
                .Parameters.AddWithValue("@NamedInsured", _NamedInsured)
                .Parameters.AddWithValue("@Address1", _Address1)
                .Parameters.AddWithValue("@Address2", _Address2)
                .Parameters.AddWithValue("@City", _City)
                .Parameters.AddWithValue("@State", _State)
                .Parameters.AddWithValue("@Zip", _Zip)
                .Parameters.AddWithValue("@ProducerID", _ProducerID)
                .Parameters.AddWithValue("@ReferenceID", _ReferenceID)
                .Parameters.AddWithValue("@VersionCounter", _VersionCounter)
                .Parameters.AddWithValue("@Effective", _Effective)
                .Parameters.AddWithValue("@Expiration", _Expiration)
                .Parameters.AddWithValue("@CoverageID", _CoverageID)
                .Parameters.AddWithValue("@TaxState", _TaxState)
                .Parameters.AddWithValue("@AcctExec", _AcctExec)
                .Parameters.AddWithValue("@CsrID", _CsrID)
                .Parameters.AddWithValue("@CompanyID", _CompanyID)
                .Parameters.AddWithValue("@BndPremium", _BndPremium)
                .Parameters.AddWithValue("@Renewal", _Renewal)
                .Parameters.AddWithValue("@ActivePolicyFlag", _ActivePolicyFlag)
                .Parameters.AddWithValue("@ClaimsFlag", _ClaimsFlag)
                .Parameters.AddWithValue("@SuspenseFlag", _SuspenseFlag)
                .Parameters.AddWithValue("@OpenItem", _OpenItem)
                .Parameters.AddWithValue("@Received", _Received)
                .Parameters.AddWithValue("@VersionBound", _VersionBound)
                .Parameters.AddWithValue("@TeamID", _TeamID)
                .Parameters.AddWithValue("@Quoted", _Quoted)
                .Parameters.AddWithValue("@PolicyID", _PolicyID)

                .Parameters.AddWithValue("@ProductID", _ProductID)
                .Parameters.AddWithValue("@SubmitTypeID", _SubmitTypeID)
                .Parameters.AddWithValue("@BndMarketID", _BndMarketID)
                .Parameters.AddWithValue("@PolicyInception", _PolicyInception)

                .ExecuteNonQuery()

                If Math.Abs(pCoverageA) > 0 Then
                    Me.Risk_CommProperty.Address1 = Me.Address1
                    Me.Risk_CommProperty.Address2 = Me.Address2
                    Me.Risk_CommProperty.City = Me.City
                    Me.Risk_CommProperty.State = Me.State
                    Me.Risk_CommProperty.ZipCode = Me.Zip
                    Me.Risk_CommProperty.BuildingValue = pCoverageA
                    Me.Risk_CommProperty.Deduct_AOP = pAPDeductible
                    Me.Risk_CommProperty.Deduct_Wind = pWindDeductible
                    If Math.Abs(pWindDeductible) > 0 Then
                        Me.Risk_CommProperty.FlagWindCOvered = "Y"
                    Else
                        Me.Risk_CommProperty.FlagWindCOvered = "N"
                    End If
                    Return Nothing
                    'Return Me.Risk_CommProperty.Save(conn)
                Else
                    Return ""
                End If
            End With
        Catch ex As Exception
            Return "Quote.Save: " & ex.Message
        End Try
    End Function
    Public Function Load(ByRef conn As SqlConnection, ByVal pQuoteID As String) As String
        Try
            Dim comm As New SqlCommand("SIU_p_GetQuote", conn)
            Dim rs As SqlDataReader

            With comm
                .CommandTimeout = 0
                .CommandType = CommandType.StoredProcedure
                .Parameters.AddWithValue("@QuoteID", pQuoteID)
                rs = .ExecuteReader
            End With
            If rs.Read Then

                _QuoteID = rs("QuoteID")
                _VersionBound = rs("VersionBound")
                _ProducerID = rs("ProducerID")
                _NamedInsured = rs("NamedInsured")
                _TypeID = rs("TypeID")
                _UserID = rs("UserID")
                _Attention = rs("Attention")
                _Received = rs("Received")
                _Acknowledged = rs("Acknowledged")
                _Quoted = rs("Quoted")
                _TeamID = rs("TeamID")
                _DivisionID = rs("DivisionID")
                _StatusID = rs("StatusID")
                _CreatedID = rs("CreatedID")
                _Renewal = rs("Renewal")
                _OldPolicyID = rs("OldPolicyID")
                _OldVersion = rs("OldVersion")
                _OldExpiration = rs("OldExpiration")
                _OpenItem = rs("OpenItem")
                _Notes = rs("Notes")
                _PolicyID = rs("PolicyID")
                _VersionCounter = rs("VersionCounter")
                _InsuredID = rs("InsuredID")
                _Description = rs("Description")
                _FileLocation = rs("FileLocation")
                _Address1 = rs("Address1")
                _Address2 = rs("Address2")
                _City = rs("City")
                _State = rs("State")
                _Zip = rs("Zip")
                _Bound = rs("Bound")
                _Submitted = rs("Submitted")
                _SubmitType = rs("SubmitType")
                _NoteAttached = rs("NoteAttached")
                _AcctExec = rs("AcctExec")
                _InsuredInterest = rs("InsuredInterest")
                _RiskInformation = rs("RiskInformation")
                _EC = rs("EC")
                _BndPremium = rs("BndPremium")
                _BndFee = rs("BndFee")
                _CompanyID = rs("CompanyID")
                _ProductID = rs("ProductID")
                _Effective = rs("Effective")
                _Expiration = rs("Expiration")
                _Setup = rs("Setup")
                _PolicyMailOut = rs("PolicyMailOut")
                _BinderRev = rs("BinderRev")
                _PriorCarrier = rs("PriorCarrier")
                _TargetPremium = rs("TargetPremium")
                _CsrID = rs("CsrID")
                _PolicyVer = rs("PolicyVer")
                _OldQuoteID = rs("OldQuoteID")
                _PolicyGrpID = rs("PolicyGrpID")
                _PendingSuspenseID = rs("PendingSuspenseID")
                _ReferenceID = rs("ReferenceID")
                _MapToID = rs("MapToID")
                _SubmitGrpID = rs("SubmitGrpID")
                _AcctAsst = rs("AcctAsst")
                _TaxState = rs("TaxState")
                _SicID = rs("SicID")
                _CoverageID = rs("CoverageID")
                _OldPremium = rs("OldPremium")
                _AddressID = rs("AddressID")
                _OldEffective = rs("OldEffective")
                _TaxBasis = rs("TaxBasis")
                _QuoteRequiredBy = rs("QuoteRequiredBy")
                _RequiredLimits = rs("RequiredLimits")
                _RequiredDeduct = rs("RequiredDeduct")
                _Retroactive = rs("Retroactive")
                _PrevCancelFlag = rs("PrevCancelFlag")
                _PrevNonRenew = rs("PrevNonRenew")
                _PriorPremium = rs("PriorPremium")
                _PriorLimits = rs("PriorLimits")
                _UWCheckList = rs("UWCheckList")
                _FileSetup = rs("FileSetup")
                _ContactID = rs("ContactID")
                _SuspenseFlag = rs("SuspenseFlag")
                _PriorDeductible = rs("PriorDeductible")
                _CategoryID = rs("CategoryID")
                _StructureID = rs("StructureID")
                _RenewalStatusID = rs("RenewalStatusID")
                _ClaimsFlag = rs("ClaimsFlag")
                _ActivePolicyFlag = rs("ActivePolicyFlag")
                _Assets = rs("Assets")
                _PublicEntity = rs("PublicEntity")
                _VentureID = rs("VentureID")
                _IncorporatedState = rs("IncorporatedState")
                _ReInsuranceFlag = rs("ReInsuranceFlag")
                _TaxedPaidBy = rs("TaxedPaidBy")
                _LayeredCoverage = rs("LayeredCoverage")
                _Employees = rs("Employees")
                _Stock_52wk = rs("Stock_52wk")
                _NetIncome = rs("NetIncome")
                _LossHistory = rs("LossHistory")
                _PriorLimitsNew = rs("PriorLimitsNew")
                _LargeLossHistory = rs("LargeLossHistory")
                _DateOfApp = rs("DateOfApp")
                _Stock_High = rs("Stock_High")
                _Stock_Low = rs("Stock_Low")
                _Stock_Current = rs("Stock_Current")
                _MarketCap = rs("MarketCap")
                _Exposures = rs("Exposures")
                _AIM_TransDate = rs("AIM_TransDate")
                _LostBusinessFlag = rs("LostBusinessFlag")
                _YearEst = rs("YearEst")
                _LostBusiness_Carrier = rs("LostBusiness_Carrier")
                _LostBusiness_Premium = rs("LostBusiness_Premium")
                _AccountKey_FK = rs("AccountKey_FK")
                _FlagRewrite = rs("FlagRewrite")
                _flagWIP = rs("flagWIP")
                _RenewalQuoteID = rs("RenewalQuoteID")
                _QuoteDueDate = rs("QuoteDueDate")
                _QuoteStatus = rs("QuoteStatus")
                _BinderExpires = rs("BinderExpires")
                _TIV = rs("TIV")
                _InvoicedPremium = rs("InvoicedPremium")
                _InvoicedFee = rs("InvoicedFee")
                _InvoicedCommRev = rs("InvoicedCommRev")
                _SplitAccount = rs("SplitAccount")
                _FileCloseReason = rs("FileCloseReason")
                _FileCloseReasonID = rs("FileCloseReasonID")
                _SourceOfLeadID = rs("SourceOfLeadID")
                _ServiceUWID = rs("ServiceUWID")
                _SubmitTypeID = rs("SubmitTypeID")
                _SubProducerID = rs("SubProducerID")
                _AgtAccountNumber = rs("AgtAccountNumber")
                _BndMarketID = rs("BndMarketID")
                _RefQuoteID = rs("RefQuoteID")
                _FlagHeldFile = rs("FlagHeldFile")
                _HeldFileMessage = rs("HeldFileMessage")
                _TermPremium = rs("TermPremium")
                _ProcessBatchKey_FK = rs("ProcessBatchKey_FK")
                _PolicyInception = rs("PolicyInception")
                _QuoteClassID = rs("ClassID")
                _ScheduleIRM = rs("ScheduleIRM")
                _ClaimExpRM = rs("ClaimExpRM")
                _DateAppRecvd = rs("DateAppRecvd")
                _DateLossRunRecvd = rs("DateLossRunRecvd")
                _CoverageEffective = rs("CoverageEffective")
                _CoverageExpired = rs("CoverageExpired")
                _SLA = rs("SLA")
                _QuoteClass = rs("Class")
                _IRFileNum = rs("IRFileNum")
                _IRDrawer = rs("IRDrawer")
                _FlagOverRideBy = rs("FlagOverRideBy")
                _RackleyQuoteID = rs("RackleyQuoteID")
                _FlagCourtesyFiling = rs("FlagCourtesyFiling")
                _FlagRPG = rs("FlagRPG")
                _CurrencyType = rs("CurrencyType")
                _CurrencySymbol = rs("CurrencySymbol")
                _FileNo = rs("FileNo")
                _UserDefinedStr1 = rs("UserDefinedStr1")
                _UserDefinedStr2 = rs("UserDefinedStr2")
                _UserDefinedStr3 = rs("UserDefinedStr3")
                _UserDefinedStr4 = rs("UserDefinedStr4")
                _UserDefinedDate1 = rs("UserDefinedDate1")
                _UserDefinedValue1 = rs("UserDefinedValue1")
                _ReservedContractID = rs("ReservedContractID")
                _CountryID = rs("CountryID")
                _RatingKey_FK = rs("RatingKey_FK")
                _eAttached = rs("eAttached")
                _NewField = rs("NewField")
                _TotalCoinsuranceLimit = rs("TotalCoinsuranceLimit")
                _TotalCoinsurancePremium = rs("TotalCoinsurancePremium")
                _CurrencyExchRate = rs("CurrencyExchRate")
                _Invoiced = rs("Invoiced")
                _OtherLead = rs("OtherLead")
                _LeadCarrierID = rs("LeadCarrierID")
                _RenewTypeID = rs("RenewTypeID")
                _IsoCode = rs("IsoCode")
                _CedingPolicyID = rs("CedingPolicyID")
                _CedingPolicyDate = rs("CedingPolicyDate")
                _ConversionStatusID = rs("ConversionStatusID")
                _FlagTaxExempt = rs("FlagTaxExempt")
                _Units = rs("Units")
                _SubUnits = rs("SubUnits")
                _LicenseAgtKey_FK = rs("LicenseAgtKey_FK")
                _ContractPlanKey_FK = rs("ContractPlanKey_FK")
                _AltStatusID = rs("AltStatusID")
                _FlagNonResidentAgt = rs("FlagNonResidentAgt")
                _FirewallTeamID = rs("FirewallTeamID")
                _CedingPolicyEndDate = rs("CedingPolicyEndDate")

            End If
            rs.Close()

            Return _Risk_CommProperty.Load(conn, _QuoteID)

        Catch ex As Exception
            Return "Quote: " & ex.Message
        End Try
    End Function

#End Region
End Class
