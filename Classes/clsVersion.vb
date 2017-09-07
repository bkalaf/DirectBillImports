Imports System.Data.SqlClient
Imports DirectBillImports.Common
Public Class clsVersion
#Region "Local Variables"
    Private _QuoteID As String
    Private _VerOriginal As String
    Private _Version As String
    Private _LobID As String
    Private _LobSubID As String
    Private _CompanyID As String
    Private _ProductID As String
    Private _Premium As Double
    Private _Non_Premium As Double
    Private _Misc_Premium As Double
    Private _NonTax_Premium As Double
    Private _Quoted As Date
    Private _Expires As Date
    Private _Limits As String
    Private _Subject As String
    Private _Endorsement As String
    Private _Financed As String
    Private _Taxed As String
    Private _MEP As String
    Private _Rate As String
    Private _GrossComm As Double
    Private _AgentComm As Double
    Private _Brokerage As String
    Private _Deductible As String
    Private _CoInsure As String
    Private _StatusID As String
    Private _ReasonID As String
    Private _SubmitDate As Date
    Private _SubmitPOC As String
    Private _MarketID As String
    Private _Apportionment As Double
    Private _Tax1 As Double
    Private _Tax2 As Double
    Private _Tax3 As Double
    Private _Tax4 As Double
    Private _FormID As String
    Private _RateInfo As String
    Private _Indicator As String
    Private _PendingSuspenseID As String
    Private _CommPaid As Double
    Private _AggregateLimits As Double
    Private _DeductibleVal As Double
    Private _BoundFlag As String
    Private _DirectBillFlag As String
    Private _ProposedEffective As Date
    Private _ProposedExpiration As Date
    Private _ProposedTerm As Long
    Private _Retroactive As Date
    Private _RetroPeriod As String
    Private _UnderLyingCoverage As String
    Private _MultiOption As String
    Private _MiscPrem1 As Double
    Private _MiscPrem2 As Double
    Private _MiscPrem3 As Double
    Private _NonTax1 As Double
    Private _NonTax2 As Double
    Private _NonPrem1 As Double
    Private _NonPrem2 As Double
    Private _PaymentRecv As Double
    Private _PremDownPayment As Double
    Private _Valuation As String
    Private _Retention As String
    Private _AIM_TransDate As Date
    Private _InvoiceCodes As String
    Private _TaxDistrib As String
    Private _PremDistrib As String
    Private _CAP_Limit As Double
    Private _EPL_Limit As Double
    Private _TakenOut_RatedTerm As Long
    Private _PolicyTerm As String
    Private _PolicyForm As String
    Private _BillToCompanyID As String
    Private _StatementKey_FK As Long
    Private _PaymentKey_FK As Long
    Private _CommRecvd As Double
    Private _VersionID As String
    Private _MarketContactKey_FK As Long
    Private _TIV As Double
    Private _CompanyFees As Double
    Private _UnderLyingLimitsSum As Double
    Private _PunitiveDamage As Double
    Private _ThirdPartyLimits As Double
    Private _AnnualPremium As Double
    Private _AnnualFees As Double
    Private _FlagCollectMuniTax As String
    Private _TrueExpire As Date
    Private _WrittenLimits As Double
    Private _AttachPoint As Double
    Private _LineSlip As Double
    Private _CoverageFormID As String
    Private _PositionID As String
    Private _LobDistrib As String
    Private _TotalTax As Double
    Private _Total As Double
    Private _TotalAmount As Double
    Private _TaxesPaidBy As String
    Private _ResubmitDate As Date
    Private _FeeSchedule As String
    Private _LobDistribSched As String
    Private _DeductType As String
    Private _PremiumFinanceFee As Double
    Private _LOB_Field1 As String
    Private _LOB_Field2 As String
    Private _LOB_Field3 As String
    Private _LOB_Flag1 As String
    Private _LOB_Prem1 As Double
    Private _LOB_Prem2 As Double
    Private _LOB_Prem3 As Double
    Private _LOB_Limit1 As String
    Private _LOB_Limit2 As String
    Private _LOB_Limit3 As String
    Private _LOB_Limit4 As String
    Private _LOB_Limit5 As String
    Private _LOB_Limit6 As String
    Private _LOB_Deduct1 As String
    Private _LOB_Deduct2 As String
    Private _LOB_Limit1Value As Double
    Private _LOB_Limit2Value As Double
    Private _LOB_Limit3Value As Double
    Private _LOB_Limit4Value As Double
    Private _LOB_Limit5Value As Double
    Private _LOB_Limit6Value As Double
    Private _LOB_Deduct1Value As Double
    Private _LOB_Deduct2Value As Double
    Private _TaxesPaidByID As String
    Private _FlagMultiStateTax As String
    Private _MultiStateDistrib As String
    Private _AdmittedPremium As Double
    Private _RatedPremium As Double
    Private _APR As Double
    Private _AmountFinanced As Double
    Private _DownPayment As Double
    Private _Payments As Double
    Private _FinCharge As Double
    Private _TotalPayment As Double
    Private _NumPayments As Long
    Private _FinanceDueDate As Date
    Private _ReferenceKey_FK As Long
    Private _RemitAmount As Double
    Private _CollectAmount As Double
    Private _DownFactor As Double
    Private _TerrorActPremium As Double
    Private _TerrorActGrossComm As Double
    Private _TerrorActAgentComm As Double
    Private _TerrorActMEP As String
    Private _TerrorActStatus As String
    Private _FlagOverrideCalc As String
    Private _TerrorTaxes As Double
    Private _FlagFinanceWithTRIA As String
    Private _FlagMultiOption As String
    Private _FlagFeeCalc As String
    Private _ParticipantCo1ID As String
    Private _ParticipantCo2ID As String
    Private _ParticipantCo3ID As String
    Private _UserDefinedStr1 As String
    Private _UserDefinedStr2 As String
    Private _UserDefinedStr3 As String
    Private _UserDefinedStr4 As String
    Private _UserDefinedDate1 As Date
    Private _UserDefinedValue1 As Double
    Private _LOB_Coverage1 As String
    Private _LOB_Coverage2 As String
    Private _LOB_Coverage3 As String
    Private _LOB_Coverage4 As String
    Private _LOB_Coverage5 As String
    Private _LOB_Coverage6 As String
    Private _LOB_DeductType1 As String
    Private _LOB_DeductType2 As String
    Private _DeclinationReasonID As String
    Private _ERPOption As String
    Private _ERPDays As Long
    Private _ERPPercent As Double
    Private _ERPPremium As Double
    Private _TaxwoTRIA1 As Double
    Private _TaxwoTRIA2 As Double
    Private _TaxwoTRIA3 As Double
    Private _TaxwoTRIA4 As Double
    Private _LOB_Prem4 As Double
    Private _LOB_Coverage7 As String
    Private _LOB_Coverage8 As String
    Private _LOB_Limit7 As String
    Private _LOB_Limit8 As String
    Private _LOB_Limit7Value As Double
    Private _LOB_Limit8Value As Double
    Private _LOB_Prem5 As Double
    Private _LOB_Prem6 As Double
    Private _LOB_Prem7 As Double
    Private _LOB_Prem8 As Double
    Private _CoverageList As String
    Private _DocucorpFormList As String
    Private _TerrorActPremium_GL As Double
    Private _FlagRecalcTaxes As String
    Private _DateMktResponseRecvd As Date
    Private _CancelClause As String
    Private _PremiumProperty As Double
    Private _PremiumLiability As Double
    Private _PremiumOther As Double
    Private _EndorsementKey_FK As Long

#End Region
#Region "Properties"

    Public Property QuoteID() As String
        Get
            Return _QuoteID
        End Get
        Set(ByVal Value As String)
            _QuoteID = Value
        End Set
    End Property
    Public Property VerOriginal() As String
        Get
            Return _VerOriginal
        End Get
        Set(ByVal Value As String)
            _VerOriginal = Value
        End Set
    End Property
    Public Property Version() As String
        Get
            Return _Version
        End Get
        Set(ByVal Value As String)
            _Version = Value
        End Set
    End Property
    Public Property LobID() As String
        Get
            Return _LobID
        End Get
        Set(ByVal Value As String)
            _LobID = Value
        End Set
    End Property
    Public Property LobSubID() As String
        Get
            Return _LobSubID
        End Get
        Set(ByVal Value As String)
            _LobSubID = Value
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
    Public Property Premium() As Double
        Get
            Return _Premium
        End Get
        Set(ByVal Value As Double)
            _Premium = Value
        End Set
    End Property
    Public Property Non_Premium() As Double
        Get
            Return _Non_Premium
        End Get
        Set(ByVal Value As Double)
            _Non_Premium = Value
        End Set
    End Property
    Public Property Misc_Premium() As Double
        Get
            Return _Misc_Premium
        End Get
        Set(ByVal Value As Double)
            _Misc_Premium = Value
        End Set
    End Property
    Public Property NonTax_Premium() As Double
        Get
            Return _NonTax_Premium
        End Get
        Set(ByVal Value As Double)
            _NonTax_Premium = Value
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
    Public Property Expires() As Date
        Get
            Return _Expires
        End Get
        Set(ByVal Value As Date)
            _Expires = Value
        End Set
    End Property
    Public Property Limits() As String
        Get
            Return _Limits
        End Get
        Set(ByVal Value As String)
            _Limits = Value
        End Set
    End Property
    Public Property Subject() As String
        Get
            Return _Subject
        End Get
        Set(ByVal Value As String)
            _Subject = Value
        End Set
    End Property
    Public Property Endorsement() As String
        Get
            Return _Endorsement
        End Get
        Set(ByVal Value As String)
            _Endorsement = Value
        End Set
    End Property
    Public Property Financed() As String
        Get
            Return _Financed
        End Get
        Set(ByVal Value As String)
            _Financed = Value
        End Set
    End Property
    Public Property Taxed() As String
        Get
            Return _Taxed
        End Get
        Set(ByVal Value As String)
            _Taxed = Value
        End Set
    End Property
    Public Property MEP() As String
        Get
            Return _MEP
        End Get
        Set(ByVal Value As String)
            _MEP = Value
        End Set
    End Property
    Public Property Rate() As String
        Get
            Return _Rate
        End Get
        Set(ByVal Value As String)
            _Rate = Value
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
    Public Property Brokerage() As String
        Get
            Return _Brokerage
        End Get
        Set(ByVal Value As String)
            _Brokerage = Value
        End Set
    End Property
    Public Property Deductible() As String
        Get
            Return _Deductible
        End Get
        Set(ByVal Value As String)
            _Deductible = Value
        End Set
    End Property
    Public Property CoInsure() As String
        Get
            Return _CoInsure
        End Get
        Set(ByVal Value As String)
            _CoInsure = Value
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
    Public Property ReasonID() As String
        Get
            Return _ReasonID
        End Get
        Set(ByVal Value As String)
            _ReasonID = Value
        End Set
    End Property
    Public Property SubmitDate() As Date
        Get
            Return _SubmitDate
        End Get
        Set(ByVal Value As Date)
            _SubmitDate = Value
        End Set
    End Property
    Public Property SubmitPOC() As String
        Get
            Return _SubmitPOC
        End Get
        Set(ByVal Value As String)
            _SubmitPOC = Value
        End Set
    End Property
    Public Property MarketID() As String
        Get
            Return _MarketID
        End Get
        Set(ByVal Value As String)
            _MarketID = Value
        End Set
    End Property
    Public Property Apportionment() As Double
        Get
            Return _Apportionment
        End Get
        Set(ByVal Value As Double)
            _Apportionment = Value
        End Set
    End Property
    Public Property Tax1() As Double
        Get
            Return _Tax1
        End Get
        Set(ByVal Value As Double)
            _Tax1 = Value
        End Set
    End Property
    Public Property Tax2() As Double
        Get
            Return _Tax2
        End Get
        Set(ByVal Value As Double)
            _Tax2 = Value
        End Set
    End Property
    Public Property Tax3() As Double
        Get
            Return _Tax3
        End Get
        Set(ByVal Value As Double)
            _Tax3 = Value
        End Set
    End Property
    Public Property Tax4() As Double
        Get
            Return _Tax4
        End Get
        Set(ByVal Value As Double)
            _Tax4 = Value
        End Set
    End Property
    Public Property FormID() As String
        Get
            Return _FormID
        End Get
        Set(ByVal Value As String)
            _FormID = Value
        End Set
    End Property
    Public Property RateInfo() As String
        Get
            Return _RateInfo
        End Get
        Set(ByVal Value As String)
            _RateInfo = Value
        End Set
    End Property
    Public Property Indicator() As String
        Get
            Return _Indicator
        End Get
        Set(ByVal Value As String)
            _Indicator = Value
        End Set
    End Property
    Public Property PendingSuspenseID() As String
        Get
            Return _PendingSuspenseID
        End Get
        Set(ByVal Value As String)
            _PendingSuspenseID = Value
        End Set
    End Property
    Public Property CommPaid() As Double
        Get
            Return _CommPaid
        End Get
        Set(ByVal Value As Double)
            _CommPaid = Value
        End Set
    End Property
    Public Property AggregateLimits() As Double
        Get
            Return _AggregateLimits
        End Get
        Set(ByVal Value As Double)
            _AggregateLimits = Value
        End Set
    End Property
    Public Property DeductibleVal() As Double
        Get
            Return _DeductibleVal
        End Get
        Set(ByVal Value As Double)
            _DeductibleVal = Value
        End Set
    End Property
    Public Property BoundFlag() As String
        Get
            Return _BoundFlag
        End Get
        Set(ByVal Value As String)
            _BoundFlag = Value
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
    Public Property ProposedEffective() As Date
        Get
            Return _ProposedEffective
        End Get
        Set(ByVal Value As Date)
            _ProposedEffective = Value
        End Set
    End Property
    Public Property ProposedExpiration() As Date
        Get
            Return _ProposedExpiration
        End Get
        Set(ByVal Value As Date)
            _ProposedExpiration = Value
        End Set
    End Property
    Public Property ProposedTerm() As Long
        Get
            Return _ProposedTerm
        End Get
        Set(ByVal Value As Long)
            _ProposedTerm = Value
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
    Public Property RetroPeriod() As String
        Get
            Return _RetroPeriod
        End Get
        Set(ByVal Value As String)
            _RetroPeriod = Value
        End Set
    End Property
    Public Property UnderLyingCoverage() As String
        Get
            Return _UnderLyingCoverage
        End Get
        Set(ByVal Value As String)
            _UnderLyingCoverage = Value
        End Set
    End Property
    Public Property MultiOption() As String
        Get
            Return _MultiOption
        End Get
        Set(ByVal Value As String)
            _MultiOption = Value
        End Set
    End Property
    Public Property MiscPrem1() As Double
        Get
            Return _MiscPrem1
        End Get
        Set(ByVal Value As Double)
            _MiscPrem1 = Value
        End Set
    End Property
    Public Property MiscPrem2() As Double
        Get
            Return _MiscPrem2
        End Get
        Set(ByVal Value As Double)
            _MiscPrem2 = Value
        End Set
    End Property
    Public Property MiscPrem3() As Double
        Get
            Return _MiscPrem3
        End Get
        Set(ByVal Value As Double)
            _MiscPrem3 = Value
        End Set
    End Property
    Public Property NonTax1() As Double
        Get
            Return _NonTax1
        End Get
        Set(ByVal Value As Double)
            _NonTax1 = Value
        End Set
    End Property
    Public Property NonTax2() As Double
        Get
            Return _NonTax2
        End Get
        Set(ByVal Value As Double)
            _NonTax2 = Value
        End Set
    End Property
    Public Property NonPrem1() As Double
        Get
            Return _NonPrem1
        End Get
        Set(ByVal Value As Double)
            _NonPrem1 = Value
        End Set
    End Property
    Public Property NonPrem2() As Double
        Get
            Return _NonPrem2
        End Get
        Set(ByVal Value As Double)
            _NonPrem2 = Value
        End Set
    End Property
    Public Property PaymentRecv() As Double
        Get
            Return _PaymentRecv
        End Get
        Set(ByVal Value As Double)
            _PaymentRecv = Value
        End Set
    End Property
    Public Property PremDownPayment() As Double
        Get
            Return _PremDownPayment
        End Get
        Set(ByVal Value As Double)
            _PremDownPayment = Value
        End Set
    End Property
    Public Property Valuation() As String
        Get
            Return _Valuation
        End Get
        Set(ByVal Value As String)
            _Valuation = Value
        End Set
    End Property
    Public Property Retention() As String
        Get
            Return _Retention
        End Get
        Set(ByVal Value As String)
            _Retention = Value
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
    Public Property InvoiceCodes() As String
        Get
            Return _InvoiceCodes
        End Get
        Set(ByVal Value As String)
            _InvoiceCodes = Value
        End Set
    End Property
    Public Property TaxDistrib() As String
        Get
            Return _TaxDistrib
        End Get
        Set(ByVal Value As String)
            _TaxDistrib = Value
        End Set
    End Property
    Public Property PremDistrib() As String
        Get
            Return _PremDistrib
        End Get
        Set(ByVal Value As String)
            _PremDistrib = Value
        End Set
    End Property
    Public Property CAP_Limit() As Double
        Get
            Return _CAP_Limit
        End Get
        Set(ByVal Value As Double)
            _CAP_Limit = Value
        End Set
    End Property
    Public Property EPL_Limit() As Double
        Get
            Return _EPL_Limit
        End Get
        Set(ByVal Value As Double)
            _EPL_Limit = Value
        End Set
    End Property
    Public Property TakenOut_RatedTerm() As Long
        Get
            Return _TakenOut_RatedTerm
        End Get
        Set(ByVal Value As Long)
            _TakenOut_RatedTerm = Value
        End Set
    End Property
    Public Property PolicyTerm() As String
        Get
            Return _PolicyTerm
        End Get
        Set(ByVal Value As String)
            _PolicyTerm = Value
        End Set
    End Property
    Public Property PolicyForm() As String
        Get
            Return _PolicyForm
        End Get
        Set(ByVal Value As String)
            _PolicyForm = Value
        End Set
    End Property
    Public Property BillToCompanyID() As String
        Get
            Return _BillToCompanyID
        End Get
        Set(ByVal Value As String)
            _BillToCompanyID = Value
        End Set
    End Property
    Public Property StatementKey_FK() As Long
        Get
            Return _StatementKey_FK
        End Get
        Set(ByVal Value As Long)
            _StatementKey_FK = Value
        End Set
    End Property
    Public Property PaymentKey_FK() As Long
        Get
            Return _PaymentKey_FK
        End Get
        Set(ByVal Value As Long)
            _PaymentKey_FK = Value
        End Set
    End Property
    Public Property CommRecvd() As Double
        Get
            Return _CommRecvd
        End Get
        Set(ByVal Value As Double)
            _CommRecvd = Value
        End Set
    End Property
    Public Property VersionID() As String
        Get
            Return _VersionID
        End Get
        Set(ByVal Value As String)
            _VersionID = Value
        End Set
    End Property
    Public Property MarketContactKey_FK() As Long
        Get
            Return _MarketContactKey_FK
        End Get
        Set(ByVal Value As Long)
            _MarketContactKey_FK = Value
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
    Public Property CompanyFees() As Double
        Get
            Return _CompanyFees
        End Get
        Set(ByVal Value As Double)
            _CompanyFees = Value
        End Set
    End Property
    Public Property UnderLyingLimitsSum() As Double
        Get
            Return _UnderLyingLimitsSum
        End Get
        Set(ByVal Value As Double)
            _UnderLyingLimitsSum = Value
        End Set
    End Property
    Public Property PunitiveDamage() As Double
        Get
            Return _PunitiveDamage
        End Get
        Set(ByVal Value As Double)
            _PunitiveDamage = Value
        End Set
    End Property
    Public Property ThirdPartyLimits() As Double
        Get
            Return _ThirdPartyLimits
        End Get
        Set(ByVal Value As Double)
            _ThirdPartyLimits = Value
        End Set
    End Property
    Public Property AnnualPremium() As Double
        Get
            Return _AnnualPremium
        End Get
        Set(ByVal Value As Double)
            _AnnualPremium = Value
        End Set
    End Property
    Public Property AnnualFees() As Double
        Get
            Return _AnnualFees
        End Get
        Set(ByVal Value As Double)
            _AnnualFees = Value
        End Set
    End Property
    Public Property FlagCollectMuniTax() As String
        Get
            Return _FlagCollectMuniTax
        End Get
        Set(ByVal Value As String)
            _FlagCollectMuniTax = Value
        End Set
    End Property
    Public Property TrueExpire() As Date
        Get
            Return _TrueExpire
        End Get
        Set(ByVal Value As Date)
            _TrueExpire = Value
        End Set
    End Property
    Public Property WrittenLimits() As Double
        Get
            Return _WrittenLimits
        End Get
        Set(ByVal Value As Double)
            _WrittenLimits = Value
        End Set
    End Property
    Public Property AttachPoint() As Double
        Get
            Return _AttachPoint
        End Get
        Set(ByVal Value As Double)
            _AttachPoint = Value
        End Set
    End Property
    Public Property LineSlip() As Double
        Get
            Return _LineSlip
        End Get
        Set(ByVal Value As Double)
            _LineSlip = Value
        End Set
    End Property
    Public Property CoverageFormID() As String
        Get
            Return _CoverageFormID
        End Get
        Set(ByVal Value As String)
            _CoverageFormID = Value
        End Set
    End Property
    Public Property PositionID() As String
        Get
            Return _PositionID
        End Get
        Set(ByVal Value As String)
            _PositionID = Value
        End Set
    End Property
    Public Property LobDistrib() As String
        Get
            Return _LobDistrib
        End Get
        Set(ByVal Value As String)
            _LobDistrib = Value
        End Set
    End Property
    Public Property TotalTax() As Double
        Get
            Return _TotalTax
        End Get
        Set(ByVal Value As Double)
            _TotalTax = Value
        End Set
    End Property
    Public Property Total() As Double
        Get
            Return _Total
        End Get
        Set(ByVal Value As Double)
            _Total = Value
        End Set
    End Property
    Public Property TotalAmount() As Double
        Get
            Return _TotalAmount
        End Get
        Set(ByVal Value As Double)
            _TotalAmount = Value
        End Set
    End Property
    Public Property TaxesPaidBy() As String
        Get
            Return _TaxesPaidBy
        End Get
        Set(ByVal Value As String)
            _TaxesPaidBy = Value
        End Set
    End Property
    Public Property ResubmitDate() As Date
        Get
            Return _ResubmitDate
        End Get
        Set(ByVal Value As Date)
            _ResubmitDate = Value
        End Set
    End Property
    Public Property FeeSchedule() As String
        Get
            Return _FeeSchedule
        End Get
        Set(ByVal Value As String)
            _FeeSchedule = Value
        End Set
    End Property
    Public Property LobDistribSched() As String
        Get
            Return _LobDistribSched
        End Get
        Set(ByVal Value As String)
            _LobDistribSched = Value
        End Set
    End Property
    Public Property DeductType() As String
        Get
            Return _DeductType
        End Get
        Set(ByVal Value As String)
            _DeductType = Value
        End Set
    End Property
    Public Property PremiumFinanceFee() As Double
        Get
            Return _PremiumFinanceFee
        End Get
        Set(ByVal Value As Double)
            _PremiumFinanceFee = Value
        End Set
    End Property
    Public Property LOB_Field1() As String
        Get
            Return _LOB_Field1
        End Get
        Set(ByVal Value As String)
            _LOB_Field1 = Value
        End Set
    End Property
    Public Property LOB_Field2() As String
        Get
            Return _LOB_Field2
        End Get
        Set(ByVal Value As String)
            _LOB_Field2 = Value
        End Set
    End Property
    Public Property LOB_Field3() As String
        Get
            Return _LOB_Field3
        End Get
        Set(ByVal Value As String)
            _LOB_Field3 = Value
        End Set
    End Property
    Public Property LOB_Flag1() As String
        Get
            Return _LOB_Flag1
        End Get
        Set(ByVal Value As String)
            _LOB_Flag1 = Value
        End Set
    End Property
    Public Property LOB_Prem1() As Double
        Get
            Return _LOB_Prem1
        End Get
        Set(ByVal Value As Double)
            _LOB_Prem1 = Value
        End Set
    End Property
    Public Property LOB_Prem2() As Double
        Get
            Return _LOB_Prem2
        End Get
        Set(ByVal Value As Double)
            _LOB_Prem2 = Value
        End Set
    End Property
    Public Property LOB_Prem3() As Double
        Get
            Return _LOB_Prem3
        End Get
        Set(ByVal Value As Double)
            _LOB_Prem3 = Value
        End Set
    End Property
    Public Property LOB_Limit1() As String
        Get
            Return _LOB_Limit1
        End Get
        Set(ByVal Value As String)
            _LOB_Limit1 = Value
        End Set
    End Property
    Public Property LOB_Limit2() As String
        Get
            Return _LOB_Limit2
        End Get
        Set(ByVal Value As String)
            _LOB_Limit2 = Value
        End Set
    End Property
    Public Property LOB_Limit3() As String
        Get
            Return _LOB_Limit3
        End Get
        Set(ByVal Value As String)
            _LOB_Limit3 = Value
        End Set
    End Property
    Public Property LOB_Limit4() As String
        Get
            Return _LOB_Limit4
        End Get
        Set(ByVal Value As String)
            _LOB_Limit4 = Value
        End Set
    End Property
    Public Property LOB_Limit5() As String
        Get
            Return _LOB_Limit5
        End Get
        Set(ByVal Value As String)
            _LOB_Limit5 = Value
        End Set
    End Property
    Public Property LOB_Limit6() As String
        Get
            Return _LOB_Limit6
        End Get
        Set(ByVal Value As String)
            _LOB_Limit6 = Value
        End Set
    End Property
    Public Property LOB_Deduct1() As String
        Get
            Return _LOB_Deduct1
        End Get
        Set(ByVal Value As String)
            _LOB_Deduct1 = Value
        End Set
    End Property
    Public Property LOB_Deduct2() As String
        Get
            Return _LOB_Deduct2
        End Get
        Set(ByVal Value As String)
            _LOB_Deduct2 = Value
        End Set
    End Property
    Public Property LOB_Limit1Value() As Double
        Get
            Return _LOB_Limit1Value
        End Get
        Set(ByVal Value As Double)
            _LOB_Limit1Value = Value
        End Set
    End Property
    Public Property LOB_Limit2Value() As Double
        Get
            Return _LOB_Limit2Value
        End Get
        Set(ByVal Value As Double)
            _LOB_Limit2Value = Value
        End Set
    End Property
    Public Property LOB_Limit3Value() As Double
        Get
            Return _LOB_Limit3Value
        End Get
        Set(ByVal Value As Double)
            _LOB_Limit3Value = Value
        End Set
    End Property
    Public Property LOB_Limit4Value() As Double
        Get
            Return _LOB_Limit4Value
        End Get
        Set(ByVal Value As Double)
            _LOB_Limit4Value = Value
        End Set
    End Property
    Public Property LOB_Limit5Value() As Double
        Get
            Return _LOB_Limit5Value
        End Get
        Set(ByVal Value As Double)
            _LOB_Limit5Value = Value
        End Set
    End Property
    Public Property LOB_Limit6Value() As Double
        Get
            Return _LOB_Limit6Value
        End Get
        Set(ByVal Value As Double)
            _LOB_Limit6Value = Value
        End Set
    End Property
    Public Property LOB_Deduct1Value() As Double
        Get
            Return _LOB_Deduct1Value
        End Get
        Set(ByVal Value As Double)
            _LOB_Deduct1Value = Value
        End Set
    End Property
    Public Property LOB_Deduct2Value() As Double
        Get
            Return _LOB_Deduct2Value
        End Get
        Set(ByVal Value As Double)
            _LOB_Deduct2Value = Value
        End Set
    End Property
    Public Property TaxesPaidByID() As String
        Get
            Return _TaxesPaidByID
        End Get
        Set(ByVal Value As String)
            _TaxesPaidByID = Value
        End Set
    End Property
    Public Property FlagMultiStateTax() As String
        Get
            Return _FlagMultiStateTax
        End Get
        Set(ByVal Value As String)
            _FlagMultiStateTax = Value
        End Set
    End Property
    Public Property MultiStateDistrib() As String
        Get
            Return _MultiStateDistrib
        End Get
        Set(ByVal Value As String)
            _MultiStateDistrib = Value
        End Set
    End Property
    Public Property AdmittedPremium() As Double
        Get
            Return _AdmittedPremium
        End Get
        Set(ByVal Value As Double)
            _AdmittedPremium = Value
        End Set
    End Property
    Public Property RatedPremium() As Double
        Get
            Return _RatedPremium
        End Get
        Set(ByVal Value As Double)
            _RatedPremium = Value
        End Set
    End Property
    Public Property APR() As Double
        Get
            Return _APR
        End Get
        Set(ByVal Value As Double)
            _APR = Value
        End Set
    End Property
    Public Property AmountFinanced() As Double
        Get
            Return _AmountFinanced
        End Get
        Set(ByVal Value As Double)
            _AmountFinanced = Value
        End Set
    End Property
    Public Property DownPayment() As Double
        Get
            Return _DownPayment
        End Get
        Set(ByVal Value As Double)
            _DownPayment = Value
        End Set
    End Property
    Public Property Payments() As Double
        Get
            Return _Payments
        End Get
        Set(ByVal Value As Double)
            _Payments = Value
        End Set
    End Property
    Public Property FinCharge() As Double
        Get
            Return _FinCharge
        End Get
        Set(ByVal Value As Double)
            _FinCharge = Value
        End Set
    End Property
    Public Property TotalPayment() As Double
        Get
            Return _TotalPayment
        End Get
        Set(ByVal Value As Double)
            _TotalPayment = Value
        End Set
    End Property
    Public Property NumPayments() As Long
        Get
            Return _NumPayments
        End Get
        Set(ByVal Value As Long)
            _NumPayments = Value
        End Set
    End Property
    Public Property FinanceDueDate() As Date
        Get
            Return _FinanceDueDate
        End Get
        Set(ByVal Value As Date)
            _FinanceDueDate = Value
        End Set
    End Property
    Public Property ReferenceKey_FK() As Long
        Get
            Return _ReferenceKey_FK
        End Get
        Set(ByVal Value As Long)
            _ReferenceKey_FK = Value
        End Set
    End Property
    Public Property RemitAmount() As Double
        Get
            Return _RemitAmount
        End Get
        Set(ByVal Value As Double)
            _RemitAmount = Value
        End Set
    End Property
    Public Property CollectAmount() As Double
        Get
            Return _CollectAmount
        End Get
        Set(ByVal Value As Double)
            _CollectAmount = Value
        End Set
    End Property
    Public Property DownFactor() As Double
        Get
            Return _DownFactor
        End Get
        Set(ByVal Value As Double)
            _DownFactor = Value
        End Set
    End Property
    Public Property TerrorActPremium() As Double
        Get
            Return _TerrorActPremium
        End Get
        Set(ByVal Value As Double)
            _TerrorActPremium = Value
        End Set
    End Property
    Public Property TerrorActGrossComm() As Double
        Get
            Return _TerrorActGrossComm
        End Get
        Set(ByVal Value As Double)
            _TerrorActGrossComm = Value
        End Set
    End Property
    Public Property TerrorActAgentComm() As Double
        Get
            Return _TerrorActAgentComm
        End Get
        Set(ByVal Value As Double)
            _TerrorActAgentComm = Value
        End Set
    End Property
    Public Property TerrorActMEP() As String
        Get
            Return _TerrorActMEP
        End Get
        Set(ByVal Value As String)
            _TerrorActMEP = Value
        End Set
    End Property
    Public Property TerrorActStatus() As String
        Get
            Return _TerrorActStatus
        End Get
        Set(ByVal Value As String)
            _TerrorActStatus = Value
        End Set
    End Property
    Public Property FlagOverrideCalc() As String
        Get
            Return _FlagOverrideCalc
        End Get
        Set(ByVal Value As String)
            _FlagOverrideCalc = Value
        End Set
    End Property
    Public Property TerrorTaxes() As Double
        Get
            Return _TerrorTaxes
        End Get
        Set(ByVal Value As Double)
            _TerrorTaxes = Value
        End Set
    End Property
    Public Property FlagFinanceWithTRIA() As String
        Get
            Return _FlagFinanceWithTRIA
        End Get
        Set(ByVal Value As String)
            _FlagFinanceWithTRIA = Value
        End Set
    End Property
    Public Property FlagMultiOption() As String
        Get
            Return _FlagMultiOption
        End Get
        Set(ByVal Value As String)
            _FlagMultiOption = Value
        End Set
    End Property
    Public Property FlagFeeCalc() As String
        Get
            Return _FlagFeeCalc
        End Get
        Set(ByVal Value As String)
            _FlagFeeCalc = Value
        End Set
    End Property
    Public Property ParticipantCo1ID() As String
        Get
            Return _ParticipantCo1ID
        End Get
        Set(ByVal Value As String)
            _ParticipantCo1ID = Value
        End Set
    End Property
    Public Property ParticipantCo2ID() As String
        Get
            Return _ParticipantCo2ID
        End Get
        Set(ByVal Value As String)
            _ParticipantCo2ID = Value
        End Set
    End Property
    Public Property ParticipantCo3ID() As String
        Get
            Return _ParticipantCo3ID
        End Get
        Set(ByVal Value As String)
            _ParticipantCo3ID = Value
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
    Public Property LOB_Coverage1() As String
        Get
            Return _LOB_Coverage1
        End Get
        Set(ByVal Value As String)
            _LOB_Coverage1 = Value
        End Set
    End Property
    Public Property LOB_Coverage2() As String
        Get
            Return _LOB_Coverage2
        End Get
        Set(ByVal Value As String)
            _LOB_Coverage2 = Value
        End Set
    End Property
    Public Property LOB_Coverage3() As String
        Get
            Return _LOB_Coverage3
        End Get
        Set(ByVal Value As String)
            _LOB_Coverage3 = Value
        End Set
    End Property
    Public Property LOB_Coverage4() As String
        Get
            Return _LOB_Coverage4
        End Get
        Set(ByVal Value As String)
            _LOB_Coverage4 = Value
        End Set
    End Property
    Public Property LOB_Coverage5() As String
        Get
            Return _LOB_Coverage5
        End Get
        Set(ByVal Value As String)
            _LOB_Coverage5 = Value
        End Set
    End Property
    Public Property LOB_Coverage6() As String
        Get
            Return _LOB_Coverage6
        End Get
        Set(ByVal Value As String)
            _LOB_Coverage6 = Value
        End Set
    End Property
    Public Property LOB_DeductType1() As String
        Get
            Return _LOB_DeductType1
        End Get
        Set(ByVal Value As String)
            _LOB_DeductType1 = Value
        End Set
    End Property
    Public Property LOB_DeductType2() As String
        Get
            Return _LOB_DeductType2
        End Get
        Set(ByVal Value As String)
            _LOB_DeductType2 = Value
        End Set
    End Property
    Public Property DeclinationReasonID() As String
        Get
            Return _DeclinationReasonID
        End Get
        Set(ByVal Value As String)
            _DeclinationReasonID = Value
        End Set
    End Property
    Public Property ERPOption() As String
        Get
            Return _ERPOption
        End Get
        Set(ByVal Value As String)
            _ERPOption = Value
        End Set
    End Property
    Public Property ERPDays() As Long
        Get
            Return _ERPDays
        End Get
        Set(ByVal Value As Long)
            _ERPDays = Value
        End Set
    End Property
    Public Property ERPPercent() As Double
        Get
            Return _ERPPercent
        End Get
        Set(ByVal Value As Double)
            _ERPPercent = Value
        End Set
    End Property
    Public Property ERPPremium() As Double
        Get
            Return _ERPPremium
        End Get
        Set(ByVal Value As Double)
            _ERPPremium = Value
        End Set
    End Property
    Public Property TaxwoTRIA1() As Double
        Get
            Return _TaxwoTRIA1
        End Get
        Set(ByVal Value As Double)
            _TaxwoTRIA1 = Value
        End Set
    End Property
    Public Property TaxwoTRIA2() As Double
        Get
            Return _TaxwoTRIA2
        End Get
        Set(ByVal Value As Double)
            _TaxwoTRIA2 = Value
        End Set
    End Property
    Public Property TaxwoTRIA3() As Double
        Get
            Return _TaxwoTRIA3
        End Get
        Set(ByVal Value As Double)
            _TaxwoTRIA3 = Value
        End Set
    End Property
    Public Property TaxwoTRIA4() As Double
        Get
            Return _TaxwoTRIA4
        End Get
        Set(ByVal Value As Double)
            _TaxwoTRIA4 = Value
        End Set
    End Property
    Public Property LOB_Prem4() As Double
        Get
            Return _LOB_Prem4
        End Get
        Set(ByVal Value As Double)
            _LOB_Prem4 = Value
        End Set
    End Property
    Public Property LOB_Coverage7() As String
        Get
            Return _LOB_Coverage7
        End Get
        Set(ByVal Value As String)
            _LOB_Coverage7 = Value
        End Set
    End Property
    Public Property LOB_Coverage8() As String
        Get
            Return _LOB_Coverage8
        End Get
        Set(ByVal Value As String)
            _LOB_Coverage8 = Value
        End Set
    End Property
    Public Property LOB_Limit7() As String
        Get
            Return _LOB_Limit7
        End Get
        Set(ByVal Value As String)
            _LOB_Limit7 = Value
        End Set
    End Property
    Public Property LOB_Limit8() As String
        Get
            Return _LOB_Limit8
        End Get
        Set(ByVal Value As String)
            _LOB_Limit8 = Value
        End Set
    End Property
    Public Property LOB_Limit7Value() As Double
        Get
            Return _LOB_Limit7Value
        End Get
        Set(ByVal Value As Double)
            _LOB_Limit7Value = Value
        End Set
    End Property
    Public Property LOB_Limit8Value() As Double
        Get
            Return _LOB_Limit8Value
        End Get
        Set(ByVal Value As Double)
            _LOB_Limit8Value = Value
        End Set
    End Property
    Public Property LOB_Prem5() As Double
        Get
            Return _LOB_Prem5
        End Get
        Set(ByVal Value As Double)
            _LOB_Prem5 = Value
        End Set
    End Property
    Public Property LOB_Prem6() As Double
        Get
            Return _LOB_Prem6
        End Get
        Set(ByVal Value As Double)
            _LOB_Prem6 = Value
        End Set
    End Property
    Public Property LOB_Prem7() As Double
        Get
            Return _LOB_Prem7
        End Get
        Set(ByVal Value As Double)
            _LOB_Prem7 = Value
        End Set
    End Property
    Public Property LOB_Prem8() As Double
        Get
            Return _LOB_Prem8
        End Get
        Set(ByVal Value As Double)
            _LOB_Prem8 = Value
        End Set
    End Property
    Public Property CoverageList() As String
        Get
            Return _CoverageList
        End Get
        Set(ByVal Value As String)
            _CoverageList = Value
        End Set
    End Property
    Public Property DocucorpFormList() As String
        Get
            Return _DocucorpFormList
        End Get
        Set(ByVal Value As String)
            _DocucorpFormList = Value
        End Set
    End Property
    Public Property TerrorActPremium_GL() As Double
        Get
            Return _TerrorActPremium_GL
        End Get
        Set(ByVal Value As Double)
            _TerrorActPremium_GL = Value
        End Set
    End Property
    Public Property FlagRecalcTaxes() As String
        Get
            Return _FlagRecalcTaxes
        End Get
        Set(ByVal Value As String)
            _FlagRecalcTaxes = Value
        End Set
    End Property
    Public Property DateMktResponseRecvd() As Date
        Get
            Return _DateMktResponseRecvd
        End Get
        Set(ByVal Value As Date)
            _DateMktResponseRecvd = Value
        End Set
    End Property
    Public Property CancelClause() As String
        Get
            Return _CancelClause
        End Get
        Set(ByVal Value As String)
            _CancelClause = Value
        End Set
    End Property
    Public Property PremiumProperty() As Double
        Get
            Return _PremiumProperty
        End Get
        Set(ByVal Value As Double)
            _PremiumProperty = Value
        End Set
    End Property
    Public Property PremiumLiability() As Double
        Get
            Return _PremiumLiability
        End Get
        Set(ByVal Value As Double)
            _PremiumLiability = Value
        End Set
    End Property
    Public Property PremiumOther() As Double
        Get
            Return _PremiumOther
        End Get
        Set(ByVal Value As Double)
            _PremiumOther = Value
        End Set
    End Property
    Public Property EndorsementKey_FK() As Long
        Get
            Return _EndorsementKey_FK
        End Get
        Set(ByVal Value As Long)
            _EndorsementKey_FK = Value
        End Set
    End Property

#End Region
#Region "Methods"
    Public Function Save(ByRef conn As sqlconnection) As String
        Dim comm As New SqlCommand("siu_p_insertversion", conn)
        Try
            With comm
                .CommandTimeout = 0
                .CommandType = CommandType.StoredProcedure
                .Parameters.AddWithValue("@QuoteID", _QuoteID)
                .Parameters.AddWithValue("@Version", _Version)
                .Parameters.AddWithValue("@VerOriginal", _VerOriginal)
                .Parameters.AddWithValue("@VersionID", _VersionID)
                .Parameters.AddWithValue("@AgentComm", _AgentComm)
                .Parameters.AddWithValue("@GrossComm", _GrossComm)
                .Parameters.AddWithValue("@MarketID", _MarketID)
                .Parameters.AddWithValue("@CompanyID", _CompanyID)
                .Parameters.AddWithValue("@ProductID", _ProductID)
                .Parameters.AddWithValue("@Premium", _Premium)
                .Parameters.AddWithValue("@Non_Premium", _Non_Premium)
                .Parameters.AddWithValue("@Taxed", _Taxed)
                .Parameters.AddWithValue("@ProposedEffective", _ProposedEffective)
                .Parameters.AddWithValue("@ProposedExpiration", _ProposedExpiration)
                .Parameters.AddWithValue("@Quoted", _Quoted)
                .Parameters.AddWithValue("@Financed", _Financed)
                .Parameters.AddWithValue("@BoundFlag", _BoundFlag)
                .Parameters.AddWithValue("@PolicyTerm", _PolicyTerm)
                .Parameters.AddWithValue("@FlagOverrideCalc", _FlagOverrideCalc)
                .Parameters.AddWithValue("@TRIA", _TerrorActPremium)
                .ExecuteNonQuery()
            End With
            Return ""
        Catch ex As Exception
            Return "Version.Save: " & ex.Message
            conn.Close()
        End Try
    End Function

    Public Function Load(ByRef conn As SqlConnection, ByVal pQuoteID As String) As String
        Try
            Dim comm As New SqlCommand("SIU_p_GetVersion", conn)
            Dim rs As SqlDataReader

            With comm
                .CommandTimeout = 0
                .CommandType = CommandType.StoredProcedure
                .Parameters.AddWithValue("@QuoteID", pQuoteID)
                rs = .ExecuteReader
            End With
            If rs.Read Then
                _QuoteID = rs("QuoteID")
                _VerOriginal = rs("VerOriginal")
                _Version = rs("Version")
                _LobID = rs("LobID")
                _LobSubID = rs("LobSubID")
                _CompanyID = rs("CompanyID")
                _ProductID = rs("ProductID")
                _Premium = rs("Premium")
                _Non_Premium = rs("Non_Premium")
                _Misc_Premium = rs("Misc_Premium")
                _NonTax_Premium = rs("NonTax_Premium")
                _Quoted = rs("Quoted")
                _Expires = rs("Expires")
                _Limits = rs("Limits")
                _Subject = rs("Subject")
                _Endorsement = rs("Endorsement")
                _Financed = rs("Financed")
                _Taxed = rs("Taxed")
                _MEP = rs("MEP")
                _Rate = rs("Rate")
                _GrossComm = rs("GrossComm")
                _AgentComm = rs("AgentComm")
                _Brokerage = rs("Brokerage")
                _Deductible = rs("Deductible")
                _CoInsure = rs("CoInsure")
                _StatusID = rs("StatusID")
                _ReasonID = rs("ReasonID")
                _SubmitDate = rs("SubmitDate")
                _SubmitPOC = rs("SubmitPOC")
                _MarketID = rs("MarketID")
                _Apportionment = rs("Apportionment")
                _Tax1 = rs("Tax1")
                _Tax2 = rs("Tax2")
                _Tax3 = rs("Tax3")
                _Tax4 = rs("Tax4")
                _FormID = rs("FormID")
                _RateInfo = rs("RateInfo")
                _Indicator = rs("Indicator")
                _PendingSuspenseID = rs("PendingSuspenseID")
                _CommPaid = rs("CommPaid")
                _AggregateLimits = rs("AggregateLimits")
                _DeductibleVal = rs("DeductibleVal")
                _BoundFlag = rs("BoundFlag")
                _DirectBillFlag = rs("DirectBillFlag")
                _ProposedEffective = rs("ProposedEffective")
                _ProposedExpiration = rs("ProposedExpiration")
                _ProposedTerm = rs("ProposedTerm")
                _Retroactive = rs("Retroactive")
                _RetroPeriod = rs("RetroPeriod")
                _UnderLyingCoverage = rs("UnderLyingCoverage")
                _MultiOption = rs("MultiOption")
                _MiscPrem1 = rs("MiscPrem1")
                _MiscPrem2 = rs("MiscPrem2")
                _MiscPrem3 = rs("MiscPrem3")
                _NonTax1 = rs("NonTax1")
                _NonTax2 = rs("NonTax2")
                _NonPrem1 = rs("NonPrem1")
                _NonPrem2 = rs("NonPrem2")
                _PaymentRecv = rs("PaymentRecv")
                _PremDownPayment = rs("PremDownPayment")
                _Valuation = rs("Valuation")
                _Retention = rs("Retention")
                _AIM_TransDate = rs("AIM_TransDate")
                _InvoiceCodes = rs("InvoiceCodes")
                _TaxDistrib = rs("TaxDistrib")
                _PremDistrib = rs("PremDistrib")
                _CAP_Limit = rs("CAP_Limit")
                _EPL_Limit = rs("EPL_Limit")
                _TakenOut_RatedTerm = rs("TakenOut_RatedTerm")
                _PolicyTerm = rs("PolicyTerm")
                _PolicyForm = rs("PolicyForm")
                _BillToCompanyID = rs("BillToCompanyID")
                _StatementKey_FK = rs("StatementKey_FK")
                _PaymentKey_FK = rs("PaymentKey_FK")
                _CommRecvd = rs("CommRecvd")
                _VersionID = rs("VersionID")
                _MarketContactKey_FK = rs("MarketContactKey_FK")
                _TIV = rs("TIV")
                _CompanyFees = rs("CompanyFees")
                _UnderLyingLimitsSum = rs("UnderLyingLimitsSum")
                _PunitiveDamage = rs("PunitiveDamage")
                _ThirdPartyLimits = rs("ThirdPartyLimits")
                _AnnualPremium = rs("AnnualPremium")
                _AnnualFees = rs("AnnualFees")
                _FlagCollectMuniTax = rs("FlagCollectMuniTax")
                _TrueExpire = rs("TrueExpire")
                _WrittenLimits = rs("WrittenLimits")
                _AttachPoint = rs("AttachPoint")
                _LineSlip = rs("LineSlip")
                _CoverageFormID = rs("CoverageFormID")
                _PositionID = rs("PositionID")
                _LobDistrib = rs("LobDistrib")
                _TotalTax = rs("TotalTax")
                _Total = rs("Total")
                _TotalAmount = rs("TotalAmount")
                _TaxesPaidBy = rs("TaxesPaidBy")
                _ResubmitDate = rs("ResubmitDate")
                _FeeSchedule = rs("FeeSchedule")
                _LobDistribSched = rs("LobDistribSched")
                _DeductType = rs("DeductType")
                _PremiumFinanceFee = rs("PremiumFinanceFee")
                _LOB_Field1 = rs("LOB_Field1")
                _LOB_Field2 = rs("LOB_Field2")
                _LOB_Field3 = rs("LOB_Field3")
                _LOB_Flag1 = rs("LOB_Flag1")
                _LOB_Prem1 = rs("LOB_Prem1")
                _LOB_Prem2 = rs("LOB_Prem2")
                _LOB_Prem3 = rs("LOB_Prem3")
                _LOB_Limit1 = rs("LOB_Limit1")
                _LOB_Limit2 = rs("LOB_Limit2")
                _LOB_Limit3 = rs("LOB_Limit3")
                _LOB_Limit4 = rs("LOB_Limit4")
                _LOB_Limit5 = rs("LOB_Limit5")
                _LOB_Limit6 = rs("LOB_Limit6")
                _LOB_Deduct1 = rs("LOB_Deduct1")
                _LOB_Deduct2 = rs("LOB_Deduct2")
                _LOB_Limit1Value = rs("LOB_Limit1Value")
                _LOB_Limit2Value = rs("LOB_Limit2Value")
                _LOB_Limit3Value = rs("LOB_Limit3Value")
                _LOB_Limit4Value = rs("LOB_Limit4Value")
                _LOB_Limit5Value = rs("LOB_Limit5Value")
                _LOB_Limit6Value = rs("LOB_Limit6Value")
                _LOB_Deduct1Value = rs("LOB_Deduct1Value")
                _LOB_Deduct2Value = rs("LOB_Deduct2Value")
                _TaxesPaidByID = rs("TaxesPaidByID")
                _FlagMultiStateTax = rs("FlagMultiStateTax")
                _MultiStateDistrib = rs("MultiStateDistrib")
                _AdmittedPremium = rs("AdmittedPremium")
                _RatedPremium = rs("RatedPremium")
                _APR = rs("APR")
                _AmountFinanced = rs("AmountFinanced")
                _DownPayment = rs("DownPayment")
                _Payments = rs("Payments")
                _FinCharge = rs("FinCharge")
                _TotalPayment = rs("TotalPayment")
                _NumPayments = rs("NumPayments")
                _FinanceDueDate = rs("FinanceDueDate")
                _ReferenceKey_FK = rs("ReferenceKey_FK")
                _RemitAmount = rs("RemitAmount")
                _CollectAmount = rs("CollectAmount")
                _DownFactor = rs("DownFactor")
                _TerrorActPremium = rs("TerrorActPremium")
                _TerrorActGrossComm = rs("TerrorActGrossComm")
                _TerrorActAgentComm = rs("TerrorActAgentComm")
                _TerrorActMEP = rs("TerrorActMEP")
                _TerrorActStatus = rs("TerrorActStatus")
                _FlagOverrideCalc = rs("FlagOverrideCalc")
                _TerrorTaxes = rs("TerrorTaxes")
                _FlagFinanceWithTRIA = rs("FlagFinanceWithTRIA")
                _FlagMultiOption = rs("FlagMultiOption")
                _FlagFeeCalc = rs("FlagFeeCalc")
                _ParticipantCo1ID = rs("ParticipantCo1ID")
                _ParticipantCo2ID = rs("ParticipantCo2ID")
                _ParticipantCo3ID = rs("ParticipantCo3ID")
                _UserDefinedStr1 = rs("UserDefinedStr1")
                _UserDefinedStr2 = rs("UserDefinedStr2")
                _UserDefinedStr3 = rs("UserDefinedStr3")
                _UserDefinedStr4 = rs("UserDefinedStr4")
                _UserDefinedDate1 = rs("UserDefinedDate1")
                _UserDefinedValue1 = rs("UserDefinedValue1")
                _LOB_Coverage1 = rs("LOB_Coverage1")
                _LOB_Coverage2 = rs("LOB_Coverage2")
                _LOB_Coverage3 = rs("LOB_Coverage3")
                _LOB_Coverage4 = rs("LOB_Coverage4")
                _LOB_Coverage5 = rs("LOB_Coverage5")
                _LOB_Coverage6 = rs("LOB_Coverage6")
                _LOB_DeductType1 = rs("LOB_DeductType1")
                _LOB_DeductType2 = rs("LOB_DeductType2")
                _DeclinationReasonID = rs("DeclinationReasonID")
                _ERPOption = rs("ERPOption")
                _ERPDays = rs("ERPDays")
                _ERPPercent = rs("ERPPercent")
                _ERPPremium = rs("ERPPremium")
                _TaxwoTRIA1 = rs("TaxwoTRIA1")
                _TaxwoTRIA2 = rs("TaxwoTRIA2")
                _TaxwoTRIA3 = rs("TaxwoTRIA3")
                _TaxwoTRIA4 = rs("TaxwoTRIA4")
                _LOB_Prem4 = rs("LOB_Prem4")
                _LOB_Coverage7 = rs("LOB_Coverage7")
                _LOB_Coverage8 = rs("LOB_Coverage8")
                _LOB_Limit7 = rs("LOB_Limit7")
                _LOB_Limit8 = rs("LOB_Limit8")
                _LOB_Limit7Value = rs("LOB_Limit7Value")
                _LOB_Limit8Value = rs("LOB_Limit8Value")
                _LOB_Prem5 = rs("LOB_Prem5")
                _LOB_Prem6 = rs("LOB_Prem6")
                _LOB_Prem7 = rs("LOB_Prem7")
                _LOB_Prem8 = rs("LOB_Prem8")
                _CoverageList = rs("CoverageList")
                _DocucorpFormList = rs("DocucorpFormList")
                _TerrorActPremium_GL = rs("TerrorActPremium_GL")
                _FlagRecalcTaxes = rs("FlagRecalcTaxes")
                _DateMktResponseRecvd = rs("DateMktResponseRecvd")
                _CancelClause = rs("CancelClause")
                _PremiumProperty = rs("PremiumProperty")
                _PremiumLiability = rs("PremiumLiability")
                _PremiumOther = rs("PremiumOther")
                _EndorsementKey_FK = rs("EndorsementKey_FK")
            End If
            rs.Close()
            Return ""
        Catch ex As Exception
            Return "Version: " & ex.Message
            conn.Close()
        End Try
    End Function

#End Region
End Class