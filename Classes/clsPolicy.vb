Imports System.Data.SqlClient
Imports DirectBillImports.Common
Public Class clsPolicy
#Region "Local Variables"
    Private _QuoteID As String
    Private _PolicyID As String
    Private _PolicyKey_PK As Long
    Private _Version As String
    Private _PolicyGrpID As String
    Private _Effective As Date
    Private _Expiration As Date
    Private _Inception As Date
    Private _Term As Integer
    Private _BillingType As String
    Private _NewAccountIndicator As String
    Private _RateDate As Date
    Private _MailToCode As String
    Private _PolicyStatusDate As Date
    Private _BillingAccountNumber As String
    Private _Bound As Date
    Private _BoundTime As Date
    Private _SetupDate As Date
    Private _MailoutDate As Date
    Private _FinanceCompanyID As String
    Private _Cancellation As Date
    Private _CancelEffective As Date
    Private _CancellationReason As String
    Private _NonRenewalCode As String
    Private _NonRenewBy As String
    Private _Reinstated As Date
    Private _Invoiced As String
    Private _Units As Long
    Private _UnitType As String
    Private _LocationZip As String
    Private _ClaimsPending As String
    Private _ClaimsMade As String
    Private _LossesPaid As Double
    Private _BillToCode As String
    Private _StatusID As String
    Private _Endorsement As String
    Private _AdditionalInsureds As String
    Private _InspectionOrdered As String
    Private _EC As String
    Private _Operations As String
    Private _SICID As String
    Private _Location As String
    Private _CancelTime_Old As Date
    Private _CancelRequestedBy As String
    Private _ReturnPrem As Double
    Private _ReturnRate As String
    Private _PolicySource As String
    Private _CompanyID As String
    Private _ProductID As String
    Private _ContractID As String
    Private _WrittenPremium As Double
    Private _Control_State As String
    Private _Financed As String
    Private _BinderType As String
    Private _RewriteCompanyID As String
    Private _RewritePolicyID As String
    Private _RewriteDate As Date
    Private _TypeID As String
    Private _BillTo As String
    Private _RenewalQuoteID As String
    Private _AuditID As String
    Private _AuditInception As Date
    Private _AuditType As String
    Private _AuditPremium As Double
    Private _AuditOutstanding As String
    Private _DeductType As String
    Private _PolicyForm As String
    Private _InstallID As String
    Private _SuspID As Long
    Private _InvoiceDate As Date
    Private _FormMakerName As String
    Private _TermPremiumAdj As Double
    Private _PolicyPrintDate As Date
    Private _EffectiveTime As String
    Private _ActivePolicyFlag As String
    Private _Limit1 As String
    Private _Coverage1 As String
    Private _Limit2 As String
    Private _Coverage2 As String
    Private _Limit3 As String
    Private _Coverage3 As String
    Private _Limit4 As String
    Private _Coverage4 As String
    Private _AIM_TransDate As Date
    Private _PolicyGroupKey_FK As Long
    Private _AccountKey_FK As Long
    Private _CancelTime As String
    Private _PolicyTerm As Integer
    Private _InspectionCo_FK As Long
    Private _LoanNumber As String
    Private _ReinsuranceCategory As String
    Private _PendingNOCKey_FK As Long
    Private _ContractName As String
    Private _ContractKey_FK As Long
    Private _DateInspectionOrdered As Date
    Private _DateNOC As Date
    Private _DateRenewalNotice As Date
    Private _AmountFinanced As Double
    Private _CountCancelled As Long
    Private _CountEndorsed As Long
    Private _CountRenewed As Long
    Private _CountClaims As Long
    Private _PremiumWritten As Double
    Private _PremiumBilled As Double
    Private _PremiumAdjustments As Double
    Private _PremiumTerm As Double
    Private _PremiumReturn As Double
    Private _DateRenewalLetter As Date
    Private _FlagFinancingFunded As String
    Private _FlagSubjectToAudit As String
    Private _FlagConfirmation As String
    Private _DateAuditReviewed As Date
    Private _AuditReceivedBy As String
    Private _DateAuditReceived As Date
    Private _AuditReviewedBy As String
    Private _InspectionOrderedBy As String
    Private _DateInspectionReceived As Date
    Private _InspectionReceivedBy As String
    Private _DateInspectionReviewed As Date
    Private _InspectionReviewedBy As String
    Private _FlagInspectionRequired As String
    Private _InspectionFile As String
    Private _DatePolicyReceived As Date
    Private _DateReceived As Date
    Private _FlagOverrideServiceUW As String
    Private _TRIAReceivedDate As Date
    Private _ERPEffective As Date
    Private _ERPExpiration As Date
    Private _DefaultBillingType As String
    Private _ProductionSplitKey_FK As Long
    Private _InvoiceKey_FK As Long
    Private _BasisPremiumTerm As Double
    Private _FlagLapseInCoverage As String
    Private _DateInspectionBilled As Date
    Private _InspectionInvoiceNumber As String
    Private _InspectionCost As Double
    Private _InspectionPhotosRecvd As Long
    Private _InspectionComments As String
    Private _InspectionInvoiceDate As Date
    Private _FlagPremFinAgentFunded As String
    Private _PremiumConvertedTerm As Double
    Private _DateInspectionReordered As Date
    Private _InspectionReorderedBy As String

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
    Public Property PolicyID() As String
        Get
            Return _PolicyID
        End Get
        Set(ByVal Value As String)
            _PolicyID = Value
        End Set
    End Property
    Public Property PolicyKey_PK() As Long
        Get
            Return _PolicyKey_PK
        End Get
        Set(ByVal Value As Long)
            _PolicyKey_PK = Value
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
    Public Property PolicyGrpID() As String
        Get
            Return _PolicyGrpID
        End Get
        Set(ByVal Value As String)
            _PolicyGrpID = Value
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
    Public Property Inception() As Date
        Get
            Return _Inception
        End Get
        Set(ByVal Value As Date)
            _Inception = Value
        End Set
    End Property
    Public Property Term() As Integer
        Get
            Return _Term
        End Get
        Set(ByVal Value As Integer)
            _Term = Value
        End Set
    End Property
    Public Property BillingType() As String
        Get
            Return _BillingType
        End Get
        Set(ByVal Value As String)
            _BillingType = Value
        End Set
    End Property
    Public Property NewAccountIndicator() As String
        Get
            Return _NewAccountIndicator
        End Get
        Set(ByVal Value As String)
            _NewAccountIndicator = Value
        End Set
    End Property
    Public Property RateDate() As Date
        Get
            Return _RateDate
        End Get
        Set(ByVal Value As Date)
            _RateDate = Value
        End Set
    End Property
    Public Property MailToCode() As String
        Get
            Return _MailToCode
        End Get
        Set(ByVal Value As String)
            _MailToCode = Value
        End Set
    End Property
    Public Property PolicyStatusDate() As Date
        Get
            Return _PolicyStatusDate
        End Get
        Set(ByVal Value As Date)
            _PolicyStatusDate = Value
        End Set
    End Property
    Public Property BillingAccountNumber() As String
        Get
            Return _BillingAccountNumber
        End Get
        Set(ByVal Value As String)
            _BillingAccountNumber = Value
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
    Public Property BoundTime() As Date
        Get
            Return _BoundTime
        End Get
        Set(ByVal Value As Date)
            _BoundTime = Value
        End Set
    End Property
    Public Property SetupDate() As Date
        Get
            Return _SetupDate
        End Get
        Set(ByVal Value As Date)
            _SetupDate = Value
        End Set
    End Property
    Public Property MailoutDate() As Date
        Get
            Return _MailoutDate
        End Get
        Set(ByVal Value As Date)
            _MailoutDate = Value
        End Set
    End Property
    Public Property FinanceCompanyID() As String
        Get
            Return IIf(Len(Trim(_FinanceCompanyID)) > 0, _FinanceCompanyID, "")
        End Get
        Set(ByVal Value As String)
            _FinanceCompanyID = Value
        End Set
    End Property
    Public Property Cancellation() As Date
        Get
            Return _Cancellation
        End Get
        Set(ByVal Value As Date)
            _Cancellation = Value
        End Set
    End Property
    Public Property CancelEffective() As Date
        Get
            Return _CancelEffective
        End Get
        Set(ByVal Value As Date)
            _CancelEffective = Value
        End Set
    End Property
    Public Property CancellationReason() As String
        Get
            Return _CancellationReason
        End Get
        Set(ByVal Value As String)
            _CancellationReason = Value
        End Set
    End Property
    Public Property NonRenewalCode() As String
        Get
            Return _NonRenewalCode
        End Get
        Set(ByVal Value As String)
            _NonRenewalCode = Value
        End Set
    End Property
    Public Property NonRenewBy() As String
        Get
            Return _NonRenewBy
        End Get
        Set(ByVal Value As String)
            _NonRenewBy = Value
        End Set
    End Property
    Public Property Reinstated() As Date
        Get
            Return _Reinstated
        End Get
        Set(ByVal Value As Date)
            _Reinstated = Value
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
    Public Property Units() As Long
        Get
            Return _Units
        End Get
        Set(ByVal Value As Long)
            _Units = Value
        End Set
    End Property
    Public Property UnitType() As String
        Get
            Return _UnitType
        End Get
        Set(ByVal Value As String)
            _UnitType = Value
        End Set
    End Property
    Public Property LocationZip() As String
        Get
            Return _LocationZip
        End Get
        Set(ByVal Value As String)
            _LocationZip = Value
        End Set
    End Property
    Public Property ClaimsPending() As String
        Get
            Return _ClaimsPending
        End Get
        Set(ByVal Value As String)
            _ClaimsPending = Value
        End Set
    End Property
    Public Property ClaimsMade() As String
        Get
            Return _ClaimsMade
        End Get
        Set(ByVal Value As String)
            _ClaimsMade = Value
        End Set
    End Property
    Public Property LossesPaid() As Double
        Get
            Return _LossesPaid
        End Get
        Set(ByVal Value As Double)
            _LossesPaid = Value
        End Set
    End Property
    Public Property BillToCode() As String
        Get
            Return _BillToCode
        End Get
        Set(ByVal Value As String)
            _BillToCode = Value
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
    Public Property Endorsement() As String
        Get
            Return _Endorsement
        End Get
        Set(ByVal Value As String)
            _Endorsement = Value
        End Set
    End Property
    Public Property AdditionalInsureds() As String
        Get
            Return _AdditionalInsureds
        End Get
        Set(ByVal Value As String)
            _AdditionalInsureds = Value
        End Set
    End Property
    Public Property InspectionOrdered() As String
        Get
            Return _InspectionOrdered
        End Get
        Set(ByVal Value As String)
            _InspectionOrdered = Value
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
    Public Property Operations() As String
        Get
            Return _Operations
        End Get
        Set(ByVal Value As String)
            _Operations = Value
        End Set
    End Property
    Public Property SICID() As String
        Get
            Return _SICID
        End Get
        Set(ByVal Value As String)
            _SICID = Value
        End Set
    End Property
    Public Property Location() As String
        Get
            Return _Location
        End Get
        Set(ByVal Value As String)
            _Location = Value
        End Set
    End Property
    Public Property CancelTime_Old() As Date
        Get
            Return _CancelTime_Old
        End Get
        Set(ByVal Value As Date)
            _CancelTime_Old = Value
        End Set
    End Property
    Public Property CancelRequestedBy() As String
        Get
            Return _CancelRequestedBy
        End Get
        Set(ByVal Value As String)
            _CancelRequestedBy = Value
        End Set
    End Property
    Public Property ReturnPrem() As Double
        Get
            Return _ReturnPrem
        End Get
        Set(ByVal Value As Double)
            _ReturnPrem = Value
        End Set
    End Property
    Public Property ReturnRate() As String
        Get
            Return _ReturnRate
        End Get
        Set(ByVal Value As String)
            _ReturnRate = Value
        End Set
    End Property
    Public Property PolicySource() As String
        Get
            Return _PolicySource
        End Get
        Set(ByVal Value As String)
            _PolicySource = Value
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
    Public Property ContractID() As String
        Get
            Return _ContractID
        End Get
        Set(ByVal Value As String)
            _ContractID = Value
        End Set
    End Property
    Public Property WrittenPremium() As Double
        Get
            Return _WrittenPremium
        End Get
        Set(ByVal Value As Double)
            _WrittenPremium = Value
        End Set
    End Property
    Public Property Control_State() As String
        Get
            Return _Control_State
        End Get
        Set(ByVal Value As String)
            _Control_State = Value
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
    Public Property BinderType() As String
        Get
            Return _BinderType
        End Get
        Set(ByVal Value As String)
            _BinderType = Value
        End Set
    End Property
    Public Property RewriteCompanyID() As String
        Get
            Return _RewriteCompanyID
        End Get
        Set(ByVal Value As String)
            _RewriteCompanyID = Value
        End Set
    End Property
    Public Property RewritePolicyID() As String
        Get
            Return _RewritePolicyID
        End Get
        Set(ByVal Value As String)
            _RewritePolicyID = Value
        End Set
    End Property
    Public Property RewriteDate() As Date
        Get
            Return _RewriteDate
        End Get
        Set(ByVal Value As Date)
            _RewriteDate = Value
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
    Public Property BillTo() As String
        Get
            Return _BillTo
        End Get
        Set(ByVal Value As String)
            _BillTo = Value
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
    Public Property AuditID() As String
        Get
            Return _AuditID
        End Get
        Set(ByVal Value As String)
            _AuditID = Value
        End Set
    End Property
    Public Property AuditInception() As Date
        Get
            Return _AuditInception
        End Get
        Set(ByVal Value As Date)
            _AuditInception = Value
        End Set
    End Property
    Public Property AuditType() As String
        Get
            Return _AuditType
        End Get
        Set(ByVal Value As String)
            _AuditType = Value
        End Set
    End Property
    Public Property AuditPremium() As Double
        Get
            Return _AuditPremium
        End Get
        Set(ByVal Value As Double)
            _AuditPremium = Value
        End Set
    End Property
    Public Property AuditOutstanding() As String
        Get
            Return _AuditOutstanding
        End Get
        Set(ByVal Value As String)
            _AuditOutstanding = Value
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
    Public Property PolicyForm() As String
        Get
            Return _PolicyForm
        End Get
        Set(ByVal Value As String)
            _PolicyForm = Value
        End Set
    End Property
    Public Property InstallID() As String
        Get
            Return _InstallID
        End Get
        Set(ByVal Value As String)
            _InstallID = Value
        End Set
    End Property
    Public Property SuspID() As Long
        Get
            Return _SuspID
        End Get
        Set(ByVal Value As Long)
            _SuspID = Value
        End Set
    End Property
    Public Property InvoiceDate() As Date
        Get
            Return _InvoiceDate
        End Get
        Set(ByVal Value As Date)
            _InvoiceDate = Value
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
    Public Property TermPremiumAdj() As Double
        Get
            Return _TermPremiumAdj
        End Get
        Set(ByVal Value As Double)
            _TermPremiumAdj = Value
        End Set
    End Property
    Public Property PolicyPrintDate() As Date
        Get
            Return _PolicyPrintDate
        End Get
        Set(ByVal Value As Date)
            _PolicyPrintDate = Value
        End Set
    End Property
    Public Property EffectiveTime() As String
        Get
            Return _EffectiveTime
        End Get
        Set(ByVal Value As String)
            _EffectiveTime = Value
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
    Public Property Limit1() As String
        Get
            Return _Limit1
        End Get
        Set(ByVal Value As String)
            _Limit1 = Value
        End Set
    End Property
    Public Property Coverage1() As String
        Get
            Return _Coverage1
        End Get
        Set(ByVal Value As String)
            _Coverage1 = Value
        End Set
    End Property
    Public Property Limit2() As String
        Get
            Return _Limit2
        End Get
        Set(ByVal Value As String)
            _Limit2 = Value
        End Set
    End Property
    Public Property Coverage2() As String
        Get
            Return _Coverage2
        End Get
        Set(ByVal Value As String)
            _Coverage2 = Value
        End Set
    End Property
    Public Property Limit3() As String
        Get
            Return _Limit3
        End Get
        Set(ByVal Value As String)
            _Limit3 = Value
        End Set
    End Property
    Public Property Coverage3() As String
        Get
            Return _Coverage3
        End Get
        Set(ByVal Value As String)
            _Coverage3 = Value
        End Set
    End Property
    Public Property Limit4() As String
        Get
            Return _Limit4
        End Get
        Set(ByVal Value As String)
            _Limit4 = Value
        End Set
    End Property
    Public Property Coverage4() As String
        Get
            Return _Coverage4
        End Get
        Set(ByVal Value As String)
            _Coverage4 = Value
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
    Public Property PolicyGroupKey_FK() As Long
        Get
            Return _PolicyGroupKey_FK
        End Get
        Set(ByVal Value As Long)
            _PolicyGroupKey_FK = Value
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
    Public Property CancelTime() As String
        Get
            Return _CancelTime
        End Get
        Set(ByVal Value As String)
            _CancelTime = Value
        End Set
    End Property
    Public Property PolicyTerm() As Integer
        Get
            Return _PolicyTerm
        End Get
        Set(ByVal Value As Integer)
            _PolicyTerm = Value
        End Set
    End Property
    Public Property InspectionCo_FK() As Long
        Get
            Return _InspectionCo_FK
        End Get
        Set(ByVal Value As Long)
            _InspectionCo_FK = Value
        End Set
    End Property
    Public Property LoanNumber() As String
        Get
            Return _LoanNumber
        End Get
        Set(ByVal Value As String)
            _LoanNumber = Value
        End Set
    End Property
    Public Property ReinsuranceCategory() As String
        Get
            Return _ReinsuranceCategory
        End Get
        Set(ByVal Value As String)
            _ReinsuranceCategory = Value
        End Set
    End Property
    Public Property PendingNOCKey_FK() As Long
        Get
            Return _PendingNOCKey_FK
        End Get
        Set(ByVal Value As Long)
            _PendingNOCKey_FK = Value
        End Set
    End Property
    Public Property ContractName() As String
        Get
            Return _ContractName
        End Get
        Set(ByVal Value As String)
            _ContractName = Value
        End Set
    End Property
    Public Property ContractKey_FK() As Long
        Get
            Return _ContractKey_FK
        End Get
        Set(ByVal Value As Long)
            _ContractKey_FK = Value
        End Set
    End Property
    Public Property DateInspectionOrdered() As Date
        Get
            Return _DateInspectionOrdered
        End Get
        Set(ByVal Value As Date)
            _DateInspectionOrdered = Value
        End Set
    End Property
    Public Property DateNOC() As Date
        Get
            Return _DateNOC
        End Get
        Set(ByVal Value As Date)
            _DateNOC = Value
        End Set
    End Property
    Public Property DateRenewalNotice() As Date
        Get
            Return _DateRenewalNotice
        End Get
        Set(ByVal Value As Date)
            _DateRenewalNotice = Value
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
    Public Property CountCancelled() As Long
        Get
            Return _CountCancelled
        End Get
        Set(ByVal Value As Long)
            _CountCancelled = Value
        End Set
    End Property
    Public Property CountEndorsed() As Long
        Get
            Return _CountEndorsed
        End Get
        Set(ByVal Value As Long)
            _CountEndorsed = Value
        End Set
    End Property
    Public Property CountRenewed() As Long
        Get
            Return _CountRenewed
        End Get
        Set(ByVal Value As Long)
            _CountRenewed = Value
        End Set
    End Property
    Public Property CountClaims() As Long
        Get
            Return _CountClaims
        End Get
        Set(ByVal Value As Long)
            _CountClaims = Value
        End Set
    End Property
    Public Property PremiumWritten() As Double
        Get
            Return _PremiumWritten
        End Get
        Set(ByVal Value As Double)
            _PremiumWritten = Value
        End Set
    End Property
    Public Property PremiumBilled() As Double
        Get
            Return _PremiumBilled
        End Get
        Set(ByVal Value As Double)
            _PremiumBilled = Value
        End Set
    End Property
    Public Property PremiumAdjustments() As Double
        Get
            Return _PremiumAdjustments
        End Get
        Set(ByVal Value As Double)
            _PremiumAdjustments = Value
        End Set
    End Property
    Public Property PremiumTerm() As Double
        Get
            Return _PremiumTerm
        End Get
        Set(ByVal Value As Double)
            _PremiumTerm = Value
        End Set
    End Property
    Public Property PremiumReturn() As Double
        Get
            Return _PremiumReturn
        End Get
        Set(ByVal Value As Double)
            _PremiumReturn = Value
        End Set
    End Property
    Public Property DateRenewalLetter() As Date
        Get
            Return _DateRenewalLetter
        End Get
        Set(ByVal Value As Date)
            _DateRenewalLetter = Value
        End Set
    End Property
    Public Property FlagFinancingFunded() As String
        Get
            Return _FlagFinancingFunded
        End Get
        Set(ByVal Value As String)
            _FlagFinancingFunded = Value
        End Set
    End Property
    Public Property FlagSubjectToAudit() As String
        Get
            Return _FlagSubjectToAudit
        End Get
        Set(ByVal Value As String)
            _FlagSubjectToAudit = Value
        End Set
    End Property
    Public Property FlagConfirmation() As String
        Get
            Return _FlagConfirmation
        End Get
        Set(ByVal Value As String)
            _FlagConfirmation = Value
        End Set
    End Property
    Public Property DateAuditReviewed() As Date
        Get
            Return _DateAuditReviewed
        End Get
        Set(ByVal Value As Date)
            _DateAuditReviewed = Value
        End Set
    End Property
    Public Property AuditReceivedBy() As String
        Get
            Return _AuditReceivedBy
        End Get
        Set(ByVal Value As String)
            _AuditReceivedBy = Value
        End Set
    End Property
    Public Property DateAuditReceived() As Date
        Get
            Return _DateAuditReceived
        End Get
        Set(ByVal Value As Date)
            _DateAuditReceived = Value
        End Set
    End Property
    Public Property AuditReviewedBy() As String
        Get
            Return _AuditReviewedBy
        End Get
        Set(ByVal Value As String)
            _AuditReviewedBy = Value
        End Set
    End Property
    Public Property InspectionOrderedBy() As String
        Get
            Return _InspectionOrderedBy
        End Get
        Set(ByVal Value As String)
            _InspectionOrderedBy = Value
        End Set
    End Property
    Public Property DateInspectionReceived() As Date
        Get
            Return _DateInspectionReceived
        End Get
        Set(ByVal Value As Date)
            _DateInspectionReceived = Value
        End Set
    End Property
    Public Property InspectionReceivedBy() As String
        Get
            Return _InspectionReceivedBy
        End Get
        Set(ByVal Value As String)
            _InspectionReceivedBy = Value
        End Set
    End Property
    Public Property DateInspectionReviewed() As Date
        Get
            Return _DateInspectionReviewed
        End Get
        Set(ByVal Value As Date)
            _DateInspectionReviewed = Value
        End Set
    End Property
    Public Property InspectionReviewedBy() As String
        Get
            Return _InspectionReviewedBy
        End Get
        Set(ByVal Value As String)
            _InspectionReviewedBy = Value
        End Set
    End Property
    Public Property FlagInspectionRequired() As String
        Get
            Return _FlagInspectionRequired
        End Get
        Set(ByVal Value As String)
            _FlagInspectionRequired = Value
        End Set
    End Property
    Public Property InspectionFile() As String
        Get
            Return _InspectionFile
        End Get
        Set(ByVal Value As String)
            _InspectionFile = Value
        End Set
    End Property
    Public Property DatePolicyReceived() As Date
        Get
            Return _DatePolicyReceived
        End Get
        Set(ByVal Value As Date)
            _DatePolicyReceived = Value
        End Set
    End Property
    Public Property DateReceived() As Date
        Get
            Return _DateReceived
        End Get
        Set(ByVal Value As Date)
            _DateReceived = Value
        End Set
    End Property
    Public Property FlagOverrideServiceUW() As String
        Get
            Return _FlagOverrideServiceUW
        End Get
        Set(ByVal Value As String)
            _FlagOverrideServiceUW = Value
        End Set
    End Property
    Public Property TRIAReceivedDate() As Date
        Get
            Return _TRIAReceivedDate
        End Get
        Set(ByVal Value As Date)
            _TRIAReceivedDate = Value
        End Set
    End Property
    Public Property ERPEffective() As Date
        Get
            Return _ERPEffective
        End Get
        Set(ByVal Value As Date)
            _ERPEffective = Value
        End Set
    End Property
    Public Property ERPExpiration() As Date
        Get
            Return _ERPExpiration
        End Get
        Set(ByVal Value As Date)
            _ERPExpiration = Value
        End Set
    End Property
    Public Property DefaultBillingType() As String
        Get
            Return _DefaultBillingType
        End Get
        Set(ByVal Value As String)
            _DefaultBillingType = Value
        End Set
    End Property
    Public Property ProductionSplitKey_FK() As Long
        Get
            Return _ProductionSplitKey_FK
        End Get
        Set(ByVal Value As Long)
            _ProductionSplitKey_FK = Value
        End Set
    End Property
    Public Property InvoiceKey_FK() As Long
        Get
            Return _InvoiceKey_FK
        End Get
        Set(ByVal Value As Long)
            _InvoiceKey_FK = Value
        End Set
    End Property
    Public Property BasisPremiumTerm() As Double
        Get
            Return _BasisPremiumTerm
        End Get
        Set(ByVal Value As Double)
            _BasisPremiumTerm = Value
        End Set
    End Property
    Public Property FlagLapseInCoverage() As String
        Get
            Return _FlagLapseInCoverage
        End Get
        Set(ByVal Value As String)
            _FlagLapseInCoverage = Value
        End Set
    End Property
    Public Property DateInspectionBilled() As Date
        Get
            Return _DateInspectionBilled
        End Get
        Set(ByVal Value As Date)
            _DateInspectionBilled = Value
        End Set
    End Property
    Public Property InspectionInvoiceNumber() As String
        Get
            Return _InspectionInvoiceNumber
        End Get
        Set(ByVal Value As String)
            _InspectionInvoiceNumber = Value
        End Set
    End Property
    Public Property InspectionCost() As Double
        Get
            Return _InspectionCost
        End Get
        Set(ByVal Value As Double)
            _InspectionCost = Value
        End Set
    End Property
    Public Property InspectionPhotosRecvd() As Long
        Get
            Return _InspectionPhotosRecvd
        End Get
        Set(ByVal Value As Long)
            _InspectionPhotosRecvd = Value
        End Set
    End Property
    Public Property InspectionComments() As String
        Get
            Return _InspectionComments
        End Get
        Set(ByVal Value As String)
            _InspectionComments = Value
        End Set
    End Property
    Public Property InspectionInvoiceDate() As Date
        Get
            Return _InspectionInvoiceDate
        End Get
        Set(ByVal Value As Date)
            _InspectionInvoiceDate = Value
        End Set
    End Property
    Public Property FlagPremFinAgentFunded() As String
        Get
            Return _FlagPremFinAgentFunded
        End Get
        Set(ByVal Value As String)
            _FlagPremFinAgentFunded = Value
        End Set
    End Property
    Public Property PremiumConvertedTerm() As Double
        Get
            Return _PremiumConvertedTerm
        End Get
        Set(ByVal Value As Double)
            _PremiumConvertedTerm = Value
        End Set
    End Property
    Public Property DateInspectionReordered() As Date
        Get
            Return _DateInspectionReordered
        End Get
        Set(ByVal Value As Date)
            _DateInspectionReordered = Value
        End Set
    End Property
    Public Property InspectionReorderedBy() As String
        Get
            Return _InspectionReorderedBy
        End Get
        Set(ByVal Value As String)
            _InspectionReorderedBy = Value
        End Set
    End Property

#End Region
#Region "Methods"
    Public Function Save(ByRef conn As SqlConnection) As String
        Dim comm As New SqlCommand("siu_p_insertpolicy", conn)
        Try
            With comm
                .CommandTimeout = 0
                .CommandType = CommandType.StoredProcedure
                .Parameters.AddWithValue("@QuoteID", _QuoteID)
                .Parameters.AddWithValue("@PolicyID", _PolicyID)
                .Parameters.AddWithValue("@Effective", _Effective)
                .Parameters.AddWithValue("@Expiration", _Expiration)
                .Parameters.AddWithValue("@Endorsement", _Endorsement)
                .Parameters.AddWithValue("@ActivePolicyFlag", _ActivePolicyFlag)
                .Parameters.AddWithValue("@PolicyKey_PK", _PolicyKey_PK)

                .Parameters.AddWithValue("@Version", _Version)
                .Parameters.AddWithValue("@Inception", _Inception)
                .Parameters.AddWithValue("@Term", _Term)
                .Parameters.AddWithValue("@Bound", _Bound)
                .Parameters.AddWithValue("@Invoiced", _Invoiced)
                .Parameters.AddWithValue("@CompanyID", _CompanyID)
                .Parameters.AddWithValue("@ProductID", _ProductID)
                .Parameters.AddWithValue("@InvoiceDate", _InvoiceDate)
                .Parameters.AddWithValue("@PremiumWritten", _PremiumWritten)
                .Parameters.AddWithValue("@PremiumTerm", _PremiumTerm)
                .Parameters.AddWithValue("@FlagInspectionRequired", _FlagInspectionRequired)
                .Parameters.AddWithValue("@FlagOverrideServiceUW", _FlagOverrideServiceUW)
                .Parameters.AddWithValue("@DefaultBillingType", _DefaultBillingType)
                .ExecuteNonQuery()
            End With
            Return ""
        Catch ex As Exception
            Return "Policy.Save: " & ex.Message
            conn.Close()
        End Try
    End Function
    Public Function Load(ByRef conn As SqlConnection, ByVal pPolicyKey_FK As Integer) As String
        Try
            Dim comm As New SqlCommand("SIU_p_GetPolicy", conn)
            Dim rs As SqlDataReader

            With comm
                .CommandTimeout = 0
                .CommandType = CommandType.StoredProcedure
                .Parameters.AddWithValue("@PolicyKey_PK", pPolicyKey_FK)
                rs = .ExecuteReader
            End With
            If rs.Read Then
                _QuoteID = rs("QuoteID")
                _PolicyID = rs("PolicyID")
                _PolicyKey_PK = rs("PolicyKey_PK")
                _Version = rs("Version")
                _PolicyGrpID = rs("PolicyGrpID")
                _Effective = rs("Effective")
                _Expiration = rs("Expiration")
                _Inception = rs("Inception")
                _Term = rs("Term")
                _BillingType = rs("BillingType")
                _NewAccountIndicator = rs("NewAccountIndicator")
                _RateDate = rs("RateDate")
                _MailToCode = rs("MailToCode")
                _PolicyStatusDate = rs("PolicyStatusDate")
                _BillingAccountNumber = rs("BillingAccountNumber")
                _Bound = rs("Bound")
                _BoundTime = rs("BoundTime")
                _SetupDate = rs("SetupDate")
                _MailoutDate = rs("MailoutDate")
                _FinanceCompanyID = rs("FinanceCompanyID")
                _Cancellation = rs("Cancellation")
                _CancelEffective = rs("CancelEffective")
                _CancellationReason = rs("CancellationReason")
                _NonRenewalCode = rs("NonRenewalCode")
                _NonRenewBy = rs("NonRenewBy")
                _Reinstated = rs("Reinstated")
                _Invoiced = rs("Invoiced")
                _Units = rs("Units")
                _UnitType = rs("UnitType")
                _LocationZip = rs("LocationZip")
                _ClaimsPending = rs("ClaimsPending")
                _ClaimsMade = rs("ClaimsMade")
                _LossesPaid = rs("LossesPaid")
                _BillToCode = rs("BillToCode")
                _StatusID = rs("StatusID")
                _Endorsement = rs("Endorsement")
                _AdditionalInsureds = rs("AdditionalInsureds")
                _InspectionOrdered = rs("InspectionOrdered")
                _EC = rs("EC")
                _Operations = rs("Operations")
                _SICID = rs("SICID")
                _Location = rs("Location")
                _CancelTime_Old = rs("CancelTime_Old")
                _CancelRequestedBy = rs("CancelRequestedBy")
                _ReturnPrem = rs("ReturnPrem")
                _ReturnRate = rs("ReturnRate")
                _PolicySource = rs("PolicySource")
                _CompanyID = rs("CompanyID")
                _ProductID = rs("ProductID")
                _ContractID = rs("ContractID")
                _WrittenPremium = rs("WrittenPremium")
                _Control_State = rs("Control_State")
                _Financed = rs("Financed")
                _BinderType = rs("BinderType")
                _RewriteCompanyID = rs("RewriteCompanyID")
                _RewritePolicyID = rs("RewritePolicyID")
                _RewriteDate = rs("RewriteDate")
                _TypeID = rs("TypeID")
                _BillTo = rs("BillTo")
                _RenewalQuoteID = rs("RenewalQuoteID")
                _AuditID = rs("AuditID")
                _AuditInception = rs("AuditInception")
                _AuditType = rs("AuditType")
                _AuditPremium = rs("AuditPremium")
                _AuditOutstanding = rs("AuditOutstanding")
                _DeductType = rs("DeductType")
                _PolicyForm = rs("PolicyForm")
                _InstallID = rs("InstallID")
                _SuspID = rs("SuspID")
                _InvoiceDate = rs("InvoiceDate")
                _FormMakerName = rs("FormMakerName")
                _TermPremiumAdj = rs("TermPremiumAdj")
                _PolicyPrintDate = rs("PolicyPrintDate")
                _EffectiveTime = rs("EffectiveTime")
                _ActivePolicyFlag = rs("ActivePolicyFlag")
                _Limit1 = rs("Limit1")
                _Coverage1 = rs("Coverage1")
                _Limit2 = rs("Limit2")
                _Coverage2 = rs("Coverage2")
                _Limit3 = rs("Limit3")
                _Coverage3 = rs("Coverage3")
                _Limit4 = rs("Limit4")
                _Coverage4 = rs("Coverage4")
                _AIM_TransDate = rs("AIM_TransDate")
                _PolicyGroupKey_FK = rs("PolicyGroupKey_FK")
                _AccountKey_FK = rs("AccountKey_FK")
                _CancelTime = rs("CancelTime")
                _PolicyTerm = rs("PolicyTerm")
                _InspectionCo_FK = rs("InspectionCo_FK")
                _LoanNumber = rs("LoanNumber")
                _ReinsuranceCategory = rs("ReinsuranceCategory")
                _PendingNOCKey_FK = rs("PendingNOCKey_FK")
                _ContractName = rs("ContractName")
                _ContractKey_FK = rs("ContractKey_FK")
                _DateInspectionOrdered = rs("DateInspectionOrdered")
                _DateNOC = rs("DateNOC")
                _DateRenewalNotice = rs("DateRenewalNotice")
                _AmountFinanced = rs("AmountFinanced")
                _CountCancelled = rs("CountCancelled")
                _CountEndorsed = rs("CountEndorsed")
                _CountRenewed = rs("CountRenewed")
                _CountClaims = rs("CountClaims")
                _PremiumWritten = rs("PremiumWritten")
                _PremiumBilled = rs("PremiumBilled")
                _PremiumAdjustments = rs("PremiumAdjustments")
                _PremiumTerm = rs("PremiumTerm")
                _PremiumReturn = rs("PremiumReturn")
                _DateRenewalLetter = rs("DateRenewalLetter")
                _FlagFinancingFunded = rs("FlagFinancingFunded")
                _FlagSubjectToAudit = rs("FlagSubjectToAudit")
                _FlagConfirmation = rs("FlagConfirmation")
                _DateAuditReviewed = rs("DateAuditReviewed")
                _AuditReceivedBy = rs("AuditReceivedBy")
                _DateAuditReceived = rs("DateAuditReceived")
                _AuditReviewedBy = rs("AuditReviewedBy")
                _InspectionOrderedBy = rs("InspectionOrderedBy")
                _DateInspectionReceived = rs("DateInspectionReceived")
                _InspectionReceivedBy = rs("InspectionReceivedBy")
                _DateInspectionReviewed = rs("DateInspectionReviewed")
                _InspectionReviewedBy = rs("InspectionReviewedBy")
                _FlagInspectionRequired = rs("FlagInspectionRequired")
                _InspectionFile = rs("InspectionFile")
                _DatePolicyReceived = rs("DatePolicyReceived")
                _DateReceived = rs("DateReceived")
                _FlagOverrideServiceUW = rs("FlagOverrideServiceUW")
                _TRIAReceivedDate = rs("TRIAReceivedDate")
                _ERPEffective = rs("ERPEffective")
                _ERPExpiration = rs("ERPExpiration")
                _DefaultBillingType = rs("DefaultBillingType")
                _ProductionSplitKey_FK = rs("ProductionSplitKey_FK")
                _InvoiceKey_FK = rs("InvoiceKey_FK")
                _BasisPremiumTerm = rs("BasisPremiumTerm")
                _FlagLapseInCoverage = rs("FlagLapseInCoverage")
                _DateInspectionBilled = rs("DateInspectionBilled")
                _InspectionInvoiceNumber = rs("InspectionInvoiceNumber")
                _InspectionCost = rs("InspectionCost")
                _InspectionPhotosRecvd = rs("InspectionPhotosRecvd")
                _InspectionComments = rs("InspectionComments")
                _InspectionInvoiceDate = rs("InspectionInvoiceDate")
                _FlagPremFinAgentFunded = rs("FlagPremFinAgentFunded")
                _PremiumConvertedTerm = rs("PremiumConvertedTerm")
                _DateInspectionReordered = rs("DateInspectionReordered")
                _InspectionReorderedBy = rs("InspectionReorderedBy")
            End If
            rs.Close()
            Return ""
        Catch ex As Exception
            Return "Policy: " & ex.Message
            conn.Close()
        End Try
    End Function

#End Region

End Class