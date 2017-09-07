Imports System.data.sqlclient
Imports DirectBillImports.Common
Public Class clsRisk_CommProperty
#Region "Local Variables"

    Private _ReferenceKey_FK As Long
    Private _RiskDetailKey_PK As Long
    Private _Description As String
    Private _Address1 As String
    Private _Address2 As String
    Private _City As String
    Private _State As String
    Private _ZipCode As String
    Private _ZipPlus As String
    Private _County As String
    Private _Units As Long
    Private _YearBuilt As Integer
    Private _BuildingValue As Double
    Private _DateAdded As Date
    Private _DateDropped As Date
    Private _BuildingNumber As Long
    Private _LocationNumber As Long
    Private _Occupancy As String
    Private _ConstructionID As String
    Private _YearManagedSince As Integer
    Private _RentableSqFootage As Long
    Private _BuildingFloors As Long
    Private _ContentsValue As Double
    Private _EDPValue As Double
    Private _RentalValue As Double
    Private _OtherValue As Double
    Private _TotalLimits As Double
    Private _ActiveFlag As String
    Private _InsuredKey_FK As Long
    Private _ContentsPremium As Double
    Private _EDPPremium As Double
    Private _RentalPremium As Double
    Private _OtherPremium As Double
    Private _FloodQuakeLimits As Double
    Private _FloodQuakePremium As Double
    Private _EmployeeDishonestPremium As Double
    Private _LiabilityTotal As Double
    Private _BuildingPremium As Double
    Private _GLRate As Double
    Private _GLBasePremium As Double
    Private _GLOtherPremium As Double
    Private _GLTotalPremium As Double
    Private _PropertyRate As Double
    Private _PropertyBasePremium As Double
    Private _PropertyOtherPremium As Double
    Private _PropertyTotalPremium As Double
    Private _MgtContactKey_FK As Long
    Private _MgtContact As String
    Private _MgtPhone As String
    Private _InspectionContact As String
    Private _InspectionPhone As String
    Private _PercentOccupied As Double
    Private _Acreage As Long
    Private _NumberElevators As Long
    Private _NumberParkingSpaces As Long
    Private _PercentSprinklered As Double
    Private _NumberCommTenants As Long
    Private _RetailSqFoot As Long
    Private _CommSqFoot As Long
    Private _ResidentalSqFoot As Long
    Private _Territory As String
    Private _TotalPremium As Double
    Private _NumberOfBuildings As Long
    Private _YearRenovated As Integer
    Private _ParkingType As String
    Private _NOTES As String
    Private _InspectionFax As String
    Private _MgtFax As String
    Private _FileReviewed As Date
    Private _ProtectionClass As String
    Private _BuildingUpdatesWire As Integer
    Private _BuildingUpdatesRoofing As Integer
    Private _BuildingUpdatesPlumbing As Integer
    Private _BuildingUpdatesHeating As Integer
    Private _FlagSmokeDetector As String
    Private _FlagLocalBurglar As String
    Private _FlagLocalFire As String
    Private _FlagCentralBurglar As String
    Private _FlagCentralFire As String
    Private _FlagSprinkled As String
    Private _FlagDeadBolt As String
    Private _RecordKey_PK As Long
    Private _ThomasGuideGridID As String
    Private _ThomasGuidePageID As String
    Private _FlagVacantPlumbingProtect As String
    Private _FlagBoardedSecured As String
    Private _FlagSecured As String
    Private _FlagAnsul As String
    Private _FlagWindCOvered As String
    Private _WindPercent As Long
    Private _Limits_BusinessIncomeUse As Double
    Private _Premium_BusinessIncomeUse As Double
    Private _Limits_InlandMarine As Double
    Private _Premium_InlandMarine As Double
    Private _Deduct_InlandMarine As Double
    Private _Deduct_Wind As Double
    Private _Deduct_AOP As Double
    Private _Limits_OtherLiab As Double
    Private _Premium_OtherLiab As Double
    Private _Deduct_OtherLiab As Double
    Private _GLIsoCode As String
    Private _Limits_GLAgg As Double
    Private _Limits_GLOccur As Double
    Private _Rate_InlandMarine As String
    Private _Rate_OtherLiab As String
    Private _OtherLiabCoverage As String
    Private _Premium_PropertyTotal As Double
    Private _StreetNumber As String
    Private _StreetName As String
    Private _FlagKeyLocation As String
    Private _FlagEQCovered As String
    Private _FlagFloodCovered As String
    Private _Deduct_Building As Double
    Private _Deduct_Contents As Double
    Private _Deduct_BI As Double
    Private _RackleyRecordKey_FK As Long
    Private _COnstruction As String
    Private _OccupancyID As String
    Private _CoverageOption As String
    Private _CoverageForm As String
    Private _Coinsurance As String
    Private _FlagTheftCovered As String
    Private _Deduct_Theft As Double
    Private _FlagVandalismExcluded As String
    Private _CSPCode As Long
    Private _FlagOpenSides As String
    Private _CauseOfLoss As String
    Private _WindHailCoverage As String
    Private _Limits_LawOrd As Double
    Private _Premium_LawOrd As Double
    Private _UserDefinedField1 As String
    Private _UserDefinedField2 As String
    Private _UserDefinedField3 As String
    Private _UserDefinedLimit1 As String
    Private _UserDefinedLimit2 As String
    Private _UserDefinedLimit3 As String
    Private _UserDefinedDate1 As Date
    Private _UserDefinedDate2 As Date
    Private _UserDefinedValue1 As Double
    Private _UserDefinedValue2 As Double
    Private _UserDefinedValue3 As Double
    Private _UserDefinedID1 As String
    Private _UserDefinedID2 As String
    Private _UserDefinedID3 As String
    Private _ExpiringPremium As Double
    Private _Deduct_Liab As Double
    Private _OrigReferenceKey_FK As Long
    Private _ProgramID As String
    Private _FlagProp As String
    Private _FlagGL As String
    Private _PlacementTypeID As String
    Private _PerilID As String
    Private _ZonePerilID As String
    Private _invoicekey_fk As Long
    Private _BuildingValueLast As Double
    Private _BuildingPremiumLast As Double
    Private _ContentsValueLast As Double
    Private _ContentsPremiumLast As Double
    Private _Limits_BusinessIncomeUseLast As Double
    Private _Premium_BusinessIncomeUseLast As Double
    Private _BuildingValueChange As Double
    Private _BuildingPremiumChange As Double
    Private _ContentsValueChange As Double
    Private _ContentsPremiumChange As Double
    Private _Limits_BusinessIncomeUseChange As Double
    Private _Premium_BusinessIncomeUseChange As Double
    Private _FlagWindPool As String

#End Region
#Region "Properties"

    Public Property ReferenceKey_FK() As Long
        Get
            Return _ReferenceKey_FK
        End Get
        Set(ByVal Value As Long)
            _ReferenceKey_FK = Value
        End Set
    End Property
    Public Property RiskDetailKey_PK() As Long
        Get
            Return _RiskDetailKey_PK
        End Get
        Set(ByVal Value As Long)
            _RiskDetailKey_PK = Value
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
    Public Property ZipCode() As String
        Get
            Return _ZipCode
        End Get
        Set(ByVal Value As String)
            _ZipCode = Value
        End Set
    End Property
    Public Property ZipPlus() As String
        Get
            Return _ZipPlus
        End Get
        Set(ByVal Value As String)
            _ZipPlus = Value
        End Set
    End Property
    Public Property County() As String
        Get
            Return _County
        End Get
        Set(ByVal Value As String)
            _County = Value
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
    Public Property YearBuilt() As Integer
        Get
            Return _YearBuilt
        End Get
        Set(ByVal Value As Integer)
            _YearBuilt = Value
        End Set
    End Property
    Public Property BuildingValue() As Double
        Get
            Return _BuildingValue
        End Get
        Set(ByVal Value As Double)
            _BuildingValue = Value
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
    Public Property DateDropped() As Date
        Get
            Return _DateDropped
        End Get
        Set(ByVal Value As Date)
            _DateDropped = Value
        End Set
    End Property
    Public Property BuildingNumber() As Long
        Get
            Return _BuildingNumber
        End Get
        Set(ByVal Value As Long)
            _BuildingNumber = Value
        End Set
    End Property
    Public Property LocationNumber() As Long
        Get
            Return _LocationNumber
        End Get
        Set(ByVal Value As Long)
            _LocationNumber = Value
        End Set
    End Property
    Public Property Occupancy() As String
        Get
            Return _Occupancy
        End Get
        Set(ByVal Value As String)
            _Occupancy = Value
        End Set
    End Property
    Public Property ConstructionID() As String
        Get
            Return _ConstructionID
        End Get
        Set(ByVal Value As String)
            _ConstructionID = Value
        End Set
    End Property
    Public Property YearManagedSince() As Integer
        Get
            Return _YearManagedSince
        End Get
        Set(ByVal Value As Integer)
            _YearManagedSince = Value
        End Set
    End Property
    Public Property RentableSqFootage() As Long
        Get
            Return _RentableSqFootage
        End Get
        Set(ByVal Value As Long)
            _RentableSqFootage = Value
        End Set
    End Property
    Public Property BuildingFloors() As Long
        Get
            Return _BuildingFloors
        End Get
        Set(ByVal Value As Long)
            _BuildingFloors = Value
        End Set
    End Property
    Public Property ContentsValue() As Double
        Get
            Return _ContentsValue
        End Get
        Set(ByVal Value As Double)
            _ContentsValue = Value
        End Set
    End Property
    Public Property EDPValue() As Double
        Get
            Return _EDPValue
        End Get
        Set(ByVal Value As Double)
            _EDPValue = Value
        End Set
    End Property
    Public Property RentalValue() As Double
        Get
            Return _RentalValue
        End Get
        Set(ByVal Value As Double)
            _RentalValue = Value
        End Set
    End Property
    Public Property OtherValue() As Double
        Get
            Return _OtherValue
        End Get
        Set(ByVal Value As Double)
            _OtherValue = Value
        End Set
    End Property
    Public Property TotalLimits() As Double
        Get
            Return _TotalLimits
        End Get
        Set(ByVal Value As Double)
            _TotalLimits = Value
        End Set
    End Property
    Public Property ActiveFlag() As String
        Get
            Return _ActiveFlag
        End Get
        Set(ByVal Value As String)
            _ActiveFlag = Value
        End Set
    End Property
    Public Property InsuredKey_FK() As Long
        Get
            Return _InsuredKey_FK
        End Get
        Set(ByVal Value As Long)
            _InsuredKey_FK = Value
        End Set
    End Property
    Public Property ContentsPremium() As Double
        Get
            Return _ContentsPremium
        End Get
        Set(ByVal Value As Double)
            _ContentsPremium = Value
        End Set
    End Property
    Public Property EDPPremium() As Double
        Get
            Return _EDPPremium
        End Get
        Set(ByVal Value As Double)
            _EDPPremium = Value
        End Set
    End Property
    Public Property RentalPremium() As Double
        Get
            Return _RentalPremium
        End Get
        Set(ByVal Value As Double)
            _RentalPremium = Value
        End Set
    End Property
    Public Property OtherPremium() As Double
        Get
            Return _OtherPremium
        End Get
        Set(ByVal Value As Double)
            _OtherPremium = Value
        End Set
    End Property
    Public Property FloodQuakeLimits() As Double
        Get
            Return _FloodQuakeLimits
        End Get
        Set(ByVal Value As Double)
            _FloodQuakeLimits = Value
        End Set
    End Property
    Public Property FloodQuakePremium() As Double
        Get
            Return _FloodQuakePremium
        End Get
        Set(ByVal Value As Double)
            _FloodQuakePremium = Value
        End Set
    End Property
    Public Property EmployeeDishonestPremium() As Double
        Get
            Return _EmployeeDishonestPremium
        End Get
        Set(ByVal Value As Double)
            _EmployeeDishonestPremium = Value
        End Set
    End Property
    Public Property LiabilityTotal() As Double
        Get
            Return _LiabilityTotal
        End Get
        Set(ByVal Value As Double)
            _LiabilityTotal = Value
        End Set
    End Property
    Public Property BuildingPremium() As Double
        Get
            Return _BuildingPremium
        End Get
        Set(ByVal Value As Double)
            _BuildingPremium = Value
        End Set
    End Property
    Public Property GLRate() As Double
        Get
            Return _GLRate
        End Get
        Set(ByVal Value As Double)
            _GLRate = Value
        End Set
    End Property
    Public Property GLBasePremium() As Double
        Get
            Return _GLBasePremium
        End Get
        Set(ByVal Value As Double)
            _GLBasePremium = Value
        End Set
    End Property
    Public Property GLOtherPremium() As Double
        Get
            Return _GLOtherPremium
        End Get
        Set(ByVal Value As Double)
            _GLOtherPremium = Value
        End Set
    End Property
    Public Property GLTotalPremium() As Double
        Get
            Return _GLTotalPremium
        End Get
        Set(ByVal Value As Double)
            _GLTotalPremium = Value
        End Set
    End Property
    Public Property PropertyRate() As Double
        Get
            Return _PropertyRate
        End Get
        Set(ByVal Value As Double)
            _PropertyRate = Value
        End Set
    End Property
    Public Property PropertyBasePremium() As Double
        Get
            Return _PropertyBasePremium
        End Get
        Set(ByVal Value As Double)
            _PropertyBasePremium = Value
        End Set
    End Property
    Public Property PropertyOtherPremium() As Double
        Get
            Return _PropertyOtherPremium
        End Get
        Set(ByVal Value As Double)
            _PropertyOtherPremium = Value
        End Set
    End Property
    Public Property PropertyTotalPremium() As Double
        Get
            Return _PropertyTotalPremium
        End Get
        Set(ByVal Value As Double)
            _PropertyTotalPremium = Value
        End Set
    End Property
    Public Property MgtContactKey_FK() As Long
        Get
            Return _MgtContactKey_FK
        End Get
        Set(ByVal Value As Long)
            _MgtContactKey_FK = Value
        End Set
    End Property
    Public Property MgtContact() As String
        Get
            Return _MgtContact
        End Get
        Set(ByVal Value As String)
            _MgtContact = Value
        End Set
    End Property
    Public Property MgtPhone() As String
        Get
            Return _MgtPhone
        End Get
        Set(ByVal Value As String)
            _MgtPhone = Value
        End Set
    End Property
    Public Property InspectionContact() As String
        Get
            Return _InspectionContact
        End Get
        Set(ByVal Value As String)
            _InspectionContact = Value
        End Set
    End Property
    Public Property InspectionPhone() As String
        Get
            Return _InspectionPhone
        End Get
        Set(ByVal Value As String)
            _InspectionPhone = Value
        End Set
    End Property
    Public Property PercentOccupied() As Double
        Get
            Return _PercentOccupied
        End Get
        Set(ByVal Value As Double)
            _PercentOccupied = Value
        End Set
    End Property
    Public Property Acreage() As Long
        Get
            Return _Acreage
        End Get
        Set(ByVal Value As Long)
            _Acreage = Value
        End Set
    End Property
    Public Property NumberElevators() As Long
        Get
            Return _NumberElevators
        End Get
        Set(ByVal Value As Long)
            _NumberElevators = Value
        End Set
    End Property
    Public Property NumberParkingSpaces() As Long
        Get
            Return _NumberParkingSpaces
        End Get
        Set(ByVal Value As Long)
            _NumberParkingSpaces = Value
        End Set
    End Property
    Public Property PercentSprinklered() As Double
        Get
            Return _PercentSprinklered
        End Get
        Set(ByVal Value As Double)
            _PercentSprinklered = Value
        End Set
    End Property
    Public Property NumberCommTenants() As Long
        Get
            Return _NumberCommTenants
        End Get
        Set(ByVal Value As Long)
            _NumberCommTenants = Value
        End Set
    End Property
    Public Property RetailSqFoot() As Long
        Get
            Return _RetailSqFoot
        End Get
        Set(ByVal Value As Long)
            _RetailSqFoot = Value
        End Set
    End Property
    Public Property CommSqFoot() As Long
        Get
            Return _CommSqFoot
        End Get
        Set(ByVal Value As Long)
            _CommSqFoot = Value
        End Set
    End Property
    Public Property ResidentalSqFoot() As Long
        Get
            Return _ResidentalSqFoot
        End Get
        Set(ByVal Value As Long)
            _ResidentalSqFoot = Value
        End Set
    End Property
    Public Property Territory() As String
        Get
            Return _Territory
        End Get
        Set(ByVal Value As String)
            _Territory = Value
        End Set
    End Property
    Public Property TotalPremium() As Double
        Get
            Return _TotalPremium
        End Get
        Set(ByVal Value As Double)
            _TotalPremium = Value
        End Set
    End Property
    Public Property NumberOfBuildings() As Long
        Get
            Return _NumberOfBuildings
        End Get
        Set(ByVal Value As Long)
            _NumberOfBuildings = Value
        End Set
    End Property
    Public Property YearRenovated() As Integer
        Get
            Return _YearRenovated
        End Get
        Set(ByVal Value As Integer)
            _YearRenovated = Value
        End Set
    End Property
    Public Property ParkingType() As String
        Get
            Return _ParkingType
        End Get
        Set(ByVal Value As String)
            _ParkingType = Value
        End Set
    End Property
    Public Property NOTES() As String
        Get
            Return _NOTES
        End Get
        Set(ByVal Value As String)
            _NOTES = Value
        End Set
    End Property
    Public Property InspectionFax() As String
        Get
            Return _InspectionFax
        End Get
        Set(ByVal Value As String)
            _InspectionFax = Value
        End Set
    End Property
    Public Property MgtFax() As String
        Get
            Return _MgtFax
        End Get
        Set(ByVal Value As String)
            _MgtFax = Value
        End Set
    End Property
    Public Property FileReviewed() As Date
        Get
            Return _FileReviewed
        End Get
        Set(ByVal Value As Date)
            _FileReviewed = Value
        End Set
    End Property
    Public Property ProtectionClass() As String
        Get
            Return _ProtectionClass
        End Get
        Set(ByVal Value As String)
            _ProtectionClass = Value
        End Set
    End Property
    Public Property BuildingUpdatesWire() As Integer
        Get
            Return _BuildingUpdatesWire
        End Get
        Set(ByVal Value As Integer)
            _BuildingUpdatesWire = Value
        End Set
    End Property
    Public Property BuildingUpdatesRoofing() As Integer
        Get
            Return _BuildingUpdatesRoofing
        End Get
        Set(ByVal Value As Integer)
            _BuildingUpdatesRoofing = Value
        End Set
    End Property
    Public Property BuildingUpdatesPlumbing() As Integer
        Get
            Return _BuildingUpdatesPlumbing
        End Get
        Set(ByVal Value As Integer)
            _BuildingUpdatesPlumbing = Value
        End Set
    End Property
    Public Property BuildingUpdatesHeating() As Integer
        Get
            Return _BuildingUpdatesHeating
        End Get
        Set(ByVal Value As Integer)
            _BuildingUpdatesHeating = Value
        End Set
    End Property
    Public Property FlagSmokeDetector() As String
        Get
            Return _FlagSmokeDetector
        End Get
        Set(ByVal Value As String)
            _FlagSmokeDetector = Value
        End Set
    End Property
    Public Property FlagLocalBurglar() As String
        Get
            Return _FlagLocalBurglar
        End Get
        Set(ByVal Value As String)
            _FlagLocalBurglar = Value
        End Set
    End Property
    Public Property FlagLocalFire() As String
        Get
            Return _FlagLocalFire
        End Get
        Set(ByVal Value As String)
            _FlagLocalFire = Value
        End Set
    End Property
    Public Property FlagCentralBurglar() As String
        Get
            Return _FlagCentralBurglar
        End Get
        Set(ByVal Value As String)
            _FlagCentralBurglar = Value
        End Set
    End Property
    Public Property FlagCentralFire() As String
        Get
            Return _FlagCentralFire
        End Get
        Set(ByVal Value As String)
            _FlagCentralFire = Value
        End Set
    End Property
    Public Property FlagSprinkled() As String
        Get
            Return _FlagSprinkled
        End Get
        Set(ByVal Value As String)
            _FlagSprinkled = Value
        End Set
    End Property
    Public Property FlagDeadBolt() As String
        Get
            Return _FlagDeadBolt
        End Get
        Set(ByVal Value As String)
            _FlagDeadBolt = Value
        End Set
    End Property
    Public Property RecordKey_PK() As Long
        Get
            Return _RecordKey_PK
        End Get
        Set(ByVal Value As Long)
            _RecordKey_PK = Value
        End Set
    End Property
    Public Property ThomasGuideGridID() As String
        Get
            Return _ThomasGuideGridID
        End Get
        Set(ByVal Value As String)
            _ThomasGuideGridID = Value
        End Set
    End Property
    Public Property ThomasGuidePageID() As String
        Get
            Return _ThomasGuidePageID
        End Get
        Set(ByVal Value As String)
            _ThomasGuidePageID = Value
        End Set
    End Property
    Public Property FlagVacantPlumbingProtect() As String
        Get
            Return _FlagVacantPlumbingProtect
        End Get
        Set(ByVal Value As String)
            _FlagVacantPlumbingProtect = Value
        End Set
    End Property
    Public Property FlagBoardedSecured() As String
        Get
            Return _FlagBoardedSecured
        End Get
        Set(ByVal Value As String)
            _FlagBoardedSecured = Value
        End Set
    End Property
    Public Property FlagSecured() As String
        Get
            Return _FlagSecured
        End Get
        Set(ByVal Value As String)
            _FlagSecured = Value
        End Set
    End Property
    Public Property FlagAnsul() As String
        Get
            Return _FlagAnsul
        End Get
        Set(ByVal Value As String)
            _FlagAnsul = Value
        End Set
    End Property
    Public Property FlagWindCOvered() As String
        Get
            Return _FlagWindCOvered
        End Get
        Set(ByVal Value As String)
            _FlagWindCOvered = Value
        End Set
    End Property
    Public Property WindPercent() As Long
        Get
            Return _WindPercent
        End Get
        Set(ByVal Value As Long)
            _WindPercent = Value
        End Set
    End Property
    Public Property Limits_BusinessIncomeUse() As Double
        Get
            Return _Limits_BusinessIncomeUse
        End Get
        Set(ByVal Value As Double)
            _Limits_BusinessIncomeUse = Value
        End Set
    End Property
    Public Property Premium_BusinessIncomeUse() As Double
        Get
            Return _Premium_BusinessIncomeUse
        End Get
        Set(ByVal Value As Double)
            _Premium_BusinessIncomeUse = Value
        End Set
    End Property
    Public Property Limits_InlandMarine() As Double
        Get
            Return _Limits_InlandMarine
        End Get
        Set(ByVal Value As Double)
            _Limits_InlandMarine = Value
        End Set
    End Property
    Public Property Premium_InlandMarine() As Double
        Get
            Return _Premium_InlandMarine
        End Get
        Set(ByVal Value As Double)
            _Premium_InlandMarine = Value
        End Set
    End Property
    Public Property Deduct_InlandMarine() As Double
        Get
            Return _Deduct_InlandMarine
        End Get
        Set(ByVal Value As Double)
            _Deduct_InlandMarine = Value
        End Set
    End Property
    Public Property Deduct_Wind() As Double
        Get
            Return _Deduct_Wind
        End Get
        Set(ByVal Value As Double)
            _Deduct_Wind = Value
        End Set
    End Property
    Public Property Deduct_AOP() As Double
        Get
            Return _Deduct_AOP
        End Get
        Set(ByVal Value As Double)
            _Deduct_AOP = Value
        End Set
    End Property
    Public Property Limits_OtherLiab() As Double
        Get
            Return _Limits_OtherLiab
        End Get
        Set(ByVal Value As Double)
            _Limits_OtherLiab = Value
        End Set
    End Property
    Public Property Premium_OtherLiab() As Double
        Get
            Return _Premium_OtherLiab
        End Get
        Set(ByVal Value As Double)
            _Premium_OtherLiab = Value
        End Set
    End Property
    Public Property Deduct_OtherLiab() As Double
        Get
            Return _Deduct_OtherLiab
        End Get
        Set(ByVal Value As Double)
            _Deduct_OtherLiab = Value
        End Set
    End Property
    Public Property GLIsoCode() As String
        Get
            Return _GLIsoCode
        End Get
        Set(ByVal Value As String)
            _GLIsoCode = Value
        End Set
    End Property
    Public Property Limits_GLAgg() As Double
        Get
            Return _Limits_GLAgg
        End Get
        Set(ByVal Value As Double)
            _Limits_GLAgg = Value
        End Set
    End Property
    Public Property Limits_GLOccur() As Double
        Get
            Return _Limits_GLOccur
        End Get
        Set(ByVal Value As Double)
            _Limits_GLOccur = Value
        End Set
    End Property
    Public Property Rate_InlandMarine() As String
        Get
            Return _Rate_InlandMarine
        End Get
        Set(ByVal Value As String)
            _Rate_InlandMarine = Value
        End Set
    End Property
    Public Property Rate_OtherLiab() As String
        Get
            Return _Rate_OtherLiab
        End Get
        Set(ByVal Value As String)
            _Rate_OtherLiab = Value
        End Set
    End Property
    Public Property OtherLiabCoverage() As String
        Get
            Return _OtherLiabCoverage
        End Get
        Set(ByVal Value As String)
            _OtherLiabCoverage = Value
        End Set
    End Property
    Public Property Premium_PropertyTotal() As Double
        Get
            Return _Premium_PropertyTotal
        End Get
        Set(ByVal Value As Double)
            _Premium_PropertyTotal = Value
        End Set
    End Property
    Public Property StreetNumber() As String
        Get
            Return _StreetNumber
        End Get
        Set(ByVal Value As String)
            _StreetNumber = Value
        End Set
    End Property
    Public Property StreetName() As String
        Get
            Return _StreetName
        End Get
        Set(ByVal Value As String)
            _StreetName = Value
        End Set
    End Property
    Public Property FlagKeyLocation() As String
        Get
            Return _FlagKeyLocation
        End Get
        Set(ByVal Value As String)
            _FlagKeyLocation = Value
        End Set
    End Property
    Public Property FlagEQCovered() As String
        Get
            Return _FlagEQCovered
        End Get
        Set(ByVal Value As String)
            _FlagEQCovered = Value
        End Set
    End Property
    Public Property FlagFloodCovered() As String
        Get
            Return _FlagFloodCovered
        End Get
        Set(ByVal Value As String)
            _FlagFloodCovered = Value
        End Set
    End Property
    Public Property Deduct_Building() As Double
        Get
            Return _Deduct_Building
        End Get
        Set(ByVal Value As Double)
            _Deduct_Building = Value
        End Set
    End Property
    Public Property Deduct_Contents() As Double
        Get
            Return _Deduct_Contents
        End Get
        Set(ByVal Value As Double)
            _Deduct_Contents = Value
        End Set
    End Property
    Public Property Deduct_BI() As Double
        Get
            Return _Deduct_BI
        End Get
        Set(ByVal Value As Double)
            _Deduct_BI = Value
        End Set
    End Property
    Public Property RackleyRecordKey_FK() As Long
        Get
            Return _RackleyRecordKey_FK
        End Get
        Set(ByVal Value As Long)
            _RackleyRecordKey_FK = Value
        End Set
    End Property
    Public Property COnstruction() As String
        Get
            Return _COnstruction
        End Get
        Set(ByVal Value As String)
            _COnstruction = Value
        End Set
    End Property
    Public Property OccupancyID() As String
        Get
            Return _OccupancyID
        End Get
        Set(ByVal Value As String)
            _OccupancyID = Value
        End Set
    End Property
    Public Property CoverageOption() As String
        Get
            Return _CoverageOption
        End Get
        Set(ByVal Value As String)
            _CoverageOption = Value
        End Set
    End Property
    Public Property CoverageForm() As String
        Get
            Return _CoverageForm
        End Get
        Set(ByVal Value As String)
            _CoverageForm = Value
        End Set
    End Property
    Public Property Coinsurance() As String
        Get
            Return _Coinsurance
        End Get
        Set(ByVal Value As String)
            _Coinsurance = Value
        End Set
    End Property
    Public Property FlagTheftCovered() As String
        Get
            Return _FlagTheftCovered
        End Get
        Set(ByVal Value As String)
            _FlagTheftCovered = Value
        End Set
    End Property
    Public Property Deduct_Theft() As Double
        Get
            Return _Deduct_Theft
        End Get
        Set(ByVal Value As Double)
            _Deduct_Theft = Value
        End Set
    End Property
    Public Property FlagVandalismExcluded() As String
        Get
            Return _FlagVandalismExcluded
        End Get
        Set(ByVal Value As String)
            _FlagVandalismExcluded = Value
        End Set
    End Property
    Public Property CSPCode() As Long
        Get
            Return _CSPCode
        End Get
        Set(ByVal Value As Long)
            _CSPCode = Value
        End Set
    End Property
    Public Property FlagOpenSides() As String
        Get
            Return _FlagOpenSides
        End Get
        Set(ByVal Value As String)
            _FlagOpenSides = Value
        End Set
    End Property
    Public Property CauseOfLoss() As String
        Get
            Return _CauseOfLoss
        End Get
        Set(ByVal Value As String)
            _CauseOfLoss = Value
        End Set
    End Property
    Public Property WindHailCoverage() As String
        Get
            Return _WindHailCoverage
        End Get
        Set(ByVal Value As String)
            _WindHailCoverage = Value
        End Set
    End Property
    Public Property Limits_LawOrd() As Double
        Get
            Return _Limits_LawOrd
        End Get
        Set(ByVal Value As Double)
            _Limits_LawOrd = Value
        End Set
    End Property
    Public Property Premium_LawOrd() As Double
        Get
            Return _Premium_LawOrd
        End Get
        Set(ByVal Value As Double)
            _Premium_LawOrd = Value
        End Set
    End Property
    Public Property UserDefinedField1() As String
        Get
            Return _UserDefinedField1
        End Get
        Set(ByVal Value As String)
            _UserDefinedField1 = Value
        End Set
    End Property
    Public Property UserDefinedField2() As String
        Get
            Return _UserDefinedField2
        End Get
        Set(ByVal Value As String)
            _UserDefinedField2 = Value
        End Set
    End Property
    Public Property UserDefinedField3() As String
        Get
            Return _UserDefinedField3
        End Get
        Set(ByVal Value As String)
            _UserDefinedField3 = Value
        End Set
    End Property
    Public Property UserDefinedLimit1() As String
        Get
            Return _UserDefinedLimit1
        End Get
        Set(ByVal Value As String)
            _UserDefinedLimit1 = Value
        End Set
    End Property
    Public Property UserDefinedLimit2() As String
        Get
            Return _UserDefinedLimit2
        End Get
        Set(ByVal Value As String)
            _UserDefinedLimit2 = Value
        End Set
    End Property
    Public Property UserDefinedLimit3() As String
        Get
            Return _UserDefinedLimit3
        End Get
        Set(ByVal Value As String)
            _UserDefinedLimit3 = Value
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
    Public Property UserDefinedDate2() As Date
        Get
            Return _UserDefinedDate2
        End Get
        Set(ByVal Value As Date)
            _UserDefinedDate2 = Value
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
    Public Property UserDefinedValue2() As Double
        Get
            Return _UserDefinedValue2
        End Get
        Set(ByVal Value As Double)
            _UserDefinedValue2 = Value
        End Set
    End Property
    Public Property UserDefinedValue3() As Double
        Get
            Return _UserDefinedValue3
        End Get
        Set(ByVal Value As Double)
            _UserDefinedValue3 = Value
        End Set
    End Property
    Public Property UserDefinedID1() As String
        Get
            Return _UserDefinedID1
        End Get
        Set(ByVal Value As String)
            _UserDefinedID1 = Value
        End Set
    End Property
    Public Property UserDefinedID2() As String
        Get
            Return _UserDefinedID2
        End Get
        Set(ByVal Value As String)
            _UserDefinedID2 = Value
        End Set
    End Property
    Public Property UserDefinedID3() As String
        Get
            Return _UserDefinedID3
        End Get
        Set(ByVal Value As String)
            _UserDefinedID3 = Value
        End Set
    End Property
    Public Property ExpiringPremium() As Double
        Get
            Return _ExpiringPremium
        End Get
        Set(ByVal Value As Double)
            _ExpiringPremium = Value
        End Set
    End Property
    Public Property Deduct_Liab() As Double
        Get
            Return _Deduct_Liab
        End Get
        Set(ByVal Value As Double)
            _Deduct_Liab = Value
        End Set
    End Property
    Public Property OrigReferenceKey_FK() As Long
        Get
            Return _OrigReferenceKey_FK
        End Get
        Set(ByVal Value As Long)
            _OrigReferenceKey_FK = Value
        End Set
    End Property
    Public Property ProgramID() As String
        Get
            Return _ProgramID
        End Get
        Set(ByVal Value As String)
            _ProgramID = Value
        End Set
    End Property
    Public Property FlagProp() As String
        Get
            Return _FlagProp
        End Get
        Set(ByVal Value As String)
            _FlagProp = Value
        End Set
    End Property
    Public Property FlagGL() As String
        Get
            Return _FlagGL
        End Get
        Set(ByVal Value As String)
            _FlagGL = Value
        End Set
    End Property
    Public Property PlacementTypeID() As String
        Get
            Return _PlacementTypeID
        End Get
        Set(ByVal Value As String)
            _PlacementTypeID = Value
        End Set
    End Property
    Public Property PerilID() As String
        Get
            Return _PerilID
        End Get
        Set(ByVal Value As String)
            _PerilID = Value
        End Set
    End Property
    Public Property ZonePerilID() As String
        Get
            Return _ZonePerilID
        End Get
        Set(ByVal Value As String)
            _ZonePerilID = Value
        End Set
    End Property
    Public Property invoicekey_fk() As Long
        Get
            Return _invoicekey_fk
        End Get
        Set(ByVal Value As Long)
            _invoicekey_fk = Value
        End Set
    End Property
    Public Property BuildingValueLast() As Double
        Get
            Return _BuildingValueLast
        End Get
        Set(ByVal Value As Double)
            _BuildingValueLast = Value
        End Set
    End Property
    Public Property BuildingPremiumLast() As Double
        Get
            Return _BuildingPremiumLast
        End Get
        Set(ByVal Value As Double)
            _BuildingPremiumLast = Value
        End Set
    End Property
    Public Property ContentsValueLast() As Double
        Get
            Return _ContentsValueLast
        End Get
        Set(ByVal Value As Double)
            _ContentsValueLast = Value
        End Set
    End Property
    Public Property ContentsPremiumLast() As Double
        Get
            Return _ContentsPremiumLast
        End Get
        Set(ByVal Value As Double)
            _ContentsPremiumLast = Value
        End Set
    End Property
    Public Property Limits_BusinessIncomeUseLast() As Double
        Get
            Return _Limits_BusinessIncomeUseLast
        End Get
        Set(ByVal Value As Double)
            _Limits_BusinessIncomeUseLast = Value
        End Set
    End Property
    Public Property Premium_BusinessIncomeUseLast() As Double
        Get
            Return _Premium_BusinessIncomeUseLast
        End Get
        Set(ByVal Value As Double)
            _Premium_BusinessIncomeUseLast = Value
        End Set
    End Property
    Public Property BuildingValueChange() As Double
        Get
            Return _BuildingValueChange
        End Get
        Set(ByVal Value As Double)
            _BuildingValueChange = Value
        End Set
    End Property
    Public Property BuildingPremiumChange() As Double
        Get
            Return _BuildingPremiumChange
        End Get
        Set(ByVal Value As Double)
            _BuildingPremiumChange = Value
        End Set
    End Property
    Public Property ContentsValueChange() As Double
        Get
            Return _ContentsValueChange
        End Get
        Set(ByVal Value As Double)
            _ContentsValueChange = Value
        End Set
    End Property
    Public Property ContentsPremiumChange() As Double
        Get
            Return _ContentsPremiumChange
        End Get
        Set(ByVal Value As Double)
            _ContentsPremiumChange = Value
        End Set
    End Property
    Public Property Limits_BusinessIncomeUseChange() As Double
        Get
            Return _Limits_BusinessIncomeUseChange
        End Get
        Set(ByVal Value As Double)
            _Limits_BusinessIncomeUseChange = Value
        End Set
    End Property
    Public Property Premium_BusinessIncomeUseChange() As Double
        Get
            Return _Premium_BusinessIncomeUseChange
        End Get
        Set(ByVal Value As Double)
            _Premium_BusinessIncomeUseChange = Value
        End Set
    End Property
    Public Property FlagWindPool() As String
        Get
            Return _FlagWindPool
        End Get
        Set(ByVal Value As String)
            _FlagWindPool = Value
        End Set
    End Property

#End Region
#Region "Methods"
    Public Function Save(ByRef conn As SqlConnection) As String
        Dim comm As New SqlCommand("SIU_p_InsertRisk_CommProperty", conn)
        Try
            With comm
                .CommandType = CommandType.StoredProcedure
                .CommandTimeout = 0
                .Parameters.Add(New SqlParameter("@ReferenceKey_FK", _ReferenceKey_FK))
                .Parameters.Add(New SqlParameter("@RiskDetailKey_PK", _RiskDetailKey_PK))
                .Parameters.Add(New SqlParameter("@CoverageA", Me.BuildingValue))
                .Parameters.Add(New SqlParameter("@APDeductible", Me.Deduct_AOP))
                .Parameters.Add(New SqlParameter("@WindDeductible", Me.Deduct_Wind))
                '.ExecuteNonQuery()
                .Dispose()
            End With
            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public Function Load(ByRef conn As SqlConnection, ByVal pQuoteID As String) As String
        Try
            Dim comm As New SqlCommand("SIU_p_GetRisk_CommProperty", conn)
            Dim rs As SqlDataReader

            With comm
                .CommandTimeout = 0
                .CommandType = CommandType.StoredProcedure
                .Parameters.AddWithValue("@QuoteID", pQuoteID)
                rs = .ExecuteReader
            End With
            If rs.Read Then
                _ReferenceKey_FK = rs("ReferenceKey_FK")
                _RiskDetailKey_PK = rs("RiskDetailKey_PK")
                _Description = rs("Description")
                _Address1 = rs("Address1")
                _Address2 = rs("Address2")
                _City = rs("City")
                _State = rs("State")
                _ZipCode = rs("ZipCode")
                _ZipPlus = rs("ZipPlus")
                _County = rs("County")
                _Units = rs("Units")
                _YearBuilt = rs("YearBuilt")
                _BuildingValue = rs("BuildingValue")
                _DateAdded = rs("DateAdded")
                _DateDropped = rs("DateDropped")
                _BuildingNumber = rs("BuildingNumber")
                _LocationNumber = rs("LocationNumber")
                _Occupancy = rs("Occupancy")
                _ConstructionID = rs("ConstructionID")
                _YearManagedSince = rs("YearManagedSince")
                _RentableSqFootage = rs("RentableSqFootage")
                _BuildingFloors = rs("BuildingFloors")
                _ContentsValue = rs("ContentsValue")
                _EDPValue = rs("EDPValue")
                _RentalValue = rs("RentalValue")
                _OtherValue = rs("OtherValue")
                _TotalLimits = rs("TotalLimits")
                _ActiveFlag = rs("ActiveFlag")
                _InsuredKey_FK = rs("InsuredKey_FK")
                _ContentsPremium = rs("ContentsPremium")
                _EDPPremium = rs("EDPPremium")
                _RentalPremium = rs("RentalPremium")
                _OtherPremium = rs("OtherPremium")
                _FloodQuakeLimits = rs("FloodQuakeLimits")
                _FloodQuakePremium = rs("FloodQuakePremium")
                _EmployeeDishonestPremium = rs("EmployeeDishonestPremium")
                _LiabilityTotal = rs("LiabilityTotal")
                _BuildingPremium = rs("BuildingPremium")
                _GLRate = rs("GLRate")
                _GLBasePremium = rs("GLBasePremium")
                _GLOtherPremium = rs("GLOtherPremium")
                _GLTotalPremium = rs("GLTotalPremium")
                _PropertyRate = rs("PropertyRate")
                _PropertyBasePremium = rs("PropertyBasePremium")
                _PropertyOtherPremium = rs("PropertyOtherPremium")
                _PropertyTotalPremium = rs("PropertyTotalPremium")
                _MgtContactKey_FK = rs("MgtContactKey_FK")
                _MgtContact = rs("MgtContact")
                _MgtPhone = rs("MgtPhone")
                _InspectionContact = rs("InspectionContact")
                _InspectionPhone = rs("InspectionPhone")
                _PercentOccupied = rs("PercentOccupied")
                _Acreage = rs("Acreage")
                _NumberElevators = rs("NumberElevators")
                _NumberParkingSpaces = rs("NumberParkingSpaces")
                _PercentSprinklered = rs("PercentSprinklered")
                _NumberCommTenants = rs("NumberCommTenants")
                _RetailSqFoot = rs("RetailSqFoot")
                _CommSqFoot = rs("CommSqFoot")
                _ResidentalSqFoot = rs("ResidentalSqFoot")
                _Territory = rs("Territory")
                _TotalPremium = rs("TotalPremium")
                _NumberOfBuildings = rs("NumberOfBuildings")
                _YearRenovated = rs("YearRenovated")
                _ParkingType = rs("ParkingType")
                _NOTES = rs("NOTES")
                _InspectionFax = rs("InspectionFax")
                _MgtFax = rs("MgtFax")
                _FileReviewed = rs("FileReviewed")
                _ProtectionClass = rs("ProtectionClass")
                _BuildingUpdatesWire = rs("BuildingUpdatesWire")
                _BuildingUpdatesRoofing = rs("BuildingUpdatesRoofing")
                _BuildingUpdatesPlumbing = rs("BuildingUpdatesPlumbing")
                _BuildingUpdatesHeating = rs("BuildingUpdatesHeating")
                _FlagSmokeDetector = rs("FlagSmokeDetector")
                _FlagLocalBurglar = rs("FlagLocalBurglar")
                _FlagLocalFire = rs("FlagLocalFire")
                _FlagCentralBurglar = rs("FlagCentralBurglar")
                _FlagCentralFire = rs("FlagCentralFire")
                _FlagSprinkled = rs("FlagSprinkled")
                _FlagDeadBolt = rs("FlagDeadBolt")
                _RecordKey_PK = rs("RecordKey_PK")
                _ThomasGuideGridID = rs("ThomasGuideGridID")
                _ThomasGuidePageID = rs("ThomasGuidePageID")
                _FlagVacantPlumbingProtect = rs("FlagVacantPlumbingProtect")
                _FlagBoardedSecured = rs("FlagBoardedSecured")
                _FlagSecured = rs("FlagSecured")
                _FlagAnsul = rs("FlagAnsul")
                _FlagWindCOvered = rs("FlagWindCOvered")
                _WindPercent = rs("WindPercent")
                _Limits_BusinessIncomeUse = rs("Limits_BusinessIncomeUse")
                _Premium_BusinessIncomeUse = rs("Premium_BusinessIncomeUse")
                _Limits_InlandMarine = rs("Limits_InlandMarine")
                _Premium_InlandMarine = rs("Premium_InlandMarine")
                _Deduct_InlandMarine = rs("Deduct_InlandMarine")
                _Deduct_Wind = rs("Deduct_Wind")
                _Deduct_AOP = rs("Deduct_AOP")
                _Limits_OtherLiab = rs("Limits_OtherLiab")
                _Premium_OtherLiab = rs("Premium_OtherLiab")
                _Deduct_OtherLiab = rs("Deduct_OtherLiab")
                _GLIsoCode = rs("GLIsoCode")
                _Limits_GLAgg = rs("Limits_GLAgg")
                _Limits_GLOccur = rs("Limits_GLOccur")
                _Rate_InlandMarine = rs("Rate_InlandMarine")
                _Rate_OtherLiab = rs("Rate_OtherLiab")
                _OtherLiabCoverage = rs("OtherLiabCoverage")
                _Premium_PropertyTotal = rs("Premium_PropertyTotal")
                _StreetNumber = rs("StreetNumber")
                _StreetName = rs("StreetName")
                _FlagKeyLocation = rs("FlagKeyLocation")
                _FlagEQCovered = rs("FlagEQCovered")
                _FlagFloodCovered = rs("FlagFloodCovered")
                _Deduct_Building = rs("Deduct_Building")
                _Deduct_Contents = rs("Deduct_Contents")
                _Deduct_BI = rs("Deduct_BI")
                _RackleyRecordKey_FK = rs("RackleyRecordKey_FK")
                _COnstruction = rs("COnstruction")
                _OccupancyID = rs("OccupancyID")
                _CoverageOption = rs("CoverageOption")
                _CoverageForm = rs("CoverageForm")
                _Coinsurance = rs("Coinsurance")
                _FlagTheftCovered = rs("FlagTheftCovered")
                _Deduct_Theft = rs("Deduct_Theft")
                _FlagVandalismExcluded = rs("FlagVandalismExcluded")
                _CSPCode = rs("CSPCode")
                _FlagOpenSides = rs("FlagOpenSides")
                _CauseOfLoss = rs("CauseOfLoss")
                _WindHailCoverage = rs("WindHailCoverage")
                _Limits_LawOrd = rs("Limits_LawOrd")
                _Premium_LawOrd = rs("Premium_LawOrd")
                _UserDefinedField1 = rs("UserDefinedField1")
                _UserDefinedField2 = rs("UserDefinedField2")
                _UserDefinedField3 = rs("UserDefinedField3")
                _UserDefinedLimit1 = rs("UserDefinedLimit1")
                _UserDefinedLimit2 = rs("UserDefinedLimit2")
                _UserDefinedLimit3 = rs("UserDefinedLimit3")
                _UserDefinedDate1 = rs("UserDefinedDate1")
                _UserDefinedDate2 = rs("UserDefinedDate2")
                _UserDefinedValue1 = rs("UserDefinedValue1")
                _UserDefinedValue2 = rs("UserDefinedValue2")
                _UserDefinedValue3 = rs("UserDefinedValue3")
                _UserDefinedID1 = rs("UserDefinedID1")
                _UserDefinedID2 = rs("UserDefinedID2")
                _UserDefinedID3 = rs("UserDefinedID3")
                _ExpiringPremium = rs("ExpiringPremium")
                _Deduct_Liab = rs("Deduct_Liab")
                _OrigReferenceKey_FK = rs("OrigReferenceKey_FK")
                _ProgramID = rs("ProgramID")
                _FlagProp = rs("FlagProp")
                _FlagGL = rs("FlagGL")
                _PlacementTypeID = rs("PlacementTypeID")
                _PerilID = rs("PerilID")
                _ZonePerilID = rs("ZonePerilID")
                _invoicekey_fk = rs("invoicekey_fk")
                _BuildingValueLast = rs("BuildingValueLast")
                _BuildingPremiumLast = rs("BuildingPremiumLast")
                _ContentsValueLast = rs("ContentsValueLast")
                _ContentsPremiumLast = rs("ContentsPremiumLast")
                _Limits_BusinessIncomeUseLast = rs("Limits_BusinessIncomeUseLast")
                _Premium_BusinessIncomeUseLast = rs("Premium_BusinessIncomeUseLast")
                _BuildingValueChange = rs("BuildingValueChange")
                _BuildingPremiumChange = rs("BuildingPremiumChange")
                _ContentsValueChange = rs("ContentsValueChange")
                _ContentsPremiumChange = rs("ContentsPremiumChange")
                _Limits_BusinessIncomeUseChange = rs("Limits_BusinessIncomeUseChange")
                _Premium_BusinessIncomeUseChange = rs("Premium_BusinessIncomeUseChange")
                _FlagWindPool = rs("FlagWindPool")
            End If
            rs.Close()

            Return ""
        Catch ex As Exception
            Return "Quote.Risk_CommProperty: " & ex.Message & ex.StackTrace
        End Try
    End Function
#End Region
End Class
