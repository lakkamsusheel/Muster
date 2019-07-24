' -------------------------------------------------------------------------------
' MUSTER.Info.FinancialCommitmentInfo
' Provides the container to persist MUSTER FinancialCommitmentInfo state
' 
' Copyright (C) 2004, 2005 CIBER, Inc.
' All rights reserved.
' 
' Release   Initials    Date        Description
' 1.0        AB       06/24/05    Original class definition.

' 2.0    Thomas Franey 02/24/09   Added Setup Install field to Info Class
' 
' Function          Description
' ---

Namespace MUSTER.Info

    Public Class FinancialCommitmentInfo
        ' Delegate for event to indicate to parent that the info object has been modified in some manner
        Public Delegate Sub FinancialCommitmentChangedEventHandler()
        ' Event that indicates to client that info object has changed in some manner
        ' 

#Region "Private Member Variables"
        Private bolDeleted As Boolean
        Private obolDeleted As Boolean

        Private bolReimburseERAC As Boolean
        Private obolReimburseERAC As Boolean

        Private nCommitmentID As Int64

        Private nFin_Event_ID As Int64
        Private nFundingType As Int64
        Private strPONumber As String
        Private strNewPONumber As String
        Private bolRollOver As Boolean
        Private bolZeroOut As Boolean
        Private dtApprovedDate As Date
        Private dtSOWDate As Date
        Private nContractType As Int64
        Private nActivityType As Int64
        Private strReimbursementCondition As String
        Private dtDueDate As Date
        Private strCase_Letter As String
        Private sERACServices As Double
        Private sLaboratoryServices As Double
        Private sFixedFee As Double
        Private nNumberofEvents As Integer
        Private sWellAbandonment As Double
        Private sFreeProductRecovery As Double
        Private sVacuumContServices As Double
        Private nVacuumContServicesCnt As Integer
        Private sPTTTesting As Double
        Private sERACVacuum As Double
        Private nERACVacuumCnt As Integer
        Private sERACSampling As Double
        Private sIRACServicesEstimate As Double
        Private sSubContractorSvcs As Double
        Private sORCContractorSvcs As Double
        Private sREMContractorSvcs As Double
        Private sPreInstallSetup As Double
        Private sInstallSetup As Double
        Private sMonthlySystemUse As Double
        Private nMonthlySystemUseCnt As Integer
        Private sMonthlyOMSampling As Double
        Private nMonthlyOMSamplingCnt As Integer
        Private sTriAnnualOMSampling As Double
        Private nTriAnnualOMSamplingCnt As Integer
        Private sEstimateTriAnnualLab As Double
        Private nEstimateTriAnnualLabCnt As Integer
        Private sEstimateUtilities As Double
        Private nEstimateUtilitiesCnt As Integer
        Private sThirdPartySettlement As Double
        Private bolThirdPartyPayment As Boolean
        Private strThirdPartyPayee As String
        Private strDueDateStatement As String
        Private strComments As String
        Private sCostRecovery As Double
        Private sMarkup As Double

        Private onFin_Event_ID As Int64
        Private onFundingType As Int64
        Private ostrPONumber As String
        Private ostrNewPONumber As String
        Private obolRollOver As Boolean
        Private obolZeroOut As Boolean
        Private odtApprovedDate As Date
        Private odtSOWDate As Date
        Private onContractType As Int64
        Private onActivityType As Int64
        Private ostrReimbursementCondition As String
        Private odtDueDate As Date
        Private ostrCase_Letter As String
        Private osERACServices As Double
        Private osLaboratoryServices As Double
        Private osFixedFee As Double
        Private onNumberofEvents As Integer
        Private osWellAbandonment As Double
        Private osFreeProductRecovery As Double
        Private osVacuumContServices As Double
        Private onVacuumContServicesCnt As Integer
        Private osPTTTesting As Double
        Private osERACVacuum As Double
        Private onERACVacuumCnt As Integer
        Private osERACSampling As Double
        Private osIRACServicesEstimate As Double
        Private osSubContractorSvcs As Double
        Private osORCContractorSvcs As Double
        Private osREMContractorSvcs As Double
        Private osPreInstallSetup As Double
        Private osInstallSetup As Double

        Private osMonthlySystemUse As Double
        Private onMonthlySystemUseCnt As Integer
        Private osMonthlyOMSampling As Double
        Private onMonthlyOMSamplingCnt As Integer
        Private osTriAnnualOMSampling As Double
        Private onTriAnnualOMSamplingCnt As Integer
        Private osEstimateTriAnnualLab As Double
        Private onEstimateTriAnnualLabCnt As Integer
        Private osEstimateUtilities As Double
        Private onEstimateUtilitiesCnt As Integer
        Private osThirdPartySettlement As Double
        Private obolThirdPartyPayment As Boolean
        Private ostrThirdPartyPayee As String
        Private ostrDueDateStatement As String
        Private ostrComments As String
        Private osCostRecovery As Double
        Private osMarkup As Double


        Private bolIsDirty As Boolean
        Private dtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private dtDataAge As DateTime
        Private dtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private nAgeThreshold As Int16 = 5
        Private nEntityID As Integer


        Private strCreatedBy As String = String.Empty
        Private strModifiedBy As String = String.Empty

        Private ostrCreatedBy As String = String.Empty
        Private ostrModifiedBy As String = String.Empty
        Private odtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private odtCreatedOn As Date = DateTime.Now.ToShortDateString

        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region

#Region "Public Events"
        Public Event FinancialCommitmentInfoChanged As FinancialCommitmentChangedEventHandler
#End Region

#Region "Constructors"
        Public Sub New()
            MyBase.New()
            Me.Init()
            dtDataAge = Now()
        End Sub
        Public Sub New(ByVal Id As Long, _
                        ByVal Fin_Event_ID As Int64, _
                        ByVal FundingType As Int64, _
                        ByVal PONumber As String, _
                        ByVal NewPONumber As String, _
                        ByVal RollOver As Boolean, _
                        ByVal ZeroOut As Boolean, _
                        ByVal ApprovedDate As Date, _
                        ByVal SOWDate As Date, _
                        ByVal ContractType As Int64, _
                        ByVal ActivityType As Int64, _
                        ByVal ReimbursementCondition As String, _
                        ByVal DueDate As Date, _
                        ByVal ThirdPartyPayment As Boolean, _
                        ByVal ThirdPartyPayee As String, _
                        ByVal DueDateStatement As String, _
                        ByVal Case_Letter As String, _
                        ByVal ERACServices As Double, _
                        ByVal LaboratoryServices As Double, _
                        ByVal FixedFee As Double, _
                        ByVal NumberofEvents As Integer, _
                        ByVal WellAbandonment As Double, _
                        ByVal FreeProductRecovery As Double, _
                        ByVal VacuumContServices As Double, _
                        ByVal VacuumContServicesCnt As Integer, _
                        ByVal PTTTesting As Double, _
                        ByVal ERACVacuum As Double, _
                        ByVal ERACVacuumCnt As Integer, _
                        ByVal ERACSampling As Double, _
                        ByVal IRACServicesEstimate As Double, _
                        ByVal SubContractorSvcs As Double, _
                        ByVal ORCContractorSvcs As Double, _
                        ByVal REMContractorSvcs As Double, _
                        ByVal PreInstallSetup As Double, _
                        ByVal MonthlySystemUse As Double, _
                        ByVal MonthlySystemUseCnt As Integer, _
                        ByVal MonthlyOMSampling As Double, _
                        ByVal MonthlyOMSamplingCnt As Integer, _
                        ByVal TriAnnualOMSampling As Double, _
                        ByVal TriAnnualOMSamplingCnt As Integer, _
                        ByVal EstimateTriAnnualLab As Double, _
                        ByVal EstimateTriAnnualLabCnt As Integer, _
                        ByVal EstimateUtilities As Double, _
                        ByVal EstimateUtilitiesCnt As Integer, _
                        ByVal ThirdPartySettlement As Double, _
                        ByVal Comments As String, _
                        ByVal CreatedBy As String, _
                        ByVal CreateDate As Date, _
                        ByVal LastEditedBy As String, _
                        ByVal LastEditDate As Date, _
                        ByVal bDeleted As Boolean, _
                        ByVal CostRecovery As Double, _
                        ByVal Markup As Double, _
                        Optional ByVal InstallSetup As Double = 0, _
                        Optional ByVal ReimburseERAC As Boolean = False)


            nCommitmentID = Id
            onFin_Event_ID = Fin_Event_ID
            onFundingType = FundingType
            ostrPONumber = PONumber
            ostrNewPONumber = NewPONumber
            obolRollOver = RollOver
            obolZeroOut = ZeroOut
            odtApprovedDate = ApprovedDate
            odtSOWDate = SOWDate
            onContractType = ContractType
            onActivityType = ActivityType
            ostrReimbursementCondition = ReimbursementCondition
            odtDueDate = DueDate
            ostrCase_Letter = Case_Letter
            osERACServices = ERACServices
            osLaboratoryServices = LaboratoryServices
            osFixedFee = FixedFee
            onNumberofEvents = NumberofEvents
            osWellAbandonment = WellAbandonment
            osFreeProductRecovery = FreeProductRecovery
            osVacuumContServices = VacuumContServices
            onVacuumContServicesCnt = VacuumContServicesCnt
            osPTTTesting = PTTTesting
            osERACVacuum = ERACVacuum
            onERACVacuumCnt = ERACVacuumCnt
            osERACSampling = ERACSampling
            osIRACServicesEstimate = IRACServicesEstimate
            osSubContractorSvcs = SubContractorSvcs
            osORCContractorSvcs = ORCContractorSvcs
            osREMContractorSvcs = REMContractorSvcs
            osPreInstallSetup = PreInstallSetup
            osMonthlySystemUse = MonthlySystemUse
            onMonthlySystemUseCnt = MonthlySystemUseCnt
            osMonthlyOMSampling = MonthlyOMSampling
            onMonthlyOMSamplingCnt = MonthlyOMSamplingCnt
            osTriAnnualOMSampling = TriAnnualOMSampling
            onTriAnnualOMSamplingCnt = TriAnnualOMSamplingCnt
            osEstimateTriAnnualLab = EstimateTriAnnualLab
            onEstimateTriAnnualLabCnt = EstimateTriAnnualLabCnt
            osEstimateUtilities = EstimateUtilities
            onEstimateUtilitiesCnt = EstimateUtilitiesCnt
            osThirdPartySettlement = ThirdPartySettlement
            obolThirdPartyPayment = ThirdPartyPayment
            ostrThirdPartyPayee = ThirdPartyPayee
            ostrDueDateStatement = DueDateStatement
            ostrComments = Comments
            osCostRecovery = CostRecovery
            osMarkup = Markup
            osInstallSetup = InstallSetup
            ostrCreatedBy = CreatedBy
            odtCreatedOn = CreateDate
            ostrModifiedBy = LastEditedBy
            odtModifiedOn = LastEditDate
            obolDeleted = bDeleted
            obolReimburseERAC = ReimburseERAC

            dtDataAge = Now()
            Me.Reset()

        End Sub
#End Region

#Region "Exposed Methods"
        ' Add other attributes as necessitated by design
        Public Sub Archive()

            onFin_Event_ID = nFin_Event_ID
            onFundingType = nFundingType
            ostrPONumber = strPONumber
            ostrNewPONumber = strNewPONumber
            obolRollOver = bolRollOver
            obolZeroOut = bolZeroOut
            odtApprovedDate = dtApprovedDate
            odtSOWDate = dtSOWDate
            onContractType = nContractType
            onActivityType = nActivityType
            ostrReimbursementCondition = strReimbursementCondition
            odtDueDate = dtDueDate
            ostrCase_Letter = strCase_Letter
            osERACServices = sERACServices
            osLaboratoryServices = sLaboratoryServices
            osFixedFee = sFixedFee
            onNumberofEvents = nNumberofEvents
            osWellAbandonment = sWellAbandonment
            osFreeProductRecovery = sFreeProductRecovery
            osVacuumContServices = sVacuumContServices
            onVacuumContServicesCnt = nVacuumContServicesCnt
            osPTTTesting = sPTTTesting
            osERACVacuum = sERACVacuum
            onERACVacuumCnt = nERACVacuumCnt
            osERACSampling = sERACSampling
            osIRACServicesEstimate = sIRACServicesEstimate
            osSubContractorSvcs = sSubContractorSvcs
            osORCContractorSvcs = sORCContractorSvcs
            osREMContractorSvcs = sREMContractorSvcs
            osPreInstallSetup = sPreInstallSetup
            osInstallSetup = sInstallSetup
            osMonthlySystemUse = sMonthlySystemUse
            onMonthlySystemUseCnt = nMonthlySystemUseCnt
            osMonthlyOMSampling = sMonthlyOMSampling
            onMonthlyOMSamplingCnt = nMonthlyOMSamplingCnt
            osTriAnnualOMSampling = sTriAnnualOMSampling
            onTriAnnualOMSamplingCnt = nTriAnnualOMSamplingCnt
            osEstimateTriAnnualLab = sEstimateTriAnnualLab
            onEstimateTriAnnualLabCnt = nEstimateTriAnnualLabCnt
            osEstimateUtilities = sEstimateUtilities
            onEstimateUtilitiesCnt = nEstimateUtilitiesCnt
            osThirdPartySettlement = sThirdPartySettlement
            obolThirdPartyPayment = bolThirdPartyPayment
            ostrThirdPartyPayee = strThirdPartyPayee
            ostrDueDateStatement = strDueDateStatement
            ostrComments = strComments
            osCostRecovery = sCostRecovery
            osMarkup = sMarkup

            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn

            obolDeleted = bolDeleted
            obolReimburseERAC = bolReimburseERAC

        End Sub

        Public Sub Reset()


            nFin_Event_ID = onFin_Event_ID
            nFundingType = onFundingType
            strPONumber = ostrPONumber
            strNewPONumber = ostrNewPONumber
            bolRollOver = obolRollOver
            bolZeroOut = obolZeroOut
            dtApprovedDate = odtApprovedDate
            dtSOWDate = odtSOWDate
            nContractType = onContractType
            nActivityType = onActivityType
            strReimbursementCondition = ostrReimbursementCondition
            dtDueDate = odtDueDate
            strCase_Letter = ostrCase_Letter
            sERACServices = osERACServices
            sLaboratoryServices = osLaboratoryServices
            sFixedFee = osFixedFee
            nNumberofEvents = onNumberofEvents
            sWellAbandonment = osWellAbandonment
            sFreeProductRecovery = osFreeProductRecovery
            sVacuumContServices = osVacuumContServices
            nVacuumContServicesCnt = onVacuumContServicesCnt
            sPTTTesting = osPTTTesting
            sERACVacuum = osERACVacuum
            nERACVacuumCnt = onERACVacuumCnt
            sERACSampling = osERACSampling
            sIRACServicesEstimate = osIRACServicesEstimate
            sSubContractorSvcs = osSubContractorSvcs
            sORCContractorSvcs = osORCContractorSvcs
            sREMContractorSvcs = osREMContractorSvcs
            sPreInstallSetup = osPreInstallSetup
            sInstallSetup = osInstallSetup
            sMonthlySystemUse = osMonthlySystemUse
            nMonthlySystemUseCnt = onMonthlySystemUseCnt
            sMonthlyOMSampling = osMonthlyOMSampling
            nMonthlyOMSamplingCnt = onMonthlyOMSamplingCnt
            sTriAnnualOMSampling = osTriAnnualOMSampling
            nTriAnnualOMSamplingCnt = onTriAnnualOMSamplingCnt
            sEstimateTriAnnualLab = osEstimateTriAnnualLab
            nEstimateTriAnnualLabCnt = onEstimateTriAnnualLabCnt
            sEstimateUtilities = osEstimateUtilities
            nEstimateUtilitiesCnt = onEstimateUtilitiesCnt
            sThirdPartySettlement = osThirdPartySettlement
            bolThirdPartyPayment = obolThirdPartyPayment
            strThirdPartyPayee = ostrThirdPartyPayee
            strDueDateStatement = ostrDueDateStatement
            strComments = ostrComments
            sCostRecovery = osCostRecovery
            sMarkup = osMarkup

            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn

            bolDeleted = obolDeleted
            bolReimburseERAC = obolReimburseERAC

        End Sub

#End Region

#Region "Private Methods"

        Private Sub CheckDirty()
            Dim bolOldState As Boolean = bolIsDirty

            bolIsDirty = (onFin_Event_ID <> nFin_Event_ID) Or _
                        (onFundingType <> nFundingType) Or _
                        (ostrPONumber <> strPONumber) Or _
                        (ostrNewPONumber <> strNewPONumber) Or _
                        (obolRollOver <> bolRollOver) Or _
                        (obolZeroOut <> bolZeroOut) Or _
                        (odtApprovedDate <> dtApprovedDate) Or _
                        (odtSOWDate <> dtSOWDate) Or _
                        (onContractType <> nContractType) Or _
                        (onActivityType <> nActivityType) Or _
                        (ostrReimbursementCondition <> strReimbursementCondition) Or _
                        (odtDueDate <> dtDueDate) Or _
                        (ostrCase_Letter <> strCase_Letter) Or _
                        (osERACServices <> sERACServices) Or _
                        (osLaboratoryServices <> sLaboratoryServices) Or _
                        (osFixedFee <> sFixedFee) Or _
                        (onNumberofEvents <> nNumberofEvents) Or _
                        (osWellAbandonment <> sWellAbandonment) Or _
                        (osFreeProductRecovery <> sFreeProductRecovery) Or _
                        (osVacuumContServices <> sVacuumContServices) Or _
                        (onVacuumContServicesCnt <> nVacuumContServicesCnt) Or _
                        (osPTTTesting <> sPTTTesting) Or _
                        (osERACVacuum <> sERACVacuum) Or _
                        (onERACVacuumCnt <> nERACVacuumCnt) Or _
                        (osERACSampling <> sERACSampling) Or _
                        (osIRACServicesEstimate <> sIRACServicesEstimate) Or _
                        (osSubContractorSvcs <> sSubContractorSvcs) Or _
                        (osORCContractorSvcs <> sORCContractorSvcs) Or _
                        (osREMContractorSvcs <> sREMContractorSvcs) Or _
                        (osPreInstallSetup <> sPreInstallSetup) Or _
                        (osInstallSetup <> sInstallSetup) Or _
                        (osMonthlySystemUse <> sMonthlySystemUse) Or _
                        (onMonthlySystemUseCnt <> nMonthlySystemUseCnt) Or _
                        (osMonthlyOMSampling <> sMonthlyOMSampling) Or _
                        (onMonthlyOMSamplingCnt <> nMonthlyOMSamplingCnt) Or _
                        (osTriAnnualOMSampling <> sTriAnnualOMSampling) Or _
                        (onTriAnnualOMSamplingCnt <> nTriAnnualOMSamplingCnt) Or _
                        (osEstimateTriAnnualLab <> sEstimateTriAnnualLab) Or _
                        (onEstimateTriAnnualLabCnt <> nEstimateTriAnnualLabCnt) Or _
                        (osEstimateUtilities <> sEstimateUtilities) Or _
                        (onEstimateUtilitiesCnt <> nEstimateUtilitiesCnt) Or _
                        (osThirdPartySettlement <> sThirdPartySettlement) Or _
                        (obolThirdPartyPayment <> bolThirdPartyPayment) Or _
                        (ostrComments <> strComments) Or _
                        (ostrDueDateStatement <> strDueDateStatement) Or _
                        (obolDeleted <> bolDeleted) Or _
                        (osCostRecovery <> sCostRecovery) Or _
                        (osMarkup <> sMarkup) Or _
                        (obolReimburseERAC <> bolReimburseERAC)


        End Sub

        Public Sub Init()
            Dim tmpDate As Date

            nCommitmentID = 0
            onFin_Event_ID = 0
            onFundingType = 0
            ostrPONumber = String.Empty
            ostrNewPONumber = String.Empty
            obolRollOver = False
            obolZeroOut = False
            odtApprovedDate = "01/01/0001"
            odtSOWDate = "01/01/0001"
            onContractType = 0
            onActivityType = 0
            ostrReimbursementCondition = String.Empty
            odtDueDate = "01/01/0001"
            ostrCase_Letter = String.Empty
            osERACServices = 0
            osLaboratoryServices = 0
            osFixedFee = 0
            onNumberofEvents = 0
            osWellAbandonment = 0
            osFreeProductRecovery = 0
            osVacuumContServices = 0
            onVacuumContServicesCnt = 0
            osPTTTesting = 0
            osERACVacuum = 0
            onERACVacuumCnt = 0
            osERACSampling = 0
            osIRACServicesEstimate = 0
            osSubContractorSvcs = 0
            osORCContractorSvcs = 0
            osREMContractorSvcs = 0
            osPreInstallSetup = 0
            osInstallSetup = 0
            osMonthlySystemUse = 0
            onMonthlySystemUseCnt = 0
            osMonthlyOMSampling = 0
            onMonthlyOMSamplingCnt = 0
            osTriAnnualOMSampling = 0
            onTriAnnualOMSamplingCnt = 0
            osEstimateTriAnnualLab = 0
            onEstimateTriAnnualLabCnt = 0
            osEstimateUtilities = 0
            onEstimateUtilitiesCnt = 0
            osThirdPartySettlement = 0
            osCostRecovery = 0
            osMarkup = 0
            ostrComments = String.Empty
            strCreatedBy = String.Empty
            dtCreatedOn = tmpDate
            strModifiedBy = String.Empty
            dtModifiedOn = tmpDate
            obolThirdPartyPayment = False
            ostrThirdPartyPayee = String.Empty
            ostrDueDateStatement = String.Empty
            obolDeleted = False
            obolReimburseERAC = False
        End Sub
#End Region

#Region "Protected Methods"
        Protected Overrides Sub Finalize()
        End Sub
#End Region

#Region "Exposed Attributes"
        ' the uniqueIdetifier for the _ProtoInfo
        Public Property CommitmentID() As Int64
            Get
                Return nCommitmentID
            End Get
            Set(ByVal Value As Int64)
                nCommitmentID = Value
            End Set
        End Property


        Public Property Fin_Event_ID() As Int64
            Get
                Return nFin_Event_ID
            End Get
            Set(ByVal Value As Int64)
                nFin_Event_ID = Value
                CheckDirty()
            End Set
        End Property



        Public Property FundingType() As Int64
            Get
                Return nFundingType
            End Get
            Set(ByVal Value As Int64)
                nFundingType = Value
                CheckDirty()
            End Set
        End Property
        Public Property PONumber() As String
            Get
                Return strPONumber
            End Get
            Set(ByVal Value As String)
                strPONumber = Value
                CheckDirty()
            End Set
        End Property
        Public Property Comments() As String
            Get
                Return strComments
            End Get
            Set(ByVal Value As String)
                strComments = Value
                CheckDirty()
            End Set
        End Property
        Public Property NewPONumber() As String
            Get
                Return strNewPONumber
            End Get
            Set(ByVal Value As String)
                strNewPONumber = Value
                CheckDirty()
            End Set
        End Property
        Public Property RollOver() As Boolean
            Get
                Return bolRollOver
            End Get
            Set(ByVal Value As Boolean)
                bolRollOver = Value
                CheckDirty()
            End Set
        End Property
        Public Property ZeroOut() As Boolean
            Get
                Return bolZeroOut
            End Get
            Set(ByVal Value As Boolean)
                bolZeroOut = Value
                CheckDirty()
            End Set
        End Property
        Public Property ApprovedDate() As Date
            Get
                Return dtApprovedDate
            End Get
            Set(ByVal Value As Date)
                dtApprovedDate = Value
                CheckDirty()
            End Set
        End Property
        Public Property SOWDate() As Date
            Get
                Return dtSOWDate
            End Get
            Set(ByVal Value As Date)
                dtSOWDate = Value
                CheckDirty()
            End Set
        End Property
        Public Property ContractType() As Int64
            Get
                Return nContractType
            End Get
            Set(ByVal Value As Int64)
                nContractType = Value
                CheckDirty()
            End Set
        End Property
        Public Property ActivityType() As Int64
            Get
                Return nActivityType
            End Get
            Set(ByVal Value As Int64)
                nActivityType = Value
                CheckDirty()
            End Set
        End Property
        Public Property ReimbursementCondition() As String
            Get
                Return strReimbursementCondition
            End Get
            Set(ByVal Value As String)
                strReimbursementCondition = Value
                CheckDirty()
            End Set
        End Property
        Public Property DueDate() As Date
            Get
                Return dtDueDate
            End Get
            Set(ByVal Value As Date)
                dtDueDate = Value
                CheckDirty()
            End Set
        End Property
        Public Property ThirdPartyPayment() As Boolean
            Get
                Return bolThirdPartyPayment
            End Get
            Set(ByVal Value As Boolean)
                bolThirdPartyPayment = Value
                CheckDirty()
            End Set
        End Property
        Public Property ThirdPartyPayee() As String
            Get
                Return strThirdPartyPayee
            End Get
            Set(ByVal Value As String)
                strThirdPartyPayee = Value
                CheckDirty()
            End Set
        End Property
        Public Property DueDateStatement() As String
            Get
                Return strDueDateStatement
            End Get
            Set(ByVal Value As String)
                strDueDateStatement = Value
                CheckDirty()
            End Set
        End Property
        Public Property Case_Letter() As String
            Get
                Return strCase_Letter
            End Get
            Set(ByVal Value As String)
                strCase_Letter = Value
                CheckDirty()
            End Set
        End Property
        Public Property ERACServices() As Double
            Get
                Return sERACServices
            End Get
            Set(ByVal Value As Double)
                sERACServices = Value
                CheckDirty()
            End Set
        End Property
        Public Property LaboratoryServices() As Double
            Get
                Return sLaboratoryServices

            End Get
            Set(ByVal Value As Double)
                sLaboratoryServices = Value
                CheckDirty()
            End Set
        End Property
        Public Property FixedFee() As Double
            Get
                Return sFixedFee
            End Get
            Set(ByVal Value As Double)
                sFixedFee = Value
                CheckDirty()
            End Set
        End Property
        Public Property NumberofEvents() As Integer
            Get
                Return nNumberofEvents
            End Get
            Set(ByVal Value As Integer)
                nNumberofEvents = Value
                CheckDirty()
            End Set
        End Property
        Public Property WellAbandonment() As Double
            Get
                Return sWellAbandonment
            End Get
            Set(ByVal Value As Double)
                sWellAbandonment = Value
                CheckDirty()
            End Set
        End Property
        Public Property FreeProductRecovery() As Double
            Get
                Return sFreeProductRecovery
            End Get
            Set(ByVal Value As Double)
                sFreeProductRecovery = Value
                CheckDirty()
            End Set
        End Property
        Public Property VacuumContServices() As Double
            Get
                Return sVacuumContServices
            End Get
            Set(ByVal Value As Double)
                sVacuumContServices = Value
                CheckDirty()
            End Set
        End Property
        Public Property VacuumContServicesCnt() As Integer
            Get
                Return nVacuumContServicesCnt
            End Get
            Set(ByVal Value As Integer)
                nVacuumContServicesCnt = Value
                CheckDirty()
            End Set
        End Property
        Public Property PTTTesting() As Double
            Get
                Return sPTTTesting
            End Get
            Set(ByVal Value As Double)
                sPTTTesting = Value
                CheckDirty()
            End Set
        End Property
        Public Property ERACVacuum() As Double
            Get
                Return sERACVacuum
            End Get
            Set(ByVal Value As Double)
                sERACVacuum = Value
                CheckDirty()
            End Set
        End Property
        Public Property ERACVacuumCnt() As Integer
            Get
                Return nERACVacuumCnt
            End Get
            Set(ByVal Value As Integer)
                nERACVacuumCnt = Value
                CheckDirty()
            End Set
        End Property
        Public Property ERACSampling() As Double
            Get
                Return sERACSampling
            End Get
            Set(ByVal Value As Double)
                sERACSampling = Value
                CheckDirty()
            End Set
        End Property
        Public Property IRACServicesEstimate() As Double
            Get
                Return sIRACServicesEstimate
            End Get
            Set(ByVal Value As Double)
                sIRACServicesEstimate = Value
                CheckDirty()
            End Set
        End Property
        Public Property SubContractorSvcs() As Double
            Get
                Return sSubContractorSvcs
            End Get
            Set(ByVal Value As Double)
                sSubContractorSvcs = Value
                CheckDirty()
            End Set
        End Property
        Public Property ORCContractorSvcs() As Double
            Get
                Return sORCContractorSvcs
            End Get
            Set(ByVal Value As Double)
                sORCContractorSvcs = Value
                CheckDirty()
            End Set
        End Property
        Public Property REMContractorSvcs() As Double
            Get
                Return sREMContractorSvcs
            End Get
            Set(ByVal Value As Double)
                sREMContractorSvcs = Value
                CheckDirty()
            End Set
        End Property
        Public Property PreInstallSetup() As Double
            Get
                Return sPreInstallSetup
            End Get
            Set(ByVal Value As Double)
                sPreInstallSetup = Value
                CheckDirty()
            End Set
        End Property

        Public Property InstallSetup() As Double
            Get
                Return sInstallSetup
            End Get
            Set(ByVal Value As Double)
                sInstallSetup = Value
                CheckDirty()
            End Set
        End Property

        Public Property MonthlySystemUse() As Double
            Get
                Return sMonthlySystemUse
            End Get
            Set(ByVal Value As Double)
                sMonthlySystemUse = Value
                CheckDirty()
            End Set
        End Property
        Public Property MonthlySystemUseCnt() As Integer
            Get
                Return nMonthlySystemUseCnt
            End Get
            Set(ByVal Value As Integer)
                nMonthlySystemUseCnt = Value
                CheckDirty()
            End Set
        End Property
        Public Property MonthlyOMSampling() As Double
            Get
                Return sMonthlyOMSampling
            End Get
            Set(ByVal Value As Double)
                sMonthlyOMSampling = Value
                CheckDirty()
            End Set
        End Property
        Public Property MonthlyOMSamplingCnt() As Integer
            Get
                Return nMonthlyOMSamplingCnt
            End Get
            Set(ByVal Value As Integer)
                nMonthlyOMSamplingCnt = Value
                CheckDirty()
            End Set
        End Property
        Public Property TriAnnualOMSampling() As Double
            Get
                Return sTriAnnualOMSampling
            End Get
            Set(ByVal Value As Double)
                sTriAnnualOMSampling = Value
                CheckDirty()
            End Set
        End Property
        Public Property TriAnnualOMSamplingCnt() As Integer
            Get
                Return nTriAnnualOMSamplingCnt
            End Get
            Set(ByVal Value As Integer)
                nTriAnnualOMSamplingCnt = Value
                CheckDirty()
            End Set
        End Property
        Public Property EstimateTriAnnualLab() As Double
            Get
                Return sEstimateTriAnnualLab
            End Get
            Set(ByVal Value As Double)
                sEstimateTriAnnualLab = Value
                CheckDirty()
            End Set
        End Property
        Public Property EstimateTriAnnualLabCnt() As Integer
            Get
                Return nEstimateTriAnnualLabCnt
            End Get
            Set(ByVal Value As Integer)
                nEstimateTriAnnualLabCnt = Value
                CheckDirty()
            End Set
        End Property
        Public Property EstimateUtilities() As Double
            Get
                Return sEstimateUtilities
            End Get
            Set(ByVal Value As Double)
                sEstimateUtilities = Value
                CheckDirty()
            End Set
        End Property
        Public Property EstimateUtilitiesCnt() As Integer
            Get
                Return nEstimateUtilitiesCnt
            End Get
            Set(ByVal Value As Integer)
                nEstimateUtilitiesCnt = Value
                CheckDirty()
            End Set
        End Property
        Public Property ThirdPartySettlement() As Double
            Get
                Return sThirdPartySettlement
            End Get
            Set(ByVal Value As Double)
                sThirdPartySettlement = Value
                CheckDirty()
            End Set
        End Property
        Public Property CostRecovery() As Double
            Get
                Return sCostRecovery
            End Get
            Set(ByVal Value As Double)
                sCostRecovery = Value
                CheckDirty()
            End Set
        End Property
        Public Property Markup() As Double
            Get
                Return sMarkup
            End Get
            Set(ByVal Value As Double)
                sMarkup = Value
                CheckDirty()
            End Set
        End Property

        ' The maximum age the info object can attain before requiring a refresh
        Public Property AgeThreshold() As Date
            Get
                Return dtDataAge
            End Get
            Set(ByVal Value As Date)
                dtDataAge = Value
            End Set
        End Property
        ' 
        ' The ID of the user that created the row
        Public Property CreatedBy() As String
            Get
                Return strCreatedBy
            End Get
            Set(ByVal Value As String)
                strCreatedBy = Value
            End Set
        End Property
        ' The date on which the row was created
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return dtCreatedOn
            End Get
        End Property
        ' The deleted flag for the TEC_ACT
        Public Property Deleted() As Boolean
            Get
                Return bolDeleted
            End Get
            Set(ByVal Value As Boolean)
                bolDeleted = Value
                CheckDirty()
            End Set
        End Property

        ' reimburse ERAC flag (to be used in reimbursements)
        Public Property ReimburseERAC() As Boolean
            Get
                Return bolReimburseERAC
            End Get
            Set(ByVal Value As Boolean)
                bolReimburseERAC = Value
                CheckDirty()
            End Set
        End Property


        ' The entity ID associated.
        Public ReadOnly Property EntityID() As Integer
            Get
            End Get
        End Property


        ' Raised when any of the _ProtoInfo attributes are modified
        Public Property IsDirty() As Boolean
            Get
                Return bolIsDirty
            End Get
            Set(ByVal Value As Boolean)
                bolIsDirty = False
            End Set
        End Property
        Public Property ModifiedBy() As String
            Get
                Return strModifiedBy
            End Get
            Set(ByVal Value As String)
                strModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return dtModifiedOn
            End Get
        End Property
        ' Returns a boolean indicating if the data has aged beyond its preset limit
        Protected ReadOnly Property IsAgedData() As Boolean
            Get
                Return IIf(DateDiff(DateInterval.Minute, dtDataAge, Now()) >= nAgeThreshold, True, False)
            End Get
        End Property
#End Region

    End Class

End Namespace
