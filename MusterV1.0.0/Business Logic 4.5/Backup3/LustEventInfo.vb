'-------------------------------------------------------------------------------
' MUSTER.Info.LustEventInfo
'   Provides the container to persist MUSTER LustEvent state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        JVC       03/02/05    Original class definition.
'
' Function          Description
'-------------------------------------------------------------------------------
' New()             Instantiates an empty LustEventInfo object.
' New(oLustEvent)   Instantiates a populated LustEventInfo object.
' New(ByVal EventID , GWS , LDR , PTT , ReportDate , StartDate , EventStatus , FacilityEventSeqNo, FlagID, MGPTFStatus, Priority, ProjMgr, ReleaseStatus, ReportSource, RelatedSites, CREATED_BY, CREATE_DATE, LAST_EDITED_BY, DATE_LAST_EDITED)
'                   Instantiates a populated LustEventInfo object.
'
' Reset()           Sets the object state to the original state when loaded from or
'                       last saved to the repository.
' Save()            Saves the object state to the repository.
' Archive()         Replaces the current value of the object to the one in collection
' CheckDirty()      Checks if the values are different in the collection and the current object
' Init()            Initializes the member variables to their default values
'
' Attribute          Description
'-------------------------------------------------------------------------------
' AgeThreshold       The maximum age the info object can attain before requiring a refresh
' CreatedBy          The ID of the user that created the row
' CreatedOn          The date on which the row was created
' Deleted            Indicates the deleted state of the row
' EventStatus        The current status of the LUST event - derived from properties table
' FacilityEventID    The LUST event ID for the facility
' FlagID             The system ID of the flag associated with the LUST event
' ID                 The system ID for this LUST event
' IsDirty            Returns a Boolean if the object has changed from its original status
' LastGWS            Date of last GWS for LUST event
' LastLDR            The date of the last LDR for the LUST event
' LastPTT            The date of the last PTT for the LUST event
' MGPTFStatus        The MGPTF status for the LUST event - derived from properties table
' ModifiedBy         ID of the user that last made changes
' ModifiedOn         The date of the last changes made 
' PM                 The current Project Manager for the LUST event
' Priority           The current priority for the LUST event - can only range 1 to 8
' ReleaseStatus      The current release status for the LUST event - derived from properties table
' ReportDate         The date the LUST event was reported to UST
' ReportSource       The source of the LUST event report to UST - derived from properties table
' Started            The date the LUST event was started
'
Namespace MUSTER.Info
    Public Class LustEventInfo
#Region "Public Events"
        ' Raised when any of the LustEventInfo attributes are modified
        Public Event LustEventInfoChanged As LustEventInfoChangedEventHandler
        Public Delegate Sub LustEventInfoChangedEventHandler()
#End Region
#Region "Private Member Variables"
        Private bolDeleted As Boolean
        Private bolIsDirty As Boolean
        Private bolIsEligibityDirty As Boolean
        Private dtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private dtDataAge As DateTime
        Private dtGWS As Date
        Private dtLDR As Date
        Private dtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private dtPTT As Date
        Private dtReportDate As Date
        Private dtCompAssDate As Date
        Private dtStartDate As Date
        Private nAgeThreshold As Int16 = 5
        Private nEventID As Int64
        Private nEventStatus As Int32
        Private nFacilityID As Int16
        Private nFlagID As Int64
        Private nMGPTFStatus As Int16
        Private nPriority As Int16
        Private nProjMgr As Int32
        Private nReleaseStatus As Int32
        Private nReportSource As Int32

        Private obolDeleted As Boolean
        Private obolIsDirty As Boolean
        Private obolIsEligibityDirty As Boolean
        Private odtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private odtGWS As Date
        Private odtLDR As Date
        Private odtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private odtPTT As Date
        Private odtReportDate As Date
        Private odtCompAssDate As Date
        Private odtStartDate As Date
        Private onEventID As Int64
        Private onEventStatus As Int32
        Private onFacilityID As Int16
        Private onFlagID As Int64
        Private onMGPTFStatus As Int32
        Private onPriority As Int16
        Private onProjMgr As Int32
        Private onReleaseStatus As Int32
        Private onReportSource As Int32
        Private ostrCreatedBy As String = String.Empty
        Private ostrModifiedBy As String = String.Empty
        Private strCreatedBy As String = String.Empty
        Private strModifiedBy As String = String.Empty
        Private strRelatedSites As String = String.Empty
        Private ostrRelatedSites As String = String.Empty
        Private obolForCommission As Boolean
        Private bolForCommission As Boolean

        Private nSuspectedSource As Int32
        'Private nHowDiscoveredID As Int32
        Private dtConfirmed As Date
        Private nIdentifiedBy As Int32
        Private nLocation As Int32
        Private nExtent As Int32
        Private nCause As Int32
        Private nEventPMId As Int32
        Private dtEventStarted As Date
        Private dtEventEnded As Date

        Private bolTOCSOIL As Boolean
        Private bolSOILBTEX As Boolean
        Private bolSOILPAH As Boolean
        Private bolSOILTPH As Boolean
        Private bolTOCGROUNDWATER As Boolean
        Private bolGWBTEX As Boolean
        Private bolGWPAH As Boolean
        Private bolGWTPH As Boolean
        Private bolFREEPRODUCT As Boolean
        Private bolFPGASOLINE As Boolean
        Private bolFPDIESEL As Boolean
        Private bolFPKEROSENE As Boolean
        Private bolFPWASTEOIL As Boolean
        Private bolFPUNKNOWN As Boolean
        Private bolTOCVAPOR As Boolean
        Private bolVAPORBTEX As Boolean
        Private bolVAPORPAH As Boolean
        Private bolHOW_DISC_FAC_LEAK_DETECTION As Boolean
        Private bolHOW_DISC_SURFACE_SHEEN As Boolean
        Private bolHOW_DISC_GW_WELL As Boolean
        Private bolHOW_DISC_GW_CONTAMINATION As Boolean
        Private bolHOW_DISC_VAPORS As Boolean
        Private bolHOW_DISC_FREE_PRODUCT As Boolean
        Private bolHOW_DISC_SOIL_CONTAMINATION As Boolean
        Private bolHOW_DISC_FAILED_PTT As Boolean
        Private bolHOW_DISC_INVENTORY_SHORTAGE As Boolean
        Private bolHOW_DISC_TANK_CLOSURE As Boolean
        Private bolHOW_DISC_INSPECTION As Boolean

        Private onSuspectedSource As Int32
        'Private onHowDiscoveredID As Int32
        Private odtConfirmed As Date
        Private onIdentifiedBy As Int32
        Private onLocation As Int32
        Private onExtent As Int32
        Private onCause As Int32
        Private onEventPMId As Int32
        Private odtEventStarted As Date
        Private odtEventEnded As Date

        Private obolTOCSOIL As Boolean
        Private obolSOILBTEX As Boolean
        Private obolSOILPAH As Boolean
        Private obolSOILTPH As Boolean
        Private obolTOCGROUNDWATER As Boolean
        Private obolGWBTEX As Boolean
        Private obolGWPAH As Boolean
        Private obolGWTPH As Boolean
        Private obolFREEPRODUCT As Boolean
        Private obolFPGASOLINE As Boolean
        Private obolFPDIESEL As Boolean
        Private obolFPKEROSENE As Boolean
        Private obolFPWASTEOIL As Boolean
        Private obolFPUNKNOWN As Boolean
        Private obolTOCVAPOR As Boolean
        Private obolVAPORBTEX As Boolean
        Private obolVAPORPAH As Boolean
        Private obolHOW_DISC_FAC_LEAK_DETECTION As Boolean
        Private obolHOW_DISC_SURFACE_SHEEN As Boolean
        Private obolHOW_DISC_GW_WELL As Boolean
        Private obolHOW_DISC_GW_CONTAMINATION As Boolean
        Private obolHOW_DISC_VAPORS As Boolean
        Private obolHOW_DISC_FREE_PRODUCT As Boolean
        Private obolHOW_DISC_SOIL_CONTAMINATION As Boolean
        Private obolHOW_DISC_FAILED_PTT As Boolean
        Private obolHOW_DISC_INVENTORY_SHORTAGE As Boolean
        Private obolHOW_DISC_TANK_CLOSURE As Boolean
        Private obolHOW_DISC_INSPECTION As Boolean

        Private onEVENTSEQUENCE As Int32
        Private nEVENTSEQUENCE As Int32

        Private strTFCheckList As String
        Private ostrTFCheckList As String
        Private strTankandPipe As String
        Private ostrTankandPipe As String

        Private nPM_HEAD_ASSESS As Int32
        Private dtPM_HEAD_DATE As DateTime
        Private strPM_HEAD_BY As String
        Private nUST_CHIEF_ASSESS As Int32
        Private dtUST_CHIEF_DATE As DateTime
        Private strUST_CHIEF_BY As String
        Private nOPC_HEAD_ASSESS As Int32
        Private dtOPC_HEAD_DATE As DateTime
        Private strOPC_HEAD_BY As String
        Private bolFOR_OPC_HEAD As Boolean
        Private nCOMMISSION_ASSESS As Int32
        Private dtCOMMISSION_DATE As DateTime
        Private strCOMMISSION_BY As String
        Private strELIGIBITY_COMMENTS As String

        Private onPM_HEAD_ASSESS As Int32
        Private odtPM_HEAD_DATE As DateTime
        Private ostrPM_HEAD_BY As String
        Private onUST_CHIEF_ASSESS As Int32
        Private odtUST_CHIEF_DATE As DateTime
        Private ostrUST_CHIEF_BY As String
        Private onOPC_HEAD_ASSESS As Int32
        Private odtOPC_HEAD_DATE As DateTime
        Private ostrOPC_HEAD_BY As String
        Private obolFOR_OPC_HEAD As Boolean
        Private onCOMMISSION_ASSESS As Int32
        Private odtCOMMISSION_DATE As DateTime
        Private ostrCOMMISSION_BY As String
        Private ostrELIGIBITY_COMMENTS As String

        Private strPMDesc As String
        Private strMGPTFStatusDesc As String
        Private strTechnicalStatusDesc As String

        Private onIRAC As Int32
        Private onERAC As Int32
        Private nIRAC As Int32
        Private nERAC As Int32


        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private WithEvents colLustActivities As New MUSTER.Info.LustActivityCollection

        Private onUserID As Integer = 0
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            Init()
            dtDataAge = Now()
        End Sub
        Sub New(ByVal EventID As Int64, _
            ByVal GWS As Date, _
            ByVal LDR As Date, _
            ByVal PTT As Date, _
            ByVal ReportDate As Date, _
            ByVal CompAssDate As Date, _
            ByVal StartDate As Date, _
            ByVal EventStatus As Int32, _
            ByVal FacilityID As Int16, _
            ByVal FlagID As Int64, _
            ByVal MGPTFStatus As Int32, _
            ByVal Priority As Int16, _
            ByVal ProjMgr As Int32, _
            ByVal ReleaseStatus As Int32, _
            ByVal ReportSource As Int32, _
            ByVal RelatedSites As String, _
            ByVal CREATED_BY As String, _
            ByVal CREATE_DATE As String, _
            ByVal LAST_EDITED_BY As String, _
            ByVal DATE_LAST_EDITED As Date, _
            ByVal SuspectedSource As Int32, _
            ByVal Confirmed As Date, _
            ByVal IdentifiedBy As Int32, _
            ByVal Location As Int32, _
            ByVal Extent As Int32, _
            ByVal EventPMId As Int32, _
            ByVal EventStarted As Date, _
            ByVal EventEnded As Date, _
            ByVal TOCSOIL As Boolean, _
            ByVal SOILBTEX As Boolean, _
            ByVal SOILPAH As Boolean, _
            ByVal SOILTPH As Boolean, _
            ByVal TOCGROUNDWATER As Boolean, _
            ByVal GWBTEX As Boolean, _
            ByVal GWPAH As Boolean, _
            ByVal GWTPH As Boolean, _
            ByVal FREEPRODUCT As Boolean, _
            ByVal FPGASOLINE As Boolean, _
            ByVal FPDIESEL As Boolean, _
            ByVal FPKEROSENE As Boolean, _
            ByVal FPWASTEOIL As Boolean, _
            ByVal FPUNKNOWN As Boolean, _
            ByVal TOCVAPOR As Boolean, _
            ByVal VAPORBTEX As Boolean, _
            ByVal VAPORPAH As Boolean, _
            ByVal EVENTSEQUENCE As Int32, _
            ByVal TFChecklist As String, _
            ByVal TankandPipe As String, _
            ByVal PM_HEAD_ASSESS As Int32, _
            ByVal PM_HEAD_DATE As DateTime, _
            ByVal PM_HEAD_BY As String, _
            ByVal UST_CHIEF_ASSESS As Int32, _
            ByVal UST_CHIEF_DATE As DateTime, _
            ByVal UST_CHIEF_BY As String, _
            ByVal OPC_HEAD_ASSESS As Int32, _
            ByVal OPC_HEAD_DATE As DateTime, _
            ByVal OPC_HEAD_BY As String, _
            ByVal FOR_OPC_HEAD As Boolean, _
            ByVal COMMISSION_ASSESS As Int32, _
            ByVal COMMISSION_DATE As DateTime, _
            ByVal COMMISSION_BY As String, _
            ByVal FOR_COMMISSION As Boolean, _
            ByVal ELIGIBITY_COMMENTS As String, _
            ByVal PMDesc As String, _
            ByVal MGPTFStatusDesc As String, _
            ByVal TechStatusDesc As String, _
            ByVal IRAC As Int32, _
            ByVal ERAC As Int32, _
            ByVal HowDiscFacLD As Boolean, _
            ByVal HowDiscSurfaceSheen As Boolean, _
            ByVal HowDiscGWWell As Boolean, _
            ByVal HowDiscGWContamination As Boolean, _
            ByVal HowDiscVapors As Boolean, _
            ByVal HowDiscFreeProduct As Boolean, _
            ByVal HowDiscSoilContamination As Boolean, _
            ByVal HowDiscFailedPTT As Boolean, _
            ByVal HowDiscInventoryShortage As Boolean, _
            ByVal HowDiscTankClosure As Boolean, _
            ByVal HowDiscInspection As Boolean, _
            ByVal Cause As Int32)
            'onPropID = PROPERTY_ID
            'ByVal HowDiscoveredID As Int32, _

            onEventID = EventID
            odtGWS = GWS
            odtLDR = LDR
            odtPTT = PTT
            odtReportDate = ReportDate
            odtCompAssDate = CompAssDate
            odtStartDate = StartDate
            onEventStatus = EventStatus
            onFacilityID = FacilityID
            onFlagID = FlagID
            onMGPTFStatus = MGPTFStatus
            onPriority = Priority
            onProjMgr = ProjMgr
            onReleaseStatus = ReleaseStatus
            onReportSource = ReportSource
            ostrRelatedSites = RelatedSites

            onSuspectedSource = SuspectedSource
            'onHowDiscoveredID = HowDiscoveredID
            odtConfirmed = Confirmed
            onIdentifiedBy = IdentifiedBy
            onLocation = Location
            onCause = Cause
            onExtent = Extent
            onEventPMId = EventPMId
            odtEventStarted = EventStarted
            odtEventEnded = EventEnded

            obolTOCSOIL = TOCSOIL
            obolSOILBTEX = SOILBTEX
            obolSOILPAH = SOILPAH
            obolSOILTPH = SOILTPH
            obolTOCGROUNDWATER = TOCGROUNDWATER
            obolGWBTEX = GWBTEX
            obolGWPAH = GWPAH
            obolGWTPH = GWTPH
            obolFREEPRODUCT = FREEPRODUCT
            obolFPGASOLINE = FPGASOLINE
            obolFPDIESEL = FPDIESEL
            obolFPKEROSENE = FPKEROSENE
            obolFPWASTEOIL = FPWASTEOIL
            obolFPUNKNOWN = FPUNKNOWN
            obolTOCVAPOR = TOCVAPOR
            obolVAPORBTEX = VAPORBTEX
            obolVAPORPAH = VAPORPAH
            obolForCommission = bolForCommission

            onEVENTSEQUENCE = EVENTSEQUENCE

            ostrTFCheckList = TFChecklist
            ostrTankandPipe = TankandPipe

            onPM_HEAD_ASSESS = PM_HEAD_ASSESS
            odtPM_HEAD_DATE = PM_HEAD_DATE
            ostrPM_HEAD_BY = PM_HEAD_BY
            onUST_CHIEF_ASSESS = UST_CHIEF_ASSESS
            odtUST_CHIEF_DATE = UST_CHIEF_DATE
            ostrUST_CHIEF_BY = UST_CHIEF_BY
            onOPC_HEAD_ASSESS = OPC_HEAD_ASSESS
            odtOPC_HEAD_DATE = OPC_HEAD_DATE
            ostrOPC_HEAD_BY = OPC_HEAD_BY
            obolFOR_OPC_HEAD = FOR_OPC_HEAD
            onCOMMISSION_ASSESS = COMMISSION_ASSESS
            odtCOMMISSION_DATE = COMMISSION_DATE
            ostrCOMMISSION_BY = COMMISSION_BY
            ostrELIGIBITY_COMMENTS = ELIGIBITY_COMMENTS

            ostrCreatedBy = CREATED_BY
            odtCreatedOn = CREATE_DATE
            ostrModifiedBy = LAST_EDITED_BY
            odtModifiedOn = DATE_LAST_EDITED


            strPMDesc = PMDesc
            strMGPTFStatusDesc = MGPTFStatusDesc
            strTechnicalStatusDesc = TechStatusDesc

            onIRAC = IRAC
            onERAC = ERAC

            obolHOW_DISC_FAC_LEAK_DETECTION = HowDiscFacLD
            obolHOW_DISC_SURFACE_SHEEN = HowDiscSurfaceSheen
            obolHOW_DISC_GW_WELL = HowDiscGWWell
            obolHOW_DISC_GW_CONTAMINATION = HowDiscGWContamination
            obolHOW_DISC_VAPORS = HowDiscVapors
            obolHOW_DISC_FREE_PRODUCT = HowDiscFreeProduct
            obolHOW_DISC_SOIL_CONTAMINATION = HowDiscSoilContamination
            obolHOW_DISC_FAILED_PTT = HowDiscFailedPTT
            obolHOW_DISC_INVENTORY_SHORTAGE = HowDiscInventoryShortage
            obolHOW_DISC_TANK_CLOSURE = HowDiscTankClosure
            obolHOW_DISC_INSPECTION = HowDiscInspection

            dtDataAge = Now()
            Me.Reset()

        End Sub

        Private Sub New(ByVal oLustEvent As MUSTER.Info.LustEventInfo)
            Try
                'onPropID = oProp.ID
                onEventID = oLustEvent.ID
                odtGWS = oLustEvent.LastGWS
                odtLDR = oLustEvent.LastLDR
                odtPTT = oLustEvent.LastPTT
                odtReportDate = oLustEvent.ReportDate
                odtCompAssDate = oLustEvent.CompAssDate
                odtStartDate = oLustEvent.Started
                onEventStatus = oLustEvent.EventStatus
                onFacilityID = oLustEvent.FacilityID
                onFlagID = oLustEvent.FlagID
                onMGPTFStatus = oLustEvent.MGPTFStatus
                onPriority = oLustEvent.Priority
                onProjMgr = oLustEvent.PM
                onReleaseStatus = oLustEvent.ReleaseStatus
                onReportSource = oLustEvent.ReportSource
                ostrRelatedSites = oLustEvent.RelatedSites
                onSuspectedSource = oLustEvent.SuspectedSource
                'onHowDiscoveredID = oLustEvent.HowDiscoveredID
                odtConfirmed = oLustEvent.Confirmed
                onIdentifiedBy = oLustEvent.IDENTIFIEDBY
                onLocation = oLustEvent.Location
                onExtent = oLustEvent.Extent
                onCause = oLustEvent.Cause
                onEventPMId = oLustEvent.EventPMId
                odtEventStarted = oLustEvent.EventStarted
                odtEventEnded = oLustEvent.EventEnded

                obolTOCSOIL = oLustEvent.TOCSOIL
                obolSOILBTEX = oLustEvent.SOILBTEX
                obolSOILPAH = oLustEvent.SOILPAH
                obolSOILTPH = oLustEvent.SOILTPH
                obolTOCGROUNDWATER = oLustEvent.TOCGROUNDWATER
                obolGWBTEX = oLustEvent.GWBTEX
                obolGWPAH = oLustEvent.GWPAH
                obolGWTPH = oLustEvent.GWTPH
                obolFREEPRODUCT = oLustEvent.FREEPRODUCT
                obolFPGASOLINE = oLustEvent.FPGASOLINE
                obolFPDIESEL = oLustEvent.FPDIESEL
                obolFPKEROSENE = oLustEvent.FPKEROSENE
                obolFPWASTEOIL = oLustEvent.FPWASTEOIL
                obolFPUNKNOWN = oLustEvent.FPUNKNOWN
                obolTOCVAPOR = oLustEvent.TOCVAPOR
                obolVAPORBTEX = oLustEvent.VAPORBTEX
                obolVAPORPAH = oLustEvent.VAPORPAH
                obolForCommission = oLustEvent.FOR_COMMISSION

                onEVENTSEQUENCE = oLustEvent.EVENTSEQUENCE
                ostrCreatedBy = oLustEvent.CreatedBy
                odtCreatedOn = oLustEvent.CreatedOn
                ostrModifiedBy = oLustEvent.ModifiedBy
                odtModifiedOn = oLustEvent.ModifiedOn

                ostrTFCheckList = oLustEvent.TFCheckList
                ostrTankandPipe = oLustEvent.TankandPipe

                onPM_HEAD_ASSESS = oLustEvent.PM_HEAD_ASSESS
                odtPM_HEAD_DATE = oLustEvent.PM_HEAD_DATE
                ostrPM_HEAD_BY = oLustEvent.PM_HEAD_BY
                onUST_CHIEF_ASSESS = oLustEvent.UST_CHIEF_ASSESS
                odtUST_CHIEF_DATE = oLustEvent.UST_CHIEF_DATE
                ostrUST_CHIEF_BY = oLustEvent.UST_CHIEF_BY
                onOPC_HEAD_ASSESS = oLustEvent.OPC_HEAD_ASSESS
                odtOPC_HEAD_DATE = oLustEvent.OPC_HEAD_DATE
                ostrOPC_HEAD_BY = oLustEvent.OPC_HEAD_BY
                obolFOR_OPC_HEAD = oLustEvent.FOR_OPC_HEAD
                onCOMMISSION_ASSESS = oLustEvent.COMMISSION_ASSESS
                odtCOMMISSION_DATE = oLustEvent.COMMISSION_DATE
                ostrCOMMISSION_BY = oLustEvent.COMMISSION_BY
                ostrELIGIBITY_COMMENTS = oLustEvent.ELIGIBITY_COMMENTS

                strPMDesc = oLustEvent.PMDesc
                strMGPTFStatusDesc = oLustEvent.MGPTFStatusDesc
                strTechnicalStatusDesc = oLustEvent.TechnicalStatusDesc

                nIRAC = oLustEvent.IRAC
                nERAC = oLustEvent.ERAC

                obolHOW_DISC_FAC_LEAK_DETECTION = oLustEvent.HowDiscFacLD
                obolHOW_DISC_SURFACE_SHEEN = oLustEvent.HowDiscSurfaceSheen
                obolHOW_DISC_GW_WELL = oLustEvent.HowDiscGWWell
                obolHOW_DISC_GW_CONTAMINATION = oLustEvent.HowDiscGWContamination
                obolHOW_DISC_VAPORS = oLustEvent.HowDiscVapors
                obolHOW_DISC_FREE_PRODUCT = oLustEvent.HowDiscFreeProduct
                obolHOW_DISC_SOIL_CONTAMINATION = oLustEvent.HowDiscSoilContamination
                obolHOW_DISC_FAILED_PTT = oLustEvent.HowDiscFailedPTT
                obolHOW_DISC_INVENTORY_SHORTAGE = oLustEvent.HowDiscInventoryShortage
                obolHOW_DISC_TANK_CLOSURE = oLustEvent.HowDiscTankClosure
                obolHOW_DISC_INSPECTION = oLustEvent.HowDiscInspection

                dtDataAge = Now()
                Me.Reset()
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
#End Region
#Region "Exposed Attributes"

        Public ReadOnly Property PMDesc() As String
            Get
                Return strPMDesc
            End Get
        End Property
        Public ReadOnly Property MGPTFStatusDesc() As String
            Get
                Return strMGPTFStatusDesc
            End Get
        End Property
        Public ReadOnly Property TechnicalStatusDesc() As String
            Get
                Return strTechnicalStatusDesc
            End Get
        End Property

        ' The maximum age the info object can attain before requiring a refresh
        Public Property AgeThreshold() As Date
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{4A717F21-3341-4208-A483-9A8BACA51984}
                Return dtDataAge
                ' #End Region ' XDEOperation End Template Expansion{4A717F21-3341-4208-A483-9A8BACA51984}
            End Get
            Set(ByVal Value As Date)
                ' #Region "XDEOperation" ' Begin Template Expansion{81C833E2-220F-421E-9A7A-3B2CC3B9FB39}
                dtDataAge = Value
                ' #End Region ' XDEOperation End Template Expansion{81C833E2-220F-421E-9A7A-3B2CC3B9FB39}
            End Set
        End Property
        ' The ID of the user that created the row
        Public Property CreatedBy() As String
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{3C051DD3-2D74-4F71-9C3C-30F06ACC7E78}
                Return strCreatedBy
                ' #End Region ' XDEOperation End Template Expansion{3C051DD3-2D74-4F71-9C3C-30F06ACC7E78}
            End Get
            Set(ByVal Value As String)
                strCreatedBy = Value
            End Set
        End Property
        Public Property CreatedOn() As Date
            Get
                Return dtCreatedOn
            End Get
            Set(ByVal Value As Date)
                dtCreatedOn = Value
            End Set
        End Property
        '' The date on which the row was created
        'Public ReadOnly Property CreatedOn() As Date
        '    Get
        '        ' #Region "XDEOperation" ' Begin Template Expansion{F0EC3ABE-53C5-4ECB-9ED0-7AACB21427B7}
        '        Return dtCreatedOn
        '        ' #End Region ' XDEOperation End Template Expansion{F0EC3ABE-53C5-4ECB-9ED0-7AACB21427B7}
        '    End Get
        'End Property
        ' Indicates the deleted state of the row
        Public Property Deleted() As Boolean
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{62C7C52B-FE4E-49F3-AB97-7A51ADFAB009}
                Return bolDeleted
                ' #End Region ' XDEOperation End Template Expansion{62C7C52B-FE4E-49F3-AB97-7A51ADFAB009}
            End Get
            Set(ByVal Value As Boolean)
                ' #Region "XDEOperation" ' Begin Template Expansion{F45D08F9-6F85-4F50-8CC1-7F38CED752CF}
                bolDeleted = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{F45D08F9-6F85-4F50-8CC1-7F38CED752CF}
            End Set
        End Property
        ' The current status of the LUST event - derived from properties table
        Public Property EventStatus() As Integer
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{B572E38D-5F27-4680-899C-46026F08BEB4}
                Return nEventStatus
                ' #End Region ' XDEOperation End Template Expansion{B572E38D-5F27-4680-899C-46026F08BEB4}
            End Get
            Set(ByVal Value As Integer)
                ' #Region "XDEOperation" ' Begin Template Expansion{25A4ACC8-3E81-4F5C-93E7-5164AC540E57}
                nEventStatus = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{25A4ACC8-3E81-4F5C-93E7-5164AC540E57}
            End Set
        End Property
        Public ReadOnly Property EventStatusOriginal() As Integer
            Get
                Return onEventStatus
            End Get
        End Property
        ' The LUST event ID for the facility
        Public Property FacilityID() As Integer
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{1BED3A4C-3268-4D81-BE82-E886FA3D8420}
                Return nFacilityID
                ' #End Region ' XDEOperation End Template Expansion{1BED3A4C-3268-4D81-BE82-E886FA3D8420}
            End Get
            Set(ByVal Value As Integer)
                nFacilityID = Value
            End Set
        End Property
        ' The system ID of the flag associated with the LUST event
        Public Property FlagID() As Long
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{A403CAF7-2EAC-4827-90F2-5E2F86E929C6}
                Return nFlagID
                ' #End Region ' XDEOperation End Template Expansion{A403CAF7-2EAC-4827-90F2-5E2F86E929C6}
            End Get
            Set(ByVal Value As Long)
                ' #Region "XDEOperation" ' Begin Template Expansion{A627C621-57A8-4626-B7A5-B0B13CA6F3B8}
                nFlagID = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{A627C621-57A8-4626-B7A5-B0B13CA6F3B8}
            End Set
        End Property
        ' The system ID for this LUST event
        Public Property ID() As Long
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{A0B81136-12BB-4719-8E07-49B06333F62F}
                Return nEventID
                ' #End Region ' XDEOperation End Template Expansion{A0B81136-12BB-4719-8E07-49B06333F62F}
            End Get
            Set(ByVal Value As Long)
                nEventID = Value
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return bolIsDirty
            End Get
            Set(ByVal value As Boolean)
                bolIsDirty = value
            End Set
        End Property

        Public Property IsEligibityDirty() As Boolean
            Get
                Return bolIsEligibityDirty
            End Get
            Set(ByVal value As Boolean)
                bolIsEligibityDirty = value
            End Set
        End Property

        Public ReadOnly Property IsNewRelease() As Boolean
            Get
                If (obolHOW_DISC_FAC_LEAK_DETECTION = False And _
                    obolHOW_DISC_SURFACE_SHEEN = False And _
                    obolHOW_DISC_GW_WELL = False And _
                    obolHOW_DISC_GW_CONTAMINATION = False And _
                    obolHOW_DISC_VAPORS = False And _
                    obolHOW_DISC_FREE_PRODUCT = False And _
                    obolHOW_DISC_SOIL_CONTAMINATION = False And _
                    obolHOW_DISC_FAILED_PTT = False And _
                    obolHOW_DISC_INVENTORY_SHORTAGE = False And _
                    obolHOW_DISC_TANK_CLOSURE = False And _
                    obolHOW_DISC_INSPECTION = False And _
                    odtConfirmed = "01/01/0001" And onIdentifiedBy = 0 And onLocation = 0 And onExtent = 0) And _
                    (obolHOW_DISC_FAC_LEAK_DETECTION = True Or _
                    obolHOW_DISC_SURFACE_SHEEN = True Or _
                    obolHOW_DISC_GW_WELL = True Or _
                    obolHOW_DISC_GW_CONTAMINATION = True Or _
                    obolHOW_DISC_VAPORS = True Or _
                    obolHOW_DISC_FREE_PRODUCT = True Or _
                    obolHOW_DISC_SOIL_CONTAMINATION = True Or _
                    obolHOW_DISC_FAILED_PTT = True Or _
                    obolHOW_DISC_INVENTORY_SHORTAGE = True Or _
                    obolHOW_DISC_TANK_CLOSURE = True Or _
                    obolHOW_DISC_INSPECTION = True Or _
                    dtConfirmed <> "01/01/0001" Or nIdentifiedBy <> 0 Or onLocation <> 0 Or nExtent <> 0) Then
                    Return True
                Else
                    Return False
                End If
                'If (onHowDiscoveredID = 0 And odtConfirmed = "01/01/0001" And onIdentifiedBy = 0 And onLocation = 0 And onExtent = 0) _
                'And (nHowDiscoveredID <> 0 Or dtConfirmed <> "01/01/0001" Or nIdentifiedBy <> 0 Or onLocation <> 0 Or nExtent <> 0) Then
                '    Return True
                'Else
                '    Return False
                'End If
            End Get
        End Property
        ' Date of last GWS for LUST event
        Public Property LastGWS() As Date
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{C4562BA4-A8CC-4382-9386-816629748037}
                Return dtGWS
                ' #End Region ' XDEOperation End Template Expansion{C4562BA4-A8CC-4382-9386-816629748037}
            End Get
            Set(ByVal Value As Date)
                ' #Region "XDEOperation" ' Begin Template Expansion{F0257B29-86E7-4152-9292-26ABA2367BE5}
                dtGWS = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{F0257B29-86E7-4152-9292-26ABA2367BE5}
            End Set
        End Property
        ' The date of the last LDR for the LUST event
        Public Property LastLDR() As Date
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{B44ADE8F-D96A-43A7-B55F-DB87049DBC50}
                Return dtLDR
                ' #End Region ' XDEOperation End Template Expansion{B44ADE8F-D96A-43A7-B55F-DB87049DBC50}
            End Get
            Set(ByVal Value As Date)
                ' #Region "XDEOperation" ' Begin Template Expansion{F78270BF-83A9-4DE2-88F0-B34A246A0127}
                dtLDR = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{F78270BF-83A9-4DE2-88F0-B34A246A0127}
            End Set
        End Property
        ' The date of the last PTT for the LUST event
        Public Property LastPTT() As Date
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{BEFBDBD6-0D3F-4520-AD97-AFCAAB5A56C8}
                Return dtPTT
                ' #End Region ' XDEOperation End Template Expansion{BEFBDBD6-0D3F-4520-AD97-AFCAAB5A56C8}
            End Get
            Set(ByVal Value As Date)
                ' #Region "XDEOperation" ' Begin Template Expansion{93C6E0C4-ABDB-4B5F-99A1-826A9DACC251}
                dtPTT = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{93C6E0C4-ABDB-4B5F-99A1-826A9DACC251}
            End Set
        End Property
        ' The MGPTF status for the LUST event - derived from properties table
        Public Property MGPTFStatus() As Integer
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{2451FC98-9B48-4F0D-8258-5BFFDCDCDB3B}
                Return nMGPTFStatus
                ' #End Region ' XDEOperation End Template Expansion{2451FC98-9B48-4F0D-8258-5BFFDCDCDB3B}
            End Get
            Set(ByVal Value As Integer)
                ' #Region "XDEOperation" ' Begin Template Expansion{AAB754B6-F493-4041-9D88-E6180C328C87}
                nMGPTFStatus = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{AAB754B6-F493-4041-9D88-E6180C328C87}
            End Set
        End Property
        Public Property ModifiedBy() As String
            Get
                Return Me.strModifiedBy
            End Get
            Set(ByVal Value As String)
                strModifiedBy = Value
            End Set
        End Property
        'Public ReadOnly Property ModifiedOn() As Date
        '    Get
        '        Return Me.dtModifiedOn
        '    End Get
        'End Property
        Public Property ModifiedOn() As Date
            Get
                Return dtModifiedOn
            End Get
            Set(ByVal Value As Date)
                dtModifiedOn = Value
            End Set
        End Property
        ' The current Project Manager for the LUST event
        Public Property PM() As Integer
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{8FBF2668-B513-4F85-BAC7-68E21D92EEDA}
                Return nProjMgr
                ' #End Region ' XDEOperation End Template Expansion{8FBF2668-B513-4F85-BAC7-68E21D92EEDA}
            End Get
            Set(ByVal Value As Integer)
                ' #Region "XDEOperation" ' Begin Template Expansion{4F8C7845-5530-41D5-A524-DA521AA28643}
                nProjMgr = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{4F8C7845-5530-41D5-A524-DA521AA28643}
            End Set
        End Property
        ' The current priority for the LUST event - can only range 1 to 8
        Public Property Priority() As Short
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{01A8A167-3D73-4D1E-A56A-2CD52DA7C1A4}
                Return nPriority
                ' #End Region ' XDEOperation End Template Expansion{01A8A167-3D73-4D1E-A56A-2CD52DA7C1A4}
            End Get
            Set(ByVal Value As Short)
                ' #Region "XDEOperation" ' Begin Template Expansion{4CA5EB6D-DD6B-48A3-A6BB-E92DF7D7164B}
                nPriority = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{4CA5EB6D-DD6B-48A3-A6BB-E92DF7D7164B}
            End Set
        End Property
        ' The current release status for the LUST event - derived from properties table
        Public Property ReleaseStatus() As Integer
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{D20FA9D0-1A82-48D7-9FC4-1667DD31120C}
                Return nReleaseStatus
                ' #End Region ' XDEOperation End Template Expansion{D20FA9D0-1A82-48D7-9FC4-1667DD31120C}
            End Get
            Set(ByVal Value As Integer)
                ' #Region "XDEOperation" ' Begin Template Expansion{650D7738-2491-4B6F-8FFB-BFA8A0B51309}
                nReleaseStatus = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{650D7738-2491-4B6F-8FFB-BFA8A0B51309}
            End Set
        End Property
        ' The date the LUST event was reported to UST
        Public Property ReportDate() As Date
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{DF329ACC-545C-49B9-B906-248378C2381A}
                Return dtReportDate
                ' #End Region ' XDEOperation End Template Expansion{DF329ACC-545C-49B9-B906-248378C2381A}
            End Get
            Set(ByVal Value As Date)
                ' #Region "XDEOperation" ' Begin Template Expansion{0FBC95DC-82D9-4C7B-B840-03AA36D6C9AB}
                dtReportDate = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{0FBC95DC-82D9-4C7B-B840-03AA36D6C9AB}
            End Set
        End Property
        Public Property CompAssDate() As Date
            Get
                Return dtCompAssDate
            End Get
            Set(ByVal Value As Date)
                dtCompAssDate = Value
                Me.CheckDirty()
            End Set
        End Property
        ' The source of the LUST event report to UST - derived from properties table
        Public Property ReportSource() As Integer
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{F447E560-B40E-4018-8A68-2EC8A0F949A9}
                Return nReportSource
                ' #End Region ' XDEOperation End Template Expansion{F447E560-B40E-4018-8A68-2EC8A0F949A9}
            End Get
            Set(ByVal Value As Integer)
                ' #Region "XDEOperation" ' Begin Template Expansion{015269C4-2C21-4597-9958-203158463AED}
                nReportSource = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{015269C4-2C21-4597-9958-203158463AED}
            End Set
        End Property
        ' The date the LUST event was started
        Public Property Started() As Date
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{DACD7EB3-FC1D-467E-AB2F-253698F33290}
                Return dtStartDate
                ' #End Region ' XDEOperation End Template Expansion{DACD7EB3-FC1D-467E-AB2F-253698F33290}
            End Get
            Set(ByVal Value As Date)
                ' #Region "XDEOperation" ' Begin Template Expansion{F4EC71D3-7F82-4E43-9C1E-711F85F65C03}
                dtStartDate = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{F4EC71D3-7F82-4E43-9C1E-711F85F65C03}
            End Set
        End Property

        Public Property SuspectedSource() As Integer
            Get
                Return nSuspectedSource
            End Get
            Set(ByVal Value As Integer)
                nSuspectedSource = Value
                Me.CheckDirty()
            End Set
        End Property

        'Public Property HowDiscoveredID() As Integer
        '    Get
        '        Return nHowDiscoveredID
        '    End Get
        '    Set(ByVal Value As Integer)
        '        nHowDiscoveredID = Value
        '        Me.CheckDirty()
        '    End Set
        'End Property
        Public Property Confirmed() As Date
            Get
                Return dtConfirmed
            End Get
            Set(ByVal Value As Date)
                dtConfirmed = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property IDENTIFIEDBY() As Integer
            Get
                Return nIdentifiedBy
            End Get
            Set(ByVal Value As Integer)
                nIdentifiedBy = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property Location() As Integer
            Get
                Return nLocation
            End Get
            Set(ByVal Value As Integer)
                nLocation = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property Extent() As Integer
            Get
                Return nExtent
            End Get
            Set(ByVal Value As Integer)
                nExtent = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property Cause() As Integer
            Get
                Return nCause
            End Get
            Set(ByVal Value As Integer)
                nCause = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property EventPMId() As Integer
            Get
                Return nEventPMId
            End Get
            Set(ByVal Value As Integer)
                nEventPMId = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property EventStarted() As Date
            Get
                Return dtEventStarted
            End Get
            Set(ByVal Value As Date)
                dtEventStarted = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property EventEnded() As Date
            Get
                Return dtEventEnded
            End Get
            Set(ByVal Value As Date)
                dtEventEnded = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property RelatedSites() As String
            Get
                Return strRelatedSites
            End Get
            Set(ByVal Value As String)
                strRelatedSites = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property TOCSOIL() As Boolean
            Get
                Return bolTOCSOIL
            End Get
            Set(ByVal Value As Boolean)
                bolTOCSOIL = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property SOILBTEX() As Boolean
            Get
                Return bolSOILBTEX
            End Get
            Set(ByVal Value As Boolean)
                bolSOILBTEX = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property SOILPAH() As Boolean
            Get
                Return bolSOILPAH
            End Get
            Set(ByVal Value As Boolean)
                bolSOILPAH = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property SOILTPH() As Boolean
            Get
                Return bolSOILTPH
            End Get
            Set(ByVal Value As Boolean)
                bolSOILTPH = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property TOCGROUNDWATER() As Boolean
            Get
                Return bolTOCGROUNDWATER
            End Get
            Set(ByVal Value As Boolean)
                bolTOCGROUNDWATER = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property GWBTEX() As Boolean
            Get
                Return bolGWBTEX
            End Get
            Set(ByVal Value As Boolean)
                bolGWBTEX = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property GWPAH() As Boolean
            Get
                Return bolGWPAH
            End Get
            Set(ByVal Value As Boolean)
                bolGWPAH = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property GWTPH() As Boolean
            Get
                Return bolGWTPH
            End Get
            Set(ByVal Value As Boolean)
                bolGWTPH = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property FREEPRODUCT() As Boolean
            Get
                Return bolFREEPRODUCT
            End Get
            Set(ByVal Value As Boolean)
                bolFREEPRODUCT = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property FPGASOLINE() As Boolean
            Get
                Return bolFPGASOLINE
            End Get
            Set(ByVal Value As Boolean)
                bolFPGASOLINE = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property FPDIESEL() As Boolean
            Get
                Return bolFPDIESEL
            End Get
            Set(ByVal Value As Boolean)
                bolFPDIESEL = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property FPKEROSENE() As Boolean
            Get
                Return bolFPKEROSENE
            End Get
            Set(ByVal Value As Boolean)
                bolFPKEROSENE = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property FPWASTEOIL() As Boolean
            Get
                Return bolFPWASTEOIL
            End Get
            Set(ByVal Value As Boolean)
                bolFPWASTEOIL = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property FPUNKNOWN() As Boolean
            Get
                Return bolFPUNKNOWN
            End Get
            Set(ByVal Value As Boolean)
                bolFPUNKNOWN = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property TOCVAPOR() As Boolean
            Get
                Return bolTOCVAPOR
            End Get
            Set(ByVal Value As Boolean)
                bolTOCVAPOR = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property VAPORBTEX() As Boolean
            Get
                Return bolVAPORBTEX
            End Get
            Set(ByVal Value As Boolean)
                bolVAPORBTEX = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property VAPORPAH() As Boolean
            Get
                Return bolVAPORPAH
            End Get
            Set(ByVal Value As Boolean)
                bolVAPORPAH = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property EVENTSEQUENCE() As Integer
            Get
                Return nEVENTSEQUENCE
            End Get
            Set(ByVal Value As Integer)
                nEVENTSEQUENCE = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property TFCheckList() As String
            Get
                Return strTFCheckList
            End Get
            Set(ByVal Value As String)
                strTFCheckList = Value

                If strTFCheckList.Length > 0 AndAlso (ostrTFCheckList Is Nothing OrElse ostrTFCheckList.Length = 0) Then
                    ostrTFCheckList = Value
                End If

                Me.CheckDirty()
            End Set
        End Property

        Public Property TankandPipe() As String
            Get
                Return strTankandPipe
            End Get
            Set(ByVal Value As String)
                strTankandPipe = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property PM_HEAD_ASSESS() As Int32
            Get
                Return nPM_HEAD_ASSESS
            End Get
            Set(ByVal Value As Int32)
                nPM_HEAD_ASSESS = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property PM_HEAD_DATE() As Date
            Get
                Return dtPM_HEAD_DATE
            End Get
            Set(ByVal Value As Date)
                dtPM_HEAD_DATE = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property PM_HEAD_BY() As String
            Get
                Return strPM_HEAD_BY
            End Get
            Set(ByVal Value As String)
                strPM_HEAD_BY = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property UST_CHIEF_ASSESS() As Int32
            Get
                Return nUST_CHIEF_ASSESS
            End Get
            Set(ByVal Value As Int32)
                nUST_CHIEF_ASSESS = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property UST_CHIEF_DATE() As Date
            Get
                Return dtUST_CHIEF_DATE
            End Get
            Set(ByVal Value As Date)
                dtUST_CHIEF_DATE = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property UST_CHIEF_BY() As String
            Get
                Return strUST_CHIEF_BY
            End Get
            Set(ByVal Value As String)
                strUST_CHIEF_BY = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property OPC_HEAD_ASSESS() As Int32
            Get
                Return nOPC_HEAD_ASSESS
            End Get
            Set(ByVal Value As Int32)
                nOPC_HEAD_ASSESS = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property OPC_HEAD_DATE() As Date
            Get
                Return dtOPC_HEAD_DATE
            End Get
            Set(ByVal Value As Date)
                dtOPC_HEAD_DATE = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property OPC_HEAD_BY() As String
            Get
                Return strOPC_HEAD_BY
            End Get
            Set(ByVal Value As String)
                strOPC_HEAD_BY = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property FOR_OPC_HEAD() As Boolean
            Get
                Return bolFOR_OPC_HEAD
            End Get
            Set(ByVal Value As Boolean)
                bolFOR_OPC_HEAD = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property COMMISSION_ASSESS() As Int32
            Get
                Return nCOMMISSION_ASSESS
            End Get
            Set(ByVal Value As Int32)
                nCOMMISSION_ASSESS = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property COMMISSION_DATE() As Date
            Get
                Return dtCOMMISSION_DATE
            End Get
            Set(ByVal Value As Date)
                dtCOMMISSION_DATE = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property COMMISSION_BY() As String
            Get
                Return strCOMMISSION_BY
            End Get
            Set(ByVal Value As String)
                strCOMMISSION_BY = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property ELIGIBITY_COMMENTS() As String
            Get
                Return strELIGIBITY_COMMENTS
            End Get
            Set(ByVal Value As String)
                strELIGIBITY_COMMENTS = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property FOR_COMMISSION() As Boolean
            Get
                Return bolForCommission
            End Get
            Set(ByVal Value As Boolean)
                bolForCommission = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property Activities() As MUSTER.Info.LustActivityCollection
            Get
                Return colLustActivities
            End Get
            Set(ByVal Value As MUSTER.Info.LustActivityCollection)
                colLustActivities = Value
            End Set
        End Property

        Public Property UserID() As Integer
            Get
                Return onUserID
            End Get
            Set(ByVal Value As Integer)
                onUserID = Value
            End Set
        End Property
        Public Property IRAC() As Integer
            Get
                Return nIRAC
            End Get
            Set(ByVal Value As Integer)
                nIRAC = Value
            End Set
        End Property
        Public Property ERAC() As Integer
            Get
                Return nERAC
            End Get
            Set(ByVal Value As Integer)
                nERAC = Value
            End Set
        End Property

        Public Property HowDiscFacLD() As Boolean
            Get
                Return bolHOW_DISC_FAC_LEAK_DETECTION
            End Get
            Set(ByVal Value As Boolean)
                bolHOW_DISC_FAC_LEAK_DETECTION = Value
                CheckDirty()
            End Set
        End Property
        Public Property HowDiscSurfaceSheen() As Boolean
            Get
                Return bolHOW_DISC_SURFACE_SHEEN
            End Get
            Set(ByVal Value As Boolean)
                bolHOW_DISC_SURFACE_SHEEN = Value
                CheckDirty()
            End Set
        End Property
        Public Property HowDiscGWWell() As Boolean
            Get
                Return bolHOW_DISC_GW_WELL
            End Get
            Set(ByVal Value As Boolean)
                bolHOW_DISC_GW_WELL = Value
                CheckDirty()
            End Set
        End Property
        Public Property HowDiscGWContamination() As Boolean
            Get
                Return bolHOW_DISC_GW_CONTAMINATION
            End Get
            Set(ByVal Value As Boolean)
                bolHOW_DISC_GW_CONTAMINATION = Value
                CheckDirty()
            End Set
        End Property
        Public Property HowDiscVapors() As Boolean
            Get
                Return bolHOW_DISC_VAPORS
            End Get
            Set(ByVal Value As Boolean)
                bolHOW_DISC_VAPORS = Value
                CheckDirty()
            End Set
        End Property
        Public Property HowDiscFreeProduct() As Boolean
            Get
                Return bolHOW_DISC_FREE_PRODUCT
            End Get
            Set(ByVal Value As Boolean)
                bolHOW_DISC_FREE_PRODUCT = Value
                CheckDirty()
            End Set
        End Property
        Public Property HowDiscSoilContamination() As Boolean
            Get
                Return bolHOW_DISC_SOIL_CONTAMINATION
            End Get
            Set(ByVal Value As Boolean)
                bolHOW_DISC_SOIL_CONTAMINATION = Value
                CheckDirty()
            End Set
        End Property
        Public Property HowDiscFailedPTT() As Boolean
            Get
                Return bolHOW_DISC_FAILED_PTT
            End Get
            Set(ByVal Value As Boolean)
                bolHOW_DISC_FAILED_PTT = Value
                CheckDirty()
            End Set
        End Property
        Public Property HowDiscInventoryShortage() As Boolean
            Get
                Return bolHOW_DISC_INVENTORY_SHORTAGE
            End Get
            Set(ByVal Value As Boolean)
                bolHOW_DISC_INVENTORY_SHORTAGE = Value
                CheckDirty()
            End Set
        End Property
        Public Property HowDiscTankClosure() As Boolean
            Get
                Return bolHOW_DISC_TANK_CLOSURE
            End Get
            Set(ByVal Value As Boolean)
                bolHOW_DISC_TANK_CLOSURE = Value
                CheckDirty()
            End Set
        End Property
        Public Property HowDiscInspection() As Boolean
            Get
                Return bolHOW_DISC_INSPECTION
            End Get
            Set(ByVal Value As Boolean)
                bolHOW_DISC_INSPECTION = Value
                CheckDirty()
            End Set
        End Property

#End Region
#Region "Protected Attributes"
        ' Returns a boolean indicating if the data has aged beyond its preset limit
        Protected ReadOnly Property IsAgedData() As Boolean
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{07B7B8F4-0B24-4EA0-B546-4C0D4AB74F19}
                Return dtDataAge < AgeThreshold
                ' #End Region ' XDEOperation End Template Expansion{07B7B8F4-0B24-4EA0-B546-4C0D4AB74F19}
            End Get
        End Property
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()


            bolIsDirty = (nEventID <> onEventID) Or _
                            (dtGWS <> odtGWS) Or _
                            (dtLDR <> odtLDR) Or _
                            (dtPTT <> odtPTT) Or _
                            (dtReportDate <> odtReportDate) Or _
                            (dtCompAssDate <> odtCompAssDate) Or _
                            (dtStartDate <> odtStartDate) Or _
                            (nEventStatus <> onEventStatus) Or _
                            (nFacilityID <> onFacilityID) Or _
                            (nFlagID <> onFlagID) Or _
                            (nMGPTFStatus <> onMGPTFStatus) Or _
                            (nPriority <> onPriority) Or _
                            (nProjMgr <> onProjMgr) Or _
                            (nReleaseStatus <> onReleaseStatus) Or _
                            (nReportSource <> onReportSource) Or _
                            (nSuspectedSource <> onSuspectedSource) Or _
                            (dtConfirmed <> odtConfirmed) Or _
                            (nIdentifiedBy <> onIdentifiedBy) Or _
                            (nLocation <> onLocation) Or _
                            (nExtent <> onExtent) Or _
                            (nCause <> onCause) Or _
                            (nEventPMId <> onEventPMId) Or _
                            (dtEventStarted <> odtEventStarted) Or _
                            (dtEventEnded <> odtEventEnded) Or _
                            (bolTOCSOIL <> obolTOCSOIL) Or _
                            (bolSOILBTEX <> obolSOILBTEX) Or _
                            (bolSOILPAH <> obolSOILPAH) Or _
                            (bolSOILTPH <> obolSOILTPH) Or _
                            (bolTOCGROUNDWATER <> obolTOCGROUNDWATER) Or _
                            (bolGWBTEX <> obolGWBTEX) Or _
                            (bolGWPAH <> obolGWPAH) Or _
                            (bolGWTPH <> obolGWTPH) Or _
                            (bolFREEPRODUCT <> obolFREEPRODUCT) Or _
                            (bolFPGASOLINE <> obolFPGASOLINE) Or _
                            (bolFPDIESEL <> obolFPDIESEL) Or _
                            (bolFPKEROSENE <> obolFPKEROSENE) Or _
                            (bolFPWASTEOIL <> obolFPWASTEOIL) Or _
                            (bolFPUNKNOWN <> obolFPUNKNOWN) Or _
                            (bolTOCVAPOR <> obolTOCVAPOR) Or _
                            (bolVAPORBTEX <> obolVAPORBTEX) Or _
                            (bolVAPORPAH <> obolVAPORPAH) Or _
                            (bolDeleted <> obolDeleted) Or _
                            (nEVENTSEQUENCE <> onEVENTSEQUENCE) Or _
                            (strTFCheckList <> ostrTFCheckList) Or _
                            (strTankandPipe <> ostrTankandPipe) Or _
                            (nPM_HEAD_ASSESS <> onPM_HEAD_ASSESS) Or _
                            (dtPM_HEAD_DATE <> odtPM_HEAD_DATE) Or _
                            (strPM_HEAD_BY <> ostrPM_HEAD_BY) Or _
                            (nUST_CHIEF_ASSESS <> onUST_CHIEF_ASSESS) Or _
                            (dtUST_CHIEF_DATE <> odtUST_CHIEF_DATE) Or _
                            (strUST_CHIEF_BY <> ostrUST_CHIEF_BY) Or _
                            (nOPC_HEAD_ASSESS <> onOPC_HEAD_ASSESS) Or _
                            (dtOPC_HEAD_DATE <> odtOPC_HEAD_DATE) Or _
                            (strOPC_HEAD_BY <> ostrOPC_HEAD_BY) Or _
                            (bolFOR_OPC_HEAD <> obolFOR_OPC_HEAD) Or _
                            (nCOMMISSION_ASSESS <> onCOMMISSION_ASSESS) Or _
                            (dtCOMMISSION_DATE <> odtCOMMISSION_DATE) Or _
                            (strCOMMISSION_BY <> ostrCOMMISSION_BY) Or _
                            (bolForCommission <> obolForCommission) Or _
                            (strELIGIBITY_COMMENTS <> ostrELIGIBITY_COMMENTS) Or _
                            (nIRAC <> onIRAC) Or _
                            (nERAC <> onERAC) Or _
                            (obolHOW_DISC_FAC_LEAK_DETECTION <> obolHOW_DISC_FAC_LEAK_DETECTION) Or _
                            (obolHOW_DISC_SURFACE_SHEEN <> obolHOW_DISC_SURFACE_SHEEN) Or _
                            (obolHOW_DISC_GW_WELL <> obolHOW_DISC_GW_WELL) Or _
                            (obolHOW_DISC_GW_CONTAMINATION <> obolHOW_DISC_GW_CONTAMINATION) Or _
                            (obolHOW_DISC_VAPORS <> obolHOW_DISC_VAPORS) Or _
                            (obolHOW_DISC_FREE_PRODUCT <> obolHOW_DISC_FREE_PRODUCT) Or _
                            (obolHOW_DISC_SOIL_CONTAMINATION <> obolHOW_DISC_SOIL_CONTAMINATION) Or _
                            (obolHOW_DISC_FAILED_PTT <> obolHOW_DISC_FAILED_PTT) Or _
                            (obolHOW_DISC_INVENTORY_SHORTAGE <> obolHOW_DISC_INVENTORY_SHORTAGE) Or _
                            (obolHOW_DISC_TANK_CLOSURE <> obolHOW_DISC_TANK_CLOSURE) Or _
                            (obolHOW_DISC_INSPECTION <> obolHOW_DISC_INSPECTION)

            bolIsEligibityDirty = (nPM_HEAD_ASSESS <> onPM_HEAD_ASSESS) Or _
                            (nUST_CHIEF_ASSESS <> onUST_CHIEF_ASSESS) Or _
                            (nOPC_HEAD_ASSESS <> onOPC_HEAD_ASSESS) Or _
                           (nCOMMISSION_ASSESS <> onCOMMISSION_ASSESS)


            If bolIsDirty Then
                RaiseEvent LustEventInfoChanged()
            End If

        End Sub
        Private Sub Init()
            onEventID = 0
            odtGWS = System.DateTime.Now
            odtLDR = System.DateTime.Now
            odtPTT = System.DateTime.Now
            odtReportDate = System.DateTime.Now
            odtCompAssDate = System.DateTime.Now
            odtStartDate = System.DateTime.Now
            onEventStatus = 0
            onFacilityID = 0
            onFlagID = 0
            onMGPTFStatus = 0
            onPriority = 0
            onProjMgr = 0
            onReleaseStatus = 0
            onReportSource = 0
            strRelatedSites = String.Empty
            onSuspectedSource = 0
            'onHowDiscoveredID = 0
            odtConfirmed = System.DateTime.Now
            onIdentifiedBy = 0
            onLocation = 0
            onExtent = 0
            onCause = 0
            onEventPMId = 0
            odtEventStarted = System.DateTime.Now
            odtEventEnded = System.DateTime.Now

            obolTOCSOIL = False
            obolSOILBTEX = False
            obolSOILPAH = False
            obolSOILTPH = False
            obolTOCGROUNDWATER = False
            obolGWBTEX = False
            obolGWPAH = False
            obolGWTPH = False
            obolFREEPRODUCT = False
            obolFPGASOLINE = False
            obolFPDIESEL = False
            obolFPKEROSENE = False
            obolFPWASTEOIL = False
            obolFPUNKNOWN = False
            obolTOCVAPOR = False
            obolVAPORBTEX = False
            obolVAPORPAH = False
            obolForCommission = False

            onEVENTSEQUENCE = 0

            onPM_HEAD_ASSESS = 0
            odtPM_HEAD_DATE = System.DateTime.Now
            ostrPM_HEAD_BY = String.Empty
            onUST_CHIEF_ASSESS = 0
            odtUST_CHIEF_DATE = System.DateTime.Now
            ostrUST_CHIEF_BY = String.Empty
            onOPC_HEAD_ASSESS = 0
            odtOPC_HEAD_DATE = System.DateTime.Now
            ostrOPC_HEAD_BY = String.Empty
            obolFOR_OPC_HEAD = False
            onCOMMISSION_ASSESS = 0
            odtCOMMISSION_DATE = System.DateTime.Now
            ostrCOMMISSION_BY = String.Empty
            ostrELIGIBITY_COMMENTS = String.Empty

            ostrCreatedBy = String.Empty
            ostrModifiedBy = String.Empty

            onIRAC = 0
            onERAC = 0

            obolHOW_DISC_FAC_LEAK_DETECTION = False
            obolHOW_DISC_SURFACE_SHEEN = False
            obolHOW_DISC_GW_WELL = False
            obolHOW_DISC_GW_CONTAMINATION = False
            obolHOW_DISC_VAPORS = False
            obolHOW_DISC_FREE_PRODUCT = False
            obolHOW_DISC_SOIL_CONTAMINATION = False
            obolHOW_DISC_FAILED_PTT = False
            obolHOW_DISC_INVENTORY_SHORTAGE = False
            obolHOW_DISC_TANK_CLOSURE = False
            obolHOW_DISC_INSPECTION = False
        End Sub
        Public Sub Reset()
            nEventID = onEventID
            dtGWS = odtGWS
            dtLDR = odtLDR
            dtPTT = odtPTT
            dtReportDate = odtReportDate
            dtCompAssDate = odtCompAssDate
            dtStartDate = odtStartDate
            nEventStatus = onEventStatus
            nFacilityID = onFacilityID
            nFlagID = onFlagID
            nMGPTFStatus = onMGPTFStatus
            nPriority = onPriority
            nProjMgr = onProjMgr
            nReleaseStatus = onReleaseStatus
            nReportSource = onReportSource
            strRelatedSites = ostrRelatedSites
            nSuspectedSource = onSuspectedSource
            'nHowDiscoveredID = onHowDiscoveredID
            dtConfirmed = odtConfirmed
            nIdentifiedBy = onIdentifiedBy
            nLocation = onLocation
            nExtent = onExtent
            nCause = onCause
            nEventPMId = onEventPMId
            dtEventStarted = odtEventStarted
            dtEventEnded = odtEventEnded

            bolTOCSOIL = obolTOCSOIL
            bolSOILBTEX = obolSOILBTEX
            bolSOILPAH = obolSOILPAH
            bolSOILTPH = obolSOILTPH
            bolTOCGROUNDWATER = obolTOCGROUNDWATER
            bolGWBTEX = obolGWBTEX
            bolGWPAH = obolGWPAH
            bolGWTPH = obolGWTPH
            bolFREEPRODUCT = obolFREEPRODUCT
            bolFPGASOLINE = obolFPGASOLINE
            bolFPDIESEL = obolFPDIESEL
            bolFPKEROSENE = obolFPKEROSENE
            bolFPWASTEOIL = obolFPWASTEOIL
            bolFPUNKNOWN = obolFPUNKNOWN
            bolTOCVAPOR = obolTOCVAPOR
            bolVAPORBTEX = obolVAPORBTEX
            bolVAPORPAH = obolVAPORPAH
            nEVENTSEQUENCE = onEVENTSEQUENCE
            bolForCommission = obolForCommission
            strTFCheckList = ostrTFCheckList
            strTankandPipe = ostrTankandPipe

            nPM_HEAD_ASSESS = onPM_HEAD_ASSESS
            dtPM_HEAD_DATE = odtPM_HEAD_DATE
            strPM_HEAD_BY = ostrPM_HEAD_BY
            nUST_CHIEF_ASSESS = onUST_CHIEF_ASSESS
            dtUST_CHIEF_DATE = odtUST_CHIEF_DATE
            strUST_CHIEF_BY = ostrUST_CHIEF_BY
            nOPC_HEAD_ASSESS = onOPC_HEAD_ASSESS
            dtOPC_HEAD_DATE = odtOPC_HEAD_DATE
            strOPC_HEAD_BY = ostrOPC_HEAD_BY
            bolFOR_OPC_HEAD = obolFOR_OPC_HEAD
            nCOMMISSION_ASSESS = onCOMMISSION_ASSESS
            dtCOMMISSION_DATE = odtCOMMISSION_DATE
            strCOMMISSION_BY = ostrCOMMISSION_BY
            strELIGIBITY_COMMENTS = ostrELIGIBITY_COMMENTS
            nIRAC = onIRAC
            nERAC = onERAC

            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn

            bolHOW_DISC_FAC_LEAK_DETECTION = obolHOW_DISC_FAC_LEAK_DETECTION
            bolHOW_DISC_SURFACE_SHEEN = obolHOW_DISC_SURFACE_SHEEN
            bolHOW_DISC_GW_WELL = obolHOW_DISC_GW_WELL
            bolHOW_DISC_GW_CONTAMINATION = obolHOW_DISC_GW_CONTAMINATION
            bolHOW_DISC_VAPORS = obolHOW_DISC_VAPORS
            bolHOW_DISC_FREE_PRODUCT = obolHOW_DISC_FREE_PRODUCT
            bolHOW_DISC_SOIL_CONTAMINATION = obolHOW_DISC_SOIL_CONTAMINATION
            bolHOW_DISC_FAILED_PTT = obolHOW_DISC_FAILED_PTT
            bolHOW_DISC_INVENTORY_SHORTAGE = obolHOW_DISC_INVENTORY_SHORTAGE
            bolHOW_DISC_TANK_CLOSURE = obolHOW_DISC_TANK_CLOSURE
            bolHOW_DISC_INSPECTION = obolHOW_DISC_INSPECTION
        End Sub
        Public Sub Archive()
            onEventID = nEventID
            odtGWS = dtGWS
            odtLDR = dtLDR
            odtPTT = dtPTT
            odtReportDate = dtReportDate
            odtCompAssDate = dtCompAssDate
            odtStartDate = dtStartDate
            onEventStatus = nEventStatus
            onFacilityID = nFacilityID
            onFlagID = nFlagID
            onMGPTFStatus = nMGPTFStatus
            onPriority = nPriority
            onProjMgr = nProjMgr
            onReleaseStatus = nReleaseStatus
            onReportSource = nReportSource
            ostrRelatedSites = strRelatedSites
            onSuspectedSource = nSuspectedSource
            'onHowDiscoveredID = nHowDiscoveredID
            odtConfirmed = dtConfirmed
            onIdentifiedBy = nIdentifiedBy
            onLocation = nLocation
            onExtent = nExtent
            onCause = nCause
            onEventPMId = nEventPMId
            odtEventStarted = dtEventStarted
            odtEventEnded = dtEventEnded

            obolTOCSOIL = bolTOCSOIL
            obolSOILBTEX = bolSOILBTEX
            obolSOILPAH = bolSOILPAH
            obolSOILTPH = bolSOILTPH
            obolTOCGROUNDWATER = bolTOCGROUNDWATER
            obolGWBTEX = bolGWBTEX
            obolGWPAH = bolGWPAH
            obolGWTPH = bolGWTPH
            obolFREEPRODUCT = bolFREEPRODUCT
            obolFPGASOLINE = bolFPGASOLINE
            obolFPDIESEL = bolFPDIESEL
            obolFPKEROSENE = bolFPKEROSENE
            obolFPWASTEOIL = bolFPWASTEOIL
            obolFPUNKNOWN = bolFPUNKNOWN
            obolTOCVAPOR = bolTOCVAPOR
            obolVAPORBTEX = bolVAPORBTEX
            obolVAPORPAH = bolVAPORPAH
            onEVENTSEQUENCE = nEVENTSEQUENCE
            obolForCommission = bolForCommission

            ostrTFCheckList = strTFCheckList
            ostrTankandPipe = strTankandPipe

            onPM_HEAD_ASSESS = nPM_HEAD_ASSESS
            odtPM_HEAD_DATE = dtPM_HEAD_DATE
            ostrPM_HEAD_BY = strPM_HEAD_BY
            onUST_CHIEF_ASSESS = nUST_CHIEF_ASSESS
            odtUST_CHIEF_DATE = dtUST_CHIEF_DATE
            ostrUST_CHIEF_BY = strUST_CHIEF_BY
            onOPC_HEAD_ASSESS = nOPC_HEAD_ASSESS
            odtOPC_HEAD_DATE = dtOPC_HEAD_DATE
            ostrOPC_HEAD_BY = strOPC_HEAD_BY
            obolFOR_OPC_HEAD = bolFOR_OPC_HEAD
            onCOMMISSION_ASSESS = nCOMMISSION_ASSESS
            odtCOMMISSION_DATE = dtCOMMISSION_DATE
            ostrCOMMISSION_BY = strCOMMISSION_BY
            ostrELIGIBITY_COMMENTS = strELIGIBITY_COMMENTS
            onIRAC = nIRAC
            onERAC = nERAC
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn

            obolHOW_DISC_FAC_LEAK_DETECTION = bolHOW_DISC_FAC_LEAK_DETECTION
            obolHOW_DISC_SURFACE_SHEEN = bolHOW_DISC_SURFACE_SHEEN
            obolHOW_DISC_GW_WELL = bolHOW_DISC_GW_WELL
            obolHOW_DISC_GW_CONTAMINATION = bolHOW_DISC_GW_CONTAMINATION
            obolHOW_DISC_VAPORS = bolHOW_DISC_VAPORS
            obolHOW_DISC_FREE_PRODUCT = bolHOW_DISC_FREE_PRODUCT
            obolHOW_DISC_SOIL_CONTAMINATION = bolHOW_DISC_SOIL_CONTAMINATION
            obolHOW_DISC_FAILED_PTT = bolHOW_DISC_FAILED_PTT
            obolHOW_DISC_INVENTORY_SHORTAGE = bolHOW_DISC_INVENTORY_SHORTAGE
            obolHOW_DISC_TANK_CLOSURE = bolHOW_DISC_TANK_CLOSURE
            obolHOW_DISC_INSPECTION = bolHOW_DISC_INSPECTION

            bolIsDirty = False
        End Sub
#End Region
#Region "Protected Operations"
        Protected Overrides Sub Finalize()
        End Sub
#End Region
    End Class
End Namespace
