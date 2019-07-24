'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.LustEvent
'   Provides the operations required to manipulate an Entity object.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0         AN       3/8/2005    Original class definition
'
' Function          Description
' Retrieve(Name)    Returns the Lust Event requested by the string arg NAME
' Retrieve(ID)      Returns the Lust Event requested by the int arg ID
' GetAll()          Returns an LustEventsCollection with all Lust Event objects
' Add(ID)           Adds the Lust Event identified by arg ID to the 
'                           internal LustEventsCollection
' Add(LustEventInfo)Adds the Lust Event passed as the argument to the internal 
'                           LustEventsCollection
' Remove(ID)        Removes the Lust Event identified by arg ID from the internal 
'                           LustEventsCollection
' Remove(NAME)      Removes the Lust Event identified by arg NAME from the 
'                           internal LustEventsCollection
' Flush()           Saves all objects in the collection
' Clear()           Clears the current object and all objects in the collection
' Reset()           Resets the current object to its original state
' EntityTable()     Returns a datatable containing all columns for the Lust Event 
'                           objects in the internal LustEventsCollection.
'
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
'-------------------------------------------------------------------------------

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pLustEvent
#Region "Public Events"
        Public Event LustEventErr(ByVal MsgStr As String)
        Public Event LustEventChanged(ByVal bolValue As Boolean)
        Public Event ColChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private Member Variables"
        Private WithEvents oLustEventInfo As New MUSTER.Info.LustEventInfo
        Private WithEvents colLustEvents As MUSTER.Info.LustEventCollection
        Private oLustEventDB As New MUSTER.DataAccess.LustEventDB
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private nID As Int64 = -1
        Private colIndex As Int64 = 0
        Private colKey As String = String.Empty
        'Private nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("Lust Event").ID
        Private WithEvents oLustActivity As MUSTER.BusinessLogic.pLustEventActivity
        Private oCalendar As New MUSTER.BusinessLogic.pCalendar
        Private oLustActivities As MUSTER.BusinessLogic.pLustEventActivity

        Private onUserID As Integer = 0
#End Region
#Region "Constructors"

        Public Sub New()
            oLustEventInfo = New MUSTER.Info.LustEventInfo
            colLustEvents = New MUSTER.Info.LustEventCollection
            oLustActivities = New MUSTER.BusinessLogic.pLustEventActivity(oLustEventInfo)
        End Sub

        Public Sub New(ByVal UserID As Integer)
            oLustEventInfo = New MUSTER.Info.LustEventInfo
            colLustEvents = New MUSTER.Info.LustEventCollection
            onUserID = UserID
            oLustActivities = New MUSTER.BusinessLogic.pLustEventActivity(oLustEventInfo)
        End Sub
        '********************************************************
        '
        ' Overloaded NEW which will populate with a single instance
        '   of the LustEvent object with the provided ID.
        '
        '********************************************************
        Public Sub New(ByVal LustEventID As Integer, ByVal UserID As Integer)
            oLustEventInfo = New MUSTER.Info.LustEventInfo
            colLustEvents = New MUSTER.Info.LustEventCollection
            onUserID = UserID
            oLustActivities = New MUSTER.BusinessLogic.pLustEventActivity(LustEventID, oLustEventInfo)
            Me.Retrieve(LustEventID)
        End Sub
        '********************************************************
        '
        ' Overloaded NEW which will populate with a single instance
        '   of the named LustEvent object.
        '
        '********************************************************
        Public Sub New(ByVal LustEventName As String, ByVal UserID As Integer)
            oLustEventInfo = New MUSTER.Info.LustEventInfo
            colLustEvents = New MUSTER.Info.LustEventCollection
            onUserID = UserID
            oLustActivities = New MUSTER.BusinessLogic.pLustEventActivity(oLustEventInfo)
            Me.Retrieve(LustEventName)
        End Sub
#End Region
#Region "Exposed Attributes"

        Public ReadOnly Property PMDesc() As String
            Get
                Return oLustEventInfo.PMDesc
            End Get
        End Property

        Public ReadOnly Property MGPTFStatusDesc() As String
            Get
                Return oLustEventInfo.MGPTFStatusDesc
            End Get
        End Property

        Public ReadOnly Property TechnicalStatusDesc() As String
            Get
                Return oLustEventInfo.TechnicalStatusDesc
            End Get
        End Property

        Public ReadOnly Property IsNewRelease() As Boolean
            Get
                Return oLustEventInfo.IsNewRelease
            End Get
        End Property

        Public Property ID() As Integer
            Get
                Return oLustEventInfo.ID
            End Get
            Set(ByVal Value As Integer)
                oLustEventInfo.ID = Integer.Parse(Value)
            End Set
        End Property

        ' The current status of the LUST event - derived from properties table
        Public Property EventStatus() As Integer
            Get
                Return oLustEventInfo.EventStatus
            End Get
            Set(ByVal Value As Integer)
                oLustEventInfo.EventStatus = Value
            End Set
        End Property
        ' The LUST event ID for the facility
        Public Property FacilityID() As Integer
            Get
                Return oLustEventInfo.FacilityID
            End Get
            Set(ByVal Value As Integer)
                oLustEventInfo.FacilityID = Value
            End Set
        End Property
        ' The system ID of the flag associated with the LUST event
        Public Property FlagID() As Long
            Get
                Return oLustEventInfo.FlagID
            End Get
            Set(ByVal Value As Long)
                oLustEventInfo.FlagID = Value
            End Set
        End Property
        ' Date of last GWS for LUST event
        Public Property LastGWS() As Date
            Get
                Return oLustEventInfo.LastGWS
            End Get
            Set(ByVal Value As Date)
                oLustEventInfo.LastGWS = Value
            End Set
        End Property
        ' The date of the last LDR for the LUST event
        Public Property LastLDR() As Date
            Get
                Return oLustEventInfo.LastLDR
            End Get
            Set(ByVal Value As Date)
                oLustEventInfo.LastLDR = Value
            End Set
        End Property
        ' The date of the last PTT for the LUST event
        Public Property LastPTT() As Date
            Get
                Return oLustEventInfo.LastPTT
            End Get
            Set(ByVal Value As Date)
                oLustEventInfo.LastPTT = Value
            End Set
        End Property
        ' The MGPTF status for the LUST event - derived from properties table
        Public Property MGPTFStatus() As Integer
            Get
                Return oLustEventInfo.MGPTFStatus
            End Get
            Set(ByVal Value As Integer)
                oLustEventInfo.MGPTFStatus = Value
            End Set
        End Property
        ' The current Project Manager for the LUST event
        Public Property PM() As Integer
            Get
                Return oLustEventInfo.PM
            End Get
            Set(ByVal Value As Integer)
                oLustEventInfo.PM = Value
            End Set
        End Property
        ' The current priority for the LUST event - can only range 1 to 8
        Public Property Priority() As Short
            Get
                Return oLustEventInfo.Priority
            End Get
            Set(ByVal Value As Short)
                oLustEventInfo.Priority = Value
            End Set
        End Property
        ' The current release status for the LUST event - derived from properties table
        Public Property ReleaseStatus() As Integer
            Get
                Return oLustEventInfo.ReleaseStatus
            End Get
            Set(ByVal Value As Integer)
                oLustEventInfo.ReleaseStatus = Value
            End Set
        End Property
        ' The date the LUST event was reported to UST
        Public Property ReportDate() As Date
            Get
                Return oLustEventInfo.ReportDate
            End Get
            Set(ByVal Value As Date)
                oLustEventInfo.ReportDate = Value
            End Set
        End Property
        Public Property CompAssDate() As Date
            Get
                Return oLustEventInfo.CompAssDate
            End Get
            Set(ByVal Value As Date)
                oLustEventInfo.CompAssDate = Value
            End Set
        End Property
        ' The source of the LUST event report to UST - derived from properties table
        Public Property ReportSource() As Integer
            Get
                Return oLustEventInfo.ReportSource
            End Get
            Set(ByVal Value As Integer)
                oLustEventInfo.ReportSource = Value
            End Set
        End Property
        ' The date the LUST event was started
        Public Property Started() As Date
            Get
                Return oLustEventInfo.Started
            End Get
            Set(ByVal Value As Date)
                oLustEventInfo.Started = Value
            End Set
        End Property


        Public Property SuspectedSource() As Integer
            Get
                Return oLustEventInfo.SuspectedSource
            End Get
            Set(ByVal Value As Integer)
                oLustEventInfo.SuspectedSource = Value
            End Set
        End Property

        'Public Property HowDiscoveredID() As Integer
        '    Get
        '        Return oLustEventInfo.HowDiscoveredID
        '    End Get
        '    Set(ByVal Value As Integer)
        '        oLustEventInfo.HowDiscoveredID = Value
        '    End Set
        'End Property
        Public Property Confirmed() As Date
            Get
                Return oLustEventInfo.Confirmed
            End Get
            Set(ByVal Value As Date)
                oLustEventInfo.Confirmed = Value
            End Set
        End Property

        Public Property IDENTIFIEDBY() As Integer
            Get
                Return oLustEventInfo.IDENTIFIEDBY
            End Get
            Set(ByVal Value As Integer)
                oLustEventInfo.IDENTIFIEDBY = Value
            End Set
        End Property

        Public Property Location() As Integer
            Get
                Return oLustEventInfo.Location
            End Get
            Set(ByVal Value As Integer)
                oLustEventInfo.Location = Value
            End Set
        End Property

        Public Property Extent() As Integer
            Get
                Return oLustEventInfo.Extent
            End Get
            Set(ByVal Value As Integer)
                oLustEventInfo.Extent = Value
            End Set
        End Property

        Public Property Cause() As Integer
            Get
                Return oLustEventInfo.Cause
            End Get
            Set(ByVal Value As Integer)
                oLustEventInfo.Cause = Value
            End Set
        End Property

        Public Property EventPMId() As Integer
            Get
                Return oLustEventInfo.EventPMId
            End Get
            Set(ByVal Value As Integer)
                oLustEventInfo.EventPMId = Value
            End Set
        End Property
        Public Property EventStarted() As Date
            Get
                Return oLustEventInfo.EventStarted
            End Get
            Set(ByVal Value As Date)
                oLustEventInfo.EventStarted = Value
            End Set
        End Property
        Public Property EventEnded() As Date
            Get
                Return oLustEventInfo.EventEnded
            End Get
            Set(ByVal Value As Date)
                oLustEventInfo.EventEnded = Value
            End Set
        End Property

        Public Property Deleted() As Boolean
            Get
                Return oLustEventInfo.Deleted
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.Deleted = Boolean.Parse(Value)
            End Set
        End Property

        Public Property RelatedSites() As String
            Get
                Return oLustEventInfo.RelatedSites
            End Get
            Set(ByVal Value As String)
                oLustEventInfo.RelatedSites = Value
            End Set
        End Property

        Public Property TOCSOIL() As Boolean
            Get
                Return oLustEventInfo.TOCSOIL
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.TOCSOIL = Value
            End Set
        End Property
        Public Property SOILBTEX() As Boolean
            Get
                Return oLustEventInfo.SOILBTEX
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.SOILBTEX = Value
            End Set
        End Property
        Public Property SOILPAH() As Boolean
            Get
                Return oLustEventInfo.SOILPAH
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.SOILPAH = Value
            End Set
        End Property
        Public Property SOILTPH() As Boolean
            Get
                Return oLustEventInfo.SOILTPH
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.SOILTPH = Value
            End Set
        End Property
        Public Property TOCGROUNDWATER() As Boolean
            Get
                Return oLustEventInfo.TOCGROUNDWATER
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.TOCGROUNDWATER = Value
            End Set
        End Property
        Public Property GWBTEX() As Boolean
            Get
                Return oLustEventInfo.GWBTEX
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.GWBTEX = Value
            End Set
        End Property
        Public Property GWPAH() As Boolean
            Get
                Return oLustEventInfo.GWPAH
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.GWPAH = Value
            End Set
        End Property
        Public Property GWTPH() As Boolean
            Get
                Return oLustEventInfo.GWTPH
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.GWTPH = Value
            End Set
        End Property
        Public Property FREEPRODUCT() As Boolean
            Get
                Return oLustEventInfo.FREEPRODUCT
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.FREEPRODUCT = Value
            End Set
        End Property
        Public Property FPGASOLINE() As Boolean
            Get
                Return oLustEventInfo.FPGASOLINE
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.FPGASOLINE = Value
            End Set
        End Property
        Public Property FPDIESEL() As Boolean
            Get
                Return oLustEventInfo.FPDIESEL
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.FPDIESEL = Value
            End Set
        End Property
        Public Property FPKEROSENE() As Boolean
            Get
                Return oLustEventInfo.FPKEROSENE
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.FPKEROSENE = Value
            End Set
        End Property
        Public Property FPWASTEOIL() As Boolean
            Get
                Return oLustEventInfo.FPWASTEOIL
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.FPWASTEOIL = Value
            End Set
        End Property
        Public Property FPUNKNOWN() As Boolean
            Get
                Return oLustEventInfo.FPUNKNOWN
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.FPUNKNOWN = Value
            End Set
        End Property
        Public Property TOCVAPOR() As Boolean
            Get
                Return oLustEventInfo.TOCVAPOR
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.TOCVAPOR = Value
            End Set
        End Property
        Public Property VAPORBTEX() As Boolean
            Get
                Return oLustEventInfo.VAPORBTEX
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.VAPORBTEX = Value
            End Set
        End Property
        Public Property VAPORPAH() As Boolean
            Get
                Return oLustEventInfo.VAPORPAH
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.VAPORPAH = Value
            End Set
        End Property

        Public Property EVENTSEQUENCE() As Integer
            Get
                Return oLustEventInfo.EVENTSEQUENCE
            End Get
            Set(ByVal Value As Integer)
                oLustEventInfo.EVENTSEQUENCE = Value
            End Set
        End Property

        Public Property TFCheckList() As String
            Get
                Return oLustEventInfo.TFCheckList
            End Get
            Set(ByVal Value As String)
                oLustEventInfo.TFCheckList = Value
            End Set
        End Property

        Public Property TankandPipe() As String
            Get
                Return oLustEventInfo.TankandPipe
            End Get
            Set(ByVal Value As String)
                oLustEventInfo.TankandPipe = Value
            End Set
        End Property

        Public Property PM_HEAD_ASSESS() As Int32
            Get
                Return oLustEventInfo.PM_HEAD_ASSESS
            End Get
            Set(ByVal Value As Int32)
                oLustEventInfo.PM_HEAD_ASSESS = Value
            End Set
        End Property

        Public Property PM_HEAD_DATE() As Date
            Get
                Return oLustEventInfo.PM_HEAD_DATE
            End Get
            Set(ByVal Value As Date)
                oLustEventInfo.PM_HEAD_DATE = Value
            End Set
        End Property

        Public Property PM_HEAD_BY() As String
            Get
                Return oLustEventInfo.PM_HEAD_BY
            End Get
            Set(ByVal Value As String)
                oLustEventInfo.PM_HEAD_BY = Value
            End Set
        End Property

        Public Property UST_CHIEF_ASSESS() As Int32
            Get
                Return oLustEventInfo.UST_CHIEF_ASSESS
            End Get
            Set(ByVal Value As Int32)
                oLustEventInfo.UST_CHIEF_ASSESS = Value
            End Set
        End Property

        Public Property UST_CHIEF_DATE() As Date
            Get
                Return oLustEventInfo.UST_CHIEF_DATE
            End Get
            Set(ByVal Value As Date)
                oLustEventInfo.UST_CHIEF_DATE = Value
            End Set
        End Property

        Public Property UST_CHIEF_BY() As String
            Get
                Return oLustEventInfo.UST_CHIEF_BY
            End Get
            Set(ByVal Value As String)
                oLustEventInfo.UST_CHIEF_BY = Value
            End Set
        End Property

        Public Property OPC_HEAD_ASSESS() As Int32
            Get
                Return oLustEventInfo.OPC_HEAD_ASSESS
            End Get
            Set(ByVal Value As Int32)
                oLustEventInfo.OPC_HEAD_ASSESS = Value
            End Set
        End Property

        Public Property OPC_HEAD_DATE() As Date
            Get
                Return oLustEventInfo.OPC_HEAD_DATE
            End Get
            Set(ByVal Value As Date)
                oLustEventInfo.OPC_HEAD_DATE = Value
            End Set
        End Property

        Public Property OPC_HEAD_BY() As String
            Get
                Return oLustEventInfo.OPC_HEAD_BY
            End Get
            Set(ByVal Value As String)
                oLustEventInfo.OPC_HEAD_BY = Value
            End Set
        End Property

        Public Property FOR_OPC_HEAD() As Boolean
            Get
                Return oLustEventInfo.FOR_OPC_HEAD
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.FOR_OPC_HEAD = Value
            End Set
        End Property

        Public Property FOR_COMMISSION() As Boolean
            Get
                Return oLustEventInfo.FOR_COMMISSION
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.FOR_COMMISSION = Value
            End Set
        End Property

        Public Property COMMISSION_ASSESS() As Int32
            Get
                Return oLustEventInfo.COMMISSION_ASSESS
            End Get
            Set(ByVal Value As Int32)
                oLustEventInfo.COMMISSION_ASSESS = Value
            End Set
        End Property

        Public Property COMMISSION_DATE() As Date
            Get
                Return oLustEventInfo.COMMISSION_DATE
            End Get
            Set(ByVal Value As Date)
                oLustEventInfo.COMMISSION_DATE = Value
            End Set
        End Property

        Public Property COMMISSION_BY() As String
            Get
                Return oLustEventInfo.COMMISSION_BY
            End Get
            Set(ByVal Value As String)
                oLustEventInfo.COMMISSION_BY = Value
            End Set
        End Property

        Public Property ELIGIBITY_COMMENTS() As String
            Get
                Return oLustEventInfo.ELIGIBITY_COMMENTS
            End Get
            Set(ByVal Value As String)
                oLustEventInfo.ELIGIBITY_COMMENTS = Value
            End Set
        End Property

        Public Property Activities() As MUSTER.Info.LustActivityCollection
            Get
                Return oLustEventInfo.Activities
            End Get
            Set(ByVal Value As MUSTER.Info.LustActivityCollection)
                oLustEventInfo.Activities = Value
            End Set
        End Property

        Public Property UserID() As Integer
            Get
                Return oLustEventInfo.UserID
            End Get
            Set(ByVal Value As Integer)
                oLustEventInfo.UserID = Value
            End Set
        End Property

        Public Property IRAC() As Integer
            Get
                Return oLustEventInfo.IRAC
            End Get
            Set(ByVal Value As Integer)
                oLustEventInfo.IRAC = Value
            End Set
        End Property
        Public Property ERAC() As Integer
            Get
                Return oLustEventInfo.ERAC
            End Get
            Set(ByVal Value As Integer)
                oLustEventInfo.ERAC = Value
            End Set
        End Property

        Public Property IsEligibityDirty() As Boolean
            Get
                Return oLustEventInfo.IsEligibityDirty
            End Get

            Set(ByVal value As Boolean)
                oLustEventInfo.IsEligibityDirty = Boolean.Parse(value)
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return oLustEventInfo.IsDirty
            End Get

            Set(ByVal value As Boolean)
                oLustEventInfo.IsDirty = Boolean.Parse(value)
            End Set
        End Property
        Public Property colIsDirty() As Boolean
            Get
                Dim xLustEventinfo As MUSTER.Info.LustEventInfo
                For Each xLustEventinfo In colLustEvents.Values
                    If xLustEventinfo.IsDirty Then
                        Return True
                        Exit Property
                    End If
                Next
                Return False
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.IsDirty = Value
            End Set
        End Property

        Public Property CreatedBy() As String
            Get
                Return oLustEventInfo.CreatedBy
            End Get
            Set(ByVal Value As String)
                oLustEventInfo.CreatedBy = Value
            End Set
        End Property
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return oLustEventInfo.CreatedOn
            End Get
        End Property
        Public Property ModifiedBy() As String
            Get
                Return oLustEventInfo.ModifiedBy
            End Get
            Set(ByVal Value As String)
                oLustEventInfo.ModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return oLustEventInfo.ModifiedOn
            End Get
        End Property

        Public Property HowDiscFacLD() As Boolean
            Get
                Return oLustEventInfo.HowDiscFacLD
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.HowDiscFacLD = Value
            End Set
        End Property
        Public Property HowDiscSurfaceSheen() As Boolean
            Get
                Return oLustEventInfo.HowDiscSurfaceSheen
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.HowDiscSurfaceSheen = Value
            End Set
        End Property
        Public Property HowDiscGWWell() As Boolean
            Get
                Return oLustEventInfo.HowDiscGWWell
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.HowDiscGWWell = Value
            End Set
        End Property
        Public Property HowDiscGWContamination() As Boolean
            Get
                Return oLustEventInfo.HowDiscGWContamination
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.HowDiscGWContamination = Value
            End Set
        End Property
        Public Property HowDiscVapors() As Boolean
            Get
                Return oLustEventInfo.HowDiscVapors
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.HowDiscVapors = Value
            End Set
        End Property
        Public Property HowDiscFreeProduct() As Boolean
            Get
                Return oLustEventInfo.HowDiscFreeProduct
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.HowDiscFreeProduct = Value
            End Set
        End Property
        Public Property HowDiscSoilContamination() As Boolean
            Get
                Return oLustEventInfo.HowDiscSoilContamination
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.HowDiscSoilContamination = Value
            End Set
        End Property
        Public Property HowDiscFailedPTT() As Boolean
            Get
                Return oLustEventInfo.HowDiscFailedPTT
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.HowDiscFailedPTT = Value
            End Set
        End Property
        Public Property HowDiscInventoryShortage() As Boolean
            Get
                Return oLustEventInfo.HowDiscInventoryShortage
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.HowDiscInventoryShortage = Value
            End Set
        End Property
        Public Property HowDiscTankClosure() As Boolean
            Get
                Return oLustEventInfo.HowDiscTankClosure
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.HowDiscTankClosure = Value
            End Set
        End Property
        Public Property HowDiscInspection() As Boolean
            Get
                Return oLustEventInfo.HowDiscInspection
            End Get
            Set(ByVal Value As Boolean)
                oLustEventInfo.HowDiscInspection = Value
            End Set
        End Property

#End Region
#Region "Exposed Operations"
#Region "Info Operations"
        'Obtains and returns an entity as called for by ID
        Public Function Retrieve(ByVal ID As Integer) As MUSTER.Info.LustEventInfo
            Dim oLustEventInfoLocal As MUSTER.Info.LustEventInfo
            Try
                For Each oLustEventInfoLocal In colLustEvents.Values
                    If oLustEventInfoLocal.ID = ID Then
                        oLustEventInfo = oLustEventInfoLocal
                        Return oLustEventInfo
                    End If
                Next
                oLustEventInfo = oLustEventDB.DBGetByID(ID)
                oLustEventInfo.UserID = onUserID
                If oLustEventInfo.ID = 0 Then
                    'oLustEventInfo.ID = nID
                    nID -= 1
                End If
                colLustEvents.Add(oLustEventInfo)
                Return oLustEventInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function Retrieve(ByVal FacilityID As Integer, ByVal Sequence As Integer) As MUSTER.Info.LustEventInfo
            Dim oLustEventInfoLocal As MUSTER.Info.LustEventInfo
            Try
                For Each oLustEventInfoLocal In colLustEvents.Values
                    If oLustEventInfoLocal.EVENTSEQUENCE = Sequence And oLustEventInfoLocal.FacilityID = FacilityID Then
                        oLustEventInfo = oLustEventInfoLocal
                        Return oLustEventInfo
                    End If
                Next
                oLustEventInfo = oLustEventDB.DBGetByFacilityAndSequence(FacilityID, Sequence)
                oLustEventInfo.UserID = onUserID
                If oLustEventInfo.ID = 0 Then
                    'oLustEventInfo.ID = nID
                    nID -= 1
                End If
                colLustEvents.Add(oLustEventInfo)
                Return oLustEventInfo
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        'Saves the data in the current Info object
        Public Sub Save(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim strModuleName As String = String.Empty
            Dim bolSubmitForCalendar As Boolean
            Try
                If Me.ValidateData(strModuleName) Then
                    If oLustEventInfo.ID > 0 Then
                        bolSubmitForCalendar = False
                    Else
                        bolSubmitForCalendar = True
                    End If

                    oLustEventDB.Put(oLustEventInfo, moduleID, staffID, returnVal)
                    If Not returnVal = String.Empty Then
                        Exit Sub
                    End If

                    'If bolSubmitForCalendar Then
                    '    CalendarEntries(oLustEventInfo.PM, "New Lust Event:" & oLustEventInfo.ID, True, False, "", oLustEventInfo.PM, Now.Date, Now.Date, oLustEventInfo.ID)
                    'End If

                    oLustEventInfo.Archive()

                End If
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Validates the data before saving
        Public Function ValidateData(Optional ByVal [module] As String = "Registration") As Boolean
            Dim errStr As String = ""
            Dim validateSuccess As Boolean = True 'False
            '********************************************************
            '
            ' Sample validation code below.  Modify as necessary to
            '  handle rules for the object.  Note that errStr should
            '  be built with all failed validation reasons as it is
            '  raised to the consumer for display to the user.  The
            '  same goes for the boolean validateSuccess.  Since
            '  validateSuccess ASSUMES failure, it must be set to
            '  TRUE if all validations are passed successfully.
            '********************************************************
            Try
                Select Case [module]
                    Case "Registration"

                        ' if any validations failed
                        Exit Select
                    Case "Technical"

                        ' if any validations failed
                        Exit Select
                End Select
                If errStr.Length > 0 Or Not validateSuccess Then
                    RaiseEvent LustEventErr(errStr)
                End If
                Return validateSuccess
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function CalendarEntries(ByVal strPMUserID As String, ByVal strTaskDesc As String, ByVal bolToDo As Boolean, ByVal bolDueToMe As Boolean, ByVal strGroupID As String, ByVal strUserID As String, ByVal dtNotificationDate As Date, ByVal dtDueDate As Date, Optional ByVal EventID As Int64 = 0, Optional ByVal ActivityID As Int64 = 0)
            Dim strSourceUserID As String = "SYSTEM"
            Dim EntityType As Int64
            Dim nEntityID As Int64

            If EventID > 0 Then
                nEntityID = EventID
                EntityType = 7
            ElseIf ActivityID > 0 Then
                nEntityID = ActivityID
                EntityType = 23
            Else
                nEntityID = oLustEventInfo.ID
                EntityType = 7
            End If



            'Create a Calendar Info object 
            Dim oCalendarInfo As MUSTER.Info.CalendarInfo

            oCalendarInfo = New MUSTER.Info.CalendarInfo(0, _
                                            dtNotificationDate, _
                                            dtDueDate, _
                                            0, _
                                            strTaskDesc, _
                                            strPMUserID, _
                                            strSourceUserID, _
                                            strGroupID, _
                                            bolDueToMe, _
                                            bolToDo, _
                                            False, _
                                            False, _
                                            strUserID, _
                                            Now(), _
                                            String.Empty, _
                                            CDate("01/01/0001"), _
                                            EntityType, _
                                            nEntityID)

            oCalendarInfo.OwningEntityID = nEntityID
            oCalendarInfo.OwningEntityType = EntityType
            oCalendarInfo.IsDirty = True
            oCalendar.Add(oCalendarInfo)
            oCalendar.Flush()

        End Function

#End Region
#Region "Collection Operations"
        'Gets all the info
        Function GetAll() As MUSTER.Info.LustEventCollection
            Try
                colLustEvents.Clear()
                colLustEvents = oLustEventDB.GetAllInfo
                Return colLustEvents
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Function GetAll(ByVal FacilityID As Int64) As MUSTER.Info.LustEventCollection
            Try
                colLustEvents.Clear()
                colLustEvents = oLustEventDB.DBGetByFacilityID(FacilityID)
                Return colLustEvents
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        'Adds an entity to the collection as called for by ID
        Public Sub Add(ByVal ID As Integer)
            Try
                oLustEventInfo = oLustEventDB.DBGetByID(ID)
                oLustEventInfo.UserID = onUserID
                If oLustEventInfo.ID = 0 Then
                    'oLustEventInfo.ID = nID
                    nID -= 1
                End If
                colLustEvents.Add(oLustEventInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Adds an entity to the collection as supplied by the caller
        Public Sub Add(ByRef oLustEvent As MUSTER.Info.LustEventInfo)
            Try
                oLustEventInfo = oLustEvent
                oLustEventInfo.UserID = onUserID
                colLustEvents.Add(oLustEventInfo)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
        'Removes the entity called for by ID from the collection
        Public Sub Remove(ByVal ID As Integer)
            Dim myIndex As Int16 = 1
            Dim oLustEventInfoLocal As MUSTER.Info.LustEventInfo

            Try
                For Each oLustEventInfoLocal In colLustEvents.Values
                    If oLustEventInfoLocal.ID = ID Then
                        colLustEvents.Remove(oLustEventInfoLocal)
                        Exit Sub
                    End If
                    myIndex += 1
                Next
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("Pipe " & ID.ToString & " is not in the collection of Pipes.")
        End Sub
        'Removes the entity supplied from the collection
        Public Sub Remove(ByVal oLustEvent As MUSTER.Info.LustEventInfo)
            Try
                colLustEvents.Remove(oLustEvent)
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
            Throw New Exception("LustEvent " & oLustEvent.ID & " is not in the collection of LustEvents.")
        End Sub
        Public Sub Flush(ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String)
            Dim xLustEventInfo As MUSTER.Info.LustEventInfo
            For Each xLustEventInfo In colLustEvents.Values
                If xLustEventInfo.IsDirty Then
                    oLustEventInfo = xLustEventInfo
                    '********************************************************
                    '
                    ' Note that if there are contained objects and the respective
                    '  contained object collections are dirty, then the contained
                    '  collections MUST BE FLUSHED before this object can be
                    '  saved.  Otherwise, there is a risk that an attempt will
                    '  be made to insert a new object to the repository without
                    '  corresponding contained information being present which 
                    '  may, in turn, cause a foreign key violation!
                    '
                    '********************************************************
                    Me.Save(moduleID, staffID, returnVal)
                End If
            Next
        End Sub
        Public Function GetNext() As String
            Return GetNextPrev(1)
        End Function
        Public Function GetPrevious() As String
            Return GetNextPrev(-1)
        End Function
        Private Function GetNextPrev(ByVal direction As Integer) As String
            Dim strArr() As String = colLustEvents.GetKeys()
            Dim nArr(strArr.GetUpperBound(0)) As Integer
            Dim y As String
            For Each y In strArr
                nArr(strArr.IndexOf(strArr, y)) = CInt(y)
            Next
            nArr.Sort(nArr)
            colIndex = Array.BinarySearch(nArr, Integer.Parse(Me.ID.ToString))
            If colIndex + direction > -1 And _
                colIndex + direction <= nArr.GetUpperBound(0) Then
                Return colLustEvents.Item(nArr.GetValue(colIndex + direction)).ID.ToString
            Else
                Return colLustEvents.Item(nArr.GetValue(colIndex)).ID.ToString
            End If
        End Function
#End Region
#Region "General Operations"
        Public Sub Clear()
            oLustEventInfo = New MUSTER.Info.LustEventInfo
        End Sub
        Public Sub Reset()
            oLustEventInfo.Reset()
        End Sub
#End Region
#Region "Miscellaneous Operations"
        Public Function EntityTable() As DataTable
            Dim oLustEventInfoLocal As New MUSTER.Info.PipeInfo
            Dim dr As DataRow
            Dim tbEntityTable As New DataTable
            Try
                tbEntityTable.Columns.Add("LustEvent ID")
                tbEntityTable.Columns.Add("Deleted")
                tbEntityTable.Columns.Add("Created By")
                tbEntityTable.Columns.Add("Date Created")
                tbEntityTable.Columns.Add("Last Edited By")
                tbEntityTable.Columns.Add("Date Last Edited")

                For Each oLustEventInfoLocal In colLustEvents.Values
                    dr = tbEntityTable.NewRow()
                    dr("Pipe ID") = oLustEventInfoLocal.ID
                    dr("Deleted") = oLustEventInfoLocal.Deleted
                    dr("Created By") = oLustEventInfoLocal.CreatedBy
                    dr("Date Created") = oLustEventInfoLocal.CreatedOn
                    dr("Last Edited By") = oLustEventInfoLocal.ModifiedBy
                    dr("Date Last Edited") = oLustEventInfoLocal.ModifiedOn
                    tbEntityTable.Rows.Add(dr)
                Next
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function


        Public Function LustActivityDocumentDataset(ByVal nEventID As Int64) As DataSet
            Dim dsActivityDoc As New DataSet
            Dim drRow As DataRow
            Dim oCol As DataColumn
            Dim dsRel As DataRelation
            Dim strSQL As String
            Try

                'strSQL = "SELECT * FROM V_LUST_TANK_DISPLAY_DATA WHERE FACILITY_ID = '" & oFacilityInfo.ID.ToString & "' ORDER BY [TANK SITE ID];" & _
                '        "SELECT * FROM V_LUST_PIPE_DISPLAY_DATA WHERE FACILITY_ID = '" & oFacilityInfo.ID.ToString & "' ORDER BY [PIPE SITE ID] "


                strSQL = "select *, (select count(relation_id) from tblTEC_ACT_DOC_RELATIONSHIP where activity_id = activity_Type_id) as ActivityDocsRelCount"
                strSQL = strSQL & " from vLUSTEVENTACTIVITY "
                strSQL = strSQL & " where Event_ID = " & oLustEventInfo.ID & ";"
                strSQL = strSQL & " "
                strSQL = strSQL & "select * from vLUSTEVENTACTIVITY_DOCUMENT "
                strSQL = strSQL & " where Event_ID = " & oLustEventInfo.ID & ";"


                dsActivityDoc = oLustEventDB.DBGetDS(strSQL)

                dsActivityDoc.Tables(0).DefaultView.Sort = "Start_Date DESC"
                dsActivityDoc.Tables(1).DefaultView.Sort = "Issued DESC"

                dsRel = New DataRelation("ActivityToDocument", dsActivityDoc.Tables(0).Columns("Event_Activity_ID"), dsActivityDoc.Tables(1).Columns("Event_Activity_ID"), False)

                dsActivityDoc.Relations.Add(dsRel)
                Return dsActivityDoc
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function LustGetLastGWS(ByVal nEventID As Int64) As DataSet
            Dim dsActivityDoc As New DataSet
            Dim drRow As DataRow
            Dim oCol As DataColumn
            Dim dsRel As DataRelation
            Dim strSQL As String
            Try

                strSQL = "select * from vLUSTEVENTACTIVITY_DOCUMENT "
                strSQL = strSQL & " where Event_ID = " & oLustEventInfo.ID & " order by  Received desc;"
                dsActivityDoc = oLustEventDB.DBGetDS(strSQL)
                dsActivityDoc.Tables(0).DefaultView.Sort = "Received DESC"

                '  dsRel = New DataRelation("ActivityToDocument", dsActivityDoc.Tables(0).Columns("Event_Activity_ID"), dsActivityDoc.Tables(1).Columns("Event_Activity_ID"), False)
                'dsActivityDoc.Relations.Add(dsRel)
                Return dsActivityDoc
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function


        Public Function LustEventRemediationDataset(ByVal nEventID As Int64) As DataSet
            Dim dsRemSys As New DataSet
            'Dim drRow As DataRow
            'Dim oCol As DataColumn
            'Dim dsRel As DataRelation
            Dim strSQL As String
            Try


                strSQL = "select * from vLUSTREMEDIATIONFOREVENT "
                strSQL = strSQL & " where Event_ID = " & oLustEventInfo.ID & " order by Start_Date;"

                dsRemSys = oLustEventDB.DBGetDS(strSQL)

                Return dsRemSys
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function LustEventPMHistory() As String
            Dim dsHistory As New DataSet
            Dim strReturn As String = ""
            Dim strSQL As String
            Dim rwData As DataRow

            strSQL = "exec spGetLustEventPMHistory " & oLustEventInfo.ID
            Try
                dsHistory = oLustEventDB.DBGetDS(strSQL)
                If dsHistory.Tables(0).Rows.Count > 0 Then
                    For Each rwData In dsHistory.Tables(0).Rows
                        strReturn &= rwData("PM") & "     " & rwData("BeginDate") & "  -  " & rwData("EndDate") & vbCrLf
                    Next
                End If
                Return strReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function


        Public Function LustEventTFChecklistPath() As String
            Dim dsFilePath As New DataSet
            Dim strReturn As String = ""
            Dim strSQL As String
            Dim rwData As DataRow

            strSQL = "Select top 1 Document_LOcation + Document_Name as FilePath from tblSYS_DOCUMENT_MANAGER "
            strSQL &= "where Entity_Type = 7 and Entity_ID = " & oLustEventInfo.ID & " order by Document_ID Desc"

            Try
                dsFilePath = oLustEventDB.DBGetDS(strSQL)
                If dsFilePath.Tables(0).Rows.Count > 0 Then
                    For Each rwData In dsFilePath.Tables(0).Rows
                        strReturn &= rwData("FilePath")
                    Next
                End If
                Return strReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
#End Region
#Region "Lookup Operations"
        Public Function PopulateLustEventStatus() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("VLUSTEVENTSTATUS")
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function PopulateLustMGPTFStatus() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("VLUSTMGPTFSTATUS")
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function PopulateLustLeakPriority() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("VLUSTLEAKPRIORITY")
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function PopulateLustReleaseStatus() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("VLUSTRELEASESTATUS")
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function PopulateLustIdentifiedBy() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable2("VLUSTIDENTIFIEDBY", True)
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function


        Public Function PopulateLustReleaseLocation() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable2("VLUSTRELEASELOCATION", True)
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function


        Public Function PopulateLustReleaseExtent() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable2("VLUSTRELEASEEXTENT", True)
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function PopulateLustCause() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable2("VLUSTCAUSE", True)
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function PopulateLustProjectManager() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable("VLUSTPROJECTMANAGER")
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function

        Public Function PopulateLustSuspectedSource() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable2("VLUSTSUSPECTEDSOURCE", True)
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function PopulateERACCompanyName() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable_Company("vTEC_ERAC_COMPANYNAME", True)
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Public Function PopulateIRACCompanyName() As DataTable
            Try
                Dim dtReturn As DataTable = GetDataTable_Company("vTEC_IRAC_COMPANYNAME", True)
                Return dtReturn
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
        Private Function GetDataTable_Company(ByVal DBViewName As String, Optional ByVal bolBlankRow As Boolean = False) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String
            Try

                If bolBlankRow Then
                    strSQL = " SELECT 0 as Company_ID,'' as Company_Name UNION "
                End If

                strSQL &= " SELECT Company_ID, Company_Name FROM " & DBViewName
                strSQL &= " order by 2 "

                dsReturn = oLustEventDB.DBGetDS(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsReturn.Tables(0)
                End If
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetDataTable(ByVal strProperty As String, Optional ByVal nVal As Int64 = 0, Optional ByVal bolDistinct As Boolean = False) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String
            If bolDistinct Then
                strSQL = "SELECT DISTINCT PROPERTY_ID, PROPERTY_NAME FROM " + strProperty
            Else
                strSQL &= " SELECT * FROM " & strProperty
            End If
            If nVal <> 0 Then
                strSQL = strSQL + " WHERE PROPERTY_ID_PARENT = " + nVal.ToString()
            End If
            Try
                dsReturn = oLustEventDB.DBGetDS(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsReturn.Tables(0)
                End If
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetDataTable2(ByVal strProperty As String, Optional ByVal bolBlankRow As Boolean = False) As DataTable
            Dim dsReturn As New DataSet
            Dim dtReturn As DataTable
            Dim strSQL As String
            strSQL = ""

            If bolBlankRow Then
                strSQL = " SELECT '' as PROPERTY_NAME, 0 as PROPERTY_ID, 0 as PROPERTY_POSITION "
                strSQL &= " UNION "
            End If

            strSQL &= " SELECT PROPERTY_NAME, PROPERTY_ID, PROPERTY_POSITION FROM " & strProperty
            strSQL &= " order by 3 "

            Try
                dsReturn = oLustEventDB.DBGetDS(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    dtReturn = dsReturn.Tables(0)
                End If
                Return dtReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
#End Region
#End Region
#Region "External Event Handlers"
        Private Sub LustEventInfoChanged() Handles oLustEventInfo.LustEventInfoChanged
            RaiseEvent LustEventChanged(True)
        End Sub
        Private Sub TamplateColChanged() Handles colLustEvents.LustEventColChanged
            RaiseEvent ColChanged(True)
        End Sub

#End Region

        Public Function CheckForSingleOpenLustEvent(ByVal OwnerID As Int64, ByVal FacilityID As Int64) As Int64
            Dim dsReturn As New DataSet
            Dim nReturn As Int64
            Dim strSQL As String
            Try

                strSQL = "select Event_ID from dbo.tblTEC_EVENT where Event_Status = 624 " 'and Event_Ended is NULL "
                If OwnerID > 0 Then
                    strSQL &= " and Facility_ID in (select Facility_ID from tblREG_Facility where deleted=0 and Owner_ID = " & OwnerID & ")"
                Else
                    strSQL &= " and  Facility_ID = " & FacilityID & " and deleted=0"
                End If

                dsReturn = oLustEventDB.DBGetDS(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    If dsReturn.Tables(0).Rows.Count > 1 Then
                        nReturn = 0
                    Else
                        nReturn = dsReturn.Tables(0).Rows(0)("Event_ID")
                    End If
                Else
                    nReturn = 0
                End If
                Return nReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function


        Public Function GetSentToPMCalendarID(ByVal PMID As String, ByVal EventID As Int64) As Int64
            Dim dsReturn As New DataSet
            Dim nReturn As Int64
            Dim strSQL As String
            Try
                strSQL = "select isnull(Calendar_Info_ID, 0) as CalendarID from tblSYS_CALENDAR_CALENDAR_INFO "
                strSQL &= " where  Owning_Entity_Type = 7"
                strSQL &= " and  Owning_Entity_ID = " & EventID
                strSQL &= " and [user_ID] = '" & PMID & "'"
                strSQL &= " and Task_Description like 'Review Lust%'"

                dsReturn = oLustEventDB.DBGetDS(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    nReturn = dsReturn.Tables(0).Rows(0)("CalendarID")
                Else
                    nReturn = 0
                End If
                Return nReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function


        Public Function OpenActivities(ByVal EventID As Int64) As Int64
            Dim dsReturn As New DataSet
            Dim nReturn As Int64
            Dim strSQL As String
            Try
                strSQL = "Select count(*) as OpenCnt from dbo.tblTEC_EVENT_ACTIVITY where deleted = 0 and Closed_Date is NULL and Event_ID = " & EventID

                dsReturn = oLustEventDB.DBGetDS(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    nReturn = dsReturn.Tables(0).Rows(0)("OpenCnt")
                Else
                    nReturn = 0
                End If
                Return nReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function


        Public Function CloseOpenNFAActivities(ByVal EventID As Int64) As Int64
            Dim dsReturn As New DataSet
            Dim nReturn As Int64
            Dim strSQL As String
            Try
                ' strSQL = "Update Set Closed_Date = Getdate() from dbo.tblTEC_EVENT_ACTIVITY where Activity_Type_ID in (Select Activity_ID from tblTEC_ACTIVITY where Activity_Name like 'NFA%')deleted = 0 and Closed_Date is NULL and Event_ID = " & EventID
                strSQL = "Update dbo.tblTEC_EVENT_ACTIVITY Set Closed_Date = convert(char(10), Getdate(), 101), TECH_COMPLETED_DATE = convert(char(10), Getdate(), 101) where Activity_Type_ID in (Select Activity_ID from tblTEC_ACTIVITY where Activity_Name= 'NFA') and deleted = 0 and Closed_Date is NULL and Event_ID = " & EventID
                dsReturn = oLustEventDB.DBExeNonQuery(strSQL)
                'If dsReturn.Tables(0).Rows.Count > 0 Then
                '    nReturn = dsReturn.Tables(0).Rows(0)("OpenCnt")
                'Else
                '    nReturn = 0
                'End If
                Return nReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function

        Public Function OpenInvoices(ByVal EventID As Int64) As Int64
            Dim dsReturn As New DataSet
            Dim nReturn As Int64
            Dim strSQL As String
            Try
                strSQL = "Select Count(*) UnpaidReimbursements from tblFIN_Reimbursements Where  deleted = 0 and Financial_EventID in (Select Fin_Event_ID from tblFIN_event where  deleted = 0 and Tec_Event_ID = " & EventID & ") and (Payment_Number = 0 or Payment_Number is NULL)"

                dsReturn = oLustEventDB.DBGetDS(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    nReturn = dsReturn.Tables(0).Rows(0)("UnpaidReimbursements")
                Else
                    nReturn = 0
                End If
                Return nReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function


        Public Function OpenDocuments(ByVal EventID As Int64) As Int64
            Dim dsReturn As New DataSet
            Dim nReturn As Int64
            Dim strSQL As String
            Try
                strSQL = "Select Count(*) as OpenDocuments from tblTEC_Event_Activity_Document Where Date_Closed is Null and Date_Sent_To_Finance is Null and deleted = 0 and Event_ID = " & EventID

                dsReturn = oLustEventDB.DBGetDS(strSQL)
                If dsReturn.Tables(0).Rows.Count > 0 Then
                    nReturn = dsReturn.Tables(0).Rows(0)("OpenDocuments")
                Else
                    nReturn = 0
                End If
                Return nReturn
            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function


        Public Function MarkDueToMeCompleted_ByDesc(ByVal EventID As Int64, ByVal strDesc As String) As String
            Dim strReturn As String
            Dim strSQL As String
            Try
                strSQL = "UPDATE dbo.tblSYS_CALENDAR_CALENDAR_INFO SET COMPLETED = 1 WHERE Owning_Entity_ID = " & EventID & " AND Owning_Entity_Type = 7 AND Due_To_Me = 1 and Task_Description like '" & strDesc & "%'''"
                oLustEventDB.DBExeNonQuery(strSQL)

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function MarkToDoCompleted_ByDesc(ByVal EventID As Int64, ByVal strDesc As String) As String
            Dim strReturn As String
            Dim strSQL As String
            Try
                strSQL = "UPDATE dbo.tblSYS_CALENDAR_CALENDAR_INFO SET COMPLETED = 1 WHERE Owning_Entity_ID = " & EventID & " AND Owning_Entity_Type = 7 AND To_Do = 1 and Task_Description like '" & strDesc & "%'''"
                oLustEventDB.DBExeNonQuery(strSQL)

            Catch Ex As Exception
                If InStr(UCase(Ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Function
        Public Function GetTecOpenFinPO(ByVal nEventID As Integer) As DataSet
            Dim dsSet As DataSet
            Try
                Return oLustEventDB.DBGetTecOpenFinPO(nEventID)
            Catch ex As Exception
                If InStr(UCase(ex.Source), UCase("MUSTER.BusinessLogic")) Then MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Function
    End Class
End Namespace
