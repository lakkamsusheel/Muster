'-------------------------------------------------------------------------------
' MUSTER.Info.InspectionInfo
'   Provides the container to persist MUSTER Template state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       MNR        06/07/05    Original class definition
'  1.1       PN         06/09/05    Added Assigned date property 

' Function          Description
'
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class InspectionInfo
#Region "Public Events"
        Public Event evtInspectionInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"
        Private nInspectionID As Int64
        Private nOwnerID As Int64
        Private nFacilityID As Int64
        Private nInspectionType As Int64
        Private strAdminComments As String
        Private dtScheduledDate As Date
        Private strScheduledTime As String
        Private dtCheckListGenDate As Date
        Private dtSubmittedDate As Date
        Private nStaffID As Int64
        Private dtAssignedDate As Date
        Private strScheduledBy As String
        Private bolLetterGenerated As Boolean
        Private dtDateLetterGenerated As Date
        Private dtDatePlanned As Date
        Private dtConductInspectionOn As Date
        Private dtRescheduledDate As Date
        Private strRescheduledTime As String
        Private strInspectorComments As String
        Private dtCompleted As Date
        Private strCreatedBy As String = String.Empty
        Private dtCreatedOn As DateTime
        Private strModifiedBy As String = String.Empty
        Private dtModifiedOn As DateTime
        Private bolDeleted As Boolean
        Private strOwnersRep As String
        Private dtCAPDatesEntered As Boolean
        Private bolInspectionAccepted As Boolean
        Private dtCAEViewed As Date
        Private dtChecklistFirstSaved As Date

        Private onInspectionID As Int64
        Private onOwnerID As Int64
        Private onFacilityID As Int64
        Private onInspectionType As Int64
        Private ostrAdminComments As String
        Private odtScheduledDate As Date
        Private ostrScheduledTime As String
        Private odtCheckListGenDate As Date
        Private odtSubmittedDate As Date
        Private onStaffID As Int64
        Private odtAssignedDate As Date
        Private ostrScheduledBy As String
        Private obolLetterGenerated As Boolean
        Private odtDateLetterGenerated As Date
        Private odtDatePlanned As Date
        Private odtConductInspectionOn As Date
        Private odtRescheduledDate As Date
        Private ostrRescheduledTime As String
        Private ostrInspectorComments As String
        Private odtCompleted As Date
        Private ostrCreatedBy As String = String.Empty
        Private odtCreatedOn As DateTime
        Private ostrModifiedBy As String = String.Empty
        Private odtModifiedOn As DateTime
        Private obolDeleted As Boolean
        Private ostrOwnersRep As String
        Private odtCAPDatesEntered As Boolean
        Private obolInspectionAccepted As Boolean
        Private odtCAEViewed As Date
        Private odtChecklistFirstSaved As Date

        Private dtDataAge As DateTime
        Private nAgeThreshold As Int16 = 5
        Private bolIsDirty As Boolean = False
        Private MusterException As New MUSTER.Exceptions.MusterExceptions

        Private colInspectionChecklistMasters As MUSTER.Info.InspectionChecklistMastersCollection
        Private colInspectionResponses As MUSTER.Info.InspectionResponsesCollection
        Private colInspectionCPReadings As MUSTER.Info.InspectionCPReadingsCollection
        Private colInspectionMonitorWells As MUSTER.Info.InspectionMonitorWellsCollection
        Private colInspectionCCATs As MUSTER.Info.InspectionCCATCollection
        Private colInspectionCitations As MUSTER.Info.InspectionCitationsCollection
        Private colInspectionDisceps As MUSTER.Info.InspectionDiscrepsCollection
        Private colInspectionRectifiers As MUSTER.Info.InspectionRectifiersCollection
        Private colInspectionSketchs As MUSTER.Info.InspectionSketchsCollection
        Private colInspectionSOCs As MUSTER.Info.InspectionSOCsCollection
        Private colInspectionComments As MUSTER.Info.InspectionCommentsCollection
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            Me.Init()
            dtDataAge = Now()
        End Sub
        Sub New(ByVal InspectionID As Int64, _
            ByVal OwnerID As Int64, _
            ByVal FacilityID As Int64, _
            ByVal InspectionType As Int64, _
            ByVal AdminComments As String, _
            ByVal ScheduledDate As Date, _
            ByVal ScheduledTime As String, _
            ByVal CheckListGenDate As Date, _
            ByVal SubmittedDate As Date, _
            ByVal StaffID As Int64, _
            ByVal AssignedDate As Date, _
            ByVal ScheduledBy As String, _
            ByVal LetterGenerated As Boolean, _
            ByVal dateLetterGenerated As Date, _
            ByVal DatePlanned As Date, _
            ByVal ConductInspectionOn As Date, _
            ByVal RescheduledDate As Date, _
            ByVal RescheduledTime As String, _
            ByVal InspectorComments As String, _
            ByVal Completed As Date, _
            ByVal CreatedBy As String, _
            ByVal CreatedOn As DateTime, _
            ByVal ModifiedBy As String, _
            ByVal ModifiedOn As DateTime, _
            ByVal Deleted As Boolean, _
            ByVal OwnersRep As String, _
            ByVal capDatesEntered As Boolean, _
            ByVal inspectionAccepted As Boolean, _
            ByVal caeViewed As Date, _
            ByVal checklistFirstSaved As Date)
            onInspectionID = InspectionID
            onOwnerID = OwnerID
            onFacilityID = FacilityID
            onInspectionType = InspectionType
            ostrAdminComments = AdminComments
            odtScheduledDate = ScheduledDate
            ostrScheduledTime = ScheduledTime
            odtCheckListGenDate = CheckListGenDate
            odtSubmittedDate = SubmittedDate
            onStaffID = StaffID
            odtAssignedDate = AssignedDate
            ostrScheduledBy = ScheduledBy
            obolLetterGenerated = LetterGenerated
            odtDateLetterGenerated = dateLetterGenerated
            odtDatePlanned = DatePlanned
            odtConductInspectionOn = ConductInspectionOn
            odtRescheduledDate = RescheduledDate
            ostrRescheduledTime = RescheduledTime
            ostrInspectorComments = InspectorComments
            odtCompleted = Completed
            ostrCreatedBy = CreatedBy
            odtCreatedOn = CreatedOn
            ostrModifiedBy = ModifiedBy
            odtModifiedOn = ModifiedOn
            obolDeleted = Deleted
            ostrOwnersRep = OwnersRep
            odtCAPDatesEntered = capDatesEntered
            obolInspectionAccepted = inspectionAccepted
            odtCAEViewed = caeViewed
            odtChecklistFirstSaved = checklistFirstSaved
            dtDataAge = Now()
            InitCollection()
            Me.Reset()
        End Sub
        Sub New(ByVal dr As DataRow)
            Try
                onInspectionID = dr.Item("INSPECTION_ID")
                onOwnerID = dr.Item("OWNER_ID")
                onFacilityID = dr.Item("FACILITY_ID")
                onInspectionType = dr.Item("INSPECTION_TYPE")
                ostrAdminComments = IIf(dr.Item("ADMIN_COMMENTS") Is DBNull.Value, String.Empty, dr.Item("ADMIN_COMMENTS"))
                odtScheduledDate = IIf(dr.Item("SCHEDULED_DATE") Is DBNull.Value, CDate(""), dr.Item("SCHEDULED_DATE"))
                ostrScheduledTime = IIf(dr.Item("SCHEDULED_TIME") Is DBNull.Value, String.Empty, dr.Item("SCHEDULED_TIME"))
                odtCheckListGenDate = IIf(dr.Item("CHECKLIST_GENERATION_DATE") Is DBNull.Value, CDate("01/01/0001"), dr.Item("CHECKLIST_GENERATION_DATE"))
                odtSubmittedDate = IIf(dr.Item("SUBMITTED_DATE") Is DBNull.Value, CDate("01/01/0001"), dr.Item("SUBMITTED_DATE"))
                onStaffID = IIf(dr.Item("STAFF_ID") Is DBNull.Value, 0, dr.Item("STAFF_ID"))
                odtAssignedDate = IIf(dr.Item("ASSIGNED_DATE") Is DBNull.Value, CDate("01/01/0001"), dr.Item("ASSIGNED_DATE"))
                ostrScheduledBy = IIf(dr.Item("SCHEDULED_BY") Is DBNull.Value, String.Empty, dr.Item("SCHEDULED_BY"))
                obolLetterGenerated = IIf(dr.Item("LETTER_GENERATED") Is DBNull.Value, False, dr.Item("LETTER_GENERATED"))
                odtDateLetterGenerated = IIf(dr.Item("DATE_GENERATED") Is DBNull.Value, CDate("01/01/0001"), dr.Item("DATE_GENERATED"))
                odtDatePlanned = IIf(dr.Item("DATE_PLANNED") Is DBNull.Value, CDate("01/01/0001"), dr.Item("DATE_PLANNED"))
                odtConductInspectionOn = IIf(dr.Item("CONDUCT_INSPECTION_ON") Is DBNull.Value, CDate("01/01/0001"), dr.Item("CONDUCT_INSPECTION_ON"))
                odtRescheduledDate = IIf(dr.Item("RESCHEDULED_DATE") Is DBNull.Value, CDate("01/01/0001"), dr.Item("RESCHEDULED_DATE"))
                ostrRescheduledTime = IIf(dr.Item("RESCHEDULED_TIME") Is DBNull.Value, String.Empty, dr.Item("RESCHEDULED_TIME"))
                ostrInspectorComments = IIf(dr.Item("INSPECTOR_COMMENTS") Is DBNull.Value, String.Empty, dr.Item("INSPECTOR_COMMENTS"))
                odtCompleted = IIf(dr.Item("COMPLETED") Is DBNull.Value, CDate("01/01/0001"), dr.Item("COMPLETED"))
                ostrCreatedBy = IIf(dr.Item("CREATED_BY") Is DBNull.Value, String.Empty, dr.Item("CREATED_BY"))
                odtCreatedOn = IIf(dr.Item("DATE_CREATED") Is DBNull.Value, CDate("01/01/0001"), dr.Item("DATE_CREATED"))
                ostrModifiedBy = IIf(dr.Item("LAST_EDITED_BY") Is DBNull.Value, String.Empty, dr.Item("LAST_EDITED_BY"))
                odtModifiedOn = IIf(dr.Item("DATE_LAST_EDITED") Is DBNull.Value, CDate("01/01/0001"), dr.Item("DATE_LAST_EDITED"))
                obolDeleted = IIf(dr.Item("DELETED") Is DBNull.Value, False, dr.Item("DELETED"))
                ostrOwnersRep = IIf(dr.Item("OWNERS_REP") Is DBNull.Value, String.Empty, dr.Item("OWNERS_REP"))
                odtCAPDatesEntered = IIf(dr.Item("CAP_DATES_ENTERED") Is DBNull.Value, False, dr.Item("CAP_DATES_ENTERED"))
                obolInspectionAccepted = IIf(dr.Item("INSPECTION_ACCPETED") Is DBNull.Value, False, dr.Item("INSPECTION_ACCPETED"))
                odtCAEViewed = IIf(dr.Item("CAE_VIEWED") Is DBNull.Value, CDate("01/01/0001"), dr.Item("CAE_VIEWED"))
                odtChecklistFirstSaved = IIf(dr.Item("CHECKLIST_FIRST_SAVED") Is DBNull.Value, CDate("01/01/0001"), dr.Item("CHECKLIST_FIRST_SAVED"))
                dtDataAge = Now()
                InitCollection()
                Me.Reset()
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "Exposed Operations"


        Public Sub Reset()
            If nInspectionID >= 0 Then
                nInspectionID = onInspectionID
            End If
            nOwnerID = onOwnerID
            nFacilityID = onFacilityID
            nInspectionType = onInspectionType
            strAdminComments = ostrAdminComments
            dtScheduledDate = odtScheduledDate
            strScheduledTime = ostrScheduledTime
            dtCheckListGenDate = odtCheckListGenDate
            dtSubmittedDate = odtSubmittedDate
            nStaffID = onStaffID
            dtAssignedDate = odtAssignedDate
            strScheduledBy = ostrScheduledBy
            bolLetterGenerated = obolLetterGenerated
            dtDateLetterGenerated = odtDateLetterGenerated
            dtDatePlanned = odtDatePlanned
            dtConductInspectionOn = odtConductInspectionOn
            dtRescheduledDate = odtRescheduledDate
            strRescheduledTime = ostrRescheduledTime
            strInspectorComments = ostrInspectorComments
            dtCompleted = odtCompleted
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            bolDeleted = obolDeleted
            strOwnersRep = ostrOwnersRep
            dtCAPDatesEntered = odtCAPDatesEntered
            bolInspectionAccepted = obolInspectionAccepted
            dtCAEViewed = odtCAEViewed
            dtChecklistFirstSaved = odtChecklistFirstSaved
            bolIsDirty = False
            RaiseEvent evtInspectionInfoChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            onInspectionID = nInspectionID
            onOwnerID = nOwnerID
            onFacilityID = nFacilityID
            onInspectionType = nInspectionType
            ostrAdminComments = strAdminComments
            odtScheduledDate = dtScheduledDate
            ostrScheduledTime = strScheduledTime
            odtCheckListGenDate = dtCheckListGenDate
            odtSubmittedDate = dtSubmittedDate
            onStaffID = nStaffID
            odtAssignedDate = dtAssignedDate
            ostrScheduledBy = strScheduledBy
            obolLetterGenerated = bolLetterGenerated
            odtDateLetterGenerated = dtDateLetterGenerated
            odtDatePlanned = dtDatePlanned
            odtConductInspectionOn = dtConductInspectionOn
            odtRescheduledDate = dtRescheduledDate
            ostrRescheduledTime = strRescheduledTime
            ostrInspectorComments = strInspectorComments
            odtCompleted = dtCompleted
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn
            obolDeleted = bolDeleted
            ostrOwnersRep = strOwnersRep
            odtCAPDatesEntered = dtCAPDatesEntered
            obolInspectionAccepted = bolInspectionAccepted
            odtCAEViewed = dtCAEViewed
            odtChecklistFirstSaved = dtChecklistFirstSaved
            bolIsDirty = False
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            Dim obolIsDirty As Boolean = bolIsDirty

            bolIsDirty = (nOwnerID <> onOwnerID) Or _
            (nFacilityID <> onFacilityID) Or _
            (nInspectionType <> onInspectionType) Or _
            (strAdminComments <> ostrAdminComments) Or _
            (dtScheduledDate <> odtScheduledDate) Or _
            (strScheduledTime <> ostrScheduledTime) Or _
            (dtCheckListGenDate <> odtCheckListGenDate) Or _
            (dtSubmittedDate <> odtSubmittedDate) Or _
            (nStaffID <> onStaffID) Or _
            (dtAssignedDate <> odtAssignedDate) Or _
            (strScheduledBy <> ostrScheduledBy) Or _
            (bolLetterGenerated <> obolLetterGenerated) Or _
            (dtDateLetterGenerated <> odtDateLetterGenerated) Or _
            (dtDatePlanned <> odtDatePlanned) Or _
            (dtConductInspectionOn <> odtConductInspectionOn) Or _
            (dtRescheduledDate <> odtRescheduledDate) Or _
            (strRescheduledTime <> ostrRescheduledTime) Or _
            (strInspectorComments <> ostrInspectorComments) Or _
            (dtCompleted <> odtCompleted) Or _
            (strCreatedBy <> ostrCreatedBy) Or _
            (dtCreatedOn <> odtCreatedOn) Or _
            (strModifiedBy <> ostrModifiedBy) Or _
            (dtModifiedOn <> odtModifiedOn) Or _
            (bolDeleted <> obolDeleted) Or _
            (strOwnersRep <> ostrOwnersRep) Or _
            (dtCAPDatesEntered <> odtCAPDatesEntered) Or _
            (bolInspectionAccepted <> obolInspectionAccepted) Or _
            (dtCAEViewed <> odtCAEViewed) Or _
            (dtChecklistFirstSaved <> odtChecklistFirstSaved)

            If obolIsDirty <> bolIsDirty Then
                RaiseEvent evtInspectionInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            onInspectionID = 0
            onOwnerID = 0
            onFacilityID = 0
            onInspectionType = 0
            ostrAdminComments = String.Empty
            odtScheduledDate = CDate("01/01/0001")
            ostrScheduledTime = String.Empty
            odtCheckListGenDate = CDate("01/01/0001")
            odtSubmittedDate = CDate("01/01/0001")
            onStaffID = 0
            odtAssignedDate = CDate("01/01/0001")
            ostrScheduledBy = String.Empty
            obolLetterGenerated = False
            odtDateLetterGenerated = CDate("01/01/0001")
            odtDatePlanned = CDate("01/01/0001")
            odtConductInspectionOn = CDate("01/01/0001")
            odtRescheduledDate = CDate("01/01/0001")
            ostrRescheduledTime = String.Empty
            ostrInspectorComments = String.Empty
            odtCompleted = CDate("01/01/0001")
            ostrCreatedBy = String.Empty
            odtCreatedOn = CDate("01/01/0001")
            ostrModifiedBy = String.Empty
            odtModifiedOn = CDate("01/01/0001")
            obolDeleted = False
            ostrOwnersRep = String.Empty
            odtCAPDatesEntered = False
            obolInspectionAccepted = False
            odtCAEViewed = CDate("01/01/0001")
            odtChecklistFirstSaved = CDate("01/01/0001")
            InitCollection()
            Me.Reset()
        End Sub
        Private Sub InitCollection()
            colInspectionResponses = New MUSTER.Info.InspectionResponsesCollection
            colInspectionCPReadings = New MUSTER.Info.InspectionCPReadingsCollection
            colInspectionMonitorWells = New MUSTER.Info.InspectionMonitorWellsCollection
            colInspectionCCATs = New MUSTER.Info.InspectionCCATCollection
            colInspectionCitations = New MUSTER.Info.InspectionCitationsCollection
            colInspectionDisceps = New MUSTER.Info.InspectionDiscrepsCollection
            colInspectionRectifiers = New MUSTER.Info.InspectionRectifiersCollection
            colInspectionSketchs = New MUSTER.Info.InspectionSketchsCollection
            colInspectionSOCs = New MUSTER.Info.InspectionSOCsCollection
            colInspectionComments = New MUSTER.Info.InspectionCommentsCollection
            colInspectionChecklistMasters = New MUSTER.Info.InspectionChecklistMastersCollection
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Int64
            Get
                Return nInspectionID
            End Get
            Set(ByVal Value As Int64)
                nInspectionID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property OwnerID() As Int64
            Get
                Return nOwnerID
            End Get
            Set(ByVal Value As Int64)
                nOwnerID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property FacilityID() As Int64
            Get
                Return nFacilityID
            End Get
            Set(ByVal Value As Int64)
                nFacilityID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property InspectionType() As Int64
            Get
                Return nInspectionType
            End Get
            Set(ByVal Value As Int64)
                nInspectionType = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property AdminComments() As String
            Get
                Return strAdminComments
            End Get
            Set(ByVal Value As String)
                strAdminComments = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ScheduledDate() As Date
            Get
                Return dtScheduledDate
            End Get
            Set(ByVal Value As Date)
                dtScheduledDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ScheduledTime() As String
            Get
                Return strScheduledTime
            End Get
            Set(ByVal Value As String)
                strScheduledTime = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CheckListGenDate() As Date
            Get
                Return dtCheckListGenDate
            End Get
            Set(ByVal Value As Date)
                dtCheckListGenDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property SubmittedDate() As Date
            Get
                Return dtSubmittedDate
            End Get
            Set(ByVal Value As Date)
                dtSubmittedDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property StaffID() As Int64
            Get
                Return nStaffID
            End Get
            Set(ByVal Value As Int64)
                nStaffID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property AssignedDate() As Date
            Get
                Return dtAssignedDate
            End Get
            Set(ByVal Value As Date)
                dtAssignedDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ScheduledBy() As String
            Get
                Return strScheduledBy
            End Get
            Set(ByVal Value As String)
                strScheduledBy = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LetterGenerated() As Boolean
            Get
                Return bolLetterGenerated
            End Get
            Set(ByVal Value As Boolean)
                bolLetterGenerated = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property DateLetterGenerated() As Date
            Get
                Return dtDateLetterGenerated
            End Get
            Set(ByVal Value As Date)
                dtDateLetterGenerated = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property DatePlanned() As Date
            Get
                Return dtDatePlanned
            End Get
            Set(ByVal Value As Date)
                dtDatePlanned = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ConductInspectionOn() As Date
            Get
                Return dtConductInspectionOn
            End Get
            Set(ByVal Value As Date)
                dtConductInspectionOn = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property RescheduledDate() As Date
            Get
                Return dtRescheduledDate
            End Get
            Set(ByVal Value As Date)
                dtRescheduledDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property RescheduledTime() As String
            Get
                Return strRescheduledTime
            End Get
            Set(ByVal Value As String)
                strRescheduledTime = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property InspectorComments() As String
            Get
                Return strInspectorComments
            End Get
            Set(ByVal Value As String)
                strInspectorComments = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Completed() As Date
            Get
                Return dtCompleted
            End Get
            Set(ByVal Value As Date)
                dtCompleted = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property OwnersRep() As String
            Get
                Return strOwnersRep
            End Get

            Set(ByVal Value As String)
                strOwnersRep = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CAPDatesEntered() As Boolean
            Get
                Return dtCAPDatesEntered
            End Get
            Set(ByVal Value As Boolean)
                dtCAPDatesEntered = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property InspectionAccepted() As Boolean
            Get
                Return bolInspectionAccepted
            End Get
            Set(ByVal Value As Boolean)
                bolInspectionAccepted = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CAEViewed() As Date
            Get
                Return dtCAEViewed
            End Get
            Set(ByVal Value As Date)
                dtCAEViewed = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ChecklistFirstSaved() As Date
            Get
                Return dtChecklistFirstSaved
            End Get
            Set(ByVal Value As Date)
                dtChecklistFirstSaved = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return bolDeleted
            End Get
            Set(ByVal value As Boolean)
                bolDeleted = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return bolIsDirty
            End Get

            Set(ByVal value As Boolean)
                bolIsDirty = value
                RaiseEvent evtInspectionInfoChanged(bolIsDirty)
            End Set
        End Property
        Public Property CreatedBy() As String
            Get
                Return strCreatedBy
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
        Public Property ModifiedBy() As String
            Get
                Return strModifiedBy
            End Get

            Set(ByVal Value As String)
                strModifiedBy = Value
            End Set
        End Property
        Public Property ModifiedOn() As Date
            Get
                Return dtModifiedOn
            End Get

            Set(ByVal Value As Date)
                dtModifiedOn = Value
            End Set
        End Property
        Public ReadOnly Property IsAgedData() As Boolean
            Get
                Return IIf(DateDiff(DateInterval.Minute, dtDataAge, Now()) >= nAgeThreshold, True, False)
            End Get
        End Property
        Public Property ChecklistMasterCollection() As MUSTER.Info.InspectionChecklistMastersCollection
            Get
                Return colInspectionChecklistMasters
            End Get
            Set(ByVal Value As MUSTER.Info.InspectionChecklistMastersCollection)
                colInspectionChecklistMasters = Value
            End Set
        End Property
        Public Property ResponsesCollection() As MUSTER.Info.InspectionResponsesCollection
            Get
                Return colInspectionResponses
            End Get
            Set(ByVal Value As MUSTER.Info.InspectionResponsesCollection)
                colInspectionResponses = Value
            End Set
        End Property
        Public Property CPReadingsCollection() As MUSTER.Info.InspectionCPReadingsCollection
            Get
                Return colInspectionCPReadings
            End Get
            Set(ByVal Value As MUSTER.Info.InspectionCPReadingsCollection)
                colInspectionCPReadings = Value
            End Set
        End Property
        Public Property MonitorWellsCollection() As MUSTER.Info.InspectionMonitorWellsCollection
            Get
                Return colInspectionMonitorWells
            End Get
            Set(ByVal Value As MUSTER.Info.InspectionMonitorWellsCollection)
                colInspectionMonitorWells = Value
            End Set
        End Property
        Public Property CCATsCollection() As MUSTER.Info.InspectionCCATCollection
            Get
                Return colInspectionCCATs
            End Get
            Set(ByVal Value As MUSTER.Info.InspectionCCATCollection)
                colInspectionCCATs = Value
            End Set
        End Property
        Public Property CitationsCollection() As MUSTER.Info.InspectionCitationsCollection
            Get
                Return colInspectionCitations
            End Get
            Set(ByVal Value As MUSTER.Info.InspectionCitationsCollection)
                colInspectionCitations = Value
            End Set
        End Property
        Public Property DiscrepsCollection() As MUSTER.Info.InspectionDiscrepsCollection
            Get
                Return colInspectionDisceps
            End Get
            Set(ByVal Value As MUSTER.Info.InspectionDiscrepsCollection)
                colInspectionDisceps = Value
            End Set
        End Property
        Public Property RectifiersCollection() As MUSTER.Info.InspectionRectifiersCollection
            Get
                Return colInspectionRectifiers
            End Get
            Set(ByVal Value As MUSTER.Info.InspectionRectifiersCollection)
                colInspectionRectifiers = Value
            End Set
        End Property
        Public Property SketchsCollection() As MUSTER.Info.InspectionSketchsCollection
            Get
                Return colInspectionSketchs
            End Get
            Set(ByVal Value As MUSTER.Info.InspectionSketchsCollection)
                colInspectionSketchs = Value
            End Set
        End Property
        Public Property SOCsCollection() As MUSTER.Info.InspectionSOCsCollection
            Get
                Return colInspectionSOCs
            End Get
            Set(ByVal Value As MUSTER.Info.InspectionSOCsCollection)
                colInspectionSOCs = Value
            End Set
        End Property
        Public Property InspectionCommentsCollection() As MUSTER.Info.InspectionCommentsCollection
            Get
                Return colInspectionComments
            End Get
            Set(ByVal Value As MUSTER.Info.InspectionCommentsCollection)
                colInspectionComments = Value
            End Set
        End Property
#End Region
#Region "Protected Operations"
        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
#End Region
    End Class
End Namespace
