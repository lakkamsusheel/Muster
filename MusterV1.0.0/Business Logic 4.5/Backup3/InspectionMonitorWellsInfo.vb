'-------------------------------------------------------------------------------
' MUSTER.Info.InspectionMonitorWellsInfo
'   Provides the container to persist MUSTER InspectionMonitorWells state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0      MNR         6/15/05     Original class definition
'
' Function          Description
' New()             Instantiates an empty InspectionMonitorWellsInfo object
' New(Deleted, CreatedBy, CreatedOn, ModifiedBy, LastEdited, OwnerL2CSnippet)
'                   Instantiates a populated InspectionMonitorWellsInfo object
' New(dr)           Instantiates a populated InspectionMonitorWellsInfo object taking member state
'                   from the datarow provided
' Reset()           Sets the object state to the original state when loaded from or
'                   last saved to the repository
'
' NOTE: This file to be used as InspectionMonitorWells to build other objects.
'       Replace keyword "InspectionMonitorWells" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class InspectionMonitorWellsInfo
#Region "Public Events"
        Public Event evtInspectionMonitorWellsInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"
        Private nInsMonWellID As Int64
        Private nInspectionID As Int64
        Private nQuestionID As Int64
        Private nLineNum As Int64
        Private bolTankLine As Boolean
        Private nWellNumber As Int64
        Private strWellDepth As String
        Private strDepthToWater As String
        Private strDepthToSlots As String
        Private nSurfaceSealed As Int64
        Private nWellCaps As Int64
        Private strInspectorsObservations As String
        Private bolDeleted As Boolean
        Private strCreatedBy As String = String.Empty
        Private dtCreatedOn As DateTime
        Private strModifiedBy As String = String.Empty
        Private dtModifiedOn As DateTime

        Private onInsMonWellID As Int64
        Private onInspectionID As Int64
        Private onQuestionID As Int64
        Private onLineNum As Int64
        Private obolTankLine As Boolean
        Private onWellNumber As Int64
        Private ostrWellDepth As String
        Private ostrDepthToWater As String
        Private ostrDepthToSlots As String
        Private onSurfaceSealed As Int64
        Private onWellCaps As Int64
        Private ostrInspectorsObservations As String
        Private obolDeleted As Boolean
        Private ostrCreatedBy As String = String.Empty
        Private odtCreatedOn As DateTime
        Private ostrModifiedBy As String = String.Empty
        Private odtModifiedOn As DateTime

        Private bolIsDirty As Boolean = False
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            Me.Init()
        End Sub
        Sub New(ByVal id As Int64, _
        ByVal InspectionID As Int64, _
        ByVal QuestionID As Int64, _
        ByVal TankLine As Boolean, _
        ByVal WellNumber As Int64, _
        ByVal WellDepth As String, _
        ByVal DepthToWater As String, _
        ByVal DepthToSlots As String, _
        ByVal SurfaceSealed As Int64, _
        ByVal WellCaps As Int64, _
        ByVal inspectorsObservations As String, _
        ByVal deleted As Boolean, _
        ByVal createdBy As String, _
        ByVal createdOn As Date, _
        ByVal modifiedBy As String, _
        ByVal modifiedOn As Date, _
        ByVal lineNum As Int64)
            onInsMonWellID = id
            onInspectionID = InspectionID
            onQuestionID = QuestionID
            onLineNum = lineNum
            obolTankLine = TankLine
            onWellNumber = WellNumber
            ostrWellDepth = WellDepth
            ostrDepthToWater = DepthToWater
            ostrDepthToSlots = DepthToSlots
            onSurfaceSealed = SurfaceSealed
            onWellCaps = WellCaps
            ostrInspectorsObservations = inspectorsObservations
            obolDeleted = deleted
            ostrCreatedBy = createdBy
            odtCreatedOn = createdOn
            ostrModifiedBy = modifiedBy
            odtModifiedOn = modifiedOn
            Me.Reset()
        End Sub
        Sub New(ByVal drInspectionMonitorWells As DataRow)
            Try
                onInsMonWellID = drInspectionMonitorWells.Item("INS_MON_WELL_ID")
                onInspectionID = IIf(drInspectionMonitorWells.Item("INSPECTION_ID") Is DBNull.Value, 0, drInspectionMonitorWells.Item("INSPECTION_ID"))
                onQuestionID = IIf(drInspectionMonitorWells.Item("QUESTION_ID") Is DBNull.Value, 0, drInspectionMonitorWells.Item("QUESTION_ID"))
                onLineNum = IIf(drInspectionMonitorWells.Item("LINE_NUMBER") Is DBNull.Value, 0, drInspectionMonitorWells.Item("LINE_NUMBER"))
                obolTankLine = IIf(drInspectionMonitorWells.Item("TANK_LINE") Is DBNull.Value, False, drInspectionMonitorWells.Item("TANK_LINE"))
                onWellNumber = IIf(drInspectionMonitorWells.Item("WELL_NUMBER") Is DBNull.Value, 0, drInspectionMonitorWells.Item("WELL_NUMBER"))
                ostrWellDepth = IIf(drInspectionMonitorWells.Item("WELL_DEPTH") Is DBNull.Value, String.Empty, drInspectionMonitorWells.Item("WELL_DEPTH"))
                ostrDepthToWater = IIf(drInspectionMonitorWells.Item("DEPTH_TO_WATER") Is DBNull.Value, String.Empty, drInspectionMonitorWells.Item("DEPTH_TO_WATER"))
                ostrDepthToSlots = IIf(drInspectionMonitorWells.Item("DEPTH_TO_SLOTS") Is DBNull.Value, String.Empty, drInspectionMonitorWells.Item("DEPTH_TO_SLOTS"))
                onSurfaceSealed = IIf(drInspectionMonitorWells.Item("SURFACE_SEALED") Is DBNull.Value, -1, drInspectionMonitorWells.Item("SURFACE_SEALED"))
                onWellCaps = IIf(drInspectionMonitorWells.Item("WELL_CAPS") Is DBNull.Value, -1, drInspectionMonitorWells.Item("WELL_CAPS"))
                ostrInspectorsObservations = IIf(drInspectionMonitorWells.Item("INSPECTORS_OBSERVTIONS") Is DBNull.Value, String.Empty, drInspectionMonitorWells.Item("INSPECTORS_OBSERVTIONS"))
                obolDeleted = IIf(drInspectionMonitorWells.Item("DELETED") Is DBNull.Value, False, drInspectionMonitorWells.Item("DELETED"))
                ostrCreatedBy = IIf(drInspectionMonitorWells.Item("CREATED_BY") Is DBNull.Value, String.Empty, drInspectionMonitorWells.Item("CREATED_BY"))
                odtCreatedOn = IIf(drInspectionMonitorWells.Item("DATE_CREATED") Is DBNull.Value, CDate("01/01/0001"), drInspectionMonitorWells.Item("DATE_CREATED"))
                ostrModifiedBy = IIf(drInspectionMonitorWells.Item("LAST_EDITED_BY") Is DBNull.Value, String.Empty, drInspectionMonitorWells.Item("LAST_EDITED_BY"))
                odtModifiedOn = IIf(drInspectionMonitorWells.Item("DATE_LAST_EDITED") Is DBNull.Value, CDate("01/01/0001"), drInspectionMonitorWells.Item("DATE_LAST_EDITED"))
                Me.Reset()
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            If nInsMonWellID >= 0 Then
                nInsMonWellID = onInsMonWellID
            End If
            nInspectionID = onInspectionID
            nQuestionID = onQuestionID
            nLineNum = onLineNum
            bolTankLine = obolTankLine
            nWellNumber = onWellNumber
            strWellDepth = ostrWellDepth
            strDepthToWater = ostrDepthToWater
            strDepthToSlots = ostrDepthToSlots
            nSurfaceSealed = onSurfaceSealed
            nWellCaps = onWellCaps
            strInspectorsObservations = ostrInspectorsObservations
            bolDeleted = obolDeleted
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            bolIsDirty = False
            RaiseEvent evtInspectionMonitorWellsInfoChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            onInsMonWellID = nInsMonWellID
            onInspectionID = nInspectionID
            onQuestionID = nQuestionID
            onLineNum = nLineNum
            obolTankLine = bolTankLine
            onWellNumber = nWellNumber
            ostrWellDepth = strWellDepth
            ostrDepthToWater = strDepthToWater
            ostrDepthToSlots = strDepthToSlots
            onSurfaceSealed = nSurfaceSealed
            onWellCaps = nWellCaps
            ostrInspectorsObservations = strInspectorsObservations
            obolDeleted = bolDeleted
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn
            bolIsDirty = False
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            Dim obolIsDirty As Boolean = bolIsDirty

            bolIsDirty = (onInspectionID <> nInspectionID) Or _
            (onQuestionID <> nQuestionID) Or _
            (nLineNum <> onLineNum) Or _
            (obolTankLine <> bolTankLine) Or _
            (onWellNumber <> nWellNumber) Or _
            (ostrWellDepth <> strWellDepth) Or _
            (ostrDepthToWater <> strDepthToWater) Or _
            (ostrDepthToSlots <> strDepthToSlots) Or _
            (onSurfaceSealed <> nSurfaceSealed) Or _
            (onWellCaps <> nWellCaps) Or _
            (ostrInspectorsObservations <> strInspectorsObservations) Or _
            (obolDeleted <> bolDeleted) Or _
            (strCreatedBy <> strCreatedBy) Or _
            (dtCreatedOn <> dtCreatedOn) Or _
            (strModifiedBy <> strModifiedBy) Or _
            (dtModifiedOn <> dtModifiedOn)

            If obolIsDirty <> bolIsDirty Then
                RaiseEvent evtInspectionMonitorWellsInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            onInsMonWellID = 0
            onInspectionID = 0
            onQuestionID = 0
            onLineNum = 0
            obolTankLine = Nothing
            onWellNumber = 0
            ostrWellDepth = String.Empty
            ostrDepthToWater = String.Empty
            ostrDepthToSlots = String.Empty
            onSurfaceSealed = -1
            onWellCaps = -1
            ostrInspectorsObservations = String.Empty
            obolDeleted = False
            ostrCreatedBy = String.Empty
            odtCreatedOn = CDate("01/01/0001")
            ostrModifiedBy = String.Empty
            odtModifiedOn = CDate("01/01/0001")
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Int64
            Get
                Return nInsMonWellID
            End Get
            Set(ByVal Value As Int64)
                nInsMonWellID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property InspectionID() As Int64
            Get
                Return nInspectionID
            End Get
            Set(ByVal Value As Int64)
                nInspectionID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property QuestionID() As Int64
            Get
                Return nQuestionID
            End Get
            Set(ByVal Value As Int64)
                nQuestionID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LineNumber() As Int64
            Get
                Return nLineNum
            End Get
            Set(ByVal Value As Int64)
                nLineNum = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property TankLine() As Boolean
            Get
                Return bolTankLine
            End Get
            Set(ByVal Value As Boolean)
                bolTankLine = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property WellNumber() As Int64
            Get
                Return nWellNumber
            End Get
            Set(ByVal Value As Int64)
                nWellNumber = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property WellDepth() As String
            Get
                Return strWellDepth
            End Get
            Set(ByVal Value As String)
                strWellDepth = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property DepthToWater() As String
            Get
                Return strDepthToWater
            End Get
            Set(ByVal Value As String)
                strDepthToWater = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property DepthToSlots() As String
            Get
                Return strDepthToSlots
            End Get
            Set(ByVal Value As String)
                strDepthToSlots = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property SurfaceSealed() As Int64
            Get
                Return nSurfaceSealed
            End Get
            Set(ByVal Value As Int64)
                nSurfaceSealed = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property WellCaps() As Int64
            Get
                Return nWellCaps
            End Get
            Set(ByVal Value As Int64)
                nWellCaps = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property InspectorsObservations() As String
            Get
                Return strInspectorsObservations
            End Get
            Set(ByVal Value As String)
                strInspectorsObservations = Value
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
                RaiseEvent evtInspectionMonitorWellsInfoChanged(bolIsDirty)
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
#End Region
#Region "Protected Operations"
        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
#End Region
    End Class
End Namespace
