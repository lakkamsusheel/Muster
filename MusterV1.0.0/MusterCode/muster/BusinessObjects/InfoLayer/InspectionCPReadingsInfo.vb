'-------------------------------------------------------------------------------
' MUSTER.Info.InspectionCPReadingsInfo
'   Provides the container to persist MUSTER InspectionCPReadings state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0      MNR         6/10/05     Original class definition
'
' Function          Description
' New()             Instantiates an empty InspectionCPReadingsInfo object
' New(Deleted, CreatedBy, CreatedOn, ModifiedBy, LastEdited, OwnerL2CSnippet)
'                   Instantiates a populated InspectionCPReadingsInfo object
' New(dr)           Instantiates a populated InspectionCPReadingsInfo object taking member state
'                   from the datarow provided
' Reset()           Sets the object state to the original state when loaded from or
'                   last saved to the repository
'
' NOTE: This file to be used as InspectionCPReadings to build other objects.
'       Replace keyword "InspectionCPReadings" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class InspectionCPReadingsInfo
#Region "Public Events"
        Public Event evtInspectionCPReadingsInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"
        Private nInsCPReadID As Integer
        Private nInspectionID As Integer
        Private nQuestionID As Integer
        Private nLineNum As Integer
        Private nTankPipeID As Integer
        Private nTankPipeIndex As Integer
        Private nTankPipeEntityID As Integer
        'Private bolTankDispenser As Boolean
        'Private bolGalvanic As Boolean
        'Private nGalvanicICResponse As Boolean
        Private strContactPoint As String
        Private strLocalReferCellPlacement As String
        Private strLocalOn As String
        Private strRemoteOff As String
        Private nPassFailIncon As Integer
        Private bolDeleted As Boolean
        Private strCreatedBy As String = String.Empty
        Private dtCreatedOn As DateTime
        Private strModifiedBy As String = String.Empty
        Private dtModifiedOn As DateTime
        Private bolRemoteReferCellPlacement As Boolean
        Private bolGalvanicIC As Boolean
        Private nGalvanicICResponse As Integer
        Private bolTestedByInspector As Boolean
        Private bolTestedByInspectorResponse As Boolean

        Private onInsCPReadID As Integer
        Private onInspectionID As Integer
        Private onQuestionID As Integer
        Private onLineNum As Integer
        Private onTankPipeID As Integer
        Private onTankPipeIndex As Integer
        Private onTankPipeEntityID As Integer
        'Private obolTankDispenser As Boolean
        'Private obolGalvanic As Boolean
        'Private onGalvanicICResponse As Boolean
        Private ostrContactPoint As String
        Private ostrLocalReferCellPlacement As String
        Private ostrLocalOn As String
        Private ostrRemoteOff As String
        Private onPassFailIncon As Integer
        Private obolDeleted As Boolean
        Private ostrCreatedBy As String = String.Empty
        Private odtCreatedOn As DateTime
        Private ostrModifiedBy As String = String.Empty
        Private odtModifiedOn As DateTime
        Private obolRemoteReferCellPlacement As Boolean
        Private obolGalvanicIC As Boolean
        Private onGalvanicICResponse As Integer
        Private obolTestedByInspector As Boolean
        Private obolTestedByInspectorResponse As Boolean

        Private bolIsDirty As Boolean = False
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            Me.Init()
        End Sub
        Sub New(ByVal id As Integer, _
        ByVal InspectionID As Integer, _
        ByVal QuestionID As Integer, _
        ByVal TankPipeID As Integer, _
        ByVal TankPipeIndex As Integer, _
        ByVal TankPipeEntityID As Integer, _
        ByVal ContactPoint As String, _
        ByVal LocalReferCellPlacement As String, _
        ByVal LocalOn As String, _
        ByVal RemoteOff As String, _
        ByVal PassFailIncon As Integer, _
        ByVal deleted As Boolean, _
        ByVal createdBy As String, _
        ByVal createdOn As DateTime, _
        ByVal modifiedBy As String, _
        ByVal modifiedOn As DateTime, _
        ByVal lineNum As Integer, _
        ByVal remoteReferCellPlacement As Boolean, _
        ByVal galvanicIC As Boolean, _
        ByVal galvanicICResp As Integer, _
        ByVal testedByInspector As Boolean, _
        ByVal testedByInspectorResp As Boolean)
            onInsCPReadID = id
            onInspectionID = InspectionID
            onQuestionID = QuestionID
            onLineNum = lineNum
            onTankPipeID = TankPipeID
            onTankPipeIndex = TankPipeIndex
            onTankPipeEntityID = TankPipeEntityID
            'obolTankDispenser = TankDispenser
            'obolGalvanic = Galvanic
            'onGalvanicICResponse = ImpressedCurrent
            ostrContactPoint = ContactPoint
            ostrLocalReferCellPlacement = LocalReferCellPlacement
            ostrLocalOn = LocalOn
            ostrRemoteOff = RemoteOff
            onPassFailIncon = PassFailIncon
            obolDeleted = deleted
            ostrCreatedBy = createdBy
            odtCreatedOn = createdOn
            ostrModifiedBy = modifiedBy
            odtModifiedOn = modifiedOn
            obolRemoteReferCellPlacement = remoteReferCellPlacement
            obolGalvanicIC = galvanicIC
            onGalvanicICResponse = galvanicICResp
            obolTestedByInspector = testedByInspector
            obolTestedByInspectorResponse = testedByInspectorResp
            Me.Reset()
        End Sub
        Sub New(ByVal drInspectionCPReadings As DataRow)
            Try
                onInsCPReadID = drInspectionCPReadings.Item("INS_CP_READ_ID")
                onInspectionID = IIf(drInspectionCPReadings.Item("INSPECTION_ID") Is DBNull.Value, 0, drInspectionCPReadings.Item("INSPECTION_ID"))
                onQuestionID = IIf(drInspectionCPReadings.Item("QUESTION_ID") Is DBNull.Value, 0, drInspectionCPReadings.Item("QUESTION_ID"))
                onLineNum = IIf(drInspectionCPReadings.Item("LINE_NUMBER") Is DBNull.Value, 0, drInspectionCPReadings.Item("LINE_NUMBER"))
                onTankPipeID = IIf(drInspectionCPReadings.Item("TANK_PIPE_ID") Is DBNull.Value, 0, drInspectionCPReadings.Item("TANK_PIPE_ID"))
                onTankPipeIndex = IIf(drInspectionCPReadings.Item("TANK_PIPE_INDEX") Is DBNull.Value, 0, drInspectionCPReadings.Item("TANK_PIPE_INDEX"))
                onTankPipeEntityID = IIf(drInspectionCPReadings.Item("TANK_PIPE_ENTITY_ID") Is DBNull.Value, 0, drInspectionCPReadings.Item("TANK_PIPE_ENTITY_ID"))
                'obolTankDispenser = IIf(drInspectionCPReadings.Item("TANK_DISPENSER") Is DBNull.Value, False, drInspectionCPReadings.Item("TANK_DISPENSER"))
                'obolGalvanic = IIf(drInspectionCPReadings.Item("GALVANIC") Is DBNull.Value, False, drInspectionCPReadings.Item("GALVANIC"))
                'onGalvanicICResponse = IIf(drInspectionCPReadings.Item("IMPRESSED_CURRENT") Is DBNull.Value, False, drInspectionCPReadings.Item("IMPRESSED_CURRENT"))
                ostrContactPoint = IIf(drInspectionCPReadings.Item("CONTACT_POINT") Is DBNull.Value, String.Empty, drInspectionCPReadings.Item("CONTACT_POINT"))
                ostrLocalReferCellPlacement = IIf(drInspectionCPReadings.Item("LOCAL_REFER_CELL_PLACEMENT") Is DBNull.Value, String.Empty, drInspectionCPReadings.Item("LOCAL_REFER_CELL_PLACEMENT"))
                ostrLocalOn = IIf(drInspectionCPReadings.Item("LOCAL_ON") Is DBNull.Value, String.Empty, drInspectionCPReadings.Item("LOCAL_ON"))
                ostrRemoteOff = IIf(drInspectionCPReadings.Item("REMOTE_OFF") Is DBNull.Value, String.Empty, drInspectionCPReadings.Item("REMOTE_OFF"))
                onPassFailIncon = IIf(drInspectionCPReadings.Item("PASS_FAIL_INCON") Is DBNull.Value, -1, drInspectionCPReadings.Item("PASS_FAIL_INCON"))
                obolDeleted = IIf(drInspectionCPReadings.Item("DELETED") Is DBNull.Value, False, drInspectionCPReadings.Item("DELETED"))
                ostrCreatedBy = IIf(drInspectionCPReadings.Item("CREATED_BY") Is DBNull.Value, String.Empty, drInspectionCPReadings.Item("CREATED_BY"))
                odtCreatedOn = IIf(drInspectionCPReadings.Item("DATE_CREATED") Is DBNull.Value, CDate("01/01/0001"), drInspectionCPReadings.Item("DATE_CREATED"))
                ostrModifiedBy = IIf(drInspectionCPReadings.Item("LAST_EDITED_BY") Is DBNull.Value, String.Empty, drInspectionCPReadings.Item("LAST_EDITED_BY"))
                odtModifiedOn = IIf(drInspectionCPReadings.Item("DATE_LAST_EDITED") Is DBNull.Value, CDate("01/01/0001"), drInspectionCPReadings.Item("DATE_LAST_EDITED"))
                obolRemoteReferCellPlacement = IIf(drInspectionCPReadings.Item("REMOTE_REFER_CELL_PLACEMENT") Is DBNull.Value, False, drInspectionCPReadings.Item("REMOTE_REFER_CELL_PLACEMENT"))
                obolGalvanicIC = IIf(drInspectionCPReadings.Item("GALVANIC_IC") Is DBNull.Value, False, drInspectionCPReadings.Item("GALVANIC_IC"))
                onGalvanicICResponse = IIf(drInspectionCPReadings.Item("GALVANIC_IC_RESPONSE") Is DBNull.Value, -1, drInspectionCPReadings.Item("GALVANIC_IC_RESPONSE"))
                obolTestedByInspector = IIf(drInspectionCPReadings.Item("TESTED_BY_INSPECTOR") Is DBNull.Value, False, drInspectionCPReadings.Item("TESTED_BY_INSPECTOR"))
                obolTestedByInspectorResponse = IIf(drInspectionCPReadings.Item("TESTED_BY_INSPECTOR_RESPONSE") Is DBNull.Value, False, drInspectionCPReadings.Item("TESTED_BY_INSPECTOR_RESPONSE"))
                Me.Reset()
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            If nInsCPReadID >= 0 Then
                nInsCPReadID = onInsCPReadID
            End If
            nInspectionID = onInspectionID
            nQuestionID = onQuestionID
            nLineNum = onLineNum
            nTankPipeID = onTankPipeID
            nTankPipeIndex = onTankPipeIndex
            nTankPipeEntityID = onTankPipeEntityID
            'bolTankDispenser = obolTankDispenser
            'bolGalvanic = obolGalvanic
            'nGalvanicICResponse = onGalvanicICResponse
            strContactPoint = ostrContactPoint
            strLocalReferCellPlacement = ostrLocalReferCellPlacement
            strLocalOn = ostrLocalOn
            strRemoteOff = ostrRemoteOff
            nPassFailIncon = onPassFailIncon
            bolDeleted = obolDeleted
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            bolRemoteReferCellPlacement = obolRemoteReferCellPlacement
            bolGalvanicIC = obolGalvanicIC
            nGalvanicICResponse = onGalvanicICResponse
            bolTestedByInspector = obolTestedByInspector
            bolTestedByInspectorResponse = obolTestedByInspectorResponse
            bolIsDirty = False
            RaiseEvent evtInspectionCPReadingsInfoChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            onInsCPReadID = nInsCPReadID
            onInspectionID = nInspectionID
            onQuestionID = nQuestionID
            onLineNum = nLineNum
            onTankPipeID = nTankPipeID
            onTankPipeIndex = nTankPipeIndex
            onTankPipeEntityID = nTankPipeEntityID
            'obolTankDispenser = bolTankDispenser
            'obolGalvanic = bolGalvanic
            'onGalvanicICResponse = nGalvanicICResponse
            ostrContactPoint = strContactPoint
            ostrLocalReferCellPlacement = strLocalReferCellPlacement
            ostrLocalOn = strLocalOn
            ostrRemoteOff = strRemoteOff
            onPassFailIncon = nPassFailIncon
            obolDeleted = bolDeleted
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn
            obolRemoteReferCellPlacement = bolRemoteReferCellPlacement
            obolGalvanicIC = bolGalvanicIC
            onGalvanicICResponse = nGalvanicICResponse
            obolTestedByInspector = bolTestedByInspector
            obolTestedByInspectorResponse = bolTestedByInspectorResponse
            bolIsDirty = False
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            Dim obolIsDirty As Boolean = bolIsDirty

            bolIsDirty = (nInspectionID <> onInspectionID) Or _
            (nQuestionID <> onQuestionID) Or _
            (nLineNum <> onLineNum) Or _
            (nTankPipeID <> onTankPipeID) Or _
            (nTankPipeIndex <> onTankPipeIndex) Or _
            (nTankPipeEntityID <> onTankPipeEntityID) Or _
            (strContactPoint <> ostrContactPoint) Or _
            (strLocalReferCellPlacement <> ostrLocalReferCellPlacement) Or _
            (strLocalOn <> ostrLocalOn) Or _
            (strRemoteOff <> ostrRemoteOff) Or _
            (nPassFailIncon <> onPassFailIncon) Or _
            (bolDeleted <> obolDeleted) Or _
            (strCreatedBy <> ostrCreatedBy) Or _
            (dtCreatedOn <> odtCreatedOn) Or _
            (strModifiedBy <> ostrModifiedBy) Or _
            (dtModifiedOn <> odtModifiedOn) Or _
            (bolRemoteReferCellPlacement <> obolRemoteReferCellPlacement) Or _
            (bolGalvanicIC <> obolGalvanicIC) Or _
            (nGalvanicICResponse <> onGalvanicICResponse) Or _
            (bolTestedByInspector <> obolTestedByInspector) Or _
            (bolTestedByInspectorResponse <> obolTestedByInspectorResponse)

            If obolIsDirty <> bolIsDirty Then
                RaiseEvent evtInspectionCPReadingsInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            onInsCPReadID = 0
            onInspectionID = 0
            onQuestionID = 0
            onLineNum = 0
            onTankPipeID = 0
            onTankPipeIndex = 0
            onTankPipeEntityID = 0
            'obolTankDispenser = Nothing
            'obolGalvanic = Nothing
            'onGalvanicICResponse = Nothing
            ostrContactPoint = String.Empty
            ostrLocalReferCellPlacement = String.Empty
            ostrLocalOn = String.Empty
            ostrRemoteOff = String.Empty
            onPassFailIncon = -1
            obolDeleted = False
            ostrCreatedBy = String.Empty
            odtCreatedOn = CDate("01/01/0001")
            ostrModifiedBy = String.Empty
            odtModifiedOn = CDate("01/01/0001")
            obolRemoteReferCellPlacement = False
            obolGalvanicIC = False
            onGalvanicICResponse = -1
            obolTestedByInspector = False
            obolTestedByInspectorResponse = False
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Integer
            Get
                Return nInsCPReadID
            End Get
            Set(ByVal Value As Integer)
                nInsCPReadID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property InspectionID() As Integer
            Get
                Return nInspectionID
            End Get
            Set(ByVal Value As Integer)
                nInspectionID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property QuestionID() As Integer
            Get
                Return nQuestionID
            End Get
            Set(ByVal Value As Integer)
                nQuestionID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LineNumber() As Integer
            Get
                Return nLineNum
            End Get
            Set(ByVal Value As Integer)
                nLineNum = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property TankPipeID() As Integer
            Get
                Return nTankPipeID
            End Get
            Set(ByVal Value As Integer)
                nTankPipeID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property TankPipeIndex() As Integer
            Get
                Return nTankPipeIndex
            End Get
            Set(ByVal Value As Integer)
                nTankPipeIndex = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property TankPipeEntityID() As Integer
            Get
                Return nTankPipeEntityID
            End Get
            Set(ByVal Value As Integer)
                nTankPipeEntityID = Value
                Me.CheckDirty()
            End Set
        End Property
        'Public Property TankDispenser() As Boolean
        '    Get
        '        Return bolTankDispenser
        '    End Get
        '    Set(ByVal Value As Boolean)
        '        bolTankDispenser = Value
        '        Me.CheckDirty()
        '    End Set
        'End Property
        'Public Property Galvanic() As Boolean
        '    Get
        '        Return bolGalvanic
        '    End Get
        '    Set(ByVal Value As Boolean)
        '        bolGalvanic = Value
        '        Me.CheckDirty()
        '    End Set
        'End Property
        'Public Property ImpressedCurrent() As Boolean
        '    Get
        '        Return nGalvanicICResponse
        '    End Get
        '    Set(ByVal Value As Boolean)
        '        nGalvanicICResponse = Value
        '        Me.CheckDirty()
        '    End Set
        'End Property
        Public Property ContactPoint() As String
            Get
                Return strContactPoint
            End Get
            Set(ByVal Value As String)
                strContactPoint = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LocalReferCellPlacement() As String
            Get
                Return strLocalReferCellPlacement
            End Get
            Set(ByVal Value As String)
                strLocalReferCellPlacement = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LocalOn() As String
            Get
                Return strLocalOn
            End Get
            Set(ByVal Value As String)
                strLocalOn = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property RemoteOff() As String
            Get
                Return strRemoteOff
            End Get
            Set(ByVal Value As String)
                strRemoteOff = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property PassFailIncon() As Integer
            Get
                Return nPassFailIncon
            End Get
            Set(ByVal Value As Integer)
                nPassFailIncon = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return bolDeleted
            End Get
            Set(ByVal Value As Boolean)
                bolDeleted = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return bolIsDirty
            End Get
            Set(ByVal value As Boolean)
                bolIsDirty = value
                RaiseEvent evtInspectionCPReadingsInfoChanged(bolIsDirty)
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
        Public Property RemoteReferCellPlacement() As Boolean
            Get
                Return bolRemoteReferCellPlacement
            End Get
            Set(ByVal Value As Boolean)
                bolRemoteReferCellPlacement = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property GalvanicIC() As Boolean
            Get
                Return bolGalvanicIC
            End Get
            Set(ByVal Value As Boolean)
                bolGalvanicIC = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property GalvanicICResponse() As Integer
            Get
                Return nGalvanicICResponse
            End Get
            Set(ByVal Value As Integer)
                nGalvanicICResponse = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property TestedByInspector() As Boolean
            Get
                Return bolTestedByInspector
            End Get
            Set(ByVal Value As Boolean)
                bolTestedByInspector = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property TestedByInspectorResponse() As Boolean
            Get
                Return bolTestedByInspectorResponse
            End Get
            Set(ByVal Value As Boolean)
                bolTestedByInspectorResponse = Value
                Me.CheckDirty()
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
