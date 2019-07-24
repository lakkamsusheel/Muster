'-------------------------------------------------------------------------------
' MUSTER.Info.InspectionCCATInfo
'   Provides the container to persist MUSTER InspectionCCAT state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       MNR        06/15/05    Original class definition
'
' Function          Description
' New()             Instantiates an empty InspectionCCATInfo object
' New(Deleted, CreatedBy, CreatedOn, ModifiedBy, LastEdited, OwnerL2CSnippet)
'                   Instantiates a populated InspectionCCATInfo object
' New(dr)           Instantiates a populated InspectionCCATInfo object taking member state
'                   from the datarow provided
' Reset()           Sets the object state to the original state when loaded from or
'                   last saved to the repository
'
' NOTE: This file to be used as InspectionCCAT to build other objects.
'       Replace keyword "InspectionCCAT" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
' 
' 6/8/2009   Thomas Franey      Added Compartment choice fot Tank CCAT
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class InspectionCCATInfo
#Region "Public Events"
        Public Event evtInspectionCCATInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"
        Private nInsCCATID As Integer
        Private nInspectionID As Integer
        Private nQuestionID As Integer
        Private nTankPipeID As Integer
        Private nTankPipeEntityID As Integer
        Private bolTankPipeResponse As Boolean
        Private bolTermination As Boolean
        Private strTankPipeResponseDetail As String
        Private bolDeleted As Boolean
        Private strCreatedBy As String = String.Empty
        Private dtCreatedOn As DateTime
        Private strModifiedBy As String = String.Empty
        Private dtModifiedOn As DateTime
        Private nCompartmentID As Integer
        Private bolFirstCompartment As Boolean = False


        Private onCompartmentID As Integer
        Private onInsCCATID As Integer
        Private onInspectionID As Integer
        Private onQuestionID As Integer
        Private onTankPipeID As Integer
        Private onTankPipeEntityID As Integer
        Private obolTankPipeResponse As Boolean
        Private obolTermination As Boolean
        Private ostrTankPipeResponseDetail As String
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
        Sub New(ByVal id As Integer, _
        ByVal InspectionID As Integer, _
        ByVal QuestionID As Integer, _
        ByVal TankPipeID As Integer, _
        ByVal TankPipeEntityID As Integer, _
        ByVal TankPipeResponse As Boolean, _
        ByVal termination As Boolean, _
        ByVal tankPipeResponseDetail As String, _
        ByVal deleted As Boolean, _
        ByVal createdBy As String, _
        ByVal createdOn As DateTime, _
        ByVal modifiedBy As String, _
        ByVal modifiedOn As DateTime, Optional ByVal compartmentID As Integer = 0)
            onInsCCATID = id
            onInspectionID = InspectionID
            onQuestionID = QuestionID
            onTankPipeID = TankPipeID
            onTankPipeEntityID = TankPipeEntityID
            obolTankPipeResponse = TankPipeResponse
            obolTermination = termination
            ostrTankPipeResponseDetail = tankPipeResponseDetail
            obolDeleted = deleted
            ostrCreatedBy = createdBy
            odtCreatedOn = createdOn
            ostrModifiedBy = modifiedBy
            odtModifiedOn = modifiedOn
            onCompartmentID = compartmentID
            Me.Reset()
        End Sub
        Sub New(ByVal drInspectionCCAT As DataRow)
            Try
                onInsCCATID = drInspectionCCAT.Item("INS_CCAT_ID")
                onInspectionID = IIf(drInspectionCCAT.Item("INSPECTION_ID") Is DBNull.Value, 0, drInspectionCCAT.Item("INSPECTION_ID"))
                onQuestionID = IIf(drInspectionCCAT.Item("QUESTION_ID") Is DBNull.Value, 0, drInspectionCCAT.Item("QUESTION_ID"))
                onTankPipeID = IIf(drInspectionCCAT.Item("TANK_PIPE_ID") Is DBNull.Value, 0, drInspectionCCAT.Item("TANK_PIPE_ID"))
                onTankPipeEntityID = IIf(drInspectionCCAT.Item("TANK_PIPE_ENTITY_ID") Is DBNull.Value, 0, drInspectionCCAT.Item("TANK_PIPE_ENTITY_ID"))
                obolTankPipeResponse = IIf(drInspectionCCAT.Item("TANK_PIPE_RESPONSE") Is DBNull.Value, False, drInspectionCCAT.Item("TANK_PIPE_RESPONSE"))
                obolTermination = IIf(drInspectionCCAT.Item("TERMINATION") Is DBNull.Value, False, drInspectionCCAT.Item("TERMINATION"))
                ostrTankPipeResponseDetail = IIf(drInspectionCCAT.Item("TANK_PIPE_RESPONSE_DETAILS") Is DBNull.Value, String.Empty, drInspectionCCAT.Item("TANK_PIPE_RESPONSE_DETAILS"))
                obolDeleted = IIf(drInspectionCCAT.Item("DELETED") Is DBNull.Value, False, drInspectionCCAT.Item("DELETED"))
                ostrCreatedBy = IIf(drInspectionCCAT.Item("CREATED_BY") Is DBNull.Value, String.Empty, drInspectionCCAT.Item("CREATED_BY"))
                odtCreatedOn = IIf(drInspectionCCAT.Item("DATE_CREATED") Is DBNull.Value, CDate("01/01/0001"), drInspectionCCAT.Item("DATE_CREATED"))
                ostrModifiedBy = IIf(drInspectionCCAT.Item("LAST_EDITED_BY") Is DBNull.Value, String.Empty, drInspectionCCAT.Item("LAST_EDITED_BY"))
                odtModifiedOn = IIf(drInspectionCCAT.Item("DATE_LAST_EDITED") Is DBNull.Value, CDate("01/01/0001"), drInspectionCCAT.Item("DATE_LAST_EDITED"))
                onCompartmentID = IIf(drInspectionCCAT.Item("CompartmentID") Is DBNull.Value, 0, drInspectionCCAT.Item("CompartmentID"))
                Me.Reset()
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            If nInsCCATID >= 0 Then
                nInsCCATID = onInsCCATID
            End If
            nInspectionID = onInspectionID
            nQuestionID = onQuestionID
            nTankPipeID = onTankPipeID
            nTankPipeEntityID = onTankPipeEntityID
            bolTankPipeResponse = obolTankPipeResponse
            bolTermination = obolTermination
            strTankPipeResponseDetail = ostrTankPipeResponseDetail
            bolDeleted = obolDeleted
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            bolIsDirty = False
            nCompartmentID = onCompartmentID

            RaiseEvent evtInspectionCCATInfoChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            onInsCCATID = nInsCCATID
            onInspectionID = nInspectionID
            onQuestionID = nQuestionID
            onTankPipeID = nTankPipeID
            onTankPipeEntityID = nTankPipeEntityID
            obolTankPipeResponse = bolTankPipeResponse
            obolTermination = bolTermination
            ostrTankPipeResponseDetail = strTankPipeResponseDetail
            obolDeleted = bolDeleted
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn
            onCompartmentID = nCompartmentID

            bolIsDirty = False
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            Dim obolIsDirty As Boolean = bolIsDirty

            bolIsDirty = (nInspectionID <> onInspectionID) Or _
            (nQuestionID <> onQuestionID) Or _
            (nTankPipeID <> onTankPipeID) Or _
            (nTankPipeEntityID <> onTankPipeEntityID) Or _
            (bolTankPipeResponse <> obolTankPipeResponse) Or _
            (bolTermination <> obolTermination) Or _
            (strTankPipeResponseDetail <> ostrTankPipeResponseDetail) Or _
            (bolDeleted <> obolDeleted) Or _
            (strCreatedBy <> ostrCreatedBy) Or _
            (dtCreatedOn <> odtCreatedOn) Or _
            (strModifiedBy <> ostrModifiedBy) Or _
            (dtModifiedOn <> odtModifiedOn) Or _
            (onCompartmentID <> nCompartmentID)


            If obolIsDirty <> bolIsDirty Then
                RaiseEvent evtInspectionCCATInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            onInsCCATID = 0
            onInspectionID = 0
            onQuestionID = 0
            onTankPipeID = 0
            onTankPipeEntityID = 0
            obolTankPipeResponse = False
            obolTermination = False
            ostrTankPipeResponseDetail = String.Empty
            obolDeleted = False
            ostrCreatedBy = String.Empty
            odtCreatedOn = CDate("01/01/0001")
            ostrModifiedBy = String.Empty
            odtModifiedOn = CDate("01/01/0001")
            onCompartmentID = 0
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Integer
            Get
                Return nInsCCATID
            End Get
            Set(ByVal Value As Integer)
                nInsCCATID = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property CompartmentID() As Integer
            Get
                Return nCompartmentID
            End Get

            Set(ByVal Value As Integer)
                nCompartmentID = Value
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

        Public Property FirstCompartment() As Boolean

            Get
                Return bolFirstCompartment
            End Get

            Set(ByVal Value As Boolean)
                bolFirstCompartment = Value
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
        Public Property TankPipeEntityID() As Integer
            Get
                Return nTankPipeEntityID
            End Get
            Set(ByVal Value As Integer)
                nTankPipeEntityID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property TankPipeResponse() As Boolean
            Get
                Return bolTankPipeResponse
            End Get
            Set(ByVal Value As Boolean)
                bolTankPipeResponse = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Termination() As Boolean
            Get
                Return bolTermination
            End Get
            Set(ByVal Value As Boolean)
                bolTermination = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property TankPipeResponseDetail() As String
            Get
                Return strTankPipeResponseDetail
            End Get
            Set(ByVal Value As String)
                strTankPipeResponseDetail = Value
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
                RaiseEvent evtInspectionCCATInfoChanged(bolIsDirty)
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
