'-------------------------------------------------------------------------------
' MUSTER.Info.InspectionDiscrepInfo
'   Provides the container to persist MUSTER InspectionDiscrep state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       MNR        06/15/05    Original class definition
'
' Function          Description
' New()             Instantiates an empty InspectionDiscrepInfo object
' New(Deleted, CreatedBy, CreatedOn, ModifiedBy, LastEdited, OwnerL2CSnippet)
'                   Instantiates a populated InspectionDiscrepInfo object
' New(dr)           Instantiates a populated InspectionDiscrepInfo object taking member state
'                   from the datarow provided
' Reset()           Sets the object state to the original state when loaded from or
'                   last saved to the repository
'
' NOTE: This file to be used as InspectionDiscrep to build other objects.
'       Replace keyword "InspectionDiscrep" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class InspectionDiscrepInfo
#Region "Public Events"
        Public Event evtInspectionDiscrepInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"
        Private nInsDiscrepID As Int64
        Private nInspectionID As Int64
        Private nQuestionID As Int64
        Private strDescription As String
        Private bolDeleted As Boolean
        Private strCreatedBy As String = String.Empty
        Private dtCreatedOn As DateTime
        Private strModifiedBy As String = String.Empty
        Private dtModifiedOn As DateTime
        Private bolRescinded As Boolean
        Private dtDiscrepReceivedDate As Date
        Private nInspCitID As Int64

        Private onInsDiscrepID As Int64
        Private onInspectionID As Int64
        Private onQuestionID As Int64
        Private ostrDescription As String
        Private obolDeleted As Boolean
        Private ostrCreatedBy As String = String.Empty
        Private odtCreatedOn As DateTime
        Private ostrModifiedBy As String = String.Empty
        Private odtModifiedOn As DateTime
        Private obolRescinded As Boolean
        Private odtDiscrepReceivedDate As Date
        Private onInspCitID As Int64

        Private bolIsDirty As Boolean = False
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            Me.Init()
        End Sub
        Sub New(ByVal id As Int64, _
        ByVal inspectionID As Int64, _
        ByVal questionID As Int64, _
        ByVal description As String, _
        ByVal deleted As Boolean, _
        ByVal createdBy As String, _
        ByVal createdOn As Date, _
        ByVal modifiedBy As String, _
        ByVal modifiedOn As Date, _
        ByVal rescind As Boolean, _
        ByVal discrepRecvd As Date, _
        ByVal inspCitID As Int64)
            onInsDiscrepID = id
            onInspectionID = inspectionID
            onQuestionID = questionID
            ostrDescription = description
            obolDeleted = deleted
            ostrCreatedBy = createdBy
            odtCreatedOn = createdOn
            ostrModifiedBy = modifiedBy
            odtModifiedOn = modifiedOn
            obolRescinded = rescind
            odtDiscrepReceivedDate = discrepRecvd
            onInspCitID = inspCitID
            Me.Reset()
        End Sub
        Sub New(ByVal drInspectionDiscrep As DataRow)
            Try
                onInsDiscrepID = drInspectionDiscrep.Item("INS_DESCREP_ID")
                onInspectionID = IIf(drInspectionDiscrep.Item("INSPECTION_ID") Is DBNull.Value, 0, drInspectionDiscrep.Item("INSPECTION_ID"))
                onQuestionID = IIf(drInspectionDiscrep.Item("QUESTION_ID") Is DBNull.Value, 0, drInspectionDiscrep.Item("QUESTION_ID"))
                ostrDescription = IIf(drInspectionDiscrep.Item("DESCRIPTION") Is DBNull.Value, String.Empty, drInspectionDiscrep.Item("DESCRIPTION"))
                obolDeleted = IIf(drInspectionDiscrep.Item("DELETED") Is DBNull.Value, False, drInspectionDiscrep.Item("DELETED"))
                ostrCreatedBy = IIf(drInspectionDiscrep.Item("CREATED_BY") Is DBNull.Value, String.Empty, drInspectionDiscrep.Item("CREATED_BY"))
                odtCreatedOn = IIf(drInspectionDiscrep.Item("DATE_CREATED") Is DBNull.Value, CDate("01/01/0001"), drInspectionDiscrep.Item("DATE_CREATED"))
                ostrModifiedBy = IIf(drInspectionDiscrep.Item("LAST_EDITED_BY") Is DBNull.Value, String.Empty, drInspectionDiscrep.Item("LAST_EDITED_BY"))
                odtModifiedOn = IIf(drInspectionDiscrep.Item("DATE_LAST_EDITED") Is DBNull.Value, CDate("01/01/0001"), drInspectionDiscrep.Item("DATE_LAST_EDITED"))
                obolRescinded = IIf(drInspectionDiscrep.Item("RESCINDED") Is DBNull.Value, False, drInspectionDiscrep.Item("RESCINDED"))
                odtDiscrepReceivedDate = IIf(drInspectionDiscrep.Item("DISCREP_RECEIVED_DATE") Is DBNull.Value, CDate("01/01/0001"), drInspectionDiscrep.Item("DISCREP_RECEIVED_DATE"))
                onInspCitID = IIf(drInspectionDiscrep.Item("INS_CIT_ID") Is DBNull.Value, 0, drInspectionDiscrep.Item("INS_CIT_ID"))
                Me.Reset()
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            If nInsDiscrepID >= 0 Then
                nInsDiscrepID = onInsDiscrepID
            End If
            nInspectionID = onInspectionID
            nQuestionID = onQuestionID
            strDescription = ostrDescription
            bolDeleted = obolDeleted
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            bolRescinded = obolRescinded
            dtDiscrepReceivedDate = odtDiscrepReceivedDate
            nInspCitID = onInspCitID
            bolIsDirty = False
            RaiseEvent evtInspectionDiscrepInfoChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            onInsDiscrepID = nInsDiscrepID
            onInspectionID = nInspectionID
            onQuestionID = nQuestionID
            ostrDescription = strDescription
            obolDeleted = bolDeleted
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn
            obolRescinded = bolRescinded
            odtDiscrepReceivedDate = dtDiscrepReceivedDate
            onInspCitID = nInspCitID
            bolIsDirty = False
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            Dim obolIsDirty As Boolean = bolIsDirty

            bolIsDirty = (nInspectionID <> onInspectionID) Or _
            (nQuestionID <> onQuestionID) Or _
            (strDescription <> ostrDescription) Or _
            (bolDeleted <> obolDeleted) Or _
            (strCreatedBy <> ostrCreatedBy) Or _
            (dtCreatedOn <> odtCreatedOn) Or _
            (strModifiedBy <> ostrModifiedBy) Or _
            (dtModifiedOn <> odtModifiedOn) Or _
            (bolRescinded <> obolRescinded) Or _
            (dtDiscrepReceivedDate <> odtDiscrepReceivedDate) Or _
            (nInspCitID <> onInspCitID)

            If obolIsDirty <> bolIsDirty Then
                RaiseEvent evtInspectionDiscrepInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            onInsDiscrepID = 0
            onInspectionID = 0
            onQuestionID = 0
            ostrDescription = String.Empty
            obolDeleted = False
            ostrCreatedBy = String.Empty
            odtCreatedOn = System.DateTime.Now
            ostrModifiedBy = String.Empty
            odtModifiedOn = System.DateTime.Now
            obolRescinded = False
            odtDiscrepReceivedDate = CDate("01/01/0001")
            onInspCitID = 0
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Int64
            Get
                Return nInsDiscrepID
            End Get
            Set(ByVal Value As Int64)
                nInsDiscrepID = Value
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
        Public Property InspCitID() As Int64
            Get
                Return nInspCitID
            End Get
            Set(ByVal Value As Int64)
                nInspCitID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Description() As String
            Get
                Return strDescription
            End Get
            Set(ByVal Value As String)
                strDescription = Value
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
                RaiseEvent evtInspectionDiscrepInfoChanged(IsDirty)
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
        Public Property Rescinded() As Boolean
            Get
                Return bolRescinded
            End Get
            Set(ByVal Value As Boolean)
                bolRescinded = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property DiscrepReceived() As Date
            Get
                Return dtDiscrepReceivedDate
            End Get
            Set(ByVal Value As Date)
                dtDiscrepReceivedDate = Value
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
