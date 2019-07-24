'-------------------------------------------------------------------------------
' MUSTER.Info.InspectionCommentsInfo
'   Provides the container to persist MUSTER InspectionComments state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       MNR        06/15/05    Original class definition
'
' Function          Description
' New()             Instantiates an empty InspectionCommentsInfo object
' New(Deleted, CreatedBy, CreatedOn, ModifiedBy, LastEdited, OwnerL2CSnippet)
'                   Instantiates a populated InspectionCommentsInfo object
' New(dr)           Instantiates a populated InspectionCommentsInfo object taking member state
'                   from the datarow provided
' Reset()           Sets the object state to the original state when loaded from or
'                   last saved to the repository
'
' NOTE: This file to be used as InspectionComments to build other objects.
'       Replace keyword "InspectionComments" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class InspectionCommentsInfo
#Region "Public Events"
        Public Event evtInspectionCommentsInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"
        Private nInsCommentsID As Int64
        Private nInspectionID As Int64
        Private strInsComments As String
        Private bolDeleted As Boolean
        Private strCreatedBy As String = String.Empty
        Private dtCreatedOn As DateTime
        Private strModifiedBy As String = String.Empty
        Private dtModifiedOn As DateTime

        Private onInsCommentsID As Int64
        Private onInspectionID As Int64
        Private ostrInsComments As String
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
        ByVal inspectionID As Int64, _
        ByVal insComments As String, _
        ByVal deleted As Boolean, _
        ByVal createdBy As String, _
        ByVal createdOn As Date, _
        ByVal modifiedBy As String, _
        ByVal modifiedOn As Date)
            onInsCommentsID = id
            onInspectionID = inspectionID
            ostrInsComments = insComments
            obolDeleted = deleted
            ostrCreatedBy = createdBy
            odtCreatedOn = createdOn
            ostrModifiedBy = modifiedBy
            odtModifiedOn = modifiedOn
            Me.Reset()
        End Sub
        Sub New(ByVal drInspectionComments As DataRow)
            Try
                onInsCommentsID = drInspectionComments.Item("INS_COMMENTS_ID")
                onInspectionID = IIf(drInspectionComments.Item("INSPECTION_ID") Is DBNull.Value, 0, drInspectionComments.Item("INSPECTION_ID"))
                ostrInsComments = IIf(drInspectionComments.Item("INS_COMMENTS") Is DBNull.Value, String.Empty, drInspectionComments.Item("INS_COMMENTS"))
                obolDeleted = IIf(drInspectionComments.Item("DELETED") Is DBNull.Value, False, drInspectionComments.Item("DELETED"))
                ostrCreatedBy = IIf(drInspectionComments.Item("CREATED_BY") Is DBNull.Value, String.Empty, drInspectionComments.Item("CREATED_BY"))
                odtCreatedOn = IIf(drInspectionComments.Item("DATE_CREATED") Is DBNull.Value, CDate("01/01/0001"), drInspectionComments.Item("DATE_CREATED"))
                ostrModifiedBy = IIf(drInspectionComments.Item("LAST_EDITED_BY") Is DBNull.Value, String.Empty, drInspectionComments.Item("LAST_EDITED_BY"))
                odtModifiedOn = IIf(drInspectionComments.Item("DATE_LAST_EDITED") Is DBNull.Value, CDate("01/01/0001"), drInspectionComments.Item("DATE_LAST_EDITED"))
                Me.Reset()
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            If nInsCommentsID >= 0 Then
                nInsCommentsID = onInsCommentsID
            End If
            nInspectionID = onInspectionID
            strInsComments = ostrInsComments
            bolDeleted = obolDeleted
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            bolIsDirty = False
            RaiseEvent evtInspectionCommentsInfoChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            onInsCommentsID = nInsCommentsID
            onInspectionID = nInspectionID
            ostrInsComments = strInsComments
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

            bolIsDirty = (nInspectionID <> onInspectionID) Or _
            (strInsComments <> ostrInsComments) Or _
            (bolDeleted <> obolDeleted) Or _
            (strCreatedBy <> ostrCreatedBy) Or _
            (dtCreatedOn <> odtCreatedOn) Or _
            (strModifiedBy <> ostrModifiedBy) Or _
            (dtModifiedOn <> odtModifiedOn)

            If obolIsDirty <> bolIsDirty Then
                RaiseEvent evtInspectionCommentsInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            onInsCommentsID = 0
            onInspectionID = 0
            ostrInsComments = String.Empty
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
                Return nInsCommentsID
            End Get
            Set(ByVal Value As Int64)
                nInsCommentsID = Value
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
        Public Property InsComments() As String
            Get
                Return strInsComments
            End Get
            Set(ByVal Value As String)
                strInsComments = Value
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
                RaiseEvent evtInspectionCommentsInfoChanged(bolIsDirty)
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
