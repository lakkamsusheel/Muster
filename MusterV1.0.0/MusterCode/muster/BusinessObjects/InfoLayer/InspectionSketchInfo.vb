'-------------------------------------------------------------------------------
' MUSTER.Info.InspectionSketchInfo
'   Provides the container to persist MUSTER InspectionSketch state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       MNR        06/15/05    Original class definition
'
' Function          Description
' New()             Instantiates an empty InspectionSketchInfo object
' New(Deleted, CreatedBy, CreatedOn, ModifiedBy, LastEdited, OwnerL2CSnippet)
'                   Instantiates a populated InspectionSketchInfo object
' New(dr)           Instantiates a populated InspectionSketchInfo object taking member state
'                   from the datarow provided
' Reset()           Sets the object state to the original state when loaded from or
'                   last saved to the repository
'
' NOTE: This file to be used as InspectionSketch to build other objects.
'       Replace keyword "InspectionSketch" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class InspectionSketchInfo
#Region "Public Events"
        Public Event evtInspectionSketchInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"
        Private nInsSketchID As Int64
        Private nInspectionID As Int64
        Private strSketchFileName As String
        Private bolDeleted As Boolean
        Private strCreatedBy As String = String.Empty
        Private dtCreatedOn As DateTime
        Private strModifiedBy As String = String.Empty
        Private dtModifiedOn As DateTime

        Private onInsSketchID As Int64
        Private onInspectionID As Int64
        Private ostrSketchFileName As String
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
        ByVal sketchFileName As String, _
        ByVal deleted As Boolean, _
        ByVal createdBy As String, _
        ByVal createdOn As Date, _
        ByVal modifiedBy As String, _
        ByVal modifiedOn As Date)
            onInsSketchID = id
            nInsSketchID = id
            onInspectionID = inspectionID
            ostrSketchFileName = sketchFileName
            obolDeleted = deleted
            ostrCreatedBy = createdBy
            odtCreatedOn = createdOn
            ostrModifiedBy = modifiedBy
            odtModifiedOn = modifiedOn
            Me.Reset()
        End Sub
        Sub New(ByVal drInspectionSketch As DataRow)
            Try
                onInsSketchID = drInspectionSketch.Item("INS_SKETCH_ID")
                onInspectionID = drInspectionSketch.Item("INSPECTION_ID")
                ostrSketchFileName = IIf(drInspectionSketch.Item("SKETCH_FILE_NAME") Is DBNull.Value, String.Empty, drInspectionSketch.Item("SKETCH_FILE_NAME"))
                obolDeleted = IIf(drInspectionSketch.Item("DELETED") Is DBNull.Value, False, drInspectionSketch.Item("DELETED"))
                ostrCreatedBy = IIf(drInspectionSketch.Item("CREATED_BY") Is DBNull.Value, String.Empty, drInspectionSketch.Item("CREATED_BY"))
                odtCreatedOn = IIf(drInspectionSketch.Item("DATE_CREATED") Is DBNull.Value, CDate("01/01/0001"), drInspectionSketch.Item("DATE_CREATED"))
                ostrModifiedBy = IIf(drInspectionSketch.Item("LAST_EDITED_BY") Is DBNull.Value, String.Empty, drInspectionSketch.Item("LAST_EDITED_BY"))
                odtModifiedOn = IIf(drInspectionSketch.Item("DATE_LAST_EDITED") Is DBNull.Value, CDate("01/01/0001"), drInspectionSketch.Item("DATE_LAST_EDITED"))
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
            strSketchFileName = ostrSketchFileName
            bolDeleted = obolDeleted
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            bolIsDirty = False
            RaiseEvent evtInspectionSketchInfoChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            onInspectionID = nInspectionID
            ostrSketchFileName = strSketchFileName
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
            (strSketchFileName <> ostrSketchFileName) Or _
            (bolDeleted <> obolDeleted) Or _
            (strCreatedBy <> ostrCreatedBy) Or _
            (dtCreatedOn <> odtCreatedOn) Or _
            (strModifiedBy <> ostrModifiedBy) Or _
            (dtModifiedOn <> odtModifiedOn)

            If obolIsDirty <> bolIsDirty Then
                RaiseEvent evtInspectionSketchInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            onInsSketchID = 0
            onInspectionID = 0
            ostrSketchFileName = String.Empty
            obolDeleted = False
            ostrCreatedBy = String.Empty
            odtCreatedOn = System.DateTime.Now
            ostrModifiedBy = String.Empty
            odtModifiedOn = System.DateTime.Now
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Int64
            Get
                Return nInsSketchID
            End Get
            Set(ByVal Value As Int64)
                nInsSketchID = Value
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
        Public Property SketchFileName() As String
            Get
                Return strSketchFileName
            End Get
            Set(ByVal Value As String)
                strSketchFileName = Value
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
                RaiseEvent evtInspectionSketchInfoChanged(bolIsDirty)
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
