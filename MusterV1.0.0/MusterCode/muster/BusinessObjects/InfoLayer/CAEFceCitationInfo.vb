'-------------------------------------------------------------------------------
' MUSTER.Info.CAEFceCitationInfo.vb
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0       MKK        08/15/05    Original class definition
'
' Function          Description
' New()             Instantiates an empty CAEFceCitationInfo object
' New(Deleted, CreatedBy, CreatedOn, ModifiedBy, LastEdited)
'                   Instantiates a populated CAEFceCitationInfo object
' New(dr)           Instantiates a populated CAEFceCitationInfo object taking member state
'                   from the datarow provided
' Reset()           Sets the object state to the original state when loaded from or
'                   last saved to the repository
'
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class CAEFceCitationInfo
#Region "Public Events"
        Public Event CAEFceCitationInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"
        Private nID As Int64
        Private nFacilityID As String
        Private nCitationID As String
        Private bolDeleted As Boolean
        Private strCreatedBy As String
        Private dtCreatedOn As DateTime
        Private strModifiedBy As String
        Private dtModifiedOn As DateTime

        Private onID As Int64
        Private onFacilityID As String
        Private onCitationID As String
        Private obolDeleted As Boolean
        Private ostrCreatedBy As String
        Private odtCreatedOn As DateTime
        Private ostrModifiedBy As String
        Private odtModifiedOn As DateTime

        Private bolShowDeleted As Boolean = False

        Private bolIsDirty As Boolean = False
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            Me.Init()
        End Sub
       
        Sub New(ByVal ID As Int64, _
        ByVal FacilityID As Integer, _
        ByVal CitationID As Integer, _
        ByVal Deleted As Boolean, _
        ByVal CreatedBy As String, _
        ByVal CreatedOn As Date, _
        ByVal ModifiedBy As String, _
        ByVal LastEdited As Date)
            onID = ID
            onFacilityID = FacilityID
            onCitationID = CitationID
            obolDeleted = Deleted
            ostrCreatedBy = CreatedBy
            odtCreatedOn = CreatedOn
            ostrModifiedBy = ModifiedBy
            odtModifiedOn = LastEdited
            Me.Reset()
        End Sub
        
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            nID = onID
            nFacilityID = onFacilityID
            nCitationID = onCitationID
            bolDeleted = obolDeleted
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            bolIsDirty = False
            RaiseEvent CAEFceCitationInfoChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            onID = nID
            onFacilityID = nFacilityID
            onCitationID = nCitationID
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

            bolIsDirty = (nID <> onID) Or _
            (nFacilityID <> onFacilityID) Or _
            (nCitationID <> onCitationID) Or _
            (bolDeleted <> obolDeleted) Or _
            (strCreatedBy <> ostrCreatedBy) Or _
            (dtCreatedOn <> odtCreatedOn) Or _
            (strModifiedBy <> ostrModifiedBy) Or _
            (dtModifiedOn <> odtModifiedOn)

            If obolIsDirty <> bolIsDirty Then
                RaiseEvent CAEFceCitationInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            onID = 0
            onCitationID = 0
            onFacilityID = 0
            bolDeleted = False
            strCreatedBy = String.Empty
            dtCreatedOn = System.DateTime.Now
            strModifiedBy = String.Empty
            dtModifiedOn = System.DateTime.Now
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Int64
            Get
                Return nID
            End Get
            Set(ByVal Value As Int64)
                nID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property FacilityID() As Integer
            Get
                Return nFacilityID
            End Get
            Set(ByVal Value As Integer)
                nFacilityID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CitationID() As Integer
            Get
                Return nCitationID
            End Get
            Set(ByVal Value As Integer)
                nCitationID = Value
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
            End Set
        End Property
        Public Property DATE_CREATED() As Date
            Get
                Return dtCreatedOn
            End Get
            Set(ByVal Value As Date)
                dtCreatedOn = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property DATE_LAST_EDITED() As Date
            Get
                Return dtModifiedOn
            End Get
            Set(ByVal Value As Date)
                dtModifiedOn = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CREATED_BY() As String
            Get
                Return strCreatedBy
            End Get
            Set(ByVal Value As String)
                strCreatedBy = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LAST_EDITED_BY() As String
            Get
                Return strModifiedBy
            End Get
            Set(ByVal Value As String)
                strModifiedBy = Value
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
