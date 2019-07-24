'-------------------------------------------------------------------------------
' MUSTER.Info.InspectorCountyAssociationInfo
'   Provides the container to persist MUSTER InspectorCountyAssociation state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0                              Original class definition
'  1.1       JC         01/12/05    Added line of code to RESET to raise
'                                       data changed event when called.
'
' Function          Description
' New()             Instantiates an empty InspectorCountyAssociationInfo object
' New(Deleted, CreatedBy, CreatedOn, ModifiedBy, LastEdited, OwnerL2CSnippet)
'                   Instantiates a populated InspectorCountyAssociationInfo object
' New(dr)           Instantiates a populated InspectorCountyAssociationInfo object taking member state
'                   from the datarow provided
' Reset()           Sets the object state to the original state when loaded from or
'                   last saved to the repository
'
' NOTE: This file to be used as InspectorCountyAssociation to build other objects.
'       Replace keyword "InspectorCountyAssociation" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class InspectorCountyAssociationInfo
#Region "Public Events"
        Public Event InspectorCountyAssociationInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"
        Private nID As Int64
        Private nStaffId As Integer
        Private nFipsCode As Integer

        Private strCountyName As String
        Private nFacilities As Integer

        Private bolDeleted As Boolean
        Private strCreatedBy As String = String.Empty
        Private dtCreatedOn As DateTime
        Private strModifiedBy As String = String.Empty
        Private dtModifiedOn As DateTime

        Private onID As Int64
        Private onStaffId As Integer
        Private onFipsCode As Integer
        Private obolDeleted As Boolean
        Private ostrCreatedBy As String = String.Empty
        Private odtCreatedOn As DateTime
        Private ostrModifiedBy As String = String.Empty
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
        ByVal StaffID As Integer, _
        ByVal fipsCode As Integer, _
        ByVal CreatedBy As String, _
        ByVal CreatedOn As Date, _
        ByVal ModifiedBy As String, _
        ByVal LastEdited As Date, _
        ByVal Deleted As Boolean)
            onID = ID
            onStaffId = StaffID
            onFipsCode = fipsCode
            ostrCreatedBy = CreatedBy
            odtCreatedOn = CreatedOn
            ostrModifiedBy = ModifiedBy
            odtModifiedOn = LastEdited
            obolDeleted = Deleted
            Me.Reset()
        End Sub
        Sub New(ByVal drInspectorOwnerAssignment As DataRow)
            Try
                onID = drInspectorOwnerAssignment.Item("INS_OWNER_ID")
                onStaffId = drInspectorOwnerAssignment.Item("STAFF_ID")
                onFipsCode = drInspectorOwnerAssignment.Item("OWNER_ID")
                obolDeleted = drInspectorOwnerAssignment.Item("DELETED")
                ostrCreatedBy = drInspectorOwnerAssignment.Item("CREATED_BY")
                odtCreatedOn = drInspectorOwnerAssignment.Item("DATE_CREATED")
                ostrModifiedBy = drInspectorOwnerAssignment.Item("LAST_EDITED_BY")
                odtModifiedOn = drInspectorOwnerAssignment.Item("DATE_LAST_EDITED")
                Me.Reset()
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            nID = onID
            nStaffId = onStaffId
            nFipsCode = onFipsCode
            bolDeleted = obolDeleted
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            bolIsDirty = False
            RaiseEvent InspectorCountyAssociationInfoChanged(bolIsDirty)
            'InspectorOwnerAssignmentInfoChanged
        End Sub
        Public Sub Archive()
            onID = nID
            onStaffId = nStaffId
            onFipsCode = nFipsCode
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
            (nStaffId <> onStaffId) Or _
            (nFipsCode <> onFipsCode) Or _
            (bolDeleted <> obolDeleted) Or _
            (strCreatedBy <> ostrCreatedBy) Or _
            (dtCreatedOn <> odtCreatedOn) Or _
            (strModifiedBy <> ostrModifiedBy) Or _
            (dtModifiedOn <> odtModifiedOn)

            If obolIsDirty <> bolIsDirty Then
                RaiseEvent InspectorCountyAssociationInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            onID = 0
            onStaffId = 0
            onFipsCode = 0
            bolDeleted = False
            strCreatedBy = String.Empty
            dtCreatedOn = System.DateTime.Now
            strModifiedBy = String.Empty
            dtModifiedOn = System.DateTime.Now
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"

        Public Property ID() As Integer
            Get
                Return nID
            End Get
            Set(ByVal Value As Integer)
                nID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property STAFF_ID() As Integer
            Get
                Return nStaffId
            End Get
            Set(ByVal Value As Integer)
                nStaffId = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property FIPS_CODE() As Integer
            Get
                Return nFipsCode
            End Get
            Set(ByVal value As Integer)
                nFipsCode = value
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
        Public Property LAST_EDITED_BY() As String
            Get
                Return strModifiedBy
            End Get
            Set(ByVal Value As String)
                strModifiedBy = Value
                Me.CheckDirty()
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
        Public Property DELETED() As Boolean
            Get
                Return bolDeleted
            End Get
            Set(ByVal Value As Boolean)
                bolDeleted = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property County() As String
            Get
                Return strCountyName
            End Get
            Set(ByVal Value As String)
                strCountyName = Value
            End Set
        End Property
        Public Property Facilities() As Integer
            Get
                Return nFacilities
            End Get
            Set(ByVal Value As Integer)
                nFacilities = Value
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
