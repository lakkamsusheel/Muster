'-------------------------------------------------------------------------------
' MUSTER.Info.ContactStrcutInfo
'   Provides the container to persist MUSTER Template state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date            Description
'  1.0       KKM        03/30/2005      Original class definition


Namespace MUSTER.Info
    <Serializable()> _
Public Class ContactStructInfo
#Region "Public Events"
        Public Event ContactStructInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"
        ' tblCON_Entity_Relationships
        Private nEntityAssocID As Integer
        Private nEntityID As Integer
        Private nEntityType As Integer
        Private nContactAssocID As Integer
        Private nmoduleID As Integer
        Private nContactTypeID As Integer
        Private bolEntityAssocActive As Boolean
        Private strccInfo As String
        Private strDisplayAs As String
        Private bolEntityAssocDeleted As Boolean

        ' tblCON_Contact_Relationships
        ' Private nContactAssocID As Integer = declared above
        Private nContactParent As Integer
        Private nContactChild As Integer
        Private bolContactAssocActive As Boolean
        Private dtAssociated As DateTime
        Private bolContactAssocDeleted As Boolean
        Private npreferredAlias As Integer
        Private npreferredAddress As Integer


        Private strCreatedBy As String = String.Empty
        Private dtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private strModifiedBy As String = String.Empty
        Private dtModifiedOn As Date = DateTime.Now.ToShortDateString

        ' tblCON_Entity_Relationships
        Private onEntityAssocID As Integer
        Private onEntityID As Integer
        Private onEntityType As Integer
        Private onContactAssocID As Integer
        Private onmoduleID As Integer
        Private onContactTypeID As Integer
        Private obolEntityAssocActive As Boolean
        Private ostrccInfo As String
        Private ostrDisplayAs As String
        Private obolEntityAssocDeleted As Boolean

        ' tblCON_Contact_Relationships
        ' Private onContactAssocID As Integer = declared above
        Private onContactParent As Integer
        Private onContactChild As Integer
        Private obolContactAssocActive As Boolean

        Private onpreferredAlias As Integer
        Private onpreferredAddress As Integer

        Private odtAssociated As DateTime
        Private obolContactAssocDeleted As Boolean

        Private ostrCreatedBy As String
        Private odtCreatedOn As Date
        Private ostrModifiedBy As String
        Private odtModifiedOn As Date

        Private bolIsPerson As Boolean
        Private strContactModule As String

        Private obolIsPerson As Boolean
        Private ostrContactModule As String

        Private bolIsDirty As Boolean
        Private dtDataAge As Date
        Private nAgeThreshold As Int16 = 5

        Private ContactChild As MUSTER.Info.ContactDatumInfo
        Private ContactParent As MUSTER.Info.ContactDatumInfo
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            Me.Init()
            dtDataAge = Now()
            ContactParent = New MUSTER.Info.ContactDatumInfo
            ContactChild = New MUSTER.Info.ContactDatumInfo
        End Sub

        Sub New(ByVal Parent_Contact As Integer, _
        ByVal Child_Contact As Integer, _
        ByVal EntityID As Integer, _
        ByVal EntityType As Integer, _
        ByVal ContactTypeID As Integer, _
        ByVal ModuleID As Integer, _
        ByVal Active As Boolean, _
        ByVal CC_Info As String, _
        ByVal DisplayAs As String, _
        ByVal EntityRelationshipDeleted As Boolean, ByVal preferredAddress As Integer, ByVal preferredAlias As Integer)

            onContactParent = Parent_Contact
            onContactChild = Child_Contact
            onEntityID = EntityID
            onEntityType = EntityType
            onContactTypeID = ContactTypeID
            onmoduleID = ModuleID
            obolEntityAssocActive = Active
            ostrccInfo = CC_Info
            ostrDisplayAs = DisplayAs
            obolEntityAssocDeleted = EntityRelationshipDeleted
            Me.onpreferredAddress = preferredAddress
            Me.onpreferredAlias = preferredAlias
            ContactParent = New MUSTER.Info.ContactDatumInfo
            ContactChild = New MUSTER.Info.ContactDatumInfo
            Me.Reset()
        End Sub
        Sub New(ByVal drContactStruct As DataRow)
            Try
                onEntityAssocID = IIf(drContactStruct.Item("EntityAssocID") Is System.DBNull.Value, 0, drContactStruct.Item("EntityAssocID"))
                onEntityID = IIf(drContactStruct.Item("EntityID") Is System.DBNull.Value, 0, drContactStruct.Item("EntityID"))
                onEntityType = IIf(drContactStruct.Item("EntityType") Is System.DBNull.Value, 0, drContactStruct.Item("EntityType"))
                onContactAssocID = IIf(drContactStruct.Item("ContactAssocID") Is System.DBNull.Value, 0, drContactStruct.Item("ContactAssocID"))
                onContactTypeID = IIf(drContactStruct.Item("ContactTypeID") Is System.DBNull.Value, 0, drContactStruct.Item("ContactTypeID"))
                onmoduleID = IIf(drContactStruct.Item("ModuleID") Is System.DBNull.Value, 0, drContactStruct.Item("ModuleID"))
                ostrccInfo = IIf(drContactStruct.Item("CC_INFO") Is System.DBNull.Value, String.Empty, drContactStruct.Item("CC_INFO"))
                ostrDisplayAs = IIf(drContactStruct.Item("DisplayAs") Is System.DBNull.Value, String.Empty, drContactStruct.Item("DisplayAs"))
                onContactParent = IIf(drContactStruct.Item("Parent_Contact") Is System.DBNull.Value, 0, drContactStruct.Item("Parent_Contact"))
                onContactChild = IIf(drContactStruct.Item("Child_Contact") Is System.DBNull.Value, 0, drContactStruct.Item("Child_Contact"))
                odtAssociated = IIf(drContactStruct.Item("DateAssociated") Is System.DBNull.Value, CDate("01/01/0001"), drContactStruct.Item("DateAssociated"))
                'to be modified - seperate the deleted and active columns in the two tables
                obolContactAssocDeleted = IIf(drContactStruct.Item("Deleted") Is System.DBNull.Value, False, drContactStruct.Item("Deleted"))
                obolContactAssocActive = IIf(drContactStruct.Item("Active") Is System.DBNull.Value, False, drContactStruct.Item("Active"))
                obolEntityAssocDeleted = IIf(drContactStruct.Item("Deleted") Is System.DBNull.Value, False, drContactStruct.Item("Deleted"))
                obolEntityAssocActive = IIf(drContactStruct.Item("Active") Is System.DBNull.Value, False, drContactStruct.Item("Active"))
                ' end modifications
                dtDataAge = Now()
                onpreferredAlias = IIf(drContactStruct.Item("PreferredAlias") Is System.DBNull.Value, 0, drContactStruct.Item("PreferredAlias"))
                onpreferredAddress = IIf(drContactStruct.Item("PreferredAddressID") Is System.DBNull.Value, 0, drContactStruct.Item("PreferredAddressID"))

                ContactParent = New MUSTER.Info.ContactDatumInfo
                ContactChild = New MUSTER.Info.ContactDatumInfo
                Me.Reset()
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            nContactParent = onContactParent
            nContactChild = onContactChild
            nEntityID = onEntityID
            nContactAssocID = onContactAssocID
            nEntityAssocID = onEntityAssocID
            nEntityType = onEntityType
            nContactTypeID = onContactTypeID
            nmoduleID = onmoduleID
            strccInfo = ostrccInfo
            strDisplayAs = ostrDisplayAs
            bolEntityAssocActive = obolEntityAssocActive
            bolContactAssocActive = obolContactAssocActive
            bolEntityAssocDeleted = obolEntityAssocDeleted
            bolContactAssocDeleted = obolContactAssocDeleted
            dtAssociated = odtAssociated
            npreferredAlias = onpreferredAlias
            npreferredAddress = onpreferredAddress

            bolIsDirty = False
            RaiseEvent ContactStructInfoChanged(bolIsDirty)
        End Sub

        Public Sub Archive()
            onContactParent = nContactParent
            onContactChild = nContactChild
            onEntityID = nEntityID
            onEntityType = nEntityType
            onContactAssocID = nContactAssocID
            onContactTypeID = nContactTypeID
            onmoduleID = nmoduleID
            onEntityAssocID = nEntityAssocID
            obolEntityAssocActive = bolEntityAssocActive
            obolContactAssocActive = bolContactAssocActive
            obolEntityAssocDeleted = bolEntityAssocDeleted
            obolContactAssocDeleted = bolContactAssocDeleted
            ostrccInfo = strccInfo
            ostrDisplayAs = strDisplayAs
            odtAssociated = dtAssociated
            onpreferredAlias = npreferredAlias
            onpreferredAddress = npreferredAddress

            bolIsDirty = False
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            ' to be modified
            Dim obolIsDirty As Boolean = bolIsDirty
            bolIsDirty = ((onEntityAssocID <> nEntityAssocID) Or _
            (ostrModifiedBy <> strModifiedBy) Or _
            (ostrCreatedBy <> strCreatedBy) Or _
            (ostrccInfo <> strccInfo) Or _
            (ostrDisplayAs <> strDisplayAs) Or _
            (onmoduleID <> nmoduleID) Or _
            (obolIsDirty <> bolIsDirty) Or _
            (obolEntityAssocDeleted <> bolEntityAssocDeleted) Or _
            (odtCreatedOn <> dtCreatedOn) Or _
            (onEntityID <> nEntityID) Or _
            (onEntityType <> nEntityType) Or _
            (odtModifiedOn <> dtModifiedOn) Or _
            (onContactAssocID <> nContactAssocID) Or _
            (onpreferredAddress <> npreferredAddress) Or _
            (onpreferredAlias <> npreferredAlias))


            If bolIsDirty <> obolIsDirty Then
                RaiseEvent ContactStructInfoChanged(bolIsDirty)
            End If
        End Sub

        Private Sub Init()
            onEntityAssocID = 0
            ostrModifiedBy = String.Empty
            ostrCreatedBy = String.Empty
            obolEntityAssocDeleted = False
            obolContactAssocDeleted = False
            obolEntityAssocActive = True
            obolContactAssocActive = True
            odtCreatedOn = DateTime.Now.ToShortDateString()
            onEntityID = 0
            onEntityType = 0
            onContactAssocID = 0
            onContactChild = 0
            onContactParent = 0
            onpreferredAlias = 0
            onpreferredAddress = 0
            odtModifiedOn = DateTime.Now.ToShortDateString()
            odtAssociated = DateTime.Now.ToShortDateString()
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property EntityAssocActive() As Boolean
            Get
                Return bolEntityAssocActive
            End Get
            Set(ByVal Value As Boolean)
                bolEntityAssocActive = Value
            End Set
        End Property
        Public Property ContactAssocActive() As Boolean
            Get
                Return bolContactAssocActive
            End Get
            Set(ByVal Value As Boolean)
                bolContactAssocActive = Value
            End Set
        End Property
        Public Property displayAs() As String
            Get
                Return strDisplayAs
            End Get
            Set(ByVal Value As String)
                strDisplayAs = Value
            End Set
        End Property
        Public Property dateAssociated() As DateTime
            Get
                Return dtAssociated
            End Get
            Set(ByVal Value As DateTime)
                dtAssociated = Value
            End Set
        End Property
        Public Property moduleID() As Integer
            Get
                Return nmoduleID
            End Get
            Set(ByVal Value As Integer)
                nmoduleID = Value
            End Set
        End Property

        Public Property PreferredAlias() As Integer
            Get
                Return npreferredAlias
            End Get
            Set(ByVal Value As Integer)
                npreferredAlias = Value
            End Set
        End Property



        Public Property PreferredAddress() As Integer
            Get
                Return npreferredAddress
            End Get
            Set(ByVal Value As Integer)
                npreferredAddress = Value
            End Set
        End Property


        Public Property entityAssocID() As Integer
            Get
                Return nEntityAssocID
            End Get
            Set(ByVal Value As Integer)
                nEntityAssocID = Value
            End Set
        End Property

        Public Property ContactAssocID() As Integer
            Get
                Return nContactAssocID
            End Get
            Set(ByVal Value As Integer)
                nContactAssocID = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property IsDirty() As Boolean
            Get
                Return bolIsDirty
            End Get
            Set(ByVal Value As Boolean)
                bolIsDirty = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property EntityAssocdeleted() As Boolean
            Get
                Return bolEntityAssocDeleted
            End Get
            Set(ByVal Value As Boolean)
                bolEntityAssocDeleted = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property ContactAssocdeleted() As Boolean
            Get
                Return bolContactAssocDeleted
            End Get
            Set(ByVal Value As Boolean)
                bolContactAssocDeleted = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property AgeThreshold() As Int16
            Get
                Return nAgeThreshold
            End Get
            Set(ByVal value As Int16)
                nAgeThreshold = Int16.Parse(value)
            End Set
        End Property

        Public Property IsAgedData() As Boolean
            Get
                Return IIf(DateDiff(DateInterval.Minute, dtDataAge, Now()) >= nAgeThreshold, True, False)
            End Get
            Set(ByVal Value As Boolean)

            End Set
        End Property

        Public Property EntityID() As Integer
            Get
                Return nEntityID
            End Get
            Set(ByVal Value As Integer)
                nEntityID = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property parentContact() As MUSTER.Info.ContactDatumInfo
            Get
                Return ContactParent
            End Get
            Set(ByVal Value As MUSTER.Info.ContactDatumInfo)
                ContactParent = Value
            End Set
        End Property

        Public Property childContact() As MUSTER.info.ContactDatumInfo
            Get
                Return ContactChild
            End Get
            Set(ByVal Value As MUSTER.Info.ContactDatumInfo)
                ContactChild = Value
            End Set
        End Property

        Public Property ParentContactID() As Integer
            Get
                Return nContactParent
            End Get
            Set(ByVal Value As Integer)
                nContactParent = Value
            End Set
        End Property

        Public Property ChildContactID() As Integer
            Get
                Return nContactChild
            End Get
            Set(ByVal Value As Integer)
                nContactChild = Value
            End Set
        End Property

        Public Property ContactTypeID() As Integer
            Get
                Return nContactTypeID
            End Get
            Set(ByVal Value As Integer)
                nContactTypeID = Value
            End Set
        End Property

        Public Property contactModule() As String
            Get
                Return strContactModule
            End Get
            Set(ByVal Value As String)
                strContactModule = Value
            End Set
        End Property

        Public Property entityType() As Integer
            Get
                Return nEntityType
            End Get
            Set(ByVal Value As Integer)
                nEntityType = Value
            End Set
        End Property

        Public Property ccInfo() As String
            Get
                Return strccInfo
            End Get
            Set(ByVal Value As String)
                strccInfo = Value
            End Set
        End Property

#Region "iAccessors"
        Public ReadOnly Property modifiedOn() As Date
            Get
                Return dtModifiedOn
            End Get
        End Property

        Public Property modifiedBy() As String
            Get
                Return strModifiedBy
            End Get
            Set(ByVal Value As String)
                strModifiedBy = Value
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

        Public ReadOnly Property CreatedOn() As Date
            Get
                Return dtCreatedOn
            End Get
        End Property

#End Region

#End Region
#Region "Protected Operations"
        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
#End Region
    End Class
End Namespace
