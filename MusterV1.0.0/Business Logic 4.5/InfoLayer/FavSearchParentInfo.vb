'-------------------------------------------------------------------------------
' MUSTER.Info.FavSearchParentInfo
'   Provides the container to persist MUSTER FavSearch Parent data
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        MR         12/5/04     Original class definition.
'  1.1        AN         12/30/04    Added Try catch and Exception Handling/Logging
'  1.2        MR         1/7/05      Added Events for data update notification.
'                                    Added firing of event in ITEM()
'  1.3        JVC2       01/21/05    Added additional NEW that take INFO as argument
'                                       for copy operations.
'                                    Modified other NEW to include obolPublic initialization.
'                                    Changed SearchDel to Deleted for consistency.
'                                    Changed private member boolSearchDel to bolDeleted
'                                    Added nCriteriaCount and attribute CriteriaCount to
'                                       track and assist in criteria organization
'
' Operations
' Function          Description
' New()             Instantiates an empty FavSearch object.
' New(sUser)        Instantiates a populated FavSearchParentInfo object.
'                   
' Reset()           Sets the object state to the original state when loaded from or
'                       last saved to the repository.
'
' Read-Write Attributes
'Attribute          Description
' ID                The primary key associated with the FavSearchInfo in the repository.
'                     
' Name              The CriterionName of the FavSearchParentInfo object.
' SearchType        The CriterionValue of the FavSearchParentInfo object.
' User              The CriterionDataType of the FavSearchParentInfo object.
' CPublic           The Public state of the FavSearchParentInfo object (True = public).
' IsDirty           Indicates if the FavSearchParentInfo state has been altered since it was
'                       last loaded from or saved to the repository.
'
' Read-Only Attributes
' CreatedBy         The name of the user that created the FavSearchParentInfo object.
' CreatedOn         The date that the FavSearchParentInfo object was created.
' ModifiedBy        The name of the user that last modified the FavSearchParentInfo object.
' ModifiedOn        The date that the FavSearchParentInfo object was last modified.
'-------------------------------------------------------------------------------
Namespace MUSTER.Info
    <Serializable()> _
Public Class FavSearchParentInfo
        Implements iAccessors
#Region "Private member variables"
        'Variables for Original values
        Private onFavSearchID As Integer
        Private ostrUser As String
        Private ostrSearchType As String
        Private ostrSearchName As String
        Private ostrLUSTStatus As String
        Private ostrTankStatus As String
        Private obolPublic As Boolean

        'Variables for Current values
        Private nFavSearchID As Integer
        Private strUser As String
        Private strSearchType As String
        Private strSearchName As String
        Private strLUSTStatus As String
        Private strTankStatus As String
        Private bolPublic As Boolean

        Private nCriteriaCount As Int32 '

        Private strCreatedBy As String
        Private dtCreatedOn As DateTime
        Private strModifiedBy As String
        Private dtModifiedOn As DateTime
        Private bolIsDirty As Boolean = False
        Private bolDeleted As Boolean
        'Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Public Events"
        Public Event FavSearchInfoChanged(ByVal DirtyState As Boolean)
#End Region
#Region "Constructors"
        Public Sub New()
            onFavSearchID = 0
            ostrUser = String.Empty
            ostrSearchType = String.Empty
            ostrSearchName = String.Empty
            ostrLUSTStatus = "UNKNOWN"
            ostrTankStatus = "UNKNOWN"
            obolPublic = False
            Me.Reset()
        End Sub

        Public Sub New(ByVal nSearchID As Integer, _
                       ByVal sUser As String, _
                       ByVal sSearchType As String, _
                       ByVal sSearchName As String, _
                       ByVal sLustStatus As String, _
                       ByVal sTankStatus As String, _
                       Optional ByVal bPublic As Boolean = False)
            onFavSearchID = nSearchID
            ostrUser = sUser
            ostrSearchType = sSearchType
            ostrSearchName = sSearchName
            ostrLUSTStatus = sLustStatus
            ostrTankStatus = sTankStatus
            obolPublic = bPublic
            Me.Reset()
        End Sub

        Public Sub New(ByRef FavSrch As FavSearchParentInfo)
            onFavSearchID = FavSrch.ID
            ostrUser = FavSrch.User
            ostrSearchType = FavSrch.SearchType
            ostrSearchName = FavSrch.Name
            ostrLUSTStatus = FavSrch.LustStatus
            ostrTankStatus = FavSrch.TankStatus
            obolPublic = FavSrch.IsPublic
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Archive()
            onFavSearchID = nFavSearchID
            ostrUser = strUser
            ostrSearchType = strSearchType
            ostrSearchName = strSearchName
            ostrLUSTStatus = strLUSTStatus
            ostrTankStatus = strTankStatus
            obolPublic = bolPublic
            bolIsDirty = False
        End Sub
        Public Sub Reset()
            nFavSearchID = onFavSearchID
            strUser = ostrUser
            strSearchType = ostrSearchType
            strSearchName = ostrSearchName
            strLUSTStatus = ostrLUSTStatus
            strTankStatus = ostrTankStatus
            bolPublic = obolPublic
            bolDeleted = False
            bolIsDirty = False
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            Dim bolOldState As Boolean = bolIsDirty
            bolIsDirty = (nFavSearchID <> onFavSearchID) Or _
                         (strUser <> ostrUser) Or _
                         (strSearchType <> ostrSearchType) Or _
                         (strSearchName <> ostrSearchName) Or _
                         (strLUSTStatus <> ostrLUSTStatus) Or _
                         (strTankStatus <> ostrTankStatus) Or _
                         (bolPublic <> obolPublic)
            If bolOldState <> bolIsDirty Then
                RaiseEvent FavSearchInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()

            ostrUser = String.Empty
            ostrSearchType = String.Empty
            ostrSearchName = String.Empty
            ostrLUSTStatus = "UNKNOWN"
            ostrTankStatus = "UNKNOWN"
            obolPublic = False
            dtCreatedOn = System.DateTime.Now
            dtModifiedOn = System.DateTime.Now
            strCreatedBy = String.Empty
            strModifiedBy = String.Empty
            nCriteriaCount = 0
            Me.Reset()

        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Integer
            Get
                Return nFavSearchID
            End Get

            Set(ByVal value As Integer)
                nFavSearchID = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Name() As String
            Get
                Return strSearchName
            End Get
            Set(ByVal Value As String)
                strSearchName = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property SearchType() As String
            Get
                Return strSearchType
            End Get
            Set(ByVal Value As String)
                strSearchType = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LustStatus() As String
            Get
                Return strLUSTStatus
            End Get
            Set(ByVal Value As String)
                strLUSTStatus = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property TankStatus() As String
            Get
                Return strTankStatus
            End Get
            Set(ByVal Value As String)
                strTankStatus = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property User() As String
            Get
                Return strUser
            End Get
            Set(ByVal Value As String)
                strUser = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property IsPublic() As Boolean
            Get
                Return bolPublic
            End Get

            Set(ByVal value As Boolean)
                bolPublic = value
                Me.CheckDirty()
            End Set
        End Property
        Public ReadOnly Property NumCriteria() As Int32
            Get
                Return nCriteriaCount
            End Get
        End Property
        Public Property Deleted() As Boolean
            Get
                Return bolDeleted
            End Get
            Set(ByVal Value As Boolean)
                bolDeleted = Value
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

#Region "iAccessors"
        Public ReadOnly Property CreatedBy() As String Implements iAccessors.CreatedBy
            Get
                Return strCreatedBy
            End Get
        End Property
        Public ReadOnly Property CreatedOn() As Date Implements iAccessors.CreatedOn
            Get
                Return dtCreatedOn
            End Get
        End Property
        Public ReadOnly Property ModifiedBy() As String Implements iAccessors.ModifiedBy
            Get
                Return strModifiedBy
            End Get
        End Property
        Public ReadOnly Property ModifiedOn() As Date Implements iAccessors.ModifiedOn
            Get
                Return dtModifiedOn
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