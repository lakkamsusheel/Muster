'-------------------------------------------------------------------------------
' MUSTER.Info.FavSearchChildInfo
'   Provides the container to persist MUSTER FavSearch Child data
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        MR         12/5/04     Original class definition.
'  1.1        AN         12/30/04    Added Try catch and Exception Handling/Logging
'  1.2        MR         1/7/05      Added Events for data update notification.
'                                    Added firing of event in ITEM()
'  1.3        JVC2      01/21/05     Changed Name to CriterionName
'                                    Added overloaded NEW operation which takes a 
'                                       FavSearchChildInfo (for copy)
'                                       Changed all NEW operations to use RESET.
'                                       Removed attribute KEY - serves no function
'                                       Changed attribute CriteriaDelete to Deleted
'                                       Changed obolCriteriaDeleted to obolDeleted
'                                       Changed bolCriteriaDeleted to bolDeleted
'                                       Added nCriterionOrder attribute
'
' Operations
' Function          Description
' New()             Instantiates an empty FavSearchChildInfo object.
' New(strCrName,strCrValue,strCrDataType)
'                   Instantiates a populated FavSearchChildInfo object.
' Reset()           Sets the object state to the original state when loaded from or
'                       last saved to the repository.
'
' Read-Write Attributes
'Attribute          Description
' CriteriaID        The primary key associated with the FavSearchChildInfo in the repository.
'                     
' CriterionName     The CriterionName of the FavSearchChildInfo object.
' CriterionValue    The CriterionValue of the FavSearchChildInfo object.
' CriterionDataType The CriterionDataType of the FavSearchChildInfo object.
' CriteriaDelete    The deleted state of the FavSearchChildInfo object (True = deleted).
' IsDirty           Indicates if the FavSearchChildInfo state has been altered since it was
'                       last loaded from or saved to the repository.
'
' Read-Only Attributes
' CreatedBy         The name of the user that created the FavSearchChildInfo object.
' CreatedOn         The date that the FavSearchChildInfo object was created.
' ModifiedBy        The name of the user that last modified the FavSearchChildInfo object.
' ModifiedOn        The date that the FavSearchChildInfo object was last modified.
'-------------------------------------------------------------------------------
'
' TODO - Integrate into BusinessObjects solution 1/23/05 - JVC 2
'
Namespace MUSTER.Info
    <Serializable()> _
Public Class FavSearchChildInfo
        Implements iAccessors
#Region "Public Events"
        Public Event CriteriaInfoChanged(ByVal DirtyState As Boolean)
#End Region
#Region "Private member variables"
        'Original Variables
        Private onCriterionID As Integer
        Private onCriterionOrder As Integer
        Private ostrCriterionName As String
        Private ostrCriterionValue As String
        Private ostrCriterionDataType As String
        Private onParentID As Integer
        Private obolDeleted As Boolean

        'Current Values
        Private nCriterionID As Integer
        Private nCriterionOrder As Integer
        Private strKeyID As String
        Private nParentID As Integer
        Private strCriterionName As String
        Private strCriterionValue As String
        Private strCriterionDataType As String
        Private bolDeleted As Boolean


        Private strCreatedBy As String
        Private dtCreatedOn As DateTime
        Private strModifiedBy As String
        Private dtModifiedOn As DateTime
        Private bolIsDirty As Boolean = False
        'Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Public Sub New()
            onCriterionID = 0
            onCriterionOrder = 0
            ostrCriterionName = ""
            ostrCriterionValue = ""
            ostrCriterionDataType = ""
            onParentID = 0
            Me.Reset()
        End Sub
        Public Sub New(ByVal CriterionID As Integer, ByVal CriterionOrder As Integer, ByVal strCrName As String, ByVal strCrValue As String, ByVal strCrDataType As String, ByVal ParentID As Integer)
            onCriterionID = CriterionID
            onCriterionOrder = CriterionOrder
            ostrCriterionName = strCrName
            ostrCriterionValue = strCrValue
            ostrCriterionDataType = strCrDataType
            onParentID = ParentID
            obolDeleted = False
            Me.Reset()
        End Sub
        Public Sub New(ByRef ChildInfo As MUSTER.Info.FavSearchChildInfo)
            onCriterionID = ChildInfo.ID
            onCriterionOrder = ChildInfo.Order
            ostrCriterionName = ChildInfo.CriterionName
            ostrCriterionValue = ChildInfo.CriterionValue
            ostrCriterionDataType = ChildInfo.CriterionDataType
            onParentID = ChildInfo.ParentID
            obolDeleted = ChildInfo.Deleted
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Archive()
            onCriterionID = nCriterionID
            onCriterionOrder = nCriterionOrder
            ostrCriterionName = strCriterionName
            ostrCriterionValue = strCriterionValue
            ostrCriterionDataType = strCriterionDataType
            obolDeleted = bolDeleted
            onParentID = nParentID
            bolIsDirty = False
        End Sub
        Public Sub Reset()
            nCriterionID = onCriterionID
            nCriterionOrder = onCriterionOrder
            strCriterionName = ostrCriterionName
            strCriterionValue = ostrCriterionValue
            strCriterionDataType = ostrCriterionDataType
            bolDeleted = obolDeleted
            nParentID = onParentID
            bolIsDirty = False
        End Sub
#End Region
#Region "Private Operations"

        Private Sub CheckDirty()
            Dim bolOldState As Boolean = bolIsDirty
            bolIsDirty = (nCriterionID <> onCriterionID) Or _
                         (nCriterionOrder <> onCriterionOrder) Or _
                         (strCriterionName <> ostrCriterionName) Or _
                         (strCriterionValue <> ostrCriterionValue) Or _
                         (strCriterionDataType <> ostrCriterionDataType) Or _
                         (nParentID <> onParentID) Or _
                         (bolDeleted <> obolDeleted)
            If bolOldState <> bolIsDirty Then
                RaiseEvent CriteriaInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            onCriterionID = 0
            onCriterionOrder = 0
            ostrCriterionName = String.Empty
            ostrCriterionValue = String.Empty
            ostrCriterionDataType = String.Empty
            onParentID = 0
            obolDeleted = False
            dtCreatedOn = System.DateTime.Now
            dtModifiedOn = System.DateTime.Now
            strCreatedBy = String.Empty
            strModifiedBy = String.Empty
            Me.Reset()

        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Integer
            Get
                Return nCriterionID
            End Get

            Set(ByVal value As Integer)
                nCriterionID = Integer.Parse(value)
                Me.CheckDirty()
            End Set
        End Property
        Public Property Order() As Integer
            Get
                Return nCriterionOrder
            End Get
            Set(ByVal Value As Integer)
                nCriterionOrder = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CriterionName() As String
            Get
                Return strCriterionName
            End Get
            Set(ByVal Value As String)
                strCriterionName = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CriterionValue() As String
            Get
                Return strCriterionValue
            End Get
            Set(ByVal Value As String)
                strCriterionValue = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CriterionDataType() As String
            Get
                Return strCriterionDataType
            End Get
            Set(ByVal Value As String)
                strCriterionDataType = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ParentID() As Integer
            Get
                Return nParentID
            End Get

            Set(ByVal value As Integer)
                nParentID = value
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

