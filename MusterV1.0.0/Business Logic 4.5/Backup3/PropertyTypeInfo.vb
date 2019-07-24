'-------------------------------------------------------------------------------
' MUSTER.Info.MusterPropertyTypeInfo
'   Provides the container to persist MUSTER Entity state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0      Elango      11/19/04    Original class definition.
'  1.1        AB        02/22/05    Added AgeThreshold and IsAgedData Attributes
'  1.2        MR        03/18/05    Removed Currect date assigned to Created On and Modified On.
'-------------------------------------------------------------------------------


Namespace MUSTER.Info
    <Serializable()> _
    Public Class PropertyTypeInfo

#Region "Private member variables"
        Private nPropTypeID As Long
        Private strPropType As String
        Private nEntityID As Long
        Private strCreatedBy As String = String.Empty
        Private dtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private strModifiedBy As String = String.Empty
        Private dtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private onPropTypeID As Long
        Private ostrPropType As String
        Private onEntityID As Long
        Private ostrCreatedBy As String
        Private odtCreatedOn As Date
        Private ostrModifiedBy As String
        Private odtModifiedOn As Date
        Private bolIsDirty As Boolean = False
        Private bolDeleted As Boolean
        Private obolDeleted As Boolean
        Private dtDataAge As DateTime
        Private nAgeThreshold As Int16 = 5

        Private WithEvents oPropertyCollection As MUSTER.Info.PropertyCollection
#End Region
#Region "Public Events"
        Public Event PropertyTypeChanged(ByVal bolValue As Boolean)
        Public Event PropertyCollectionChanged(ByVal bolValue As Boolean)
#End Region
#Region "constructors"
        ' The prototype New method
        Public Sub New(ByVal EntityId As Integer, ByVal PropertyTypeId As Integer, ByVal propType As String, ByVal createdby As String, ByVal createdon As String)
            onEntityID = EntityId
            onPropTypeID = PropertyTypeId
            ostrPropType = propType
            ostrCreatedBy = createdby
            odtCreatedOn = createdon
            ostrModifiedBy = String.Empty
            odtModifiedOn = CDate("01/01/0001")
            dtDataAge = Now()
            Me.Reset()
            oPropertyCollection = New MUSTER.Info.PropertyCollection
        End Sub
        Public Sub New()
            MyBase.New()
            dtDataAge = Now()
            oPropertyCollection = New MUSTER.Info.PropertyCollection
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            nEntityID = onEntityID
            nPropTypeID = onPropTypeID
            strPropType = ostrPropType
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            bolDeleted = obolDeleted
        End Sub
        Public Sub Archive()
            onEntityID = nEntityID
            onPropTypeID = nPropTypeID
            ostrPropType = strPropType
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn
            obolDeleted = bolDeleted
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            Dim bolOldValue As Boolean = bolIsDirty

            bolIsDirty = (nEntityID <> onEntityID) Or _
                         (nPropTypeID <> onPropTypeID) Or _
                         (strPropType <> ostrPropType) Or _
                         (bolDeleted <> obolDeleted)

            If bolOldValue <> bolIsDirty Then
                RaiseEvent PropertyTypeChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            onPropTypeID = 0
            ostrPropType = String.Empty
            onEntityID = 0
            ostrCreatedBy = String.Empty
            odtCreatedOn = CDate("01/01/0001")
            ostrModifiedBy = String.Empty
            odtModifiedOn = CDate("01/01/0001")
            obolDeleted = False
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        'The entity associated with the property type
        Public Property EntityId() As Integer
            Get
                Return nEntityID
            End Get
            Set(ByVal Value As Integer)
                nEntityID = Value
                Me.CheckDirty()
            End Set
        End Property
        ' The "common name" used to refer to the property type.
        Public Property Name() As String
            Get
                Return strPropType
            End Get
            Set(ByVal Value As String)
                strPropType = Value
                Me.CheckDirty()
            End Set
        End Property
        ' The identifier used by the repository to identify the class of the property type
        Public Property ID() As Long
            Get
                Return nPropTypeID
            End Get
            Set(ByVal Value As Long)
                nPropTypeID = Value
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

        Public Property Properties() As MUSTER.Info.PropertyCollection
            Get
                Return oPropertyCollection
            End Get
            Set(ByVal Value As MUSTER.Info.PropertyCollection)
                oPropertyCollection = Value
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

        Public ReadOnly Property IsAgedData() As Boolean
            Get
                Return IIf(DateDiff(DateInterval.Minute, dtDataAge, Now()) >= nAgeThreshold, True, False)
            End Get

        End Property

#End Region
#Region "iAccessors"
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
        Public Property ModifiedBy() As String
            Get
                Return strModifiedBy
            End Get
            Set(ByVal Value As String)
                strModifiedBy = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return dtModifiedOn
            End Get
        End Property
#End Region
#Region "Protected Operations"
        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
#End Region
#Region "External Event Handlers"
        Private Sub PropertyCollectionChangedSub() Handles oPropertyCollection.InfoChanged
            RaiseEvent PropertyTypeChanged(True)
            RaiseEvent PropertyCollectionChanged(True)
        End Sub
#End Region
    End Class
End Namespace


