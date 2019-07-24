'-------------------------------------------------------------------------------
' MUSTER.Info.MusterPropertyInfo
'   Provides the container to persist MUSTER Entity state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0      EN      11/19/04    Original class definition.
'  1.1      AN      12/30/04    Added Try catch and Exception Handling/Logging
'  1.2      AB      02/22/05    Added AgeThreshold and IsAgedData Attributes
'  1.3      MR      03/18/05    Removed Currect date assigned to Created On and Modified On.
'
'
' Function          Description
' New()             Instantiates an empty EntityInfo object.
' New(ID, Name, CreatedBy, CreatedOn, ModifiedBy, ModifiedOn)

'New(ByVal PROPERTY_ID,PROPERTY_TYPE_ID,PROPERTY_NAME,PROPERTY_DESCRIPTION,PROPERTY_POSITION,BUSINESS_TAG,PROPERTY_ACTIVE,CREATED_BY,CREATE_DATE,LAST_EDITED_BY,DATE_LAST_EDITED
'                   Instantiates a populated MusterPropertyInfo object.
' New(ds)           Instantiates a populated MusterPropertyInfo object taking member state
'                       from the first row in the first table in the dataset provided
' Reset()           Sets the object state to the original state when loaded from or
'                       last saved to the repository.
' Save()            Saves the object state to the repository.
'
'Attribute          Description
' PROPERTY_ID       The unique identifier associated with the Property in the repository.
' PROPERTY_NAME     The name of the MusterProperty.
' IsDirty            Indicates if the MusterProperty state has been altered since it was
'                       last loaded from or saved to the repository.
'-------------------------------------------------------------------------------


Namespace MUSTER.Info
    <Serializable()> _
      Public Class PropertyInfo

#Region "Private member variables"

        Private bolPropIsActive As Boolean
        Private nPropID As Long
        Private nPropParentID As Long = 0
        Private nPropPos As Integer
        Private nPropTypeID As Long
        Private strPropDesc As String
        Private strPropName As String
        Private strPropType As String
        Private strCreatedBy As String = String.Empty
        Private dtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private strModifiedBy As String = String.Empty
        Private dtModifiedOn As Date = DateTime.Now.ToShortDateString

        Private obolPropIsActive As Boolean
        Private onPropID As Long
        Private onPropParentID As Long = 0
        Private onPropPos As Integer
        Private onPropTypeID As Long
        Private ostrPropDesc As String
        Private ostrPropName As String
        Private ostrPropType As String
        Private ostrCreatedBy As String
        Private odtCreatedOn As Date
        Private ostrModifiedBy As String
        Private odtModifiedOn As Date
        Private onBUSINESS_TAG As Integer
        Private nBUSINESS_TAG As Integer
        Private bolIsDirty As Boolean = False
        Private bolDeleted As Boolean
        Private strAVAILABLE_PROPERTY_DISPLAY As String
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
        Private dtDataAge As DateTime
        Private nAgeThreshold As Int16 = 5
        Private WithEvents PropertiesChildCollection As New MUSTER.Info.PropertyCollection
#End Region
#Region "Public Events"
        Public Event PropertyChanged(ByVal bolValue As Boolean)
        Public Event PropertiesChildCollectionChanged(ByVal bolValue As Boolean)
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            dtDataAge = Now()
        End Sub
        Sub New(ByVal PROPERTY_ID As Integer, _
            ByVal PROPERTY_PARENT_ID As Integer, _
            ByVal PROPERTY_TYPE_ID As Integer, _
            ByVal PROPERTY_NAME As String, _
            ByVal PROPERTY_DESCRIPTION As String, _
            ByVal PROPERTY_POSITION As Integer, _
            ByVal BUSINESS_TAG As Integer, _
            ByVal PROPERTY_ACTIVE As Boolean, _
            ByVal CREATED_BY As String, _
            ByVal CREATE_DATE As Date, _
            ByVal LAST_EDITED_BY As String, _
            ByVal DATE_LAST_EDITED As Date)

            onPropID = PROPERTY_ID
            onPropParentID = PROPERTY_PARENT_ID
            onPropTypeID = PROPERTY_TYPE_ID
            ostrPropName = PROPERTY_NAME
            ostrPropDesc = PROPERTY_DESCRIPTION
            onPropPos = PROPERTY_POSITION
            onBUSINESS_TAG = nBUSINESS_TAG
            obolPropIsActive = PROPERTY_ACTIVE
            ostrCreatedBy = AltIsDBNull(CREATED_BY, String.Empty)
            odtCreatedOn = AltIsDBNull(CREATE_DATE, CDate("01/01/0001"))
            ostrModifiedBy = AltIsDBNull(LAST_EDITED_BY, String.Empty)
            odtModifiedOn = AltIsDBNull(DATE_LAST_EDITED, CDate("01/01/0001"))
            dtDataAge = Now()
            Me.Reset()

        End Sub

        Private Sub New(ByVal oProp As MUSTER.Info.PropertyInfo)
            Try
                onPropID = oProp.ID
                onPropParentID = oProp.Parent_ID
                onPropTypeID = oProp.PropType_ID
                ostrPropName = oProp.Name
                ostrPropDesc = oProp.PropDesc
                onPropPos = oProp.PropPos
                onBUSINESS_TAG = oProp.nBUSINESS_TAG
                obolPropIsActive = oProp.PropIsActive
                ostrCreatedBy = oProp.CreatedBy
                odtCreatedOn = oProp.CreatedOn
                ostrModifiedBy = oProp.ModifiedBy
                odtModifiedOn = oProp.ModifiedOn
                dtDataAge = Now()
                Me.Reset()
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub

        Private Function AltIsDBNull(ByVal ObjectToCheck As Object, ByVal AlternateObject As Object) As Object
            If ObjectToCheck Is System.DBNull.Value Then
                AltIsDBNull = AlternateObject
            Else
                AltIsDBNull = ObjectToCheck
            End If
        End Function

        Sub New(ByVal dr As DataRow)
            Try
                onPropID = AltIsDBNull(dr.Item("PROPERTY_ID"), 0)
                onPropParentID = AltIsDBNull(dr.Item("PARENT_ID"), 0)
                onPropTypeID = AltIsDBNull(dr.Item("PROPERTY_TYPE_ID"), 0)
                ostrPropName = AltIsDBNull(dr.Item("PROPERTY_NAME"), String.Empty)
                ostrPropDesc = AltIsDBNull(dr.Item("PROPERTY_DESCRIPTION"), String.Empty)
                onPropPos = AltIsDBNull(dr.Item("PROPERTY_POSITION"), 0)
                onBUSINESS_TAG = AltIsDBNull(dr.Item("BUSINESS_TAG"), 0)
                obolPropIsActive = IIf(dr.Item("PROPERTY_ACTIVE").ToString.ToUpper = "YES", True, False)
                ostrCreatedBy = AltIsDBNull(dr.Item("CREATED_BY"), String.Empty)
                odtCreatedOn = AltIsDBNull(dr.Item("CREATE_DATE"), CDate("01/01/0001"))
                ostrModifiedBy = AltIsDBNull(dr.Item("LAST_EDITED_BY"), String.Empty)
                odtModifiedOn = AltIsDBNull(dr.Item("DATE_LAST_EDITED"), CDate("01/01/0001"))
                dtDataAge = Now()
                Me.Reset()
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            nPropID = onPropID
            nPropParentID = onPropParentID
            nPropTypeID = onPropTypeID
            strPropName = ostrPropName
            strPropDesc = ostrPropDesc
            nPropPos = onPropPos
            nBUSINESS_TAG = onBUSINESS_TAG
            bolPropIsActive = obolPropIsActive
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            bolIsDirty = False
        End Sub
        Public Sub Archive()
            onPropID = nPropID
            onPropParentID = nPropParentID
            onPropTypeID = nPropTypeID
            ostrPropName = strPropName
            ostrPropDesc = strPropDesc
            onPropPos = nPropPos
            onBUSINESS_TAG = nBUSINESS_TAG
            obolPropIsActive = bolPropIsActive
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn
            bolIsDirty = False
        End Sub

#End Region
#Region "Private Operations"
        Private Sub CheckDirty()

            Dim bolOldValue As Boolean = bolIsDirty

            bolIsDirty = (nPropID <> onPropID) Or _
                        (nPropTypeID <> onPropTypeID) Or _
                        (bolPropIsActive <> obolPropIsActive) Or _
                        (nPropPos <> onPropPos) Or _
                        (strPropDesc <> ostrPropDesc) Or _
                        (strPropName <> ostrPropName) Or _
                        (strPropType <> ostrPropType)

            If bolOldValue <> bolIsDirty Then
                RaiseEvent PropertyChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            obolPropIsActive = False
            onPropID = 0
            onBUSINESS_TAG = 0
            onPropPos = 0
            onPropTypeID = 0
            ostrPropDesc = String.Empty
            ostrPropName = String.Empty
            ostrPropType = String.Empty
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        Public ReadOnly Property Key() As String
            Get
                If Me.Parent_ID = 0 Then
                    Return Me.ID
                Else
                    Return Me.Parent_ID.ToString & "|" & Me.ID
                End If
            End Get
        End Property
        Public Property ID() As String
            Get
                Return nPropID
            End Get
            Set(ByVal Value As String)
                nPropID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public ReadOnly Property Parent_ID() As String
            Get
                Return nPropParentID
            End Get
        End Property
        Public Property PropType_ID() As Long
            Get
                Return nPropTypeID
            End Get
            Set(ByVal Value As Long)
                nPropTypeID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Name() As String
            Get
                Return strPropName
            End Get
            Set(ByVal Value As String)
                strPropName = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property PropDesc() As String
            Get
                Return strPropDesc
            End Get
            Set(ByVal Value As String)
                strPropDesc = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property PropPos() As Integer
            Get
                Return nPropPos
            End Get
            Set(ByVal Value As Integer)
                nPropPos = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property BUSINESSTAG() As Integer
            Get
                Return nBUSINESS_TAG
            End Get
            Set(ByVal Value As Integer)
                nBUSINESS_TAG = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property PropIsActive() As Boolean
            Get
                Return bolPropIsActive
            End Get
            Set(ByVal Value As Boolean)
                bolPropIsActive = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Prop_Type() As String
            Get
                Return strPropType
            End Get
            Set(ByVal Value As String)
                strPropType = Value
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

        Public Property ChildProperties() As MUSTER.Info.PropertyCollection
            Get
                Return PropertiesChildCollection
            End Get
            Set(ByVal value As MUSTER.Info.PropertyCollection)
                PropertiesChildCollection = value
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
#End Region
#Region "Protected Operations"
        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
#End Region
#Region "External Event Handlers"
        Private Sub PropertiesChildCollectionChangedSub() Handles PropertiesChildCollection.InfoChanged
            RaiseEvent PropertyChanged(True)
            RaiseEvent PropertiesChildCollectionChanged(True)
        End Sub
#End Region
    End Class

End Namespace
