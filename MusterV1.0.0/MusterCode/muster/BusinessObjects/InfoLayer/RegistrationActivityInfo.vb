
Namespace MUSTER.Info
    ' -------------------------------------------------------------------------------
    '    MUSTER.Info.RegistrationInfo
    '          Provides the container for Registration Activity Informaiton
    ' 
    '    Copyright (C) 2004 CIBER, Inc.
    '    All rights reserved.
    ' 
    '    Release   Initials    Date        Description
    '       1.0        JVC2      02/08/2005  Original framework from Rational XDE.
    '       1.1        AB        02/22/05    Added AgeThreshold and IsAgedData Attributes
    ' 
    '    Operations
    '    Function          Description
    '  
    '    Read-Only Attributes
    '    CreatedBy         The name of the user that created the ProfileInfo object.
    '    CreatedOn         The date that the ProfileInfo object was created.
    '    ModifiedBy        The name of the user that last modified the ProfileInfo object.
    '    ModifiedOn        The date that the ProfileInfo object was last modified.
    ' -------------------------------------------------------------------------------
    Public Class RegistrationActivityInfo
#Region "Private Member Variables"
        Private bolDeleted As Boolean
        Private bolIsDirty As Boolean
        Private bolProcessed As Boolean
        Private dtCreatedOn As Date
        Private dtDateAdded As Date
        Private dtDateProcessed As Date
        Private dtModifiedOn As Date
        Private nEntityID As Integer
        Private nEntityType As Integer
        Private nRegActionIndex As Integer
        Private strRegistrationActivity As String
        Private nRegistrationID As Integer
        Private obolDeleted As Boolean
        Private obolIsDirty As Boolean
        Private obolProcessed As Boolean
        Private odtDateAdded As Date
        Private odtDateProcessed As Date
        Private onEntityID As Integer
        Private onEntityType As Integer
        Private onRegActionIndex As Integer
        Private ostrRegistrationActivity As String
        Private oRegistrationID As Integer
        Private ostrUserID As String
        Private strCreatedBy As String
        Private strModifiedBy As String
        Private strUserID As String
        Private dtDataAge As DateTime
        Private nAgeThreshold As Int16 = 5
        Private nCalendarID As Integer
        Private onCalendarID As Integer
#End Region
#Region "Public Events"
        Public Delegate Sub RegistrationInfoChangedEventHandler(ByVal bolDirtyState As Boolean)
        ' Raised when any of the attributes of the RegistrationInfo is altered.
        Public Event RegistrationInfoChanged As RegistrationInfoChangedEventHandler
#End Region
#Region "Constructors"
        Public Sub New()
            dtDataAge = Now()
        End Sub
        Public Sub New(ByVal RegActionIndex As Integer, _
                        ByVal RegID As Integer, _
                        ByVal EntityType As Integer, _
                        ByVal EntityID As Integer, _
                        ByVal UserID As String, _
                        ByVal RegistrationActivity As String, _
                        ByVal Processed As Boolean, _
                        ByVal DateAdded As DateTime, _
                        ByVal CalID As Integer)
            oRegistrationID = RegID
            onEntityType = EntityType
            onEntityID = EntityID
            ostrUserID = UserID
            ostrRegistrationActivity = RegistrationActivity
            onRegActionIndex = RegActionIndex
            odtDateAdded = DateAdded
            obolProcessed = Processed
            onCalendarID = CalID
            dtDataAge = Now()
            Reset()
        End Sub
        Sub New(ByVal drTemplate As DataRow)
            Try
                oRegistrationID = drTemplate.Item("RegID")
                onEntityType = drTemplate.Item("EntityType")
                onEntityID = drTemplate.Item("EntityID")
                ostrUserID = drTemplate.Item("UserID")
                ostrRegistrationActivity = drTemplate.Item("RegActivity")
                'nRegActionIndex = drTemplate.Item("COMPLETED")
                odtDateAdded = drTemplate.Item("DateAdded")
                obolProcessed = drTemplate.Item("Processed")
                '********************************************************
                '
                ' Other private member variables for prior state here
                '
                '********************************************************
                obolDeleted = drTemplate.Item("DELETED")
                strCreatedBy = drTemplate.Item("CREATED_BY")
                dtCreatedOn = drTemplate.Item("DATE_CREATED")
                strModifiedBy = drTemplate.Item("LAST_EDITED_BY")
                dtModifiedOn = drTemplate.Item("DATE_LAST_EDITED")
                dtDataAge = Now()
                onCalendarID = IIf(drTemplate.Item("CALENDAR_INFO_ID") Is DBNull.Value, 0, drTemplate.Item("CALENDAR_INFO_ID"))
                Me.Reset()
            Catch ex As Exception
                'MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "Exposed Attributes"
        ' Gets/Sets the Registration Activity Value
        Public Property ActivityDesc() As String
            Get
                Return strRegistrationActivity
            End Get
            Set(ByVal Value As String)
                strRegistrationActivity = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property RegistrationID() As Integer
            Get
                Return nRegistrationID
            End Get
            Set(ByVal Value As Integer)
                nRegistrationID = Value
                Me.CheckDirty()
            End Set
        End Property
        ' Gets/Sets the Entity Type for the registration activity
        Public Property EntityType() As Integer
            Get
                Return nEntityType
            End Get
            Set(ByVal Value As Integer)
                nEntityType = Value
                Me.CheckDirty()
            End Set
        End Property
        ' Gets/Sets the Entity ID associated with the registration activity
        Public Property EntityId() As Integer
            Get
                Return nEntityID
            End Get
            Set(ByVal Value As Integer)
                nEntityID = Value
                Me.CheckDirty()
            End Set
        End Property
        ' Gets/Sets the User ID associated with the registration activity
        Public Property UserID() As String
            Get
                Return strUserID
            End Get
            Set(ByVal Value As String)
                strUserID = Value
                Me.CheckDirty()
            End Set
        End Property
        ' Gets/Sets the boolean indicating whether or not the registration activity has been processed
        Public Property Processed() As Boolean
            Get
                Return bolProcessed
            End Get
            Set(ByVal Value As Boolean)
                bolProcessed = Value
                Me.CheckDirty()
            End Set
        End Property
        ' Gets/Sets the date on which the registration activity was processed
        Public Property DateAdded() As DateTime
            Get
                Return dtDateAdded
            End Get
            Set(ByVal Value As DateTime)
                dtDateAdded = Value
                Me.CheckDirty()
            End Set
        End Property

        ' Gets the string indicating the user that created the activity
        Public ReadOnly Property CreatedBy() As String
            Get
                Return strCreatedBy
            End Get
        End Property
        ' Gets the date on which the activity was created
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return dtCreatedOn
            End Get
        End Property
        ' Gets/Sets the date on which the registration activity was added
        Public Property DateAdded(ByVal dtDate As Date) As Date
            Get
                Return dtDateAdded
            End Get
            Set(ByVal Value As Date)
                dtDateAdded = Value
            End Set
        End Property

        ' Gets/Sets the deleted state for the registration activity
        Public Property Deleted() As Boolean
            Get
                Return Me.bolDeleted
            End Get
            Set(ByVal Value As Boolean)
                Me.bolDeleted = Value
                Me.CheckDirty()
            End Set
        End Property

        ' Gets the dirty state of the info object
        Public ReadOnly Property IsDirty() As Boolean
            Get
                Return Me.bolIsDirty
            End Get
        End Property
        ' Gets the string indicating the last user to modify the activity
        Public ReadOnly Property ModifiedBy() As String
            Get
                Return strModifiedBy
            End Get
        End Property
        ' Gets the date on which the activity was last modified
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return dtModifiedOn
            End Get
        End Property


        'Gets/Sets the Registration Action Index
        Public Property RegActionIndex() As Long
            Get
                Return nRegActionIndex
            End Get
            Set(ByVal Value As Long)
                nRegActionIndex = Value
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

        Public Property CalendarID() As Integer
            Get
                Return nCalendarID
            End Get
            Set(ByVal Value As Integer)
                nCalendarID = Value
                Me.CheckDirty()
            End Set
        End Property
#End Region
#Region "Exposed Operations"
        ' Resets the info attributes to their original states
        Public Sub Reset()
            nRegActionIndex = onRegActionIndex
            nRegistrationID = oRegistrationID
            nEntityType = onEntityType
            nEntityID = onEntityID
            strUserID = ostrUserID
            strRegistrationActivity = ostrRegistrationActivity
            dtDateAdded = odtDateAdded
            bolProcessed = obolProcessed
            nCalendarID = onCalendarID
            '********************************************************
            '
            ' Other assignments of current state to prior state here
            '
            '********************************************************
            bolDeleted = obolDeleted
            bolIsDirty = False
            RaiseEvent RegistrationInfoChanged(bolIsDirty)
        End Sub
        ' Sets the baseline private members to the current state private members
        ' Performs a standard copy
        Public Sub Archive()
            onRegActionIndex = nRegActionIndex
            oRegistrationID = nRegistrationID
            onEntityType = nEntityType
            onEntityID = nEntityID
            ostrUserID = strUserID
            ostrRegistrationActivity = strRegistrationActivity
            odtDateAdded = dtDateAdded
            obolProcessed = bolProcessed
            '********************************************************
            '
            ' Other assignments of current state to prior state here
            '
            '********************************************************
            obolDeleted = bolDeleted
            onCalendarID = nCalendarID
            bolIsDirty = False
        End Sub
#End Region
#Region "Protected Operations"
        ' Overrides the base finalize method
        Protected Overrides Sub Finalize()
        End Sub
#End Region
#Region "Private Operations"
        ' Checks all attributes and returns boolean indicating if any data has changed
        Private Function CheckDirty() As Boolean
            Dim obolIsDirty As Boolean = bolIsDirty


            bolIsDirty = (nRegistrationID <> oRegistrationID) Or _
            (nRegActionIndex <> onRegActionIndex) Or _
            (nEntityType <> onEntityType) Or _
            (nEntityID <> onEntityID) Or _
            (strUserID <> ostrUserID) Or _
            (strRegistrationActivity <> ostrRegistrationActivity) Or _
            (dtDateAdded <> odtDateAdded) Or _
            (bolProcessed <> obolProcessed) Or _
            (bolDeleted <> obolDeleted) Or _
            (nCalendarID <> onCalendarID)


            If obolIsDirty <> bolIsDirty Then
                RaiseEvent RegistrationInfoChanged(bolIsDirty)
            End If
        End Function
        ' Initializes the info structure
        Private Sub Init()
            onRegActionIndex = 0
            oRegistrationID = 0
            onEntityType = 0
            onEntityID = 0
            ostrUserID = 0
            ostrRegistrationActivity = 0
            odtDateAdded = System.DateTime.Now
            obolProcessed = False
            '********************************************************
            '
            ' Other assignments of current state to empty/false/etc here
            '
            '********************************************************
            obolDeleted = False
            strCreatedBy = String.Empty
            dtCreatedOn = System.DateTime.Now
            strModifiedBy = String.Empty
            dtModifiedOn = System.DateTime.Now
            onCalendarID = 0
            Me.Reset()
        End Sub
#End Region
    End Class
End Namespace
