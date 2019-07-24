'-------------------------------------------------------------------------------
' MUSTER.Info.CalendarInfo
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        KJ      12/10/04    Original class definition.
'  1.1        KJ      12/23/04    Added the Archive Method and also added more description to header.
'  1.2        AN      12/30/04    Added Try catch and Exception Handling/Logging
'  1.3        KJ      01/06/05    Added Events for InfoBecameDirty.
'  1.4        AB      02/17/05    Added AgeThreshold and IsAgedData Attributes
'  1.5        JVC2    03/25/05    Added OwningEntityType and OwningEntityID so that any object can inspect the
'                                  calendar and determine which entries it originated.
'
'   
' Function                  Description
'  New()                Instantiates an empty CalendarInfo object.
'  New(nCalendarInfoId, dtNotificationDate, dtDateDue, nCurrentColorCode, strTaskDescription, strUserId, strSourceUserId, strGroupId, bolDueToMe, bolToDo, bolCompleted, bolDeleted, CreatedBy, CreatedOn, ModifiedBy, ModifiedOn)
'                       Instantiates a populated CalendarInfo object.
'  New(dr)              Instantiates a populated CalendarInfo object taking member state
'                           from the datarow provided
' Archive()             Sets the object state to the new state 
'  Reset()              Sets the object state to the original state when loaded from or last saved to the repository.
'
'Attribute                      Description
'  CalendarInfoId       The unique identifier associated with the Calendar in the repository.
'  NotificationDate     The Date at which a Notification has to be given to a user
'  DateDue              The date at which the event will be due
'  CurrentColorCode     The color code associated with the calendar event
'  TaskDescription      The description of the calendar event
'  UserID               The user ID asssociated with the calendar event
'  SourceUserID         The Source ID associated with the calendar event
'  GroupID              The Group to which the calendar event belongs
'  DueToMe              The Boolean value indicating all the calendar events due to me
'  ToDo                 The Boolean value indicating all the calendar events To Do
'  Completed            The Boolean value indicating all the calendar events completed
'  Deleted              The Boolean value indicating all the calendar events which are marked deleted
'  IsDirty              Indicates if the Calendar state has been altered since it was
'                           last loaded from or saved to the repository.
'  AgeThreshold         Indicates the number of minutes old data can be before it should be 
'                           refreshed from the DB.  Data should only be refreshed when Retrieved
'                           and when IsDirty is false
'  IsAgedData           Will return true if the data has been held longer than the AgeThreshold
'
'-------------------------------------------------------------------------------
'
' TODO - Remove CreatedBy, CreatedOn, ModifiedBy, ModifiedOn from Constructor - JVC 2
'
Namespace MUSTER.Info
    <Serializable()> _
Public Class CalendarInfo

#Region "Private member variables"
        'Original Values
        Private onCalendarInfoId As Integer = 0       'added new
        Private odtNotificationDate As DateTime
        Private odtDateDue As DateTime
        Private onCurrentColorCode As Integer
        Private ostrTaskDescription As String = String.Empty
        Private ostrUserID As String = String.Empty
        Private ostrSourceUserID As String = String.Empty
        Private ostrGroupID As String = String.Empty
        Private obolDueToMe As Boolean
        Private obolToDo As Boolean
        Private obolCompleted As Boolean
        Private obolDeleted As Boolean
        Private onOwningEntityType As Long = 0
        Private onOwningEntityID As Long = 0

        'Current Values
        Private nCalendarInfoId As Integer = 0       'added new
        Private dtNotificationDate As DateTime
        Private dtDateDue As DateTime
        Private nCurrentColorCode As Integer
        Private strTaskDescription As String = String.Empty
        Private strUserID As String = String.Empty
        Private strSourceUserID As String = String.Empty
        Private strGroupID As String = String.Empty
        Private bolDueToMe As Boolean
        Private bolToDo As Boolean
        Private bolCompleted As Boolean
        Private bolDeleted As Boolean
        Private nOwningEntityType As Long = 0
        Private nOwningEntityID As Long = 0

        Private strCreatedBy As String = String.Empty
        Private dtCreatedOn As DateTime = DateTime.Now.ToShortDateString
        Private strModifiedBy As String = String.Empty
        Private dtModifiedOn As DateTime = DateTime.Now.ToShortDateString
        Private dtDataAge As DateTime

        Private ostrCreatedBy As String = String.Empty
        Private odtCreatedOn As DateTime = DateTime.Now.ToShortDateString
        Private ostrModifiedBy As String = String.Empty
        Private odtModifiedOn As DateTime = DateTime.Now.ToShortDateString

        Private nAgeThreshold As Int16 = 5

        Private bolIsDirty As Boolean = False
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Public Events"
        Public Event InfoBecameDirty(ByVal DirtyState As Boolean)
#End Region

#Region "Constructors"
        Sub New()
            MyBase.New()
            dtDataAge = Now()
        End Sub
        Sub New(ByVal CalendarInfoId As Integer, _
            ByVal NotificationDate As Date, _
            ByVal DateDue As Date, _
            ByVal CurrentColorCode As Integer, _
            ByVal TaskDescription As String, _
            ByVal UserId As String, _
            ByVal SourceUserId As String, _
            ByVal GroupId As String, _
            ByVal DueToMe As Boolean, _
            ByVal ToDo As Boolean, _
            ByVal Completed As Boolean, _
            ByVal Deleted As Boolean, _
            ByVal CreatedBy As String, _
            ByVal CreatedOn As Date, _
            ByVal ModifiedBy As String, _
            ByVal LastEdited As Date, _
            Optional ByVal OwningEntityType As Int32 = 0, _
            Optional ByVal OwningEntityID As Int64 = 0)
            onCalendarInfoId = CalendarInfoId
            odtNotificationDate = NotificationDate
            odtDateDue = DateDue
            onCurrentColorCode = CurrentColorCode
            ostrTaskDescription = TaskDescription
            ostrUserID = UserId
            ostrSourceUserID = SourceUserId
            ostrGroupID = GroupId
            obolDueToMe = DueToMe
            obolToDo = ToDo
            obolCompleted = Completed
            obolDeleted = Deleted
            ostrCreatedBy = CreatedBy
            odtCreatedOn = CreatedOn
            ostrModifiedBy = ModifiedBy
            odtModifiedOn = LastEdited
            dtDataAge = Now()
            onOwningEntityType = OwningEntityType
            onOwningEntityID = OwningEntityID

            Me.Reset()
        End Sub
        Sub New(ByVal drAddress As DataRow)
            Try
                onCalendarInfoId = drAddress.Item("CALENDAR_INFO_ID")
                odtNotificationDate = drAddress.Item("NOTIFICATION_DATE")
                odtDateDue = drAddress.Item("DATE_DUE")
                onCurrentColorCode = drAddress.Item("CURRENT_COLOR_CODE")
                ostrTaskDescription = drAddress.Item("TASK_DESCRIPTION")
                ostrUserID = drAddress.Item("USER_ID")
                ostrSourceUserID = drAddress.Item("SOURCE_USER_ID")
                ostrGroupID = drAddress.Item("GROUP_ID")
                obolDueToMe = drAddress.Item("DUE_TO_ME")
                obolToDo = drAddress.Item("TO_DO")
                obolCompleted = drAddress.Item("COMPLETED")
                obolDeleted = Not drAddress.Item("DELETED")
                ostrCreatedBy = drAddress.Item("CREATED_BY")
                odtCreatedOn = drAddress.Item("DATE_CREATED")
                ostrModifiedBy = drAddress.Item("LAST_EDITED_BY")
                odtModifiedOn = drAddress.Item("DATE_LAST_EDITED")
                dtDataAge = Now()
            Catch Ex As Exception
                MusterException.Publish(Ex, Nothing, Nothing)
                Throw Ex
            End Try
        End Sub
#End Region

#Region "Exposed Operations"

        Public Sub Reset()
            nCalendarInfoId = onCalendarInfoId
            dtNotificationDate = odtNotificationDate
            dtDateDue = odtDateDue
            nCurrentColorCode = onCurrentColorCode
            strTaskDescription = ostrTaskDescription
            strUserID = ostrUserID
            strSourceUserID = ostrSourceUserID
            strGroupID = ostrGroupID
            bolDueToMe = obolDueToMe
            bolToDo = obolToDo
            bolCompleted = obolCompleted
            bolDeleted = obolDeleted
            bolIsDirty = False
            nOwningEntityType = onOwningEntityType
            nOwningEntityID = onOwningEntityID

            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn

        End Sub

        Public Sub Archive()
            onCalendarInfoId = nCalendarInfoId
            odtNotificationDate = dtNotificationDate
            odtDateDue = dtDateDue
            onCurrentColorCode = nCurrentColorCode
            ostrTaskDescription = strTaskDescription
            ostrUserID = strUserID
            ostrSourceUserID = strSourceUserID
            ostrGroupID = strGroupID
            obolDueToMe = bolDueToMe
            obolToDo = bolToDo
            obolCompleted = bolCompleted
            obolDeleted = bolDeleted
            bolIsDirty = False
            onOwningEntityType = nOwningEntityType
            onOwningEntityID = nOwningEntityID

            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = odtModifiedOn

        End Sub
#End Region

#Region "Private Operations"
        Private Sub CheckDirty()
            Dim bolOldState As Boolean = bolIsDirty
            bolIsDirty = (nCalendarInfoId <> onCalendarInfoId Or _
                        dtNotificationDate <> odtNotificationDate Or _
                        dtDateDue <> odtDateDue Or _
                        nCurrentColorCode <> onCurrentColorCode Or _
                        strTaskDescription <> ostrTaskDescription Or _
                        strUserID <> ostrUserID Or _
                        strSourceUserID <> ostrSourceUserID Or _
                        strGroupID <> ostrGroupID Or _
                        bolDueToMe <> bolDueToMe Or _
                        bolToDo <> obolToDo Or _
                        bolCompleted <> obolCompleted Or _
                        bolDeleted <> obolDeleted Or _
                        nOwningEntityType <> onOwningEntityType Or _
                        nOwningEntityID <> nOwningEntityID)
            If bolOldState <> bolIsDirty Then
                RaiseEvent InfoBecameDirty(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            onCalendarInfoId = 0
            odtNotificationDate = System.DateTime.Now
            odtDateDue = System.DateTime.Now
            onCurrentColorCode = 0
            ostrTaskDescription = String.Empty
            ostrUserID = String.Empty
            ostrSourceUserID = String.Empty
            ostrGroupID = String.Empty
            obolDueToMe = False
            obolToDo = False
            obolCompleted = False
            obolDeleted = False
            dtCreatedOn = System.DateTime.Now
            dtModifiedOn = System.DateTime.Now
            strCreatedBy = String.Empty
            strModifiedBy = String.Empty
            onOwningEntityType = 0
            onOwningEntityID = 0
            Me.Reset()
        End Sub
#End Region

#Region "Exposed Attributes"
        Public Property CalendarInfoId() As Integer
            Get
                Return nCalendarInfoId
            End Get

            Set(ByVal value As Integer)
                nCalendarInfoId = Integer.Parse(value)
                Me.CheckDirty()
            End Set
        End Property

        Public Property OwningEntityType() As Integer
            Get
                Return nOwningEntityType
            End Get
            Set(ByVal Value As Integer)
                nOwningEntityType = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property OwningEntityID() As Long
            Get
                Return nOwningEntityID
            End Get
            Set(ByVal Value As Long)
                nOwningEntityID = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property NotificationDate() As DateTime
            Get
                Return dtNotificationDate
            End Get

            Set(ByVal value As DateTime)
                dtNotificationDate = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property DateDue() As DateTime
            Get
                Return dtDateDue
            End Get

            Set(ByVal value As DateTime)
                dtDateDue = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property CurrentColorCode() As Integer
            Get
                Return nCurrentColorCode
            End Get

            Set(ByVal value As Integer)
                nCurrentColorCode = Integer.Parse(value)
                Me.CheckDirty()
            End Set
        End Property

        Public Property TaskDescription() As String
            Get
                Return strTaskDescription
            End Get

            Set(ByVal value As String)
                strTaskDescription = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property UserId() As String
            Get
                Return strUserID
            End Get

            Set(ByVal value As String)
                strUserID = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property SourceUserId() As String
            Get
                Return strSourceUserID
            End Get

            Set(ByVal value As String)
                strSourceUserID = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property GroupId() As String
            Get
                Return strGroupID
            End Get

            Set(ByVal value As String)
                strGroupID = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property DueToMe() As Boolean
            Get
                Return bolDueToMe
            End Get

            Set(ByVal value As Boolean)
                bolDueToMe = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property ToDo() As Boolean
            Get
                Return bolToDo
            End Get

            Set(ByVal value As Boolean)
                bolToDo = value
                Me.CheckDirty()
            End Set
        End Property

        Public Property Completed() As Boolean
            Get
                Return bolCompleted
            End Get

            Set(ByVal value As Boolean)
                bolCompleted = value
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
    End Class

End Namespace


