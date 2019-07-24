'-------------------------------------------------------------------------------
' MUSTER.Info.TicklerMessageInfo
'   Provides the container to persist MUSTER Tickler Message info 
'
' Copyright (C) 2009 CIBER, Inc.
' All rights reserved.
'
' Release        Initials         Date        Description
'  1.0        Thomas Franey       05/29/09    Original class definition.
'
' Function          Description
'-------------------------------------------------------------------------------
'

Namespace MUSTER.Info
    Public Class TicklerMessageInfo
#Region "Private Member Variables"

        Private strFromID As String
        Private strToID As String
        Private nMsgID As String
        Private bolRead As Boolean
        Private bolCompleted As Boolean
        Private bolIsIssue As Boolean
        Private strSubject As String
        Private strMessage As String
        Private nModuleID As Integer
        Private strObjectID As String
        Private strKeyword As String
        Private strImageFile As String
        Private dtPostDate As DateTime
        Private dtDateCompleted As DateTime
        Private dtDateRead As DateTime


        Private oStrFromID As String
        Private oStrToID As String
        Private oBolRead As Boolean
        Private oBolCompleted As Boolean
        Private oBolIsIssue As Boolean
        Private oStrSubject As String
        Private oStrMessage As String
        Private onModuleID As Integer
        Private oStrObjectID As String
        Private oStrKeyword As String
        Private oStrImageFile As String
        Private odtPostDate As DateTime
        Private odtDateCompleted As DateTime
        Private odtDateRead As DateTime


        Private bolIsDirty As Boolean

        Private dtCreatedOn As Date = DateTime.Now.ToShortDateString
#End Region
#Region "Public Events"
        Public Delegate Sub InfoChangedEventHandler()
        Public Event InfoChanged As InfoChangedEventHandler
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()

            Init()
            Archive()

        End Sub

        Public Sub New(ByVal Id As String, _
            ByVal fromID As String, _
            ByVal toID As String, _
            ByVal read As Boolean, _
            ByVal completed As Boolean, _
            ByVal isIssue As Boolean, _
            ByVal subject As String, _
            ByVal message As String, _
            ByVal moduleID As Integer, _
            ByVal objectID As String, _
            ByVal keyword As String, _
            ByVal imageFile As String, _
            ByVal CreateDate As Date, _
            ByVal PostDate As Date, _
            ByVal dateRead As DateTime, _
            ByVal dateCompleted As DateTime)


            MyBase.new()

            nMsgID = Id
            oStrFromID = fromID
            oStrToID = toID
            oBolRead = read
            oBolCompleted = completed
            oBolIsIssue = isIssue
            oStrSubject = subject
            oStrMessage = message
            onModuleID = moduleID
            oStrObjectID = objectID
            oStrKeyword = keyword
            oStrImageFile = imageFile
            odtDateCompleted = dateCompleted
            odtDateRead = dateRead
            odtPostDate = PostDate


            dtCreatedOn = CreateDate



            Me.Reset()

        End Sub
#End Region

#Region "Exposed Attributes"
        ' the message read flag for the SYS_TICKLER
        Public Property Read() As Boolean
            Get
                Return bolRead
            End Get
            Set(ByVal Value As Boolean)
                bolRead = Value
                CheckDirty()
            End Set
        End Property


        ' the message completed flag for the SYS_TICKLER
        Public Property Completed() As Boolean
            Get
                Return bolCompleted
            End Get
            Set(ByVal Value As Boolean)
                bolCompleted = Value
                CheckDirty()
            End Set
        End Property

        ' the 'message reported to issue tracker' flag for the SYS_TICKLER
        Public Property IsIssue() As Boolean
            Get
                Return bolIsIssue
            End Get
            Set(ByVal Value As Boolean)
                bolIsIssue = Value
                CheckDirty()
            End Set
        End Property

        ' The date on which the row was created
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return dtCreatedOn
            End Get
        End Property

        ' The date on which was last read
        Public ReadOnly Property DateRead() As DateTime
            Get
                Return dtDateRead
            End Get
        End Property

        ' The date on which was completed (task is done by user)
        Public ReadOnly Property DateCompleted() As DateTime
            Get
                Return dtDateCompleted
            End Get
        End Property

        ' The date set to be posted
        Public Property PostDate() As DateTime
            Get
                Return dtPostDate
            End Get

            Set(ByVal Value As DateTime)
                dtPostDate = Value
            End Set
        End Property


        ' The system ID for this Tickler Message
        Public Property ID() As String
            Get
                Return nMsgID
            End Get
            Set(ByVal Value As String)
                nMsgID = Value
                CheckDirty()

            End Set
        End Property

        'message creator ID (user ID or SYSTEM)
        Public Property FromID() As String
            Get
                Return strFromID
            End Get
            Set(ByVal Value As String)
                strFromID = Value
                CheckDirty()
            End Set

        End Property

        'receiver ID (User ID or user's group ID)
        Public Property toID() As String
            Get
                Return strToID
            End Get

            Set(ByVal Value As String)
                strToID = Value
                CheckDirty()
            End Set
        End Property


        'tickler message Subject
        Public Property Subject() As String
            Get
                Return strSubject
            End Get

            Set(ByVal Value As String)
                strSubject = Value
                CheckDirty()
            End Set
        End Property

        'Message text
        Public Property Message() As String
            Get
                Return strMessage
            End Get

            Set(ByVal Value As String)
                strMessage = Value
                CheckDirty()
            End Set
        End Property


        'Object ID of the entity of message reference
        Public Property ObjectID() As String
            Get
                Return strObjectID
            End Get

            Set(ByVal Value As String)
                strObjectID = Value
                CheckDirty()
            End Set
        End Property

        'keyword referencing the entity type of the entity being referenced for this message
        Public Property Keyword() As String
            Get
                Return strKeyword
            End Get

            Set(ByVal Value As String)
                strKeyword = Value
                CheckDirty()
            End Set
        End Property


        ' The name of the file holding a possible clip image referencing the message
        Public Property ImageFile() As String
            Get
                Return strImageFile
            End Get
            Set(ByVal Value As String)
                strImageFile = Value
                CheckDirty()
            End Set
        End Property

        ' The module id of message reference
        Public Property ModuleID() As Integer
            Get
                Return nModuleID
            End Get
            Set(ByVal Value As Integer)
                nModuleID = Value
                CheckDirty()
            End Set
        End Property

#End Region
#Region "Exposed Methods"

        Public Sub Archive()

            oStrFromID = strFromID
            oStrToID = strToID
            oBolRead = bolRead
            oBolCompleted = bolCompleted
            oBolIsIssue = bolIsIssue
            oStrSubject = strSubject
            oStrMessage = strMessage
            onModuleID = nModuleID
            oStrObjectID = strObjectID
            oStrKeyword = strKeyword
            oStrImageFile = strImageFile
            odtDateCompleted = dtDateCompleted
            odtDateRead = dtDateRead
            odtPostDate = dtPostDate


            bolIsDirty = False

        End Sub
        Public Sub Reset()

            strFromID = oStrFromID
            strToID = oStrToID
            bolRead = oBolRead
            bolCompleted = oBolCompleted
            bolIsIssue = oBolIsIssue
            strSubject = oStrSubject
            strMessage = oStrMessage
            nModuleID = onModuleID
            strObjectID = oStrObjectID
            strKeyword = oStrKeyword
            strImageFile = oStrImageFile
            dtDateCompleted = odtDateCompleted
            dtDateRead = odtDateRead
            dtPostDate = odtPostDate


            bolIsDirty = False

        End Sub

        Public ReadOnly Property IsDirty() As Boolean
            Get
                Return Me.bolIsDirty
            End Get
        End Property
#End Region
#Region "Private Methods"
        Private Sub CheckDirty()
            Dim bolOldState As Boolean = bolIsDirty

            bolIsDirty = (strFromID <> oStrFromID) Or _
            (strToID <> oStrToID) Or _
            (bolRead <> oBolRead) Or _
            (bolCompleted <> oBolCompleted) Or _
            (bolIsIssue <> oBolIsIssue) Or _
            (strSubject <> oStrSubject) Or _
            (strMessage <> oStrMessage) Or _
            (nModuleID <> onModuleID) Or _
            (strObjectID <> oStrObjectID) Or _
            (strKeyword <> oStrKeyword) Or _
            (strImageFile <> oStrImageFile) Or _
            (odtDateCompleted <> dtDateCompleted) Or _
            (odtDateRead <> dtDateRead) Or _
            (odtPostDate <> dtPostDate)


        End Sub

        Sub Init()

            strFromID = "SYSTEM"
            strToID = String.Empty
            bolRead = False
            bolCompleted = False
            bolIsIssue = False
            strSubject = "New Message"
            strMessage = String.Empty
            nModuleID = -1
            strObjectID = String.Empty
            strKeyword = String.Empty
            strImageFile = String.Empty
            dtCreatedOn = DateTime.Now.ToShortDateString
            odtDateCompleted = Nothing
            odtDateRead = Nothing
            odtPostDate = Nothing
            nMsgID = String.Empty

        End Sub
#End Region
#Region "Protected Methods"
        Protected Overrides Sub Finalize()
        End Sub
#End Region
    End Class
End Namespace
