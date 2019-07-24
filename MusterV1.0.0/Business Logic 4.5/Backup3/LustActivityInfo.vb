'-------------------------------------------------------------------------------
' MUSTER.Info.LustActivityInfo
'   Provides the container to persist MUSTER LustActivity state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        AN        03/10/05    Original class definition.
'
' Function          Description
'-------------------------------------------------------------------------------
' New()             Instantiates an empty LustActivityInfo object.
' New(oLustEvent)   Instantiates a populated LustActivityInfo object.
' New()
'                   Instantiates a populated LustActivityInfo object.
'
' Reset()           Sets the object state to the original state when loaded from or
'                       last saved to the repository.
' Save()            Saves the object state to the repository.
' Archive()         Replaces the current value of the object to the one in collection
' CheckDirty()      Checks if the values are different in the collection and the current object
' Init()            Initializes t
'
' Attribute          Description
'-------------------------------------------------------------------------------
' ActivityID         The system id for this Lust Event Activity object
' EventID            The id of the lust event that this Activity is associated to
' Closed             The date the item was closed
' Completed          The date the item was completed
' First_GWS_Below    Date
' Second_GWS_Below   Date
' Started            Date item was started
' Type               Activity type id for the item
'
' AgeThreshold       The maximum age the info object can attain before requiring a refresh
' CreatedBy          The ID of the user that created the row
' CreatedOn          The date on which the row was created
' Deleted            Indicates the deleted state of the row
' ModifiedBy         ID of the user that last made changes
' ModifiedOn         The date of the last changes made 
'-------------------------------------------------------------------------------
'
Namespace MUSTER.Info
    ' The container for the LUST Activity
    Public Class LustActivityInfo
#Region "Public Events"
        Public Delegate Sub LustActivityInfoChangedEventHandler()
        ' Fired when CheckDirty determines that an attribute of the activity has been modified
        Public Event LustActivityInfoChanged As LustActivityInfoChangedEventHandler
#End Region
#Region "Private Member Variables"
        Private bolDeleted As Boolean
        Private dtClosed As Date
        Private dtCompleted As Date
        Private dtFirst_GWS_Below As Date
        Private dtSecond_GWS_Below As Date
        Private dtStarted As Date
        Private nActivityID As Integer
        Private nActivityType As Integer
        Private nEventID As Integer
        Private nEntityID As Integer
        Private nRem_Sys_ID As Long
        Private odtClosed As Date
        Private odtCompleted As Date
        Private odtFirst_GWS_Below As Date
        Private odtSecond_GWS_Below As Date
        Private odtStarted As Date
        Private onActivityID As Integer
        Private onEventID As Integer
        Private onActivityType As Integer
        Private onRem_Sys_ID As Long
        Private obolDeleted As Boolean

        Private ostrCreatedBy As String
        Private ostrModifiedBy As String
        Private odtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private odtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private strCreatedBy As String
        Private strModifiedBy As String
        Private dtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private dtCreatedOn As Date = DateTime.Now.ToShortDateString

        Private dtDataAge As DateTime
        Private nAgeThreshold As Int16 = 5
        Private colComments As MUSTER.Info.CommentsCollection
        Private WithEvents colDocuments As MUSTER.Info.LustDocumentCollection
        Private bolIsDirty As Boolean = False
        Private MusterException As New MUSTER.Exceptions.MusterExceptions

        Private onUserID As Integer = 0
        Private onFacilityID As Integer = 0
#End Region
#Region "Constructors"
        Public Sub New()
            'colComments = New MUSTER.Info.CommentsCollection
            colDocuments = New MUSTER.Info.LustDocumentCollection
            Me.Init()
            dtDataAge = Now()
        End Sub

        Public Sub New(ByVal ACTIVITY_ID As Integer, _
                        ByVal EVENT_ID As Integer, _
                        ByVal START_DATE As Date, _
                        ByVal FIRST_GWS_BELOW As Date, _
                        ByVal SECOND_GWS_BELOW As Date, _
                        ByVal TECH_COMPLETED_DATE As Date, _
                        ByVal CLOSED_DATE As Date, _
                        ByVal ACTIVITY_TYPE_ID As Integer, _
                        ByVal CREATED_BY As String, _
                        ByVal CREATE_DATE As String, _
                        ByVal LAST_EDITED_BY As String, _
                        ByVal DATE_LAST_EDITED As Date, _
                        ByVal DELETED As Integer, _
                        ByVal REM_SYS_ID As Int64)

            odtClosed = CLOSED_DATE
            odtCompleted = TECH_COMPLETED_DATE
            odtFirst_GWS_Below = FIRST_GWS_BELOW
            odtSecond_GWS_Below = SECOND_GWS_BELOW
            odtStarted = START_DATE
            onActivityID = ACTIVITY_ID
            onActivityType = ACTIVITY_TYPE_ID
            obolDeleted = DELETED
            ostrCreatedBy = CREATED_BY
            odtCreatedOn = CREATE_DATE
            ostrModifiedBy = LAST_EDITED_BY
            odtModifiedOn = DATE_LAST_EDITED
            onEventID = EVENT_ID
            onRem_Sys_ID = REM_SYS_ID
            dtDataAge = Now()
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Methods"
        Public Sub Reset()
            dtClosed = odtClosed
            dtCompleted = odtCompleted
            dtFirst_GWS_Below = odtFirst_GWS_Below
            dtSecond_GWS_Below = odtSecond_GWS_Below
            dtStarted = odtStarted
            nActivityID = onActivityID
            nActivityType = onActivityType
            nRem_Sys_ID = onRem_Sys_ID
            nEventID = onEventID
            bolDeleted = obolDeleted

            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn

            IsDirty = False

        End Sub
        Public Sub Archive()
            odtClosed = dtClosed
            odtCompleted = dtCompleted
            odtFirst_GWS_Below = dtFirst_GWS_Below
            odtSecond_GWS_Below = dtSecond_GWS_Below
            odtStarted = dtStarted
            onActivityID = nActivityID
            onActivityType = nActivityType
            onRem_Sys_ID = nRem_Sys_ID
            obolDeleted = bolDeleted

            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn

            IsDirty = False
        End Sub
#End Region
#Region "Private Methods"
        Private Sub Init()
            odtClosed = CDate("01/01/0001")
            odtCompleted = CDate("01/01/0001")
            odtFirst_GWS_Below = CDate("01/01/0001")
            odtSecond_GWS_Below = CDate("01/01/0001")
            odtStarted = CDate("01/01/0001")
            onActivityID = -1
            onActivityType = -1
            onRem_Sys_ID = -1
            onEventID = -1
            obolDeleted = False
            Me.Reset()
        End Sub
        Private Sub CheckDirty()
            Dim bolOldState As Boolean = bolIsDirty

            bolIsDirty = (bolDeleted <> obolDeleted) Or _
                            (dtClosed <> odtClosed) Or _
                            (dtCompleted <> odtCompleted) Or _
                            (dtFirst_GWS_Below <> odtFirst_GWS_Below) Or _
                            (dtSecond_GWS_Below <> odtSecond_GWS_Below) Or _
                            (dtStarted <> odtStarted) Or _
                            (nActivityID <> onActivityID) Or _
                            (nEventID <> onEventID) Or _
                            (nActivityType <> onActivityType) Or _
                            (nRem_Sys_ID <> onRem_Sys_ID)

            If bolOldState <> bolIsDirty Then
                RaiseEvent LustActivityInfoChanged()
            End If
        End Sub
#End Region
#Region "Protected Methods"
        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
#End Region
#Region "Exposed Attributes"
        ' The system generated ID for the LUST Activity (auto-increment in DB)
        Public Property ActivityID() As Integer
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{D3FA26B9-368B-4820-B5DE-F6A2A86A6598}
                Return nActivityID
                ' #End Region ' XDEOperation End Template Expansion{D3FA26B9-368B-4820-B5DE-F6A2A86A6598}
            End Get
            Set(ByVal Value As Integer)
                ' #Region "XDEOperation" ' Begin Template Expansion{7D68FC7C-E2B8-401A-A45B-CD81EFD2F266}
                nActivityID = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{7D68FC7C-E2B8-401A-A45B-CD81EFD2F266}
            End Set
        End Property
        Public Property EventID() As Integer
            Get
                Return Me.nEventID
            End Get
            Set(ByVal Value As Integer)
                Me.nEventID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property AgeThreshold() As Integer
            Get
                Return Me.nAgeThreshold
            End Get
            Set(ByVal Value As Integer)
                Me.nAgeThreshold = Value
            End Set
        End Property
        Public ReadOnly Property IsAgedData() As Boolean
            Get
                Return IIf(DateDiff(DateInterval.Minute, dtDataAge, Now()) >= nAgeThreshold, True, False)
            End Get
        End Property

        ' The date the LUST Activity was closed
        Public Property Closed() As Date
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{C0B53E7A-8084-4A31-8460-025F758611BF}
                Return dtClosed
                ' #End Region ' XDEOperation End Template Expansion{C0B53E7A-8084-4A31-8460-025F758611BF}
            End Get
            Set(ByVal Value As Date)
                ' #Region "XDEOperation" ' Begin Template Expansion{92FDC681-7D30-4170-AE42-3298B4B0D5EA}
                dtClosed = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{92FDC681-7D30-4170-AE42-3298B4B0D5EA}
            End Set
        End Property
        ' The date the LUST Activity was completed
        Public Property Completed() As Date
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{3BDE85E2-5429-4228-A5CB-667BF75AB2F9}
                Return dtCompleted
                ' #End Region ' XDEOperation End Template Expansion{3BDE85E2-5429-4228-A5CB-667BF75AB2F9}
            End Get
            Set(ByVal Value As Date)
                ' #Region "XDEOperation" ' Begin Template Expansion{BD072E5E-DC62-4177-BC04-374A14ED2A92}
                dtCompleted = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{BD072E5E-DC62-4177-BC04-374A14ED2A92}
            End Set
        End Property
        ' The "First GWS Below" for the LUST Activity (associated with REM and GWS activities)
        Public Property First_GWS_Below() As Date
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{DFBF20DA-3106-44CB-B706-2514DF7A11F0}
                Return dtFirst_GWS_Below
                ' #End Region ' XDEOperation End Template Expansion{DFBF20DA-3106-44CB-B706-2514DF7A11F0}
            End Get
            Set(ByVal Value As Date)
                ' #Region "XDEOperation" ' Begin Template Expansion{E17220ED-537D-48E9-AF77-CFFA90BFAADF}
                dtFirst_GWS_Below = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{E17220ED-537D-48E9-AF77-CFFA90BFAADF}
            End Set
        End Property
        ' (see Frist_GWS_Below)
        Public Property Second_GWS_Below() As Date
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{8E4B0E7D-530C-474A-A3B0-A342C1FE84DA}
                Return dtSecond_GWS_Below
                ' #End Region ' XDEOperation End Template Expansion{8E4B0E7D-530C-474A-A3B0-A342C1FE84DA}
            End Get
            Set(ByVal Value As Date)
                ' #Region "XDEOperation" ' Begin Template Expansion{A95484A7-B29C-4873-B2E4-F5AB1BA247F8}
                dtSecond_GWS_Below = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{A95484A7-B29C-4873-B2E4-F5AB1BA247F8}
            End Set
        End Property
        ' The start date for the LUST Activity
        Public Property Started() As Date
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{DA567DA1-27C2-41A4-831E-D519301186EA}
                Return dtStarted
                ' #End Region ' XDEOperation End Template Expansion{DA567DA1-27C2-41A4-831E-D519301186EA}
            End Get
            Set(ByVal Value As Date)
                ' #Region "XDEOperation" ' Begin Template Expansion{3ACC7769-F281-4D7C-B069-115C08FE2587}
                dtStarted = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{3ACC7769-F281-4D7C-B069-115C08FE2587}
            End Set
        End Property
        ' The type of the LUST Activity (from tblSYS_PROPERTY_MASTER)
        Public Property Type() As Integer
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{94C44F8D-28D3-4219-9489-B687D92C6D88}
                Return nActivityType
                ' #End Region ' XDEOperation End Template Expansion{94C44F8D-28D3-4219-9489-B687D92C6D88}
            End Get
            Set(ByVal Value As Integer)
                ' #Region "XDEOperation" ' Begin Template Expansion{9166BCF3-BAC8-4F5C-BBC2-CB044E033257}
                nActivityType = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{9166BCF3-BAC8-4F5C-BBC2-CB044E033257}
            End Set
        End Property
        ' The deleted flag for the LUST Activity
        Public Property Deleted() As Boolean
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{F476C393-899C-4EEB-95B4-64F2FACC24F5}
                Return bolDeleted
                ' #End Region ' XDEOperation End Template Expansion{F476C393-899C-4EEB-95B4-64F2FACC24F5}
            End Get
            Set(ByVal Value As Boolean)
                ' #Region "XDEOperation" ' Begin Template Expansion{7492E909-CF10-4092-8235-6C64FD66AEB4}
                bolDeleted = Value
                Me.CheckDirty()
                ' #End Region ' XDEOperation End Template Expansion{7492E909-CF10-4092-8235-6C64FD66AEB4}
            End Set
        End Property
        ' The comments associated with the LUST Activity
        Public Property Comments() As MUSTER.Info.CommentsCollection
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{166F7FBB-72AF-472C-B19C-E72A317F69A9}
                Return colComments
                ' #End Region ' XDEOperation End Template Expansion{166F7FBB-72AF-472C-B19C-E72A317F69A9}
            End Get
            Set(ByVal Value As MUSTER.Info.CommentsCollection)
                ' #Region "XDEOperation" ' Begin Template Expansion{D85DD6C0-6A26-4BA4-9EBD-3D295FB3A56F}
                colComments = Value
                ' #End Region ' XDEOperation End Template Expansion{D85DD6C0-6A26-4BA4-9EBD-3D295FB3A56F}
            End Set
        End Property
        ' The documents associated with the LUST Activity
        Public Property Documents() As MUSTER.Info.LustDocumentCollection
            Get
                Return colDocuments
            End Get
            Set(ByVal Value As MUSTER.Info.LustDocumentCollection)
                colDocuments = Value
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

        Public Property UserID() As Integer
            Get
                Return onUserID
            End Get
            Set(ByVal Value As Integer)
                onUserID = Value
            End Set
        End Property

        Public Property FacilityID() As Integer
            Get
                Return onFacilityID
            End Get
            Set(ByVal Value As Integer)
                onFacilityID = Value
            End Set
        End Property

        Public Property RemSystemID() As Integer
            Get
                Return onRem_Sys_ID
            End Get
            Set(ByVal Value As Integer)
                onRem_Sys_ID = Value
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
#Region "Protected Attributes"
        ' The Entity ID associated with a LUST Activity (from tblSYS_ENTITY)
        Protected ReadOnly Property EntityID() As Integer
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{EC81C992-5189-4468-9CB9-B50B0D0B5906}
                Return nEntityID
                ' #End Region ' XDEOperation End Template Expansion{EC81C992-5189-4468-9CB9-B50B0D0B5906}
            End Get
        End Property
#End Region
#Region "Private Attributes"
        ' The system generated ID of the remediation system associated with the LUST Activity
        Private Property Rem_Sys_ID() As Long
            Get
                ' #Region "XDEOperation" ' Begin Template Expansion{E677D852-26FA-4B0A-8F1F-55D94098665F}
                Return nRem_Sys_ID
                ' #End Region ' XDEOperation End Template Expansion{E677D852-26FA-4B0A-8F1F-55D94098665F}
            End Get
            Set(ByVal Value As Long)
                ' #Region "XDEOperation" ' Begin Template Expansion{FF64EEF1-9280-4963-881A-57910850E9C8}
                nRem_Sys_ID = Value
                ' #End Region ' XDEOperation End Template Expansion{FF64EEF1-9280-4963-881A-57910850E9C8}
            End Set
        End Property
#End Region
    End Class
End Namespace
