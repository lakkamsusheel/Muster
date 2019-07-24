'-------------------------------------------------------------------------------
' MUSTER.Info.FlagInfo
'   Provides the container to persist MUSTER Flag state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        MNR       12/13/04    Original class definition.
'  1.1        AN        12/30/04    Added Try catch and Exception Handling/Logging
'  1.4       JVC2       01/26/2005  Added SourceUserID attribute, strSourceUser, ostrSourceUser
'                                       updated CheckDirty to include check on new attribute.
'  1.5       JVC2       01/31/2005  Added code to accomodate new column SOURCE_USER_ID
'  1.6       JVC2       02/03/2005  Added CreatedBy, CreatedOn, ModifiedBy, ModifiedOn to NEW()
'  1.7        AB        02/18/05    Added AgeThreshold and IsAgedData Attributes
'
' Function          Description
' New()             Instantiates an empty FlagInfo object
' New(flagID, entityID, entityType, flagDesc, deleted, createdBy,
'       createdOn, modifiedBy, modifiedOn, moduleID, calendarInfoID)
'                   Instantiates a populated FlagInfo object
' Reset()           Sets the object state to the original state when loaded from or
'                   last saved to the repository
' AgeThreshold      Indicates the number of minutes old data can be before it should be 
'                        refreshed from the DB.  Data should only be refreshed when Retrieved
'                        and when IsDirty is false
' IsAgedData        Will return true if the data has been held longer than the AgeThreshold
'
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class FlagInfo

#Region "Private member variables"
        Private nFlagID As Integer
        Private nEntityID As Integer
        Private nEntityType As Integer
        Private strFlagDescription As String
        Private bolDeleted As Boolean
        Private dtDueDate As Date
        Private strModuleID As String
        Private nCalendarInfoID As Integer
        Private strCreatedBy As String
        Private strSourceUser As String
        Private dtCreatedOn As DateTime
        Private strModifiedBy As String
        Private dtModifiedOn As DateTime
        Private strFlagColor As String
        Private dtTurnsRedOn As Date

        Private onFlagID As Integer
        Private onEntityID As Integer
        Private onEntityType As Integer
        Private ostrFlagDescription As String
        Private obolDeleted As Boolean
        Private odtDueDate As Date
        Private ostrModuleID As String
        Private onCalendarInfoID As Integer
        Private ostrCreatedBy As String
        Private ostrSourceUser As String
        Private odtCreatedOn As DateTime
        Private ostrModifiedBy As String
        Private odtModifiedOn As DateTime
        Private ostrFlagColor As String
        Private odtTurnsRedOn As Date

        Private bolIsDirty As Boolean
        Private dtDataAge As DateTime
        Private nAgeThreshold As Int16 = 5
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Sub New()
            MyBase.new()
            dtDataAge = Now()
            Me.Init()
        End Sub
        Sub New(ByVal flagID As Integer, _
                ByVal entityID As Integer, _
                ByVal entityType As Integer, _
                ByVal flagDesc As String, _
                ByVal deleted As Boolean, _
                ByVal dueDate As Date, _
                ByVal moduleID As String, _
                ByVal calendarInfoID As Integer, _
                ByVal CreatedBy As String, _
                ByVal CreatedOn As DateTime, _
                ByVal ModifiedBy As String, _
                ByVal ModifiedOn As DateTime, _
                ByVal turnsRedon As Date, _
                Optional ByVal SourceUserID As String = "", _
                Optional ByVal flagColor As String = "YELLOW")
            onFlagID = flagID
            onEntityID = entityID
            onEntityType = entityType
            ostrFlagDescription = flagDesc
            obolDeleted = deleted
            odtDueDate = dueDate
            ostrSourceUser = IIf(SourceUserID = String.Empty, ostrCreatedBy, SourceUserID)
            ostrCreatedBy = CreatedBy
            odtCreatedOn = CreatedOn
            ostrModifiedBy = ModifiedBy
            odtModifiedOn = ModifiedOn
            ostrModuleID = moduleID
            onCalendarInfoID = calendarInfoID
            ostrFlagColor = flagColor
            odtTurnsRedOn = turnsRedon
            dtDataAge = Now()
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            nFlagID = onFlagID
            nEntityID = onEntityID
            nEntityType = onEntityType
            strFlagDescription = ostrFlagDescription
            bolDeleted = obolDeleted
            dtDueDate = odtDueDate
            strModuleID = ostrModuleID
            nCalendarInfoID = onCalendarInfoID
            strCreatedBy = ostrCreatedBy
            strSourceUser = ostrSourceUser
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            strFlagColor = ostrFlagColor
            dtTurnsRedOn = odtTurnsRedOn
            bolIsDirty = False
        End Sub
        Public Sub Archive()
            onFlagID = nFlagID
            onEntityID = nEntityID
            onEntityType = nEntityType
            ostrFlagDescription = strFlagDescription
            obolDeleted = bolDeleted
            odtDueDate = dtDueDate
            ostrModuleID = strModuleID
            onCalendarInfoID = nCalendarInfoID
            ostrCreatedBy = strCreatedBy
            ostrSourceUser = strSourceUser
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn
            ostrFlagColor = strFlagColor
            odtTurnsRedOn = dtTurnsRedOn
            bolIsDirty = False
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            '(nFlagID <> onFlagID) Or _
            bolIsDirty = (nEntityID <> onEntityID) Or _
                        (nEntityType <> onEntityType) Or _
                        (strFlagDescription <> ostrFlagDescription) Or _
                        (bolDeleted <> obolDeleted) Or _
                        (dtDueDate <> odtDueDate) Or _
                        (strSourceUser <> ostrSourceUser) Or _
                        (strModuleID <> ostrModuleID) Or _
                        (nCalendarInfoID <> onCalendarInfoID) Or _
                        (strFlagColor <> ostrFlagColor) Or _
                        (dtTurnsRedOn <> dtTurnsRedOn)
        End Sub
        Private Sub Init()
            onFlagID = 0
            onEntityID = 0
            onEntityType = 0
            ostrFlagDescription = String.Empty
            obolDeleted = False
            odtDueDate = System.DateTime.Now
            ostrCreatedBy = String.Empty
            ostrSourceUser = String.Empty
            odtCreatedOn = System.DateTime.Now
            ostrModifiedBy = String.Empty
            odtModifiedOn = System.DateTime.Now
            ostrModuleID = String.Empty
            onCalendarInfoID = 0
            ostrFlagColor = "YELLOW"
            odtTurnsRedOn = DateAdd(DateInterval.Day, 120, Now.Date)
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Integer
            Get
                Return nFlagID
            End Get
            Set(ByVal Value As Integer)
                nFlagID = Value
                Me.CheckDirty()
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
        Public Property EntityType() As Integer
            Get
                Return nEntityType
            End Get
            Set(ByVal Value As Integer)
                nEntityType = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property FlagDescription() As String
            Get
                Return strFlagDescription
            End Get
            Set(ByVal Value As String)
                strFlagDescription = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return bolDeleted
            End Get
            Set(ByVal Value As Boolean)
                bolDeleted = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property DueDate() As Date
            Get
                Return dtDueDate
            End Get
            Set(ByVal Value As Date)
                dtDueDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ModuleID() As String
            Get
                Return strModuleID
            End Get
            Set(ByVal Value As String)
                strModuleID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CalendarInfoID() As Integer
            Get
                Return nCalendarInfoID
            End Get
            Set(ByVal Value As Integer)
                nCalendarInfoID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property SourceUserID() As String
            Get
                Return strSourceUser
            End Get
            Set(ByVal Value As String)
                strSourceUser = Value
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
        Public Property FlagColor() As String
            Get
                Return strFlagColor
            End Get
            Set(ByVal Value As String)
                ostrFlagColor = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property TurnsRedOn() As Date
            Get
                Return dtTurnsRedOn
            End Get
            Set(ByVal Value As Date)
                dtTurnsRedOn = Value
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

        Public ReadOnly Property IsAgedData() As Boolean
            Get
                Return IIf(DateDiff(DateInterval.Minute, dtDataAge, Now()) >= nAgeThreshold, True, False)
            End Get

        End Property

        'Public Property CreatedBy() As String
        '    Get
        '        Return strCreatedBy
        '    End Get
        '    Set(ByVal Value As String)
        '        strCreatedBy = Value
        '    End Set
        'End Property
        'Public Property DateCreated() As Date
        '    Get
        '        Return dtDateCreated
        '    End Get
        '    Set(ByVal Value As Date)
        '        dtDateCreated = Value
        '    End Set
        'End Property
        'Public Property LastEditedBy() As String
        '    Get
        '        Return strLastEditedBy
        '    End Get
        '    Set(ByVal Value As String)
        '        strLastEditedBy = Value
        '    End Set
        'End Property
        'Public Property DateLastEdited() As Date
        '    Get
        '        Return dtDateLastEdited
        '    End Get
        '    Set(ByVal Value As Date)
        '        dtDateLastEdited = Value
        '    End Set
        'End Property

#Region "IAccessors"
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
