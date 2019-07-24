'-------------------------------------------------------------------------------
' MUSTER.Info.RegistrationInfo
'   Provides the container to persist MUSTER Template state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0                              Original class definition
'  1.1       JC         01/12/05    Added line of code to RESET to raise
'                                       data changed event when called.
'  1.2       AB         02/22/05    Added AgeThreshold and IsAgedData Attributes
'
' Function          Description
' New()             Instantiates an empty RegistrationInfo object
' New(Deleted, CreatedBy, CreatedOn, ModifiedBy, LastEdited, OwnerL2CSnippet)
'                   Instantiates a populated RegistrationInfo object
' New(dr)           Instantiates a populated RegistrationInfo object taking member state
'                   from the datarow provided
' Reset()           Sets the object state to the original state when loaded from or
'                   last saved to the repository
'
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class RegistrationInfo
        Implements iAccessors
#Region "Public Events"
        Public Event RegistrationInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"
        Private nRegID As Int64
        Private nOWNER_ID As Int64
        Private dtDATE_STARTED As DateTime
        Private dtDATE_COMPLETED As DateTime
        Private bolCOMPLETED As Boolean

        Private colActivity As MUSTER.Info.RegistrationActivityCollection
        Private ActivityInfo As MUSTER.Info.RegistrationActivityInfo
        '********************************************************
        '
        ' Other private member variables for current state here
        '
        '********************************************************
        Private bolDeleted As Boolean
        Private strCreatedBy As String
        Private dtCreatedOn As DateTime
        Private strModifiedBy As String
        Private dtModifiedOn As DateTime

        Private onRegID As Int64
        Private onOWNER_ID As Int64
        Private odtDATE_STARTED As DateTime
        Private odtDATE_COMPLETED As DateTime
        Private obolCOMPLETED As Boolean
        '********************************************************
        '
        ' Other private member variables for previous state here
        '
        '********************************************************
        Private obolDeleted As Boolean
        Private ostrCreatedBy As String
        Private odtCreatedOn As DateTime
        Private ostrModifiedBy As String
        Private odtModifiedOn As DateTime

        Private bolShowDeleted As Boolean = False
        Private dtDataAge As DateTime
        Private nAgeThreshold As Int16 = 5

        Private bolIsDirty As Boolean = False
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            dtDataAge = Now()
            Me.Init()
        End Sub
        '********************************************************
        '
        ' Update signature to reflect other attributes persisted
        '   by the object.
        '
        '********************************************************
        Sub New(ByVal ID As Int64, _
        ByVal OWNER_ID As Int64, _
        ByVal DATE_STARTED As DateTime, _
        ByVal DATE_COMPLETED As DateTime, _
        ByVal bolCOMPLETED As Boolean, _
        ByVal Deleted As Boolean, _
        ByVal CreatedBy As String, _
        ByVal CreatedOn As Date, _
        ByVal ModifiedBy As String, _
        ByVal LastEdited As Date)
            onRegID = ID
            onOWNER_ID = OWNER_ID
            odtDATE_STARTED = DATE_STARTED
            odtDATE_COMPLETED = DATE_COMPLETED
            obolCOMPLETED = bolCOMPLETED
            '********************************************************
            '
            ' Other private member variables for prior state here
            '
            '********************************************************
            obolDeleted = Deleted
            ostrCreatedBy = CreatedBy
            odtCreatedOn = CreatedOn
            ostrModifiedBy = ModifiedBy
            odtModifiedOn = LastEdited
            dtDataAge = Now()
            Me.Reset()
        End Sub
        Sub New(ByVal drTemplate As DataRow)
            Try
                onRegID = drTemplate.Item("ID")
                onOWNER_ID = drTemplate.Item("OWNER_ID")
                odtDATE_STARTED = drTemplate.Item("DATE_STARTED")
                odtDATE_COMPLETED = drTemplate.Item("DATE_COMPLETED")
                obolCOMPLETED = drTemplate.Item("COMPLETED")
                '********************************************************
                '
                ' Other private member variables for prior state here
                '
                '********************************************************
                obolDeleted = drTemplate.Item("DELETED")
                ostrCreatedBy = drTemplate.Item("CREATED_BY")
                odtCreatedOn = drTemplate.Item("DATE_CREATED")
                ostrModifiedBy = drTemplate.Item("LAST_EDITED_BY")
                odtModifiedOn = drTemplate.Item("DATE_LAST_EDITED")
                dtDataAge = Now()
                Me.Reset()
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            nRegID = onRegID
            nOWNER_ID = onOWNER_ID
            dtDATE_STARTED = odtDATE_STARTED
            dtDATE_COMPLETED = odtDATE_COMPLETED
            bolCOMPLETED = obolCOMPLETED
            '********************************************************
            '
            ' Other assignments of current state to prior state here
            '
            '********************************************************
            bolDeleted = obolDeleted
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            bolIsDirty = False
            RaiseEvent RegistrationInfoChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            onRegID = nRegID
            onOWNER_ID = nOWNER_ID
            odtDATE_STARTED = dtDATE_STARTED
            odtDATE_COMPLETED = dtDATE_COMPLETED
            obolCOMPLETED = bolCOMPLETED
            '********************************************************
            '
            ' Other assignments of prior state to current state here
            '
            '********************************************************
            obolDeleted = bolDeleted
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn
            bolIsDirty = False
        End Sub
#End Region

#Region "Private Operations"
        '********************************************************
        '
        ' Update eval in CheckDirty with other current and prior
        '    state variables.
        '
        '********************************************************
        Private Sub CheckDirty()
            Dim obolIsDirty As Boolean = bolIsDirty


            bolIsDirty = (nRegID <> onRegID) Or _
            (nOWNER_ID <> onOWNER_ID) Or _
            (dtDATE_STARTED <> odtDATE_STARTED) Or _
            (dtDATE_COMPLETED <> odtDATE_COMPLETED) Or _
            (bolCOMPLETED <> obolCOMPLETED) Or _
            (bolDeleted <> obolDeleted) Or _
            (strCreatedBy <> ostrCreatedBy) Or _
            (dtCreatedOn <> odtCreatedOn) Or _
            (strModifiedBy <> ostrModifiedBy) Or _
            (dtModifiedOn <> odtModifiedOn)


            If obolIsDirty <> bolIsDirty Then
                RaiseEvent RegistrationInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            onRegID = 0
            onOWNER_ID = 0
            odtDATE_STARTED = System.DateTime.Now
            odtDATE_COMPLETED = System.DateTime.Now
            bolCOMPLETED = False
            '********************************************************
            '
            ' Other assignments of current state to empty/false/etc here
            '
            '********************************************************
            bolDeleted = False
            strCreatedBy = String.Empty
            dtCreatedOn = System.DateTime.Now
            strModifiedBy = String.Empty
            dtModifiedOn = System.DateTime.Now
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        '********************************************************
        '
        ' Add properties to expose the persisted current state vars
        '
        '********************************************************
        Public Property ID() As Int64
            Get
                Return nRegID
            End Get
            Set(ByVal Value As Int64)
                nRegID = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property OWNER_ID() As Int64
            Get
                Return nOWNER_ID
            End Get
            Set(ByVal Value As Int64)
                nOWNER_ID = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property DATE_STARTED() As DateTime
            Get
                Return dtDATE_STARTED
            End Get
            Set(ByVal Value As DateTime)
                dtDATE_STARTED = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property DATE_COMPLETED() As DateTime
            Get
                Return dtDATE_COMPLETED
            End Get
            Set(ByVal Value As DateTime)
                dtDATE_COMPLETED = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property COMPLETED() As Boolean
            Get
                Return bolCOMPLETED
            End Get

            Set(ByVal value As Boolean)
                bolCOMPLETED = value
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
