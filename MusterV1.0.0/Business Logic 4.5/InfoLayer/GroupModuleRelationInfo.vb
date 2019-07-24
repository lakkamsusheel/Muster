Namespace MUSTER.Info
    <Serializable()> _
    Public Class GroupModuleRelationInfo
#Region "Public Events"
        Public Event GroupModuleRelationChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"
        Private nGroupID As Integer
        Private nModuleID As Integer
        Private strModule As String
        Private bolWriteAccess As Boolean
        Private bolReadAccess As Boolean
        Private bolDeleted As Boolean

        Private onGroupID As Integer
        Private onModuleID As Integer
        Private ostrModule As String
        Private obolWriteAccess As Boolean
        Private obolReadAccess As Boolean
        Private obolDeleted As Boolean

        Private bolIsDirty As Boolean
        Private dtDataAge As DateTime
        Private nAgeThreshold As Int16 = 5
        'Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Sub New()
            MyBase.new()
            dtDataAge = Now()
            Me.Init()
        End Sub
        Sub New(ByVal groupID As Integer, _
                ByVal moduleID As Integer, _
                ByVal writeAccess As Boolean, _
                ByVal readAccess As Boolean, _
                ByVal deleted As Boolean, _
                ByVal moduleName As String)
            onGroupID = groupID
            onModuleID = moduleID
            obolWriteAccess = writeAccess
            obolReadAccess = readAccess
            obolDeleted = deleted
            ostrModule = moduleName
            dtDataAge = Now()
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            nGroupID = onGroupID
            nModuleID = onModuleID
            bolWriteAccess = obolWriteAccess
            bolReadAccess = obolReadAccess
            bolDeleted = obolDeleted
            strModule = ostrModule
            bolIsDirty = False
            RaiseEvent GroupModuleRelationChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            onGroupID = nGroupID
            onModuleID = nModuleID
            obolWriteAccess = bolWriteAccess
            obolReadAccess = bolReadAccess
            obolDeleted = bolDeleted
            ostrModule = strModule
            bolIsDirty = False
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            Dim bolOldValue As Boolean = bolIsDirty
            bolIsDirty = (nGroupID <> onGroupID) Or _
                        (nModuleID <> onModuleID) Or _
                        (bolWriteAccess <> obolWriteAccess) Or _
                        (bolReadAccess <> obolReadAccess) Or _
                        (bolDeleted <> obolDeleted)
            If bolOldValue <> bolIsDirty Then
                RaiseEvent GroupModuleRelationChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            onGroupID = 0
            onModuleID = 0
            obolWriteAccess = False
            obolReadAccess = False
            obolDeleted = False
            ostrModule = String.Empty
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As String
            Get
                Return nGroupID.ToString + "|" + nModuleID.ToString
            End Get
            Set(ByVal Value As String)
                nGroupID = Value.Split("|")(0)
                nModuleID = Value.Split("|")(1)
                Me.CheckDirty()
            End Set
        End Property
        Public Property GroupID() As Integer
            Get
                Return nGroupID
            End Get
            Set(ByVal Value As Integer)
                nGroupID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ModuleID() As Integer
            Get
                Return nModuleID
            End Get
            Set(ByVal Value As Integer)
                nModuleID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property WriteAccess() As Boolean
            Get
                Return bolWriteAccess
            End Get
            Set(ByVal Value As Boolean)
                bolWriteAccess = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ReadAccess() As Boolean
            Get
                Return bolReadAccess
            End Get
            Set(ByVal Value As Boolean)
                bolReadAccess = Value
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
        Public Property IsDirty() As Boolean
            Get
                Return bolIsDirty
            End Get

            Set(ByVal value As Boolean)
                bolIsDirty = value
            End Set
        End Property
        Public Property ModuleName() As String
            Get
                Return strModule
            End Get
            Set(ByVal Value As String)
                strModule = Value
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
#Region "Protected Operations"
        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
#End Region
    End Class
End Namespace
