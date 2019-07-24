Namespace MUSTER.Info
    <Serializable()> _
    Public Class ReportGroupRelationInfo
#Region "Public Events"
        Public Event GroupModuleRelationChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"
        Private nGroupID As Integer
        Private nReportID As Integer
        Private strGroupName As String
        Private bolActive As Boolean = False
        Private bolDeleted As Boolean

        Private onReportID As Integer
        Private obolDeleted As Boolean

        Private bolIsDirty As Boolean
        Private bolIsNew As Boolean
        Private dtDataAge As DateTime
        Private nAgeThreshold As Int16 = 5
        'Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Sub New()
            MyBase.new()
            Me.Init()
            dtDataAge = Now()
        End Sub
        Sub New(ByVal reportID As Integer, _
                ByVal groupID As Integer, _
                ByVal groupName As String, _
                ByVal active As Boolean, _
                ByVal deleted As Boolean)
            onReportID = reportID
            nGroupID = groupID
            strGroupName = groupName
            bolActive = active
            obolDeleted = deleted
            If onReportID > 0 Then
                bolIsNew = False
            Else
                bolIsNew = True
            End If
            dtDataAge = Now()
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            nReportID = onReportID
            bolDeleted = obolDeleted
            bolIsDirty = False
            RaiseEvent GroupModuleRelationChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            onReportID = nReportID
            obolDeleted = bolDeleted
            bolIsNew = False
            bolIsDirty = False
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            Dim bolOldValue As Boolean = bolIsDirty
            bolIsDirty = (nReportID <> onReportID) Or _
                        (bolDeleted <> obolDeleted)
            If bolOldValue <> bolIsDirty Then
                RaiseEvent GroupModuleRelationChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            nGroupID = 0
            onReportID = 0
            obolDeleted = False
            strGroupName = String.Empty
            bolActive = False
            bolIsNew = True
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As String
            Get
                Return nReportID.ToString + "|" + nGroupID.ToString
            End Get
            Set(ByVal Value As String)
                nReportID = Value.Split("|")(0)
                nGroupID = Value.Split("|")(1)
                Me.CheckDirty()
            End Set
        End Property
        Public Property ReportID() As Integer
            Get
                Return nReportID
            End Get
            Set(ByVal Value As Integer)
                nReportID = Value
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
        Public Property isNew() As Boolean
            Get
                Return bolIsNew
            End Get
            Set(ByVal Value As Boolean)
                bolIsNew = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property GroupName() As String
            Get
                Return strGroupName
            End Get
            Set(ByVal Value As String)
                strGroupName = Value
            End Set
        End Property
        Public Property Active() As Boolean
            Get
                Return bolActive
            End Get
            Set(ByVal Value As Boolean)
                bolActive = Value
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
#End Region
#Region "Protected Operations"
        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
#End Region
    End Class
End Namespace
