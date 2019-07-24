Namespace MUSTER.Info
    Public Class FinancialInfo
        ' Delegate for event to indicate to parent that the info object has been modified in some manner
        Public Delegate Sub FinancialChangedEventHandler()
        ' Event that indicates to client that info object has changed in some manner
        ' 

#Region "Private Member Variables"
        Private bolDeleted As Boolean
        Private obolDeleted As Boolean
        Private nSequence As Integer
        Private onSequence As Integer
        Private nTecEvent As Integer
        Private onTecEvent As Integer
        Private dtStartDate As Date
        Private odtStartDate As Date
        Private dtClosedDate As Date
        Private odtClosedDate As Date
        Private nVendorID As Integer
        Private onVendorID As Integer
        Private nStatus As Integer
        Private onStatus As Integer
        Private nTecEventDesc As Integer

        Private bolIsDirty As Boolean
        Private dtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private dtDataAge As DateTime
        Private dtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private nAgeThreshold As Int16 = 5
        Private nEntityID As Integer
        Private nID As Int64

        Private strCreatedBy As String = String.Empty
        Private strModifiedBy As String = String.Empty

        Private ostrCreatedBy As String = String.Empty
        Private ostrModifiedBy As String = String.Empty
        Private odtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private odtCreatedOn As Date = DateTime.Now.ToShortDateString

        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region

#Region "Public Events"
        Public Event FinancialInfoChanged As FinancialChangedEventHandler
#End Region

#Region "Constructors"
        Public Sub New()
            MyBase.New()
            Me.Init()
            dtDataAge = Now()
        End Sub
        Public Sub New(ByVal Id As Long, _
                        ByVal Sequence As Integer, _
                        ByVal TecEvent As Int64, _
                        ByVal StartDate As Date, _
                        ByVal VendorID As Int64, _
                        ByVal Status As Int64, _
                        ByVal ClosedDate As Date, _
                        ByVal CreatedBy As String, _
                        ByVal CreateDate As Date, _
                        ByVal LastEditedBy As String, _
                        ByVal LastEditDate As Date, _
                        ByVal bDeleted As Boolean, _
                        ByVal TecEventDesc As Int64)


            nID = Id
            onSequence = Sequence
            onTecEvent = TecEvent
            odtStartDate = StartDate
            odtClosedDate = ClosedDate
            onVendorID = VendorID
            onStatus = Status
            ostrCreatedBy = CreatedBy
            odtCreatedOn = CreateDate
            ostrModifiedBy = LastEditedBy
            odtModifiedOn = LastEditDate
            obolDeleted = bDeleted
            nTecEventDesc = TecEventDesc



            dtDataAge = Now()
            Me.Reset()

        End Sub
#End Region

#Region "Exposed Methods"
        ' Add other attributes as necessitated by design
        Public Sub Archive()

            onSequence = nSequence
            onTecEvent = nTecEvent
            odtStartDate = dtStartDate
            odtClosedDate = dtClosedDate
            onVendorID = nVendorID
            onStatus = nStatus
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn

            obolDeleted = bolDeleted

        End Sub

        Public Sub Reset()


            nSequence = onSequence
            nTecEvent = onTecEvent
            dtStartDate = odtStartDate
            dtClosedDate = odtClosedDate
            nVendorID = onVendorID
            nStatus = onStatus
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            bolDeleted = obolDeleted

        End Sub

#End Region

#Region "Private Methods"

        Private Sub CheckDirty()
            Dim bolOldState As Boolean = bolIsDirty

            bolIsDirty = (onSequence <> nSequence) Or _
                        (onTecEvent <> nTecEvent) Or _
                        (odtStartDate <> dtStartDate) Or _
                        (odtClosedDate <> dtClosedDate) Or _
                        (onVendorID <> nVendorID) Or _
                        (onStatus <> nStatus) Or _
                        (obolDeleted <> bolDeleted)

        End Sub

        Public Sub Init()
            Dim tmpDate As Date

            nID = 0
            onSequence = 0
            onTecEvent = 0
            odtStartDate = "01/01/0001"
            odtClosedDate = "01/01/0001"
            onVendorID = 0
            onStatus = 0
            strCreatedBy = String.Empty
            dtCreatedOn = tmpDate
            strModifiedBy = String.Empty
            dtModifiedOn = tmpDate
            obolDeleted = False
        End Sub
#End Region

#Region "Protected Methods"
        Protected Overrides Sub Finalize()
        End Sub
#End Region

#Region "Exposed Attributes"

        Public Property Sequence() As Integer
            Get
                Return nSequence
            End Get
            Set(ByVal Value As Integer)
                nSequence = Value
            End Set
        End Property

        Public ReadOnly Property TecEventIDDesc() As Int64
            Get
                Return nTecEventDesc
            End Get
        End Property

        Public Property TecEventID() As Int64
            Get
                Return nTecEvent
            End Get
            Set(ByVal Value As Int64)
                nTecEvent = Value
            End Set
        End Property
        Public Property StartDate() As Date
            Get
                Return dtStartDate
            End Get
            Set(ByVal Value As Date)
                dtStartDate = Value
                CheckDirty()
            End Set
        End Property

        Public Property ClosedDate() As Date
            Get
                Return dtClosedDate
            End Get
            Set(ByVal Value As Date)
                dtClosedDate = Value
                CheckDirty()
            End Set
        End Property

        Public Property VendorID() As Int64
            Get
                Return nVendorID
            End Get
            Set(ByVal Value As Int64)
                nVendorID = Value
            End Set
        End Property
        Public Property Status() As Int64
            Get
                Return nStatus
            End Get
            Set(ByVal Value As Int64)
                nStatus = Value
                CheckDirty()
            End Set
        End Property

        ' The maximum age the info object can attain before requiring a refresh
        Public Property AgeThreshold() As Date
            Get
                Return dtDataAge
            End Get
            Set(ByVal Value As Date)
                dtDataAge = Value
            End Set
        End Property
        ' 
        ' The ID of the user that created the row
        Public Property CreatedBy() As String
            Get
                Return strCreatedBy
            End Get
            Set(ByVal Value As String)
                strCreatedBy = Value
            End Set
        End Property
        ' The date on which the row was created
        'Public ReadOnly Property CreatedOn() As Date
        '    Get
        '        Return dtCreatedOn
        '    End Get
        'End Property
        Public Property CreatedOn() As Date
            Get
                Return dtCreatedOn
            End Get
            Set(ByVal Value As Date)
                dtCreatedOn = Value
            End Set
        End Property
        ' The deleted flag for the TEC_ACT
        Public Property Deleted() As Boolean
            Get
                Return bolDeleted
            End Get
            Set(ByVal Value As Boolean)
                bolDeleted = Value
                CheckDirty()
            End Set
        End Property

        ' The entity ID associated.
        Public ReadOnly Property EntityID() As Integer
            Get
            End Get
        End Property
        ' the uniqueIdetifier for the _ProtoInfo
        Public Property ID() As Int64
            Get
                Return nID
            End Get
            Set(ByVal Value As Int64)
                nID = Value
            End Set
        End Property

        ' Raised when any of the _ProtoInfo attributes are modified
        Public Property IsDirty() As Boolean
            Get
                Return bolIsDirty
            End Get
            Set(ByVal Value As Boolean)
                bolIsDirty = False
            End Set
        End Property
        Public Property ModifiedBy() As String
            Get
                Return strModifiedBy
            End Get
            Set(ByVal Value As String)
                strModifiedBy = Value
            End Set
        End Property
        'Public ReadOnly Property ModifiedOn() As Date
        '    Get
        '        Return dtModifiedOn
        '    End Get
        'End Property
        Public Property ModifiedOn() As Date
            Get
                Return dtModifiedOn
            End Get
            Set(ByVal Value As Date)
                dtModifiedOn = Value
            End Set
        End Property
        ' Returns a boolean indicating if the data has aged beyond its preset limit
        Protected ReadOnly Property IsAgedData() As Boolean
            Get
                Return IIf(DateDiff(DateInterval.Minute, dtDataAge, Now()) >= nAgeThreshold, True, False)
            End Get
        End Property
#End Region

    End Class
End Namespace

