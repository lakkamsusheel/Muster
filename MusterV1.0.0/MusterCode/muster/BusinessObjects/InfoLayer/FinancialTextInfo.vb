
Namespace MUSTER.Info
    ' -------------------------------------------------------------------------------
    ' MUSTER.Info._ProtoInfo
    ' Provides the container to persist MUSTER _Proto state
    ' 
    ' Copyright (C) 2004, 2005 CIBER, Inc.
    ' All rights reserved.
    ' 
    ' Release   Initials    Date        Description
    ' 1.0        JVC       06/08/05    Original class definition.
    ' 
    ' Function          Description
    ' -------------------------------------------------------------------------------
    ' 
    Public Class FinancialTextInfo
        ' Delegate for event to indicate to parent that the info object has been modified in some manner
        Public Delegate Sub FinancialTextChangedEventHandler()
        ' Event that indicates to client that info object has changed in some manner
        ' 

#Region "Private Member Variables"
        Private bolActive As Boolean
        Private bolDeleted As Boolean
        Private bolIsDirty As Boolean
        Private dtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private dtDataAge As DateTime
        Private dtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private nAgeThreshold As Int16 = 5
        Private nEntityID As Integer
        Private nID As Int64
        Private nReasonType As Integer
        Private obolActive As Boolean
        Private obolDeleted As Boolean
        Private strCreatedBy As String
        Private strModifiedBy As String
        Private strReasonName As String
        Private strReasonText As String
        Private ostrReasonName As String
        Private ostrReasonText As String
        Private onReasonType As Integer

        Private ostrCreatedBy As String = String.Empty
        Private odtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private ostrModifiedBy As String = String.Empty
        Private odtModifiedOn As Date = DateTime.Now.ToShortDateString

        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region

#Region "Public Events"
        Public Event FinancialInfoChanged As FinancialTextChangedEventHandler
#End Region

#Region "Constructors"
        Public Sub New()
            MyBase.New()
            Me.Init()
            dtDataAge = Now()
        End Sub
        Public Sub New(ByVal Id As Long, _
                        ByVal ReasonType As Integer, _
                        ByVal ReasonName As String, _
                        ByVal ReasonText As String, _
                        ByVal CreatedBy As String, _
                        ByVal CreateDate As Date, _
                        ByVal LastEditedBy As String, _
                        ByVal LastEditDate As Date, _
                        ByVal bActive As Boolean, _
                        ByVal bDeleted As Boolean)


            nID = Id
            onReasonType = ReasonType
            ostrReasonName = ReasonName
            ostrReasonText = ReasonText
            ostrCreatedBy = CreatedBy
            odtCreatedOn = CreateDate
            ostrModifiedBy = LastEditedBy
            odtModifiedOn = LastEditDate
            obolActive = bActive
            obolDeleted = bDeleted

            dtDataAge = Now()
            Me.Reset()

        End Sub
#End Region

#Region "Exposed Methods"
        ' Add other attributes as necessitated by design
        Public Sub Archive()

            onReasonType = nReasonType
            ostrReasonName = strReasonName
            ostrReasonText = strReasonText
            obolActive = bolActive
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn
            obolDeleted = bolDeleted

        End Sub

        Public Sub Reset()

            nReasonType = onReasonType
            strReasonName = ostrReasonName
            strReasonText = ostrReasonText
            bolActive = obolActive

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

            bolIsDirty = (obolActive <> bolActive) Or _
                        (obolDeleted <> bolDeleted) Or _
                        (onReasonType <> nReasonType) Or _
                        (ostrReasonName <> strReasonName) Or _
                        (ostrReasonText <> strReasonText)

        End Sub

        Public Sub Init()
            Dim tmpDate As Date

            nID = 0
            onReasonType = 0
            ostrReasonName = String.Empty
            ostrReasonText = String.Empty
            strCreatedBy = String.Empty
            dtCreatedOn = tmpDate
            strModifiedBy = String.Empty
            dtModifiedOn = tmpDate
            obolActive = True
            obolDeleted = False
        End Sub
#End Region

#Region "Protected Methods"
        Protected Overrides Sub Finalize()
        End Sub
#End Region

#Region "Exposed Attributes"

        ' the Active/Inactive flag
        Public Property Active() As Boolean
            Get
                Return bolActive
            End Get
            Set(ByVal Value As Boolean)
                bolActive = Value
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
        Public ReadOnly Property CreatedOn() As Date
            Get
                Return dtCreatedOn
            End Get
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
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return dtModifiedOn
            End Get
        End Property
        ' The "common name" of the reason text.
        Public Property Reason_Name() As String
            Get
                Return strReasonName
            End Get
            Set(ByVal Value As String)
                strReasonName = Value
                CheckDirty()
            End Set
        End Property
        ' The actual text for the reason_type/ID combination - is stored as a TEXT field in the database.
        Public Property Reason_Text() As String
            Get
                Return strReasonText
            End Get
            Set(ByVal Value As String)
                strReasonText = Value
                CheckDirty()
            End Set
        End Property
        ' The property ID of the financial text "type" with which the text is associated (e.g. Conditions for Reimbursement, Additional Conditions for Reimbursement, Deduction Reasons, etc.)
        Public Property Reason_Type() As Integer
            Get
                Return nReasonType
            End Get
            Set(ByVal Value As Integer)
                nReasonType = Value
                CheckDirty()
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
