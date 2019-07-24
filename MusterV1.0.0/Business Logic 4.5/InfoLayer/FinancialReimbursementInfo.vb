' -------------------------------------------------------------------------------
' MUSTER.Info.FinancialReimbursementInfo
' Provides the container to persist MUSTER FinancialReimbursementInfo state
' 
' Copyright (C) 2004, 2005 CIBER, Inc.
' All rights reserved.
' 
' Release   Initials      Date        Description
' 1.0        AB           06/24/05    Original class definition.
' 2.0     Thomas Franey   02/25/09    Added Comment field to info
' 2.1     Thomas Franey   02/25/09    Noticed that PONumber was not in info, added it. 
' 
' Function          Description
' ---
Namespace MUSTER.Info


    Public Class FinancialReimbursementInfo
        ' Delegate for event to indicate to parent that the info object has been modified in some manner
        Public Delegate Sub FinancialReimbursementChangedEventHandler()
        ' Event that indicates to client that info object has changed in some manner
        ' 

#Region "Public Events"
        Public Event FinancialReimbursementInfoChanged As FinancialReimbursementChangedEventHandler
#End Region

#Region "Private Member Variables"

        Private nFinancialEventID As Int64
        Private nCommitmentID As Int64
        Private nPaymentNumber As Integer
        Private dtReceivedDate As Date
        Private dtPaymentDate As Date
        Private nRequestedAmount As Double
        Private strIncompleteReason As String
        Private bolIncomplete As Boolean
        Private strIncompleteOther As String
        Private bolDeleted As Boolean
        Private strPONumber As String
        Private strComment As String

        Private onFinancialEventID As Int64
        Private onCommitmentID As Int64
        Private onPaymentNumber As Integer
        Private odtReceivedDate As Date
        Private odtPaymentDate As Date
        Private onRequestedAmount As Double
        Private ostrIncompleteReason As String
        Private obolIncomplete As Boolean
        Private ostrIncompleteOther As String
        Private ostrComment As String
        Private oStrPONumber As String
        Private obolDeleted As Boolean


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
        Private odtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private ostrModifiedBy As String = String.Empty
        Private odtModifiedOn As Date = DateTime.Now.ToShortDateString

        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region

#Region "Constructors"
        Public Sub New()
            MyBase.New()
            Me.Init()
            dtDataAge = Now()
        End Sub
        Public Sub New(ByVal Id As Long, _
                        ByVal FinancialEventID As Int64, _
                        ByVal CommitmentID As Int64, _
                        ByVal PaymentNumber As Int64, _
                        ByVal ReceivedDate As Date, _
                        ByVal PaymentDate As Date, _
                        ByVal RequestedAmount As Double, _
                        ByVal IncompleteReason As String, _
                        ByVal Incomplete As Boolean, _
                        ByVal IncompleteOther As String, _
                        ByVal CreatedBy As String, _
                        ByVal CreateDate As Date, _
                        ByVal LastEditedBy As String, _
                        ByVal LastEditDate As Date, _
                        ByVal bDeleted As Boolean, _
                        Optional ByVal PONumber As String = "", _
                        Optional ByVal Comment As String = "")


            nID = Id
            onFinancialEventID = FinancialEventID
            onCommitmentID = CommitmentID
            onPaymentNumber = PaymentNumber
            odtReceivedDate = ReceivedDate
            odtPaymentDate = PaymentDate
            onRequestedAmount = RequestedAmount
            ostrIncompleteReason = IncompleteReason
            obolIncomplete = Incomplete
            ostrIncompleteOther = IncompleteOther
            ostrCreatedBy = CreatedBy
            odtCreatedOn = CreateDate
            ostrModifiedBy = LastEditedBy
            odtModifiedOn = LastEditDate
            ostrComment = Comment
            obolDeleted = bDeleted
            oStrPONumber = PONumber

            dtDataAge = Now()
            Me.Reset()

        End Sub
#End Region

#Region "Exposed Methods"
        ' Add other attributes as necessitated by design
        Public Sub Archive()

            onFinancialEventID = nFinancialEventID
            onCommitmentID = nCommitmentID
            onPaymentNumber = nPaymentNumber
            odtReceivedDate = dtReceivedDate
            odtPaymentDate = dtPaymentDate
            onRequestedAmount = nRequestedAmount
            ostrIncompleteReason = strIncompleteReason
            obolIncomplete = bolIncomplete
            ostrIncompleteOther = strIncompleteOther
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn
            ostrComment = strComment
            obolDeleted = bolDeleted
            oStrPONumber = strPONumber


        End Sub

        Public Sub Reset()

            nFinancialEventID = onFinancialEventID
            nCommitmentID = onCommitmentID
            nPaymentNumber = onPaymentNumber
            dtReceivedDate = odtReceivedDate
            dtPaymentDate = odtPaymentDate
            nRequestedAmount = onRequestedAmount
            strIncompleteReason = ostrIncompleteReason
            bolIncomplete = obolIncomplete
            strIncompleteOther = ostrIncompleteOther
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            strComment = ostrComment
            bolDeleted = obolDeleted
            strPONumber = oStrPONumber

        End Sub

#End Region

#Region "Private Methods"

        Private Sub CheckDirty()
            Dim bolOldState As Boolean = bolIsDirty

            bolIsDirty = (onCommitmentID <> nCommitmentID) Or _
                        (onFinancialEventID <> nFinancialEventID) Or _
                        (onPaymentNumber <> nPaymentNumber) Or _
                        (odtReceivedDate <> dtReceivedDate) Or _
                        (dtPaymentDate <> odtPaymentDate) Or _
                        (onRequestedAmount <> nRequestedAmount) Or _
                        (ostrIncompleteReason <> strIncompleteReason) Or _
                        (obolIncomplete <> bolIncomplete) Or _
                        (ostrIncompleteOther <> strIncompleteOther) Or _
                        (obolDeleted <> bolDeleted) Or _
                        (strPONumber <> oStrPONumber) Or _
                        (ostrComment <> strComment)


        End Sub

        Public Sub Init()
            Dim tmpDate As Date

            nID = 0
            nCommitmentID = 0
            nFinancialEventID = 0
            nPaymentNumber = 0
            dtReceivedDate = tmpDate
            dtPaymentDate = tmpDate
            nRequestedAmount = 0
            strIncompleteReason = String.Empty
            bolIncomplete = False
            strIncompleteOther = String.Empty
            strCreatedBy = String.Empty
            dtCreatedOn = tmpDate
            strModifiedBy = String.Empty
            dtModifiedOn = tmpDate
            ostrComment = String.Empty
            obolDeleted = False
            oStrPONumber = String.Empty

        End Sub
#End Region

#Region "Protected Methods"
        Protected Overrides Sub Finalize()
        End Sub
#End Region

#Region "Exposed Attributes"
        ' the uniqueIdetifier for the _ProtoInfo
        Public Property ID() As Int64
            Get
                Return nID
            End Get
            Set(ByVal Value As Int64)
                nID = Value
            End Set
        End Property
        ' = 0

        'nFinancialEventID
        Public Property FinancialEventID() As Int64
            Get
                Return nFinancialEventID
            End Get
            Set(ByVal Value As Int64)
                nFinancialEventID = Value
            End Set
        End Property




        Public Property CommitmentID() As Int64
            Get
                Return nCommitmentID
            End Get
            Set(ByVal Value As Int64)
                nCommitmentID = Value
            End Set
        End Property

        Public Property PaymentNumber() As Int64
            Get
                Return nPaymentNumber
            End Get
            Set(ByVal Value As Int64)
                nPaymentNumber = Value
            End Set
        End Property
        Public Property ReceivedDate() As Date
            Get
                Return dtReceivedDate
            End Get
            Set(ByVal Value As Date)
                dtReceivedDate = Value
            End Set
        End Property
        Public Property PaymentDate() As Date
            Get
                Return dtPaymentDate
            End Get
            Set(ByVal Value As Date)
                dtPaymentDate = Value
            End Set
        End Property
        Public Property RequestedAmount() As Double
            Get
                Return nRequestedAmount
            End Get
            Set(ByVal Value As Double)
                nRequestedAmount = Value
            End Set
        End Property
        Public Property IncompleteReason() As String
            Get
                Return strIncompleteReason
            End Get
            Set(ByVal Value As String)
                strIncompleteReason = Value
            End Set
        End Property
        Public Property Incomplete() As Boolean
            Get
                Return bolIncomplete
            End Get
            Set(ByVal Value As Boolean)
                bolIncomplete = Value
            End Set
        End Property

        Public Property IncompleteOther() As String
            Get
                Return strIncompleteOther
            End Get
            Set(ByVal Value As String)
                strIncompleteOther = Value
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


        ' The reimburse ERAC flag


        ' The entity ID associated.
        Public ReadOnly Property EntityID() As Integer
            Get
            End Get
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

        Public Property Comment() As String
            Get
                Return strComment
            End Get
            Set(ByVal Value As String)
                strComment = Value
            End Set
        End Property

        Public Property PONumber() As String
            Get
                Return strPONumber
            End Get
            Set(ByVal Value As String)
                strPONumber = Value
            End Set
        End Property
        Public ReadOnly Property ModifiedOn() As Date
            Get
                Return dtModifiedOn
            End Get
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