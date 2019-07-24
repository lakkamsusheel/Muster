' -------------------------------------------------------------------------------
' MUSTER.Info.FinancialCommitmentCollection
' Provides the container to persist MUSTER FinancialInvoiceInfo state
' 
' Copyright (C) 2004, 2005 CIBER, Inc.
' All rights reserved.
' 
' Release   Initials    Date        Description
' 1.0        AB       06/24/05    Original class definition.
' 
' Function          Description
' ---

Namespace MUSTER.Info


    Public Class FinancialInvoiceInfo


        ' Delegate for event to indicate to parent that the info object has been modified in some manner
        Public Delegate Sub FinancialChangedEventHandler()
        ' Event that indicates to client that info object has changed in some manner
        ' 

#Region "Private Member Variables"
        Private nInvoiceID As Int64

        Private nReimbursementID As Integer
        Private nSequence As Integer
        Private strVendorInvoice As String
        Private nInvoicedAmount As Double
        Private nPaidAmount As Double
        Private strDeductionReason As String
        Private bolOnHold As Boolean
        Private bolFinal As Boolean
        Private strComment As String
        Private bolDeleted As Boolean
        Private strPONumber As String

        Private onReimbursementID As Integer
        Private onSequence As Integer
        Private ostrVendorInvoice As String
        Private onInvoicedAmount As Double
        Private onPaidAmount As Double
        Private ostrDeductionReason As String
        Private obolOnHold As Boolean
        Private obolFinal As Boolean
        Private ostrComment As String
        Private obolDeleted As Boolean
        Private ostrPONumber As String

        Private strCreatedBy As String = String.Empty
        Private dtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private strModifiedBy As String = String.Empty
        Private dtModifiedOn As Date = DateTime.Now.ToShortDateString

        Private ostrCreatedBy As String = String.Empty
        Private odtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private ostrModifiedBy As String = String.Empty
        Private odtModifiedOn As Date = DateTime.Now.ToShortDateString

        Private bolIsDirty As Boolean
        Private dtDataAge As DateTime
        Private nAgeThreshold As Int16 = 5
        Private nEntityID As Integer
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
                        ByVal ReimbursementID As Int64, _
                        ByVal Sequence As Int64, _
                        ByVal VendorInvoice As String, _
                        ByVal InvoicedAmount As Double, _
                        ByVal PaidAmount As Double, _
                        ByVal DeductionReason As String, _
                        ByVal OnHold As Boolean, _
                        ByVal Final As Boolean, _
                        ByVal Comment As String, _
                        ByVal CreatedBy As String, _
                        ByVal CreateDate As Date, _
                        ByVal LastEditedBy As String, _
                        ByVal LastEditDate As Date, _
                        ByVal bDeleted As Boolean, _
                        ByVal PONumber As String)


            nInvoiceID = Id

            onReimbursementID = ReimbursementID
            onSequence = Sequence
            ostrVendorInvoice = VendorInvoice
            onInvoicedAmount = InvoicedAmount
            onPaidAmount = PaidAmount
            ostrDeductionReason = DeductionReason
            obolOnHold = OnHold
            obolFinal = Final
            ostrComment = Comment
            ostrCreatedBy = CreatedBy
            odtCreatedOn = CreateDate
            ostrModifiedBy = LastEditedBy
            odtModifiedOn = LastEditDate
            obolDeleted = bDeleted
            ostrPONumber = PONumber
            dtDataAge = Now()
            Me.Reset()

        End Sub
#End Region

#Region "Exposed Methods"
        ' Add other attributes as necessitated by design
        Public Sub Archive()

            onReimbursementID = nReimbursementID
            onSequence = nSequence
            ostrVendorInvoice = strVendorInvoice
            onInvoicedAmount = nInvoicedAmount
            onPaidAmount = nPaidAmount
            ostrDeductionReason = strDeductionReason
            obolOnHold = bolOnHold
            obolFinal = bolFinal
            ostrComment = strComment

            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn

            obolDeleted = bolDeleted
            ostrPONumber = strPONumber

        End Sub

        Public Sub Reset()


            nReimbursementID = onReimbursementID
            nSequence = onSequence
            strVendorInvoice = ostrVendorInvoice
            nInvoicedAmount = onInvoicedAmount
            nPaidAmount = onPaidAmount
            strDeductionReason = ostrDeductionReason
            bolOnHold = obolOnHold
            bolFinal = obolFinal
            strComment = ostrComment
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn

            bolDeleted = obolDeleted
            strPONumber = ostrPONumber

        End Sub

#End Region

#Region "Private Methods"

        Private Sub CheckDirty()
            Dim bolOldState As Boolean = bolIsDirty

            bolIsDirty = (onReimbursementID <> nReimbursementID) Or _
                        (onSequence <> nSequence) Or _
                        (ostrVendorInvoice <> strVendorInvoice) Or _
                        (onPaidAmount <> nPaidAmount) Or _
                        (onInvoicedAmount <> nInvoicedAmount) Or _
                        (ostrDeductionReason <> strDeductionReason) Or _
                        (obolOnHold <> bolOnHold) Or _
                        (obolFinal <> bolFinal) Or _
                        (ostrComment <> strComment) Or _
                        (obolDeleted <> bolDeleted) Or _
                        (ostrPONumber <> strPONumber)

        End Sub

        Public Sub Init()

            nInvoiceID = 0

            onReimbursementID = 0
            onSequence = 1
            ostrVendorInvoice = ""
            onInvoicedAmount = 0
            onPaidAmount = 0
            ostrDeductionReason = ""
            obolOnHold = False
            obolFinal = False
            ostrComment = ""
            obolDeleted = False

            nReimbursementID = 0
            nSequence = 1
            strVendorInvoice = ""
            nInvoicedAmount = 0
            nPaidAmount = 0
            strDeductionReason = ""
            bolOnHold = False
            bolFinal = False
            strComment = ""
            bolDeleted = False
            strPONumber = ""
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
                Return nInvoiceID
            End Get
            Set(ByVal Value As Int64)
                nInvoiceID = Value
            End Set
        End Property

        Public Property ReimbursementID() As Int64
            Get
                Return nReimbursementID
            End Get
            Set(ByVal Value As Int64)
                nReimbursementID = Value
                CheckDirty()
            End Set
        End Property
        Public Property PaymentSequence() As Integer
            Get
                Return nSequence
            End Get
            Set(ByVal Value As Integer)
                nSequence = Value
                CheckDirty()
            End Set
        End Property

        Public Property VendorInvoice() As String
            Get
                Return strVendorInvoice
            End Get
            Set(ByVal Value As String)
                strVendorInvoice = Value
                CheckDirty()
            End Set
        End Property

        Public Property InvoicedAmount() As Double
            Get
                Return nInvoicedAmount
            End Get
            Set(ByVal Value As Double)
                nInvoicedAmount = Value
                CheckDirty()
            End Set
        End Property

        Public Property PaidAmount() As Double
            Get
                Return nPaidAmount
            End Get
            Set(ByVal Value As Double)
                nPaidAmount = Value
                CheckDirty()
            End Set
        End Property

        Public Property DeductionReason() As String
            Get
                Return strDeductionReason
            End Get
            Set(ByVal Value As String)
                strDeductionReason = Value
                CheckDirty()
            End Set
        End Property
        Public Property OnHold() As Boolean
            Get
                Return bolOnHold
            End Get
            Set(ByVal Value As Boolean)
                bolOnHold = Value
                CheckDirty()
            End Set
        End Property

        Public Property Final() As Boolean
            Get
                Return bolFinal
            End Get
            Set(ByVal Value As Boolean)
                bolFinal = Value
                CheckDirty()
            End Set
        End Property
        Public Property Comment() As String
            Get
                Return strComment
            End Get
            Set(ByVal Value As String)
                strComment = Value
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

        Public Property PONumber() As String
            Get
                Return strPONumber
            End Get
            Set(ByVal Value As String)
                strPONumber = Value
                CheckDirty()
            End Set
        End Property

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
