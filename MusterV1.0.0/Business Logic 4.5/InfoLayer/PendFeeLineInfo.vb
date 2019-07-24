Namespace muster.info
    ' -------------------------------------------------------------------------------
    '    MUSTER.Info.PendFeeLineInfo
    '          Provides the container to persist MUSTER Pending Fee Line Items
    ' 
    '    Copyright (C) 2004, 2005 CIBER, Inc.
    '    All rights reserved.
    ' 
    '    Release   Initials    Date        Description
    '       1.0        AN      06/28/05    Original class definition.
    ' 
    '    Function          Description
    ' -------------------------------------------------------------------------------
    '
    Public Class PendFeeLineInfo
#Region "Private member variables"

        Private bolDeleted As Boolean
        Private bolIsDirty As Boolean = False
        Private nID As Integer
        Private dtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private dtDataAge As DateTime
        Private dtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private nAgeThreshold As Int16 = 5
        Private obolDeleted As Boolean
        Private odtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private odtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private ostrCreatedBy As String
        Private ostrModifiedBy As String
        Private strCreatedBy As String
        Private strModifiedBy As String

        Private nInvoiceAdviceID As Integer
        Private nItemSequenceNumber As Integer
        Private strInvoiceNumber As String
        Private nFacilityID As Integer
        Private nOwnerID As Integer
        Private strFiscalYear As String
        Private dtInvoiceDate As DateTime
        Private nQuantity As Integer
        Private dUnitPrice As Decimal
        Private nInvoiceLineAmount As Integer
        Private strFeeType As String
        Private nInvoiceType As Integer
        Private dtDueDate As DateTime
        Private strDescription As String

        Private onInvoiceAdviceID As Integer
        Private onItemSequenceNumber As Integer
        Private ostrInvoiceNumber As String
        Private onFacilityID As Integer
        Private onOwnerID As Integer
        Private ostrFiscalYear As Integer
        Private odtInvoiceDate As DateTime
        Private onQuantity As Integer
        Private odUnitPrice As Decimal
        Private onInvoiceLineAmount As Integer
        Private ostrFeeType As String
        Private onInvoiceType As Integer
        Private odtDueDate As DateTime
        Private ostrDescription As String

#End Region
#Region "Constructors"
        Public Sub New()
            MyBase.new()
            dtDataAge = Now()
            Me.Init()
        End Sub
        Public Sub New(ByVal ID As Integer, _
                        ByVal InvoiceAdviceID As Integer, _
                        ByVal ItemSequenceNumber As Integer, _
                        ByVal InvoiceNumber As String, _
                        ByVal FacilityID As Integer, _
                        ByVal OwnerID As Integer, _
                        ByVal FiscalYear As String, _
                        ByVal InvoiceDate As DateTime, _
                        ByVal Quantity As Integer, _
                        ByVal UnitPrice As Decimal, _
                        ByVal InvoiceLineAmount As Integer, _
                        ByVal FeeType As String, _
                        ByVal InvoiceType As Integer, _
                        ByVal DueDate As DateTime, _
                        ByVal DESCRIPTION As String)

            nInvoiceAdviceID = InvoiceAdviceID
            nItemSequenceNumber = ItemSequenceNumber
            strInvoiceNumber = InvoiceNumber
            nFacilityID = FacilityID
            nOwnerID = OwnerID
            strFiscalYear = FiscalYear
            dtInvoiceDate = InvoiceDate
            nQuantity = Quantity
            dUnitPrice = UnitPrice
            nInvoiceLineAmount = InvoiceLineAmount
            strFeeType = FeeType
            nInvoiceType = InvoiceType
            dtDueDate = DueDate
            strDescription = DESCRIPTION

            'odtGenerateDate = GENERATION_DATE
            nID = ID
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Archive()
            obolDeleted = bolDeleted
        End Sub
        Public Sub Reset()
            bolDeleted = obolDeleted
            bolIsDirty = False
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            bolIsDirty = (bolDeleted <> obolDeleted) Or _
                         (nInvoiceAdviceID = onInvoiceAdviceID) Or _
                         (nItemSequenceNumber = onItemSequenceNumber) Or _
                         (strInvoiceNumber = ostrInvoiceNumber) Or _
                         (nFacilityID = onFacilityID) Or _
                         (nOwnerID = onOwnerID) Or _
                         (strFiscalYear = ostrFiscalYear) Or _
                         (dtInvoiceDate = odtInvoiceDate) Or _
                         (nQuantity = onQuantity) Or _
                         (dUnitPrice = odUnitPrice) Or _
                         (nInvoiceLineAmount = onInvoiceLineAmount) Or _
                         (strFeeType = ostrFeeType) Or _
                         (nInvoiceType = onInvoiceType) Or _
                         (dtDueDate = odtDueDate) Or _
                         (strDescription = ostrDescription)
        End Sub
        Private Sub Init()
            nID = 0
            obolDeleted = False
            bolIsDirty = False

            nInvoiceAdviceID = 0
            nItemSequenceNumber = 0
            strInvoiceNumber = String.Empty
            nFacilityID = 0
            nOwnerID = 0
            strFiscalYear = 0
            dtInvoiceDate = Now()
            nQuantity = 0
            dUnitPrice = 0
            nInvoiceLineAmount = 0
            strFeeType = String.Empty
            nInvoiceType = 0
            dtDueDate = Now()
            strDescription = String.Empty
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Integer
            Get
                Return nID
            End Get
            Set(ByVal Value As Integer)
                nID = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property InvoiceAdviceId() As Integer
            Get
                Return nInvoiceAdviceID
            End Get
            Set(ByVal Value As Integer)
                nInvoiceAdviceID = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property ItemSequenceNumber() As Integer
            Get
                Return nItemSequenceNumber
            End Get
            Set(ByVal Value As Integer)
                Value = nItemSequenceNumber
                Me.CheckDirty()
            End Set
        End Property

        Public Property InvoiceNumber() As String
            Get
                Return strInvoiceNumber
            End Get
            Set(ByVal Value As String)
                strInvoiceNumber = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property FacilityId() As Integer
            Get
                Return nFacilityID
            End Get
            Set(ByVal Value As Integer)
                nFacilityID = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property OwnerId() As Integer
            Get
                Return nOwnerID
            End Get
            Set(ByVal Value As Integer)
                nOwnerID = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property FiscalYear() As String
            Get
                Return strFiscalYear
            End Get
            Set(ByVal Value As String)
                strFiscalYear = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property InvoiceDate() As DateTime
            Get
                Return dtInvoiceDate
            End Get
            Set(ByVal Value As DateTime)
                dtInvoiceDate = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property Quantity() As Integer
            Get
                Return nQuantity
            End Get
            Set(ByVal Value As Integer)
                nQuantity = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property UnitPrice() As Decimal
            Get
                Return dUnitPrice
            End Get
            Set(ByVal Value As Decimal)
                dUnitPrice = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property InvoiceLineAmount() As Integer
            Get
                Return nInvoiceLineAmount
            End Get
            Set(ByVal Value As Integer)
                nInvoiceLineAmount = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property FeeType() As String
            Get
                Return strFeeType
            End Get
            Set(ByVal Value As String)
                strFeeType = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property InvoiceType() As Integer
            Get
                Return nInvoiceType
            End Get
            Set(ByVal Value As Integer)
                nInvoiceType = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property DueDate() As DateTime
            Get
                Return dtDueDate
            End Get
            Set(ByVal Value As DateTime)
                dtDueDate = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property Description() As String
            Get
                Return strDescription
            End Get
            Set(ByVal Value As String)
                strDescription = Value
                Me.CheckDirty()
            End Set
        End Property



        Public Property AgeThreshold() As Integer
            Get
                Return nAgeThreshold
            End Get
            Set(ByVal Value As Integer)
                nAgeThreshold = Value
            End Set
        End Property
        ' The deleted flag for the LUST Activity
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
            Set(ByVal Value As Boolean)
                bolIsDirty = Value
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
        End Sub
#End Region

        Public Delegate Sub PendFeeLineInfoChangedEventHandler()


        ' Fired when CheckDirty determines that an attribute of the activity has been modified
        Public Event PendFeeLineInfoChanged As PendFeeLineInfoChangedEventHandler

    End Class
End Namespace
