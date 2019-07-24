
Namespace MUSTER.Info
    ' -------------------------------------------------------------------------------
    '    MUSTER.Info.FeeInvoiceInfo
    '          Provides the container to persist MUSTER Fee Basis state
    ' 
    '    Copyright (C) 2004, 2005 CIBER, Inc.
    '    All rights reserved.
    ' 
    '    Release   Initials    Date        Description
    '       1.0        JVC       06/14/05    Original class definition.
    ' 
    '    Function          Description
    ' -------------------------------------------------------------------------------
    '
    Public Class FeeInvoiceInfo
#Region "Private member variables"

        Private WithEvents colInvoiceLineItems As New MUSTER.Info.FeeInvoiceCollection

        Private bolIsDirty As Boolean = False
        Private dtDataAge As DateTime

        Private nID As Integer
        Private nAgeThreshold As Integer = 5
        Private bolDeleted As Boolean
        Private obolDeleted As Boolean

        Private strRecType As String
        Private strFeeType As String
        Private strAdviceID As String
        Private nOwnerID As Int64
        Private sInvoiceAmount As Single
        Private sInvoiceLineAmount As Single
        Private strWarrantNumber As String
        Private nFiscalYear As Int16
        Private nFacilityID As Int64
        Private strDescription As String
        Private nSeqNumber As Int16
        Private sUnitPrice As Single
        Private strQuantity As String
        Private dtWarrantDate As Date
        Private dtDueDate As Date
        Private strCreditApplyTo As String
        Private bolTypeGeneration As Boolean
        Private bolProcessed As Boolean

        Private ostrRecType As String
        Private ostrFeeType As String
        Private ostrAdviceID As String
        Private onOwnerID As Int64
        Private osInvoiceAmount As Single
        Private osInvoiceLineAmount As Single
        Private ostrWarrantNumber As String
        Private onFiscalYear As Int16
        Private onFacilityID As Int64
        Private ostrDescription As String
        Private onSeqNumber As Int16
        Private osUnitPrice As Single
        Private ostrQuantity As String
        Private odtWarrantDate As Date
        Private odtDueDate As Date
        Private ostrCreditApplyTo As String
        Private obolTypeGeneration As Boolean
        Private obolProcessed As Boolean




        Private strInvoiceType As String
        Private strIssueName As String
        Private strIssueAddr1 As String
        Private strIssueAddr2 As String
        Private strIssueCity As String
        Private strIssueState As String
        Private strIssueZip As String
        Private strCheckTransID As String
        Private strCheckNumber As String


        Private ostrInvoiceType As String
        Private ostrIssueName As String
        Private ostrIssueAddr1 As String
        Private ostrIssueAddr2 As String
        Private ostrIssueCity As String
        Private ostrIssueState As String
        Private ostrIssueZip As String
        Private ostrCheckTransID As String
        Private ostrCheckNumber As String

        Private strCreatedBy As String = String.Empty
        Private dtCreateDate As Date = DateTime.Now.ToShortDateString
        Private strModifiedBy As String = String.Empty
        Private dtModifiedDate As Date = DateTime.Now.ToShortDateString

        Private ostrCreatedBy As String = String.Empty
        Private odtCreateDate As Date = DateTime.Now.ToShortDateString
        Private ostrModifiedBy As String = String.Empty
        Private odtModifiedDate As Date = DateTime.Now.ToShortDateString

#End Region
#Region "Constructors"
        Public Sub New()
            MyBase.new()
            dtDataAge = Now()
            Me.Init()
        End Sub
        Public Sub New(ByVal ID As Integer, _
                        ByVal CREATED_BY As String, _
                        ByVal CREATE_DATE As String, _
                        ByVal LAST_EDITED_BY As String, _
                        ByVal DATE_LAST_EDITED As Date, _
                        ByVal RecType As String, _
                        ByVal FeeType As String, _
                        ByVal AdviceID As String, _
                        ByVal OwnerID As Int64, _
                        ByVal InvoiceAmount As Single, _
                        ByVal InvoiceLineAmount As Single, _
                        ByVal WarrantNumber As String, _
                        ByVal FiscalYear As Int16, _
                        ByVal FacilityID As Int64, _
                        ByVal Description As String, _
                        ByVal SequenceNumber As Int16, _
                        ByVal UnitPrice As Single, _
                        ByVal Quantity As String, _
                        ByVal WarrantDate As Date, _
                        ByVal DueDate As Date, _
                        ByVal CreditApplyTo As String, _
                        ByVal TypeGeneration As Boolean, _
                        ByVal Processed As Boolean, _
                        ByVal InvoiceType As String, _
                        ByVal IssueName As String, _
                        ByVal IssueAddr1 As String, _
                        ByVal IssueAddr2 As String, _
                        ByVal IssueCity As String, _
                        ByVal IssueState As String, _
                        ByVal IssueZip As String, _
                        ByVal CheckTransID As String, _
                        ByVal CheckNumber As String, _
                        ByVal DELETED As Integer)


            nID = ID


            ostrInvoiceType = InvoiceType
            ostrIssueName = IssueName
            ostrIssueAddr1 = IssueAddr1
            ostrIssueAddr2 = IssueAddr2
            ostrIssueCity = IssueCity
            ostrIssueState = IssueState
            ostrIssueZip = IssueZip
            ostrCheckTransID = CheckTransID
            ostrCheckNumber = CheckNumber

            obolDeleted = DELETED
            ostrRecType = RecType
            ostrFeeType = FeeType
            ostrAdviceID = AdviceID
            onOwnerID = OwnerID
            osInvoiceAmount = InvoiceAmount
            osInvoiceLineAmount = InvoiceLineAmount
            ostrWarrantNumber = WarrantNumber
            onFiscalYear = FiscalYear
            onFacilityID = FacilityID
            ostrDescription = Description
            onSeqNumber = SequenceNumber
            osUnitPrice = UnitPrice
            ostrQuantity = Quantity
            odtWarrantDate = WarrantDate
            odtDueDate = DueDate
            ostrCreditApplyTo = CreditApplyTo
            obolTypeGeneration = TypeGeneration
            obolProcessed = Processed

            ostrCreatedBy = CREATED_BY
            ostrModifiedBy = LAST_EDITED_BY
            odtCreateDate = CREATE_DATE
            odtModifiedDate = DATE_LAST_EDITED

            Me.Reset()
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Archive()
            obolDeleted = bolDeleted
            ostrRecType = strRecType
            ostrFeeType = strFeeType
            ostrAdviceID = strAdviceID
            onOwnerID = nOwnerID
            osInvoiceAmount = sInvoiceAmount
            osInvoiceLineAmount = sInvoiceLineAmount
            ostrWarrantNumber = strWarrantNumber
            onFiscalYear = nFiscalYear
            onFacilityID = nFacilityID
            ostrDescription = strDescription
            onSeqNumber = nSeqNumber
            osUnitPrice = sUnitPrice
            ostrQuantity = strQuantity
            odtWarrantDate = dtWarrantDate
            odtDueDate = dtDueDate
            ostrCreditApplyTo = strCreditApplyTo
            obolTypeGeneration = bolTypeGeneration
            obolProcessed = bolProcessed

            ostrInvoiceType = strInvoiceType
            ostrIssueName = strIssueName
            ostrIssueAddr1 = strIssueAddr1
            ostrIssueAddr2 = strIssueAddr2
            ostrIssueCity = strIssueCity
            ostrIssueState = strIssueState
            ostrIssueZip = strIssueZip
            ostrCheckTransID = strCheckTransID
            ostrCheckNumber = strCheckNumber

            ostrCreatedBy = strCreatedBy
            ostrModifiedBy = strModifiedBy
            odtCreateDate = dtCreateDate
            odtModifiedDate = dtModifiedDate

        End Sub
        Public Sub Reset()
            bolDeleted = obolDeleted
            strRecType = ostrRecType
            strFeeType = ostrFeeType
            strAdviceID = ostrAdviceID
            nOwnerID = onOwnerID
            sInvoiceAmount = osInvoiceAmount
            sInvoiceLineAmount = osInvoiceLineAmount
            strWarrantNumber = ostrWarrantNumber
            nFiscalYear = onFiscalYear
            nFacilityID = onFacilityID
            strDescription = ostrDescription
            nSeqNumber = onSeqNumber
            sUnitPrice = osUnitPrice
            strQuantity = ostrQuantity
            dtWarrantDate = odtWarrantDate
            dtDueDate = odtDueDate
            strCreditApplyTo = ostrCreditApplyTo
            bolTypeGeneration = obolTypeGeneration
            bolProcessed = obolProcessed

            strInvoiceType = ostrInvoiceType
            strIssueName = ostrIssueName
            strIssueAddr1 = ostrIssueAddr1
            strIssueAddr2 = ostrIssueAddr2
            strIssueCity = ostrIssueCity
            strIssueState = ostrIssueState
            strIssueZip = ostrIssueZip
            strCheckTransID = ostrCheckTransID
            strCheckNumber = ostrCheckNumber

            strCreatedBy = ostrCreatedBy
            strModifiedBy = ostrModifiedBy
            dtCreateDate = odtCreateDate
            dtModifiedDate = odtModifiedDate

        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            bolIsDirty = (bolDeleted <> obolDeleted) Or _
                            (strRecType <> ostrRecType) Or _
                            (strFeeType <> ostrFeeType) Or _
                            (strAdviceID <> ostrAdviceID) Or _
                            (nOwnerID <> onOwnerID) Or _
                            (sInvoiceAmount <> osInvoiceAmount) Or _
                            (sInvoiceLineAmount <> osInvoiceLineAmount) Or _
                            (strWarrantNumber <> ostrWarrantNumber) Or _
                            (nFacilityID <> onFacilityID) Or _
                            (nSeqNumber <> onSeqNumber) Or _
                            (nFiscalYear <> onFiscalYear) Or _
                            (sUnitPrice <> osUnitPrice) Or _
                            (strQuantity <> ostrQuantity) Or _
                            (strDescription <> ostrDescription) Or _
                            (dtWarrantDate <> odtWarrantDate) Or _
                            (dtDueDate <> odtDueDate) Or _
                            (strCreditApplyTo <> ostrCreditApplyTo) Or _
                            (bolProcessed <> obolProcessed) Or _
                            (strInvoiceType = ostrInvoiceType) Or _
                            (strIssueName = ostrIssueName) Or _
                            (strIssueAddr1 = ostrIssueAddr1) Or _
                            (strIssueAddr2 = ostrIssueAddr2) Or _
                            (strIssueCity = ostrIssueCity) Or _
                            (strIssueState = ostrIssueState) Or _
                            (strIssueZip = ostrIssueZip) Or _
                            (strCheckTransID = ostrCheckTransID) Or _
                            (strCheckNumber = ostrCheckNumber) Or _
                            (bolTypeGeneration <> obolTypeGeneration)
        End Sub
        Private Sub Init()
            Dim tmpdate As Date

            nID = 0
            obolDeleted = False
            ostrRecType = ""
            ostrFeeType = ""
            ostrAdviceID = 0
            onOwnerID = 0
            osInvoiceAmount = 0
            osInvoiceLineAmount = 0
            ostrWarrantNumber = 0
            onFiscalYear = 0
            onFacilityID = 0
            ostrDescription = ""
            onSeqNumber = 0
            osUnitPrice = 0
            ostrQuantity = 0
            odtWarrantDate = tmpdate
            odtDueDate = tmpdate
            ostrCreditApplyTo = ""
            obolTypeGeneration = False
            obolProcessed = False
            ostrInvoiceType = ""
            ostrIssueName = ""
            ostrIssueAddr1 = ""
            ostrIssueAddr2 = ""
            ostrIssueCity = ""
            ostrIssueState = ""
            ostrIssueZip = ""
            ostrCheckTransID = ""
            ostrCheckNumber = ""
            ostrCreatedBy = String.Empty
            odtCreateDate = DateTime.Now.ToShortDateString
            ostrModifiedBy = String.Empty
            odtModifiedDate = DateTime.Now.ToShortDateString

        End Sub
#End Region
#Region "Exposed Attributes"

        Public Property CheckNumber() As String
            Get
                Return strCheckNumber
            End Get
            Set(ByVal Value As String)
                strCheckNumber = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CheckTransID() As String
            Get
                Return strCheckTransID
            End Get
            Set(ByVal Value As String)
                strCheckTransID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property IssueZip() As String
            Get
                Return strIssueZip
            End Get
            Set(ByVal Value As String)
                strIssueZip = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property IssueState() As String
            Get
                Return strIssueState
            End Get
            Set(ByVal Value As String)
                strIssueState = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property IssueCity() As String
            Get
                Return strIssueCity
            End Get
            Set(ByVal Value As String)
                strIssueCity = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property IssueAddr2() As String
            Get
                Return strIssueAddr2
            End Get
            Set(ByVal Value As String)
                strIssueAddr2 = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property IssueAddr1() As String
            Get
                Return strIssueAddr1
            End Get
            Set(ByVal Value As String)
                strIssueAddr1 = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property IssueName() As String
            Get
                Return strIssueName
            End Get
            Set(ByVal Value As String)
                strIssueName = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property InvoiceType() As String
            Get
                Return strInvoiceType
            End Get
            Set(ByVal Value As String)
                strInvoiceType = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ID() As Integer
            Get
                Return nID
            End Get
            Set(ByVal Value As Integer)
                nID = Value
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

        Public Property RecType() As String
            Get
                Return strRecType
            End Get
            Set(ByVal Value As String)
                strRecType = Value
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

        Public Property Description() As String
            Get
                Return strDescription
            End Get
            Set(ByVal Value As String)
                strDescription = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property FiscalYear() As Int16
            Get
                Return nFiscalYear
            End Get
            Set(ByVal Value As Int16)
                nFiscalYear = Value
                Me.CheckDirty()
            End Set
        End Property
        Public ReadOnly Property IsAgedData() As Boolean
            Get
                Return IIf(DateDiff(DateInterval.Minute, dtDataAge, Now()) >= nAgeThreshold, True, False)
            End Get
        End Property
        Public Property IsDirty() As Boolean
            Get
                Return bolIsDirty
            End Get
            Set(ByVal Value As Boolean)
                bolIsDirty = Value
            End Set
        End Property

        Public Property InvoiceAdviceID() As String
            Get
                Return strAdviceID
            End Get
            Set(ByVal Value As String)
                strAdviceID = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property OwnerID() As Int64
            Get
                Return nOwnerID
            End Get
            Set(ByVal Value As Int64)
                nOwnerID = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property InvoiceAmount() As Single
            Get
                Return sInvoiceAmount
            End Get
            Set(ByVal Value As Single)
                sInvoiceAmount = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property InvoiceLineAmount() As Single
            Get
                Return sInvoiceLineAmount
            End Get
            Set(ByVal Value As Single)
                sInvoiceLineAmount = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property WarrantNumber() As String
            Get
                Return strWarrantNumber
            End Get
            Set(ByVal Value As String)
                strWarrantNumber = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property FacilityID() As Int64
            Get
                Return nFacilityID
            End Get
            Set(ByVal Value As Int64)
                nFacilityID = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property SequenceNumber() As Int16
            Get
                Return nSeqNumber
            End Get
            Set(ByVal Value As Int16)
                nSeqNumber = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property UnitPrice() As Single
            Get
                Return sUnitPrice
            End Get
            Set(ByVal Value As Single)
                sUnitPrice = Value
                Me.CheckDirty()
            End Set
        End Property


        Public Property Quantity() As String
            Get
                Return strQuantity
            End Get
            Set(ByVal Value As String)
                strQuantity = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property WarrantDate() As Date
            Get
                Return dtWarrantDate
            End Get
            Set(ByVal Value As Date)
                dtWarrantDate = Value
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
        Public Property CreditApplyTo() As String
            Get
                Return strCreditApplyTo
            End Get
            Set(ByVal Value As String)
                strCreditApplyTo = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property TypeGeneration() As Boolean
            Get
                Return bolTypeGeneration
            End Get
            Set(ByVal Value As Boolean)
                bolTypeGeneration = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Processed() As Boolean
            Get
                Return bolProcessed
            End Get
            Set(ByVal Value As Boolean)
                bolProcessed = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property InvoiceLineItems() As MUSTER.Info.FeeInvoiceCollection
            Get
                Return colInvoiceLineItems
            End Get
            Set(ByVal Value As MUSTER.Info.FeeInvoiceCollection)
                colInvoiceLineItems = Value
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
                Return dtCreateDate
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
                Return dtModifiedDate
            End Get
        End Property

#End Region
#Region "Protected Operations"
        Protected Overrides Sub Finalize()
        End Sub
#End Region

        Public Delegate Sub FeeInvoiceInfoChangedEventHandler()
        ' Fired when CheckDirty determines that an attribute of the activity has been modified
        Public Event FeeInvoiceInfoChanged As FeeInvoiceInfoChangedEventHandler

    End Class
End Namespace


