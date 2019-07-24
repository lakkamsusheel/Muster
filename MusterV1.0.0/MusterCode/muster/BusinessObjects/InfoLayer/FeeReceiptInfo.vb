
Namespace MUSTER.Info
    ' -------------------------------------------------------------------------------
    '    MUSTER.Info.FeeReceiptInfo
    '          Provides the container to persist MUSTER Fee Receipt state
    ' 
    '    Copyright (C) 2004, 2005 CIBER, Inc.
    '    All rights reserved.
    ' 
    '    Release   Initials    Date        Description
    '       1.0        AB       09/26/05    Original class definition.
    ' 
    '    Function          Description
    ' -------------------------------------------------------------------------------
    '
    Public Class FeeReceiptInfo
#Region "Private member variables"


        Private bolIsDirty As Boolean = False
        Private dtDataAge As DateTime

        Private nID As Integer
        Private nAgeThreshold As Integer = 5
        Private bolDeleted As Boolean
        Private obolDeleted As Boolean
        Private strFiscalYear As String
        Private strReturnType As String
        Private strCheckTransID As String
        Private nOwnerID As Int64
        Private nFacilityID As Int64
        Private strInvoiceID As String
        Private strCheckNumber As String
        Private strMisapplyFlag As String
        Private strMisapplyReason As String
        Private strOverpaymentReason As String
        Private strIssuingCompany As String
        Private nSeqNumber As Int16
        Private sAmountReceived As Single
        Private dtReceiptDate As Date

        '-------------------

        Private ostrFiscalYear As String
        Private ostrReturnType As String
        Private ostrCheckTransID As String
        Private onOwnerID As Int64
        Private onFacilityID As Int64
        Private ostrInvoiceID As String
        Private ostrCheckNumber As String
        Private ostrMisapplyFlag As String
        Private ostrMisapplyReason As String
        Private ostrOverpaymentReason As String
        Private ostrIssuingCompany As String
        Private onSeqNumber As Int16
        Private osAmountReceived As Single
        Private odtReceiptDate As Date

        Private strCreatedBy As String = String.Empty
        Private strModifiedBy As String = String.Empty
        Private dtCreateDate As Date = DateTime.Now.ToShortDateString
        Private dtModifiedDate As Date = DateTime.Now.ToShortDateString

        Private ostrCreatedBy As String = String.Empty
        Private ostrModifiedBy As String = String.Empty
        Private odtCreateDate As Date = DateTime.Now.ToShortDateString
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
                        ByVal FiscalYear As String, _
                        ByVal ReturnType As String, _
                        ByVal CheckTransID As String, _
                        ByVal OwnerID As Int64, _
                        ByVal FacilityID As Int64, _
                        ByVal InvoiceID As String, _
                        ByVal CheckNumber As String, _
                        ByVal MisapplyFlag As String, _
                        ByVal MisapplyReason As String, _
                        ByVal IssuingCompany As String, _
                        ByVal SequenceNumber As Int16, _
                        ByVal AmountReceived As Single, _
                        ByVal ReceiptDate As Date, _
                        ByVal OverpaymentReason As String, _
                        ByVal DELETED As Integer)


            nID = ID

            obolDeleted = DELETED
            ostrFiscalYear = FiscalYear
            ostrReturnType = ReturnType
            ostrCheckTransID = CheckTransID
            onOwnerID = OwnerID
            onFacilityID = FacilityID
            ostrInvoiceID = InvoiceID
            ostrCheckNumber = CheckNumber
            ostrMisapplyFlag = MisapplyFlag
            ostrMisapplyReason = MisapplyReason
            ostrOverpaymentReason = OverpaymentReason
            ostrIssuingCompany = IssuingCompany
            onSeqNumber = SequenceNumber
            osAmountReceived = AmountReceived
            odtReceiptDate = ReceiptDate

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

            ostrFiscalYear = strFiscalYear
            ostrReturnType = strReturnType
            ostrCheckTransID = strCheckTransID
            onOwnerID = nOwnerID
            onFacilityID = nFacilityID
            ostrInvoiceID = strInvoiceID
            ostrCheckNumber = strCheckNumber
            ostrMisapplyFlag = strMisapplyFlag
            ostrMisapplyReason = strMisapplyReason
            ostrOverpaymentReason = strOverpaymentReason
            ostrIssuingCompany = strIssuingCompany
            onSeqNumber = nSeqNumber
            osAmountReceived = sAmountReceived
            odtReceiptDate = dtReceiptDate

            ostrCreatedBy = strCreatedBy
            ostrModifiedBy = strModifiedBy
            odtCreateDate = dtCreateDate
            odtModifiedDate = dtModifiedDate

        End Sub
        Public Sub Reset()
            bolDeleted = obolDeleted
            strFiscalYear = ostrFiscalYear
            strReturnType = ostrReturnType
            strCheckTransID = ostrCheckTransID
            nOwnerID = onOwnerID
            nFacilityID = onFacilityID
            strInvoiceID = ostrInvoiceID
            strCheckNumber = ostrCheckNumber
            strMisapplyFlag = ostrMisapplyFlag
            strMisapplyReason = ostrMisapplyReason
            strOverpaymentReason = ostrOverpaymentReason
            strIssuingCompany = ostrIssuingCompany
            nSeqNumber = onSeqNumber
            sAmountReceived = osAmountReceived
            dtReceiptDate = odtReceiptDate

            strCreatedBy = ostrCreatedBy
            strModifiedBy = ostrModifiedBy
            dtCreateDate = odtCreateDate
            dtModifiedDate = odtModifiedDate

        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            bolIsDirty = (bolDeleted <> obolDeleted) Or _
                            (strFiscalYear <> ostrFiscalYear) Or _
                            (strReturnType <> ostrReturnType) Or _
                            (strInvoiceID <> ostrInvoiceID) Or _
                            (nOwnerID <> onOwnerID) Or _
                            (strCheckTransID <> ostrCheckTransID) Or _
                            (nFacilityID <> onFacilityID) Or _
                            (nSeqNumber <> onSeqNumber) Or _
                            (strIssuingCompany <> ostrIssuingCompany) Or _
                            (strMisapplyReason <> ostrMisapplyReason) Or _
                            (strOverpaymentReason <> ostrOverpaymentReason) Or _
                            (strCheckNumber <> ostrCheckNumber) Or _
                            (dtReceiptDate <> odtReceiptDate) Or _
                            (sAmountReceived <> osAmountReceived) Or _
                            (strMisapplyFlag <> ostrMisapplyFlag)

        End Sub
        Private Sub Init()
            Dim tmpdate As Date

            nID = 0
            obolDeleted = False
            ostrFiscalYear = DatePart(DateInterval.Year, Now.Date)
            ostrReturnType = String.Empty
            ostrCheckTransID = String.Empty
            onOwnerID = 0
            onFacilityID = 0
            ostrInvoiceID = 0
            ostrCheckNumber = ""
            ostrMisapplyFlag = ""
            ostrMisapplyReason = ""
            ostrOverpaymentReason = ""
            ostrIssuingCompany = ""
            onSeqNumber = 0
            osAmountReceived = 0
            odtReceiptDate = "01/01/0001"

            ostrCreatedBy = String.Empty
            ostrModifiedBy = String.Empty
            odtCreateDate = DateTime.Now.ToShortDateString
            odtModifiedDate = DateTime.Now.ToShortDateString

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
        Public Property AgeThreshold() As Integer
            Get
                Return nAgeThreshold
            End Get
            Set(ByVal Value As Integer)
                nAgeThreshold = Value
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
        Public Property Deleted() As Boolean
            Get
                Return bolDeleted
            End Get
            Set(ByVal Value As Boolean)
                bolDeleted = Value
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
        Public Property ReturnType() As String
            Get
                Return strReturnType
            End Get
            Set(ByVal Value As String)
                strReturnType = Value
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
        Public Property OwnerID() As Int64
            Get
                Return nOwnerID
            End Get
            Set(ByVal Value As Int64)
                nOwnerID = Value
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
        Public Property InvoiceNumber() As String
            Get
                Return strInvoiceID
            End Get
            Set(ByVal Value As String)
                strInvoiceID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CheckNumber() As String
            Get
                Return strCheckNumber
            End Get
            Set(ByVal Value As String)
                strCheckNumber = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property MisapplyFlag() As String
            Get
                Return strMisapplyFlag
            End Get
            Set(ByVal Value As String)
                strMisapplyFlag = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property MisapplyReason() As String
            Get
                Return strMisapplyReason
            End Get
            Set(ByVal Value As String)
                strMisapplyReason = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property OverpaymentReason() As String
            Get
                Return strOverpaymentReason
            End Get
            Set(ByVal Value As String)
                strOverpaymentReason = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property IssuingCompany() As String
            Get
                Return strIssuingCompany
            End Get
            Set(ByVal Value As String)
                strIssuingCompany = Value
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
        Public Property AmountReceived() As Single
            Get
                Return sAmountReceived
            End Get
            Set(ByVal Value As Single)
                sAmountReceived = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ReceiptDate() As Date
            Get
                Return dtReceiptDate
            End Get
            Set(ByVal Value As Date)
                dtReceiptDate = Value
                Me.CheckDirty()
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

        Public Delegate Sub FeeReceiptInfoChangedEventHandler()
        ' Fired when CheckDirty determines that an attribute of the activity has been modified
        Public Event FeeReceiptInfoChanged As FeeReceiptInfoChangedEventHandler

    End Class
End Namespace