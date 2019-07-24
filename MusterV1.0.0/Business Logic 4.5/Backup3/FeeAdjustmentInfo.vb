
Namespace MUSTER.Info
    ' -------------------------------------------------------------------------------
    '    MUSTER.Info.FeeAdjustmentInfo
    '          Provides the container to persist MUSTER Fee Adjustment state
    ' 
    '    Copyright (C) 2004, 2005 CIBER, Inc.
    '    All rights reserved.
    ' 
    '    Release   Initials    Date        Description
    '       1.0      AB       01/09/06    Original class definition.
    ' 
    '    Function          Description
    ' -------------------------------------------------------------------------------
    '
    Public Class FeeAdjustmentInfo
#Region "Private member variables"

        Private bolDeleted As Boolean
        Private obolDeleted As Boolean
        Private bolIsDirty As Boolean = False
        Private nID As Integer

        Private bolReturnedFromBP2K As Boolean

        Private nOwnerID As Int64
        Private strCreditCode As String
        Private nFiscalYear As Integer
        Private nFacilityID As Int64
        Private strInvoiceNumber As String
        Private sAmount As Decimal
        Private dtApplied As Date
        Private strCheckNumber As String
        Private strReason As String
        Private nItemSeqNumber As Int64

        Private onOwnerID As Int64
        Private ostrCreditCode As String
        Private onFiscalYear As Integer
        Private onFacilityID As Int64
        Private ostrInvoiceNumber As String
        Private osAmount As Decimal
        Private odtApplied As Date
        Private ostrCheckNumber As String
        Private ostrReason As String
        Private onItemSeqNumber As Int64
        Private dtDataAge As DateTime
        Private nAgeThreshold As Int16 = 5

        Private dtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private dtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private strCreatedBy As String = String.Empty
        Private strModifiedBy As String = String.Empty

        Private odtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private odtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private ostrCreatedBy As String = String.Empty
        Private ostrModifiedBy As String = String.Empty

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
                        ByVal DELETED As Integer, _
                        ByVal OwnerID As Int64, _
                        ByVal CreditCode As String, _
                        ByVal FiscalYear As Long, _
                        ByVal FacilityID As Int64, _
                        ByVal InvoiceNumber As String, _
                        ByVal ItemSeqNumber As Int64, _
                        ByVal Amount As Decimal, _
                        ByVal Applied As Date, _
                        ByVal CheckNumber As String, _
                        ByVal Reason As String, _
                        ByVal ReturnedFromBP2K As Boolean)


            onOwnerID = OwnerID
            ostrCreditCode = CreditCode
            onFiscalYear = FiscalYear
            onFacilityID = FacilityID
            ostrInvoiceNumber = InvoiceNumber
            osAmount = Amount
            odtApplied = Applied
            ostrCheckNumber = CheckNumber
            ostrReason = Reason
            onItemSeqNumber = ItemSeqNumber
            bolReturnedFromBP2K = ReturnedFromBP2K
            obolDeleted = DELETED

            ostrCreatedBy = CREATED_BY
            odtCreatedOn = CREATE_DATE
            ostrModifiedBy = LAST_EDITED_BY
            odtModifiedOn = DATE_LAST_EDITED
            nID = ID
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Archive()
            obolDeleted = bolDeleted
            onOwnerID = nOwnerID
            ostrCreditCode = strCreditCode
            onFiscalYear = nFiscalYear
            onFacilityID = nFacilityID
            ostrInvoiceNumber = strInvoiceNumber
            osAmount = sAmount
            odtApplied = dtApplied
            ostrCheckNumber = strCheckNumber
            ostrReason = strReason
            onItemSeqNumber = nItemSeqNumber

            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn

            bolIsDirty = False
        End Sub
        Public Sub Reset()
            bolDeleted = obolDeleted
            bolIsDirty = False
            nOwnerID = onOwnerID
            strCreditCode = ostrCreditCode
            nFiscalYear = onFiscalYear
            nFacilityID = onFacilityID
            strInvoiceNumber = ostrInvoiceNumber
            sAmount = osAmount
            dtApplied = odtApplied
            strCheckNumber = ostrCheckNumber
            strReason = ostrReason
            nItemSeqNumber = onItemSeqNumber

            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn

        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            bolIsDirty = (bolDeleted <> obolDeleted) Or _
                            (nOwnerID <> onOwnerID) Or _
                            (nFiscalYear <> onFiscalYear) Or _
                            (nFacilityID <> onFacilityID) Or _
                            (strInvoiceNumber <> ostrInvoiceNumber) Or _
                            (sAmount <> osAmount) Or _
                            (dtApplied <> odtApplied) Or _
                            (strCheckNumber <> ostrCheckNumber) Or _
                            (nItemSeqNumber <> onItemSeqNumber) Or _
                            (strReason <> ostrReason)

        End Sub
        Private Sub Init()
            nID = 0
            obolDeleted = False
            bolIsDirty = False
            onOwnerID = 0
            ostrCreditCode = String.Empty
            onFiscalYear = 0
            onFacilityID = 0
            ostrInvoiceNumber = String.Empty
            osAmount = 0
            ostrCheckNumber = String.Empty
            ostrReason = String.Empty
            onItemSeqNumber = 0

            ostrCreatedBy = String.Empty
            odtCreatedOn = DateTime.Now.ToShortDateString
            ostrModifiedBy = String.Empty
            odtModifiedOn = DateTime.Now.ToShortDateString

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
        ' The base fee for the billing period
        Public Property Amount() As Decimal
            Get
                Return sAmount
            End Get
            Set(ByVal Value As Decimal)
                sAmount = Value
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


        Public Property ItemSeqNumber() As Int64
            Get
                Return nItemSeqNumber
            End Get
            Set(ByVal Value As Int64)
                nItemSeqNumber = Value
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

        Public Property CreditCode() As String
            Get
                Return strCreditCode
            End Get
            Set(ByVal Value As String)
                strCreditCode = Value
                Me.CheckDirty()
            End Set
        End Property
        ' The fiscal year for the billing period
        Public Property FiscalYear() As Integer
            Get
                Return nFiscalYear
            End Get
            Set(ByVal Value As Integer)
                nFiscalYear = Value
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
                Return strInvoiceNumber
            End Get
            Set(ByVal Value As String)
                strInvoiceNumber = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Applied() As Date
            Get
                Return dtApplied
            End Get
            Set(ByVal Value As Date)
                dtApplied = Value
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
        Public Property Reason() As String
            Get
                Return strReason
            End Get
            Set(ByVal Value As String)
                strReason = Value
                Me.CheckDirty()
            End Set
        End Property

        Public ReadOnly Property ReturnedFromBP2K() As Boolean
            Get
                Return bolReturnedFromBP2K
            End Get
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
                Return dtCreatedOn
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
                Return dtModifiedOn
            End Get
        End Property

#End Region
#Region "Protected Operations"
        Protected Overrides Sub Finalize()
        End Sub
#End Region

        Public Delegate Sub FeeAdjustmentInfoChangedEventHandler()


        ' Fired when CheckDirty determines that an attribute of the activity has been modified
        Public Event FeeAdjustmentInfoChanged As FeeAdjustmentInfoChangedEventHandler

    End Class
End Namespace
