
Namespace MUSTER.Info
    ' -------------------------------------------------------------------------------
    '    MUSTER.Info.FeeLateFeeInfo
    '          Provides the container to persist MUSTER Late Fee state
    ' 
    '    Copyright (C) 2004, 2005 CIBER, Inc.
    '    All rights reserved.
    ' 
    '    Release   Initials    Date        Description
    '       1.0        AB       12/05/05    Original class definition.
    ' 
    '    Function          Description
    ' -------------------------------------------------------------------------------
    '
    Public Class FeeLateFeeInfo
#Region "Private member variables"

        Private nID As Integer
        Private dtDataAge As DateTime
        Private nAgeThreshold As Int16 = 5
        Private bolIsDirty As Boolean = False
        Private dtCreatedOn As Date = DateTime.Now.ToShortDateString
        Private dtModifiedOn As Date = DateTime.Now.ToShortDateString
        Private strCreatedBy As String = String.Empty
        Private strModifiedBy As String = String.Empty

        Private nFiscalYear As Integer
        Private strInvoiceNumber As String
        Private Charges As Decimal
        Private strCertLetterNumber As String
        Private bolWaiveApprovalRec As Boolean
        Private bolWaiveApprovalStatus As Boolean
        Private nWaiveReason As Long
        Private dtWaiverFinalizedOn As Date
        Private bolProcessCertification As Boolean
        Private bolProcessWaiver As Boolean
        Private bolDeleted As Boolean


        Private onFiscalYear As Integer
        Private ostrInvoiceNumber As String
        Private oCharges As Decimal
        Private ostrCertLetterNumber As String
        Private obolWaiveApprovalRec As Boolean
        Private obolWaiveApprovalStatus As Boolean
        Private onWaiveReason As Long
        Private obolProcessCertification As Boolean
        Private obolProcessWaiver As Boolean
        Private odtWaiverFinalizedOn As Date
        Private obolDeleted As Boolean

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
                        ByVal FISCAL_YEAR As Long, _
                        ByVal CurrentCharges As Decimal, _
                        ByVal InvoiceNumber As String, _
                        ByVal CertifiedLetterNumber As String, _
                        ByVal WaiveApprovalRec As Boolean, _
                        ByVal WaiveApprovalStatus As Boolean, _
                        ByVal WaiveReason As Int64, _
                        ByVal ProcessedCert As Boolean, _
                        ByVal ProcessedWaiver As Boolean, _
                        ByVal WaiverFinalizedOn As Date, _
                        ByVal DELETED As Boolean)

            obolDeleted = DELETED
            oCharges = CurrentCharges
            ostrInvoiceNumber = InvoiceNumber
            ostrCertLetterNumber = CertifiedLetterNumber
            obolWaiveApprovalRec = WaiveApprovalRec
            obolWaiveApprovalStatus = WaiveApprovalStatus
            onWaiveReason = WaiveReason
            obolProcessCertification = ProcessedCert
            obolProcessWaiver = ProcessedWaiver
            onFiscalYear = FISCAL_YEAR
            odtWaiverFinalizedOn = WaiverFinalizedOn
            odtCreatedOn = CREATE_DATE
            odtModifiedOn = DATE_LAST_EDITED
            ostrCreatedBy = CREATED_BY
            ostrModifiedBy = LAST_EDITED_BY

            nID = ID
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Archive()
            obolDeleted = bolDeleted
            bolIsDirty = False
            oCharges = Charges
            onFiscalYear = nFiscalYear
            ostrInvoiceNumber = strInvoiceNumber
            ostrCertLetterNumber = strCertLetterNumber
            obolWaiveApprovalRec = bolWaiveApprovalRec
            obolWaiveApprovalStatus = bolWaiveApprovalStatus
            onWaiveReason = nWaiveReason
            obolProcessCertification = bolProcessCertification
            obolProcessWaiver = bolProcessWaiver
            odtWaiverFinalizedOn = dtWaiverFinalizedOn

            odtCreatedOn = dtCreatedOn
            odtModifiedOn = dtModifiedOn
            ostrCreatedBy = strCreatedBy
            ostrModifiedBy = strModifiedBy

        End Sub
        Public Sub Reset()
            bolDeleted = obolDeleted
            bolIsDirty = False
            Charges = oCharges
            nFiscalYear = onFiscalYear
            strInvoiceNumber = ostrInvoiceNumber
            strCertLetterNumber = ostrCertLetterNumber
            bolWaiveApprovalRec = obolWaiveApprovalRec
            bolWaiveApprovalStatus = obolWaiveApprovalStatus
            nWaiveReason = onWaiveReason
            bolProcessCertification = obolProcessCertification
            bolProcessWaiver = obolProcessWaiver
            dtWaiverFinalizedOn = odtWaiverFinalizedOn

            dtCreatedOn = odtCreatedOn
            dtModifiedOn = odtModifiedOn
            strCreatedBy = ostrCreatedBy
            strModifiedBy = ostrModifiedBy

        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            bolIsDirty = (bolDeleted <> obolDeleted) Or _
                            (Charges <> oCharges) Or _
                            (nFiscalYear <> onFiscalYear) Or _
                            (strCertLetterNumber <> ostrCertLetterNumber) Or _
                            (bolWaiveApprovalRec <> obolWaiveApprovalRec) Or _
                            (obolWaiveApprovalStatus <> bolWaiveApprovalStatus) Or _
                            (nWaiveReason <> onWaiveReason) Or _
                            (bolProcessCertification <> obolProcessCertification) Or _
                            (bolProcessWaiver <> obolProcessWaiver) Or _
                            (dtWaiverFinalizedOn <> odtWaiverFinalizedOn) Or _
                            (strInvoiceNumber <> ostrInvoiceNumber)
        End Sub
        Private Sub Init()
            nID = 0
            obolDeleted = False
            bolIsDirty = False
            oCharges = 0
            onFiscalYear = 0
            ostrInvoiceNumber = String.Empty
            ostrCertLetterNumber = String.Empty
            obolWaiveApprovalRec = False
            obolWaiveApprovalStatus = False
            onWaiveReason = 0
            obolProcessCertification = False
            obolProcessWaiver = False
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

        ' The base fee for the billing period
        Public Property LateCharges() As Decimal
            Get
                Return Charges
            End Get
            Set(ByVal Value As Decimal)
                Charges = Value
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
        Public Property InvoiceNumber() As String
            Get
                Return strInvoiceNumber
            End Get
            Set(ByVal Value As String)
                strInvoiceNumber = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CertLetterNumber() As String
            Get
                Return strCertLetterNumber
            End Get
            Set(ByVal Value As String)
                strCertLetterNumber = Value
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

        Public Property WaiveApprovalRecommendation() As Boolean
            Get
                Return bolWaiveApprovalRec
            End Get
            Set(ByVal Value As Boolean)
                bolWaiveApprovalRec = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property WaiveApprovalStatus() As Boolean
            Get
                Return bolWaiveApprovalStatus
            End Get
            Set(ByVal Value As Boolean)
                bolWaiveApprovalStatus = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property WaiveReason() As Int64
            Get
                Return nWaiveReason
            End Get
            Set(ByVal Value As Int64)
                nWaiveReason = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property ProcessCertification() As Boolean
            Get
                Return bolProcessCertification
            End Get
            Set(ByVal Value As Boolean)
                bolProcessCertification = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ProcessWaiver() As Boolean
            Get
                Return bolProcessWaiver
            End Get
            Set(ByVal Value As Boolean)
                bolProcessWaiver = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property WaiverFinalizedOn() As Date
            Get
                Return dtWaiverFinalizedOn
            End Get
            Set(ByVal Value As Date)
                dtWaiverFinalizedOn = Value
                Me.CheckDirty()
            End Set
        End Property

        Public ReadOnly Property CreateDate() As Date
            Get
                Return dtCreatedOn
            End Get
        End Property

        Public ReadOnly Property ModifiedDate() As Date
            Get
                Return dtModifiedOn
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
        Public Property ModifiedBy() As String
            Get
                Return strModifiedBy
            End Get
            Set(ByVal Value As String)
                strModifiedBy = Value
            End Set
        End Property
#End Region
#Region "Protected Operations"
        Protected Overrides Sub Finalize()
        End Sub
#End Region

        Public Delegate Sub FeeLateFeeInfoChangedEventHandler()


        ' Fired when CheckDirty determines that an attribute of the activity has been modified
        Public Event FeeLateFeeInfoChanged As FeeLateFeeInfoChangedEventHandler

    End Class
End Namespace
