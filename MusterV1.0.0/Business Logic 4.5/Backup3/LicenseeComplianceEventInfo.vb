'-------------------------------------------------------------------------------
' MUSTER.Info.LicenseeComplianceEventInfo
'   Provides the container to persist MUSTER LicenseeComplianceEvent state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0                              Original class definition
'  1.1       JC         01/12/05    Added line of code to RESET to raise
'                                       data changed event when called.
'
' Function          Description
' New()             Instantiates an empty LicenseeComplianceEventInfo object
' New(Deleted, CreatedBy, CreatedOn, ModifiedBy, LastEdited,)
'                   Instantiates a populated FacilityComplianceEventInfo object
' New(dr)           Instantiates a populated FacilityComplianceEventInfo object taking member state
'                   from the datarow provided
' Reset()           Sets the object state to the original state when loaded from or
'                   last saved to the repository
'
' NOTE: This file to be used as LicenseeComplianceEvent to build other objects.
'       Replace keyword "LicenseeComplianceEvent" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class LicenseeComplianceEventInfo
#Region "Public Events"
        Public Event LCEInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"

        Private nLCEID As Int32
        Private nLicenseeID As Int32
        Private nFacilityID As Int32
        Private nLicenseeCitationID As Int32
        Private dtCitationDueDate As Date
        Private dtCitationReceivedDate As Date
        Private bolRescinded As Boolean = False
        Private dtLCEDate As Date
        Private dtLCEProcessDate As Date
        Private dtNextDueDate As Date
        Private dtOverrideDueDate As Date
        Private nLCEStatus As Integer
        Private strStatus As String
        Private nEscalation As Integer
        Private dPolicyAmount As Decimal
        Private dOverrideAmount As Decimal
        Private dSettlementAmount As Decimal
        Private dPaidAmount As Decimal
        Private dtDateReceived As Date
        Private dtWorkShopDate As Date
        Private nWorkshopResult As Integer
        Private dtShowCauseDate As Date
        Private nShowCauseResults As Integer
        Private dtCommissionDate As Date
        Private nCommissionResults As Integer
        Private nPendingLetter As Integer
        Private dtLetterGenerated As Date
        Private bolLetterPrinted As Boolean
        Private bolDeleted As Boolean
        Private strCreatedBy As String
        Private dtCreatedOn As Date
        Private strModifiedBy As String
        Private dtModifiedOn As Date
        Private strLicenseeName As String
        Private strFacilityName As String
        Private strPendingLetter As String
        Private strShowCauseResult As String
        Private strCommissionResult As String
        Private strWorkshopResult As String
        Private strOwner As String
        Private nowner As Integer
        Private strCitationText As String
        Private strEscalationName As String
        Private nPendingLetterTemplateNum As Integer

        Private onLCEID As Int32
        Private onLicenseeID As Int32
        Private onFacilityID As Int32
        Private onLicenseeCitationID As Int32
        Private odtCitationDueDate As Date
        Private odtCitationReceivedDate As Date
        Private obolRescinded As Boolean = False
        Private odtLCEDate As Date
        Private odtLCEProcessDate As Date
        Private odtNextDueDate As Date
        Private odtOverrideDueDate As Date
        Private onLCEStatus As Integer
        Private ostrStatus As String
        Private onEscalation As Integer
        Private odPolicyAmount As Decimal
        Private odOverrideAmount As Decimal
        Private odSettlementAmount As Decimal
        Private odPaidAmount As Decimal
        Private odtDateReceived As Date
        Private odtWorkShopDate As Date
        Private onWorkshopResult As Integer
        Private odtShowCauseDate As Date
        Private onShowCauseResults As Integer
        Private odtCommissionDate As Date
        Private onCommissionResults As Integer
        Private onPendingLetter As Integer
        Private odtLetterGenerated As Date
        Private obolLetterPrinted As Boolean
        Private obolDeleted As Boolean
        Private ostrCreatedBy As String
        Private odtCreatedOn As Date
        Private ostrModifiedBy As String
        Private odtModifiedOn As Date
        Private ostrLicenseeName As String
        Private ostrFacilityName As String
        Private ostrPendingLetter As String
        Private ostrShowCauseResult As String
        Private ostrCommissionResult As String
        Private ostrWorkshopResult As String
        Private ostrOwner As String
        Private onowner As Integer
        Private ostrCitationText As String
        Private ostrEscalationName As String
        Private onPendingLetterTemplateNum As Integer

        Private bolShowDeleted As Boolean = False

        Private bolIsDirty As Boolean = False
        Private MusterException As New MUSTER.Exceptions.MusterExceptions
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            Me.Init()
        End Sub
        Sub New(ByVal ID As Integer, _
        ByVal LicenseeID As Integer, _
        ByVal ownerID As Integer, _
        ByVal ownerName As String, _
        ByVal Facility As String, _
        ByVal Licensee As String, _
        ByVal FacilityID As Integer, _
        ByVal LicenseeCitationID As Integer, _
        ByVal CitationDueDate As Date, _
        ByVal CitationReceivedDate As Date, _
        ByVal Rescinded As Boolean, _
        ByVal LCEDate As Date, _
        ByVal LCEProcessDate As Date, _
        ByVal NextDueDate As Date, _
        ByVal OverrideDueDate As Date, _
        ByVal LCEStatus As Integer, _
        ByVal Status As String, _
        ByVal Escalation As Integer, _
        ByVal EscalationName As String, _
        ByVal PolicyAmount As Decimal, _
        ByVal OverrideAmount As Decimal, _
        ByVal SettlementAmount As Decimal, _
        ByVal PaidAmount As Decimal, _
        ByVal DateReceived As Date, _
        ByVal WorkShopDate As Date, _
        ByVal WorkshopResult As Integer, _
        ByVal WorkShopResultString As String, _
        ByVal ShowCauseDate As Date, _
        ByVal ShowCauseResults As Integer, _
        ByVal ShowCauseResultString As String, _
        ByVal CommissionDate As Date, _
        ByVal CommissionResults As Integer, _
        ByVal CommissionResultString As String, _
        ByVal PendingLetter As Integer, _
        ByVal PendingLetterString As String, _
        ByVal LetterGenerated As Date, _
        ByVal LetterPrinted As Boolean, _
        ByVal CitationText As String, _
        ByVal LastEdited As Date, _
        ByVal CreatedBy As String, _
        ByVal CreatedOn As Date, _
        ByVal ModifiedBy As String, _
        ByVal Deleted As Boolean, _
        ByVal pendingLetterTemplateNum As Integer)

            onLCEID = ID
            ostrCitationText = CitationText
            onowner = ownerID
            ostrOwner = ownerName
            ostrLicenseeName = Licensee
            ostrFacilityName = Facility
            onLicenseeID = LicenseeID
            onFacilityID = FacilityID
            onLicenseeCitationID = LicenseeCitationID
            odtCitationDueDate = CitationDueDate
            odtCitationReceivedDate = CitationReceivedDate
            obolRescinded = Rescinded
            odtLCEDate = LCEDate
            odtLCEProcessDate = LCEProcessDate
            odtNextDueDate = NextDueDate
            odtOverrideDueDate = OverrideDueDate
            onLCEStatus = LCEStatus
            ostrStatus = Status.ToUpper
            onEscalation = Escalation
            ostrEscalationName = EscalationName
            odPolicyAmount = PolicyAmount
            odOverrideAmount = OverrideAmount
            odSettlementAmount = SettlementAmount
            odPaidAmount = PaidAmount
            odtDateReceived = DateReceived
            odtWorkShopDate = WorkShopDate
            onWorkshopResult = WorkshopResult
            odtShowCauseDate = ShowCauseDate
            onShowCauseResults = ShowCauseResults
            odtCommissionDate = CommissionDate
            onCommissionResults = CommissionResults
            onPendingLetter = PendingLetter
            ostrPendingLetter = PendingLetterString.ToUpper
            ostrWorkshopResult = WorkShopResultString.ToUpper
            ostrCommissionResult = CommissionResultString.ToUpper
            ostrShowCauseResult = ShowCauseResultString.ToUpper
            odtLetterGenerated = LetterGenerated
            obolLetterPrinted = LetterPrinted
            ostrCreatedBy = CreatedBy
            odtCreatedOn = CreatedOn
            ostrModifiedBy = ModifiedBy
            odtModifiedOn = ModifiedOn
            obolDeleted = Deleted
            onPendingLetterTemplateNum = pendingLetterTemplateNum
            Me.Reset()

        End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()

            If onLCEID > 0 Then
                nLCEID = onLCEID
            End If

            strCitationText = ostrCitationText
            nowner = onowner
            strOwner = ostrOwner
            strLicenseeName = ostrLicenseeName
            strFacilityName = ostrFacilityName
            nLicenseeID = onLicenseeID
            nFacilityID = onFacilityID
            nLicenseeCitationID = onLicenseeCitationID
            dtCitationDueDate = odtCitationDueDate
            dtCitationReceivedDate = odtCitationReceivedDate
            bolRescinded = obolRescinded
            dtLCEDate = odtLCEDate
            dtLCEProcessDate = odtLCEProcessDate
            dtNextDueDate = odtNextDueDate
            dtOverrideDueDate = odtOverrideDueDate
            nLCEStatus = onLCEStatus
            strStatus = ostrStatus
            nEscalation = onEscalation
            strEscalationName = ostrEscalationName
            dPolicyAmount = odPolicyAmount
            dOverrideAmount = odOverrideAmount
            dSettlementAmount = odSettlementAmount
            dPaidAmount = odPaidAmount
            dtDateReceived = odtDateReceived
            dtWorkShopDate = odtWorkShopDate
            nWorkshopResult = onWorkshopResult
            dtShowCauseDate = odtShowCauseDate
            nShowCauseResults = onShowCauseResults
            dtCommissionDate = odtCommissionDate
            nCommissionResults = onCommissionResults
            nPendingLetter = onPendingLetter
            strPendingLetter = ostrPendingLetter
            strWorkshopResult = ostrWorkshopResult
            strCommissionResult = ostrCommissionResult
            strShowCauseResult = ostrShowCauseResult
            dtLetterGenerated = odtLetterGenerated
            bolLetterPrinted = obolLetterPrinted
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            bolDeleted = obolDeleted
            nPendingLetterTemplateNum = onPendingLetterTemplateNum

            bolIsDirty = False
            RaiseEvent LCEInfoChanged(bolIsDirty)
        End Sub
        Public Sub Archive()

            onLCEID = nLCEID
            ostrCitationText = strCitationText
            ostrOwner = strOwner
            onowner = nowner
            ostrLicenseeName = strLicenseeName
            ostrFacilityName = strFacilityName
            onLicenseeID = nLicenseeID
            onFacilityID = nFacilityID
            onLicenseeCitationID = nLicenseeCitationID
            odtCitationDueDate = dtCitationDueDate
            odtCitationReceivedDate = dtCitationReceivedDate
            obolRescinded = bolRescinded
            odtLCEDate = dtLCEDate
            odtLCEProcessDate = dtLCEProcessDate
            odtNextDueDate = dtNextDueDate
            odtOverrideDueDate = dtOverrideDueDate
            onLCEStatus = nLCEStatus
            ostrStatus = strStatus
            onEscalation = nEscalation
            ostrEscalationName = strEscalationName
            odPolicyAmount = dPolicyAmount
            odOverrideAmount = dOverrideAmount
            odSettlementAmount = dSettlementAmount
            odPaidAmount = dPaidAmount
            odtDateReceived = dtDateReceived
            odtWorkShopDate = dtWorkShopDate
            onWorkshopResult = nWorkshopResult
            odtShowCauseDate = dtShowCauseDate
            onShowCauseResults = nShowCauseResults
            odtCommissionDate = dtCommissionDate
            onCommissionResults = nCommissionResults
            onPendingLetter = nPendingLetter
            ostrPendingLetter = strPendingLetter
            ostrWorkshopResult = strWorkshopResult
            ostrCommissionResult = strCommissionResult
            ostrShowCauseResult = strShowCauseResult
            odtLetterGenerated = dtLetterGenerated
            obolLetterPrinted = bolLetterPrinted
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn
            obolDeleted = bolDeleted
            onPendingLetterTemplateNum = nPendingLetterTemplateNum

            bolIsDirty = False

        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            Dim obolIsDirty As Boolean = bolIsDirty
            '(nLCEID <> onLCEID) Or _
            bolIsDirty = (strCitationText <> ostrCitationText) Or _
            (nowner <> onowner) Or _
            (strOwner <> ostrOwner) Or _
            (strLicenseeName <> ostrLicenseeName) Or _
            (strFacilityName <> ostrFacilityName) Or _
            (nLicenseeID <> onLicenseeID) Or _
            (nFacilityID <> onFacilityID) Or _
            (nLicenseeCitationID <> onLicenseeCitationID) Or _
            (dtCitationDueDate <> odtCitationDueDate) Or _
            (dtCitationReceivedDate <> odtCitationReceivedDate) Or _
            (bolRescinded <> obolRescinded) Or _
            (dtLCEDate <> odtLCEDate) Or _
            (dtLCEProcessDate <> odtLCEProcessDate) Or _
            (dtNextDueDate <> odtNextDueDate) Or _
            (dtOverrideDueDate <> odtOverrideDueDate) Or _
            (nLCEStatus <> onLCEStatus) Or _
            (strStatus <> ostrStatus) Or _
            (nEscalation <> onEscalation) Or _
            (strEscalationName <> ostrEscalationName) Or _
            (dPolicyAmount <> odPolicyAmount) Or _
            (dOverrideAmount <> odOverrideAmount) Or _
            (dSettlementAmount <> odSettlementAmount) Or _
            (dPaidAmount <> odPaidAmount) Or _
            (dtDateReceived <> odtDateReceived) Or _
            (dtWorkShopDate <> odtWorkShopDate) Or _
            (nWorkshopResult <> onWorkshopResult) Or _
            (dtShowCauseDate <> odtShowCauseDate) Or _
            (nShowCauseResults <> onShowCauseResults) Or _
            (dtCommissionDate <> odtCommissionDate) Or _
            (nCommissionResults <> onCommissionResults) Or _
            (nPendingLetter <> onPendingLetter) Or _
            (ostrPendingLetter <> strPendingLetter) Or _
            (ostrWorkshopResult <> strWorkshopResult) Or _
            (ostrCommissionResult <> strCommissionResult) Or _
            (ostrShowCauseResult <> strShowCauseResult) Or _
            (dtLetterGenerated <> odtLetterGenerated) Or _
            (bolLetterPrinted <> obolLetterPrinted) Or _
            (dtModifiedOn <> odtModifiedOn) Or _
            (strCreatedBy <> ostrCreatedBy) Or _
            (dtCreatedOn <> odtCreatedOn) Or _
            (strModifiedBy <> ostrModifiedBy) Or _
            (bolDeleted <> obolDeleted) Or _
            (nPendingLetterTemplateNum <> onPendingLetterTemplateNum)

            If obolIsDirty <> bolIsDirty Then
                RaiseEvent LCEInfoChanged(bolIsDirty)
            End If

        End Sub
        Private Sub Init()
            onLCEID = 0
            ostrCitationText = String.Empty
            ostrOwner = String.Empty
            onowner = 0
            ostrLicenseeName = String.Empty
            ostrFacilityName = String.Empty
            onLicenseeID = 0
            onFacilityID = 0
            onLicenseeCitationID = 0
            odtCitationDueDate = CDate("01/01/0001")
            odtCitationReceivedDate = CDate("01/01/0001")
            obolRescinded = False
            odtLCEDate = CDate("01/01/0001")
            odtLCEProcessDate = CDate("01/01/0001")
            odtNextDueDate = CDate("01/01/0001")
            odtOverrideDueDate = CDate("01/01/0001")
            onLCEStatus = 0
            onEscalation = 0
            ostrEscalationName = String.Empty
            odPolicyAmount = -1.0
            odOverrideAmount = -1.0
            odSettlementAmount = -1.0
            odPaidAmount = -1.0
            odtDateReceived = CDate("01/01/0001")
            odtWorkShopDate = CDate("01/01/0001")
            onWorkshopResult = WorkshopResult
            odtShowCauseDate = CDate("01/01/0001")
            onShowCauseResults = 0
            odtCommissionDate = CDate("01/01/0001")
            onCommissionResults = 0
            onPendingLetter = 0
            ostrPendingLetter = String.Empty
            ostrWorkshopResult = String.Empty
            ostrCommissionResult = String.Empty
            ostrShowCauseResult = String.Empty
            odtLetterGenerated = CDate("01/01/0001")
            obolLetterPrinted = False
            bolDeleted = False
            strCreatedBy = String.Empty
            dtCreatedOn = System.DateTime.Now
            strModifiedBy = String.Empty
            dtModifiedOn = System.DateTime.Now
            onPendingLetterTemplateNum = 0
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property EscalationName() As String
            Get
                Return strEscalationName
            End Get
            Set(ByVal Value As String)
                strEscalationName = Value
            End Set
        End Property
        Public Property CitationText() As String
            Get
                Return strCitationText
            End Get
            Set(ByVal Value As String)
                strCitationText = Value
            End Set
        End Property
        Public Property OwnerID() As Integer
            Get
                Return nowner
            End Get
            Set(ByVal Value As Integer)
                nowner = Value
            End Set
        End Property
        Public Property OwnerName() As String
            Get
                Return strOwner
            End Get
            Set(ByVal Value As String)
                strOwner = Value
            End Set
        End Property
        Public Property LicenseeName() As String
            Get
                Return strLicenseeName
            End Get
            Set(ByVal Value As String)
                strLicenseeName = Value
            End Set
        End Property
        Public Property facilityName() As String
            Get
                Return strFacilityName
            End Get
            Set(ByVal Value As String)
                strFacilityName = Value
            End Set
        End Property
        Public Property ID() As Int32
            Get
                Return nLCEID
            End Get
            Set(ByVal Value As Int32)
                nLCEID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LicenseeID() As Int32
            Get
                Return nLicenseeID
            End Get
            Set(ByVal Value As Int32)
                nLicenseeID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property FacilityID() As Int32
            Get
                Return nFacilityID
            End Get
            Set(ByVal Value As Int32)
                nFacilityID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LicenseeCitationID() As Int32
            Get
                Return nLicenseeCitationID
            End Get
            Set(ByVal Value As Int32)
                nLicenseeCitationID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CitationDueDate() As Date
            Get
                Return dtCitationDueDate
            End Get

            Set(ByVal Value As Date)
                dtCitationDueDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CitationReceivedDate() As Date
            Get
                Return dtCitationReceivedDate
            End Get
            Set(ByVal Value As Date)
                dtCitationReceivedDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Rescinded() As Boolean
            Get
                Return bolRescinded
            End Get
            Set(ByVal Value As Boolean)
                bolRescinded = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LCEDate() As Date
            Get
                Return dtLCEDate
            End Get
            Set(ByVal Value As Date)
                dtLCEDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LCEProcessDate() As Date
            Get
                Return dtLCEProcessDate
            End Get
            Set(ByVal Value As Date)
                dtLCEProcessDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property NextDueDate() As Date
            Get
                Return dtNextDueDate
            End Get
            Set(ByVal Value As Date)
                dtNextDueDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property OverrideDueDate() As Date
            Get
                Return dtOverrideDueDate
            End Get
            Set(ByVal Value As Date)
                dtOverrideDueDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LCEStatus() As Integer
            Get
                Return nLCEStatus
            End Get
            Set(ByVal Value As Integer)
                nLCEStatus = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Status() As String
            Get
                Return strStatus
            End Get
            Set(ByVal Value As String)
                strStatus = Value
            End Set
        End Property
        Public Property Escalation() As Integer
            Get
                Return nEscalation
            End Get
            Set(ByVal Value As Integer)
                nEscalation = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property PolicyAmount() As Decimal
            Get
                Return dPolicyAmount
            End Get
            Set(ByVal Value As Decimal)
                dPolicyAmount = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property OverrideAmount() As Decimal
            Get
                Return dOverrideAmount
            End Get
            Set(ByVal Value As Decimal)
                dOverrideAmount = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property SettlementAmount() As Decimal
            Get
                Return dSettlementAmount
            End Get
            Set(ByVal Value As Decimal)
                dSettlementAmount = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property PaidAmount() As Decimal
            Get
                Return dPaidAmount
            End Get
            Set(ByVal Value As Decimal)
                dPaidAmount = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property DateReceived() As Date
            Get
                Return dtDateReceived
            End Get
            Set(ByVal Value As Date)
                dtDateReceived = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property WorkShopDate() As Date
            Get
                Return dtWorkShopDate
            End Get
            Set(ByVal Value As Date)
                dtWorkShopDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property WorkshopResult() As Integer
            Get
                Return nWorkshopResult
            End Get
            Set(ByVal Value As Integer)
                nWorkshopResult = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ShowCauseDate() As Date
            Get
                Return dtShowCauseDate
            End Get
            Set(ByVal Value As Date)
                dtShowCauseDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ShowCauseResults() As Integer
            Get
                Return nShowCauseResults
            End Get
            Set(ByVal Value As Integer)
                nShowCauseResults = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CommissionDate() As Date
            Get
                Return dtCommissionDate
            End Get
            Set(ByVal Value As Date)
                dtCommissionDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CommissionResults() As Integer
            Get
                Return nCommissionResults
            End Get
            Set(ByVal Value As Integer)
                nCommissionResults = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property PendingLetter() As Integer
            Get
                Return nPendingLetter
            End Get
            Set(ByVal Value As Integer)
                nPendingLetter = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LetterGenerated() As Date
            Get
                Return dtLetterGenerated
            End Get
            Set(ByVal Value As Date)
                dtLetterGenerated = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property LetterPrinted() As Boolean
            Get
                Return bolLetterPrinted
            End Get
            Set(ByVal Value As Boolean)
                bolLetterPrinted = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CreatedBy() As String
            Get
                Return strCreatedBy
            End Get
            Set(ByVal Value As String)
                strCreatedBy = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property CreatedOn() As Date
            Get
                Return dtCreatedOn
            End Get
            Set(ByVal Value As Date)
                dtCreatedOn = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ModifiedBy() As String
            Get
                Return strModifiedBy
            End Get
            Set(ByVal Value As String)
                strModifiedBy = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property ModifiedOn() As Date
            Get
                Return dtModifiedOn
            End Get
            Set(ByVal Value As Date)
                dtModifiedOn = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Deleted() As Boolean
            Get
                Return bolDeleted
            End Get

            Set(ByVal value As Boolean)
                bolDeleted = value
                Me.CheckDirty()
            End Set
        End Property
        Public Property PendingLetterTemplateNum() As Integer
            Get
                Return nPendingLetterTemplateNum
            End Get
            Set(ByVal Value As Integer)
                nPendingLetterTemplateNum = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property PendingLetterName() As String
            Get
                Return strPendingLetter
            End Get
            Set(ByVal Value As String)
                strPendingLetter = Value
            End Set
        End Property

        Public Property WorkshopResultName() As String
            Get
                Return strWorkshopResult
            End Get
            Set(ByVal Value As String)
                strWorkshopResult = Value
            End Set
        End Property

        Public Property ShowCauseResultName() As String
            Get
                Return strShowCauseResult
            End Get
            Set(ByVal Value As String)
                strShowCauseResult = Value
            End Set
        End Property
        Public Property CommissionResultName() As String
            Get
                Return strCommissionResult
            End Get
            Set(ByVal Value As String)
                strCommissionResult = Value
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
#End Region
#Region "Protected Operations"
        Protected Overrides Sub Finalize()
            MyBase.Finalize()
        End Sub
#End Region
    End Class
End Namespace
