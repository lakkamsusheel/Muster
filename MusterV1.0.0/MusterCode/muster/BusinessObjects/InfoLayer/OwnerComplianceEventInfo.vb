'-------------------------------------------------------------------------------
' MUSTER.Info.OwnerComplianceEventInfo
'   Provides the container to persist MUSTER OwnerComplianceEvent state
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'  1.0        MR        7/1/2005         Original class definition
'  1.1  Thomas Franey   5/21/2009         added comments to info
'
' Function          Description
' New()             Instantiates an empty OwnerComplianceEventInfo object
' New(Deleted, CreatedBy, CreatedOn, ModifiedBy, LastEdited, OwnerL2CSnippet)
'                   Instantiates a populated OwnerComplianceEventInfo object
' New(dr)           Instantiates a populated OwnerComplianceEventInfo object taking member state
'                   from the datarow provided
' Reset()           Sets the object state to the original state when loaded from or
'                   last saved to the repository
'
' NOTE: This file to be used as OwnerComplianceEvent to build other objects.
'       Replace keyword "OwnerComplianceEvent" with respective Object name.
'       Don't forget to update the information in this header to reflect the
'       attributes and operations of the Info object!
'-------------------------------------------------------------------------------

Namespace MUSTER.Info
    <Serializable()> _
    Public Class OwnerComplianceEventInfo
#Region "Public Events"
        Public Event OwnerComplianceEventInfoChanged(ByVal bolValue As Boolean)
#End Region
#Region "Private member variables"

        Private nOCEID As Int32
        Private nOwnerID As Int32
        Private nCitation As Integer
        Private dtCitationDueDate As Date
        Private bolRescinded As Boolean
        Private nOCEPath As Integer
        Private dtOCEDate As Date
        Private dtOCEProcessDate As Date
        Private dtNextDueDate As Date
        Private dtOverrideDueDate As Date
        Private nOCEStatus As Integer
        Private nEscalation As Integer
        Private dPolicyAmount As Decimal
        Private dOverRideAmount As Decimal
        Private dSettlementAmount As Decimal
        Private dPaidAmount As Decimal
        Private dtDateReceived As Date
        Private dtWorkShopDate As Date
        Private nWorkShopResult As Integer
        Private bolWorkshopRequired As Boolean
        Private dtShowCauseDate As Date
        Private nShowCauseResult As Integer
        Private dtAdminDate As Date
        Private nAdminResult As Integer

        Private dtCommissionDate As Date
        Private nCommissionResult As Integer
        Private strAgreedOrder As String
        Private strAdministrativeOrder As String
        Private nPendingLetter As Integer
        Private dtLetterGenerated As Date
        Private dtRedTagDate As Date
        Private bolLetterPrinted As Boolean
        Private strCreatedBy As String
        Private dtCreatedOn As DateTime
        Private strModifiedBy As String
        Private dtModifiedOn As DateTime
        Private bolDeleted As Boolean
        Private nPendingLetterTemplateNum As Integer
        Private nEnsiteID As Integer
        Private strComments As String


        Private onOCEID As Int32
        Private onOwnerID As Int32
        Private onCitation As Integer
        Private odtCitationDueDate As Date
        Private obolRescinded As Boolean
        Private onOCEPath As Integer
        Private odtOCEDate As Date
        Private odtOCEProcessDate As Date
        Private odtNextDueDate As Date
        Private odtOverrideDueDate As Date
        Private onOCEStatus As Integer
        Private onEscalation As Integer
        Private odPolicyAmount As Decimal
        Private odOverRideAmount As Decimal
        Private odSettlementAmount As Decimal
        Private odPaidAmount As Decimal
        Private odtDateReceived As Date
        Private odtWorkShopDate As Date
        Private onWorkShopResult As Integer
        Private obolWorkshopRequired As Boolean
        Private odtShowCauseDate As Date
        Private onShowCauseResult As Integer
        Private odtCommissionDate As Date
        Private onCommissionResult As Integer
        Private ostrAgreedOrder As String
        Private ostrAdministrativeOrder As String
        Private onPendingLetter As Integer
        Private odtLetterGenerated As Date
        Private odtRedTagDate As Date
        Private obolLetterPrinted As Boolean
        Private ostrCreatedBy As String
        Private odtCreatedOn As DateTime
        Private ostrModifiedBy As String
        Private odtModifiedOn As DateTime
        Private obolDeleted As Boolean
        Private onPendingLetterTemplateNum As Integer
        Private onEnsiteID As Integer
        Private ostrComments As String

        Private bolShowDeleted As Boolean = False
        Private strEscalation As String

        Private odtAdminDate As Date
        Private onAdminResult As Integer

        Private bolIsDirty As Boolean = False
        Private MusterException As New MUSTER.Exceptions.MusterExceptions

        'Private strPendingLetter As String
        'Private strShowCauseResult As String
        'Private strCommissionResult As String
        'Private strWorkshopResult As String
        'Private strOwner As String
        'Private strStatus As String

        'Private ostrPendingLetter As String
        'Private ostrShowCauseResult As String
        'Private ostrCommissionResult As String
        'Private ostrWorkshopResult As String
        'Private ostrOwner As String
        'Private ostrStatus As String
#End Region
#Region "Constructors"
        Sub New()
            MyBase.New()
            Me.Init()
        End Sub
        Sub New(ByVal OCEID As Int32, _
                ByVal OwnerID As Int32, _
                ByVal Citation As Int32, _
                ByVal CitationDueDate As Date, _
                ByVal Rescinded As Boolean, _
                ByVal OCEPath As Integer, _
                ByVal OCEDate As Date, _
                ByVal OCEProcessDate As Date, _
                ByVal NextDueDate As Date, _
                ByVal OverrideDueDate As Date, _
                ByVal OCEStatus As Integer, _
                ByVal Escalation As Integer, _
                ByVal PolicyAmount As Decimal, _
                ByVal OverRideAmount As Decimal, _
                ByVal SettlementAmount As Decimal, _
                ByVal PaidAmount As Decimal, _
                ByVal DateReceived As Date, _
                ByVal WorkShopDate As Date, _
                ByVal WorkshopResult As Integer, _
                ByVal WorkShopRequired As Boolean, _
                ByVal ShowCauseDate As Date, _
                ByVal ShowCauseResults As Integer, _
                ByVal CommissionDate As Date, _
                ByVal CommissionResults As Integer, _
                ByVal AgreedOrder As String, _
                ByVal AdministrativeOrder As String, _
                ByVal PendingLetter As Integer, _
                ByVal LetterGenerated As Date, _
                ByVal RedTagDate As Date, _
                ByVal LetterPrinted As Boolean, _
                ByVal CreatedBy As String, _
                ByVal CreatedOn As Date, _
                ByVal ModifiedBy As String, _
                ByVal ModifiedOn As Date, _
                ByVal Deleted As Boolean, _
                ByVal escalationString As String, _
                ByVal pendingLetterTemplateNum As Integer, _
                ByVal ensiteID As Integer, Optional ByVal comments As String = "", Optional ByVal admin_hear_date As Object = Nothing, Optional ByVal admin_result As Object = Nothing)
            onOCEID = OCEID
            onOwnerID = OwnerID
            onCitation = Citation
            odtCitationDueDate = CitationDueDate
            obolRescinded = Rescinded
            onOCEPath = OCEPath
            odtOCEDate = OCEDate
            odtOCEProcessDate = OCEProcessDate
            odtNextDueDate = NextDueDate
            odtOverrideDueDate = OverrideDueDate
            onOCEStatus = OCEStatus
            onEscalation = Escalation
            odPolicyAmount = PolicyAmount
            odOverRideAmount = OverRideAmount
            odSettlementAmount = SettlementAmount
            odPaidAmount = PaidAmount
            odtDateReceived = DateReceived
            obolWorkshopRequired = bolWorkshopRequired
            odtWorkShopDate = WorkShopDate
            onWorkShopResult = WorkshopResult
            odtShowCauseDate = ShowCauseDate
            onShowCauseResult = ShowCauseResult
            odtCommissionDate = CommissionDate
            onCommissionResult = CommissionResult
            ostrAgreedOrder = AgreedOrder
            ostrAdministrativeOrder = AdministrativeOrder
            onPendingLetter = PendingLetter
            odtLetterGenerated = LetterGenerated
            odtRedTagDate = RedTagDate
            obolLetterPrinted = LetterPrinted
            ostrCreatedBy = CreatedBy
            odtCreatedOn = CreatedOn
            ostrModifiedBy = ModifiedBy
            odtModifiedOn = ModifiedOn
            obolDeleted = Deleted
            strEscalation = escalationString
            onPendingLetterTemplateNum = pendingLetterTemplateNum
            onEnsiteID = ensiteID
            ostrComments = comments
            If Not admin_hear_date Is Nothing AndAlso TypeOf admin_hear_date Is Date Then
                odtAdminDate = admin_hear_date
            Else
                odtAdminDate = CDate("01/01/0001")
            End If

            If Not admin_result Is Nothing AndAlso TypeOf admin_result Is Integer Then
                onAdminResult = admin_result
            Else
                onAdminResult = 0
            End If

            Me.Reset()
        End Sub
        Sub New(ByVal dr As DataRow)
            Try
                onOCEID = dr.Item("OCE_ID")
                onOwnerID = dr.Item("OWNER_ID")
                onCitation = IIf(dr.Item("CITATION") Is DBNull.Value, 0, dr.Item("CITATION"))
                odtCitationDueDate = IIf(dr.Item("CITATION_DUE_DATE") Is DBNull.Value, CDate("01/01/0001"), dr.Item("CITATION_DUE_DATE"))
                obolRescinded = IIf(dr.Item("RESCINDED") Is DBNull.Value, False, dr.Item("RESCINDED"))
                onOCEPath = IIf(dr.Item("OCE_PATH") Is DBNull.Value, 0, dr.Item("OCE_PATH"))
                odtOCEDate = IIf(dr.Item("OCE_DATE") Is DBNull.Value, CDate("01/01/0001"), dr.Item("OCE_DATE"))
                odtOCEProcessDate = IIf(dr.Item("OCE_PROCESS_DATE") Is DBNull.Value, CDate("01/01/0001"), dr.Item("OCE_PROCESS_DATE"))
                odtNextDueDate = IIf(dr.Item("NEXT_DUE_DATE") Is DBNull.Value, CDate("01/01/0001"), dr.Item("NEXT_DUE_DATE"))
                odtOverrideDueDate = IIf(dr.Item("OVERRIDE_DUE_DATE") Is DBNull.Value, CDate("01/01/0001"), dr.Item("OVERRIDE_DUE_DATE"))
                onOCEStatus = IIf(dr.Item("OCE_STATUS") Is DBNull.Value, 0, dr.Item("OCE_STATUS"))
                onEscalation = IIf(dr.Item("ESCALATION") Is DBNull.Value, 0, dr.Item("ESCALATION"))
                odPolicyAmount = IIf(dr.Item("POLICY_AMOUNT") Is DBNull.Value, -1.0, dr.Item("POLICY_AMOUNT"))
                odOverRideAmount = IIf(dr.Item("OVERRIDE_AMOUNT") Is DBNull.Value, -1.0, dr.Item("OVERRIDE_AMOUNT"))
                odSettlementAmount = IIf(dr.Item("SETTLEMENT_AMOUNT") Is DBNull.Value, -1.0, dr.Item("SETTLEMENT_AMOUNT"))
                odPaidAmount = IIf(dr.Item("PAID_AMOUNT") Is DBNull.Value, -1.0, dr.Item("PAID_AMOUNT"))
                odtDateReceived = IIf(dr.Item("DATE_RECEIVED") Is DBNull.Value, CDate("01/01/0001"), dr.Item("DATE_RECEIVED"))
                odtWorkShopDate = IIf(dr.Item("WORKSHOP_DATE") Is DBNull.Value, CDate("01/01/0001"), dr.Item("WORKSHOP_DATE"))
                onWorkShopResult = IIf(dr.Item("WORKSHOP_RESULT") Is DBNull.Value, 0, dr.Item("WORKSHOP_RESULT"))
                obolWorkshopRequired = IIf(dr.Item("WORKSHOP_REQUIRED") Is DBNull.Value, False, dr.Item("WORKSHOP_REQUIRED"))
                odtShowCauseDate = IIf(dr.Item("SHOW_CAUSE_DATE") Is DBNull.Value, CDate("01/01/0001"), dr.Item("SHOW_CAUSE_DATE"))
                onShowCauseResult = IIf(dr.Item("SHOW_CAUSE_RESULTS") Is DBNull.Value, 0, dr.Item("SHOW_CAUSE_RESULTS"))
                odtCommissionDate = IIf(dr.Item("COMMISSION_DATE") Is DBNull.Value, CDate("01/01/0001"), dr.Item("COMMISSION_DATE"))
                onCommissionResult = IIf(dr.Item("COMMISSION_RESULTS") Is DBNull.Value, 0, dr.Item("COMMISSION_RESULTS"))
                ostrAgreedOrder = IIf(dr.Item("AGREED_ORDER") Is DBNull.Value, String.Empty, dr.Item("AGREED_ORDER"))
                ostrAdministrativeOrder = IIf(dr.Item("ADMINISTRATIVE_ORDER") Is DBNull.Value, String.Empty, dr.Item("ADMINISTRATIVE_ORDER"))
                onPendingLetter = IIf(dr.Item("PENDING_LETTER") Is DBNull.Value, 0, dr.Item("PENDING_LETTER"))
                odtLetterGenerated = IIf(dr.Item("LETTER_GENERATED") Is DBNull.Value, CDate("01/01/0001"), dr.Item("LETTER_GENERATED"))
                odtRedTagDate = IIf(dr.Item("REDTAG_DATE") Is DBNull.Value, CDate("01/01/0001"), dr.Item("REDTAG_DATE"))
                obolLetterPrinted = IIf(dr.Item("LETTER_PRINTED") Is DBNull.Value, False, dr.Item("LETTER_PRINTED"))
                ostrCreatedBy = IIf(dr.Item("CREATED_BY") Is DBNull.Value, String.Empty, dr.Item("CREATED_BY"))
                odtCreatedOn = IIf(dr.Item("DATE_CREATED") Is DBNull.Value, CDate("01/01/0001"), dr.Item("DATE_CREATED"))
                ostrModifiedBy = IIf(dr.Item("LAST_EDITED_BY") Is DBNull.Value, String.Empty, dr.Item("LAST_EDITED_BY"))
                odtModifiedOn = IIf(dr.Item("DATE_LAST_EDITED") Is DBNull.Value, CDate("01/01/0001"), dr.Item("DATE_LAST_EDITED"))
                obolDeleted = dr.Item("DELETED")
                strEscalation = IIf(dr.Item("STRESCALATION") Is DBNull.Value, String.Empty, dr.Item("STRESCALATION"))
                onPendingLetterTemplateNum = IIf(dr.Item("PENDING_LETTER_TEMPLATE_NUM") Is DBNull.Value, 0, dr.Item("PENDING_LETTER_TEMPLATE_NUM"))
                onEnsiteID = IIf(dr.Item("ENSITE ID") Is DBNull.Value, 0, dr.Item("ENSITE ID"))
                odtAdminDate = IIf(dr.Item("ADMIN_HEARING_DATE") Is DBNull.Value, CDate("01/01/0001"), dr.Item("ADMIN_HEARING_DATE"))
                onAdminResult = IIf(dr.Item("ADMIN_HEARING_RESULT") Is DBNull.Value, 0, dr.Item("ADMIN_HEARING_RESULT"))

                If Not dr.Item("COMMENTS") Is Nothing Then
                    ostrComments = IIf(dr.Item("COMMENTS") Is DBNull.Value, String.Empty, dr.Item("COMMENTS"))
                End If


                Me.Reset()
            Catch ex As Exception
                MusterException.Publish(ex, Nothing, Nothing)
                Throw ex
            End Try
        End Sub
        '    Sub New(ByVal OCEID As Int32, _
        'ByVal OwnerID As Int32, _
        'ByVal OwnerName As String, _
        'ByVal Citation As Int32, _
        'ByVal CitationDueDate As Date, _
        'ByVal Rescinded As Boolean, _
        'ByVal OCEPath As Integer, _
        'ByVal OCEDate As Date, _
        'ByVal OCEProcessDate As Date, _
        'ByVal NextDueDate As Date, _
        'ByVal OverrideDueDate As Date, _
        'ByVal OCEStatus As Integer, _
        'ByVal status As String, _
        'ByVal Escalation As Integer, _
        'ByVal PolicyAmount As Decimal, _
        'ByVal OverRideAmount As Decimal, _
        'ByVal SettlementAmount As Decimal, _
        'ByVal PaidAmount As Decimal, _
        'ByVal DateReceived As Date, _
        'ByVal WorkShopDate As Date, _
        'ByVal WorkshopResult As Integer, _
        'ByVal WorkShopRequired As Boolean, _
        'ByVal WorkShopResultName As String, _
        'ByVal ShowCauseDate As Date, _
        'ByVal ShowCauseResults As Integer, _
        'ByVal ShowCauseResultName As String, _
        'ByVal CommissionDate As Date, _
        'ByVal CommissionResults As Integer, _
        'ByVal CommissionResultName As String, _
        'ByVal AgreedOrder As Integer, _
        'ByVal AdministrativeOrder As Integer, _
        'ByVal PendingLetter As Integer, _
        'ByVal PendingLetterName As String, _
        'ByVal LetterGenerated As Date, _
        'ByVal LetterPrinted As Boolean, _
        'ByVal CreatedBy As String, _
        'ByVal CreatedOn As Date, _
        'ByVal ModifiedBy As String, _
        'ByVal ModifiedOn As Date, _
        'ByVal Deleted As Boolean)


        '        onOCEID = OCEID
        '        onOwnerID = OwnerID
        '        onCitation = Citation
        '        odtCitationDueDate = CitationDueDate
        '        obolRescinded = Rescinded
        '        onOCEPath = OCEPath
        '        odtOCEDate = OCEDate
        '        odtOCEProcessDate = OCEProcessDate
        '        odtNextDueDate = NextDueDate
        '        odtOverrideDueDate = OverrideDueDate
        '        onOCEStatus = OCEStatus
        '        onEscalation = Escalation
        '        odPolicyAmount = PolicyAmount
        '        odOverRideAmount = OverRideAmount
        '        odSettlementAmount = Settlement
        '        odPaidAmount = PaidAmount
        '        odtDateReceived = DateReceived
        '        obolWorkshopRequired = bolWorkshopRequired
        '        odtWorkShopDate = WorkShopDate
        '        onWorkShopResult = WorkshopResult
        '        odtShowCauseDate = ShowCauseDate
        '        onShowCauseResult = ShowCauseResult
        '        odtCommissionDate = CommissionDate
        '        onCommissionResult = CommissionResult
        '        ostrAgreedOrder = AgreedOrder
        '        ostrAdministrativeOrder = AdministrativeOrder
        '        onPendingLetter = PendingLetter
        '        odtLetterGenerated = LetterGenerated
        '        obolLetterPrinted = LetterPrinted
        '        ostrCreatedBy = CreatedBy
        '        odtCreatedOn = CreatedOn
        '        ostrModifiedBy = ModifiedBy
        '        odtModifiedOn = ModifiedOn
        '        obolDeleted = Deleted

        '        ostrPendingLetter = PendingLetterName
        '        ostrShowCauseResult = ShowCauseResultName
        '        ostrCommissionResult = CommissionResultName
        '        ostrWorkshopResult = WorkShopResultName
        '        ostrOwner = OwnerName
        '        ostrStatus = status
        '        Me.Reset()
        '    End Sub
#End Region
#Region "Exposed Operations"
        Public Sub Reset()
            If nOCEID >= 0 Then
                nOCEID = onOCEID
            End If
            nOwnerID = onOwnerID
            nCitation = onCitation
            dtCitationDueDate = odtCitationDueDate
            bolRescinded = obolRescinded
            nOCEPath = onOCEPath
            dtOCEDate = odtOCEDate
            dtOCEProcessDate = odtOCEProcessDate
            dtNextDueDate = odtNextDueDate
            dtOverrideDueDate = odtOverrideDueDate
            nOCEStatus = onOCEStatus
            nEscalation = onEscalation
            dPolicyAmount = odPolicyAmount
            dOverRideAmount = odOverRideAmount
            dSettlementAmount = odSettlementAmount
            dPaidAmount = odPaidAmount
            dtDateReceived = odtDateReceived
            bolWorkshopRequired = obolWorkshopRequired
            dtWorkShopDate = odtWorkShopDate
            nWorkShopResult = onWorkShopResult
            dtShowCauseDate = odtShowCauseDate
            nShowCauseResult = onShowCauseResult
            dtCommissionDate = odtCommissionDate
            nCommissionResult = onCommissionResult
            strAgreedOrder = ostrAgreedOrder
            strAdministrativeOrder = ostrAdministrativeOrder
            nPendingLetter = onPendingLetter
            dtLetterGenerated = odtLetterGenerated
            dtRedTagDate = odtRedTagDate
            bolLetterPrinted = obolLetterPrinted
            strCreatedBy = ostrCreatedBy
            dtCreatedOn = odtCreatedOn
            strModifiedBy = ostrModifiedBy
            dtModifiedOn = odtModifiedOn
            bolDeleted = obolDeleted
            nPendingLetterTemplateNum = onPendingLetterTemplateNum
            nEnsiteID = onEnsiteID
            strComments = ostrComments
            dtAdminDate = odtAdminDate
            nAdminResult = onAdminResult

            bolIsDirty = False
            RaiseEvent OwnerComplianceEventInfoChanged(bolIsDirty)
        End Sub
        Public Sub Archive()
            onOCEID = nOCEID
            onOwnerID = nOwnerID
            onCitation = nCitation
            odtCitationDueDate = dtCitationDueDate
            obolRescinded = bolRescinded
            onOCEPath = nOCEPath
            odtOCEDate = dtOCEDate
            odtOCEProcessDate = dtOCEProcessDate
            odtNextDueDate = dtNextDueDate
            odtOverrideDueDate = dtOverrideDueDate
            onOCEStatus = nOCEStatus
            onEscalation = nEscalation
            odPolicyAmount = dPolicyAmount
            odOverRideAmount = dOverRideAmount
            odSettlementAmount = dSettlementAmount
            odPaidAmount = dPaidAmount
            odtDateReceived = dtDateReceived
            obolWorkshopRequired = bolWorkshopRequired
            odtWorkShopDate = dtWorkShopDate
            onWorkShopResult = nWorkShopResult
            odtShowCauseDate = dtShowCauseDate
            onShowCauseResult = nShowCauseResult
            odtCommissionDate = dtCommissionDate
            onCommissionResult = nCommissionResult
            ostrAgreedOrder = strAgreedOrder
            ostrAdministrativeOrder = strAdministrativeOrder
            onPendingLetter = nPendingLetter
            odtLetterGenerated = dtLetterGenerated
            odtRedTagDate = dtRedTagDate
            obolLetterPrinted = bolLetterPrinted
            ostrCreatedBy = strCreatedBy
            odtCreatedOn = dtCreatedOn
            ostrModifiedBy = strModifiedBy
            odtModifiedOn = dtModifiedOn
            obolDeleted = bolDeleted
            onPendingLetterTemplateNum = nPendingLetterTemplateNum
            onEnsiteID = nEnsiteID
            ostrComments = strComments
            odtAdminDate = dtAdminDate
            onAdminResult = nAdminResult

            bolIsDirty = False
        End Sub
#End Region
#Region "Private Operations"
        Private Sub CheckDirty()
            Dim obolIsDirty As Boolean = bolIsDirty

            bolIsDirty = (nOwnerID <> onOwnerID) Or _
                        (nCitation <> onCitation) Or _
                        (dtCitationDueDate <> odtCitationDueDate) Or _
                        (bolRescinded <> obolRescinded) Or _
                        (nOCEPath <> onOCEPath) Or _
                        (dtOCEDate <> odtOCEDate) Or _
                        (dtOCEProcessDate <> odtOCEProcessDate) Or _
                        (dtNextDueDate <> odtNextDueDate) Or _
                        (dtOverrideDueDate <> odtOverrideDueDate) Or _
                        (nOCEStatus <> onOCEStatus) Or _
                        (nEscalation <> onEscalation) Or _
                        (dPolicyAmount <> odPolicyAmount) Or _
                        (odOverRideAmount <> dOverRideAmount) Or _
                        (dSettlementAmount <> odSettlementAmount) Or _
                        (dPaidAmount <> odPaidAmount) Or _
                        (dtDateReceived <> odtDateReceived) Or _
                        (dtWorkShopDate <> odtWorkShopDate) Or _
                        (nWorkShopResult <> onWorkShopResult) Or _
                        (bolWorkshopRequired <> obolWorkshopRequired) Or _
                        (dtShowCauseDate <> odtShowCauseDate) Or _
                        (nShowCauseResult <> onShowCauseResult) Or _
                        (dtCommissionDate <> odtCommissionDate) Or _
                        (nCommissionResult <> onCommissionResult) Or _
                        (strAgreedOrder <> ostrAgreedOrder) Or _
                        (strAdministrativeOrder <> ostrAdministrativeOrder) Or _
                        (nPendingLetter <> onPendingLetter) Or _
                        (dtLetterGenerated <> odtLetterGenerated) Or _
                        (dtRedTagDate <> odtRedTagDate) Or _
                        (bolLetterPrinted <> obolLetterPrinted) Or _
                        (dtModifiedOn <> odtModifiedOn) Or _
                        (strCreatedBy <> ostrCreatedBy) Or _
                        (dtCreatedOn <> odtCreatedOn) Or _
                        (strModifiedBy <> ostrModifiedBy) Or _
                        (bolDeleted <> obolDeleted) Or _
                        (nPendingLetterTemplateNum <> onPendingLetterTemplateNum) Or _
                        (nEnsiteID <> onEnsiteID) Or _
                        (strComments <> ostrComments) Or _
                        (dtAdminDate <> odtAdminDate) Or _
                        (nAdminResult <> onAdminResult)


            If obolIsDirty <> bolIsDirty Then
                RaiseEvent OwnerComplianceEventInfoChanged(bolIsDirty)
            End If
        End Sub
        Private Sub Init()
            onOCEID = 0
            onOwnerID = 0
            onCitation = 0
            odtCitationDueDate = CDate("01/01/0001")
            obolRescinded = False
            onOCEPath = 0
            odtOCEDate = CDate("01/01/0001")
            odtOCEProcessDate = CDate("01/01/0001")
            odtNextDueDate = CDate("01/01/0001")
            odtOverrideDueDate = CDate("01/01/0001")
            onOCEStatus = 0
            onEscalation = 0
            odPolicyAmount = -1.0
            odSettlementAmount = -1.0
            odPaidAmount = -1.0
            odOverRideAmount = -1.0
            odtDateReceived = CDate("01/01/0001")
            odtWorkShopDate = CDate("01/01/0001")
            onWorkShopResult = 0
            odtAdminDate = CDate("01/01/0001")
            onAdminResult = 0
            odtShowCauseDate = CDate("01/01/0001")
            onShowCauseResult = 0
            odtCommissionDate = CDate("01/01/0001")
            onCommissionResult = 0
            ostrAgreedOrder = String.Empty
            ostrAdministrativeOrder = String.Empty
            onPendingLetter = 0
            odtLetterGenerated = CDate("01/01/0001")
            odtRedTagDate = CDate("01/01/0001")
            obolLetterPrinted = False
            obolDeleted = False
            ostrCreatedBy = String.Empty
            odtCreatedOn = CDate("01/01/0001")
            ostrModifiedBy = String.Empty
            odtModifiedOn = CDate("01/01/0001")
            strEscalation = String.Empty
            onPendingLetterTemplateNum = 0
            onEnsiteID = 0
            ostrComments = String.Empty
            Me.Reset()
        End Sub
#End Region
#Region "Exposed Attributes"
        Public Property ID() As Int32
            Get
                Return nOCEID
            End Get
            Set(ByVal Value As Int32)
                nOCEID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property OwnerID() As Int32
            Get
                Return nOwnerID
            End Get
            Set(ByVal Value As Int32)
                nOwnerID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Citation() As Integer
            Get
                Return nCitation
            End Get
            Set(ByVal Value As Integer)
                nCitation = Value
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
        Public Property Rescinded() As Boolean
            Get
                Return bolRescinded
            End Get
            Set(ByVal Value As Boolean)
                bolRescinded = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property OCEPath() As Integer
            Get
                Return nOCEPath
            End Get
            Set(ByVal Value As Integer)
                nOCEPath = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property OCEDate() As Date
            Get
                Return dtOCEDate
            End Get
            Set(ByVal Value As Date)
                dtOCEDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property OCEProcessDate() As Date
            Get
                Return dtOCEProcessDate
            End Get
            Set(ByVal Value As Date)
                dtOCEProcessDate = Value
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
        Public Property OCEStatus() As Integer
            Get
                Return nOCEStatus
            End Get
            Set(ByVal Value As Integer)
                nOCEStatus = Value
                Me.CheckDirty()
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
        Public Property OverRideAmount() As Decimal
            Get
                Return dOverRideAmount
            End Get
            Set(ByVal Value As Decimal)
                dOverRideAmount = Value
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
        Public Property WorkShopResult() As Integer
            Get
                Return nWorkShopResult
            End Get
            Set(ByVal Value As Integer)
                nWorkShopResult = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property WorkshopRequired() As Boolean
            Get
                Return bolWorkshopRequired
            End Get
            Set(ByVal Value As Boolean)
                bolWorkshopRequired = Value
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
        Public Property ShowCauseResult() As Integer
            Get
                Return nShowCauseResult
            End Get
            Set(ByVal Value As Integer)
                nShowCauseResult = Value
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
        Public Property CommissionResult() As Integer
            Get
                Return nCommissionResult
            End Get
            Set(ByVal Value As Integer)
                nCommissionResult = Value
                Me.CheckDirty()
            End Set
        End Property

        Public Property AdminHearingDate() As Date
            Get
                Return dtAdminDate
            End Get
            Set(ByVal Value As Date)
                dtAdminDate = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property AdminHearingResult() As Integer
            Get
                Return nAdminResult
            End Get
            Set(ByVal Value As Integer)
                nAdminResult = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property AgreedOrder() As String
            Get
                Return strAgreedOrder
            End Get
            Set(ByVal Value As String)
                strAgreedOrder = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property AdministrativeOrder() As String
            Get
                Return strAdministrativeOrder
            End Get
            Set(ByVal Value As String)
                strAdministrativeOrder = Value
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
        Public Property RedTagDate() As Date
            Get
                Return dtRedTagDate
            End Get
            Set(ByVal Value As Date)
                dtRedTagDate = Value
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
        Public Property EscalationString() As String
            Get
                Return strEscalation
            End Get
            Set(ByVal Value As String)
                strEscalation = Value
                ' no need to check for isdirty as this variable is not included in the condition
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
        Public Property EnsiteID() As Integer
            Get
                Return nEnsiteID
            End Get
            Set(ByVal Value As Integer)
                nEnsiteID = Value
                Me.CheckDirty()
            End Set
        End Property
        Public Property Comments() As String
            Get
                Return strComments
            End Get

            Set(ByVal Value As String)
                strComments = Value
                Me.CheckDirty()
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
