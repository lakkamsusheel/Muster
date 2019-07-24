Imports Infragistics.Win.UltraWinGrid
Imports System.Collections.Generic

'   Changes
'   05/21/2009    Thomas Franey          Added Comments to Enforcement
'   12/20/2010    Hua Cao                Added Red Tag Date to Enforcement
'                                         Hide Score and Potential Inspector from Inspection grid as well as clean SQL code
Public Class CandEManagement
    Inherits System.Windows.Forms.Form


#Region "Events"
    Friend WithEvents pnlAssignedInspectionsBottom As System.Windows.Forms.Panel
    Public Event evtWorkShopDate()
#End Region

#Region "Private Member Variables"
    Private FixingStr As Boolean = False
    Private pFCE As New MUSTER.BusinessLogic.pFacilityComplianceEvent
    Private pOCE As New MUSTER.BusinessLogic.pOwnerComplianceEvent
    Private pInspection As New MUSTER.BusinessLogic.pInspection
    Private pInsCitation As New MUSTER.BusinessLogic.pInspectionCitation
    Private pInsDiscrep As New MUSTER.BusinessLogic.pInspectionDiscrep
    Private WithEvents pLCE As New MUSTER.BusinessLogic.pLicenseeComplianceEvent
    Private WithEvents objLCE As LicenseeComplianceEvent
    Friend WithEvents objEscalationDates As UserDate
    Private frmWorkshopDate As WorkShopDate
    Private EnforcementToolTip As ToolTip
    Private WithEvents frmEnforcementHistory As EnforcementHistory
    Private dtWorkshopDate As Date = CDate("01/01/0001")
    Private OCECreationError As Boolean = False
    Private lastUGRowOwner As Integer
    Private OCECancelEscalation As Boolean = False
    Private bolLoading As Boolean
    Private intWorkingFacility As Integer = 0
    Private intWorkingOwnerID As Integer = 0
    Private arrayWorkingvalues(1) As Integer
    Private bolCitationModified As Boolean = False
    Private bolOCEModified As Boolean = False
    Private bolOCEDueDateModified As Boolean = False
    Private bolCitationDateModified As Boolean = False
    Private bolDiscrepModified As Boolean = False
    Private bolDiscrepDateModified As Boolean = False
    Private strOCEGeneratedLetterName As String = String.Empty

    Private vListWorkShopResult, vListShowCauseResult, vListCommissionResult, vListAdminHearingResult As Infragistics.Win.ValueList
    Private vListShowCauseHearingResult, vListCommissionHearingResult As Infragistics.Win.ValueList

    Dim rp As New Remove_Pencil

    Public MyGuid As New System.Guid

    Private ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
    Private ugChildRow As Infragistics.Win.UltraWinGrid.UltraGridRow
    Private ugGrandChildRow As Infragistics.Win.UltraWinGrid.UltraGridRow
    Private ugGreatGrandChildRow As Infragistics.Win.UltraWinGrid.UltraGridRow
    Private returnVal As String = String.Empty

    Private Enum OCELetterType
        ShowCauseHearing = 1006
        ShowCauseAgreedOrder = 1007
        CommissionHearing = 1008
        CommissionHearingNFARescinded = 1009
        AdministrativeOrder = 1010
        NFALetter = 1119
        Discrepancy = 1172
        NOV = 1173
        NOVWorkshop = 1174
        NOVAgreedOrder = 1175
        NOVWorkshopAgreedOrder = 1176
        SecondNotice = 1249
        AgreedOrder = 1250
        NFARescind = 1251
    End Enum
    'Public WithEvents pIns As New MUSTER.BusinessLogic.pInspection
    'Dim ugcitationrow As Infragistics.Win.UltraWinGrid.UltraGridRow
    'Dim ugcitationband As Infragistics.Win.UltraWinGrid.UltraGridChildBand
    'Dim objWorkshopDate As New WorkShopDate
#End Region

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()
        bolLoading = True
        MyGuid = System.Guid.NewGuid
        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        MusterContainer.AppUser.LogEntry("CandEManagement", MyGuid.ToString)
        MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "CandEManagement")
        bolLoading = False
    End Sub


    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)

        If Not EnforcementToolTip Is Nothing Then
            EnforcementToolTip.Dispose()
        End If

        MusterContainer.AppSemaphores.Remove(MyGuid.ToString)
        MusterContainer.AppUser.LogExit(MyGuid.ToString)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents tabCntrlCandE As System.Windows.Forms.TabControl
    Friend WithEvents tbPageInspections As System.Windows.Forms.TabPage
    Friend WithEvents pnlInspectionsContainer As System.Windows.Forms.Panel
    Friend WithEvents ugInspections As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlInspectionsBottom As System.Windows.Forms.Panel
    Friend WithEvents btnInsViewEditCheckList As System.Windows.Forms.Button
    Friend WithEvents tbPageCompliance As System.Windows.Forms.TabPage
    Friend WithEvents tbPageEnforcement As System.Windows.Forms.TabPage
    Friend WithEvents tbPageAssignedInspections As System.Windows.Forms.TabPage
    Friend WithEvents tbPageLicensees As System.Windows.Forms.TabPage
    Friend WithEvents lblRenewalLicensees As System.Windows.Forms.Label
    Friend WithEvents pnlInspectionsTop As System.Windows.Forms.Panel
    Friend WithEvents btnInsGenerateFCEs As System.Windows.Forms.Button
    Friend WithEvents pnlComplianceTop As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents pnlComplianceBottom As System.Windows.Forms.Panel
    Friend WithEvents btnCompGenerateOCEs As System.Windows.Forms.Button
    Friend WithEvents pnlComplianceContainer As System.Windows.Forms.Panel
    Friend WithEvents ugCompliance As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnCompAddFCE As System.Windows.Forms.Button
    Friend WithEvents grpCompAdminFCEFunctions As System.Windows.Forms.GroupBox
    Friend WithEvents btnCompEditFCE As System.Windows.Forms.Button
    Friend WithEvents btnCompDeleteFCE As System.Windows.Forms.Button
    Friend WithEvents pnlEnforcementContainer As System.Windows.Forms.Panel
    Friend WithEvents ugEnforcement As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlEnforcementBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlEnforcementTop As System.Windows.Forms.Panel
    Friend WithEvents lblEnforceOCEs As System.Windows.Forms.Label
    Friend WithEvents btnEnforceProcessEscalations As System.Windows.Forms.Button
    Friend WithEvents btnEnforceViewEnforceHistory As System.Windows.Forms.Button
    Friend WithEvents btnEnforceRefresh As System.Windows.Forms.Button
    Friend WithEvents btnEnforceProcessRescissions As System.Windows.Forms.Button
    Friend WithEvents btnEnforceGenerateLetter As System.Windows.Forms.Button
    Friend WithEvents btnAssignedInspectionsAccept As System.Windows.Forms.Button
    Friend WithEvents btnAssignedInspectionsAdd As System.Windows.Forms.Button
    Friend WithEvents pnlLicenseesTop As System.Windows.Forms.Panel
    Friend WithEvents lblLicenseesComplianceEvents As System.Windows.Forms.Label
    Friend WithEvents pnlLicenseesBottom As System.Windows.Forms.Panel
    Friend WithEvents btnLicenseeDeleteLCE As System.Windows.Forms.Button
    Friend WithEvents btnLicenseeEditLCE As System.Windows.Forms.Button
    Friend WithEvents btnLicenseeAddLCE As System.Windows.Forms.Button
    Friend WithEvents btnLicenseeGenerateLetter As System.Windows.Forms.Button
    Friend WithEvents btnLicenseeProcessRescissions As System.Windows.Forms.Button
    Friend WithEvents btnLicenseeRefresh As System.Windows.Forms.Button
    Friend WithEvents btnLicenseeViewEnforcementHistory As System.Windows.Forms.Button
    Friend WithEvents btnLicenseeProcessEscalations As System.Windows.Forms.Button
    Friend WithEvents pnlLicenseesContainer As System.Windows.Forms.Panel
    Friend WithEvents ugLicensees As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlAssignedInspectionsTop As System.Windows.Forms.Panel
    Friend WithEvents lblAssignedInspections As System.Windows.Forms.Label
    Friend WithEvents pnlAssignedInspectionsContainer As System.Windows.Forms.Panel
    Friend WithEvents ugAssignedInspections As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnEditAssignedInspection As System.Windows.Forms.Button
    Friend WithEvents btnAssignedInspExpandCollapseAll As System.Windows.Forms.Button
    Friend WithEvents btnComplianceExpandCollapseAll As System.Windows.Forms.Button
    Friend WithEvents btnLicenseesExpandCollapseAll As System.Windows.Forms.Button
    Friend WithEvents btnEnforcementExpandCollapseAll As System.Windows.Forms.Button
    Friend WithEvents btnInspectionsExpandCollapseAll As System.Windows.Forms.Button
    Friend WithEvents btnInspectionsRefresh As System.Windows.Forms.Button
    Friend WithEvents btnComplianceRefresh As System.Windows.Forms.Button
    Friend WithEvents btnAssignedRefresh As System.Windows.Forms.Button
    Friend WithEvents btnAssignedInspectionsDelete As System.Windows.Forms.Button
    Friend WithEvents btnInsViewEditCheckList2 As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents cmbManager As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents brnAgreedOrder As System.Windows.Forms.Button
    Friend WithEvents btnManualEsc As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance5 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance6 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance7 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance8 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance9 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance10 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.tabCntrlCandE = New System.Windows.Forms.TabControl
        Me.tbPageInspections = New System.Windows.Forms.TabPage
        Me.pnlInspectionsContainer = New System.Windows.Forms.Panel
        Me.ugInspections = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlInspectionsBottom = New System.Windows.Forms.Panel
        Me.btnInsGenerateFCEs = New System.Windows.Forms.Button
        Me.btnInsViewEditCheckList = New System.Windows.Forms.Button
        Me.btnInspectionsRefresh = New System.Windows.Forms.Button
        Me.pnlInspectionsTop = New System.Windows.Forms.Panel
        Me.btnInspectionsExpandCollapseAll = New System.Windows.Forms.Button
        Me.lblRenewalLicensees = New System.Windows.Forms.Label
        Me.tbPageCompliance = New System.Windows.Forms.TabPage
        Me.pnlComplianceContainer = New System.Windows.Forms.Panel
        Me.ugCompliance = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlComplianceBottom = New System.Windows.Forms.Panel
        Me.grpCompAdminFCEFunctions = New System.Windows.Forms.GroupBox
        Me.btnCompDeleteFCE = New System.Windows.Forms.Button
        Me.btnCompEditFCE = New System.Windows.Forms.Button
        Me.btnCompAddFCE = New System.Windows.Forms.Button
        Me.btnCompGenerateOCEs = New System.Windows.Forms.Button
        Me.btnComplianceRefresh = New System.Windows.Forms.Button
        Me.pnlComplianceTop = New System.Windows.Forms.Panel
        Me.btnComplianceExpandCollapseAll = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.tbPageEnforcement = New System.Windows.Forms.TabPage
        Me.pnlEnforcementContainer = New System.Windows.Forms.Panel
        Me.ugEnforcement = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlEnforcementBottom = New System.Windows.Forms.Panel
        Me.brnAgreedOrder = New System.Windows.Forms.Button
        Me.btnInsViewEditCheckList2 = New System.Windows.Forms.Button
        Me.btnEnforceGenerateLetter = New System.Windows.Forms.Button
        Me.btnEnforceProcessRescissions = New System.Windows.Forms.Button
        Me.btnEnforceRefresh = New System.Windows.Forms.Button
        Me.btnEnforceViewEnforceHistory = New System.Windows.Forms.Button
        Me.btnEnforceProcessEscalations = New System.Windows.Forms.Button
        Me.pnlEnforcementTop = New System.Windows.Forms.Panel
        Me.btnEnforcementExpandCollapseAll = New System.Windows.Forms.Button
        Me.lblEnforceOCEs = New System.Windows.Forms.Label
        Me.tbPageAssignedInspections = New System.Windows.Forms.TabPage
        Me.pnlAssignedInspectionsContainer = New System.Windows.Forms.Panel
        Me.ugAssignedInspections = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlAssignedInspectionsTop = New System.Windows.Forms.Panel
        Me.btnAssignedInspExpandCollapseAll = New System.Windows.Forms.Button
        Me.lblAssignedInspections = New System.Windows.Forms.Label
        Me.pnlAssignedInspectionsBottom = New System.Windows.Forms.Panel
        Me.btnEditAssignedInspection = New System.Windows.Forms.Button
        Me.btnAssignedInspectionsAccept = New System.Windows.Forms.Button
        Me.btnAssignedInspectionsAdd = New System.Windows.Forms.Button
        Me.btnAssignedRefresh = New System.Windows.Forms.Button
        Me.btnAssignedInspectionsDelete = New System.Windows.Forms.Button
        Me.tbPageLicensees = New System.Windows.Forms.TabPage
        Me.pnlLicenseesContainer = New System.Windows.Forms.Panel
        Me.ugLicensees = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlLicenseesBottom = New System.Windows.Forms.Panel
        Me.btnLicenseeGenerateLetter = New System.Windows.Forms.Button
        Me.btnLicenseeProcessRescissions = New System.Windows.Forms.Button
        Me.btnLicenseeRefresh = New System.Windows.Forms.Button
        Me.btnLicenseeViewEnforcementHistory = New System.Windows.Forms.Button
        Me.btnLicenseeProcessEscalations = New System.Windows.Forms.Button
        Me.btnLicenseeDeleteLCE = New System.Windows.Forms.Button
        Me.btnLicenseeEditLCE = New System.Windows.Forms.Button
        Me.btnLicenseeAddLCE = New System.Windows.Forms.Button
        Me.pnlLicenseesTop = New System.Windows.Forms.Panel
        Me.btnLicenseesExpandCollapseAll = New System.Windows.Forms.Button
        Me.lblLicenseesComplianceEvents = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Label2 = New System.Windows.Forms.Label
        Me.cmbManager = New System.Windows.Forms.ComboBox
        Me.btnManualEsc = New System.Windows.Forms.Button
        Me.tabCntrlCandE.SuspendLayout()
        Me.tbPageInspections.SuspendLayout()
        Me.pnlInspectionsContainer.SuspendLayout()
        CType(Me.ugInspections, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlInspectionsBottom.SuspendLayout()
        Me.pnlInspectionsTop.SuspendLayout()
        Me.tbPageCompliance.SuspendLayout()
        Me.pnlComplianceContainer.SuspendLayout()
        CType(Me.ugCompliance, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlComplianceBottom.SuspendLayout()
        Me.grpCompAdminFCEFunctions.SuspendLayout()
        Me.pnlComplianceTop.SuspendLayout()
        Me.tbPageEnforcement.SuspendLayout()
        Me.pnlEnforcementContainer.SuspendLayout()
        CType(Me.ugEnforcement, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlEnforcementBottom.SuspendLayout()
        Me.pnlEnforcementTop.SuspendLayout()
        Me.tbPageAssignedInspections.SuspendLayout()
        Me.pnlAssignedInspectionsContainer.SuspendLayout()
        CType(Me.ugAssignedInspections, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlAssignedInspectionsTop.SuspendLayout()
        Me.pnlAssignedInspectionsBottom.SuspendLayout()
        Me.tbPageLicensees.SuspendLayout()
        Me.pnlLicenseesContainer.SuspendLayout()
        CType(Me.ugLicensees, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlLicenseesBottom.SuspendLayout()
        Me.pnlLicenseesTop.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'tabCntrlCandE
        '
        Me.tabCntrlCandE.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.tabCntrlCandE.Controls.Add(Me.tbPageInspections)
        Me.tabCntrlCandE.Controls.Add(Me.tbPageCompliance)
        Me.tabCntrlCandE.Controls.Add(Me.tbPageEnforcement)
        Me.tabCntrlCandE.Controls.Add(Me.tbPageAssignedInspections)
        Me.tabCntrlCandE.Controls.Add(Me.tbPageLicensees)
        Me.tabCntrlCandE.Location = New System.Drawing.Point(0, 25)
        Me.tabCntrlCandE.Name = "tabCntrlCandE"
        Me.tabCntrlCandE.SelectedIndex = 0
        Me.tabCntrlCandE.Size = New System.Drawing.Size(928, 488)
        Me.tabCntrlCandE.TabIndex = 3
        '
        'tbPageInspections
        '
        Me.tbPageInspections.Controls.Add(Me.pnlInspectionsContainer)
        Me.tbPageInspections.Controls.Add(Me.pnlInspectionsBottom)
        Me.tbPageInspections.Controls.Add(Me.pnlInspectionsTop)
        Me.tbPageInspections.Location = New System.Drawing.Point(4, 22)
        Me.tbPageInspections.Name = "tbPageInspections"
        Me.tbPageInspections.Size = New System.Drawing.Size(920, 462)
        Me.tbPageInspections.TabIndex = 0
        Me.tbPageInspections.Text = "Inspections"
        '
        'pnlInspectionsContainer
        '
        Me.pnlInspectionsContainer.Controls.Add(Me.ugInspections)
        Me.pnlInspectionsContainer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlInspectionsContainer.Location = New System.Drawing.Point(0, 32)
        Me.pnlInspectionsContainer.Name = "pnlInspectionsContainer"
        Me.pnlInspectionsContainer.Size = New System.Drawing.Size(920, 382)
        Me.pnlInspectionsContainer.TabIndex = 1
        '
        'ugInspections
        '
        Me.ugInspections.Cursor = System.Windows.Forms.Cursors.Default
        Appearance1.BorderColor = System.Drawing.SystemColors.ControlLight
        Me.ugInspections.DisplayLayout.Override.CellAppearance = Appearance1
        Me.ugInspections.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Appearance2.BorderColor = System.Drawing.SystemColors.ControlLight
        Me.ugInspections.DisplayLayout.Override.RowAppearance = Appearance2
        Me.ugInspections.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugInspections.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugInspections.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugInspections.Location = New System.Drawing.Point(0, 0)
        Me.ugInspections.Name = "ugInspections"
        Me.ugInspections.Size = New System.Drawing.Size(920, 382)
        Me.ugInspections.TabIndex = 0
        '
        'pnlInspectionsBottom
        '
        Me.pnlInspectionsBottom.Controls.Add(Me.btnInsGenerateFCEs)
        Me.pnlInspectionsBottom.Controls.Add(Me.btnInsViewEditCheckList)
        Me.pnlInspectionsBottom.Controls.Add(Me.btnInspectionsRefresh)
        Me.pnlInspectionsBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlInspectionsBottom.Location = New System.Drawing.Point(0, 414)
        Me.pnlInspectionsBottom.Name = "pnlInspectionsBottom"
        Me.pnlInspectionsBottom.Size = New System.Drawing.Size(920, 48)
        Me.pnlInspectionsBottom.TabIndex = 2
        '
        'btnInsGenerateFCEs
        '
        Me.btnInsGenerateFCEs.Location = New System.Drawing.Point(176, 16)
        Me.btnInsGenerateFCEs.Name = "btnInsGenerateFCEs"
        Me.btnInsGenerateFCEs.Size = New System.Drawing.Size(132, 24)
        Me.btnInsGenerateFCEs.TabIndex = 1
        Me.btnInsGenerateFCEs.Text = "Generate FCE(s)"
        '
        'btnInsViewEditCheckList
        '
        Me.btnInsViewEditCheckList.Location = New System.Drawing.Point(40, 16)
        Me.btnInsViewEditCheckList.Name = "btnInsViewEditCheckList"
        Me.btnInsViewEditCheckList.Size = New System.Drawing.Size(132, 24)
        Me.btnInsViewEditCheckList.TabIndex = 0
        Me.btnInsViewEditCheckList.Text = "View/Edit CheckList"
        '
        'btnInspectionsRefresh
        '
        Me.btnInspectionsRefresh.Location = New System.Drawing.Point(328, 16)
        Me.btnInspectionsRefresh.Name = "btnInspectionsRefresh"
        Me.btnInspectionsRefresh.Size = New System.Drawing.Size(132, 24)
        Me.btnInspectionsRefresh.TabIndex = 1
        Me.btnInspectionsRefresh.Text = "Refresh"
        '
        'pnlInspectionsTop
        '
        Me.pnlInspectionsTop.Controls.Add(Me.btnInspectionsExpandCollapseAll)
        Me.pnlInspectionsTop.Controls.Add(Me.lblRenewalLicensees)
        Me.pnlInspectionsTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlInspectionsTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlInspectionsTop.Name = "pnlInspectionsTop"
        Me.pnlInspectionsTop.Size = New System.Drawing.Size(920, 32)
        Me.pnlInspectionsTop.TabIndex = 0
        '
        'btnInspectionsExpandCollapseAll
        '
        Me.btnInspectionsExpandCollapseAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnInspectionsExpandCollapseAll.BackColor = System.Drawing.SystemColors.Control
        Me.btnInspectionsExpandCollapseAll.Location = New System.Drawing.Point(824, 5)
        Me.btnInspectionsExpandCollapseAll.Name = "btnInspectionsExpandCollapseAll"
        Me.btnInspectionsExpandCollapseAll.Size = New System.Drawing.Size(88, 23)
        Me.btnInspectionsExpandCollapseAll.TabIndex = 5
        Me.btnInspectionsExpandCollapseAll.Text = "Expand All"
        '
        'lblRenewalLicensees
        '
        Me.lblRenewalLicensees.Location = New System.Drawing.Point(8, 16)
        Me.lblRenewalLicensees.Name = "lblRenewalLicensees"
        Me.lblRenewalLicensees.Size = New System.Drawing.Size(64, 16)
        Me.lblRenewalLicensees.TabIndex = 1
        Me.lblRenewalLicensees.Text = "Inspections"
        '
        'tbPageCompliance
        '
        Me.tbPageCompliance.Controls.Add(Me.pnlComplianceContainer)
        Me.tbPageCompliance.Controls.Add(Me.pnlComplianceBottom)
        Me.tbPageCompliance.Controls.Add(Me.pnlComplianceTop)
        Me.tbPageCompliance.Location = New System.Drawing.Point(4, 22)
        Me.tbPageCompliance.Name = "tbPageCompliance"
        Me.tbPageCompliance.Size = New System.Drawing.Size(920, 462)
        Me.tbPageCompliance.TabIndex = 1
        Me.tbPageCompliance.Text = "Compliance"
        Me.tbPageCompliance.Visible = False
        '
        'pnlComplianceContainer
        '
        Me.pnlComplianceContainer.Controls.Add(Me.ugCompliance)
        Me.pnlComplianceContainer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlComplianceContainer.Location = New System.Drawing.Point(0, 32)
        Me.pnlComplianceContainer.Name = "pnlComplianceContainer"
        Me.pnlComplianceContainer.Size = New System.Drawing.Size(920, 358)
        Me.pnlComplianceContainer.TabIndex = 1
        '
        'ugCompliance
        '
        Me.ugCompliance.Cursor = System.Windows.Forms.Cursors.Default
        Appearance3.BorderColor = System.Drawing.SystemColors.ControlLight
        Me.ugCompliance.DisplayLayout.Override.CellAppearance = Appearance3
        Me.ugCompliance.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Appearance4.BorderColor = System.Drawing.SystemColors.ControlLight
        Me.ugCompliance.DisplayLayout.Override.RowAppearance = Appearance4
        Me.ugCompliance.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugCompliance.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugCompliance.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugCompliance.Location = New System.Drawing.Point(0, 0)
        Me.ugCompliance.Name = "ugCompliance"
        Me.ugCompliance.Size = New System.Drawing.Size(920, 358)
        Me.ugCompliance.TabIndex = 0
        '
        'pnlComplianceBottom
        '
        Me.pnlComplianceBottom.Controls.Add(Me.grpCompAdminFCEFunctions)
        Me.pnlComplianceBottom.Controls.Add(Me.btnCompGenerateOCEs)
        Me.pnlComplianceBottom.Controls.Add(Me.btnComplianceRefresh)
        Me.pnlComplianceBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlComplianceBottom.Location = New System.Drawing.Point(0, 390)
        Me.pnlComplianceBottom.Name = "pnlComplianceBottom"
        Me.pnlComplianceBottom.Size = New System.Drawing.Size(920, 72)
        Me.pnlComplianceBottom.TabIndex = 2
        '
        'grpCompAdminFCEFunctions
        '
        Me.grpCompAdminFCEFunctions.Controls.Add(Me.btnCompDeleteFCE)
        Me.grpCompAdminFCEFunctions.Controls.Add(Me.btnCompEditFCE)
        Me.grpCompAdminFCEFunctions.Controls.Add(Me.btnCompAddFCE)
        Me.grpCompAdminFCEFunctions.Location = New System.Drawing.Point(168, 8)
        Me.grpCompAdminFCEFunctions.Name = "grpCompAdminFCEFunctions"
        Me.grpCompAdminFCEFunctions.Size = New System.Drawing.Size(344, 56)
        Me.grpCompAdminFCEFunctions.TabIndex = 1
        Me.grpCompAdminFCEFunctions.TabStop = False
        Me.grpCompAdminFCEFunctions.Text = "Administrative FCE Functions"
        '
        'btnCompDeleteFCE
        '
        Me.btnCompDeleteFCE.Location = New System.Drawing.Point(208, 24)
        Me.btnCompDeleteFCE.Name = "btnCompDeleteFCE"
        Me.btnCompDeleteFCE.Size = New System.Drawing.Size(86, 24)
        Me.btnCompDeleteFCE.TabIndex = 2
        Me.btnCompDeleteFCE.Text = "Delete FCE"
        '
        'btnCompEditFCE
        '
        Me.btnCompEditFCE.Location = New System.Drawing.Point(112, 24)
        Me.btnCompEditFCE.Name = "btnCompEditFCE"
        Me.btnCompEditFCE.Size = New System.Drawing.Size(86, 24)
        Me.btnCompEditFCE.TabIndex = 1
        Me.btnCompEditFCE.Text = "Edit FCE"
        '
        'btnCompAddFCE
        '
        Me.btnCompAddFCE.Location = New System.Drawing.Point(16, 24)
        Me.btnCompAddFCE.Name = "btnCompAddFCE"
        Me.btnCompAddFCE.Size = New System.Drawing.Size(86, 24)
        Me.btnCompAddFCE.TabIndex = 0
        Me.btnCompAddFCE.Text = "Add FCE"
        '
        'btnCompGenerateOCEs
        '
        Me.btnCompGenerateOCEs.Location = New System.Drawing.Point(24, 32)
        Me.btnCompGenerateOCEs.Name = "btnCompGenerateOCEs"
        Me.btnCompGenerateOCEs.Size = New System.Drawing.Size(132, 24)
        Me.btnCompGenerateOCEs.TabIndex = 0
        Me.btnCompGenerateOCEs.Text = "Generate OCE(s)"
        '
        'btnComplianceRefresh
        '
        Me.btnComplianceRefresh.Location = New System.Drawing.Point(520, 32)
        Me.btnComplianceRefresh.Name = "btnComplianceRefresh"
        Me.btnComplianceRefresh.Size = New System.Drawing.Size(132, 24)
        Me.btnComplianceRefresh.TabIndex = 0
        Me.btnComplianceRefresh.Text = "Refresh"
        '
        'pnlComplianceTop
        '
        Me.pnlComplianceTop.Controls.Add(Me.btnComplianceExpandCollapseAll)
        Me.pnlComplianceTop.Controls.Add(Me.Label1)
        Me.pnlComplianceTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlComplianceTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlComplianceTop.Name = "pnlComplianceTop"
        Me.pnlComplianceTop.Size = New System.Drawing.Size(920, 32)
        Me.pnlComplianceTop.TabIndex = 0
        '
        'btnComplianceExpandCollapseAll
        '
        Me.btnComplianceExpandCollapseAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnComplianceExpandCollapseAll.Location = New System.Drawing.Point(824, 5)
        Me.btnComplianceExpandCollapseAll.Name = "btnComplianceExpandCollapseAll"
        Me.btnComplianceExpandCollapseAll.Size = New System.Drawing.Size(88, 23)
        Me.btnComplianceExpandCollapseAll.TabIndex = 3
        Me.btnComplianceExpandCollapseAll.Text = "Expand All"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(192, 16)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Facility Compliance Events (FCEs)"
        '
        'tbPageEnforcement
        '
        Me.tbPageEnforcement.Controls.Add(Me.pnlEnforcementContainer)
        Me.tbPageEnforcement.Controls.Add(Me.pnlEnforcementBottom)
        Me.tbPageEnforcement.Controls.Add(Me.pnlEnforcementTop)
        Me.tbPageEnforcement.Location = New System.Drawing.Point(4, 22)
        Me.tbPageEnforcement.Name = "tbPageEnforcement"
        Me.tbPageEnforcement.Size = New System.Drawing.Size(920, 462)
        Me.tbPageEnforcement.TabIndex = 2
        Me.tbPageEnforcement.Text = "Enforcement"
        Me.tbPageEnforcement.Visible = False
        '
        'pnlEnforcementContainer
        '
        Me.pnlEnforcementContainer.Controls.Add(Me.ugEnforcement)
        Me.pnlEnforcementContainer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlEnforcementContainer.Location = New System.Drawing.Point(0, 32)
        Me.pnlEnforcementContainer.Name = "pnlEnforcementContainer"
        Me.pnlEnforcementContainer.Size = New System.Drawing.Size(920, 382)
        Me.pnlEnforcementContainer.TabIndex = 1
        '
        'ugEnforcement
        '
        Me.ugEnforcement.Cursor = System.Windows.Forms.Cursors.Default
        Appearance5.BorderColor = System.Drawing.SystemColors.ControlLight
        Me.ugEnforcement.DisplayLayout.Override.CellAppearance = Appearance5
        Me.ugEnforcement.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Appearance6.BorderColor = System.Drawing.SystemColors.ControlLight
        Me.ugEnforcement.DisplayLayout.Override.RowAppearance = Appearance6
        Me.ugEnforcement.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugEnforcement.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugEnforcement.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugEnforcement.Location = New System.Drawing.Point(0, 0)
        Me.ugEnforcement.Name = "ugEnforcement"
        Me.ugEnforcement.Size = New System.Drawing.Size(920, 382)
        Me.ugEnforcement.TabIndex = 0
        '
        'pnlEnforcementBottom
        '
        Me.pnlEnforcementBottom.Controls.Add(Me.btnManualEsc)
        Me.pnlEnforcementBottom.Controls.Add(Me.brnAgreedOrder)
        Me.pnlEnforcementBottom.Controls.Add(Me.btnInsViewEditCheckList2)
        Me.pnlEnforcementBottom.Controls.Add(Me.btnEnforceGenerateLetter)
        Me.pnlEnforcementBottom.Controls.Add(Me.btnEnforceProcessRescissions)
        Me.pnlEnforcementBottom.Controls.Add(Me.btnEnforceRefresh)
        Me.pnlEnforcementBottom.Controls.Add(Me.btnEnforceViewEnforceHistory)
        Me.pnlEnforcementBottom.Controls.Add(Me.btnEnforceProcessEscalations)
        Me.pnlEnforcementBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlEnforcementBottom.Location = New System.Drawing.Point(0, 414)
        Me.pnlEnforcementBottom.Name = "pnlEnforcementBottom"
        Me.pnlEnforcementBottom.Size = New System.Drawing.Size(920, 48)
        Me.pnlEnforcementBottom.TabIndex = 2
        '
        'brnAgreedOrder
        '
        Me.brnAgreedOrder.Location = New System.Drawing.Point(584, 8)
        Me.brnAgreedOrder.Name = "brnAgreedOrder"
        Me.brnAgreedOrder.Size = New System.Drawing.Size(104, 32)
        Me.brnAgreedOrder.TabIndex = 6
        Me.brnAgreedOrder.Text = "Agreed Order Letter"
        '
        'btnInsViewEditCheckList2
        '
        Me.btnInsViewEditCheckList2.Location = New System.Drawing.Point(696, 8)
        Me.btnInsViewEditCheckList2.Name = "btnInsViewEditCheckList2"
        Me.btnInsViewEditCheckList2.Size = New System.Drawing.Size(96, 32)
        Me.btnInsViewEditCheckList2.TabIndex = 5
        Me.btnInsViewEditCheckList2.Text = "View/Edit CheckList"
        '
        'btnEnforceGenerateLetter
        '
        Me.btnEnforceGenerateLetter.Location = New System.Drawing.Point(472, 8)
        Me.btnEnforceGenerateLetter.Name = "btnEnforceGenerateLetter"
        Me.btnEnforceGenerateLetter.Size = New System.Drawing.Size(104, 32)
        Me.btnEnforceGenerateLetter.TabIndex = 4
        Me.btnEnforceGenerateLetter.Text = "Generate Letter"
        '
        'btnEnforceProcessRescissions
        '
        Me.btnEnforceProcessRescissions.Location = New System.Drawing.Point(352, 8)
        Me.btnEnforceProcessRescissions.Name = "btnEnforceProcessRescissions"
        Me.btnEnforceProcessRescissions.Size = New System.Drawing.Size(104, 32)
        Me.btnEnforceProcessRescissions.TabIndex = 3
        Me.btnEnforceProcessRescissions.Text = "Process Rescissions"
        '
        'btnEnforceRefresh
        '
        Me.btnEnforceRefresh.Location = New System.Drawing.Point(240, 8)
        Me.btnEnforceRefresh.Name = "btnEnforceRefresh"
        Me.btnEnforceRefresh.Size = New System.Drawing.Size(96, 32)
        Me.btnEnforceRefresh.TabIndex = 2
        Me.btnEnforceRefresh.Text = "Refresh"
        '
        'btnEnforceViewEnforceHistory
        '
        Me.btnEnforceViewEnforceHistory.Location = New System.Drawing.Point(120, 8)
        Me.btnEnforceViewEnforceHistory.Name = "btnEnforceViewEnforceHistory"
        Me.btnEnforceViewEnforceHistory.Size = New System.Drawing.Size(104, 32)
        Me.btnEnforceViewEnforceHistory.TabIndex = 1
        Me.btnEnforceViewEnforceHistory.Text = "View Enforcement History"
        '
        'btnEnforceProcessEscalations
        '
        Me.btnEnforceProcessEscalations.Location = New System.Drawing.Point(0, 8)
        Me.btnEnforceProcessEscalations.Name = "btnEnforceProcessEscalations"
        Me.btnEnforceProcessEscalations.Size = New System.Drawing.Size(104, 32)
        Me.btnEnforceProcessEscalations.TabIndex = 0
        Me.btnEnforceProcessEscalations.Text = "Process Escalations"
        '
        'pnlEnforcementTop
        '
        Me.pnlEnforcementTop.Controls.Add(Me.btnEnforcementExpandCollapseAll)
        Me.pnlEnforcementTop.Controls.Add(Me.lblEnforceOCEs)
        Me.pnlEnforcementTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlEnforcementTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlEnforcementTop.Name = "pnlEnforcementTop"
        Me.pnlEnforcementTop.Size = New System.Drawing.Size(920, 32)
        Me.pnlEnforcementTop.TabIndex = 0
        '
        'btnEnforcementExpandCollapseAll
        '
        Me.btnEnforcementExpandCollapseAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnEnforcementExpandCollapseAll.Location = New System.Drawing.Point(824, 5)
        Me.btnEnforcementExpandCollapseAll.Name = "btnEnforcementExpandCollapseAll"
        Me.btnEnforcementExpandCollapseAll.Size = New System.Drawing.Size(88, 23)
        Me.btnEnforcementExpandCollapseAll.TabIndex = 4
        Me.btnEnforcementExpandCollapseAll.Text = "Expand All"
        '
        'lblEnforceOCEs
        '
        Me.lblEnforceOCEs.Location = New System.Drawing.Point(8, 16)
        Me.lblEnforceOCEs.Name = "lblEnforceOCEs"
        Me.lblEnforceOCEs.Size = New System.Drawing.Size(200, 16)
        Me.lblEnforceOCEs.TabIndex = 1
        Me.lblEnforceOCEs.Text = "Owner Compliance Events (OCEs)"
        '
        'tbPageAssignedInspections
        '
        Me.tbPageAssignedInspections.Controls.Add(Me.pnlAssignedInspectionsContainer)
        Me.tbPageAssignedInspections.Controls.Add(Me.pnlAssignedInspectionsTop)
        Me.tbPageAssignedInspections.Controls.Add(Me.pnlAssignedInspectionsBottom)
        Me.tbPageAssignedInspections.Location = New System.Drawing.Point(4, 22)
        Me.tbPageAssignedInspections.Name = "tbPageAssignedInspections"
        Me.tbPageAssignedInspections.Size = New System.Drawing.Size(920, 462)
        Me.tbPageAssignedInspections.TabIndex = 3
        Me.tbPageAssignedInspections.Text = "Assigned Inspections"
        '
        'pnlAssignedInspectionsContainer
        '
        Me.pnlAssignedInspectionsContainer.Controls.Add(Me.ugAssignedInspections)
        Me.pnlAssignedInspectionsContainer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlAssignedInspectionsContainer.Location = New System.Drawing.Point(0, 32)
        Me.pnlAssignedInspectionsContainer.Name = "pnlAssignedInspectionsContainer"
        Me.pnlAssignedInspectionsContainer.Size = New System.Drawing.Size(920, 382)
        Me.pnlAssignedInspectionsContainer.TabIndex = 1
        '
        'ugAssignedInspections
        '
        Me.ugAssignedInspections.Cursor = System.Windows.Forms.Cursors.Default
        Appearance7.BorderColor = System.Drawing.SystemColors.ControlLight
        Me.ugAssignedInspections.DisplayLayout.Override.CellAppearance = Appearance7
        Me.ugAssignedInspections.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Appearance8.BorderColor = System.Drawing.SystemColors.ControlLight
        Me.ugAssignedInspections.DisplayLayout.Override.RowAppearance = Appearance8
        Me.ugAssignedInspections.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugAssignedInspections.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugAssignedInspections.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugAssignedInspections.Location = New System.Drawing.Point(0, 0)
        Me.ugAssignedInspections.Name = "ugAssignedInspections"
        Me.ugAssignedInspections.Size = New System.Drawing.Size(920, 382)
        Me.ugAssignedInspections.TabIndex = 0
        '
        'pnlAssignedInspectionsTop
        '
        Me.pnlAssignedInspectionsTop.Controls.Add(Me.btnAssignedInspExpandCollapseAll)
        Me.pnlAssignedInspectionsTop.Controls.Add(Me.lblAssignedInspections)
        Me.pnlAssignedInspectionsTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlAssignedInspectionsTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlAssignedInspectionsTop.Name = "pnlAssignedInspectionsTop"
        Me.pnlAssignedInspectionsTop.Size = New System.Drawing.Size(920, 32)
        Me.pnlAssignedInspectionsTop.TabIndex = 0
        '
        'btnAssignedInspExpandCollapseAll
        '
        Me.btnAssignedInspExpandCollapseAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAssignedInspExpandCollapseAll.Location = New System.Drawing.Point(824, 5)
        Me.btnAssignedInspExpandCollapseAll.Name = "btnAssignedInspExpandCollapseAll"
        Me.btnAssignedInspExpandCollapseAll.Size = New System.Drawing.Size(88, 23)
        Me.btnAssignedInspExpandCollapseAll.TabIndex = 2
        Me.btnAssignedInspExpandCollapseAll.Text = "Expand All"
        '
        'lblAssignedInspections
        '
        Me.lblAssignedInspections.Location = New System.Drawing.Point(8, 16)
        Me.lblAssignedInspections.Name = "lblAssignedInspections"
        Me.lblAssignedInspections.Size = New System.Drawing.Size(112, 16)
        Me.lblAssignedInspections.TabIndex = 1
        Me.lblAssignedInspections.Text = "Assigned Inspections"
        '
        'pnlAssignedInspectionsBottom
        '
        Me.pnlAssignedInspectionsBottom.Controls.Add(Me.btnEditAssignedInspection)
        Me.pnlAssignedInspectionsBottom.Controls.Add(Me.btnAssignedInspectionsAccept)
        Me.pnlAssignedInspectionsBottom.Controls.Add(Me.btnAssignedInspectionsAdd)
        Me.pnlAssignedInspectionsBottom.Controls.Add(Me.btnAssignedRefresh)
        Me.pnlAssignedInspectionsBottom.Controls.Add(Me.btnAssignedInspectionsDelete)
        Me.pnlAssignedInspectionsBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlAssignedInspectionsBottom.Location = New System.Drawing.Point(0, 414)
        Me.pnlAssignedInspectionsBottom.Name = "pnlAssignedInspectionsBottom"
        Me.pnlAssignedInspectionsBottom.Size = New System.Drawing.Size(920, 48)
        Me.pnlAssignedInspectionsBottom.TabIndex = 2
        '
        'btnEditAssignedInspection
        '
        Me.btnEditAssignedInspection.Location = New System.Drawing.Point(192, 8)
        Me.btnEditAssignedInspection.Name = "btnEditAssignedInspection"
        Me.btnEditAssignedInspection.Size = New System.Drawing.Size(160, 24)
        Me.btnEditAssignedInspection.TabIndex = 2
        Me.btnEditAssignedInspection.Text = "Edit Assigned Inspections"
        '
        'btnAssignedInspectionsAccept
        '
        Me.btnAssignedInspectionsAccept.Location = New System.Drawing.Point(368, 8)
        Me.btnAssignedInspectionsAccept.Name = "btnAssignedInspectionsAccept"
        Me.btnAssignedInspectionsAccept.Size = New System.Drawing.Size(160, 24)
        Me.btnAssignedInspectionsAccept.TabIndex = 1
        Me.btnAssignedInspectionsAccept.Text = "Accept Assigned Inspections"
        '
        'btnAssignedInspectionsAdd
        '
        Me.btnAssignedInspectionsAdd.Location = New System.Drawing.Point(24, 8)
        Me.btnAssignedInspectionsAdd.Name = "btnAssignedInspectionsAdd"
        Me.btnAssignedInspectionsAdd.Size = New System.Drawing.Size(144, 24)
        Me.btnAssignedInspectionsAdd.TabIndex = 0
        Me.btnAssignedInspectionsAdd.Text = "Add Assigned Inspections"
        '
        'btnAssignedRefresh
        '
        Me.btnAssignedRefresh.Location = New System.Drawing.Point(720, 8)
        Me.btnAssignedRefresh.Name = "btnAssignedRefresh"
        Me.btnAssignedRefresh.Size = New System.Drawing.Size(160, 24)
        Me.btnAssignedRefresh.TabIndex = 1
        Me.btnAssignedRefresh.Text = "Refresh"
        '
        'btnAssignedInspectionsDelete
        '
        Me.btnAssignedInspectionsDelete.Location = New System.Drawing.Point(544, 8)
        Me.btnAssignedInspectionsDelete.Name = "btnAssignedInspectionsDelete"
        Me.btnAssignedInspectionsDelete.Size = New System.Drawing.Size(160, 24)
        Me.btnAssignedInspectionsDelete.TabIndex = 1
        Me.btnAssignedInspectionsDelete.Text = "Delete Assigned Inspections"
        '
        'tbPageLicensees
        '
        Me.tbPageLicensees.Controls.Add(Me.pnlLicenseesContainer)
        Me.tbPageLicensees.Controls.Add(Me.pnlLicenseesBottom)
        Me.tbPageLicensees.Controls.Add(Me.pnlLicenseesTop)
        Me.tbPageLicensees.Location = New System.Drawing.Point(4, 22)
        Me.tbPageLicensees.Name = "tbPageLicensees"
        Me.tbPageLicensees.Size = New System.Drawing.Size(920, 462)
        Me.tbPageLicensees.TabIndex = 4
        Me.tbPageLicensees.Text = "Licensees"
        '
        'pnlLicenseesContainer
        '
        Me.pnlLicenseesContainer.Controls.Add(Me.ugLicensees)
        Me.pnlLicenseesContainer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlLicenseesContainer.Location = New System.Drawing.Point(0, 32)
        Me.pnlLicenseesContainer.Name = "pnlLicenseesContainer"
        Me.pnlLicenseesContainer.Size = New System.Drawing.Size(920, 382)
        Me.pnlLicenseesContainer.TabIndex = 1
        '
        'ugLicensees
        '
        Me.ugLicensees.Cursor = System.Windows.Forms.Cursors.Default
        Appearance9.BorderColor = System.Drawing.SystemColors.ControlLight
        Me.ugLicensees.DisplayLayout.Override.CellAppearance = Appearance9
        Me.ugLicensees.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Appearance10.BorderColor = System.Drawing.SystemColors.ControlLight
        Me.ugLicensees.DisplayLayout.Override.RowAppearance = Appearance10
        Me.ugLicensees.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Extended
        Me.ugLicensees.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugLicensees.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugLicensees.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugLicensees.Location = New System.Drawing.Point(0, 0)
        Me.ugLicensees.Name = "ugLicensees"
        Me.ugLicensees.Size = New System.Drawing.Size(920, 382)
        Me.ugLicensees.TabIndex = 0
        '
        'pnlLicenseesBottom
        '
        Me.pnlLicenseesBottom.Controls.Add(Me.btnLicenseeGenerateLetter)
        Me.pnlLicenseesBottom.Controls.Add(Me.btnLicenseeProcessRescissions)
        Me.pnlLicenseesBottom.Controls.Add(Me.btnLicenseeRefresh)
        Me.pnlLicenseesBottom.Controls.Add(Me.btnLicenseeViewEnforcementHistory)
        Me.pnlLicenseesBottom.Controls.Add(Me.btnLicenseeProcessEscalations)
        Me.pnlLicenseesBottom.Controls.Add(Me.btnLicenseeDeleteLCE)
        Me.pnlLicenseesBottom.Controls.Add(Me.btnLicenseeEditLCE)
        Me.pnlLicenseesBottom.Controls.Add(Me.btnLicenseeAddLCE)
        Me.pnlLicenseesBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlLicenseesBottom.Location = New System.Drawing.Point(0, 414)
        Me.pnlLicenseesBottom.Name = "pnlLicenseesBottom"
        Me.pnlLicenseesBottom.Size = New System.Drawing.Size(920, 48)
        Me.pnlLicenseesBottom.TabIndex = 2
        '
        'btnLicenseeGenerateLetter
        '
        Me.btnLicenseeGenerateLetter.Location = New System.Drawing.Point(752, 8)
        Me.btnLicenseeGenerateLetter.Name = "btnLicenseeGenerateLetter"
        Me.btnLicenseeGenerateLetter.Size = New System.Drawing.Size(112, 32)
        Me.btnLicenseeGenerateLetter.TabIndex = 7
        Me.btnLicenseeGenerateLetter.Text = "Generate Letter"
        '
        'btnLicenseeProcessRescissions
        '
        Me.btnLicenseeProcessRescissions.Location = New System.Drawing.Point(640, 8)
        Me.btnLicenseeProcessRescissions.Name = "btnLicenseeProcessRescissions"
        Me.btnLicenseeProcessRescissions.Size = New System.Drawing.Size(104, 32)
        Me.btnLicenseeProcessRescissions.TabIndex = 6
        Me.btnLicenseeProcessRescissions.Text = "Process Rescissions"
        '
        'btnLicenseeRefresh
        '
        Me.btnLicenseeRefresh.Location = New System.Drawing.Point(520, 8)
        Me.btnLicenseeRefresh.Name = "btnLicenseeRefresh"
        Me.btnLicenseeRefresh.Size = New System.Drawing.Size(112, 32)
        Me.btnLicenseeRefresh.TabIndex = 5
        Me.btnLicenseeRefresh.Text = "Refresh"
        '
        'btnLicenseeViewEnforcementHistory
        '
        Me.btnLicenseeViewEnforcementHistory.Location = New System.Drawing.Point(392, 8)
        Me.btnLicenseeViewEnforcementHistory.Name = "btnLicenseeViewEnforcementHistory"
        Me.btnLicenseeViewEnforcementHistory.Size = New System.Drawing.Size(120, 32)
        Me.btnLicenseeViewEnforcementHistory.TabIndex = 4
        Me.btnLicenseeViewEnforcementHistory.Text = "View Enforcement History"
        '
        'btnLicenseeProcessEscalations
        '
        Me.btnLicenseeProcessEscalations.Location = New System.Drawing.Point(296, 8)
        Me.btnLicenseeProcessEscalations.Name = "btnLicenseeProcessEscalations"
        Me.btnLicenseeProcessEscalations.Size = New System.Drawing.Size(88, 32)
        Me.btnLicenseeProcessEscalations.TabIndex = 3
        Me.btnLicenseeProcessEscalations.Text = "Process Escalations"
        '
        'btnLicenseeDeleteLCE
        '
        Me.btnLicenseeDeleteLCE.Location = New System.Drawing.Point(200, 12)
        Me.btnLicenseeDeleteLCE.Name = "btnLicenseeDeleteLCE"
        Me.btnLicenseeDeleteLCE.Size = New System.Drawing.Size(86, 24)
        Me.btnLicenseeDeleteLCE.TabIndex = 2
        Me.btnLicenseeDeleteLCE.Text = "Delete LCE"
        '
        'btnLicenseeEditLCE
        '
        Me.btnLicenseeEditLCE.Location = New System.Drawing.Point(104, 12)
        Me.btnLicenseeEditLCE.Name = "btnLicenseeEditLCE"
        Me.btnLicenseeEditLCE.Size = New System.Drawing.Size(86, 24)
        Me.btnLicenseeEditLCE.TabIndex = 1
        Me.btnLicenseeEditLCE.Text = "Edit LCE"
        '
        'btnLicenseeAddLCE
        '
        Me.btnLicenseeAddLCE.Location = New System.Drawing.Point(8, 12)
        Me.btnLicenseeAddLCE.Name = "btnLicenseeAddLCE"
        Me.btnLicenseeAddLCE.Size = New System.Drawing.Size(86, 24)
        Me.btnLicenseeAddLCE.TabIndex = 0
        Me.btnLicenseeAddLCE.Text = "Add LCE"
        '
        'pnlLicenseesTop
        '
        Me.pnlLicenseesTop.Controls.Add(Me.btnLicenseesExpandCollapseAll)
        Me.pnlLicenseesTop.Controls.Add(Me.lblLicenseesComplianceEvents)
        Me.pnlLicenseesTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlLicenseesTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlLicenseesTop.Name = "pnlLicenseesTop"
        Me.pnlLicenseesTop.Size = New System.Drawing.Size(920, 32)
        Me.pnlLicenseesTop.TabIndex = 0
        '
        'btnLicenseesExpandCollapseAll
        '
        Me.btnLicenseesExpandCollapseAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnLicenseesExpandCollapseAll.Location = New System.Drawing.Point(824, 5)
        Me.btnLicenseesExpandCollapseAll.Name = "btnLicenseesExpandCollapseAll"
        Me.btnLicenseesExpandCollapseAll.Size = New System.Drawing.Size(88, 23)
        Me.btnLicenseesExpandCollapseAll.TabIndex = 4
        Me.btnLicenseesExpandCollapseAll.Text = "Expand All"
        '
        'lblLicenseesComplianceEvents
        '
        Me.lblLicenseesComplianceEvents.Location = New System.Drawing.Point(8, 16)
        Me.lblLicenseesComplianceEvents.Name = "lblLicenseesComplianceEvents"
        Me.lblLicenseesComplianceEvents.Size = New System.Drawing.Size(200, 16)
        Me.lblLicenseesComplianceEvents.TabIndex = 1
        Me.lblLicenseesComplianceEvents.Text = "Licensee Compliance Events (LCEs)"
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.cmbManager)
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(928, 26)
        Me.Panel1.TabIndex = 4
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(0, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 16)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "CNE Manager"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'cmbManager
        '
        Me.cmbManager.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbManager.Location = New System.Drawing.Point(104, 0)
        Me.cmbManager.Name = "cmbManager"
        Me.cmbManager.Size = New System.Drawing.Size(392, 21)
        Me.cmbManager.TabIndex = 0
        '
        'btnManualEsc
        '
        Me.btnManualEsc.Location = New System.Drawing.Point(800, 8)
        Me.btnManualEsc.Name = "btnManualEsc"
        Me.btnManualEsc.Size = New System.Drawing.Size(104, 32)
        Me.btnManualEsc.TabIndex = 7
        Me.btnManualEsc.Text = "Manual Escalation"
        '
        'CandEManagement
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(928, 518)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.tabCntrlCandE)
        Me.Name = "CandEManagement"
        Me.Text = "C and E Management"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.tabCntrlCandE.ResumeLayout(False)
        Me.tbPageInspections.ResumeLayout(False)
        Me.pnlInspectionsContainer.ResumeLayout(False)
        CType(Me.ugInspections, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlInspectionsBottom.ResumeLayout(False)
        Me.pnlInspectionsTop.ResumeLayout(False)
        Me.tbPageCompliance.ResumeLayout(False)
        Me.pnlComplianceContainer.ResumeLayout(False)
        CType(Me.ugCompliance, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlComplianceBottom.ResumeLayout(False)
        Me.grpCompAdminFCEFunctions.ResumeLayout(False)
        Me.pnlComplianceTop.ResumeLayout(False)
        Me.tbPageEnforcement.ResumeLayout(False)
        Me.pnlEnforcementContainer.ResumeLayout(False)
        CType(Me.ugEnforcement, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlEnforcementBottom.ResumeLayout(False)
        Me.pnlEnforcementTop.ResumeLayout(False)
        Me.tbPageAssignedInspections.ResumeLayout(False)
        Me.pnlAssignedInspectionsContainer.ResumeLayout(False)
        CType(Me.ugAssignedInspections, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlAssignedInspectionsTop.ResumeLayout(False)
        Me.pnlAssignedInspectionsBottom.ResumeLayout(False)
        Me.tbPageLicensees.ResumeLayout(False)
        Me.pnlLicenseesContainer.ResumeLayout(False)
        CType(Me.ugLicensees, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlLicenseesBottom.ResumeLayout(False)
        Me.pnlLicenseesTop.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "UI Support Routines"
    Private Sub ExpandAll(ByVal bol As Boolean, ByRef ug As Infragistics.Win.UltraWinGrid.UltraGrid, ByRef btn As Button)
        If bol Then
            btn.Text = "Collapse All"
            ug.Rows.ExpandAll(True)
        Else
            btn.Text = "Expand All"
            ug.Rows.CollapseAll(True)
        End If
    End Sub
    Private Sub PopulateInspections()
        Try
            ugInspections.DataSource = pFCE.GetInspections(Me.intWorkingFacility, cmbManager.SelectedValue)
            ugInspections.DrawFilter = rp



            For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In ugInspections.Rows
                If Not row.HasChild Then
                    row.Hidden = True
                Else
                    row.Hidden = False
                End If


            Next



        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub PopulateCompliances()
        Try
            ugCompliance.DataSource = pFCE.GetCompliances(0, False, cmbManager.SelectedValue)
            ugCompliance.DrawFilter = rp

            For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In ugCompliance.Rows
                If Not row.HasChild Then
                    row.Hidden = True
                Else
                    row.Hidden = False
                End If

            Next

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub PopulateEnforcements()
        Try

            Dim ds As DataSet = Nothing
            bolOCEModified = False
            bolOCEDueDateModified = False
            bolCitationModified = False

            GridMaster.GlobalInstance.ReadyForBotJobs(Me, Me.arrayWorkingvalues)

            intWorkingOwnerID = arrayWorkingvalues(0)
            intWorkingFacility = arrayWorkingvalues(1)

            ds = pOCE.GetEnforcements(intWorkingOwnerID, False, , , intWorkingFacility, True, cmbManager.SelectedValue)



            ugEnforcement.DataSource = ds

            ugEnforcement.DrawFilter = rp

            For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In Me.ugEnforcement.Rows

                If row.HasChild Then

                    For Each row2 As Infragistics.Win.UltraWinGrid.UltraGridRow In row.ChildBands(0).Rows

                        If Not row2.Cells("WORKSHOP DATE").Value Is DBNull.Value AndAlso row2.Cells("WORKSHOP DATE").Value > CDate("1/1/1900") AndAlso CDate(row2.Cells("WORKSHOP DATE").Value) < Today.Date AndAlso (row2.Cells("WORKSHOP RESULT").Value Is DBNull.Value OrElse row2.Cells("WORKSHOP RESULT").Value = 0) Then
                            row2.Cells("Workshop Result").Appearance.BackColor2 = Color.Red
                            row2.Cells("Workshop Result").Appearance.BackColor = Color.Red

                        End If


                        '   GetOCEEscalation(row2, True)

                        If row2.HasChild Then

                            row2.Hidden = False

                            If lastUGRowOwner > 0 AndAlso lastUGRowOwner = row2.Cells("OWNER_ID").Value Then

                                row.Expanded = True
                                row2.Expanded = True
                                row2.Cells("SELECTED").Value = True

                            End If

                        ElseIf Not row2.Hidden Then
                            row2.Hidden = True

                        End If

                    Next

                    lastUGRowOwner = 0
                ElseIf Not row.Hidden Then
                    row.Hidden = True

                End If

            Next


            If intWorkingFacility > 0 OrElse intWorkingOwnerID > 0 Then

                ugEnforcement.Rows.ExpandAll(False)

                For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In Me.ugEnforcement.Rows


                    If row.HasChild Then

                        row.ChildBands(0).Rows.ExpandAll(False)

                    End If

                Next

            End If


        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub PopulateAssignedInspections()
        Try
            ugAssignedInspections.DataSource = pInspection.GetAssignedFacilities(0, False)
            ugAssignedInspections.DrawFilter = rp
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub PopulateLicensees()
        Try
            ugLicensees.DataSource = Nothing
            pLCE.GetAll()
            ugLicensees.DataSource = pLCE.EntityTable(True)
            ugLicensees.DrawFilter = rp
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub SetupTabs()
        Try
            Me.Cursor = Cursors.WaitCursor

            Select Case tabCntrlCandE.SelectedTab.Name
                Case tbPageInspections.Name
                    PopulateInspections()
                    btnInspectionsExpandCollapseAll.Text = "Expand All"
                    'ExpandAll(False, ugInspections, btnInspectionsExpandCollapseAll)
                Case tbPageCompliance.Name
                    PopulateCompliances()
                    btnComplianceExpandCollapseAll.Text = "Expand All"
                    'ExpandAll(False, ugCompliance, btnComplianceExpandCollapseAll)
                Case tbPageEnforcement.Name
                    PopulateEnforcements()

                    If Me.intWorkingFacility > 0 Then
                        btnEnforcementExpandCollapseAll.Text = "Collapse All"
                    Else
                        btnEnforcementExpandCollapseAll.Text = "Expand All"
                    End If

                    'ExpandAll(False, ugEnforcement, btnEnforcementExpandCollapseAll)

                Case tbPageAssignedInspections.Name
                    PopulateAssignedInspections()
                    btnAssignedInspExpandCollapseAll.Text = "Expand All"
                    'ExpandAll(False, ugAssignedInspections, btnAssignedInspExpandCollapseAll)
                Case tbPageLicensees.Name
                    PopulateLicensees()
                    btnLicenseesExpandCollapseAll.Text = "Expand All"
                    'ExpandAll(False, ugLicensees, btnLicenseesExpandCollapseAll)

            End Select

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub


    Public Function getSelectedRow() As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            Select Case tabCntrlCandE.SelectedTab.Name

                Case tbPageInspections.Name

                    If ugInspections.Rows.Count > 0 Then

                        If Not ugInspections.ActiveRow Is Nothing Then

                            If ugInspections.ActiveRow.Band.Index = 0 Then

                                Dim facSelectedcount As Integer = 0

                                For Each ugChildRow In ugInspections.ActiveRow.ChildBands(0).Rows

                                    If ugChildRow.Cells("SELECTED").Value = True Or _
                                        ugChildRow.Cells("SELECTED").Text = True Then
                                        facSelectedcount += 1
                                        ' if multiple facilities are selected, which facility's row do u want
                                        If facSelectedcount > 1 Then
                                            Return Nothing
                                        End If
                                        ugRow = ugChildRow
                                    End If

                                Next

                            ElseIf ugInspections.ActiveRow.Band.Index = 1 Then

                                ugRow = ugInspections.ActiveRow
                            ElseIf ugInspections.ActiveRow.Band.Index = 2 Then
                                ugRow = ugInspections.ActiveRow.ParentRow
                            End If
                        End If
                    End If

                Case tbPageCompliance.Name
                    If ugCompliance.Rows.Count > 0 Then
                        If Not ugCompliance.ActiveRow Is Nothing Then
                            If ugCompliance.ActiveRow.Band.Index = 0 Or _
                                ugCompliance.ActiveRow.Band.Index = 1 Then
                                ugRow = ugCompliance.ActiveRow
                            ElseIf ugCompliance.ActiveRow.Band.Index = 2 Then
                                ugRow = ugCompliance.ActiveRow.ParentRow
                            End If
                        End If
                    End If

                Case tbPageEnforcement.Name
                    If ugEnforcement.ActiveRow.Band.Index = 2 Then
                        ugRow = ugEnforcement.ActiveRow
                    ElseIf ugEnforcement.ActiveRow.Band.Index = 3 Then
                        ugRow = ugEnforcement.ActiveRow.ParentRow
                    Else

                        ugRow = Nothing

                    End If

                Case tbPageAssignedInspections.Name
                    If ugAssignedInspections.Rows.Count > 0 Then
                        If Not ugAssignedInspections.ActiveRow Is Nothing Then
                            ugRow = ugAssignedInspections.ActiveRow
                        End If
                    End If
                Case tbPageLicensees.Name
                    ugRow = Nothing
                Case Else
                    ugRow = Nothing
            End Select
            Return ugRow
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function


    'Public Function getSelectedRows() As Infragistics.Win.UltraWinGrid.SelectedRowsCollection
    '    Dim ugRows As Infragistics.Win.UltraWinGrid.SelectedRowsCollection
    '    Dim ug As Infragistics.Win.UltraWinGrid.UltraGrid
    '    Try
    '        ug = getSelectedGrid()
    '        If Not ug Is Nothing Then
    '            If ug.Rows.Count > 0 Then
    '                If ug.Selected.Rows.Count > 0 Then
    '                    ugRows = ug.Selected.Rows
    '                End If
    '            End If
    '        End If
    '        Return ugRows
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Function
#End Region


#Region "Inspections"

    'Private Sub btnInsGenerateFCEs_ClickOld(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    ' 7/30 starts
    '    Dim strOwnerIDs As String = String.Empty
    '    Dim strInspectionIDs As String = String.Empty
    '    Dim strFacilityIDs As String = String.Empty
    '    Dim bolStatus As Boolean = False
    '    Dim dtNullDate As Date
    '    Dim bolNoDiscrepancyLetter As Boolean = False
    '    Dim bolAtleastOneInspection As Boolean = False
    '    Dim oUser As New MUSTER.BusinessLogic.pUser
    '    Dim dtView As DataView
    '    Dim rowCol As Infragistics.Win.UltraWinGrid.RowsCollection
    '    Try
    '        ''For Owner Band
    '        'For Each ugrow In ugInspections.Rows

    '        '    If ugrow.Cells("Selected").Text.ToUpper = "TRUE" Then
    '        '        bolStatus = True
    '        '        If strOwnerIDs <> String.Empty Then
    '        '            strOwnerIDs += "," + ugrow.Cells("Owner_ID").Value.ToString
    '        '        Else
    '        '            strOwnerIDs += ugrow.Cells("Owner_ID").Value.ToString
    '        '        End If
    '        '    End If

    '        For Each dr As Infragistics.Win.UltraWinGrid.UltraGridRow In ugInspections.Rows
    '            If dr.Cells("Selected").Value Then
    '                bolAtleastOneInspection = True
    '                bolNoDiscrepancyLetter = False

    '                rowCol = dr.ChildBands("OwnerToFacility".ToUpper).Rows
    '                ' loop through the inspections/facilities
    '                For Each dr1 As Infragistics.Win.UltraWinGrid.UltraGridRow In rowCol
    '                    If dr1.Cells("Selected").Value Then
    '                        pIns.RetrieveCheckListInfo(CInt(dr1.Cells("Inspection_ID").Value), CInt(dr1.Cells("facility_id").Value), CInt(dr1.Cells("Owner_ID").Value))
    '                        ' Clearing all Submitted Date
    '                        pIns.Completed = Now

    '                        ' For Inspections containing no citations - generate No discrepancy letter
    '                        If dr1.ChildBands Is Nothing Then
    '                            bolNoDiscrepancyLetter = True
    '                            If strFacilityIDs <> String.Empty Then
    '                                strFacilityIDs += "," + dr1.Cells("Facility_ID").Value.ToString
    '                            Else
    '                                strFacilityIDs += dr1.Cells("Facility_ID").Value.ToString
    '                            End If
    '                        Else
    '                            'Create FCE if it has One or more Citations.
    '                            pFCE.Retrieve(0)
    '                            pFCE.InspectionID = CInt(dr1.Cells("Inspection_ID").Value)
    '                            pFCE.OwnerID = CInt(dr1.Cells("Owner_ID").Value)
    '                            pFCE.FacilityID = CInt(dr1.Cells("Facility_ID").Value)
    '                            pFCE.FCEDate = Now
    '                            pFCE.Source = "CAE"
    '                            pFCE.Deleted = False
    '                            pFCE.Save(False)
    '                        End If

    '                        'Happy Scenario 3.f
    '                        Dim bolCheckListAnswer As Boolean = False
    '                        For Each checkList As MUSTER.Info.InspectionChecklistMasterInfo In pIns.CheckListMaster.InspectionInfo.ChecklistMasterCollection.Values
    '                            If checkList.CheckListItemNumber = "4.2.7" Or checkList.CheckListItemNumber = "5.2.7" Then
    '                                For Each response As MUSTER.Info.InspectionResponsesInfo In pIns.CheckListMaster.InspectionInfo.ResponsesCollection.Values
    '                                    If response.QuestionID = checkList.ID Then
    '                                        If response.Response = 0 Then
    '                                            bolCheckListAnswer = True
    '                                        End If
    '                                        Exit For
    '                                    End If
    '                                Next
    '                            End If
    '                            If bolCheckListAnswer Then
    '                                Exit For
    '                            End If
    '                        Next
    '                        If bolCheckListAnswer Then
    '                            ' Create a LUST Event
    '                            Dim pLust As New MUSTER.BusinessLogic.pLustEvent
    '                            pLust.Retrieve(0)
    '                            pLust.Started = Now
    '                            ' Date of report = Inspection Date
    '                            ' Inspection date is the most recent date for a given inspection in tblINS_INSPECTION_DATES
    '                            pIns.Retrieve(dr1.Cells("Inspection_ID").Value)

    '                            dtView = pIns.CheckListMaster.GetCLInspectionHistory().Tables(0).DefaultView
    '                            dtView.Sort = "[DATE INSPECTED] DESC"
    '                            pLust.ReportDate = dtView.Item(0)("DATE INSPECTED")
    '                            'plust.ReportDate = pIns.
    '                            ' Release status = Confirmed
    '                            pLust.ReleaseStatus = 623
    '                            'Dim oUser As New MUSTER.BusinessLogic.pUser
    '                            Dim oUserInfo As MUSTER.Info.UserInfo
    '                            oUserInfo = oUser.RetrievePMHead()
    '                            pLust.PM = oUserInfo.ID

    '                            'Event Status = Open
    '                            pLust.EventStatus = 624
    '                            'MGPTFStatus = EUD
    '                            pLust.MGPTFStatus = 617
    '                            'Confirmed on  = Inspection Date
    '                            pLust.Confirmed = dtView.Item(0)("DATE INSPECTED")
    '                            'identified by = Inspector
    '                            pLust.IDENTIFIEDBY = "Inspector"
    '                            'How discovered  = Inspection
    '                            pLust.HowDiscoveredID = 654
    '                            'Create a LUST Event
    '                            pLust.Save()

    '                            'Create a TO DO Calender entry on the current date for the PM-Head user
    '                            'indicating "New LUST Event Created from Inspection"
    '                            Dim ocalInfo As MUSTER.Info.CalendarInfo

    '                            Dim pCal As New MUSTER.BusinessLogic.pCalendar
    '                            oUser.RetrievePMHead()
    '                            Dim nEntityTypeID As Integer = New MUSTER.BusinessLogic.pEntity("C&E").ID
    '                            ocalInfo = New MUSTER.Info.CalendarInfo(0, _
    '                                        Now(), _
    '                                        Now(), _
    '                                        0, _
    '                                        "New LUST Event created from Inspection" + pIns.ID.ToString, _
    '                                        oUser.ID, _
    '                                        "SYSTEM", _
    '                                        String.Empty, _
    '                                        False, _
    '                                        True, _
    '                                        False, _
    '                                        False, _
    '                                        MusterContainer.AppUser.ID, _
    '                                        Now(), _
    '                                        String.Empty, _
    '                                        CDate("01/01/0001"), _
    '                                        nEntityTypeID, _
    '                                        pIns.ID)
    '                            pCal.Add(ocalInfo)
    '                            pCal.Save()
    '                        End If
    '                    End If
    '                Next
    '                If bolNoDiscrepancyLetter Then
    '                    'Generate NO discrepancy letter for each owner
    '                    MsgBox("Owner has one or more selected Inspections containing no citations. Please wait untill the No discrepancy letters are generated")

    '                End If
    '            End If
    '        Next
    '        If bolAtleastOneInspection Then
    '            pIns.Flush()
    '        End If

    '        MsgBox("FCE is generated successfully")
    '        tabCntrlCandE.SelectedTab = tbPageCompliance
    '        SetupTabs()

    '        '    'For Facility Band
    '        '    For Each ugchildband In ugrow.ChildBands
    '        '        For Each ugchildrow In ugchildband.Rows
    '        '            If ugrow.Cells("Selected").Text.ToUpper = "TRUE" Then
    '        '                If Not IsDBNull(ugchildrow.Cells("ScheduledDate").Value) And IsDBNull(ugchildrow.Cells("SubmittedDate").Value) Then
    '        '                    Dim msgResult = MsgBox("WARNING!" & vbCrLf & "Selected Facility has been Scheduled but not Submitted. Do you want to Continue?", MsgBoxStyle.Exclamation Or MsgBoxStyle.YesNo, "Warning")

    '        '                    If msgResult = MsgBoxResult.No Then
    '        '                        Exit Sub
    '        '                    Else
    '        '                        ' Clearing all Submitted Date
    '        '                        pIns.Retrieve(CInt(ugchildrow.Cells("Inspection_ID").Value))
    '        '                        pIns.Completed = Now

    '        '                        'To generate DiscrepancyLetter
    '        '                        If CInt(ugchildrow.Cells("Inspection_ID").Value) <= 0 Then
    '        '                            bolNoDiscrepancyLetter = True
    '        '                        End If

    '        '                        If strInspectionIDs <> String.Empty Then
    '        '                            strInspectionIDs += "," + ugchildrow.Cells("Inspection_ID").Value.ToString
    '        '                        Else
    '        '                            strInspectionIDs += ugchildrow.Cells("Inspection_ID").Value.ToString
    '        '                        End If



    '        '                        If strFacilityIDs <> String.Empty Then
    '        '                            strFacilityIDs += "," + ugchildrow.Cells("Facility_ID").Value.ToString
    '        '                        Else
    '        '                            strFacilityIDs += ugchildrow.Cells("Facility_ID").Value.ToString
    '        '                        End If

    '        '                    End If
    '        '                End If

    '        '                'Create FCE if it has One or more Citations.
    '        '                If Not IsDBNull(ugchildrow.Cells("SubmittedDate").Value) Then
    '        '                    If CInt(ugchildrow.Cells("Citations").Value) > 0 Then
    '        '                        pFCE.Retrieve(0)
    '        '                        pFCE.InspectionID = CInt(ugchildrow.Cells("Inspection_ID").Value)
    '        '                        pFCE.OwnerID = CInt(ugchildrow.Cells("Owner_ID").Value)
    '        '                        pFCE.FacilityID = CInt(ugchildrow.Cells("Facility_ID").Value)
    '        '                        pFCE.FCEDate = Now
    '        '                        pFCE.Source = 1
    '        '                        pFCE.Deleted = False
    '        '                        pFCE.Save(False)
    '        '                    End If

    '        '                    ' Citation Band -- To Insert FCE Citation.
    '        '                    'For Each ugcitationband In ugchildrow.ChildBands
    '        '                    '    For Each ugcitationrow In ugcitationband.Rows
    '        '                    '        pInsCitation.Add(0)
    '        '                    '        pInsCitation.InspectionID = CInt(ugcitationrow.Cells("Inspection_ID").Value)
    '        '                    '        pInsCitation.FacilityID = CInt(ugcitationrow.Cells("Facility_ID").Value)
    '        '                    '        pInsCitation.QuestionID = CInt(ugcitationrow.Cells("Question_ID").Value)
    '        '                    '        pInsCitation.CitationID = CInt(ugcitationrow.Cells("Citation_ID").Value)
    '        '                    '        ' Need Clarification for the following fields.
    '        '                    '        'pInsCitation.Rescinded = False
    '        '                    '        'pInsCitation.CitationDueDate = Now
    '        '                    '        'pInsCitation.CitationReceivedDate = now
    '        '                    '        pInsCitation.Save(True, True)
    '        '                    '    Next
    '        '                    'Next
    '        '                End If
    '        '            End If
    '        '        Next
    '        '    Next

    '        '    If bolNoDiscrepancyLetter Then
    '        '        'Generate Discrepancy Letter for Owner.
    '        '    End If
    '        'Next

    '        '' Save All Inspections after clearing Submitted Date.
    '        'If strInspectionIDs <> String.Empty Then
    '        '    pIns.Flush()
    '        'End If

    '        ''7/30 - Ends

    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub


    Private Sub btnInsGenerateFCEs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInsGenerateFCEs.Click
        Dim bolAnsIsNo, bolHasCitation10 As Boolean
        Dim dtOwners As New DataTable
        Dim dtFacs As New DataTable
        Dim alOwners As New ArrayList
        Dim drOwner As DataRow
        Dim drFac As DataRow
        Dim dr As DataRow
        Dim bolFCEGenerated, bolLetterGenerated, bolLUSTEventCreated, bolUnsubmittedInsp, bolAnySelected, bolContinue As Boolean

        Try

            bolFCEGenerated = False
            bolLetterGenerated = False
            bolLUSTEventCreated = False
            bolHasCitation10 = False
            ' datatable to store owner id's and facility id's to
            ' generate No Discrepancy Letter (Happy Scenario 3.c)
            dtOwners.Columns.Add("OWNER_ID", GetType(Integer))
            dtOwners.Columns.Add("OWNERNAME", GetType(String))
            dtOwners.Columns.Add("ADDRESS_LINE_ONE", GetType(String))
            dtOwners.Columns.Add("ADDRESS_TWO", GetType(String))
            dtOwners.Columns.Add("CITY", GetType(String))
            dtOwners.Columns.Add("STATE", GetType(String))
            dtOwners.Columns.Add("ZIP", GetType(String))

            dtFacs.Columns.Add("OWNER_ID", GetType(Integer))
            dtFacs.Columns.Add("FACILITY_ID", GetType(Integer))
            dtFacs.Columns.Add("INSPECTION_ID", GetType(Integer))
            dtFacs.Columns.Add("LETTER_GENERATED", GetType(Boolean))
            dtFacs.Columns.Add("FACILITY", GetType(String))
            dtFacs.Columns.Add("ADDRESS", GetType(String))

            For Each ugRow In ugInspections.Rows ' owner
                bolAnySelected = False
                bolUnsubmittedInsp = False
                bolContinue = False
                If Not bolUnsubmittedInsp Then
                    If ugRow.Cells("SELECTED").Value = True Or _
                        ugRow.Cells("SELECTED").Text = True Then
                        bolAnySelected = True
                    End If
                    For Each ugChildRow In ugRow.ChildBands(0).Rows  'facs
                        If ugChildRow.Cells("SELECTED").Value = True Or _
                            ugChildRow.Cells("SELECTED").Text = True Then
                            bolAnySelected = True
                        End If
                        If Not (ugChildRow.Cells("SCHEDULED").Value Is DBNull.Value) And _
                            ugChildRow.Cells("SUBMITTED").Value Is DBNull.Value Then
                            bolUnsubmittedInsp = True
                        End If
                        If bolAnySelected And bolUnsubmittedInsp Then Exit For
                    Next
                    ' if there is any selected row and any inspection which is scheduled but not submitted
                    If bolAnySelected And bolUnsubmittedInsp Then
                        If MsgBox("There are Scheduled but not Submitted Inspections for Owner: " + vbCrLf + _
                            ugRow.Cells("OWNERNAME").Text + vbCrLf + _
                            "Do you want to continue Generating FCE(s) for this Owner?", _
                            MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                            bolContinue = True
                        End If
                    ElseIf bolAnySelected And Not bolUnsubmittedInsp Then
                        bolContinue = True
                    End If
                End If
                If bolContinue Then
                    For Each ugChildRow In ugRow.ChildBands(0).Rows
                        If ugChildRow.Cells("SELECTED").Value = True Or _
                            ugChildRow.Cells("SELECTED").Text = True Then
                            bolAnsIsNo = False
                            bolHasCitation10 = False
                            ' Happy Scenario 3.b
                            ' setting Completed date removes the inspecton from inspector and inspections grid




                            Dim id As Integer = ugChildRow.Cells("INSPECTION_ID").Value

                            pInspection.Retrieve(id)

                            Dim newowner As Integer = pInspection.OwnerID

                            If Not ugChildRow.Cells("Owner Status").Value Is Nothing AndAlso ugChildRow.Cells("Owner Status").Value.ToString.Length > 0 Then

                                newowner = ugChildRow.Cells("NewOwner").Value

                                If MsgBox("This facility has been recently transferred to a new owner, would you like to transfer the Inspection as well.", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                                    pInspection.OwnerID = newowner
                                    pInspection.ModifiedBy = MusterContainer.AppUser.ID
                                    pInspection.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal, True)
                                    pInspection.Retrieve(id, , , newowner)
                                Else
                                    newowner = pInspection.OwnerID
                                End If

                            End If



                            pInspection.Completed = Today.Date
                            ' will save after 3.c, 3.d, 3.e and 3.f are complete
                            ' sort of rollback if there is an error
                            'pIns.Save()
                            If ugChildRow.ChildBands(0).Rows.Count = 0 Then
                                ' Happy Scenario 3.c
                                ' No Discrepancy letter for each owner that has one or more selected inspections
                                ' with no citations
                                If Not alOwners.Contains(newowner) Then
                                    alOwners.Add(newowner)
                                    drOwner = dtOwners.NewRow
                                    drOwner("OWNER_ID") = newowner
                                    drOwner("OWNERNAME") = ugChildRow.ParentRow.Cells("OWNERNAME").Value
                                    drOwner("ADDRESS_LINE_ONE") = ugChildRow.Cells("OWNADDRESS_LINE_ONE").Value
                                    drOwner("ADDRESS_TWO") = ugChildRow.Cells("OWNADDRESS_TWO").Value
                                    drOwner("CITY") = ugChildRow.Cells("OWNCITY").Value
                                    drOwner("STATE") = ugChildRow.Cells("OWNSTATE").Value
                                    drOwner("ZIP") = ugChildRow.Cells("OWNZIP").Value
                                    dtOwners.Rows.Add(drOwner)
                                End If

                                drFac = dtFacs.NewRow
                                drFac("OWNER_ID") = newowner
                                drFac("FACILITY_ID") = ugChildRow.Cells("FACILITY_ID").Value
                                drFac("INSPECTION_ID") = ugChildRow.Cells("INSPECTION_ID").Value
                                drFac("LETTER_GENERATED") = False
                                drFac("FACILITY") = ugChildRow.Cells("FACILITY").Text
                                drFac("ADDRESS") = ugChildRow.Cells("ADDRESS_LINE_ONE").Text + ", "
                                'If Not ugChildRow.Cells("ADDRESS_TWO").Value Is DBNull.Value Then
                                '    drFac("ADDRESS") += ugChildRow.Cells("ADDRESS_TWO").Text.Trim + ", "
                                'ElseIf ugChildRow.Cells("ADDRESS_TWO").Text.Trim.Length > 0 Then
                                '    drFac("ADDRESS") += ugChildRow.Cells("ADDRESS_TWO").Text.Trim + ", "
                                'End If
                                drFac("ADDRESS") += ugChildRow.Cells("CITY").Text.Trim + ", "
                                drFac("ADDRESS") += ugChildRow.Cells("STATE").Text.Trim + " "
                                drFac("ADDRESS") += ugChildRow.Cells("ZIP").Text.Trim
                                dtFacs.Rows.Add(drFac)
                            Else
                                ' Happy Scenario 3.d
                                ' Create FCE if it has One or more Citations
                                ' if fce already created, it returns that record
                                pFCE.Retrieve(, ugChildRow.Cells("INSPECTION_ID").Value, ugChildRow.Cells("OWNER_ID").Value, ugChildRow.Cells("FACILITY_ID").Value)
                                pFCE.InspectionID = ugChildRow.Cells("INSPECTION_ID").Value
                                pFCE.OwnerID = newowner
                                pFCE.FacilityID = ugChildRow.Cells("FACILITY_ID").Value
                                pFCE.FCEDate = Today.Date
                                pFCE.Source = "INSPECTION"
                                If pFCE.ID <= 0 Then
                                    pFCE.CreatedBy = MusterContainer.AppUser.ID
                                Else
                                    pFCE.ModifiedBy = MusterContainer.AppUser.ID
                                End If
                                pFCE.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal, True)
                                If Not UIUtilsGen.HasRights(returnVal) Then
                                    Exit Sub
                                End If

                                bolFCEGenerated = True

                                ' Happy Scenario 3.e
                                ' create a FCE Citation for each submitted inspection citation
                                ' reusing inspection citation. just entering fceid to the citation row in db
                                For Each ugGrandChildRow In ugChildRow.ChildBands(0).Rows
                                    If ugGrandChildRow.Cells("SELECTED").Value = True Or _
                                        ugGrandChildRow.Cells("SELECTED").Text = True Then
                                        pInsCitation.Retrieve(pInspection.InspectionInfo, ugGrandChildRow.Cells("INS_CIT_ID").Value)
                                        pInsCitation.FCEID = pFCE.ID
                                        If pInsCitation.ID <= 0 Then
                                            pInsCitation.CreatedBy = MusterContainer.AppUser.ID
                                        Else
                                            pInsCitation.ModifiedBy = MusterContainer.AppUser.ID
                                        End If
                                        pInsCitation.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal, True)
                                        If Not UIUtilsGen.HasRights(returnVal) Then
                                            Exit Sub
                                        End If
                                        ' #2927
                                        If pInsCitation.CitationID = 10 Then
                                            bolHasCitation10 = True
                                        End If
                                    End If

                                    ' Happy Scenario 3.f
                                    ' checking to see if answer to question 4.2.7 or 5.2.7 on the checklist is NO
                                    ' according to Inspection use case, if answer to question 4.2.7 or 5.2.7 on the checklist is NO
                                    ' a citation is created. Hence, if there are citations for question 4.2.7 or 5.2.7
                                    ' then the answer to the questions on the checklist is NO
                                    If ugGrandChildRow.Cells("QUESTION#").Value = "4.2.7" Or ugGrandChildRow.Cells("QUESTION#").Value = "5.2.7" Then
                                        bolAnsIsNo = True
                                    End If
                                Next
                            End If

                            ' Happy Scenario 3.f
                            ' If answer to question 4.2.7 or 5.2.7 on the checklist is NO, create LUST event and TODO cal entry
                            If bolAnsIsNo Then
                                Dim pLustEvent As New MUSTER.BusinessLogic.pLustEvent
                                Dim oUserInfo As MUSTER.Info.UserInfo
                                Dim oUser As New MUSTER.BusinessLogic.pUser
                                oUserInfo = oUser.RetrievePMHead()
                                pLustEvent.Add( _
                                    New MUSTER.Info.LustEventInfo(0, CDate("01/01/0001"), CDate("01/01/0001"), CDate("01/01/0001"), CDate("01/01/0001"), _
                                    ugChildRow.Cells("SCHEDULED").Value, Today.Date, 624, ugChildRow.Cells("FACILITY_ID").Value, 0, 617, 0, oUserInfo.UserKey, _
                                    623, 0, String.Empty, String.Empty, CDate("01/01/0001"), String.Empty, CDate("01/01/0001"), _
                                    0, ugChildRow.Cells("SCHEDULED").Value, 658, 0, 0, oUserInfo.UserKey, Today.Date, CDate("01/01/0001"), _
                                    False, False, False, False, False, False, False, False, False, False, False, False, False, False, False, _
                                    False, False, 0, String.Empty, String.Empty, 0, CDate("01/01/0001"), String.Empty, 0, CDate("01/01/0001"), _
                                    String.Empty, 0, CDate("01/01/0001"), String.Empty, False, 0, CDate("01/01/0001"), String.Empty, False, _
                                    String.Empty, String.Empty, String.Empty, String.Empty, 0, 0, False, False, False, False, False, False, False, False, False, False, True, 0))
                                pLustEvent.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal)
                                If Not UIUtilsGen.HasRights(returnVal) Then
                                    Exit Sub
                                End If

                                Dim oLustEventActivity As New MUSTER.BusinessLogic.pLustEventActivity
                                oLustEventActivity.Add(New MUSTER.Info.LustActivityInfo(0, _
                                    pLustEvent.ID, _
                                    Now.Date, _
                                    CDate("01/01/0001"), _
                                    CDate("01/01/0001"), _
                                    CDate("01/01/0001"), _
                                    CDate("01/01/0001"), _
                                    682, _
                                    MusterContainer.AppUser.ID, _
                                    Now.Date, _
                                    String.Empty, _
                                    CDate("01/01/0001"), _
                                    0, 0))
                                oLustEventActivity.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), CType(MusterContainer.AppUser.UserKey, Integer), returnVal)
                                If Not UIUtilsGen.HasRights(returnVal) Then
                                    Exit Sub
                                End If

                                bolLUSTEventCreated = True
                                ' Create TODO cal entry for pm head
                                Dim ocalInfo As New MUSTER.Info.CalendarInfo(0, _
                                                                    Now(), _
                                                                    Now(), _
                                                                    0, _
                                                                    "New LUST Event created from Inspection for Facility: " + ugChildRow.Cells("FACILITY_ID").Value.ToString, _
                                                                    oUserInfo.ID, _
                                                                    "SYSTEM", _
                                                                    String.Empty, _
                                                                    False, _
                                                                    True, _
                                                                    False, _
                                                                    False, _
                                                                    String.Empty, _
                                                                    CDate("01/01/0001"), _
                                                                    String.Empty, _
                                                                    CDate("01/01/0001"), _
                                                                    UIUtilsGen.EntityTypes.LustActivity, _
                                                                    oLustEventActivity.ActivityID)
                                MusterContainer.pCalendar.Add(ocalInfo)
                                MusterContainer.pCalendar.Save()
                                Dim mc As MusterContainer
                                mc = Me.MdiParent
                                If Not mc Is Nothing Then
                                    mc.RefreshCalendarInfo()
                                    mc.LoadDueToMeCalendar()
                                    mc.LoadToDoCalendar()
                                End If
                            End If

                            ' #2927
                            ' need to set only tos/tosi tanks/pipes dirty as TOS/TOSI rules for tanks/pipes whose status
                            ' has not changed applies only to tos/tosi tanks/pipes
                            If bolHasCitation10 Then
                                pInspection.RetrieveOwnerFacTanksPipes()
                                Dim tnk As MUSTER.Info.TankInfo
                                Dim pipe As MUSTER.Info.PipeInfo
                                For Each tnk In pInspection.CheckListMaster.Owner.Facility.TankCollection.Values
                                    If tnk.TankStatus = 425 Or tnk.TankStatus = 429 Then
                                        tnk.IsDirty = True
                                    End If
                                    For Each pipe In tnk.pipesCollection.Values
                                        If pipe.PipeStatusDesc = 425 Or pipe.PipeStatusDesc = 429 Then
                                            pipe.IsDirty = True
                                        End If
                                    Next
                                Next
                            End If

                            ' saving inspection after all the HappyScenario points are done
                            pInspection.ModifiedBy = MusterContainer.AppUser.ID
                            pInspection.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal)
                            If Not UIUtilsGen.HasRights(returnVal) Then
                                Exit Sub
                            End If
                            ' #2900 update inspection soc
                            pInspection.UpdateInspectionSOC(UIUtilsGen.ModuleID.CAE, MusterContainer.AppUser.UserKey, returnVal, pInspection.ID, MusterContainer.AppUser.ID)
                            If Not UIUtilsGen.HasRights(returnVal) Then
                                Exit Sub
                            End If

                            ' to remove owner, fac, tanks, and pipes from collection
                            pInspection = New MUSTER.BusinessLogic.pInspection

                        End If
                    Next
                End If
            Next

            'Happy Scenario 3.c
            If dtOwners.Rows.Count > 0 Then
                ' generate letter
                Dim regletter As New Reg_Letters
                For Each drOwner In dtOwners.Rows
                    Dim dt As New DataTable
                    dt = dtFacs.Clone
                    For Each drFac In dtFacs.Select("OWNER_ID = " + drOwner("OWNER_ID").ToString)
                        dr = dt.NewRow
                        For Each dtCol As DataColumn In dtFacs.Columns
                            dr(dtCol.ColumnName) = drFac(dtCol.ColumnName)
                        Next
                        dt.Rows.Add(dr)
                    Next
                    ' generate letter
                    If regletter.GenerateCAEFCENoDiscrepancyLetter(drOwner, dt) Then
                        For Each drFac In dtFacs.Select("OWNER_ID = " + drOwner("OWNER_ID").ToString)
                            drFac("LETTER_GENERATED") = True
                        Next
                    End If
                Next
                bolLetterGenerated = True
            End If

            Dim strMsg As String = String.Empty
            If bolFCEGenerated Then
                strMsg += "FCE(s), "
            End If
            If bolLUSTEventCreated Then
                strMsg += "Lust Event(s), "
            End If
            If bolLetterGenerated Then
                strMsg += "Letters, "
            End If
            If strMsg.Length > 0 Then
                MsgBox("Generated " + strMsg.Substring(0, strMsg.Length - 2) + " Successfully")
                SetupTabs()
            Else
                MsgBox("No FCE(s) / LUST Event(s) / Letter(s) created")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            ' if any facilities letter was not geenrated, clear completed date so they show up on inspections grid
            RollBackInspection(dtFacs)
        End Try
    End Sub
    Private Sub RollBackInspection(ByVal dtFacs As DataTable)
        Try
            If Not dtFacs Is Nothing Then
                If dtFacs.Select("LETTER_GENERATED = False").Length > 0 Then
                    For Each drFac As DataRow In dtFacs.Select("LETTER_GENERATED = False")
                        pInspection.Retrieve(drFac("INSPECTION_ID"))
                        pInspection.Completed = CDate("01/01/0001")
                        pInspection.ModifiedBy = MusterContainer.AppUser.ID
                        pInspection.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal)
                        If Not UIUtilsGen.HasRights(returnVal) Then
                            Exit Sub
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnInsViewEditCheckList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInsViewEditCheckList.Click
        Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            ugRow = getSelectedRow()
            If ugRow Is Nothing Then
                MsgBox("Please select a Facilty")
                Exit Sub
            End If
            If ugRow.Cells("SELECTED").Value = True Or _
                ugRow.Cells("SELECTED").Text = True Then
                If ugRow.Cells("SUBMITTED").Value Is DBNull.Value Then
                    MsgBox("Cannot Edit / View checklist - Inspection not Submitted to C&E")
                    Exit Sub
                Else
                    Dim id As Integer = ugRow.Cells("INSPECTION_ID").Value
                    Dim newOwner As Integer = ugRow.Cells("NewOwner").Value

                    pInspection.Retrieve(id)

                    If Not ugRow.Cells("Owner Status").Value Is Nothing AndAlso ugRow.Cells("Owner Status").Value.ToString.Length > 0 Then

                        If MsgBox("This facility has been recently transferred to a new owner, would you like to transfer the Inspection as well.", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                            pInspection.OwnerID = newOwner
                            pInspection.ModifiedBy = MusterContainer.AppUser.ID
                            pInspection.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal, True)

                            pInspection.Retrieve(id, , , newOwner)

                        End If

                    End If


                    Dim bolReadOnly As Boolean = True
                    If MusterContainer.AppUser.HasAccess(UIUtilsGen.ModuleID.CAE, MusterContainer.AppUser.UserKey, UIUtilsGen.EntityTypes.Inspection) Then
                        bolReadOnly = False
                    End If
                    Dim frmChecklist As New CheckList(pInspection, bolReadOnly, , UIUtilsGen.ModuleID.CAE)
                    frmChecklist.WindowState = FormWindowState.Maximized
                    frmChecklist.CallingForm = Me
                    Me.Tag = "0"




                    frmChecklist.ShowDialog()
                    If Me.Tag = "1" Then
                        ' set inspection viewed date as today
                        pInspection.CAEViewed = Today.Date
                        pInspection.ModifiedBy = MusterContainer.AppUser.ID
                        pInspection.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal, True)
                        If Not UIUtilsGen.HasRights(returnVal) Then
                            Exit Sub
                        End If

                        'refresh grid
                        'SetupTabs()
                        ' bug #2317
                        ' update row's viewed cell
                        ugRow.Cells("VIEWED").Value = Today.Date
                    End If
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnInspectionsExpandCollapseAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInspectionsExpandCollapseAll.Click
        Try
            If btnInspectionsExpandCollapseAll.Text = "Expand All" Then
                ExpandAll(True, ugInspections, btnInspectionsExpandCollapseAll)
            Else
                ExpandAll(False, ugInspections, btnInspectionsExpandCollapseAll)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugInspections_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugInspections.InitializeLayout
        Try

#If DEBUG Then
            e.Layout.Bands(0).Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.True
            e.Layout.Bands(0).Override.RowFilterMode = RowFilterMode.AllRowsInBand
            e.Layout.Bands(0).Override.RowFilterAction = RowFilterAction.HideFilteredOutRows
#End If

            e.Layout.Bands(0).Columns("SCORE").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Descending
            e.Layout.Bands(0).Columns("OWNERNAME").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            e.Layout.Bands(1).Columns("SCHEDULED").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            e.Layout.Bands(1).Columns("SUBMITTED").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            e.Layout.Bands(1).Columns("FACILITY").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending

            e.Layout.Bands(0).Override.RowAppearance.BackColor = Color.White
            e.Layout.Bands(1).Override.RowAppearance.BackColor = Color.RosyBrown
            e.Layout.Bands(2).Override.RowAppearance.BackColor = Color.Khaki

            e.Layout.Bands(0).Columns("ADDRESS_ID").Hidden = True
            e.Layout.Bands(0).Columns("OWNER_ID").Hidden = True

            If e.Layout.Bands(0).Summaries.Count = 0 Then

                e.Layout.Bands(0).Summaries.Add(Infragistics.Win.UltraWinGrid.SummaryType.Count, e.Layout.Bands(0).Columns("OWNERNAME"), Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn)
                e.Layout.Bands(0).Summaries.Add(Infragistics.Win.UltraWinGrid.SummaryType.Sum, e.Layout.Bands(0).Columns("FACILITIES"), Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn)
                e.Layout.Bands(0).Summaries.Add(Infragistics.Win.UltraWinGrid.SummaryType.Sum, e.Layout.Bands(0).Columns("SUBMITTED"), Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn)

            End If


            e.Layout.Bands(1).Columns("INSPECTION_ID").Hidden = True
            e.Layout.Bands(1).Columns("OWNER_ID").Hidden = True
            e.Layout.Bands(1).Columns("OWNERNAME").Hidden = True
            e.Layout.Bands(1).Columns("OWNADDRESS_LINE_ONE").Hidden = True
            e.Layout.Bands(1).Columns("OWNADDRESS_TWO").Hidden = True
            e.Layout.Bands(1).Columns("OWNCITY").Hidden = True
            e.Layout.Bands(1).Columns("OWNSTATE").Hidden = True
            e.Layout.Bands(1).Columns("OWNZIP").Hidden = True
            e.Layout.Bands(1).Columns("ADDRESS_LINE_ONE").Hidden = True
            e.Layout.Bands(1).Columns("ADDRESS_TWO").Hidden = True
            e.Layout.Bands(1).Columns("CITY").Hidden = True
            e.Layout.Bands(1).Columns("STATE").Hidden = True
            e.Layout.Bands(1).Columns("ZIP").Hidden = True
            'e.Layout.Bands(1).Columns("COUNTY").Hidden = True

            e.Layout.Bands(2).Columns("CHKLST_POSITION").Hidden = True
            e.Layout.Bands(2).Columns("INSPECTION_ID").Hidden = True
            e.Layout.Bands(2).Columns("FACILITY_ID").Hidden = True
            e.Layout.Bands(2).Columns("CITATION_ID").Hidden = True
            e.Layout.Bands(2).Columns("QUESTION_ID").Hidden = True
            e.Layout.Bands(2).Columns("INS_CIT_ID").Hidden = True

            e.Layout.Bands(0).Columns("SELECTED").CellActivation = Activation.AllowEdit
            e.Layout.Bands(0).Columns("OWNERNAME").CellActivation = Activation.NoEdit
            e.Layout.Bands(0).Columns("SCHEDULED").CellActivation = Activation.NoEdit
            e.Layout.Bands(0).Columns("SUBMITTED").CellActivation = Activation.NoEdit
            e.Layout.Bands(0).Columns("VIEWED").CellActivation = Activation.NoEdit
            e.Layout.Bands(0).Columns("VIEWED").Hidden = True
            e.Layout.Bands(0).Columns("FACILITIES").CellActivation = Activation.NoEdit
            e.Layout.Bands(0).Columns("DUE").CellActivation = Activation.NoEdit
            e.Layout.Bands(0).Columns("POTENTIAL").CellActivation = Activation.NoEdit
            e.Layout.Bands(0).Columns("INSPECTOR POTENTIAL").CellActivation = Activation.NoEdit
            e.Layout.Bands(0).Columns("SCORE").CellActivation = Activation.NoEdit

            'hide bad data
            e.Layout.Bands(0).Columns("INSPECTOR POTENTIAL").Hidden = True
            e.Layout.Bands(0).Columns("SCORE").Hidden = True
            e.Layout.Bands(1).Columns("NewOwner").Hidden = True
            e.Layout.Bands(1).Columns("Owner Status").CellActivation = Activation.NoEdit
            e.Layout.Bands(1).Columns("SELECTED").CellActivation = Activation.AllowEdit
            e.Layout.Bands(1).Columns("FACILITY_ID").CellActivation = Activation.NoEdit
            e.Layout.Bands(1).Columns("FACILITY").CellActivation = Activation.NoEdit
            e.Layout.Bands(1).Columns("SCHEDULED").CellActivation = Activation.NoEdit
            e.Layout.Bands(1).Columns("SUBMITTED").CellActivation = Activation.NoEdit
            e.Layout.Bands(1).Columns("VIEWED").CellActivation = Activation.NoEdit
            e.Layout.Bands(1).Columns("VIEWED").Hidden = True

            e.Layout.Bands(1).Columns("INSPECTOR").CellActivation = Activation.NoEdit
            e.Layout.Bands(1).Columns("OWNER NOT REG").CellActivation = Activation.NoEdit
            e.Layout.Bands(1).Columns("CITATIONS").CellActivation = Activation.NoEdit

            e.Layout.Bands(2).Columns("SELECTED").CellActivation = Activation.AllowEdit
            e.Layout.Bands(2).Columns("QUESTION#").CellActivation = Activation.NoEdit
            e.Layout.Bands(2).Columns("CITATION").CellActivation = Activation.NoEdit
            e.Layout.Bands(2).Columns("CITATIONTEXT").CellActivation = Activation.NoEdit
            e.Layout.Bands(2).Columns("CATEGORY").CellActivation = Activation.NoEdit
            e.Layout.Bands(2).Columns("CCAT").CellActivation = Activation.NoEdit
            e.Layout.Bands(2).Columns("CCAT COMMENTS").CellActivation = Activation.NoEdit

            e.Layout.Bands(2).Columns("CITATIONTEXT").Header.Caption = "CITATION TEXT"

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugInspections_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugInspections.CellChange
        If bolLoading Then Exit Sub
        Try
            bolLoading = True
            If "SELECTED".Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Band.Index = 0 Then
                    Dim facCount, facSelectedCount As Integer
                    facCount = 0
                    facSelectedCount = 0
                    For Each ugChildRow In e.Cell.Row.ChildBands(0).Rows ' facs
                        facCount += 1
                        If Not ugChildRow.Cells("SUBMITTED").Value Is DBNull.Value Then
                            facSelectedCount += 1
                            ugChildRow.Cells("SELECTED").Value = e.Cell.Text
                            For Each ugGrandChildRow In ugChildRow.ChildBands(0).Rows ' citations
                                ugGrandChildRow.Cells("SELECTED").Value = e.Cell.Text
                            Next
                        End If
                    Next
                    If e.Cell.Text.ToUpper = "TRUE" And facCount > 0 Then
                        If facSelectedCount = 0 Then
                            MsgBox("No Submitted Inspections")
                            e.Cell.CancelUpdate()
                        ElseIf facCount <> facSelectedCount Then
                            MsgBox("Not all Facilities Inspections have been submitted")
                            e.Cell.Value = e.Cell.Text
                        Else
                            e.Cell.Value = e.Cell.Text
                        End If
                    Else
                        e.Cell.Value = e.Cell.Text
                    End If
                ElseIf e.Cell.Row.Band.Index = 1 Then
                    If Not e.Cell.Row.Cells("SUBMITTED").Value Is DBNull.Value Then
                        For Each ugGrandChildRow In e.Cell.Row.ChildBands(0).Rows
                            ugGrandChildRow.Cells("SELECTED").Value = e.Cell.Text
                        Next
                        e.Cell.Value = e.Cell.Text
                    Else
                        MsgBox("Inspection not Submitted")
                        e.Cell.CancelUpdate()
                    End If
                ElseIf e.Cell.Row.Band.Index = 2 Then
                    If e.Cell.Row.ParentRow.Cells("SUBMITTED").Value Is DBNull.Value Then
                        MsgBox("Inspection not Submitted")
                        e.Cell.CancelUpdate()
                    Else
                        e.Cell.Row.ParentRow.Cells("SELECTED").Value = e.Cell.Text
                        For Each ugGrandChildRow In e.Cell.Row.ParentRow.ChildBands(0).Rows
                            ugGrandChildRow.Cells("SELECTED").Value = e.Cell.Text
                        Next
                    End If
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            bolLoading = False
        End Try
    End Sub
    Private Sub btnInspectionsRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInspectionsRefresh.Click
        SetupTabs()
    End Sub
#End Region

#Region "OCE Creating / Modification Logic"
    Private Sub OCECreationLogic(ByVal ug As Infragistics.Win.UltraWinGrid.UltraGridRow, ByVal fromCompliance As Boolean, Optional ByVal excludeOceID As Integer = 0)
        Dim bolAllCitCategoryDisc, bolHavePriorViolations, bolIsAnyCat1, bolIsAnyCat2a2b2c, bolContinue As Boolean
        Try
            bolAllCitCategoryDisc = True
            bolHavePriorViolations = False
            bolContinue = True
            ' Is the category of all citations for the owner = Discrepancy
            If fromCompliance Then
                For Each ugChildRow In ug.ChildBands(0).Rows ' facs
                    If ugChildRow.Cells("SELECTED").Value = True Or _
                        ugChildRow.Cells("SELECTED").Text = True Then
                        For Each ugGrandChildRow In ugChildRow.ChildBands(0).Rows ' citations
                            If ugGrandChildRow.Cells("CATEGORY").Value Is DBNull.Value Then
                                bolAllCitCategoryDisc = False
                                Exit For
                            Else
                                If ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper <> "DISCREPANCY" Then
                                    bolAllCitCategoryDisc = False
                                    Exit For
                                End If
                            End If
                        Next
                        If Not bolAllCitCategoryDisc Then Exit For
                    End If
                Next
            Else
                For Each ugGrandChildRow In ug.ChildBands(0).Rows ' citations
                    If ugGrandChildRow.Cells("SELECTED").Value = False Then
                        If ugGrandChildRow.Cells("CATEGORY").Value Is DBNull.Value Then
                            bolAllCitCategoryDisc = False
                            Exit For
                        Else
                            If ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper <> "DISCREPANCY" Then
                                bolAllCitCategoryDisc = False
                                Exit For
                            End If
                        End If
                    End If
                Next
            End If

            ' Does Owner have existing oce's with workshop date within 90 days prior and today and workshop result <> "No Show"
            If pOCE.OwnerHasWorkshopOCEDuringPast90Days(ug.Cells("OWNER_ID").Value, excludeOceID) Then
                bolContinue = False
                ViolationWithin90DaysSheet11(ug, fromCompliance, bolAllCitCategoryDisc, excludeOceID)
            End If

            If bolAllCitCategoryDisc And bolContinue Then
                DiscrepanciesOnlySheet2(ug, fromCompliance)
            ElseIf bolContinue Then
                Dim ds As DataSet = pOCE.OwnersPriorViolations(ug.Cells("OWNER_ID").Value, excludeOceID)
                ' does the owner have any prior violations
                If Not ds Is Nothing Then
                    If ds.Tables.Count > 0 Then
                        If ds.Tables(0).Rows.Count > 0 Then
                            bolHavePriorViolations = True
                        End If
                    End If
                End If
                If bolHavePriorViolations Then
                    PriorViolationsSheet5(ug, ds, fromCompliance, excludeOceID)
                Else
                    ' is the category of any citations for the owner = 1
                    If fromCompliance Then
                        For Each ugChildRow In ug.ChildBands(0).Rows ' facs
                            If ugChildRow.Cells("SELECTED").Value = True Or _
                                ugChildRow.Cells("SELECTED").Text = True Then
                                For Each ugGrandChildRow In ugChildRow.ChildBands(0).Rows ' citations
                                    If Not ugGrandChildRow.Cells("CATEGORY").Value Is DBNull.Value Then
                                        If ugGrandChildRow.Cells("CATEGORY").Value = "1" Then
                                            bolIsAnyCat1 = True
                                            Exit For
                                        End If
                                    End If
                                Next
                                If bolIsAnyCat1 Then Exit For
                            End If
                        Next
                    Else
                        For Each ugGrandChildRow In ug.ChildBands(0).Rows ' citations
                            If ugGrandChildRow.Cells("SELECTED").Value = False Then
                                If Not ugGrandChildRow.Cells("CATEGORY").Value Is DBNull.Value Then
                                    If ugGrandChildRow.Cells("CATEGORY").Value = "1" Then
                                        bolIsAnyCat1 = True
                                        Exit For
                                    End If
                                End If
                            End If
                        Next
                    End If
                    If bolIsAnyCat1 Then
                        Cat1NoPriorViolationsSheet7(ug, fromCompliance, False, excludeOceID)
                    Else
                        ' is the category of any citation for the owner = 2a / 2b / 2c
                        If fromCompliance Then
                            For Each ugChildRow In ug.ChildBands(0).Rows ' facs
                                If ugChildRow.Cells("SELECTED").Value = True Or _
                                    ugChildRow.Cells("SELECTED").Text = True Then
                                    For Each ugGrandChildRow In ugChildRow.ChildBands(0).Rows ' citations
                                        If Not ugGrandChildRow.Cells("CATEGORY").Value Is DBNull.Value Then
                                            If ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper = "2A" Or _
                                                ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper = "2B" Or _
                                                    ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper = "2C" Then
                                                bolIsAnyCat2a2b2c = True
                                                Exit For
                                            End If
                                        End If
                                    Next
                                    If bolIsAnyCat2a2b2c Then Exit For
                                End If
                            Next
                        Else
                            For Each ugGrandChildRow In ug.ChildBands(0).Rows ' citations
                                If ugGrandChildRow.Cells("SELECTED").Value = False Then
                                    If Not ugGrandChildRow.Cells("CATEGORY").Value Is DBNull.Value Then
                                        If ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper = "2A" Or _
                                            ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper = "2B" Or _
                                                ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper = "2C" Then
                                            bolIsAnyCat2a2b2c = True
                                            Exit For
                                        End If
                                    End If
                                End If
                            Next
                        End If
                        If bolIsAnyCat2a2b2c Then
                            Cat2NoPriorViolationsSheet4(ug, fromCompliance, False, excludeOceID)
                        Else
                            Cat3NoPriorViolationsSheet3(ug, fromCompliance)
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub DiscrepanciesOnlySheet2(ByVal ug As Infragistics.Win.UltraWinGrid.UltraGridRow, ByVal fromCompliance As Boolean)
        Dim strOwnerSize As String
        Dim policyPenalty As Integer
        Try
            policyPenalty = 0
            strOwnerSize = pOCE.GetOwnerSize(ug.Cells("OWNER_ID").Value)
            If strOwnerSize.Length > 0 Then
                If fromCompliance Then
                    For Each ugChildRow In ug.ChildBands(0).Rows ' facs
                        If ugChildRow.Cells("SELECTED").Value = True Or _
                            ugChildRow.Cells("SELECTED").Text = True Then
                            For Each ugGrandChildRow In ugChildRow.ChildBands(0).Rows
                                If Not ugGrandChildRow.Cells("CATEGORY").Value Is DBNull.Value Then
                                    If ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper = "DISCREPANCY" Then
                                        policyPenalty = ugGrandChildRow.Cells(strOwnerSize).Value
                                        Exit For
                                    End If
                                End If
                            Next
                            If policyPenalty <> 0 Then Exit For
                        End If
                    Next
                Else
                    For Each ugGrandChildRow In ug.ChildBands(0).Rows ' citations
                        If ugGrandChildRow.Cells("SELECTED").Value = False Then
                            If Not ugGrandChildRow.Cells("CATEGORY").Value Is DBNull.Value Then
                                If ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper = "DISCREPANCY" Then
                                    policyPenalty = ugGrandChildRow.Cells(strOwnerSize).Value
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                End If
                pOCE.PolicyAmount = policyPenalty
                pOCE.OCEPath = 1168 ' NOV(A)
                pOCE.OCEDate = Today.Date
                pOCE.NextDueDate = DateAdd(DateInterval.Day, 90, Today.Date)
                If pOCE.NextDueDate.DayOfWeek = DayOfWeek.Saturday Then
                    pOCE.NextDueDate = DateAdd(DateInterval.Day, 2, pOCE.NextDueDate)
                ElseIf pOCE.NextDueDate.DayOfWeek = DayOfWeek.Sunday Then
                    pOCE.NextDueDate = DateAdd(DateInterval.Day, 1, pOCE.NextDueDate)
                End If
                pOCE.OCEStatus = 1124 ' New
                pOCE.SettlementAmount = 0
                pOCE.WorkshopRequired = False
                pOCE.PendingLetter = 1172 ' Discrepancy
                pOCE.PendingLetterTemplateNum = UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.DiscrepanciesOnly)
            Else
                MsgBox("Invalid Owner Size")
                OCECreationError = True
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub PriorViolationsSheet5(ByVal ug As Infragistics.Win.UltraWinGrid.UltraGridRow, ByVal ds As DataSet, ByVal fromCompliance As Boolean, Optional ByVal excludeOceID As Integer = 0)
        Dim dr As DataRow
        Dim bolNoCat1Cat2Prior, bolStopCat1, bolStopCat2 As Boolean
        Dim nCat3PriorCount As Integer
        Try
            bolNoCat1Cat2Prior = True
            nCat3PriorCount = 0
            bolStopCat1 = False
            bolStopCat2 = False
            For Each dr In ds.Tables(0).Rows
                If Not dr("Category") Is DBNull.Value Then
                    If dr("Category") = "3" Then
                        nCat3PriorCount += 1
                    ElseIf dr("Category").ToString.StartsWith("1") Or dr("Category").ToString.StartsWith("2") Then
                        bolNoCat1Cat2Prior = False
                    End If
                End If
            Next

            If fromCompliance Then
                For Each ugChildRow In ug.ChildBands(0).Rows
                    If ugChildRow.Cells("SELECTED").Value = True Or _
                        ugChildRow.Cells("SELECTED").Text = True Then
                        For Each ugGrandChildRow In ugChildRow.ChildBands(0).Rows
                            If Not ugGrandChildRow.Cells("CATEGORY").Value Is DBNull.Value Then
                                If ugGrandChildRow.Cells("CATEGORY").Value.ToString = "1" Then
                                    bolStopCat1 = True
                                    Exit For
                                ElseIf ugGrandChildRow.Cells("CATEGORY").Value.ToString.StartsWith("2") Then
                                    bolStopCat2 = True
                                End If
                            End If
                        Next
                        If bolStopCat1 Then Exit For
                    End If
                Next
            Else
                For Each ugGrandChildRow In ug.ChildBands(0).Rows
                    If ugGrandChildRow.Cells("SELECTED").Value = False Then
                        If Not ugGrandChildRow.Cells("CATEGORY").Value Is DBNull.Value Then
                            If ugGrandChildRow.Cells("CATEGORY").Value.ToString = "1" Then
                                bolStopCat1 = True
                                Exit For
                            ElseIf ugGrandChildRow.Cells("CATEGORY").Value.ToString.StartsWith("2") Then
                                bolStopCat2 = True
                            End If
                        End If
                    End If
                Next
            End If
            If bolStopCat1 Then
                If bolNoCat1Cat2Prior Or nCat3PriorCount = 1 Then
                    Cat1PriorOneCat3Sheet7(ug, fromCompliance, excludeOceID)
                Else
                    Cat1PriorCat1Cat2OneCat3Sheet6(ug, fromCompliance)
                End If
            ElseIf bolStopCat2 Then
                If bolNoCat1Cat2Prior Or nCat3PriorCount = 1 Then
                    Cat2PriorOneCat3Sheet4(ug, fromCompliance, excludeOceID)
                Else
                    Cat2PriorCat1Cat2OneCat3Sheet8(ug, fromCompliance)
                End If
            Else
                If bolNoCat1Cat2Prior Or nCat3PriorCount = 1 Then
                    Cat3PriorOneCat3Sheet10(ug, fromCompliance, excludeOceID)
                Else
                    Cat3PriorCat1Cat2OneCat3Sheet9(ug, fromCompliance)
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub Cat1NoPriorViolationsSheet7(ByVal ug As Infragistics.Win.UltraWinGrid.UltraGridRow, ByVal fromCompliance As Boolean, ByVal isCat1_1Cat3 As Boolean, Optional ByVal excludeOceID As Integer = 0)
        Dim strOwnerSize As String
        Dim policyPenalty As Integer
        Try
            policyPenalty = 0
            strOwnerSize = pOCE.GetOwnerSize(ug.Cells("OWNER_ID").Value)
            If strOwnerSize.Length > 0 Then
                If fromCompliance Then
                    For Each ugChildRow In ug.ChildBands(0).Rows ' facs
                        If ugChildRow.Cells("SELECTED").Value = True Or _
                            ugChildRow.Cells("SELECTED").Text = True Then
                            For Each ugGrandChildRow In ugChildRow.ChildBands(0).Rows
                                If Not ugGrandChildRow.Cells("CATEGORY").Value Is DBNull.Value Then
                                    If ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper = "1" Then
                                        policyPenalty = ugGrandChildRow.Cells(strOwnerSize).Value
                                        Exit For
                                    End If
                                End If
                            Next
                            If policyPenalty <> 0 Then Exit For
                        End If
                    Next
                Else
                    For Each ugGrandChildRow In ug.ChildBands(0).Rows ' citations
                        If ugGrandChildRow.Cells("SELECTED").Value = False Then
                            If Not ugGrandChildRow.Cells("CATEGORY").Value Is DBNull.Value Then
                                If ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper = "1" Then
                                    policyPenalty = ugGrandChildRow.Cells(strOwnerSize).Value
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                End If
                If pOCE.OwnerHasPrevWorkshopDate(pOCE.OwnerID, False, excludeOceID) Then
                    Cat1PriorCat1Cat2OneCat3Sheet6(ug, fromCompliance)
                Else
                    pOCE.WorkshopRequired = True
                    pOCE.PolicyAmount = policyPenalty
                    pOCE.OCEPath = 1171 ' NOV + Workshop + Agreed Order(D)
                    pOCE.OCEDate = Today.Date
                    pOCE.NextDueDate = DateAdd(DateInterval.Day, 90, Today.Date)
                    If pOCE.NextDueDate.DayOfWeek = DayOfWeek.Saturday Then
                        pOCE.NextDueDate = DateAdd(DateInterval.Day, 2, pOCE.NextDueDate)
                    ElseIf pOCE.NextDueDate.DayOfWeek = DayOfWeek.Sunday Then
                        pOCE.NextDueDate = DateAdd(DateInterval.Day, 1, pOCE.NextDueDate)
                    End If
                    pOCE.OCEStatus = 1124 ' New
                    pOCE.SettlementAmount = policyPenalty / 2.0
                    pOCE.PendingLetter = 1176 ' NOV + Workshop + Agreed Order
                    If isCat1_1Cat3 Then
                        pOCE.PendingLetterTemplateNum = UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.CAT1_NoPrior_NOV_Workshop_AgreedOrder, isCat1_1Cat3)
                    Else
                        pOCE.PendingLetterTemplateNum = UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.CAT1_NoPrior_NOV_Workshop_AgreedOrder, isCat1_1Cat3)
                    End If
                End If
            Else
                MsgBox("Invalid Owner Size")
                OCECreationError = True
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub Cat2NoPriorViolationsSheet4(ByVal ug As Infragistics.Win.UltraWinGrid.UltraGridRow, ByVal fromCompliance As Boolean, ByVal isCat2_1Cat3 As Boolean, Optional ByVal excludeOceID As Integer = 0)
        Dim strOwnerSize As String
        Dim policyPenalty2A, policyPenalty2B, policyPenalty2C As Integer
        Try
            policyPenalty2A = 0
            policyPenalty2B = 0
            policyPenalty2C = 0
            strOwnerSize = pOCE.GetOwnerSize(ug.Cells("OWNER_ID").Value)
            If strOwnerSize.Length > 0 Then
                If fromCompliance Then
                    For Each ugChildRow In ug.ChildBands(0).Rows ' facs
                        If ugChildRow.Cells("SELECTED").Value = True Or _
                            ugChildRow.Cells("SELECTED").Text = True Then
                            For Each ugGrandChildRow In ugChildRow.ChildBands(0).Rows ' citations
                                If Not ugGrandChildRow.Cells("CATEGORY").Value Is DBNull.Value Then
                                    If ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper = "2A" Then
                                        policyPenalty2A = ugGrandChildRow.Cells(strOwnerSize).Value
                                    ElseIf ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper = "2B" Then
                                        policyPenalty2B = ugGrandChildRow.Cells(strOwnerSize).Value
                                    ElseIf ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper = "2C" Then
                                        policyPenalty2C = ugGrandChildRow.Cells(strOwnerSize).Value
                                    End If
                                End If
                            Next
                            If policyPenalty2A <> 0 And policyPenalty2B <> 0 And policyPenalty2C <> 0 Then Exit For
                        End If
                    Next
                Else
                    For Each ugGrandChildRow In ug.ChildBands(0).Rows ' citations
                        If ugGrandChildRow.Cells("SELECTED").Value = False Then
                            If Not ugGrandChildRow.Cells("CATEGORY").Value Is DBNull.Value Then
                                If ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper = "2A" Then
                                    policyPenalty2A = ugGrandChildRow.Cells(strOwnerSize).Value
                                ElseIf ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper = "2B" Then
                                    policyPenalty2B = ugGrandChildRow.Cells(strOwnerSize).Value
                                ElseIf ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper = "2C" Then
                                    policyPenalty2C = ugGrandChildRow.Cells(strOwnerSize).Value
                                End If
                            End If
                            If policyPenalty2A <> 0 And policyPenalty2B <> 0 And policyPenalty2C <> 0 Then Exit For
                        End If
                    Next
                End If
                If pOCE.OwnerHasPrevWorkshopDate(pOCE.OwnerID, False, excludeOceID) Then
                    Cat1PriorCat1Cat2OneCat3Sheet6(ug, fromCompliance)
                Else
                    pOCE.WorkshopRequired = True
                    pOCE.PolicyAmount = IIf(policyPenalty2A = 0, IIf(policyPenalty2B = 0, policyPenalty2C, policyPenalty2B), policyPenalty2A)
                    pOCE.OCEPath = 1169 ' NOV + Workshop (C)
                    pOCE.OCEDate = Today.Date
                    pOCE.NextDueDate = DateAdd(DateInterval.Day, 90, Today.Date)
                    If pOCE.NextDueDate.DayOfWeek = DayOfWeek.Saturday Then
                        pOCE.NextDueDate = DateAdd(DateInterval.Day, 2, pOCE.NextDueDate)
                    ElseIf pOCE.NextDueDate.DayOfWeek = DayOfWeek.Sunday Then
                        pOCE.NextDueDate = DateAdd(DateInterval.Day, 1, pOCE.NextDueDate)
                    End If
                    pOCE.OCEStatus = 1124 ' New

                    ' If workshop result = pass, settlement amount = $0 else 1/2 of Policy amount
                    If pOCE.WorkShopResult = 1011 Then
                        pOCE.SettlementAmount = 0
                    Else
                        pOCE.SettlementAmount = pOCE.PolicyAmount / 2.0
                    End If

                    pOCE.PendingLetter = 1174 ' NOV + Workshop
                    If isCat2_1Cat3 Then
                        pOCE.PendingLetterTemplateNum = UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.CAT2_1_CAT3_NOV_Workshop, isCat2_1Cat3)
                    Else
                        pOCE.PendingLetterTemplateNum = UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.CAT2_NoPrior_NOV_Workshop, isCat2_1Cat3)
                    End If
                End If
            Else
                MsgBox("Invalid Owner Size")
                OCECreationError = True
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub Cat3NoPriorViolationsSheet3(ByVal ug As Infragistics.Win.UltraWinGrid.UltraGridRow, ByVal fromCompliance As Boolean)
        Dim strOwnerSize As String
        Dim policyPenalty As Integer
        Try
            policyPenalty = 0
            strOwnerSize = pOCE.GetOwnerSize(ug.Cells("OWNER_ID").Value)
            If strOwnerSize.Length > 0 Then
                If fromCompliance Then
                    For Each ugChildRow In ug.ChildBands(0).Rows ' facs
                        If ugChildRow.Cells("SELECTED").Value = True Or _
                            ugChildRow.Cells("SELECTED").Text = True Then
                            For Each ugGrandChildRow In ugChildRow.ChildBands(0).Rows ' citations
                                If Not ugGrandChildRow.Cells("CATEGORY").Value Is DBNull.Value Then
                                    If ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper = "3" Then
                                        policyPenalty = ugGrandChildRow.Cells(strOwnerSize).Value
                                        Exit For
                                    End If
                                End If
                            Next
                            If policyPenalty <> 0 Then Exit For
                        End If
                    Next
                Else
                    For Each ugGrandChildRow In ug.ChildBands(0).Rows ' citations
                        If ugGrandChildRow.Cells("SELECTED").Value = False Then
                            If Not ugGrandChildRow.Cells("CATEGORY").Value Is DBNull.Value Then
                                If ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper = "3" Then
                                    policyPenalty = ugGrandChildRow.Cells(strOwnerSize).Value
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                End If
                pOCE.PolicyAmount = policyPenalty
                pOCE.OCEPath = 1168 ' NOV(A)
                pOCE.OCEDate = Today.Date
                pOCE.NextDueDate = DateAdd(DateInterval.Day, 90, Today.Date)
                If pOCE.NextDueDate.DayOfWeek = DayOfWeek.Saturday Then
                    pOCE.NextDueDate = DateAdd(DateInterval.Day, 2, pOCE.NextDueDate)
                ElseIf pOCE.NextDueDate.DayOfWeek = DayOfWeek.Sunday Then
                    pOCE.NextDueDate = DateAdd(DateInterval.Day, 1, pOCE.NextDueDate)
                End If
                pOCE.OCEStatus = 1124 ' New
                pOCE.SettlementAmount = 0
                pOCE.WorkshopRequired = False
                pOCE.PendingLetter = 1173 ' NOV
                pOCE.PendingLetterTemplateNum = UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.CAT3_NoPrior_NOV)
            Else
                MsgBox("Invalid Owner Size")
                OCECreationError = True
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub ChangeTOSItoTOS(ByVal ownerID As Integer, ByVal facID As Integer)
        Dim pOwn As New MUSTER.BusinessLogic.pOwner
        Dim pTank As MUSTER.BusinessLogic.pTank
        Dim pPipe As MUSTER.BusinessLogic.pPipe
        Try
            pOwn.Facilities.RetrieveAll(ownerID, , False, facID, )
            'pOwn.RetrieveAll(ownerID, , False, facID, )
            pTank = pOwn.Facilities.FacilityTanks
            For Each tnk As MUSTER.Info.TankInfo In pOwn.Facility.TankCollection.Values
                pTank.TankInfo = tnk
                If pTank.TankStatus = 429 Then ' TOSI
                    ' pTank.TankStatus = 425 ' TOS
                    pTank.ModifiedBy = MusterContainer.AppUser.ID
                    pTank.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID, True, False, False)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If
                End If
                pPipe = pTank.Pipes
                For Each pipe As MUSTER.Info.PipeInfo In tnk.pipesCollection.Values
                    pPipe.Pipe = pipe
                    If pPipe.PipeStatusDesc = 429 Then ' TOSI

                        ' pPipe.PipeStatusDesc = 425 ' TOS
                        If pPipe.PipeID <= 0 Then
                            pPipe.CreatedBy = MusterContainer.AppUser.ID
                        Else
                            pPipe.ModifiedBy = MusterContainer.AppUser.ID
                        End If
                        pPipe.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID, True, False)
                        If Not UIUtilsGen.HasRights(returnVal) Then
                            Exit Sub
                        End If

                    End If
                Next
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub ChangeTOStoTOSI(ByVal ownerID As Integer, ByVal facID As Integer, Optional ByVal excludeOceID As String = "")
        Dim pOwn As New MUSTER.BusinessLogic.pOwner
        Dim pTank As MUSTER.BusinessLogic.pTank
        Dim pPipe As MUSTER.BusinessLogic.pPipe
        Try
            pOwn.Facilities.RetrieveAll(ownerID, , False, facID, )
            'pOwn.RetrieveAll(ownerID, , False, facID, )
            pTank = pOwn.Facilities.FacilityTanks
            For Each tnk As MUSTER.Info.TankInfo In pOwn.Facility.TankCollection.Values
                pTank.TankInfo = tnk
                If pTank.TankStatus = 425 Then ' TOS
                    '  pTank.TankStatus = 429 ' TOSI
                    pTank.ModifiedBy = MusterContainer.AppUser.ID
                    pTank.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID, True, False, False, excludeOceID.ToString)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If
                End If
                pPipe = pTank.Pipes
                For Each pipe As MUSTER.Info.PipeInfo In tnk.pipesCollection.Values
                    pPipe.Pipe = pipe
                    If pPipe.PipeStatusDesc = 425 Then ' TOS
                        '    pPipe.PipeStatusDesc = 429 ' TOSI
                        If pPipe.PipeID <= 0 Then
                            pPipe.CreatedBy = MusterContainer.AppUser.ID
                        Else
                            pPipe.ModifiedBy = MusterContainer.AppUser.ID
                        End If
                        pPipe.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID, True, False, excludeOceID.ToString)
                        If Not UIUtilsGen.HasRights(returnVal) Then
                            Exit Sub
                        End If
                    End If
                Next
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub ViolationWithin90DaysSheet11(ByVal ug As Infragistics.Win.UltraWinGrid.UltraGridRow, ByVal fromCompliance As Boolean, ByVal bolAllCitCategoryDisc As Boolean, Optional ByVal excludeOceID As Integer = 0)
        Dim strOwnerSize As String
        Try
            strOwnerSize = pOCE.GetOwnerSize(ug.Cells("OWNER_ID").Value)
            If strOwnerSize.Length > 0 Then
                If bolAllCitCategoryDisc Then
                    pOCE.PolicyAmount = pOCE.GetPenaltyByOwnerSizeCitCategory(strOwnerSize, "Discrepancy")
                    pOCE.OCEPath = 1168 ' NOV(A)
                    pOCE.OCEDate = Today.Date
                    pOCE.NextDueDate = DateAdd(DateInterval.Day, 90, Today.Date)
                    If pOCE.NextDueDate.DayOfWeek = DayOfWeek.Saturday Then
                        pOCE.NextDueDate = DateAdd(DateInterval.Day, 2, pOCE.NextDueDate)
                    ElseIf pOCE.NextDueDate.DayOfWeek = DayOfWeek.Sunday Then
                        pOCE.NextDueDate = DateAdd(DateInterval.Day, 1, pOCE.NextDueDate)
                    End If
                    pOCE.OCEStatus = 1124 ' New
                    pOCE.SettlementAmount = 0
                    pOCE.WorkshopRequired = False
                    pOCE.PendingLetter = 1172 ' Discrepancy
                    pOCE.PendingLetterTemplateNum = UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.ViolationWithin90Days_Discrepancy_WhenOwnerHasDiscrepCitationsOnly)
                Else
                    pOCE.PolicyAmount = pOCE.GetPenaltyByOwnerSizeCitCategory(strOwnerSize, "1")
                    pOCE.OCEPath = 1182 ' NOV + Agreed Order
                    pOCE.OCEDate = Today.Date
                    pOCE.NextDueDate = DateAdd(DateInterval.Day, 90, Today.Date)
                    If pOCE.NextDueDate.DayOfWeek = DayOfWeek.Saturday Then
                        pOCE.NextDueDate = DateAdd(DateInterval.Day, 2, pOCE.NextDueDate)
                    ElseIf pOCE.NextDueDate.DayOfWeek = DayOfWeek.Sunday Then
                        pOCE.NextDueDate = DateAdd(DateInterval.Day, 1, pOCE.NextDueDate)
                    End If
                    pOCE.OCEStatus = 1124 ' New
                    pOCE.SettlementAmount = pOCE.PolicyAmount / 2.0
                    pOCE.OverRideAmount = 0
                    pOCE.WorkshopRequired = False
                    pOCE.PendingLetter = 1172 ' Discrepancy
                    pOCE.PendingLetterTemplateNum = UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.ViolationWithin90Days_Discrepancy_WhenOwnerHasNonDiscrepCitation)
                End If
            Else
                MsgBox("Invalid Owner Size")
                OCECreationError = True
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    ' Begin branch off from Prior Violations
    Private Sub Cat1PriorOneCat3Sheet7(ByVal ug As Infragistics.Win.UltraWinGrid.UltraGridRow, ByVal fromCompliance As Boolean, Optional ByVal excludeOceID As Integer = 0)
        Cat1NoPriorViolationsSheet7(ug, fromCompliance, True, excludeOceID)
    End Sub
    Private Sub Cat1PriorCat1Cat2OneCat3Sheet6(ByVal ug As Infragistics.Win.UltraWinGrid.UltraGridRow, ByVal fromCompliance As Boolean)
        Dim strOwnerSize As String
        Dim policyPenalty As Integer
        Try
            policyPenalty = 0
            strOwnerSize = pOCE.GetOwnerSize(ug.Cells("OWNER_ID").Value)
            If strOwnerSize.Length > 0 Then
                If fromCompliance Then
                    For Each ugChildRow In ug.ChildBands(0).Rows ' facs
                        If ugChildRow.Cells("SELECTED").Value = True Or _
                            ugChildRow.Cells("SELECTED").Text = True Then
                            For Each ugGrandChildRow In ugChildRow.ChildBands(0).Rows ' citations
                                If Not ugGrandChildRow.Cells("CATEGORY").Value Is DBNull.Value Then
                                    If ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper = "1" Then
                                        policyPenalty = ugGrandChildRow.Cells(strOwnerSize).Value
                                        Exit For
                                    End If
                                End If
                            Next
                            If policyPenalty <> 0 Then Exit For
                        End If
                    Next
                Else
                    For Each ugGrandChildRow In ug.ChildBands(0).Rows ' citations
                        If ugGrandChildRow.Cells("SELECTED").Value = False Then
                            If Not ugGrandChildRow.Cells("CATEGORY").Value Is DBNull.Value Then
                                If ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper = "1" Then
                                    policyPenalty = ugGrandChildRow.Cells(strOwnerSize).Value
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                End If

                If policyPenalty = 0 Then
                    policyPenalty = GetPenaltyforCategory(strOwnerSize, "1")
                End If

                pOCE.PolicyAmount = policyPenalty
                pOCE.OCEPath = 1170 ' NOV + Agreed Order (B)
                pOCE.OCEDate = Today.Date
                pOCE.NextDueDate = DateAdd(DateInterval.Day, 90, Today.Date)
                If pOCE.NextDueDate.DayOfWeek = DayOfWeek.Saturday Then
                    pOCE.NextDueDate = DateAdd(DateInterval.Day, 2, pOCE.NextDueDate)
                ElseIf pOCE.NextDueDate.DayOfWeek = DayOfWeek.Sunday Then
                    pOCE.NextDueDate = DateAdd(DateInterval.Day, 1, pOCE.NextDueDate)
                End If
                pOCE.OCEStatus = 1124 ' New
                pOCE.SettlementAmount = pOCE.PolicyAmount / 2.0
                pOCE.WorkshopRequired = False
                pOCE.PendingLetter = 1175 ' NOV + Agreed Order
                pOCE.PendingLetterTemplateNum = UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.CAT1_CAT1_CAT2_1_CAT3_NOV_AgreedOrder)
            Else
                MsgBox("Invalid Owner Size")
                OCECreationError = True
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub Cat2PriorOneCat3Sheet4(ByVal ug As Infragistics.Win.UltraWinGrid.UltraGridRow, ByVal fromCompliance As Boolean, Optional ByVal excludeOceID As Integer = 0)
        Cat2NoPriorViolationsSheet4(ug, fromCompliance, True, excludeOceID)
    End Sub
    Private Sub Cat2PriorCat1Cat2OneCat3Sheet8(ByVal ug As Infragistics.Win.UltraWinGrid.UltraGridRow, ByVal fromCompliance As Boolean)
        Dim strOwnerSize As String
        Dim policyPenalty2A, policyPenalty2B, policyPenalty2C As Integer
        Try
            policyPenalty2A = 0
            policyPenalty2B = 0
            policyPenalty2C = 0
            strOwnerSize = pOCE.GetOwnerSize(ug.Cells("OWNER_ID").Value)
            If strOwnerSize.Length > 0 Then
                If fromCompliance Then
                    For Each ugChildRow In ug.ChildBands(0).Rows ' facs
                        If ugChildRow.Cells("SELECTED").Value = True Or _
                            ugChildRow.Cells("SELECTED").Text = True Then
                            For Each ugGrandChildRow In ugChildRow.ChildBands(0).Rows ' citations
                                If Not ugGrandChildRow.Cells("CATEGORY").Value Is DBNull.Value Then
                                    If ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper = "2A" Then
                                        policyPenalty2A = ugGrandChildRow.Cells(strOwnerSize).Value
                                    ElseIf ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper = "2B" Then
                                        policyPenalty2B = ugGrandChildRow.Cells(strOwnerSize).Value
                                    ElseIf ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper = "2C" Then
                                        policyPenalty2C = ugGrandChildRow.Cells(strOwnerSize).Value
                                    End If
                                End If
                            Next
                            If policyPenalty2A <> 0 And policyPenalty2B <> 0 And policyPenalty2C <> 0 Then Exit For
                        End If
                    Next
                Else
                    For Each ugGrandChildRow In ug.ChildBands(0).Rows ' citations
                        If ugGrandChildRow.Cells("SELECTED").Value = False Then
                            If Not ugGrandChildRow.Cells("CATEGORY").Value Is DBNull.Value Then
                                If ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper = "2A" Then
                                    policyPenalty2A = ugGrandChildRow.Cells(strOwnerSize).Value
                                ElseIf ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper = "2B" Then
                                    policyPenalty2B = ugGrandChildRow.Cells(strOwnerSize).Value
                                ElseIf ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper = "2C" Then
                                    policyPenalty2C = ugGrandChildRow.Cells(strOwnerSize).Value
                                End If
                            End If
                            If policyPenalty2A <> 0 And policyPenalty2B <> 0 And policyPenalty2C <> 0 Then Exit For
                        End If
                    Next
                End If
                pOCE.PolicyAmount = IIf(policyPenalty2A = 0, IIf(policyPenalty2B = 0, policyPenalty2C, policyPenalty2B), policyPenalty2A)
                pOCE.OCEPath = 1170 ' NOV + Agreed Order (B)
                pOCE.OCEDate = Today.Date
                pOCE.NextDueDate = DateAdd(DateInterval.Day, 90, Today.Date)
                If pOCE.NextDueDate.DayOfWeek = DayOfWeek.Saturday Then
                    pOCE.NextDueDate = DateAdd(DateInterval.Day, 2, pOCE.NextDueDate)
                ElseIf pOCE.NextDueDate.DayOfWeek = DayOfWeek.Sunday Then
                    pOCE.NextDueDate = DateAdd(DateInterval.Day, 1, pOCE.NextDueDate)
                End If
                pOCE.OCEStatus = 1124 ' New
                pOCE.SettlementAmount = pOCE.PolicyAmount / 2.0
                pOCE.WorkshopRequired = False
                pOCE.PendingLetter = 1175 ' NOV + Agreed Order
                pOCE.PendingLetterTemplateNum = UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.CAT2_CAT1_CAT2_1_CAT3_NOV_AgreedOrder)
            Else
                MsgBox("Invalid Owner Size")
                OCECreationError = True
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub Cat3PriorOneCat3Sheet10(ByVal ug As Infragistics.Win.UltraWinGrid.UltraGridRow, ByVal fromCompliance As Boolean, Optional ByVal excludeOceID As Integer = 0)
        Dim strOwnerSize As String
        Dim policyPenalty As Integer
        Try
            policyPenalty = 0
            strOwnerSize = pOCE.GetOwnerSize(ug.Cells("OWNER_ID").Value)
            If strOwnerSize.Length > 0 Then
                If fromCompliance Then
                    For Each ugChildRow In ug.ChildBands(0).Rows ' facs
                        If ugChildRow.Cells("SELECTED").Value = True Or _
                            ugChildRow.Cells("SELECTED").Text = True Then
                            For Each ugGrandChildRow In ugChildRow.ChildBands(0).Rows ' citations
                                If Not ugGrandChildRow.Cells("CATEGORY").Value Is DBNull.Value Then
                                    If ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper = "3" Then
                                        policyPenalty = ugGrandChildRow.Cells(strOwnerSize).Value
                                        Exit For
                                    End If
                                End If
                            Next
                            If policyPenalty <> 0 Then Exit For
                        End If
                    Next
                Else
                    For Each ugGrandChildRow In ug.ChildBands(0).Rows ' citations
                        If ugGrandChildRow.Cells("SELECTED").Value = False Then
                            If Not ugGrandChildRow.Cells("CATEGORY").Value Is DBNull.Value Then
                                If ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper = "3" Then
                                    policyPenalty = ugGrandChildRow.Cells(strOwnerSize).Value
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                End If
                If pOCE.OwnerHasPrevWorkshopDate(pOCE.OwnerID, False, excludeOceID) Then
                    Cat1PriorCat1Cat2OneCat3Sheet6(ug, fromCompliance)
                Else
                    pOCE.WorkshopRequired = True
                    pOCE.PolicyAmount = policyPenalty
                    pOCE.OCEPath = 1169 ' NOV + Workshop (C)
                    pOCE.OCEDate = Today.Date
                    pOCE.NextDueDate = DateAdd(DateInterval.Day, 90, Today.Date)
                    If pOCE.NextDueDate.DayOfWeek = DayOfWeek.Saturday Then
                        pOCE.NextDueDate = DateAdd(DateInterval.Day, 2, pOCE.NextDueDate)
                    ElseIf pOCE.NextDueDate.DayOfWeek = DayOfWeek.Sunday Then
                        pOCE.NextDueDate = DateAdd(DateInterval.Day, 1, pOCE.NextDueDate)
                    End If
                    pOCE.OCEStatus = 1124 ' New
                    ' If workshop result = pass, settlement amount = $0 else 1/2 of Policy amount
                    If pOCE.WorkShopResult = 1011 Then
                        pOCE.SettlementAmount = 0
                    Else
                        pOCE.SettlementAmount = pOCE.PolicyAmount / 2.0
                    End If
                    pOCE.PendingLetter = 1174 ' NOV + Workshop
                    pOCE.PendingLetterTemplateNum = UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.CAT3_1_CAT3_NOV_Workshop)
                End If
            Else
                MsgBox("Invalid Owner Size")
                OCECreationError = True
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub Cat3PriorCat1Cat2OneCat3Sheet9(ByVal ug As Infragistics.Win.UltraWinGrid.UltraGridRow, ByVal fromCompliance As Boolean)
        Dim strOwnerSize As String
        Dim policyPenalty As Integer
        Try
            policyPenalty = 0
            strOwnerSize = pOCE.GetOwnerSize(ug.Cells("OWNER_ID").Value)
            If strOwnerSize.Length > 0 Then
                If fromCompliance Then
                    For Each ugChildRow In ug.ChildBands(0).Rows ' facs
                        If ugChildRow.Cells("SELECTED").Value = True Or _
                            ugChildRow.Cells("SELECTED").Text = True Then
                            For Each ugGrandChildRow In ugChildRow.ChildBands(0).Rows ' citations
                                If Not ugGrandChildRow.Cells("CATEGORY").Value Is DBNull.Value Then
                                    If ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper = "3" Then
                                        policyPenalty = ugGrandChildRow.Cells(strOwnerSize).Value
                                        Exit For
                                    End If
                                End If
                            Next
                            If policyPenalty <> 0 Then Exit For
                        End If
                    Next
                Else
                    For Each ugGrandChildRow In ug.ChildBands(0).Rows ' citations
                        If ugGrandChildRow.Cells("SELECTED").Value = False Then
                            If Not ugGrandChildRow.Cells("CATEGORY").Value Is DBNull.Value Then
                                If ugGrandChildRow.Cells("CATEGORY").Value.ToString.ToUpper = "3" Then
                                    policyPenalty = ugGrandChildRow.Cells(strOwnerSize).Value
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                End If
                pOCE.PolicyAmount = policyPenalty
                pOCE.OCEPath = 1170 ' NOV + Agreed Order (B)
                pOCE.OCEDate = Today.Date
                pOCE.NextDueDate = DateAdd(DateInterval.Day, 90, Today.Date)
                If pOCE.NextDueDate.DayOfWeek = DayOfWeek.Saturday Then
                    pOCE.NextDueDate = DateAdd(DateInterval.Day, 2, pOCE.NextDueDate)
                ElseIf pOCE.NextDueDate.DayOfWeek = DayOfWeek.Sunday Then
                    pOCE.NextDueDate = DateAdd(DateInterval.Day, 1, pOCE.NextDueDate)
                End If
                pOCE.OCEStatus = 1124 ' New
                pOCE.SettlementAmount = pOCE.PolicyAmount / 2.0
                pOCE.WorkshopRequired = False
                pOCE.PendingLetter = 1175 ' NOV + Agreed Order
                pOCE.PendingLetterTemplateNum = UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.CAT3_CAT1_CAT2_1_CAT3_NOV_AgreedOrder)
            Else
                MsgBox("Invalid Owner Size")
                OCECreationError = True
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Function GetPenaltyforCategory(ByVal strOwnerSize As String, ByVal strCategory As String) As Integer
        Dim returnVal As Integer = 0
        Dim ds As DataSet = pOCE.GetCitationCategoryPolicyPenalties()
        Try
            If Not ds Is Nothing Then
                If ds.Tables.Count > 0 Then
                    If ds.Tables(0).Rows.Count > 0 Then
                        For Each dr As DataRow In ds.Tables(0).Rows
                            If dr("CATEGORY") = strCategory Then
                                returnVal = dr(strOwnerSize)
                            End If
                            If returnVal <> 0 Then Exit For
                        Next
                    End If
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
        Return returnVal
    End Function
    ' End branch off from Prior Violations
#End Region

#Region "Compliance"
    'Private Sub btnCompGenerateOCEs_ClickOld(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCompGenerateOCEs.Click
    '    Dim rowCol As Infragistics.Win.UltraWinGrid.RowsCollection
    '    Dim strOwners As New ArrayList
    '    Dim strFCE As New ArrayList
    '    Try
    '        For Each dr As Infragistics.Win.UltraWinGrid.UltraGridRow In ugCompliance.Rows
    '            If dr.Cells("Selected").Value Then
    '                'strOwners.Add(dr.Cells("OWNER_ID").Value)
    '                rowCol = dr.ChildBands("OwnertoFCE".ToUpper).Rows
    '                ' loop through the FCE's
    '                For Each dr1 As Infragistics.Win.UltraWinGrid.UltraGridRow In rowCol
    '                    If dr1.Cells("Selected").Value Then
    '                        strFCE.Add(dr1.Cells("FCEID").Value)
    '                        'pOCE.CreationLogic(dr1.Cells("FCEID").Value)
    '                        pFCE.Retrieve(dr1.Cells("FCEID").Value)
    '                        pFCE.OCEGenerated = True
    '                        pFCE.Save()
    '                    End If
    '                Next
    '                'pOCE.OCECreationLogic(dr.Cells("OWNER_ID").Value, strFCE)
    '            End If
    '        Next
    '        Dim bolLoading As Boolean = True
    '        For Each oceinfo As MUSTER.Info.OwnerComplianceEventInfo In pOCE.OCECollection.Values
    '            If oceinfo.WorkshopRequired = True Then
    '                If bolLoading Then
    '                    frmWorkshopDate.ShowDialog()
    '                    bolLoading = False
    '                End If
    '                oceinfo.WorkShopDate = frmWorkshopDate.WorkshopDate
    '            End If
    '        Next
    '        pOCE.Flush()
    '        MsgBox("OCEs generated successfully")
    '        tabCntrlCandE.SelectedTab = tbPageCompliance
    '        SetupTabs()
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    Private Sub btnCompGenerateOCEs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCompGenerateOCEs.Click
        If bolLoading Then Exit Sub
        Dim bolOCECreated, bolOCEGenerated As Boolean
        Dim alChangeTosiToTos As New ArrayList
        Dim alFCEIDs As New ArrayList
        Try
            If Not MusterContainer.AppUser.HasAccess(UIUtilsGen.ModuleID.CAE, MusterContainer.AppUser.UserKey, UIUtilsGen.EntityTypes.CAEOwnerComplianceEvent) Then
                MessageBox.Show("You do not have rights to save Owner Compliance Event.", "ACCESS DENIED", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            bolLoading = True
            OCECreationError = False
            bolOCEGenerated = False
            ' reset date to null
            dtWorkshopDate = CDate("01/01/0001")
            ' Scenario 2 - 3
            For Each ugRow In ugCompliance.Rows ' owner
                bolOCECreated = False
                alFCEIDs = New ArrayList
                For Each ugChildRow In ugRow.ChildBands(0).Rows ' facs
                    If ugChildRow.Cells("SELECTED").Value = True Or _
                        ugChildRow.Cells("SELECTED").Text = True Then
                        ' Scenarion 2 - 3.a
                        If Not bolOCECreated Then
                            pOCE.Retrieve(0)
                            pOCE.OwnerID = ugRow.Cells("OWNER_ID").Value
                            bolOCECreated = True
                        End If
                        If Not alFCEIDs.Contains(ugChildRow.Cells("FCE_ID").Value) Then alFCEIDs.Add(ugChildRow.Cells("FCE_ID").Value)
                        ' Scenario 2 - 3.d
                        ' If FCE contain Citation "280.20, 280.21 & 280.31" (CITATION_ID = 10) change status of all tanks and pipes that are tosi to tos
                        For Each ugGrandChildRow In ugChildRow.ChildBands(0).Rows
                            If Not ugGrandChildRow.Cells("CITATION_ID").Value Is DBNull.Value Then
                                If ugGrandChildRow.Cells("CITATION_ID").Value = 10 Then
                                    If Not alChangeTosiToTos.Contains(ugChildRow.Cells("FACILITY_ID").Value) Then
                                        alChangeTosiToTos.Add(ugChildRow.Cells("FACILITY_ID").Value)
                                    End If
                                    'ChangeTOSItoTOS(ugRow.Cells("OWNER_ID").Value, ugChildRow.Cells("FACILITY_ID").Value)
                                    Exit For
                                End If
                            End If
                        Next
                    End If
                Next
                If bolOCECreated Then
                    ' Execute OCE Creation logic
                    OCECreationLogic(ugRow, True)
                    If OCECreationError Then
                        OCECreationError = False
                        Exit Sub
                    End If
                    ' Scenario 2 - 3.b
                    If pOCE.WorkshopRequired Then
                        If Date.Compare(dtWorkshopDate, CDate("01/01/0001")) = 0 Then
                            frmWorkshopDate = New WorkShopDate
                            frmWorkshopDate.ShowDialog()
                            pOCE.WorkShopDate = frmWorkshopDate.WorkshopDate
                            dtWorkshopDate = frmWorkshopDate.WorkshopDate
                        Else
                            pOCE.WorkShopDate = dtWorkshopDate
                        End If
                    Else
                        pOCE.WorkShopDate = CDate("01/01/0001")
                    End If
                    If pOCE.ID <= 0 Then
                        pOCE.CreatedBy = MusterContainer.AppUser.ID
                    Else
                        pOCE.ModifiedBy = MusterContainer.AppUser.ID
                    End If
                    pOCE.Escalation = pOCE.OCEStatus

                    pOCE.Save(0, CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal, True)
                    For i As Integer = 0 To alFCEIDs.Count - 1
                        pFCE.Retrieve(alFCEIDs.Item(i))
                        pFCE.OCEGenerated = True
                        pFCE.OCEID = pOCE.ID
                        If pFCE.ID <= 0 Then
                            pFCE.CreatedBy = MusterContainer.AppUser.ID
                        Else
                            pFCE.ModifiedBy = MusterContainer.AppUser.ID
                        End If

                        pFCE.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal, True)
                        If Not UIUtilsGen.HasRights(returnVal) Then
                            Exit Sub
                        End If
                    Next
                    'pFCE.Retrieve(ugChildRow.Cells("FCE_ID").Value)
                    'pFCE.OCEGenerated = True
                    'If pFCE.ID <= 0 Then
                    '    pFCE.CreatedBy = MusterContainer.AppUser.ID
                    'Else
                    '    pFCE.ModifiedBy = MusterContainer.AppUser.ID
                    'End If
                    'pFCE.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal, True)
                    'If Not UIUtilsGen.HasRights(returnVal) Then
                    '    Exit Sub
                    'End If

                    ' add OCE_ID and Due Date to Citations
                    For Each ugChildRow In ugRow.ChildBands(0).Rows
                        If ugChildRow.Cells("SELECTED").Value = True Or _
                            ugChildRow.Cells("SELECTED").Text = True Then
                            For Each ugGrandChildRow In ugChildRow.ChildBands(0).Rows
                                pInsCitation.Retrieve(New MUSTER.Info.InspectionInfo, ugGrandChildRow.Cells("INS_CIT_ID").Value)
                                pInsCitation.OCEID = pOCE.ID
                                pInsCitation.CitationDueDate = pOCE.NextDueDate

                                If pInsCitation.ID <= 0 Then
                                    pInsCitation.CreatedBy = MusterContainer.AppUser.ID
                                Else
                                    pInsCitation.ModifiedBy = MusterContainer.AppUser.ID
                                End If

                                pInsCitation.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal, True)
                                If Not UIUtilsGen.HasRights(returnVal) Then
                                    Exit Sub
                                End If
                            Next
                        End If
                    Next
                    bolOCEGenerated = True
                    ' change tosi tanks/pipes to tos
                    If alChangeTosiToTos.Count > 0 Then
                        For i As Integer = 0 To alChangeTosiToTos.Count - 1
                            ChangeTOSItoTOS(ugRow.Cells("OWNER_ID").Value, alChangeTosiToTos.Item(i))
                        Next
                    End If
                End If
            Next
            ' reset date to null
            dtWorkshopDate = CDate("01/01/0001")
            If bolOCEGenerated Then
                MsgBox("OCE(s) generated successfully")
                SetupTabs()
            Else
                MsgBox("No OCE(s) created")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            bolLoading = False
        End Try
    End Sub
    Private Sub btnCompAddFCE_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCompAddFCE.Click
        Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim strOwnerName As String = String.Empty
        Dim nOwnerID As Integer = 0
        Dim nFacilityID As Integer = 0
        Dim nInspID As Integer = 0
        Dim nFCEID As Integer = 0
        Try
            If Not MusterContainer.AppUser.HasAccess(UIUtilsGen.ModuleID.CAE, MusterContainer.AppUser.UserKey, UIUtilsGen.EntityTypes.CAEFacilityCompliantEvent) Then
                MessageBox.Show("You do not have rights to save Facility Compliance Event.", "ACCESS DENIED", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            ugRow = getSelectedRow()
            If Not ugRow Is Nothing Then
                If ugRow.Cells("SELECTED").Value = True Or _
                    ugRow.Cells("SELECTED").Text = True Then
                    nOwnerID = ugRow.Cells("OWNER_ID").Value
                    If ugRow.Band.Index = 1 Then
                        nFacilityID = ugRow.Cells("FACILITY_ID").Value
                        ' Since its adding a new FCE, no point in sending the fceid / inspectionid
                        'nFCEID = ugRow.Cells("FCE_ID").Value
                        'nInspID = ugRow.Cells("INSPECTION_ID").Value
                        strOwnerName = ugRow.ParentRow.Cells("OWNERNAME").Value
                    Else
                        strOwnerName = ugRow.Cells("OWNERNAME").Value
                    End If
                End If
            End If
            Dim frmFCE As New FacilityComplianceEvent(nOwnerID, nInspID, nFacilityID, nFCEID, strOwnerName)
            frmFCE.CallingForm = Me
            Me.Tag = "0"
            frmFCE.ShowDialog()
            If Me.Tag = "1" Then
                ' refresh grid
                SetupTabs()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCompEditFCE_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCompEditFCE.Click
        Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim strOwnerName As String = String.Empty
        Dim nOwnerID As Integer = 0
        Dim nFacilityID As Integer = 0
        Dim nInspID As Integer
        Dim nFCEID As Integer = 0
        Try
            If Not MusterContainer.AppUser.HasAccess(UIUtilsGen.ModuleID.CAE, MusterContainer.AppUser.UserKey, UIUtilsGen.EntityTypes.CAEFacilityCompliantEvent) Then
                MessageBox.Show("You do not have rights to save Facility Compliance Event.", "ACCESS DENIED", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            ugRow = getSelectedRow()
            If Not ugRow Is Nothing Then
                If ugRow.Band.Index = 1 Then
                    If ugRow.Cells("SOURCE").Value = "ADMIN" Then
                        nOwnerID = ugRow.Cells("OWNER_ID").Value
                        nFacilityID = ugRow.Cells("FACILITY_ID").Value
                        nFCEID = ugRow.Cells("FCE_ID").Value
                        nInspID = ugRow.Cells("INSPECTION_ID").Value
                        strOwnerName = ugRow.ParentRow.Cells("OWNERNAME").Value
                    Else
                        MsgBox("Cannot edit FCE Generated from " + ugRow.Cells("SOURCE").Text + vbCrLf + _
                                "Please select FCE with Source 'ADMIN'")
                        Exit Sub
                    End If
                Else
                    MsgBox("Please select FCE with Source 'ADMIN' to edit")
                    Exit Sub
                End If
            Else
                MsgBox("Please select FCE with Source 'ADMIN' to edit")
                Exit Sub
            End If
            Dim frmFCE As New FacilityComplianceEvent(nOwnerID, nInspID, nFacilityID, nFCEID, strOwnerName)
            frmFCE.CallingForm = Me
            Me.Tag = "0"
            frmFCE.ShowDialog()
            If Me.Tag = "1" Then
                ' refresh grid
                SetupTabs()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCompDeleteFCE_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCompDeleteFCE.Click
        Dim bolFCEDeleted, bolOwnerDelete As Boolean
        Dim nFCEDeletedCount, nFCECount As Integer
        Try
            If Not MusterContainer.AppUser.HasAccess(UIUtilsGen.ModuleID.CAE, MusterContainer.AppUser.UserKey, UIUtilsGen.EntityTypes.CAEFacilityCompliantEvent) Then
                MessageBox.Show("You do not have rights to save Facility Compliance Event.", "ACCESS DENIED", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            bolFCEDeleted = False
            For Each ugRow In ugCompliance.Rows
                bolOwnerDelete = False
                nFCEDeletedCount = 0
                nFCECount = ugRow.ChildBands(0).Rows.Count
                If ugRow.Cells("SELECTED").Value = True Or _
                    ugRow.Cells("SELECTED").Text = True Then
                    bolOwnerDelete = True
                End If
                For Each ugChildRow In ugRow.ChildBands(0).Rows
                    If ugChildRow.Cells("SELECTED").Value = True Or _
                        ugChildRow.Cells("SELECTED").Text = True Then
                        If ugChildRow.Cells("SOURCE").Value = "ADMIN" Then
                            Dim result As DialogResult
                            result = MsgBox("Do you want to delete FCE for Facility : " + ugChildRow.Cells("FACILITY").Value.ToString + " (" + ugChildRow.Cells("FACILITY_ID").Value.ToString + ")", MsgBoxStyle.YesNo)
                            If result = DialogResult.Yes Then
                                nFCEDeletedCount += 1
                                bolFCEDeleted = True
                                pFCE.Retrieve(ugChildRow.Cells("FCE_ID").Value)
                                pFCE.Deleted = True
                                ' delete inspection
                                If pFCE.InspectionID > 0 Then
                                    pInspection.Retrieve(pFCE.InspectionID)
                                    pInspection.Deleted = True
                                    pInspection.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal, True)
                                    If Not UIUtilsGen.HasRights(returnVal) Then
                                        Exit Sub
                                    End If
                                    ' delete all citations related to this fce id
                                    If pInsCitation Is Nothing Then
                                        pInsCitation = New MUSTER.BusinessLogic.pInspectionCitation
                                    End If
                                    Dim colCitations As MUSTER.Info.InspectionCitationsCollection = pInsCitation.RetrieveByOtherID(, pFCE.ID, , False)
                                    For Each inspCitationInfo As MUSTER.Info.InspectionCitationInfo In colCitations.Values
                                        pInsCitation.InspectionCitationInfo = inspCitationInfo
                                        pInsCitation.Deleted = True

                                        If pInsCitation.ID <= 0 Then
                                            pInsCitation.CreatedBy = MusterContainer.AppUser.ID
                                        Else
                                            pInsCitation.ModifiedBy = MusterContainer.AppUser.ID
                                        End If

                                        pInsCitation.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal, True)
                                        If Not UIUtilsGen.HasRights(returnVal) Then
                                            Exit Sub
                                        End If
                                    Next
                                End If
                                If pFCE.ID <= 0 Then
                                    pFCE.CreatedBy = MusterContainer.AppUser.ID
                                Else
                                    pFCE.ModifiedBy = MusterContainer.AppUser.ID
                                End If
                                pFCE.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal, True)
                                If Not UIUtilsGen.HasRights(returnVal) Then
                                    Exit Sub
                                End If

                                ugChildRow.Delete(False)
                            End If
                        Else
                            MsgBox("Cannot delete System Generated FCE. Please choose FCE with source 'ADMIN'")
                            Exit Sub
                        End If
                    End If
                Next
                If bolOwnerDelete Or nFCEDeletedCount = nFCECount Then
                    ugRow.Delete(False)
                End If
            Next
            If bolFCEDeleted Then
                MsgBox("FCE(s) Deleted Successfully")
            Else
                MsgBox("Please select a FCE to Delete")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnComplianceExpandCollapseAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnComplianceExpandCollapseAll.Click
        Try
            If btnComplianceExpandCollapseAll.Text = "Expand All" Then
                ExpandAll(True, ugCompliance, btnComplianceExpandCollapseAll)
            Else
                ExpandAll(False, ugCompliance, btnComplianceExpandCollapseAll)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugCompliance_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugCompliance.InitializeLayout
        Try
            e.Layout.Bands(0).Override.RowAppearance.BackColor = Color.White
            e.Layout.Bands(1).Override.RowAppearance.BackColor = Color.RosyBrown
            e.Layout.Bands(1).Override.RowAlternateAppearance.BackColor = Color.PeachPuff
            e.Layout.Bands(2).Override.RowAppearance.BackColor = Color.Khaki

            e.Layout.Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Free

            e.Layout.Bands(0).Columns("OWNER_ID").Hidden = True

            e.Layout.Bands(1).Columns("FCE_ID").Hidden = True
            e.Layout.Bands(1).Columns("OWNER_ID").Hidden = True
            e.Layout.Bands(1).Columns("INSPECTION_ID").Hidden = True
            e.Layout.Bands(1).Columns("OWNER @ INSPECTION").Hidden = True

            e.Layout.Bands(2).Columns("INSPECTION_ID").Hidden = True
            e.Layout.Bands(2).Columns("FCE_ID").Hidden = True
            e.Layout.Bands(2).Columns("FACILITY_ID").Hidden = True
            e.Layout.Bands(2).Columns("INS_CIT_ID").Hidden = True
            e.Layout.Bands(2).Columns("QUESTION_ID").Hidden = True
            e.Layout.Bands(2).Columns("CITATION_ID").Hidden = True
            e.Layout.Bands(2).Columns("SMALL").Hidden = True
            e.Layout.Bands(2).Columns("MEDIUM").Hidden = True
            e.Layout.Bands(2).Columns("LARGE").Hidden = True

            e.Layout.Bands(3).Columns("INS_CIT_ID").Hidden = True
            e.Layout.Bands(3).Columns("INSPECTION_ID").Hidden = True
            e.Layout.Bands(3).Columns("QUESTION_ID").Hidden = True
            e.Layout.Bands(3).Columns("CITATION_ID").Hidden = True

            e.Layout.Bands(0).Columns("SELECTED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit
            e.Layout.Bands(0).Columns("OWNERNAME").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            If e.Layout.Bands(0).Summaries.Count = 0 Then
                e.Layout.Bands(0).Summaries.Add(Infragistics.Win.UltraWinGrid.SummaryType.Sum, e.Layout.Bands(0).Columns("FACILITIES"), SummaryPosition.UseSummaryPositionColumn)
                e.Layout.Bands(0).Summaries.Add(Infragistics.Win.UltraWinGrid.SummaryType.Count, e.Layout.Bands(0).Columns("OWNERNAME"), SummaryPosition.UseSummaryPositionColumn)
            End If

            e.Layout.Bands(1).Columns("SELECTED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit
            e.Layout.Bands(1).Columns("DATE GENERATED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("FACILITY_ID").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("FACILITY").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("SOURCE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("INSPECTOR").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("INSPECTED ON").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("CITATIONS").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            'e.Layout.Bands(1).Columns("OWNER @ INSPECTION").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            e.Layout.Bands(2).Columns("CITATION").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(2).Columns("CITATIONTEXT").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(2).Columns("CATEGORY").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(2).Columns("DUE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(2).Columns("RECEIVED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(2).Columns("CCAT").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(2).Columns("CCAT COMMENTS").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            e.Layout.Bands(2).Columns("CITATIONTEXT").Header.Caption = "CITATION TEXT"

            e.Layout.Bands(3).Columns("DISCREP TEXT").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(3).Columns("DISCREP TEXT").Width = 400
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugCompliance_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugCompliance.CellChange
        If bolLoading Then Exit Sub
        Try
            bolLoading = True
            If "SELECTED".Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Band.Index = 0 Then
                    For Each ugChildRow In e.Cell.Row.ChildBands(0).Rows
                        ugChildRow.Cells("SELECTED").Value = e.Cell.Text
                    Next
                End If
                e.Cell.Value = e.Cell.Text
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            bolLoading = False
        End Try
    End Sub
    Private Sub btnComplianceRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnComplianceRefresh.Click
        SetupTabs()
    End Sub
#End Region

#Region "OCE Escalation"
    Private Sub ProcessOCEEscalationLogic(ByRef ug As Infragistics.Win.UltraWinGrid.UltraGridRow, Optional ByVal dtStr As String = "", Optional ByVal strWhatDate As String = "", Optional ByRef skipQuestion As Integer = 0)
        Dim oceID, oceStatus, ownerID, workshopResult, showCauseResult, commissionResult, pendingLetter, ocePath, pendingLetterTemplateNum, escalationID As Integer
        Dim nextDueDate, nextDueDateOld, overrideDueDate, workshopDate, showCauseDate, commissionDate, citationDueDate, dateReceived, userProvidedDate As Date
        Dim workshopRequired As Boolean
        Dim policyAmount, overrideAmount, settlementAmount, paidAmount As Decimal
        Dim stroceStatus, strpendingLetter, strEscalation As String
        Dim strReturn As String = String.Empty
        Dim adminDate, adminResult As Object
        Dim dt As Date
        Dim ds As New DataSet
        Try
            oceID = ug.Cells("OCE_ID").Value
            oceStatus = ug.Cells("OCE_STATUS").Value
            ownerID = ug.Cells("OWNER_ID").Value
            nextDueDate = IIf(ug.Cells("NEXT DUE DATE").Value Is DBNull.Value, CDate("01/01/0001"), ug.Cells("NEXT DUE DATE").Value)
            nextDueDate = nextDueDate
            overrideDueDate = IIf(ug.Cells("OVERRIDE DUE DATE").Value Is DBNull.Value, CDate("01/01/0001"), ug.Cells("OVERRIDE DUE DATE").Value)
            policyAmount = IIf(ug.Cells("POLICY AMOUNT").Value Is DBNull.Value, -1.0, ug.Cells("POLICY AMOUNT").Value)
            overrideAmount = IIf(ug.Cells("OVERRIDE AMOUNT").Value Is DBNull.Value, -1.0, ug.Cells("OVERRIDE AMOUNT").Value)
            settlementAmount = IIf(ug.Cells("SETTLEMENT AMOUNT").Value Is DBNull.Value, -1.0, ug.Cells("SETTLEMENT AMOUNT").Value)
            paidAmount = IIf(ug.Cells("PAID AMOUNT").Value Is DBNull.Value, -1.0, ug.Cells("PAID AMOUNT").Value)
            workshopRequired = IIf(ug.Cells("WORKSHOP_REQUIRED").Value Is DBNull.Value, False, ug.Cells("WORKSHOP_REQUIRED").Value)
            workshopDate = IIf(ug.Cells("WORKSHOP DATE").Value Is DBNull.Value, CDate("01/01/0001"), ug.Cells("WORKSHOP DATE").Value)
            workshopResult = IIf(ug.Cells("WORKSHOP RESULT").Value Is DBNull.Value, 0, ug.Cells("WORKSHOP RESULT").Value)
            showCauseDate = IIf(ug.Cells("SHOW CAUSE HEARING DATE").Value Is DBNull.Value, CDate("01/01/0001"), ug.Cells("SHOW CAUSE HEARING DATE").Value)
            showCauseResult = IIf(ug.Cells("SHOW CAUSE HEARING RESULT").Value Is DBNull.Value, 0, ug.Cells("SHOW CAUSE HEARING RESULT").Value)
            commissionDate = IIf(ug.Cells("COMMISSION HEARING DATE").Value Is DBNull.Value, CDate("01/01/0001"), ug.Cells("COMMISSION HEARING DATE").Value)
            commissionResult = IIf(ug.Cells("COMMISSION HEARING RESULT").Value Is DBNull.Value, 0, ug.Cells("COMMISSION HEARING RESULT").Value)
            pendingLetter = IIf(ug.Cells("PENDING_LETTER").Value Is DBNull.Value, 0, ug.Cells("PENDING_LETTER").Value)
            citationDueDate = IIf(ug.Cells("CITATION_DUE_DATE").Value Is DBNull.Value, CDate("01/01/0001"), ug.Cells("CITATION_DUE_DATE").Value)
            ocePath = IIf(ug.Cells("OCE_PATH").Value Is DBNull.Value, 0, ug.Cells("OCE_PATH").Value)
            dateReceived = IIf(ug.Cells("DATE RECEIVED").Value Is DBNull.Value, CDate("01/01/0001"), ug.Cells("DATE RECEIVED").Value)
            stroceStatus = IIf(ug.Cells("STATUS").Value Is DBNull.Value, String.Empty, ug.Cells("STATUS").Text)
            strpendingLetter = IIf(ug.Cells("PENDING LETTER").Value Is DBNull.Value, String.Empty, ug.Cells("PENDING LETTER").Text)
            strEscalation = IIf(ug.Cells("ESCALATION").Value Is DBNull.Value, String.Empty, ug.Cells("ESCALATION").Text)
            pendingLetterTemplateNum = IIf(ug.Cells("PENDING_LETTER_TEMPLATE_NUM").Value Is DBNull.Value, 0, ug.Cells("PENDING_LETTER_TEMPLATE_NUM").Value)
            escalationID = IIf(ug.Cells("ESCALATION_ID").Value Is DBNull.Value, 0, ug.Cells("ESCALATION_ID").Value)
            adminDate = IIf(ug.Cells("ADMIN HEARING DATE").Value Is DBNull.Value, Nothing, ug.Cells("ADMIN HEARING DATE").Value)
            adminResult = IIf(ug.Cells("ADMIN HEARING RESULT").Value Is DBNull.Value, Nothing, ug.Cells("ADMIN HEARING RESULT").Value)

            If dtStr = String.Empty Then
                userProvidedDate = CDate("01/01/0001")
            Else
                userProvidedDate = CDate(dtStr)
            End If

            If Date.Compare(userProvidedDate, CDate("01/01/0001")) = 0 And strWhatDate <> String.Empty Then
                ' cancel
                OCECancelEscalation = True
            Else

                If pendingLetter > 0 AndAlso (skipQuestion = 1 OrElse (skipQuestion = 0 AndAlso MsgBox("You still have letters that are not printed. Do you wish to still proceed with the process escalations", MsgBoxStyle.YesNo) = MsgBoxResult.Yes)) Then
                    pendingLetter = 0
                    skipQuestion = 1
                ElseIf pendingLetter > 0 AndAlso skipQuestion = 0 Then
                    skipQuestion = 2
                End If
                Dim flagNFA As Short
                If escalationID = "1248" Then
                    Dim strFees As String
                    ds = pOCE.RunSQLQuery("SELECT dbo.udfGetOwnerPastDueFees(" + ownerID.ToString + ",0,NULL)")
                    If ds.Tables(0).Rows(0)(0) > 0 Then
                        strFees = ds.Tables(0).Rows(0)(0)
                        strFees = strFees.Split(".")(0)
                        If MsgBox("Fees are owed, do you still want to escalate to NFA?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                            flagNFA = 1
                        Else
                            flagNFA = 0
                        End If
                    End If
                End If

                strReturn = pOCE.ExecuteCAEOCEEscalation(flagNFA, oceID, oceStatus, ownerID, nextDueDate, _
                                            overrideDueDate, policyAmount, overrideAmount, _
                                            settlementAmount, paidAmount, workshopRequired, _
                                            workshopDate, workshopResult, showCauseDate, _
                                            showCauseResult, commissionDate, commissionResult, _
                                            pendingLetter, citationDueDate, ocePath, _
                                            dateReceived, userProvidedDate, stroceStatus, _
                                            strpendingLetter, strEscalation, pendingLetterTemplateNum, escalationID, adminDate, adminResult, MusterContainer.AppUser.ID)
            End If
            If strReturn = String.Empty And Not OCECancelEscalation Then
                ' update grid
                ug.Cells("OCE_STATUS").Value = oceStatus
                ug.Cells("NEXT DUE DATE").Value = nextDueDate
                If Date.Compare(overrideDueDate, CDate("01/01/0001")) = 0 Then
                    ug.Cells("OVERRIDE DUE DATE").Value = DBNull.Value
                Else
                    ug.Cells("OVERRIDE DUE DATE").Value = overrideDueDate
                End If
                ug.Cells("SETTLEMENT AMOUNT").Value = settlementAmount
                If strWhatDate = "GET SHOW CAUSE HEARING DATE FROM USER" Then
                    ug.Cells("SHOW CAUSE HEARING DATE").Value = showCauseDate
                    ug.Cells("SHOW CAUSE HEARING RESULT").Value = showCauseResult
                ElseIf strWhatDate = "GET COMMISSION HEARING DATE FROM USER" Then
                    ug.Cells("COMMISSION HEARING DATE").Value = commissionDate
                    ug.Cells("COMMISSION HEARING RESULT").Value = commissionResult
                End If
                ug.Cells("PENDING_LETTER").Value = pendingLetter
                If ug.Cells("PENDING_LETTER").Value <> pendingLetter Then
                    ug.Cells("LETTER PRINTED").Value = False
                End If

                ug.Cells("STATUS").Value = stroceStatus
                ug.Cells("PENDING LETTER").Value = strpendingLetter
                ug.Cells("ESCALATION").Value = strEscalation
                ug.Cells("ESCALATION_ID").Value = escalationID
                ug.Cells("PENDING_LETTER_TEMPLATE_NUM").Value = pendingLetterTemplateNum
                SetugRowComboValue(ug)
                ' if Next Due Date is modified, change the citation(s) due date
                If Date.Compare(nextDueDateOld, nextDueDate) <> 0 Then
                    bolOCEDueDateModified = True

                    If pendingLetter > 0 AndAlso (skipQuestion OrElse MsgBox("You still have letters that are not printed. Do you wish to still proceed with the process escalations", MsgBoxStyle.YesNo) = MsgBoxResult.Yes) Then
                        pendingLetter = 0
                        skipQuestion = True
                    End If

                    ModifyOCECitationDueDate(ug, skipQuestion)
                End If
            ElseIf strReturn <> String.Empty And Not OCECancelEscalation Then
                ' get date and call function again
                frmWorkshopDate = New WorkShopDate
                frmWorkshopDate.DateLabel = IIf(strReturn = "GET SHOW CAUSE HEARING DATE FROM USER", "Show Cause Hearing Date", "Commission Hearing Date")
                frmWorkshopDate.frmText = frmWorkshopDate.DateLabel
                frmWorkshopDate.ShowDialog()
                dt = frmWorkshopDate.WorkshopDate
                ProcessOCEEscalationLogic(ug, dt.ToShortDateString, strReturn, skipQuestion)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub


    Private Sub ProcessOCEEscalationLogicManual(ByRef ug As Infragistics.Win.UltraWinGrid.UltraGridRow, Optional ByVal dtStr As String = "", Optional ByVal strWhatDate As String = "", Optional ByRef skipQuestion As Integer = 0, Optional ByVal code As Int16 = -1)
        Dim oceID, oceStatus, ownerID, workshopResult, showCauseResult, commissionResult, pendingLetter, ocePath, pendingLetterTemplateNum, escalationID As Integer
        Dim nextDueDate, nextDueDateOld, overrideDueDate, workshopDate, showCauseDate, commissionDate, citationDueDate, dateReceived, userProvidedDate As Date
        Dim workshopRequired As Boolean
        Dim policyAmount, overrideAmount, settlementAmount, paidAmount As Decimal
        Dim stroceStatus, strpendingLetter, strEscalation As String
        Dim strReturn As String = String.Empty
        Dim adminDate, adminResult As Object
        Dim dt As Date
        Try
            oceID = ug.Cells("OCE_ID").Value
            oceStatus = ug.Cells("OCE_STATUS").Value
            ownerID = ug.Cells("OWNER_ID").Value
            nextDueDate = IIf(ug.Cells("NEXT DUE DATE").Value Is DBNull.Value, CDate("01/01/0001"), ug.Cells("NEXT DUE DATE").Value)
            nextDueDate = nextDueDate
            overrideDueDate = IIf(ug.Cells("OVERRIDE DUE DATE").Value Is DBNull.Value, CDate("01/01/0001"), ug.Cells("OVERRIDE DUE DATE").Value)
            policyAmount = IIf(ug.Cells("POLICY AMOUNT").Value Is DBNull.Value, -1.0, ug.Cells("POLICY AMOUNT").Value)
            overrideAmount = IIf(ug.Cells("OVERRIDE AMOUNT").Value Is DBNull.Value, -1.0, ug.Cells("OVERRIDE AMOUNT").Value)
            settlementAmount = IIf(ug.Cells("SETTLEMENT AMOUNT").Value Is DBNull.Value, -1.0, ug.Cells("SETTLEMENT AMOUNT").Value)
            paidAmount = IIf(ug.Cells("PAID AMOUNT").Value Is DBNull.Value, -1.0, ug.Cells("PAID AMOUNT").Value)
            workshopRequired = IIf(ug.Cells("WORKSHOP_REQUIRED").Value Is DBNull.Value, False, ug.Cells("WORKSHOP_REQUIRED").Value)
            workshopDate = IIf(ug.Cells("WORKSHOP DATE").Value Is DBNull.Value, CDate("01/01/0001"), ug.Cells("WORKSHOP DATE").Value)
            workshopResult = IIf(ug.Cells("WORKSHOP RESULT").Value Is DBNull.Value, 0, ug.Cells("WORKSHOP RESULT").Value)
            showCauseDate = IIf(ug.Cells("SHOW CAUSE HEARING DATE").Value Is DBNull.Value, CDate("01/01/0001"), ug.Cells("SHOW CAUSE HEARING DATE").Value)
            showCauseResult = IIf(ug.Cells("SHOW CAUSE HEARING RESULT").Value Is DBNull.Value, 0, ug.Cells("SHOW CAUSE HEARING RESULT").Value)
            commissionDate = IIf(ug.Cells("COMMISSION HEARING DATE").Value Is DBNull.Value, CDate("01/01/0001"), ug.Cells("COMMISSION HEARING DATE").Value)
            commissionResult = IIf(ug.Cells("COMMISSION HEARING RESULT").Value Is DBNull.Value, 0, ug.Cells("COMMISSION HEARING RESULT").Value)
            pendingLetter = IIf(ug.Cells("PENDING_LETTER").Value Is DBNull.Value, 0, ug.Cells("PENDING_LETTER").Value)
            citationDueDate = IIf(ug.Cells("CITATION_DUE_DATE").Value Is DBNull.Value, CDate("01/01/0001"), ug.Cells("CITATION_DUE_DATE").Value)
            ocePath = IIf(ug.Cells("OCE_PATH").Value Is DBNull.Value, 0, ug.Cells("OCE_PATH").Value)
            dateReceived = IIf(ug.Cells("DATE RECEIVED").Value Is DBNull.Value, CDate("01/01/0001"), ug.Cells("DATE RECEIVED").Value)
            stroceStatus = IIf(ug.Cells("STATUS").Value Is DBNull.Value, String.Empty, ug.Cells("STATUS").Text)
            strpendingLetter = IIf(ug.Cells("PENDING LETTER").Value Is DBNull.Value, String.Empty, ug.Cells("PENDING LETTER").Text)
            strEscalation = IIf(ug.Cells("ESCALATION").Value Is DBNull.Value, String.Empty, ug.Cells("ESCALATION").Text)
            pendingLetterTemplateNum = IIf(ug.Cells("PENDING_LETTER_TEMPLATE_NUM").Value Is DBNull.Value, 0, ug.Cells("PENDING_LETTER_TEMPLATE_NUM").Value)
            escalationID = code
            adminDate = IIf(ug.Cells("ADMIN HEARING DATE").Value Is DBNull.Value, Nothing, ug.Cells("ADMIN HEARING DATE").Value)
            adminResult = IIf(ug.Cells("ADMIN HEARING RESULT").Value Is DBNull.Value, Nothing, ug.Cells("ADMIN HEARING RESULT").Value)

            If dtStr = String.Empty Then
                userProvidedDate = CDate("01/01/0001")
            Else
                userProvidedDate = CDate(dtStr)
            End If

            If Date.Compare(userProvidedDate, CDate("01/01/0001")) = 0 And strWhatDate <> String.Empty Then
                ' cancel
                OCECancelEscalation = True
            Else

                If pendingLetter > 0 AndAlso (skipQuestion = 1 OrElse (skipQuestion = 0 AndAlso MsgBox("You still have letters that are not printed. Do you wish to still proceed with the process escalations", MsgBoxStyle.YesNo) = MsgBoxResult.Yes)) Then
                    pendingLetter = 0
                    skipQuestion = 1
                ElseIf pendingLetter > 0 AndAlso skipQuestion = 0 Then
                    skipQuestion = 2
                End If


                Dim flagNFA As Short = 0
                strReturn = pOCE.ExecuteCAEOCEEscalation(flagNFA, oceID, oceStatus, ownerID, nextDueDate, _
                                            overrideDueDate, policyAmount, overrideAmount, _
                                            settlementAmount, paidAmount, workshopRequired, _
                                            workshopDate, workshopResult, showCauseDate, _
                                            showCauseResult, commissionDate, commissionResult, _
                                            pendingLetter, citationDueDate, ocePath, _
                                            dateReceived, userProvidedDate, stroceStatus, _
                                            strpendingLetter, strEscalation, pendingLetterTemplateNum, escalationID, adminDate, adminResult, MusterContainer.AppUser.ID)
            End If
            If strReturn = String.Empty And Not OCECancelEscalation Then
                ' update grid
                ug.Cells("OCE_STATUS").Value = oceStatus
                ug.Cells("NEXT DUE DATE").Value = nextDueDate
                If Date.Compare(overrideDueDate, CDate("01/01/0001")) = 0 Then
                    ug.Cells("OVERRIDE DUE DATE").Value = DBNull.Value
                Else
                    ug.Cells("OVERRIDE DUE DATE").Value = overrideDueDate
                End If
                ug.Cells("SETTLEMENT AMOUNT").Value = settlementAmount
                If strWhatDate = "GET SHOW CAUSE HEARING DATE FROM USER" Then
                    ug.Cells("SHOW CAUSE HEARING DATE").Value = showCauseDate
                    ug.Cells("SHOW CAUSE HEARING RESULT").Value = showCauseResult
                ElseIf strWhatDate = "GET COMMISSION HEARING DATE FROM USER" Then
                    ug.Cells("COMMISSION HEARING DATE").Value = commissionDate
                    ug.Cells("COMMISSION HEARING RESULT").Value = commissionResult
                End If
                ug.Cells("PENDING_LETTER").Value = pendingLetter
                If ug.Cells("PENDING_LETTER").Value <> pendingLetter Then
                    ug.Cells("LETTER PRINTED").Value = False
                End If

                ug.Cells("STATUS").Value = stroceStatus
                ug.Cells("PENDING LETTER").Value = strpendingLetter
                ug.Cells("ESCALATION").Value = strEscalation
                ug.Cells("ESCALATION_ID").Value = escalationID
                ug.Cells("PENDING_LETTER_TEMPLATE_NUM").Value = pendingLetterTemplateNum
                SetugRowComboValue(ug)
                ' if Next Due Date is modified, change the citation(s) due date
                If Date.Compare(nextDueDateOld, nextDueDate) <> 0 Then
                    bolOCEDueDateModified = True

                    If pendingLetter > 0 AndAlso (skipQuestion OrElse MsgBox("You still have letters that are not printed. Do you wish to still proceed with the process escalations", MsgBoxStyle.YesNo) = MsgBoxResult.Yes) Then
                        pendingLetter = 0
                        skipQuestion = True
                    End If

                    ModifyOCECitationDueDate(ug, skipQuestion)
                End If
            ElseIf strReturn <> String.Empty And Not OCECancelEscalation Then
                ' get date and call function again
                frmWorkshopDate = New WorkShopDate
                frmWorkshopDate.DateLabel = IIf(strReturn = "GET SHOW CAUSE HEARING DATE FROM USER", "Show Cause Hearing Date", "Commission Hearing Date")
                frmWorkshopDate.frmText = frmWorkshopDate.DateLabel
                frmWorkshopDate.ShowDialog()
                dt = frmWorkshopDate.WorkshopDate
                ProcessOCEEscalationLogic(ug, dt.ToShortDateString, strReturn, skipQuestion)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub GetOCEEscalation(ByRef ug As Infragistics.Win.UltraWinGrid.UltraGridRow, ByRef skipQuestion As Integer)
        Dim oceID, oceStatus, ownerID, workshopResult, showCauseResult, commissionResult, pendingLetter, ocePath, escalationID As Integer
        Dim nextDueDate, overrideDueDate, workshopDate, showCauseDate, commissionDate, citationDueDate, dateReceived, userProvidedDate As Date
        Dim workshopRequired As Boolean
        Dim policyAmount, overrideAmount, settlementAmount, paidAmount As Decimal
        Dim adminDate, adminResult As Object

        Dim strReturn As String = String.Empty
        Try


            If (DirectCast(IIf(ug.Cells("OVERRIDE DUE DATE").Value Is DBNull.Value OrElse ug.Cells("OVERRIDE DUE DATE").Value < CDate("1900-01-01"), IIf(ug.Cells("NEXT DUE DATE").Value Is DBNull.Value, CDate("1900-01-01"), ug.Cells("NEXT DUE DATE").Value), ug.Cells("OVERRIDE DUE DATE").Value), DateTime).Date < Today.Date) AndAlso ug.Cells("OCE_STATUS").Value = IIf(ug.Cells("ESCALATION_ID").Value Is DBNull.Value, 0, ug.Cells("ESCALATION_ID").Value) Then

                oceID = ug.Cells("OCE_ID").Value
                oceStatus = ug.Cells("OCE_STATUS").Value
                ownerID = ug.Cells("OWNER_ID").Value
                nextDueDate = IIf(ug.Cells("NEXT DUE DATE").Value Is DBNull.Value, CDate("01/01/0001"), ug.Cells("NEXT DUE DATE").Value)
                overrideDueDate = IIf(ug.Cells("OVERRIDE DUE DATE").Value Is DBNull.Value, CDate("01/01/0001"), ug.Cells("OVERRIDE DUE DATE").Value)
                policyAmount = IIf(ug.Cells("POLICY AMOUNT").Value Is DBNull.Value, -1.0, ug.Cells("POLICY AMOUNT").Value)
                overrideAmount = IIf(ug.Cells("OVERRIDE AMOUNT").Value Is DBNull.Value, -1.0, ug.Cells("OVERRIDE AMOUNT").Value)
                settlementAmount = IIf(ug.Cells("SETTLEMENT AMOUNT").Value Is DBNull.Value, -1.0, ug.Cells("SETTLEMENT AMOUNT").Value)
                paidAmount = IIf(ug.Cells("PAID AMOUNT").Value Is DBNull.Value, -1.0, ug.Cells("PAID AMOUNT").Value)
                workshopRequired = IIf(ug.Cells("WORKSHOP_REQUIRED").Value Is DBNull.Value, False, ug.Cells("WORKSHOP_REQUIRED").Value)
                workshopDate = IIf(ug.Cells("WORKSHOP DATE").Value Is DBNull.Value, CDate("01/01/0001"), ug.Cells("WORKSHOP DATE").Value)
                workshopResult = IIf(ug.Cells("WORKSHOP RESULT").Value Is DBNull.Value, 0, ug.Cells("WORKSHOP RESULT").Value)
                showCauseDate = IIf(ug.Cells("SHOW CAUSE HEARING DATE").Value Is DBNull.Value, CDate("01/01/0001"), ug.Cells("SHOW CAUSE HEARING DATE").Value)
                showCauseResult = IIf(ug.Cells("SHOW CAUSE HEARING RESULT").Value Is DBNull.Value, 0, ug.Cells("SHOW CAUSE HEARING RESULT").Value)
                commissionDate = IIf(ug.Cells("COMMISSION HEARING DATE").Value Is DBNull.Value, CDate("01/01/0001"), ug.Cells("COMMISSION HEARING DATE").Value)
                commissionResult = IIf(ug.Cells("COMMISSION HEARING RESULT").Value Is DBNull.Value, 0, ug.Cells("COMMISSION HEARING RESULT").Value)
                pendingLetter = IIf(ug.Cells("PENDING_LETTER").Value Is DBNull.Value, 0, ug.Cells("PENDING_LETTER").Value)
                citationDueDate = IIf(ug.Cells("CITATION_DUE_DATE").Value Is DBNull.Value, CDate("01/01/0001"), ug.Cells("CITATION_DUE_DATE").Value)
                ocePath = IIf(ug.Cells("OCE_PATH").Value Is DBNull.Value, 0, ug.Cells("OCE_PATH").Value)
                dateReceived = IIf(ug.Cells("DATE RECEIVED").Value Is DBNull.Value, CDate("01/01/0001"), ug.Cells("DATE RECEIVED").Value)
                escalationID = IIf(ug.Cells("ESCALATION_ID").Value Is DBNull.Value, 0, ug.Cells("ESCALATION_ID").Value)
                adminResult = IIf(ug.Cells("ADMIN HEARING RESULT").Value Is DBNull.Value, Nothing, ug.Cells("ADMIN HEARING RESULT").Value)
                adminDate = IIf(ug.Cells("ADMIN HEARING DATE").Value Is DBNull.Value, Nothing, ug.Cells("ADMIN HEARING DATE").Value)

                If pendingLetter > 0 AndAlso (skipQuestion = 1 OrElse (skipQuestion = 0 AndAlso MsgBox("You have a letter that is no generated. Will you like to proceed to set the next possible escalation?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes)) Then
                    pendingLetter = 0
                    skipQuestion = 1
                ElseIf pendingLetter > 0 AndAlso skipQuestion = 0 Then
                    skipQuestion = 2
                End If

                Dim flagNFA As Short = 0
                If escalationID = 1248 Then
                    flagNFA = 1
                End If
                strReturn = pOCE.GetCAEOCEEscalation(flagNFA, oceID, oceStatus, ownerID, nextDueDate, _
                                            overrideDueDate, policyAmount, overrideAmount, _
                                            settlementAmount, paidAmount, workshopRequired, _
                                            workshopDate, workshopResult, showCauseDate, _
                                            showCauseResult, commissionDate, commissionResult, _
                                            pendingLetter, citationDueDate, ocePath, _
                                            dateReceived, escalationID, adminDate, adminResult)
                If strReturn = String.Empty Then
                    strReturn = IIf(ug.Cells("STATUS").Value Is DBNull.Value, String.Empty, ug.Cells("STATUS").Text)
                End If
            Else
                strReturn = IIf(ug.Cells("ESCALATION").Value Is DBNull.Value, String.Empty, ug.Cells("ESCALATION").Value)
                escalationID = IIf(ug.Cells("ESCALATION_ID").Value Is DBNull.Value, 0, ug.Cells("ESCALATION_ID").Value)
            End If

            ug.Cells("ESCALATION").Value = strReturn
            ug.Cells("ESCALATION_ID").Value = escalationID
            SetugRowComboValue(ug)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Enforcement"
    Private Sub ModifyOCECitationDueDate(ByVal ug As Infragistics.Win.UltraWinGrid.UltraGridRow, ByRef skipquestion As Boolean)
        Dim dt As Date
        Dim bolLoadingLocal As Boolean = bolLoading
        Try
            If bolOCEDueDateModified OrElse Me.bolOCEModified Then
                bolLoading = True
                If Not bolOCEDueDateModified OrElse MsgBox("OCE's Due Date was modified, Do you want to apply change to all the Citation(s) Due Date?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    If Not ug.Cells("NEXT DUE DATE").Value Is DBNull.Value Then
                        dt = IIf(ug.Cells("OVERRIDE DUE DATE").Value Is DBNull.Value, ug.Cells("NEXT DUE DATE").Value, ug.Cells("OVERRIDE DUE DATE").Value)
                    End If
                    For Each ugGrandChildRow In ug.ChildBands(0).Rows
                        pInsCitation.Retrieve(New MUSTER.Info.InspectionInfo, ugGrandChildRow.Cells("INS_CIT_ID").Value)
                        pInsCitation.OCEID = ug.Cells("OCE_ID").Value
                        pInsCitation.CitationDueDate = dt

                        If pInsCitation.ID <= 0 Then
                            pInsCitation.CreatedBy = MusterContainer.AppUser.ID
                        Else
                            pInsCitation.ModifiedBy = MusterContainer.AppUser.ID
                        End If

                        SaveOCECitation(ugGrandChildRow, skipquestion)

                        pInsCitation.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal, True)
                        If Not UIUtilsGen.HasRights(returnVal) Then
                            Exit Sub
                        End If

                        ugGrandChildRow.Cells("DUE").Value = dt
                    Next
                End If
                bolOCEDueDateModified = False
            End If
        Catch ex As Exception
            Throw ex
        Finally
            bolLoading = bolLoadingLocal
        End Try
    End Sub


    Private Sub btnInsViewEditCheckList2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInsViewEditCheckList2.Click
        Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            ugRow = getSelectedRow()
            If ugRow Is Nothing Then
                MsgBox("Please select a Facilty")
                Exit Sub
            End If
            If ugRow.Cells("SELECTED").Value = True Or _
                ugRow.Cells("SELECTED").Text = True Then

                pInspection.Retrieve(ugRow.Cells("INSPECTION_ID").Value)

                If pInspection.SubmittedDate = Nothing OrElse pInspection.SubmittedDate < "1/1/1910" Then

                    MsgBox("Cannot Edit / View checklist - Inspection not Submitted to C&E")
                    Exit Sub
                Else
                    Dim bolReadOnly As Boolean = True
                    If MusterContainer.AppUser.HasAccess(UIUtilsGen.ModuleID.CAE, MusterContainer.AppUser.UserKey, UIUtilsGen.EntityTypes.Inspection) Then
                        bolReadOnly = False
                    End If
                    Dim frmChecklist As New CheckList(pInspection, True, , UIUtilsGen.ModuleID.CAE)
                    frmChecklist.WindowState = FormWindowState.Maximized
                    frmChecklist.CallingForm = Me
                    Me.Tag = "0"
                    frmChecklist.ShowDialog()
                    If Me.Tag = "1" Then
                        ' set inspection viewed date as today
                        pInspection.CAEViewed = Today.Date
                        pInspection.ModifiedBy = MusterContainer.AppUser.ID
                        pInspection.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal, True)
                        If Not UIUtilsGen.HasRights(returnVal) Then
                            Exit Sub
                        End If

                        'refresh grid
                        'SetupTabs()
                        ' bug #2317
                        ' update row's viewed cell
                    End If
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ModifyOCECitationReceivedDate(ByVal ug As Infragistics.Win.UltraWinGrid.UltraGridRow, ByVal dt As Date)
        Dim bolLoadingLocal As Boolean = bolLoading
        Try
            If bolCitationDateModified Then
                bolLoading = True
                If MsgBox("Citation's Received Date was modified, Do you want to apply change to all the Citation(s) Received Date?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    For Each ugGrandChildRow In ug.ChildBands(0).Rows
                        pInsCitation.Retrieve(New MUSTER.Info.InspectionInfo, ugGrandChildRow.Cells("INS_CIT_ID").Value)
                        pInsCitation.OCEID = ug.Cells("OCE_ID").Value
                        pInsCitation.CitationReceivedDate = dt

                        If pInsCitation.ID <= 0 Then
                            pInsCitation.CreatedBy = MusterContainer.AppUser.ID
                        Else
                            pInsCitation.ModifiedBy = MusterContainer.AppUser.ID
                        End If

                        pInsCitation.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal, True)
                        If Not UIUtilsGen.HasRights(returnVal) Then
                            Exit Sub
                        End If

                        If Date.Compare(dt, CDate("01/01/0001")) = 0 Then
                            ugGrandChildRow.Cells("RECEIVED").Value = DBNull.Value
                        Else
                            ugGrandChildRow.Cells("RECEIVED").Value = dt
                        End If
                    Next
                End If
                bolCitationDateModified = False
            End If
        Catch ex As Exception
            Throw ex
        Finally
            bolLoading = bolLoadingLocal
        End Try
    End Sub
    Private Sub ModifyOCEDiscrepReceivedDate(ByVal ug As Infragistics.Win.UltraWinGrid.UltraGridRow, ByVal dt As Date)
        Dim bolLoadingLocal As Boolean = bolLoading
        Try
            If bolDiscrepDateModified Then
                bolLoading = True
                If MsgBox("Discrep's Received Date was modified, Do you want to apply change to all the Discrep(s) Received Date?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    For Each ugGreatGrandChildRow In ug.ChildBands(0).Rows
                        pInsDiscrep.Retrieve(New MUSTER.Info.InspectionInfo, ugGreatGrandChildRow.Cells("INS_DESCREP_ID").Value)
                        pInsDiscrep.DiscrepReceived = dt
                        pInsDiscrep.ModifiedBy = MusterContainer.AppUser.ID

                        pInsDiscrep.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal, True)
                        If Not UIUtilsGen.HasRights(returnVal) Then
                            Exit Sub
                        End If

                        ugGreatGrandChildRow.Cells("RECEIVED").Value = dt
                    Next
                End If
                bolDiscrepDateModified = False
            End If
        Catch ex As Exception
            Throw ex
        Finally
            bolLoading = bolLoadingLocal
        End Try
    End Sub
    Private Sub SetugRowComboValue(ByRef ug As Infragistics.Win.UltraWinGrid.UltraGridRow)
        Try


            If ug.Cells("WORKSHOP_REQUIRED").Value Is DBNull.Value Then
                If ug.Cells("WORKSHOP_REQUIRED").Row.IsAlternate Then
                    ug.Cells("WORKSHOP RESULT").Appearance.BackColor = Color.PeachPuff
                    ug.Cells("WORKSHOP RESULT").Activation = Activation.NoEdit

                    ug.Cells("WORKSHOP DATE").Appearance.BackColor = Color.PeachPuff
                    ug.Cells("WORKSHOP DATE").Activation = Activation.NoEdit
                Else
                    ug.Cells("WORKSHOP RESULT").Appearance.BackColor = Color.RosyBrown
                    ug.Cells("WORKSHOP RESULT").Activation = Activation.NoEdit

                    ug.Cells("WORKSHOP DATE").Appearance.BackColor = Color.RosyBrown
                    ug.Cells("WORKSHOP DATE").Activation = Activation.NoEdit
                End If
            ElseIf ug.Cells("WORKSHOP_REQUIRED").Value Then
                ug.Cells("WORKSHOP DATE").Appearance.BackColor = Color.Yellow
                ug.Cells("WORKSHOP RESULT").Activation = Activation.AllowEdit

                If Not ug.Cells("WORKSHOP DATE").Value Is DBNull.Value AndAlso ug.Cells("WORKSHOP DATE").Value > CDate("1/1/1900") AndAlso ug.Cells("WORKSHOP DATE").Value < Today.Date Then
                    ug.Cells("WORKSHOP RESULT").Appearance.BackColor = Color.Red
                Else
                    ug.Cells("WORKSHOP RESULT").Appearance.BackColor = Color.Yellow
                End If
                ug.Cells("WORKSHOP DATE").Activation = Activation.AllowEdit
            Else
                If ug.Cells("WORKSHOP_REQUIRED").Row.IsAlternate Then
                    ug.Cells("WORKSHOP RESULT").Appearance.BackColor = Color.PeachPuff
                    ug.Cells("WORKSHOP RESULT").Activation = Activation.NoEdit

                    ug.Cells("WORKSHOP DATE").Appearance.BackColor = Color.PeachPuff
                    ug.Cells("WORKSHOP DATE").Activation = Activation.NoEdit
                Else
                    ug.Cells("WORKSHOP RESULT").Appearance.BackColor = Color.RosyBrown
                    ug.Cells("WORKSHOP RESULT").Activation = Activation.NoEdit

                    ug.Cells("WORKSHOP DATE").Appearance.BackColor = Color.RosyBrown
                    ug.Cells("WORKSHOP DATE").Activation = Activation.NoEdit
                End If
            End If

            If ug.Cells("SHOW CAUSE HEARING DATE").Value Is DBNull.Value Then
                If ug.Cells("SHOW CAUSE HEARING DATE").Row.IsAlternate Then
                    ug.Cells("SHOW CAUSE HEARING RESULT").Appearance.BackColor = Color.PeachPuff
                    ug.Cells("SHOW CAUSE HEARING RESULT").Activation = Activation.NoEdit

                    ug.Cells("SHOW CAUSE HEARING DATE").Appearance.BackColor = Color.PeachPuff
                    ug.Cells("SHOW CAUSE HEARING DATE").Activation = Activation.NoEdit
                Else
                    ug.Cells("SHOW CAUSE HEARING RESULT").Appearance.BackColor = Color.RosyBrown
                    ug.Cells("SHOW CAUSE HEARING RESULT").Activation = Activation.NoEdit

                    ug.Cells("SHOW CAUSE HEARING DATE").Appearance.BackColor = Color.RosyBrown
                    ug.Cells("SHOW CAUSE HEARING DATE").Activation = Activation.NoEdit
                End If
            Else
                ug.Cells("SHOW CAUSE HEARING RESULT").Appearance.BackColor = Color.Yellow
                ug.Cells("SHOW CAUSE HEARING RESULT").Activation = Activation.AllowEdit

                ug.Cells("SHOW CAUSE HEARING DATE").Appearance.BackColor = Color.Yellow
                ug.Cells("SHOW CAUSE HEARING DATE").Activation = Activation.AllowEdit
            End If



            If ug.Cells("OCE_STATUS").Value <> 1667 Then
                If ug.Cells("ADMIN HEARING DATE").Row.IsAlternate Then
                    ug.Cells("ADMIN HEARING RESULT").Appearance.BackColor = Color.PeachPuff
                    ug.Cells("ADMIN HEARING RESULT").Activation = Activation.NoEdit

                    ug.Cells("ADMIN HEARING DATE").Appearance.BackColor = Color.PeachPuff
                    ug.Cells("ADMIN HEARING DATE").Activation = Activation.NoEdit
                Else
                    ug.Cells("ADMIN HEARING RESULT").Appearance.BackColor = Color.RosyBrown
                    ug.Cells("ADMIN HEARING RESULT").Activation = Activation.NoEdit

                    ug.Cells("ADMIN HEARING DATE").Appearance.BackColor = Color.RosyBrown
                    ug.Cells("ADMIN HEARING DATE").Activation = Activation.NoEdit
                End If
            Else
                ug.Cells("ADMIN HEARING RESULT").Appearance.BackColor = Color.Yellow
                ug.Cells("ADMIN HEARING RESULT").Activation = Activation.AllowEdit

                ug.Cells("ADMIN HEARING DATE").Appearance.BackColor = Color.Yellow
                ug.Cells("ADMIN HEARING DATE").Activation = Activation.AllowEdit
            End If


            If ug.Cells("COMMISSION HEARING DATE").Value Is DBNull.Value Then
                If ug.Cells("COMMISSION HEARING DATE").Row.IsAlternate Then
                    ug.Cells("COMMISSION HEARING RESULT").Appearance.BackColor = Color.PeachPuff
                    ug.Cells("COMMISSION HEARING RESULT").Activation = Activation.NoEdit

                    ug.Cells("COMMISSION HEARING DATE").Appearance.BackColor = Color.PeachPuff
                    ug.Cells("COMMISSION HEARING DATE").Activation = Activation.NoEdit
                Else
                    ug.Cells("COMMISSION HEARING RESULT").Appearance.BackColor = Color.RosyBrown
                    ug.Cells("COMMISSION HEARING RESULT").Activation = Activation.NoEdit

                    ug.Cells("COMMISSION HEARING DATE").Appearance.BackColor = Color.RosyBrown
                    ug.Cells("COMMISSION HEARING DATE").Activation = Activation.NoEdit
                End If
            Else
                ug.Cells("COMMISSION HEARING RESULT").Appearance.BackColor = Color.Yellow
                ug.Cells("COMMISSION HEARING RESULT").Activation = Activation.AllowEdit

                ug.Cells("COMMISSION HEARING DATE").Appearance.BackColor = Color.Yellow
                ug.Cells("COMMISSION HEARING DATE").Activation = Activation.AllowEdit
            End If

            ug.Cells("AGREED ORDER #").Activation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit
            ug.Cells("AGREED ORDER #").Appearance.BackColor = Color.Yellow

            ug.Cells("ADMINISTRATIVE ORDER #").Activation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit
            ug.Cells("ADMINISTRATIVE ORDER #").Appearance.BackColor = Color.Yellow
            'If ug.Cells("STATUS").Text.ToUpper.IndexOf("ORDER") > -1 Then
            '    ug.Cells("AGREED ORDER #").Activation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit
            '    ug.Cells("AGREED ORDER #").Appearance.BackColor = Color.Yellow

            '    ug.Cells("ADMINISTRATIVE ORDER #").Activation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit
            '    ug.Cells("ADMINISTRATIVE ORDER #").Appearance.BackColor = Color.Yellow
            'Else
            '    If ug.Cells("AGREED ORDER #").Row.IsAlternate Then
            '        ug.Cells("AGREED ORDER #").Activation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            '        ug.Cells("AGREED ORDER #").Appearance.BackColor = Color.PeachPuff

            '        ug.Cells("ADMINISTRATIVE ORDER #").Activation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            '        ug.Cells("ADMINISTRATIVE ORDER #").Appearance.BackColor = Color.PeachPuff
            '    Else
            '        ug.Cells("AGREED ORDER #").Activation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            '        ug.Cells("AGREED ORDER #").Appearance.BackColor = Color.RosyBrown

            '        ug.Cells("ADMINISTRATIVE ORDER #").Activation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            '        ug.Cells("ADMINISTRATIVE ORDER #").Appearance.BackColor = Color.RosyBrown
            '    End If
            'End If

            If Not ug.Cells("AGREED ORDER #").Value Is DBNull.Value Then
                If ug.Cells("AGREED ORDER #").Value = String.Empty Then
                    ug.Cells("AGREED ORDER #").Value = DBNull.Value
                End If
            End If
            If Not ug.Cells("ADMINISTRATIVE ORDER #").Value Is DBNull.Value Then
                If ug.Cells("ADMINISTRATIVE ORDER #").Value = String.Empty Then
                    ug.Cells("ADMINISTRATIVE ORDER #").Value = DBNull.Value
                End If
            End If

            If ug.Cells("STATUS").Text.Trim <> String.Empty And ug.Cells("ESCALATION").Text.Trim <> String.Empty Then
                If ug.Cells("STATUS").Text.Trim <> ug.Cells("ESCALATION").Text.Trim Then
                    ug.Cells("ESCALATION").Appearance.BackColor = Color.Red
                Else
                    ug.Cells("ESCALATION").Appearance.BackColor = ug.Cells("STATUS").Appearance.BackColor
                End If
            Else
                ug.Cells("ESCALATION").Appearance.BackColor = ug.Cells("STATUS").Appearance.BackColor
            End If

            ' WorkShopResult
            If vListWorkShopResult.FindByDataValue(ug.Cells("WORKSHOP RESULT").Value) Is Nothing Then
                ug.Cells("WORKSHOP RESULT").Value = DBNull.Value
            End If
            ' SHOW CAUSE HEARING RESULT
            If vListShowCauseResult.FindByDataValue(ug.Cells("SHOW CAUSE HEARING RESULT").Value) Is Nothing Then
                ug.Cells("SHOW CAUSE HEARING RESULT").Value = DBNull.Value
            End If
            ' COMMISSION HEARING RESULT
            If vListCommissionResult.FindByDataValue(ug.Cells("COMMISSION HEARING RESULT").Value) Is Nothing Then
                ug.Cells("COMMISSION HEARING RESULT").Value = DBNull.Value
            End If
            ' COMMISSION HEARING RESULT
            If vListAdminHearingResult.FindByDataValue(ug.Cells("ADMIN HEARING RESULT").Value) Is Nothing Then
                ug.Cells("ADMIN HEARING RESULT").Value = DBNull.Value
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Function SaveOCE(ByRef ug As Infragistics.Win.UltraWinGrid.UltraGridRow) As Boolean
        Dim ds As New DataSet
        Dim dt As New DataTable("OCE")
        Dim dr As DataRow
        Dim thisdate As DateTime
        Dim bolReturn As Boolean = True
        Try
            thisdate = New Date(Now.Year, Now.Month, Now.Day, Now.Hour, Now.Minute, Now.Second)

            For Each ugRowCell As Infragistics.Win.UltraWinGrid.UltraGridCell In ug.Cells
                dt.Columns.Add(ugRowCell.Column.Key)
            Next
            dr = dt.NewRow
            For Each ugRowCell As Infragistics.Win.UltraWinGrid.UltraGridCell In ug.Cells
                dr(ugRowCell.Column.Key) = ugRowCell.Value
            Next
            dr("LAST_EDITED_BY") = MusterContainer.AppUser.ID
            dr("DATE_LAST_EDITED") = Now

            dt.Rows.Add(dr)

            dt.Columns("ESCALATION").ColumnName = "STRESCALATION"
            dt.Columns("OCE DATE").ColumnName = "OCE_DATE"
            dt.Columns("LAST PROCESS DATE").ColumnName = "OCE_PROCESS_DATE"
            dt.Columns("NEXT DUE DATE").ColumnName = "NEXT_DUE_DATE"
            dt.Columns("OVERRIDE DUE DATE").ColumnName = "OVERRIDE_DUE_DATE"
            dt.Columns("POLICY AMOUNT").ColumnName = "POLICY_AMOUNT"
            dt.Columns("OVERRIDE AMOUNT").ColumnName = "OVERRIDE_AMOUNT"
            dt.Columns("SETTLEMENT AMOUNT").ColumnName = "SETTLEMENT_AMOUNT"
            dt.Columns("PAID AMOUNT").ColumnName = "PAID_AMOUNT"
            dt.Columns("DATE RECEIVED").ColumnName = "DATE_RECEIVED"
            dt.Columns("WORKSHOP DATE").ColumnName = "WORKSHOP_DATE"
            dt.Columns("WORKSHOP RESULT").ColumnName = "WORKSHOP_RESULT"
            dt.Columns("SHOW CAUSE HEARING DATE").ColumnName = "SHOW_CAUSE_DATE"
            dt.Columns("SHOW CAUSE HEARING RESULT").ColumnName = "SHOW_CAUSE_RESULTS"
            dt.Columns("COMMISSION HEARING DATE").ColumnName = "COMMISSION_DATE"
            dt.Columns("COMMISSION HEARING RESULT").ColumnName = "COMMISSION_RESULTS"
            dt.Columns("AGREED ORDER #").ColumnName = "AGREED_ORDER"
            dt.Columns("ADMINISTRATIVE ORDER #").ColumnName = "ADMINISTRATIVE_ORDER"
            dt.Columns("LETTER GENERATED").ColumnName = "LETTER_GENERATED"
            dt.Columns("RED TAG DATE").ColumnName = "REDTAG_DATE"
            dt.Columns("LETTER PRINTED").ColumnName = "LETTER_PRINTED"
            dt.Columns("ADMIN HEARING DATE").ColumnName = "ADMIN_HEARING_DATE"
            dt.Columns("ADMIN HEARING RESULT").ColumnName = "ADMIN_HEARING_RESULT"


            dt.Columns.Add("ESCALATION")
            dt.Rows(0)("ESCALATION") = dt.Rows(0)("ESCALATION_ID")

            ds.Tables.Add(dt)

            If pOCE.OCECollection.Contains(ds.Tables(0).Rows(0)("OCE_ID").ToString) Then
                pOCE.OCECollection.Remove(ds.Tables(0).Rows(0)("OCE_ID").ToString)
            End If
            pOCE.Load(ds)
            If pOCE.ID <= 0 Then
                pOCE.CreatedBy = MusterContainer.AppUser.ID
            Else
                pOCE.ModifiedBy = MusterContainer.AppUser.ID
            End If
            Dim flagNFA As Short = 0
            If dt.Rows(0)("ESCALATION") = "1248" Then ' 1248 - NFA pending
                Dim strFees As String
                ds = pOCE.RunSQLQuery("SELECT dbo.udfGetOwnerPastDueFees(" + pOCE.OwnerID.ToString + ",0,NULL)")
                If ds.Tables(0).Rows(0)(0) > 0 Then
                    strFees = ds.Tables(0).Rows(0)(0)
                    strFees = strFees.Split(".")(0)
                    If MsgBox("Fees are owed, do you still want to escalate to NFA?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                        flagNFA = 1
                    Else
                        flagNFA = 0
                    End If
                End If
            End If

            bolReturn = pOCE.Save(flagNFA, CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal, True)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Function
            End If

            ' Update Escalation 
            Dim update As Boolean = False
            ug.Cells("ESCALATION").Value = pOCE.EscalationString

            If ug.Cells("ESCALATION_ID").Value <> pOCE.Escalation Then
                update = True
            End If

            ug.Cells("ESCALATION_ID").Value = pOCE.Escalation
            ' if OVERRIDE DUE DATE is modified, change the citation(s) due date
            ModifyOCECitationDueDate(ug, False)
            pOCE.Remove(pOCE.ID)
            pOCE = New MUSTER.BusinessLogic.pOwnerComplianceEvent
            SetugRowComboValue(ug)


            Return bolReturn
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            thisdate = Nothing
        End Try
    End Function
    Private Function SaveOCECitation(ByRef ug As Infragistics.Win.UltraWinGrid.UltraGridRow, ByRef skipQuestion As Integer) As Boolean
        Dim ds As New DataSet
        Dim dt As New DataTable("Citation")
        Dim dr As DataRow
        Dim bolReturn As Boolean = True
        Try
            For Each ugRowCell As Infragistics.Win.UltraWinGrid.UltraGridCell In ug.Cells
                dt.Columns.Add(ugRowCell.Column.Key)
            Next
            dt.Columns.Add("CREATED_BY")
            dt.Columns.Add("DATE_CREATED")
            dt.Columns.Add("LAST_EDITED_BY")
            dt.Columns.Add("DATE_LAST_EDITED")
            dr = dt.NewRow
            For Each ugRowCell As Infragistics.Win.UltraWinGrid.UltraGridCell In ug.Cells
                dr(ugRowCell.Column.Key) = ugRowCell.Value
            Next
            dr("CREATED_BY") = DBNull.Value
            dr("DATE_CREATED") = DBNull.Value
            dr("LAST_EDITED_BY") = MusterContainer.AppUser.ID
            dr("DATE_LAST_EDITED") = Now

            dt.Rows.Add(dr)

            dt.Columns("DUE").ColumnName = "CITATION_DUE_DATE"
            dt.Columns("RECEIVED").ColumnName = "CITATION_RECEIVED_DATE"

            ds.Tables.Add(dt)

            pInsCitation.Load(New MUSTER.Info.InspectionInfo, ds)


            If pInsCitation.ID <= 0 Then
                pInsCitation.CreatedBy = MusterContainer.AppUser.ID
            Else
                pInsCitation.ModifiedBy = MusterContainer.AppUser.ID
            End If

            bolReturn = pInsCitation.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal, True)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Function
            End If

            pInsCitation.Remove(pInsCitation.ID)
            pInsCitation = New MUSTER.BusinessLogic.pInspectionCitation

            GetOCEEscalation(ug.ParentRow, skipQuestion)

            Return bolReturn
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function
    Private Function SaveOCEDiscrep(ByRef ug As Infragistics.Win.UltraWinGrid.UltraGridRow, ByRef skipQuestion As Integer) As Boolean
        Dim bolReturn As Boolean = True
        Try
            pInsDiscrep.Retrieve(New MUSTER.Info.InspectionInfo, ug.Cells("INS_DESCREP_ID").Value)
            pInsDiscrep.Rescinded = ug.Cells("RESCINDED").Value
            If ug.Cells("RECEIVED").Value Is DBNull.Value Then
                pInsDiscrep.DiscrepReceived = CDate("01/01/0001")
            Else
                pInsDiscrep.DiscrepReceived = ug.Cells("RECEIVED").Value
            End If
            pInsDiscrep.ModifiedBy = MusterContainer.AppUser.ID

            bolReturn = pInsDiscrep.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal, True)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Function
            End If

            pInsDiscrep.Remove(pInsDiscrep.ID)
            pInsDiscrep = New MUSTER.BusinessLogic.pInspectionDiscrep

            GetOCEEscalation(ug.ParentRow.ParentRow, skipQuestion)

            Return bolReturn
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function
    Private Sub RescindOCE(ByRef ug As Infragistics.Win.UltraWinGrid.UltraGridRow)
        Try
            ' Scenario 5 - 2
            ug.Cells("POLICY AMOUNT").Value = -1.0
            ug.Cells("SETTLEMENT AMOUNT").Value = -1.0
            ug.Cells("OVERRIDE AMOUNT").Value = -1.0
            ug.Cells("WORKSHOP DATE").Value = CDate("01/01/0001")
            ug.Cells("SHOW CAUSE HEARING DATE").Value = CDate("01/01/0001")
            ug.Cells("COMMISSION HEARING DATE").Value = CDate("01/01/0001")
            ug.Cells("RESCINDED").Value = True
            ug.Cells("PENDING_LETTER").Value = 1251 ' NFA Rescind
            ug.Cells("PENDING_LETTER_TEMPLATE_NUM").Value = UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.NFARescind)
            ug.Cells("OCE_STATUS").Value = 1130 ' NFA Rescind
            ' if there is a row for "nfa rescind" under status band, no need to add row with status "nfa rescind"
            'Dim bolAddNFARescindRow As Boolean = True
            'For Each ugRow In ugEnforcement.Rows ' status
            '    If ugRow.Cells("OCE_STATUS").Value = 1130 Then
            '        bolAddNFARescindRow = False
            '        Exit For
            '    End If
            'Next
            'If bolAddNFARescindRow Then
            '    ugEnforcement.Rows.Band.AddNew()

            '    ugEnforcement.ActiveRow.Cells("OCE_STATUS").Value = 1130
            '    ugEnforcement.ActiveRow.Cells("PROPERTY_POSITION").Value = 11
            '    ugEnforcement.ActiveRow.Cells("SELECTED").Value = 0
            '    ugEnforcement.ActiveRow.Cells("STATUS").Value = "NFA Rescind"
            'End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub RescindOCECitation(ByRef ug As Infragistics.Win.UltraWinGrid.UltraGridRow)
        Try
            ug.Cells("RESCINDED").Value = True
            ug.Cells("RECEIVED").Value = Today.Date
            'If Not ug.ChildBands Is Nothing Then
            '    For Each ugRow In ug.ChildBands(0).Rows
            '        ugRow.Cells("RECEIVED").Value = Today.Date
            '    Next
            'End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Function GenerateOCELetter(ByRef ugOwnerRow As Infragistics.Win.UltraWinGrid.UltraGridRow, Optional ByVal ForceAgreedOrder As Boolean = False) As Boolean

        Dim [continue] As Boolean = False
        Dim dontUpdate As Boolean = False
        Dim letterNum As Integer

        If ForceAgreedOrder Then
            [continue] = True
            letterNum = 48

        ElseIf (Not ugOwnerRow.Cells("OldLetter").Value Is DBNull.Value AndAlso (ugOwnerRow.Cells("PENDING_LETTER_TEMPLATE_NUM").Value Is DBNull.Value OrElse ugOwnerRow.Cells("PENDING_LETTER_TEMPLATE_NUM").Value = 0)) AndAlso Not ForceAgreedOrder Then
            Dim p As New BusinessLogic.pProperty
            [continue] = (MsgBox(String.Format("No Pending Letter to Generate. {0}Would you like to regenerate the last letter (of type {1})?", _
                        vbCrLf, p.GetPropertyNameByID(ugOwnerRow.Cells("OldLetter").Value)), MsgBoxStyle.YesNo) = MsgBoxResult.Yes)

            p = Nothing

            If [continue] Then
                letterNum = ugOwnerRow.Cells("OldLetter").Value
                dontUpdate = True
            End If

            ' Return False

        ElseIf ugOwnerRow.Cells("PENDING_LETTER_TEMPLATE_NUM").Value Is DBNull.Value OrElse ugOwnerRow.Cells("PENDING_LETTER_TEMPLATE_NUM").Value = 0 Then
            MsgBox("No pending Letter to generate")
        Else
            [continue] = True
            letterNum = ugOwnerRow.Cells("PENDING_LETTER_TEMPLATE_NUM").Value
        End If

        If [continue] Then
            Dim bolReturnValue As Boolean = True
            Dim dtFacs, dtCitations, dtDiscreps, dtCOIFacs, dtCorActions As DataTable
            Dim ugCell As Infragistics.Win.UltraWinGrid.UltraGridCell
            Dim ugChildRowLocal, ugGrandChildRowLocal As Infragistics.Win.UltraWinGrid.UltraGridRow
            Dim dr, dr1 As DataRow
            Dim dr2 As DataRow
            Dim nCitationIndex, nDiscrepIndex, nPrevFac As Integer
            Try

                dtFacs = New DataTable
                dtFacs.Columns.Add("FACILITY_ID", GetType(Integer))
                dtFacs.Columns.Add("FACILITY", GetType(String))
                dtFacs.Columns.Add("ADDRESS", GetType(String))
                dtFacs.Columns.Add("ADDRESS_STATE_COUNTY", GetType(String))
                dtFacs.Columns.Add("INSPECTEDON", GetType(Date))
                dtFacs.Columns.Add("ADDRESS_LINE_ONE", GetType(String))
                dtFacs.Columns.Add("CITY", GetType(String))
                dtFacs.Columns.Add("STATE", GetType(String))
                dtFacs.Columns.Add("ZIP", GetType(String))
                dtFacs.Columns.Add("Manager", GetType(String))

                dtCitations = New DataTable
                dtCitations.Columns.Add("FACILITY_ID", GetType(Integer))
                dtCitations.Columns.Add("CITATION_ID", GetType(Integer))

                dtCitations.Columns.Add("CITATION_INDEX", GetType(Integer))
                dtCitations.Columns.Add("CITATIONTEXT", GetType(String))
                dtCitations.Columns.Add("CorrectiveAction", GetType(String))
                dtCitations.Columns.Add("General Corrective Action", GetType(String))

                dtCitations.Columns.Add("DUE", GetType(String))
                dtCitations.Columns.Add("CCAT", GetType(String))
                dtCitations.Columns.Add("CCAT_COMMENTS", GetType(String))
                dtCitations.Columns.Add("INDX", GetType(String))


                dtCorActions = New DataTable

                With dtCorActions
                    .Columns.Add("FACILITY_ID", GetType(Integer))
                    .Columns.Add("CorrectiveAction", GetType(String))
                    .Columns.Add("CITATION_ID", GetType(Integer))

                    .Columns.Add("General Corrective Action", GetType(String))
                    .Columns.Add("DUE", GetType(String))
                    .Columns.Add("INDX", GetType(String))

                End With


                dtDiscreps = New DataTable
                dtDiscreps.Columns.Add("FACILITY_ID", GetType(Integer))

                dtDiscreps.Columns.Add("DISCREP_INDEX", GetType(Integer))
                dtDiscreps.Columns.Add("CITATION_ID", GetType(Integer))
                dtDiscreps.Columns.Add("DISCREP TEXT", GetType(String))
                dtDiscreps.Columns.Add("CorrectiveAction", GetType(String))
                dtDiscreps.Columns.Add("CCAT", GetType(String))
                dtDiscreps.Columns.Add("General Corrective Action", GetType(String))
                dtDiscreps.Columns.Add("CCAT_COMMENTS", GetType(String))
                dtDiscreps.Columns.Add("QUESTION_ID", GetType(Integer))
                dtDiscreps.Columns.Add("INDX", GetType(String))

                dtCOIFacs = New DataTable
                dtCOIFacs.Columns.Add("FACILITY_ID", GetType(Integer))

                strOCEGeneratedLetterName = String.Empty
                nCitationIndex = 0
                nDiscrepIndex = 0
                nPrevFac = 0

                ' get facs for which to include coi letter
                Dim ds As DataSet = pOCE.GetCOIFacs(ugOwnerRow.Cells("OCE_ID").Value, False)
                If ds.Tables.Count > 0 Then
                    If ds.Tables(0).Rows.Count > 0 Then
                        dtCOIFacs = ds.Tables(0)
                        'For Each dr In ds.Tables(0).Rows
                        '    dr1 = dtCOIFacs.NewRow
                        '    dr1("FACILITY_ID") = dr("FACILITY_ID")
                        '    dtCOIFacs.Rows.Add(dr)
                        'Next
                    End If
                End If

                If Not ugOwnerRow.ChildBands Is Nothing Then
                    If Not ugOwnerRow.ChildBands(0).Rows Is Nothing Then
                        For Each ugChildRowLocal In ugOwnerRow.ChildBands(0).Rows ' facilities / citations
                            ' if not rescinded

                            If ugChildRowLocal.Cells("INS_CIT_ID").Value = 16586 Then
                                Dim a
                                a = 1

                            End If
                            If (Not ugChildRowLocal.Cells("RESCINDED").Value) Or (ugChildRowLocal.Cells("RESCINDED").Value And ugChildRowLocal.ParentRow.Cells("STATUS").Text.ToUpper = "NFA RESCIND") Then
                                ' if facility not present in table, add facility details
                                If dtFacs.Select("FACILITY_ID = " + ugChildRowLocal.Cells("FACILITY_ID").Value.ToString).Length = 0 Then
                                    dr = dtFacs.NewRow
                                    dr("FACILITY_ID") = ugChildRowLocal.Cells("FACILITY_ID").Value
                                    dr("FACILITY") = ugChildRowLocal.Cells("FACILITY").Value
                                    dr("ADDRESS") = ugChildRowLocal.Cells("ADDRESS_LINE_ONE").Text + ", " + _
                                                    ugChildRowLocal.Cells("CITY").Text

                                    dr("ADDRESS_STATE_COUNTY") = ugChildRowLocal.Cells("ADDRESS_LINE_ONE").Text + ", " + _
                                                    ugChildRowLocal.Cells("CITY").Text + ", " + _
                                                    IIf(ugChildRowLocal.Cells("STATELONG").Value Is DBNull.Value, ugChildRowLocal.Cells("STATE").Text, ugChildRowLocal.Cells("STATELONG").Text) + ", " + _
                                                    ugChildRowLocal.Cells("COUNTY").Text + " County"

                                    dr("ADDRESS_LINE_ONE") = ugChildRowLocal.Cells("ADDRESS_LINE_ONE").Text
                                    dr("CITY") = ugChildRowLocal.Cells("CITY").Text
                                    dr("STATE") = ugChildRowLocal.Cells("STATE").Text
                                    dr("ZIP") = ugChildRowLocal.Cells("ZIP").Text

                                    If ugChildRowLocal.Cells("INSPECTEDON").Value Is DBNull.Value Then
                                        dr("INSPECTEDON") = CDate("01/01/0001")
                                    ElseIf ugChildRowLocal.Cells("INSPECTEDON").Text = String.Empty Then
                                        dr("INSPECTEDON") = CDate("01/01/0001")
                                    Else
                                        dr("INSPECTEDON") = ugChildRowLocal.Cells("INSPECTEDON").Value
                                    End If

                                    dr("manager") = ugChildRowLocal.Cells("Manager Name").Value

                                    dtFacs.Rows.Add(dr)
                                End If

                                If nPrevFac <> ugChildRowLocal.Cells("FACILITY_ID").Value Then
                                    nPrevFac = ugChildRowLocal.Cells("FACILITY_ID").Value
                                    nCitationIndex = 0
                                    If dtCitations.Select("FACILITY_ID = " + nPrevFac.ToString).Length > 0 Then
                                        For Each dr In dtCitations.Select("FACILITY_ID = " + nPrevFac.ToString)
                                            If nCitationIndex > dr("CITATION_INDEX") Then
                                                nCitationIndex = dr("CITATION_INDEX")
                                            End If
                                        Next
                                    End If
                                    nDiscrepIndex = 0
                                    If dtDiscreps.Select("FACILITY_ID = " + nPrevFac.ToString).Length > 0 Then
                                        For Each dr In dtDiscreps.Select("FACILITY_ID = " + nPrevFac.ToString)
                                            If nDiscrepIndex > dr("DISCREP_INDEX") Then
                                                nDiscrepIndex = dr("DISCREP_INDEX")
                                            End If
                                        Next
                                    End If
                                End If

                                ' if citation is discrepancy - use discrepancy text as citation text. if discrepancy does not exists, do not add citation
                                ' if citation not present in table, add citation
                                If ugChildRowLocal.Cells("CITATION_ID").Value = 19 Or ugChildRowLocal.Cells("CITATION_ID").Value = 27 Or ugChildRowLocal.Cells("CITATION_ID").Value = 28 Then
                                    If Not ugChildRowLocal.ChildBands Is Nothing Then
                                        If Not ugChildRowLocal.ChildBands(0).Rows Is Nothing Then
                                            For Each ugGrandChildRowLocal In ugChildRowLocal.ChildBands(0).Rows
                                                If Not ugGrandChildRowLocal.Cells("RESCINDED").Value Then
                                                    If ugGrandChildRowLocal.Cells("CITATION_ID").Value = 19 Or ugGrandChildRowLocal.Cells("CITATION_ID").Value = 27 Or ugGrandChildRowLocal.Cells("CITATION_ID").Value = 28 Then ' Category Discrepancy
                                                        If ugGrandChildRowLocal.Cells("DISCREP TEXT").Text.Length > 0 Then

                                                            Dim rows As DataRow() = dtCitations.Select("FACILITY_ID = " + ugChildRowLocal.Cells("FACILITY_ID").Text + " AND CITATIONTEXT = '" + ugGrandChildRowLocal.Cells("DISCREP TEXT").Text.Replace("'", "") + "'")

                                                            dr2 = dtCorActions.NewRow
                                                            dr2("FACILITY_ID") = ugChildRowLocal.Cells("FACILITY_ID").Value
                                                            dr2("CorrectiveAction") = ugGrandChildRowLocal.Cells("CorrectiveAction").Text.Trim
                                                            dr2("INDX") = ugChildRowLocal.Cells("INDX").Text.Trim


                                                            dr2("General Corrective Action") = ugChildRowLocal.Cells("General Corrective Action").Text.Trim
                                                            dr2("Citation_ID") = ugChildRowLocal.Cells("Citation_ID").Text.Trim

                                                            If rows.Length = 0 Then
                                                                nCitationIndex += 1
                                                                dr = dtCitations.NewRow

                                                                dr("FACILITY_ID") = ugChildRowLocal.Cells("FACILITY_ID").Value

                                                                dr("CITATION_INDEX") = nCitationIndex
                                                                dr("CITATIONTEXT") = ugGrandChildRowLocal.Cells("DISCREP TEXT").Text.Trim
                                                                dr("Citation_ID") = ugChildRowLocal.Cells("Citation_ID").Text.Trim
                                                                dr("CorrectiveAction") = ugGrandChildRowLocal.Cells("CorrectiveAction").Text.Trim
                                                                dr("General Corrective Action") = ugChildRowLocal.Cells("General Corrective Action").Text.Trim
                                                                dr("INDX") = ugChildRowLocal.Cells("INDX").Text.Trim

                                                                dr("CCAT") = ugGrandChildRowLocal.Cells("CCAT").Text.Trim
                                                                dr("CCAT_COMMENTS") = ugChildRowLocal.Cells("CCAT_COMMENTS").Text.Trim

                                                                If Not ugChildRowLocal.Cells("DUE").Value Is DBNull.Value Then
                                                                    If ugChildRowLocal.Cells("DUE").Text = String.Empty Then
                                                                        dr("DUE") = "-N/A-"
                                                                        dr2("DUE") = "-N/A-"

                                                                    ElseIf Date.Compare(ugChildRowLocal.Cells("DUE").Value, CDate("01/01/0001")) = 0 Then
                                                                        dr("DUE") = "-N/A-"
                                                                        dr2("DUE") = "-N/A-"

                                                                    Else
                                                                        dr("DUE") = CDate(ugChildRowLocal.Cells("DUE").Value).ToShortDateString
                                                                        dr2("DUE") = CDate(ugChildRowLocal.Cells("DUE").Value).ToShortDateString

                                                                    End If
                                                                Else
                                                                    dr("DUE") = "-N/A-"
                                                                    dr2("DUE") = "-N/A-"

                                                                End If
                                                                dtCitations.Rows.Add(dr)
                                                                dtCorActions.Rows.Add(dr2)
                                                            Else
                                                                ' rows(0).Item("CCAT") = String.Format("{0}{1}{2}", rows(0).Item("CCAT"), IIf(rows(0).Item("CCAT") <> String.Empty, ", ", String.Empty), ugGrandChildRowLocal.Cells("CCAT").Text.Trim)
                                                                ' rows(0).Item("CCAT_COMMENTS") = String.Format("{0}{1}{2}", rows(0).Item("CCAT_COMMENTS"), IIf(rows(0).Item("CCAT_COMMENTS") <> String.Empty, ", ", String.Empty), ugChildRowLocal.Cells("CCAT_COMMENTS").Text.Trim)

                                                                If Not ugChildRowLocal.Cells("DUE").Value Is DBNull.Value Then
                                                                    If ugChildRowLocal.Cells("DUE").Text = String.Empty Then
                                                                        dr2("DUE") = "-N/A-"

                                                                    ElseIf Date.Compare(ugChildRowLocal.Cells("DUE").Value, CDate("01/01/0001")) = 0 Then
                                                                        dr2("DUE") = "-N/A-"

                                                                    Else
                                                                        dr2("DUE") = CDate(ugChildRowLocal.Cells("DUE").Value).ToShortDateString

                                                                    End If
                                                                Else
                                                                    dr("DUE") = "-N/A-"
                                                                    dr2("DUE") = "-N/A-"
                                                                End If
                                                                dtCorActions.Rows.Add(dr2)


                                                            End If ' If dtCitations.Select("FACILITY_ID = " + ugChildRowLocal.Cells("FACILITY_ID").Text + " AND CITATIONTEXT = '" + ugGrandChildRowLocal.Cells("DISCREP TEXT").Text + "'").Length = 0 Then

                                                        End If ' If ugGrandChildRowLocal.Cells("DISCREP TEXT").Text.Length > 0 Then
                                                    End If ' If ugGrandChildRowLocal.Cells("CITATION_ID").Value = 19 Then ' Category Discrepancy
                                                End If ' If Not ugGrandChildRowLocal.Cells("RESCINDED").Value Then
                                            Next ' For Each ugGrandChildRowLocal In ugChildRowLocal.ChildBands(0).Rows
                                        End If ' If Not ugChildRowLocal.ChildBands(0).Rows Is Nothing Then
                                    End If ' If Not ugChildRowLocal.ChildBands Is Nothing Then
                                Else
                                    Dim rows As DataRow() = dtCitations.Select("FACILITY_ID = " + ugChildRowLocal.Cells("FACILITY_ID").Text + " AND CITATIONTEXT = '" + ugChildRowLocal.Cells("CITATIONTEXT").Text.Replace("'", "") + "'")

                                    dr2 = dtCorActions.NewRow
                                    dr2("CorrectiveAction") = ugChildRowLocal.Cells("CorrectiveAction").Text.Trim
                                    dr2("General Corrective Action") = ugChildRowLocal.Cells("General Corrective Action").Text.Trim
                                    dr2("Citation_ID") = ugChildRowLocal.Cells("Citation_ID").Text.Trim

                                    dr2("INDX") = ugChildRowLocal.Cells("INDX").Text.Trim
                                    dr2("FACILITY_ID") = ugChildRowLocal.Cells("FACILITY_ID").Value

                                    If rows.Length = 0 Then
                                        nCitationIndex += 1
                                        dr = dtCitations.NewRow

                                        dr("FACILITY_ID") = ugChildRowLocal.Cells("FACILITY_ID").Value

                                        dr("CITATION_INDEX") = nCitationIndex
                                        dr("INDX") = ugChildRowLocal.Cells("INDX").Text.Trim

                                        dr("CITATIONTEXT") = ugChildRowLocal.Cells("CITATIONTEXT").Text.Trim
                                        dr("CorrectiveAction") = ugChildRowLocal.Cells("CorrectiveAction").Text.Trim
                                        dr("General Corrective Action") = ugChildRowLocal.Cells("General Corrective Action").Text.Trim

                                        dr("Citation_ID") = ugChildRowLocal.Cells("Citation_ID").Text.Trim

                                        dr("CCAT") = ugChildRowLocal.Cells("CCAT").Text.Trim
                                        dr("CCAT_COMMENTS") = ugChildRowLocal.Cells("CCAT_COMMENTS").Text.Trim
                                        If Not ugChildRowLocal.Cells("DUE").Value Is DBNull.Value Then
                                            If ugChildRowLocal.Cells("DUE").Text = String.Empty Then
                                                dr("DUE") = "-N/A-"
                                                dr2("DUE") = "-N/A-"

                                            ElseIf Date.Compare(ugChildRowLocal.Cells("DUE").Value, CDate("01/01/0001")) = 0 Then
                                                dr("DUE") = "-N/A-"
                                                dr2("DUE") = "-N/A-"

                                            Else
                                                dr("DUE") = CDate(ugChildRowLocal.Cells("DUE").Value).ToShortDateString
                                                dr2("DUE") = CDate(ugChildRowLocal.Cells("DUE").Value).ToShortDateString

                                            End If
                                        Else
                                            dr("DUE") = "-N/A-"
                                            dr2("DUE") = "-N/A-"

                                        End If
                                        dtCitations.Rows.Add(dr)
                                        dtCorActions.Rows.Add(dr2)
                                    Else

                                        'rows(0).Item("CCAT") = String.Format("{0}{1}{2}", rows(0).Item("CCAT"), IIf(rows(0).Item("CCAT") <> String.Empty, ", ", String.Empty), ugGrandChildRowLocal.Cells("CCAT").Text.Trim)
                                        'rows(0).Item("CCAT_COMMENTS") = String.Format("{0}{1}{2}", rows(0).Item("CCAT_COMMENTS"), IIf(rows(0).Item("CCAT_COMMENTS") <> String.Empty, ", ", String.Empty), ugChildRowLocal.Cells("CCAT_COMMENTS").Text.Trim)


                                        If Not ugChildRowLocal.Cells("DUE").Value Is DBNull.Value Then
                                            If ugChildRowLocal.Cells("DUE").Text = String.Empty Then
                                                dr2("DUE") = "-N/A-"

                                            ElseIf Date.Compare(ugChildRowLocal.Cells("DUE").Value, CDate("01/01/0001")) = 0 Then
                                                dr2("DUE") = "-N/A-"

                                            Else
                                                dr2("DUE") = CDate(ugChildRowLocal.Cells("DUE").Value).ToShortDateString

                                            End If
                                        Else
                                            dr2("DUE") = "-N/A-"

                                        End If

                                        dtCorActions.Rows.Add(dr2)

                                    End If
                                End If ' If ugChildRowLocal.Cells("CITATION_ID").Value = 19 Then

                                If Not ugChildRowLocal.ChildBands Is Nothing Then
                                    If Not ugChildRowLocal.ChildBands(0).Rows Is Nothing Then
                                        For Each ugGrandChildRowLocal In ugChildRowLocal.ChildBands(0).Rows
                                            If Not ugGrandChildRowLocal.Cells("RESCINDED").Value Then
                                                If ugGrandChildRowLocal.Cells("CITATION_ID").Value = 19 Or ugGrandChildRowLocal.Cells("CITATION_ID").Value = 27 Or ugGrandChildRowLocal.Cells("CITATION_ID").Value = 28 Then ' Category Discrepancy
                                                    Dim rows() As DataRow = dtDiscreps.Select("FACILITY_ID = " + ugChildRowLocal.Cells("FACILITY_ID").Text + " AND CITATION_ID = " + ugGrandChildRowLocal.Cells("CITATION_ID").Text + " AND [DISCREP TEXT] = '" + ugGrandChildRowLocal.Cells("DISCREP TEXT").Text.Replace("'", "") + "'")

                                                    dr2 = dtCorActions.NewRow
                                                    dr2("FACILITY_ID") = ugChildRowLocal.Cells("FACILITY_ID").Value
                                                    dr2("CorrectiveAction") = ugGrandChildRowLocal.Cells("CorrectiveAction").Text.Trim
                                                    dr2("General Corrective Action") = ugChildRowLocal.Cells("General Corrective Action").Text.Trim
                                                    dr2("INDX") = ugChildRowLocal.Cells("INDX").Text.Trim
                                                    dr2("Citation_ID") = ugChildRowLocal.Cells("Citation_ID").Text.Trim
                                                    dr2("DUE") = "-N/A-"
                                                    If rows.Length = 0 Then
                                                        nDiscrepIndex += 1
                                                        dr = dtDiscreps.NewRow
                                                        dr("FACILITY_ID") = ugChildRowLocal.Cells("FACILITY_ID").Value
                                                        dr("INDX") = ugChildRowLocal.Cells("INDX").Text.Trim
                                                        dr("DISCREP_INDEX") = nDiscrepIndex
                                                        dr("CITATION_ID") = ugGrandChildRowLocal.Cells("CITATION_ID").Value
                                                        dr("QUESTION_ID") = ugGrandChildRowLocal.Cells("QUESTION_ID").Value
                                                        dr("DISCREP TEXT") = ugGrandChildRowLocal.Cells("DISCREP TEXT").Text.Trim
                                                        dr("CorrectiveAction") = ugGrandChildRowLocal.Cells("CorrectiveAction").Text.Trim
                                                        dr("General Corrective Action") = ugChildRowLocal.Cells("CorrectiveAction").Text.Trim

                                                        dr("Citation_ID") = ugChildRowLocal.Cells("Citation_ID").Text.Trim


                                                        dr("CCAT") = ugGrandChildRowLocal.Cells("CCAT").Text.Trim
                                                        dr("CCAT_COMMENTS") = ugChildRowLocal.Cells("CCAT_COMMENTS").Text.Trim
                                                        dtDiscreps.Rows.Add(dr)
                                                    Else
                                                        ' rows(0).Item("CCAT") = String.Format("{0}{1}{2}", rows(0).Item("CCAT"), IIf(rows(0).Item("CCAT") <> String.Empty, ", ", String.Empty), ugGrandChildRowLocal.Cells("CCAT").Text.Trim)
                                                        ' rows(0).Item("CCAT_COMMENTS") = String.Format("{0}{1}{2}", rows(0).Item("CCAT_COMMENTS"), IIf(rows(0).Item("CCAT_COMMENTS") <> String.Empty, ", ", String.Empty), ugChildRowLocal.Cells("CCAT_COMMENTS").Text.Trim)
                                                    End If

                                                    If dtCorActions.Select(String.Format("CorrectiveAction = '{0}' and Facility_ID ={1}", dr("CorrectiveAction"), ugChildRowLocal.Cells("FACILITY_ID").Value)).GetUpperBound(0) < 0 Then
                                                        dtCorActions.Rows.Add(dr2)
                                                    End If


                                                End If
                                            End If
                                        Next
                                    End If
                                End If
                            End If ' If Not ugChildRowLocal.Cells("RESCINDED").Value Then
                        Next
                    End If
                End If

                ' generate letter
                Dim regletter As New Reg_Letters
                'Dim documentID As Integer = 0
                bolReturnValue = regletter.GenerateCAEOCELetter(ugOwnerRow, dtFacs, dtCitations, dtDiscreps, dtCorActions, dtCOIFacs, letterNum, strOCEGeneratedLetterName, pOCE.GetLetterGeneratedDate(ugOwnerRow.Cells("OCE_ID").Value, UIUtilsGen.EntityTypes.CAEOwnerComplianceEvent, , False), returnVal, pOCE)

                'Exit Function

                If bolReturnValue AndAlso Not dontUpdate AndAlso Not ForceAgreedOrder Then
                    '    ' letter generated date
                    '    ' only for the first created letter (OCE Creation / Modification), the date is saved / referred
                    '    Dim bolSaveLetterGenDate As Boolean = False
                    '    Select Case ugOwnerRow.Cells("PENDING_LETTER_TEMPLATE_NUM").Value
                    '        Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.DiscrepanciesOnly) ' 1
                    '            bolSaveLetterGenDate = True
                    '        Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.CAT3_NoPrior_NOV) ' 2
                    '            bolSaveLetterGenDate = True
                    '        Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.CAT2_NoPrior_NOV_Workshop) ' 3
                    '            bolSaveLetterGenDate = True
                    '        Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.CAT2_1_CAT3_NOV_Workshop, True) ' 3
                    '            bolSaveLetterGenDate = True
                    '        Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.CAT1_CAT1_CAT2_1_CAT3_NOV_AgreedOrder) ' 4
                    '            bolSaveLetterGenDate = True
                    '        Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.CAT1_1_CAT3_NOV_Workshop_AgreedOrder) ' 5
                    '            bolSaveLetterGenDate = True
                    '        Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.CAT1_NoPrior_NOV_Workshop_AgreedOrder, True) ' 5
                    '            bolSaveLetterGenDate = True
                    '        Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.CAT2_CAT1_CAT2_1_CAT3_NOV_AgreedOrder) ' 6
                    '            bolSaveLetterGenDate = True
                    '        Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.CAT3_CAT1_CAT2_1_CAT3_NOV_AgreedOrder) ' 7
                    '            bolSaveLetterGenDate = True
                    '        Case UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.OCELetterTemplateNum.CAT3_1_CAT3_NOV_Workshop) ' 8
                    '            bolSaveLetterGenDate = True
                    '        Case Else
                    '            bolSaveLetterGenDate = False
                    '    End Select

                    '    If bolSaveLetterGenDate Then
                    '        pOCE.SaveLetterGeneratedDate(0, ugOwnerRow.Cells("OCE_ID").Value, UIUtilsGen.EntityTypes.CAEOwnerComplianceEvent, ugOwnerRow.Cells("PENDING_LETTER_TEMPLATE_NUM").Value, Today.Date, False, documentID, UIUtilsGen.ModuleID.CAE, MusterContainer.AppUser.UserKey, returnVal)
                    '        If Not UIUtilsGen.HasRights(returnVal) Then
                    '            Return False
                    '        End If
                    '    End If

                    ' If there is an error, need a way to rollback / revert changes
                    ugOwnerRow.Cells("PENDING_LETTER").Tag = ugOwnerRow.Cells("PENDING_LETTER").Value
                    ugOwnerRow.Cells("PENDING LETTER").Tag = ugOwnerRow.Cells("PENDING LETTER").Value
                    ugOwnerRow.Cells("LETTER PRINTED").Tag = ugOwnerRow.Cells("LETTER PRINTED").Value
                    ugOwnerRow.Cells("LETTER GENERATED").Tag = ugOwnerRow.Cells("LETTER GENERATED").Value
                    '   ugOwnerRow.Cells("RED TAG DATE").Tag = ugOwnerRow.Cells("RED TAG DATE").Value
                    ugOwnerRow.Cells("PENDING_LETTER_TEMPLATE_NUM").Tag = ugOwnerRow.Cells("PENDING_LETTER_TEMPLATE_NUM").Value
                    ' Setting Letter Generated = current date
                    ' Setting Letter Printed = False
                    ugOwnerRow.Cells("PENDING_LETTER").Value = DBNull.Value
                    ugOwnerRow.Cells("PENDING LETTER").Value = String.Empty
                    ugOwnerRow.Cells("LETTER PRINTED").Value = False
                    ugOwnerRow.Cells("LETTER GENERATED").Value = Today.Date
                    ugOwnerRow.Cells("PENDING_LETTER_TEMPLATE_NUM").Value = DBNull.Value
                End If
                Return bolReturnValue
            Catch ex As Exception
                Dim MyErr As New ErrorReport(ex)
                MyErr.ShowDialog()
            End Try
        End If

    End Function
    Private Function ReverseGenerateOCELetter(ByRef ugOwnerRow As Infragistics.Win.UltraWinGrid.UltraGridRow, ByVal bolWasLetterGenerated As Boolean) As Boolean
        Dim bolReturnValue As Boolean = True
        Try
            ' If there is an error, need a way to rollback / revert changes
            ugOwnerRow.Cells("PENDING_LETTER").Value = ugOwnerRow.Cells("PENDING_LETTER").Tag
            ugOwnerRow.Cells("PENDING LETTER").Value = ugOwnerRow.Cells("PENDING LETTER").Tag
            ugOwnerRow.Cells("LETTER PRINTED").Value = ugOwnerRow.Cells("LETTER PRINTED").Tag
            ugOwnerRow.Cells("LETTER GENERATED").Value = ugOwnerRow.Cells("LETTER GENERATED").Tag
            '  ugOwnerRow.Cells("RED TAG DATE").Value = ugOwnerRow.Cells("RED TAG DATE").Tag
            ugOwnerRow.Cells("PENDING_LETTER_TEMPLATE_NUM").Value = ugOwnerRow.Cells("PENDING_LETTER_TEMPLATE_NUM").Tag

            ' need to delete the letter if it was generated
            If bolWasLetterGenerated Then
                If strOCEGeneratedLetterName <> String.Empty Then
                    ' before deleting, make sure the letter is not open. if open, close letter and then delete
                    Dim strDOCPATH As String = MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_SystemGenerated).ProfileValue & "\"
                    If System.IO.File.Exists(strDOCPATH + "\" + strOCEGeneratedLetterName) Then
                        Dim wordApp As Word.Application = MusterContainer.GetWordApp

                        If Not wordApp Is Nothing Then
                            wordApp.Documents.Open(strOCEGeneratedLetterName)
                            wordApp.ActiveDocument.Close(False)
                            System.IO.File.Delete(strDOCPATH + "\" + strOCEGeneratedLetterName)
                            ' delete from doc manager
                            UIUtilsGen.DeleteDocument(strOCEGeneratedLetterName, MusterContainer.AppUser.ID, False)
                        End If
                        wordApp = Nothing
                    End If
                End If
            End If
            Return bolReturnValue
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function

    'Private Sub btnEnforceProcessEscalations_ClickOld(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim drow As Infragistics.Win.UltraWinGrid.UltraGridRow
    '    Dim bol As Boolean
    '    Try
    '        For Each drow In ugEnforcement.Rows
    '            For Each childrow As Infragistics.Win.UltraWinGrid.UltraGridRow In drow.ChildBands("StatustoOCE".ToUpper).Rows
    '                If childrow.Cells("Selected").Value Then
    '                    bol = True
    '                End If
    '            Next
    '        Next
    '        If Not bol Then
    '            MsgBox("Please select the OCE to be processed")
    '            Exit Sub
    '        End If
    '        For Each drow In ugLicensees.Rows
    '            For Each childrow As Infragistics.Win.UltraWinGrid.UltraGridRow In drow.ChildBands("StatustoOCE".ToUpper).Rows
    '                If childrow.Cells("Selected").Value Then
    '                    ' TO DO
    '                    pOCE.Retrieve(childrow.Cells("OCE_ID").Value)
    '                    pOCE.EscalationLogic(pOCE.OCEInfo)
    '                    pOCE.Save()
    '                End If
    '            Next
    '        Next
    '        ugEnforcement.DataSource = Nothing
    '        ugEnforcement.DataSource = pOCE.EntityTable()
    '        ugEnforcement.Rows.ExpandAll(True)
    '        MsgBox("Processing(escalating) the selected OCE's is done successfully")
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub ugEnforcement_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugEnforcement.InitializeLayout

    '    If ugEnforcement.DisplayLayout.Bands(1).Groups.Count = 0 Then
    '        ugEnforcement.DisplayLayout.Bands(1).Groups.Add("Row1")
    '        ugEnforcement.DisplayLayout.Bands(1).Groups.Add("Row2")
    '    End If

    '    'After you have bound your grid to a DataSource you should create an unbound column that will be used as your CheckBox column. In the InitializeLayout event add the following code to create an unbound column:
    '    Me.ugEnforcement.DisplayLayout.Bands(0).Columns.Add("Selected").DataType = GetType(Boolean)
    '    Me.ugEnforcement.DisplayLayout.Bands(0).Columns("Selected").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
    '    Me.ugEnforcement.DisplayLayout.Bands(0).Columns("Selected").Header.VisiblePosition = 0

    '    Me.ugEnforcement.DisplayLayout.Bands(1).Columns.Add("Selected").DataType = GetType(Boolean)
    '    Me.ugEnforcement.DisplayLayout.Bands(1).Columns("Selected").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
    '    Me.ugEnforcement.DisplayLayout.Bands(1).Columns("Selected").Header.VisiblePosition = 0


    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Selected").Group = ugEnforcement.DisplayLayout.Bands(1).Groups("Row1")
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("FILLER5").Group = ugEnforcement.DisplayLayout.Bands(1).Groups("Row1")
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Owner").Group = ugEnforcement.DisplayLayout.Bands(1).Groups("Row1")
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("FILLER1").Group = ugEnforcement.DisplayLayout.Bands(1).Groups("Row1")
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Rescinded").Group = ugEnforcement.DisplayLayout.Bands(1).Groups("Row1")
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("FILLER2").Group = ugEnforcement.DisplayLayout.Bands(1).Groups("Row1")
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("OCE" + vbCrLf + "Date").Group = ugEnforcement.DisplayLayout.Bands(1).Groups("Row1")
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Last" + vbCrLf + "Process Date").Group = ugEnforcement.DisplayLayout.Bands(1).Groups("Row1")
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Next" + vbCrLf + "Due Date").Group = ugEnforcement.DisplayLayout.Bands(1).Groups("Row1")
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Override" + vbCrLf + "Due Date").Group = ugEnforcement.DisplayLayout.Bands(1).Groups("Row1")
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("OCE_Status").Group = ugEnforcement.DisplayLayout.Bands(1).Groups("Row1")
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Escalation").Group = ugEnforcement.DisplayLayout.Bands(1).Groups("Row1")
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Policy" + vbCrLf + "Amount").Group = ugEnforcement.DisplayLayout.Bands(1).Groups("Row1")
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Override" + vbCrLf + "Amount").Group = ugEnforcement.DisplayLayout.Bands(1).Groups("Row1")
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Settlement" + vbCrLf + "Amount").Group = ugEnforcement.DisplayLayout.Bands(1).Groups("Row1")
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("FILLER3").Group = ugEnforcement.DisplayLayout.Bands(1).Groups("Row1")
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Paid" + vbCrLf + "Amount").Group = ugEnforcement.DisplayLayout.Bands(1).Groups("Row1")
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Date" + vbCrLf + "Received").Group = ugEnforcement.DisplayLayout.Bands(1).Groups("Row1")

    '    ugEnforcement.DisplayLayout.Bands(1).Columns("WorkShop" + vbCrLf + "Date").Group = ugEnforcement.DisplayLayout.Bands(1).Groups("Row2")
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("WorkShop" + vbCrLf + "Result").Group = ugEnforcement.DisplayLayout.Bands(1).Groups("Row2")
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Show Cause" + vbCrLf + "Hearing Date").Group = ugEnforcement.DisplayLayout.Bands(1).Groups("Row2")
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Show Cause" + vbCrLf + "Hearing Results").Group = ugEnforcement.DisplayLayout.Bands(1).Groups("Row2")
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Commission" + vbCrLf + "Hearing Date").Group = ugEnforcement.DisplayLayout.Bands(1).Groups("Row2")
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Commission" + vbCrLf + "Hearing Results").Group = ugEnforcement.DisplayLayout.Bands(1).Groups("Row2")
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Agreed" + vbCrLf + "Order #").Group = ugEnforcement.DisplayLayout.Bands(1).Groups("Row2")
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Administrative" + vbCrLf + "Order #").Group = ugEnforcement.DisplayLayout.Bands(1).Groups("Row2")
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Pending" + vbCrLf + "Letter").Group = ugEnforcement.DisplayLayout.Bands(1).Groups("Row2")
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Letter" + vbCrLf + "Printed").Group = ugEnforcement.DisplayLayout.Bands(1).Groups("Row2")
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Letter" + vbCrLf + "Generated").Group = ugEnforcement.DisplayLayout.Bands(1).Groups("Row2")
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("FILLER4").Group = ugEnforcement.DisplayLayout.Bands(1).Groups("Row2")

    '    ugEnforcement.DisplayLayout.Bands(1).LevelCount = 2
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Selected").Level = 0
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("FILLER5").Level = 1
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Owner").Level = 0
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("FILLER1").Level = 1
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Rescinded").Level = 0
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("FILLER2").Level = 1
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("OCE" + vbCrLf + "Date").Level = 0
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Last" + vbCrLf + "Process Date").Level = 1
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Next" + vbCrLf + "Due Date").Level = 0
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Override" + vbCrLf + "Due Date").Level = 1
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("OCE_Status").Level = 0
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Escalation").Level = 1
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Policy" + vbCrLf + "Amount").Level = 0
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Override" + vbCrLf + "Amount").Level = 1
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Settlement" + vbCrLf + "Amount").Level = 0
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("FILLER3").Level = 1
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Paid" + vbCrLf + "Amount").Level = 0
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Date" + vbCrLf + "Received").Level = 1
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("WorkShop" + vbCrLf + "Date").Level = 0
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("WorkShop" + vbCrLf + "Result").Level = 1
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Show Cause" + vbCrLf + "Hearing Date").Level = 0
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Show Cause" + vbCrLf + "Hearing Results").Level = 1
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Commission" + vbCrLf + "Hearing Results").Level = 1
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Commission" + vbCrLf + "Hearing Date").Level = 0
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Agreed" + vbCrLf + "Order #").Level = 0
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Administrative" + vbCrLf + "Order #").Level = 1
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Pending" + vbCrLf + "Letter").Level = 0
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Letter" + vbCrLf + "Printed").Level = 1
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Letter" + vbCrLf + "Generated").Level = 0
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("FILLER4").Level = 1

    '    ugEnforcement.DisplayLayout.Bands(1).Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Free
    '    ugEnforcement.DisplayLayout.Bands(1).Override.CellMultiLine = Infragistics.Win.DefaultableBoolean.True
    '    ugEnforcement.DisplayLayout.Bands(1).Override.DefaultColWidth = 80
    '    ugEnforcement.DisplayLayout.Bands(1).ColHeaderLines = 2
    '    ugEnforcement.DisplayLayout.Bands(1).Override.RowSelectorWidth = 2
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("Owner").CellMultiLine = Infragistics.Win.DefaultableBoolean.False

    '    ugEnforcement.DisplayLayout.Bands(1).Columns("FILLER1").Header.Caption = ""
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("FILLER1").CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("FILLER1").CellAppearance.BackColorDisabled = Color.LightGoldenrodYellow

    '    ugEnforcement.DisplayLayout.Bands(1).Columns("FILLER2").Header.Caption = ""
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("FILLER2").CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("FILLER2").CellAppearance.BackColorDisabled = Color.LightGoldenrodYellow

    '    ugEnforcement.DisplayLayout.Bands(1).Columns("FILLER3").Header.Caption = ""
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("FILLER3").CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("FILLER3").CellAppearance.BackColorDisabled = Color.LightGoldenrodYellow
    '    ugEnforcement.DisplayLayout.Bands(1).GroupHeadersVisible = False

    '    ugEnforcement.DisplayLayout.Bands(1).Columns("FILLER4").Header.Caption = ""
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("FILLER4").CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("FILLER4").CellAppearance.BackColorDisabled = Color.LightGoldenrodYellow

    '    ugEnforcement.DisplayLayout.Bands(1).Columns("FILLER5").Header.Caption = ""
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("FILLER5").CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled
    '    ugEnforcement.DisplayLayout.Bands(1).Columns("FILLER5").CellAppearance.BackColorDisabled = Color.LightGoldenrodYellow

    '    ugEnforcement.DisplayLayout.Bands(1).Columns("OCE_ID").Hidden = True

    'End Sub
    Private Sub btnEnforceProcessEscalations_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEnforceProcessEscalations.Click
        If bolLoading Then Exit Sub
        Dim bolOCEEscalated As Boolean
        Dim skipQuestion As Integer = 0
        Dim thisDate As DateTime
        Try
            thisDate = New Date(Now.Year, Now.Month, Now.Day, Now.Hour, Now.Minute, Now.Second)


            If Not MusterContainer.AppUser.HasAccess(UIUtilsGen.ModuleID.CAE, MusterContainer.AppUser.UserKey, UIUtilsGen.EntityTypes.CAEOwnerComplianceEvent) Then
                MessageBox.Show("You do not have rights to save Owner Compliance Event.", "ACCESS DENIED", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            bolLoading = True
            'OCEEscalationError = False
            bolOCEEscalated = False
            For Each ugRow In ugEnforcement.Rows ' status
                For Each ugChildRow In ugRow.ChildBands(0).Rows ' owner
                    If ugChildRow.Cells("SELECTED").Value = True Or _
                        ugChildRow.Cells("SELECTED").Text = True Then
                        OCECancelEscalation = False
                        ' Execute OCE Escalation Logic and update the Grid Row
                        ProcessOCEEscalationLogic(ugChildRow, , , skipQuestion)
                        If OCECancelEscalation Then
                            OCECancelEscalation = False
                            Exit Sub
                        Else
                            bolOCEEscalated = True
                        End If
                    End If
                Next
            Next
            If bolOCEEscalated Then

                ProcessRedtagChanges(thisDate)

                MsgBox("OCE(s) escalated Successfully")
                ' refresh grid
                SetupTabs()
            Else
                MsgBox("No OCE(s) selected / escalated")
            End If
            OCECancelEscalation = False
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            thisDate = Nothing
            bolLoading = False
        End Try
    End Sub


    Private Sub btnEnforceProcessManualEsc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnManualEsc.Click
        If bolLoading Then Exit Sub
        Dim bolOCEEscalated As Boolean
        Dim skipQuestion As Integer = 0
        Dim thisDate As DateTime
        Try
            thisDate = New Date(Now.Year, Now.Month, Now.Day, Now.Hour, Now.Minute, Now.Second)


            If Not MusterContainer.AppUser.HasAccess(UIUtilsGen.ModuleID.CAE, MusterContainer.AppUser.UserKey, UIUtilsGen.EntityTypes.CAEOwnerComplianceEvent) Then
                MessageBox.Show("You do not have rights to save Owner Compliance Event.", "ACCESS DENIED", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            Dim code As Integer = 0
            Dim frmDiaglog As New CandEStatusSelection
            frmDiaglog.ShowDialog()

            If frmDiaglog.DialogResult = DialogResult.OK AndAlso IsNumeric(frmDiaglog.SelectedButton.Tag) Then
                code = -1 - Convert.ToInt16(frmDiaglog.SelectedButton.Tag)
            End If

            frmDiaglog = Nothing

            If code = 0 Then
                Exit Sub
            End If


            bolLoading = True
            'OCEEscalationError = False
            bolOCEEscalated = False

            For Each ugRow In ugEnforcement.Rows ' status
                For Each ugChildRow In ugRow.ChildBands(0).Rows ' owner
                    If ugChildRow.Cells("SELECTED").Value = True Or _
                        ugChildRow.Cells("SELECTED").Text = True Then
                        OCECancelEscalation = False
                        ' Execute OCE Escalation Logic and update the Grid Row
                        ProcessOCEEscalationLogicManual(ugChildRow, , , skipQuestion, code)
                        If OCECancelEscalation Then
                            OCECancelEscalation = False
                            Exit Sub
                        Else
                            bolOCEEscalated = True
                        End If
                    End If
                Next
            Next
            If bolOCEEscalated Then

                ProcessRedtagChanges(thisDate)

                MsgBox("OCE(s) Manually Escalated Successfully")
                ' refresh grid
                SetupTabs()
            Else
                MsgBox("No OCE(s) selected / escalated")
            End If
            OCECancelEscalation = False
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            thisDate = Nothing
            bolLoading = False
        End Try
    End Sub

    Private Sub ProcessRedtagChanges(ByVal thisDate As Date)

        Dim ds As DataSet
        Dim redtagProhib As RedTagProhibitionChangeManager
        Try
            ds = Me.pOCE.GetMyOCERedTagChanges(thisDate, MusterContainer.AppUser.ID)

            If Not ds Is Nothing AndAlso ds.Tables.Count > 1 AndAlso (ds.Tables(0).Rows.Count + ds.Tables(1).Rows.Count) > 0 Then

                redtagProhib = New RedTagProhibitionChangeManager(ds)

                redtagProhib.ShowDialog()

            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally

            If Not ds Is Nothing Then
                ds.Dispose()
            End If

            If Not redtagProhib Is Nothing Then
                redtagProhib.Dispose()
            End If
        End Try



    End Sub

    Private Sub btnEnforceViewEnforceHistory_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEnforceViewEnforceHistory.Click
        Dim bolOwnerSelected As Boolean = False
        Try
            For Each ugRow In ugEnforcement.Rows
                For Each ugChildRow In ugRow.ChildBands(0).Rows
                    If ugChildRow.Cells("SELECTED").Value = True Or _
                        ugChildRow.Cells("SELECTED").Text = True Then
                        bolOwnerSelected = True
                        frmEnforcementHistory = New EnforcementHistory(True, ugChildRow.Cells("OWNER_ID").Value, ugChildRow.Cells("OWNERNAME").Value.ToString, pOCE)
                        frmEnforcementHistory.ShowDialog()
                        ' to prevent from opening enforcement history of selected owners one after another
                        Exit Sub
                    End If
                Next
            Next
            If Not bolOwnerSelected Then
                MsgBox("Please select an Owner to View Enforcement History")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnEnforceRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEnforceRefresh.Click
        SetupTabs()
    End Sub
    Private Sub btnEnforceProcessRescissions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEnforceProcessRescissions.Click
        If bolLoading Then Exit Sub
        Dim bolOCERecinded, bolChangeTOStoTOSI, bolChangeTOSItoTOS, bolCitationRescinded, bolRescindConfirmed As Boolean
        Dim nCitationRescindedCount, nRescindFacilityID As Integer
        Dim slPrevRescindedFacCitCount, slRescindedFacCitCount, slFacCitCount As SortedList
        Dim dtFacs As DataTable
        Dim dr As DataRow
        Dim thisdate As Date = New Date(Now.Year, Now.Month, Now.Day, Now.Hour, Now.Minute, Now.Second)


        dtFacs = New DataTable
        dtFacs.Columns.Add("FACILITY_ID", GetType(Integer))
        dtFacs.Columns.Add("FACILITY", GetType(String))
        dtFacs.Columns.Add("ADDRESS", GetType(String))
        dtFacs.Columns.Add("ADDRESS_STATE_COUNTY", GetType(String))
        dtFacs.Columns.Add("INSPECTEDON", GetType(Date))

        Try
            If Not MusterContainer.AppUser.HasAccess(UIUtilsGen.ModuleID.CAE, MusterContainer.AppUser.UserKey, UIUtilsGen.EntityTypes.CAEOwnerComplianceEvent) Then
                MessageBox.Show("You do not have rights to save Owner Compliance Event.", "ACCESS DENIED", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If

            bolLoading = True
            dtWorkshopDate = CDate("01/01/0001")
            bolOCERecinded = False
            bolRescindConfirmed = False
            Dim skipQuestion As Integer = 0

            For Each ugRow In ugEnforcement.Rows ' status
                For Each ugChildRow In ugRow.ChildBands(0).Rows ' owner
                    bolCitationRescinded = False
                    bolChangeTOStoTOSI = False
                    bolChangeTOSItoTOS = False
                    nRescindFacilityID = 0
                    If ugChildRow.Cells("SELECTED").Value = True Or _
                        ugChildRow.Cells("SELECTED").Text = "True" Then
                        ' rescind oce
                        If Not bolRescindConfirmed Then
                            If MsgBox("Are you sure you want to rescind the selected OCE(s) / Citations?", MsgBoxStyle.YesNo, "Confirm Rescission") = MsgBoxResult.No Then
                                Exit Sub
                            Else
                                bolRescindConfirmed = True
                            End If
                        End If
                        bolOCERecinded = True
                        ' Scenario 5 - 2
                        RescindOCE(ugChildRow)
                        SaveOCE(ugChildRow)
                        SetugRowComboValue(ugChildRow)

                        ' Scenario 5 - 2.i
                        ' If OCE contain Citation "280.20, 280.21 & 280.31" (CITATION_ID = 10) change status of all tanks and pipes that are tos to tosi
                        For Each ugGrandChildRow In ugChildRow.ChildBands(0).Rows
                            RescindOCECitation(ugGrandChildRow)
                            SaveOCECitation(ugGrandChildRow, skipQuestion)
                            If Not bolChangeTOStoTOSI Then
                                If Not ugGrandChildRow.Cells("CITATION_ID").Value Is DBNull.Value Then
                                    If ugGrandChildRow.Cells("CITATION_ID").Value = 10 And ugGrandChildRow.Cells("RESCINDED").OriginalValue = False And ugGrandChildRow.Cells("RESCINDED").Value = True Then
                                        bolChangeTOStoTOSI = True
                                        'ChangeTOStoTOSI(ugGrandChildRow.Cells("OWNER_ID").Value, ugGrandChildRow.Cells("FACILITY_ID").Value)
                                        nRescindFacilityID = ugGrandChildRow.Cells("FACILITY_ID").Value
                                    End If
                                End If
                            End If
                        Next
                        If bolChangeTOStoTOSI Then
                            ChangeTOStoTOSI(ugGrandChildRow.Cells("OWNER_ID").Value, nRescindFacilityID, ugGrandChildRow.Cells("OCE_ID").Value)
                        End If
                    Else
                        If Not ugChildRow.ChildBands(0).Rows Is Nothing Then
                            nCitationRescindedCount = 0
                            slRescindedFacCitCount = New SortedList
                            slPrevRescindedFacCitCount = New SortedList
                            slFacCitCount = New SortedList
                            ' check if citations are selected

                            For Each ugGrandChildRow In ugChildRow.ChildBands(0).Rows ' citation / facility
                                If Not slFacCitCount.ContainsKey(ugGrandChildRow.Cells("FACILITY_ID").Value) Then
                                    slFacCitCount.Add(ugGrandChildRow.Cells("FACILITY_ID").Value, 0)
                                End If
                                slFacCitCount.Item(ugGrandChildRow.Cells("FACILITY_ID").Value) += 1
                                If ugGrandChildRow.Cells("RESCINDED").Value = True Then
                                    If Not slPrevRescindedFacCitCount.ContainsKey(ugGrandChildRow.Cells("FACILITY_ID").Value) Then
                                        slPrevRescindedFacCitCount.Add(ugGrandChildRow.Cells("FACILITY_ID").Value, 0)
                                    End If
                                    slPrevRescindedFacCitCount.Item(ugGrandChildRow.Cells("FACILITY_ID").Value) += 1
                                End If
                                If ((ugGrandChildRow.Cells("SELECTED").Value = True Or _
                                    ugGrandChildRow.Cells("SELECTED").Text = "True") And _
                                    ugGrandChildRow.Cells("RESCINDED").Value = False) Then

                                    If Not bolRescindConfirmed Then
                                        If MsgBox("Are you sure you want to rescind the selected OCE(s) / Citations?", MsgBoxStyle.YesNo, "Confirm Rescission") = MsgBoxResult.No Then
                                            Exit Sub
                                        Else
                                            bolRescindConfirmed = True
                                        End If
                                    End If
                                    bolCitationRescinded = True
                                    bolOCERecinded = True
                                    RescindOCECitation(ugGrandChildRow)
                                    SaveOCECitation(ugGrandChildRow, skipQuestion)
                                    If Not bolChangeTOStoTOSI Then
                                        If Not ugGrandChildRow.Cells("CITATION_ID").Value Is DBNull.Value Then
                                            If ugGrandChildRow.Cells("CITATION_ID").Value = 10 And ugGrandChildRow.Cells("RESCINDED").OriginalValue = False And ugGrandChildRow.Cells("RESCINDED").Value = True Then
                                                bolChangeTOStoTOSI = True
                                                'ChangeTOStoTOSI(ugGrandChildRow.Cells("OWNER_ID").Value, ugGrandChildRow.Cells("FACILITY_ID").Value)
                                                nRescindFacilityID = ugGrandChildRow.Cells("FACILITY_ID").Value
                                            End If
                                        End If
                                    End If

                                    nCitationRescindedCount += 1
                                    If Not slRescindedFacCitCount.ContainsKey(ugGrandChildRow.Cells("FACILITY_ID").Value) Then
                                        slRescindedFacCitCount.Add(ugGrandChildRow.Cells("FACILITY_ID").Value, 0)
                                    End If
                                    slRescindedFacCitCount.Item(ugGrandChildRow.Cells("FACILITY_ID").Value) += 1

                                    If dtFacs.Select("FACILITY_ID = " + ugGrandChildRow.Cells("FACILITY_ID").Value.ToString).Length <= 0 Then
                                        dr = dtFacs.NewRow
                                        dr("FACILITY_ID") = ugGrandChildRow.Cells("FACILITY_ID").Value
                                        dr("FACILITY") = ugGrandChildRow.Cells("FACILITY").Value
                                        dr("ADDRESS") = ugGrandChildRow.Cells("ADDRESS_LINE_ONE").Text + ", " + _
                                                        ugGrandChildRow.Cells("CITY").Text + " " + _
                                                        ugGrandChildRow.Cells("STATE").Text

                                        dr("ADDRESS_STATE_COUNTY") = ugGrandChildRow.Cells("ADDRESS_LINE_ONE").Text + ", " + _
                                                        ugGrandChildRow.Cells("CITY").Text + ", " + _
                                                        IIf(ugGrandChildRow.Cells("STATELONG").Value Is DBNull.Value, ugGrandChildRow.Cells("STATE").Text, ugGrandChildRow.Cells("STATELONG").Text) + ", " + _
                                                        ugGrandChildRow.Cells("COUNTY").Text + " County"

                                        If ugGrandChildRow.Cells("INSPECTEDON").Value Is DBNull.Value Then
                                            dr("INSPECTEDON") = CDate("01/01/0001")
                                        ElseIf ugGrandChildRow.Cells("INSPECTEDON").Text = String.Empty Then
                                            dr("INSPECTEDON") = CDate("01/01/0001")
                                        Else
                                            dr("INSPECTEDON") = ugGrandChildRow.Cells("INSPECTEDON").Value
                                        End If
                                        dtFacs.Rows.Add(dr)
                                    End If

                                End If
                            Next ' For Each ugGrandChildRow In ugChildRow.ChildBands(0).Rows
                            If bolChangeTOStoTOSI Then
                                ChangeTOStoTOSI(ugGrandChildRow.Cells("OWNER_ID").Value, nRescindFacilityID, ugGrandChildRow.Cells("OCE_ID").Value)
                            End If
                            Dim nPrevRescindedCount As Integer = 0
                            For Each nFacID As Integer In slPrevRescindedFacCitCount.Keys
                                nPrevRescindedCount += slPrevRescindedFacCitCount.Item(nFacID)
                            Next
                            If (nPrevRescindedCount + nCitationRescindedCount) = ugChildRow.ChildBands(0).Rows.Count And nCitationRescindedCount > 0 Then
                                If Not bolRescindConfirmed Then
                                    If MsgBox("Are you sure you want to rescind the selected OCE(s) / Citations?", MsgBoxStyle.YesNo, "Confirm Rescission") = MsgBoxResult.No Then
                                        Exit Sub
                                    Else
                                        bolRescindConfirmed = True
                                    End If
                                End If
                                ' Rescind OCE
                                bolOCERecinded = True
                                ' Scenario 6 - 2b
                                RescindOCE(ugChildRow)
                                SaveOCE(ugChildRow)
                                SetugRowComboValue(ugChildRow)
                            Else
                                If bolCitationRescinded Then
                                    ' #2456
                                    Dim bolGenerateNFARescindLetter As Boolean = False
                                    For Each nFacID As Integer In slRescindedFacCitCount.Keys
                                        If slPrevRescindedFacCitCount.Contains(nFacID) Then
                                            nPrevRescindedCount = slPrevRescindedFacCitCount.Item(nFacID)
                                        Else
                                            nPrevRescindedCount = 0
                                        End If
                                        If (nPrevRescindedCount + slRescindedFacCitCount.Item(nFacID)) = slFacCitCount.Item(nFacID) Then
                                            bolGenerateNFARescindLetter = True
                                        Else
                                            If dtFacs.Select("FACILITY_ID = " + nFacID.ToString).Length > 0 Then
                                                dr = dtFacs.Select("FACILITY_ID = " + nFacID.ToString)(0)
                                                dtFacs.Rows.Remove(dr)
                                            End If
                                        End If
                                    Next
                                    If bolGenerateNFARescindLetter Then
                                        ' generate letter
                                        Dim regletter As New Reg_Letters
                                        'Dim documentID As Integer = 0
                                        regletter.GenerateCAENFARescindLetter(ugChildRow, dtFacs, pOCE.GetLetterGeneratedDate(ugChildRow.Cells("OCE_ID").Value, UIUtilsGen.EntityTypes.CAEOwnerComplianceEvent, , False), ugChildRow.ChildBands(0).Rows.Count > 1)
                                    End If

                                    ' Scenario 6 - 2.c
                                    ' Execute OCE Creation / Modification logic
                                    'pOCE.Retrieve(ugChildRow.Cells("OCE_ID").Value)
                                    ' reset oce to new
                                    pOCE = New MUSTER.BusinessLogic.pOwnerComplianceEvent
                                    pOCE.OwnerID = ugChildRow.Cells("OWNER_ID").Value
                                    pOCE.ID = ugChildRow.Cells("OCE_ID").Value

                                    ugChildRow.Cells("SELECTED").Value = True
                                    OCECreationError = False
                                    OCECreationLogic(ugChildRow, False, ugChildRow.Cells("OCE_ID").Value)
                                    If OCECreationError Then
                                        OCECreationError = False
                                        Exit Sub
                                    End If
                                    ' Scenario 2 - 3.b
                                    If pOCE.WorkshopRequired Then
                                        If Date.Compare(dtWorkshopDate, CDate("01/01/0001")) = 0 Then
                                            frmWorkshopDate = New WorkShopDate
                                            frmWorkshopDate.ShowDialog()
                                            pOCE.WorkShopDate = frmWorkshopDate.WorkshopDate
                                            dtWorkshopDate = frmWorkshopDate.WorkshopDate
                                        Else
                                            pOCE.WorkShopDate = dtWorkshopDate
                                        End If
                                    End If
                                    If pOCE.ID <= 0 Then
                                        pOCE.CreatedBy = MusterContainer.AppUser.ID
                                    Else
                                        pOCE.ModifiedBy = MusterContainer.AppUser.ID
                                    End If
                                    pOCE.Save(0, CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal, True)
                                    If Not UIUtilsGen.HasRights(returnVal) Then
                                        Exit Sub
                                    End If

                                    ' If OCE contain Citation "280.20, 280.21 & 280.31" (CITATION_ID = 10) change status of all tanks and pipes that are tosi to tos
                                    If Not bolChangeTOSItoTOS Then
                                        For Each ugGrandChildRow In ugChildRow.ChildBands(0).Rows
                                            If ugGrandChildRow.Cells("SELECTED").Value = False Then
                                                If Not ugGrandChildRow.Cells("CITATION_ID").Value Is DBNull.Value Then
                                                    If ugGrandChildRow.Cells("CITATION_ID").Value = 10 And ugGrandChildRow.Cells("RESCINDED").Value = False Then
                                                        bolChangeTOSItoTOS = True
                                                        ChangeTOSItoTOS(ugRow.Cells("OWNER_ID").Value, ugChildRow.Cells("FACILITY_ID").Value)
                                                        Exit For
                                                    End If
                                                End If
                                            End If
                                        Next
                                    End If ' If Not bolChangeTOSItoTOS Then

                                End If ' If bolCitationRescinded Then

                            End If ' If (nCitationSelectedCount = ugChildRow.ChildBands(0).Rows.Count And nCitationSelectedCount > 0) Or _
                            '(nCitationRescindedCount = ugChildRow.ChildBands(0).Rows.Count And nCitationRescindedCount > 0) Then

                        End If ' If Not ugChildRow.ChildBands(0).Rows Is Nothing Then

                    End If ' If ugChildRow.Cells("SELECTED").Value = True Or _
                    ' ugChildRow.Cells("SELECTED").Text = True Then

                Next ' For Each ugChildRow In ugRow.ChildBands(0).Rows ' owner
            Next ' For Each ugRow In ugEnforcement.Rows ' status

            dtWorkshopDate = CDate("01/01/0001")

            If bolOCERecinded Then

                ProcessRedtagChanges(thisdate)

                ' refresh grid
                MsgBox("Process Rescissions Successful")
                SetupTabs()
            Else
                MsgBox("No OCE(s) / Citations selected to Process Rescissions")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            bolLoading = False
        End Try
    End Sub
    Private Sub btnEnforceGenerateLetter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEnforceGenerateLetter.Click
        If bolLoading Then Exit Sub
        Dim bolLetterGenerated As Boolean
        Dim bolNFARescindLetterGenerated As Boolean = False
        Dim thisDate As Date
        Try

            thisDate = New Date(Now.Year, Now.Month, Now.Day, Now.Hour, Now.Minute, Now.Second)

            If Not MusterContainer.AppUser.HasAccess(UIUtilsGen.ModuleID.CAE, MusterContainer.AppUser.UserKey, UIUtilsGen.EntityTypes.CAEOwnerComplianceEvent) Then
                MessageBox.Show("You do not have rights to save Owner Compliance Event.", "ACCESS DENIED", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            bolLoading = True
            bolLetterGenerated = False
            For Each ugRow In ugEnforcement.Rows ' status
                For Each ugChildRow In ugRow.ChildBands(0).Rows ' owner
                    If ugChildRow.Cells("SELECTED").Value = True Or _
                        ugChildRow.Cells("SELECTED").Text = True Then
                        If Not ugChildRow.Cells("PENDING_LETTER").Value Is DBNull.Value Or Not ugChildRow.Cells("OldLetter").Value Is DBNull.Value Then
                            If (Not ugChildRow.Cells("OldLetter").Value Is DBNull.Value AndAlso ugChildRow.Cells("OldLetter").Value > 0) OrElse Not ugChildRow.Cells("PENDING_LETTER").Value = 0 Then
                                ' Generate Letter
                                lastUGRowOwner = ugChildRow.Cells("OWNER_ID").Value

                                If GenerateOCELetter(ugChildRow) Then
                                    If ugChildRow.Cells("STATUS").Text.ToUpper = "NFA RESCIND" Then
                                        bolNFARescindLetterGenerated = True
                                    End If
                                    If SaveOCE(ugChildRow) Then
                                        ' Get OCE Escalation and update the Grid Row
                                        ' oce escalation is retrieved when oce is saved
                                        ' GetOCEEscalation(ugChildRow)
                                        bolLetterGenerated = True
                                    Else
                                        ReverseGenerateOCELetter(ugChildRow, True)
                                        MsgBox("There was an error Saving OCE for Owner: " + ugChildRow.Cells("OWNERNAME").Value.ToString)
                                        Exit Sub
                                    End If
                                Else
                                    'MsgBox("There was an error Generating: " + ugChildRow.Cells("PENDING LETTER").Text + " for Owner: " + ugChildRow.Cells("OWNERNAME").Value.ToString)
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                Next
            Next

            ProcessRedtagChanges(thisDate)

            If Not bolLetterGenerated Then
                MsgBox("No Letter(s) Generated")

            ElseIf bolNFARescindLetterGenerated Then
                SetupTabs()
            Else

                If Not ugChildRow Is Nothing Then
                    btnEnforceRefresh.PerformClick()
                End If

            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            bolLoading = False
            thisDate = Nothing
        End Try
    End Sub
    Private Sub btnEnforcementExpandCollapseAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEnforcementExpandCollapseAll.Click
        Try
            If btnEnforcementExpandCollapseAll.Text = "Expand All" Then
                ExpandAll(True, ugEnforcement, btnEnforcementExpandCollapseAll)
            Else
                ExpandAll(False, ugEnforcement, btnEnforcementExpandCollapseAll)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugEnforcement_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugEnforcement.InitializeLayout
        Try
            e.Layout.UseFixedHeaders = True
            e.Layout.Override.FixedHeaderIndicator = Infragistics.Win.UltraWinGrid.FixedHeaderIndicator.None
            e.Layout.Bands(0).Columns("SELECTED").Header.Fixed = True
            e.Layout.Bands(0).Columns("STATUS").Header.Fixed = True
            e.Layout.Bands(0).Columns("StatusOrder").Hidden = True

            If e.Layout.Bands(0).Summaries.Count = 0 Then
                e.Layout.Bands(0).Summaries.Add(Infragistics.Win.UltraWinGrid.SummaryType.Sum, e.Layout.Bands(0).Columns("FACILITIES"), SummaryPosition.UseSummaryPositionColumn)
                e.Layout.Bands(0).Summaries.Add(Infragistics.Win.UltraWinGrid.SummaryType.Sum, e.Layout.Bands(0).Columns("OWNERS"), SummaryPosition.UseSummaryPositionColumn)
            End If

            e.Layout.Override.RowSizing = RowSizing.Fixed


            ' cannot have fixed header if row header is split into two rows
            'e.Layout.Bands(1).Columns("SELECTED").Header.Fixed = True
            'e.Layout.Bands(1).Columns("OWNERNAME").Header.Fixed = True
            'e.Layout.Bands(1).Columns("ENSITE ID").Header.Fixed = True
            'e.Layout.Bands(1).Columns("FILLER1").Header.Fixed = True
#If DEBUG Then
            e.Layout.Bands(1).Override.AllowRowFiltering = Infragistics.Win.DefaultableBoolean.True
            e.Layout.Bands(1).Override.RowFilterMode = RowFilterMode.AllRowsInBand
            e.Layout.Bands(1).Override.RowFilterAction = RowFilterAction.HideFilteredOutRows
#End If

            e.Layout.Bands(2).Columns("SELECTED").Header.Fixed = True
            e.Layout.Bands(2).Columns("OWNER_ID").Header.Fixed = True
            e.Layout.Bands(2).Columns("FACILITY_ID").Header.Fixed = True
            e.Layout.Bands(2).Columns("FACILITY").Header.Fixed = True

            'e.Layout.Bands(1).Columns("ENSITE ID").MaskInput = "nnnnnnnnn"
            'e.Layout.Bands(1).Columns("ENSITE ID").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
            'e.Layout.Bands(1).Columns("ENSITE ID").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw

            e.Layout.Bands(1).Columns("POLICY AMOUNT").MaskInput = "$nnn,nnn,nnn.nn"
            e.Layout.Bands(1).Columns("POLICY AMOUNT").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
            e.Layout.Bands(1).Columns("POLICY AMOUNT").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw

            e.Layout.Bands(1).Columns("OVERRIDE AMOUNT").MaskInput = "$nnn,nnn,nnn.nn"
            e.Layout.Bands(1).Columns("OVERRIDE AMOUNT").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
            e.Layout.Bands(1).Columns("OVERRIDE AMOUNT").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw

            e.Layout.Bands(1).Columns("SETTLEMENT AMOUNT").MaskInput = "$nnn,nnn,nnn.nn"
            e.Layout.Bands(1).Columns("SETTLEMENT AMOUNT").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
            e.Layout.Bands(1).Columns("SETTLEMENT AMOUNT").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw

            e.Layout.Bands(1).Columns("PAID AMOUNT").MaskInput = "$nnn,nnn,nnn.nn"
            e.Layout.Bands(1).Columns("PAID AMOUNT").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
            e.Layout.Bands(1).Columns("PAID AMOUNT").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw

            e.Layout.Bands(1).Columns("AGREED ORDER #").MaskInput = "&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            e.Layout.Bands(1).Columns("AGREED ORDER #").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
            e.Layout.Bands(1).Columns("AGREED ORDER #").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw

            e.Layout.Bands(1).Columns("ADMINISTRATIVE ORDER #").MaskInput = "&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            e.Layout.Bands(1).Columns("ADMINISTRATIVE ORDER #").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
            e.Layout.Bands(1).Columns("ADMINISTRATIVE ORDER #").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw

            e.Layout.Bands(0).Override.RowAppearance.BackColor = Color.White
            e.Layout.Bands(1).Override.RowAppearance.BackColor = Color.RosyBrown
            e.Layout.Bands(1).Override.RowAlternateAppearance.BackColor = Color.PeachPuff
            e.Layout.Bands(2).Override.RowAppearance.BackColor = Color.Khaki

            ' setting editable cells color to yellow
            'e.Layout.Bands(1).Columns("PAID AMOUNT").CellAppearance.BackColor = Color.Yellow
            'e.Layout.Bands(1).Columns("DATE RECEIVED").CellAppearance.BackColor = Color.Yellow
            'e.Layout.Bands(1).Columns("ENSITE ID").CellAppearance.BackColor = Color.Yellow
            e.Layout.Bands(1).Columns("WORKSHOP DATE").CellAppearance.BackColor = Color.Yellow
            e.Layout.Bands(1).Columns("WORKSHOP RESULT").CellAppearance.BackColor = Color.Yellow
            e.Layout.Bands(1).Columns("SHOW CAUSE HEARING DATE").CellAppearance.BackColor = Color.Yellow
            e.Layout.Bands(1).Columns("SHOW CAUSE HEARING RESULT").CellAppearance.BackColor = Color.Yellow
            e.Layout.Bands(1).Columns("COMMISSION HEARING RESULT").CellAppearance.BackColor = Color.Yellow
            e.Layout.Bands(1).Columns("COMMISSION HEARING DATE").CellAppearance.BackColor = Color.Yellow
            e.Layout.Bands(1).Columns("COMMENTS").CellAppearance.BackColor = Color.Ivory
            e.Layout.Bands(1).Columns("COMMENTS").CellAppearance.ForeColor = Color.Black
            e.Layout.Bands(1).Columns("COMMENTS").CellAppearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            e.Layout.Bands(1).Columns("COMMENTS").CellDisplayStyle = CellDisplayStyle.PlainText
            e.Layout.Bands(1).Columns("COMMENTS").CellMultiLine = Infragistics.Win.DefaultableBoolean.True
            e.Layout.Bands(1).Columns("COMMENTS").AutoSizeMode = ColumnAutoSizeMode.SiblingRowsOnly
            e.Layout.Bands(1).Columns("COMMENTS").Header.Caption = "COMMENTS"



            'e.Layout.Bands(1).Columns("PAID AMOUNT").CellAppearance.BackColor = Color.Yellow
            'e.Layout.Bands(1).Columns("DATE RECEIVED").CellAppearance.BackColor = Color.Yellow
            e.Layout.Bands(1).Columns("LETTER PRINTED").CellAppearance.BackColor = Color.Yellow
            e.Layout.Bands(1).Columns("OVERRIDE DUE DATE").CellAppearance.BackColor = Color.Yellow
            e.Layout.Bands(1).Columns("OVERRIDE AMOUNT").CellAppearance.BackColor = Color.Yellow
            e.Layout.Bands(2).Columns("DUE").CellAppearance.BackColor = Color.Yellow
            e.Layout.Bands(2).Columns("RECEIVED").CellAppearance.BackColor = Color.Yellow
            e.Layout.Bands(3).Columns("RESCINDED").CellAppearance.BackColor = Color.Yellow
            e.Layout.Bands(3).Columns("RECEIVED").CellAppearance.BackColor = Color.Yellow

            e.Layout.Bands(0).Columns("OCE_STATUS").Hidden = True
            e.Layout.Bands(0).Columns("PROPERTY_POSITION").Hidden = True
            e.Layout.Bands(1).Columns("COMMENTS").Hidden = False
            e.Layout.Bands(0).Columns("STATUSORDER").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending

            e.Layout.Bands(1).Columns("OWNERNAME").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            e.Layout.Bands(2).Columns("FACILITY").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            e.Layout.Bands(3).Columns("DISCREP TEXT").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending

            e.Layout.Bands(1).Columns("OCE_STATUS").Hidden = True
            e.Layout.Bands(1).Columns("OCE_ID").Hidden = True
            e.Layout.Bands(1).Columns("OWNER_ID").Hidden = True
            e.Layout.Bands(1).Columns("OCE_PATH").Hidden = True
            e.Layout.Bands(1).Columns("CITATION").Hidden = True
            e.Layout.Bands(1).Columns("CITATION_DUE_DATE").Hidden = True
            e.Layout.Bands(1).Columns("WORKSHOP_REQUIRED").Hidden = True

            e.Layout.Bands(1).Columns("PENDING_LETTER").Hidden = True

            e.Layout.Bands(1).Columns("CREATED_BY").Hidden = True
            e.Layout.Bands(1).Columns("DATE_CREATED").Hidden = True
            e.Layout.Bands(1).Columns("LAST_EDITED_BY").Hidden = True
            e.Layout.Bands(1).Columns("DATE_LAST_EDITED").Hidden = True
            e.Layout.Bands(1).Columns("DELETED").Hidden = True
            e.Layout.Bands(1).Columns("PENDING_LETTER_TEMPLATE_NUM").Hidden = True
            e.Layout.Bands(1).Columns("ADDRESS_LINE_ONE").Hidden = True
            e.Layout.Bands(1).Columns("ADDRESS_TWO").Hidden = True
            e.Layout.Bands(1).Columns("CITY").Hidden = True
            e.Layout.Bands(1).Columns("STATE").Hidden = True
            e.Layout.Bands(1).Columns("ZIP").Hidden = True
            e.Layout.Bands(1).Columns("ORGANIZATION_ID").Hidden = True
            e.Layout.Bands(1).Columns("PERSON_ID").Hidden = True
            e.Layout.Bands(1).Columns("ESCALATION_ID").Hidden = True

            e.Layout.Bands(2).Columns("INS_CIT_ID").Hidden = True
            e.Layout.Bands(2).Columns("INSPECTION_ID").Hidden = True
            e.Layout.Bands(2).Columns("QUESTION_ID").Hidden = True
            e.Layout.Bands(2).Columns("OCE_ID").Hidden = True
            e.Layout.Bands(2).Columns("FCE_ID").Hidden = True
            e.Layout.Bands(2).Columns("CITATION_ID").Hidden = True
            e.Layout.Bands(2).Columns("NFA_DATE").Hidden = True
            e.Layout.Bands(2).Columns("DELETED").Hidden = True
            e.Layout.Bands(2).Columns("SMALL").Hidden = True
            e.Layout.Bands(2).Columns("MEDIUM").Hidden = True
            e.Layout.Bands(2).Columns("LARGE").Hidden = True
            e.Layout.Bands(2).Columns("CITATION").Hidden = True
            e.Layout.Bands(2).Columns("ADDRESS_LINE_ONE").Hidden = True
            e.Layout.Bands(2).Columns("ADDRESS_TWO").Hidden = True
            e.Layout.Bands(2).Columns("CITY").Hidden = True
            e.Layout.Bands(2).Columns("STATE").Hidden = True
            e.Layout.Bands(2).Columns("ZIP").Hidden = True
            e.Layout.Bands(2).Columns("STATELONG").Hidden = True
            e.Layout.Bands(2).Columns("COUNTY").Hidden = True
            'e.Layout.Bands(1).Columns("FILLER4").Hidden = True
            'e.Layout.Bands(1).Columns("FILLER5").Hidden = True
            e.Layout.Bands(1).Columns("FILLER3").Hidden = True
            e.Layout.Bands(2).Columns("CorrectiveAction").Hidden = True
            e.Layout.Bands(2).Columns("CorrectiveAction").MaxWidth = 4000

            e.Layout.Bands(2).Columns("INSPECTEDON").Hidden = True
            e.Layout.Bands(2).Columns("CITATIONTEXT_FOR_LETTER").Hidden = True
            e.Layout.Bands(2).Columns("CCAT").Hidden = False
            e.Layout.Bands(2).Columns("General Corrective Action").MaxWidth = 4000

            e.Layout.Bands(3).Columns("INS_CIT_ID").Hidden = True
            e.Layout.Bands(3).Columns("INSPECTION_ID").Hidden = True
            e.Layout.Bands(3).Columns("CITATION_ID").Hidden = True
            e.Layout.Bands(3).Columns("QUESTION_ID").Hidden = True
            e.Layout.Bands(1).Columns("COMMISSION HEARING DATE").Hidden = True
            e.Layout.Bands(1).Columns("COMMISSION HEARING RESULT").Hidden = True

            e.Layout.Bands(3).Columns("INS_DESCREP_ID").Hidden = True
            e.Layout.Bands(3).Columns("CorrectiveAction").Hidden = True
            e.Layout.Bands(3).Columns("CorrectiveAction").MaxWidth = 4000

            e.Layout.Bands(3).Columns("CitationText").Hidden = True
            e.Layout.Bands(3).Columns("CCAT").Hidden = False
            e.Layout.Bands(1).Columns("RED TAG DATE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit
            e.Layout.Bands(0).Columns("SELECTED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit
            e.Layout.Bands(0).Columns("STATUS").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            e.Layout.Bands(1).Columns("SELECTED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit
            e.Layout.Bands(1).Columns("OWNERNAME").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("RESCINDED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("OCE DATE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("NEXT DUE DATE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("STATUS").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("POLICY AMOUNT").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("SETTLEMENT AMOUNT").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("PAID AMOUNT").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            'e.Layout.Bands(1).Columns("AGREED ORDER #").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("PENDING LETTER").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("LETTER GENERATED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            e.Layout.Bands(1).Columns("FILLER1").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            'e.Layout.Bands(1).Columns("FILLER2").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("ENSITE ID").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("FILLER3").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("LAST PROCESS DATE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("ESCALATION").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("DATE RECEIVED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            'e.Layout.Bands(1).Columns("FILLER4").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            'e.Layout.Bands(1).Columns("ADMINISTRATIVE ORDER #").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            'e.Layout.Bands(1).Columns("FILLER5").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            e.Layout.Bands(2).Columns("SELECTED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit
            e.Layout.Bands(2).Columns("OWNER_ID").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(2).Columns("FACILITY_ID").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(2).Columns("FACILITY").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(2).Columns("RESCINDED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(2).Columns("SOURCE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(2).Columns("STATECITATION").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(2).Columns("CATEGORY").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(2).Columns("CCAT").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(2).Columns("DUE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit
            e.Layout.Bands(2).Columns("RECEIVED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit
            e.Layout.Bands(2).Columns("CITATIONTEXT").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            e.Layout.Bands(3).Columns("DISCREP TEXT").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(3).Columns("DISCREP TEXT").Width = 400
            e.Layout.Bands(3).Columns("RESCINDED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit
            e.Layout.Bands(3).Columns("RECEIVED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit

            e.Layout.Bands(1).Columns("FILLER1").Header.Caption = "Red Tag Status"
            'e.Layout.Bands(1).Columns("FILLER2").Header.Caption = ""
            e.Layout.Bands(1).Columns("ENSITE ID").Header.Caption = ""
            ' e.Layout.Bands(1).Columns("FILLER3").Header.Caption = ""
            ' e.Layout.Bands(1).Columns("FILLER4").Header.Caption = ""
            'e.Layout.Bands(1).Columns("FILLER5").Header.Caption = ""

            e.Layout.Bands(1).Columns("OCE DATE").Header.Caption = "OCE" + vbCrLf + "DATE"
            e.Layout.Bands(1).Columns("LAST PROCESS DATE").Header.Caption = "LAST" + vbCrLf + "PROCESS DATE"
            e.Layout.Bands(1).Columns("NEXT DUE DATE").Header.Caption = "NEXT" + vbCrLf + "DUE DATE"
            e.Layout.Bands(1).Columns("OVERRIDE DUE DATE").Header.Caption = "OVERRIDE" + vbCrLf + "DUE DATE"
            e.Layout.Bands(1).Columns("POLICY AMOUNT").Header.Caption = "POLICY" + vbCrLf + "AMOUNT"
            e.Layout.Bands(1).Columns("OVERRIDE AMOUNT").Header.Caption = "OVERRIDE" + vbCrLf + "AMOUNT"
            e.Layout.Bands(1).Columns("SETTLEMENT AMOUNT").Header.Caption = "SETTLEMENT" + vbCrLf + "AMOUNT"

            e.Layout.Bands(1).Columns("PAID AMOUNT").Header.Caption = "PAID" + vbCrLf + "AMOUNT"
            e.Layout.Bands(1).Columns("DATE RECEIVED").Header.Caption = "PAYMENT" + vbCrLf + "RECEIVED"
            e.Layout.Bands(1).Columns("WORKSHOP DATE").Header.Caption = "WORKSHOP" + vbCrLf + "DATE"
            e.Layout.Bands(1).Columns("WORKSHOP RESULT").Header.Caption = "WORKSHOP" + vbCrLf + "RESULT"
            e.Layout.Bands(1).Columns("ADMIN HEARING DATE").Header.Caption = "ADMIN HEARING" + vbCrLf + "DATE"
            e.Layout.Bands(1).Columns("ADMIN HEARING RESULT").Header.Caption = "ADMIN HEARING" + vbCrLf + "RESULT"
            e.Layout.Bands(1).Columns("SHOW CAUSE HEARING DATE").Header.Caption = "SHOW CAUSE" + vbCrLf + "HEARING DATE"
            e.Layout.Bands(1).Columns("SHOW CAUSE HEARING RESULT").Header.Caption = "SHOW CAUSE" + vbCrLf + "HEARING RESULT"
            e.Layout.Bands(1).Columns("COMMISSION HEARING RESULT").Header.Caption = "COMMISSION" + vbCrLf + "HEARING RESULT"
            e.Layout.Bands(1).Columns("COMMISSION HEARING DATE").Header.Caption = "COMMISSION" + vbCrLf + "HEARING DATE"
            e.Layout.Bands(1).Columns("AGREED ORDER #").Header.Caption = "AGREED" + vbCrLf + "ORDER #"
            e.Layout.Bands(1).Columns("ADMINISTRATIVE ORDER #").Header.Caption = "ADMINISTRATIVE" + vbCrLf + "ORDER #"
            e.Layout.Bands(1).Columns("PENDING LETTER").Header.Caption = "PENDING" + vbCrLf + "LETTER"
            e.Layout.Bands(1).Columns("LETTER PRINTED").Header.Caption = "LETTER" + vbCrLf + "PRINTED"
            e.Layout.Bands(1).Columns("LETTER GENERATED").Header.Caption = "LETTER" + vbCrLf + "GENERATED"
            e.Layout.Bands(1).Columns("RED TAG DATE").Header.Caption = "RED TAG" + vbCrLf + "DATE"

            e.Layout.Bands(2).Columns("STATECITATION").Header.Caption = "STATE CITATION"
            e.Layout.Bands(2).Columns("CITATIONTEXT").Header.Caption = "CITATION TEXT"

            e.Layout.Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Free
            e.Layout.Bands(1).Override.DefaultColWidth = 100
            e.Layout.Bands(1).ColHeaderLines = 2
            e.Layout.Bands(1).Override.RowSelectorWidth = 1

            'e.Layout.Bands(1).Columns("PENDING LETTER").CellMultiLine = Infragistics.Win.DefaultableBoolean.True
            'e.Layout.Bands(1).Columns("PENDING LETTER").RowLayoutColumnInfo.SpanY = 2
            e.Layout.Bands(1).Override.RowSizing = RowSizing.Fixed
            e.Layout.Bands(1).Columns("COMMENTS").Layout.Override.RowSizing = RowSizing.AutoFixed

            If e.Layout.Bands(1).Groups.Count = 0 Then
                e.Layout.Bands(1).Groups.Add("ROW1")
                'e.layout.Bands(1).Groups.Add("ROW2")
            End If

            e.Layout.Bands(1).GroupHeadersVisible = False

            e.Layout.Bands(1).Columns("SELECTED").Group = e.Layout.Bands(1).Groups("ROW1")
            e.Layout.Bands(1).Columns("FILLER1").Group = e.Layout.Bands(1).Groups("ROW1")
            e.Layout.Bands(1).Columns("OWNERNAME").Group = e.Layout.Bands(1).Groups("ROW1")
            e.Layout.Bands(1).Columns("RESCINDED").Group = e.Layout.Bands(1).Groups("ROW1")

            'e.Layout.Bands(1).Columns("ENSITE ID").Group = e.Layout.Bands(1).Groups("ROW1")
            'e.Layout.Bands(1).Columns("FILLER2").Group = e.Layout.Bands(1).Groups("ROW1")
            'e.Layout.Bands(1).Columns("FILLER3").Group = e.Layout.Bands(1).Groups("ROW1")
            e.Layout.Bands(1).Columns("OCE DATE").Group = e.Layout.Bands(1).Groups("ROW1")
            e.Layout.Bands(1).Columns("LAST PROCESS DATE").Group = e.Layout.Bands(1).Groups("ROW1")
            e.Layout.Bands(1).Columns("NEXT DUE DATE").Group = e.Layout.Bands(1).Groups("ROW1")
            e.Layout.Bands(1).Columns("OVERRIDE DUE DATE").Group = e.Layout.Bands(1).Groups("ROW1")
            e.Layout.Bands(1).Columns("STATUS").Group = e.Layout.Bands(1).Groups("ROW1")
            e.Layout.Bands(1).Columns("ESCALATION").Group = e.Layout.Bands(1).Groups("ROW1")
            e.Layout.Bands(1).Columns("POLICY AMOUNT").Group = e.Layout.Bands(1).Groups("ROW1")
            e.Layout.Bands(1).Columns("PENDING LETTER").Group = e.Layout.Bands(1).Groups("ROW1")
            e.Layout.Bands(1).Columns("SETTLEMENT AMOUNT").Group = e.Layout.Bands(1).Groups("ROW1")
            e.Layout.Bands(1).Columns("OVERRIDE AMOUNT").Group = e.Layout.Bands(1).Groups("ROW1")

            'e.Layout.Bands(1).Columns("FILLER4").Group = e.Layout.Bands(1).Groups("ROW1")

            e.Layout.Bands(1).Columns("PAID AMOUNT").Group = e.Layout.Bands(1).Groups("ROW1")
            e.Layout.Bands(1).Columns("DATE RECEIVED").Group = e.Layout.Bands(1).Groups("ROW1")
            e.Layout.Bands(1).Columns("WORKSHOP DATE").Group = e.Layout.Bands(1).Groups("ROW1")
            e.Layout.Bands(1).Columns("WORKSHOP RESULT").Group = e.Layout.Bands(1).Groups("ROW1")
            e.Layout.Bands(1).Columns("ADMIN HEARING DATE").Group = e.Layout.Bands(1).Groups("ROW1")
            e.Layout.Bands(1).Columns("ADMIN HEARING RESULT").Group = e.Layout.Bands(1).Groups("ROW1")
            e.Layout.Bands(1).Columns("SHOW CAUSE HEARING DATE").Group = e.Layout.Bands(1).Groups("ROW1")
            e.Layout.Bands(1).Columns("SHOW CAUSE HEARING RESULT").Group = e.Layout.Bands(1).Groups("ROW1")
            'e.Layout.Bands(1).Columns("COMMISSION HEARING DATE").Group = e.Layout.Bands(1).Groups("ROW1")
            'e.Layout.Bands(1).Columns("COMMISSION HEARING RESULT").Group = e.Layout.Bands(1).Groups("ROW1")

            e.Layout.Bands(1).Columns("AGREED ORDER #").Group = e.Layout.Bands(1).Groups("ROW1")
            e.Layout.Bands(1).Columns("ADMINISTRATIVE ORDER #").Group = e.Layout.Bands(1).Groups("ROW1")
            ' e.Layout.Bands(1).Columns("FILLER5").Group = e.Layout.Bands(1).Groups("ROW1")
            e.Layout.Bands(1).Columns("LETTER GENERATED").Group = e.Layout.Bands(1).Groups("ROW1")
            e.Layout.Bands(1).Columns("RED TAG DATE").Group = e.Layout.Bands(1).Groups("ROW1")
            e.Layout.Bands(1).Columns("LETTER PRINTED").Group = e.Layout.Bands(1).Groups("ROW1")
            e.Layout.Bands(1).Columns("COMMENTS").Group = e.Layout.Bands(1).Groups("ROW1")

            e.Layout.Bands(1).Columns("LETTER GENERATED").Width = 100
            e.Layout.Bands(1).Columns("RED TAG DATE").Width = 100




            e.Layout.Bands(1).LevelCount = 3
            e.Layout.Bands(1).Columns("SELECTED").Level = 0
            e.Layout.Bands(1).Columns("FILLER1").Level = 1
            e.Layout.Bands(1).Columns("OWNERNAME").Level = 0
            e.Layout.Bands(1).Columns("ENSITE ID").Level = 1
            'e.Layout.Bands(1).Columns("FILLER2").Level = 1
            e.Layout.Bands(1).Columns("RESCINDED").Level = 1
            e.Layout.Bands(1).Columns("FILLER3").Level = 1
            e.Layout.Bands(1).Columns("OCE DATE").Level = 0
            e.Layout.Bands(1).Columns("LAST PROCESS DATE").Level = 1
            e.Layout.Bands(1).Columns("NEXT DUE DATE").Level = 0
            e.Layout.Bands(1).Columns("OVERRIDE DUE DATE").Level = 1
            e.Layout.Bands(1).Columns("STATUS").Level = 0
            e.Layout.Bands(1).Columns("ESCALATION").Level = 1
            e.Layout.Bands(1).Columns("POLICY AMOUNT").Level = 0
            e.Layout.Bands(1).Columns("PENDING LETTER").Level = 1
            e.Layout.Bands(1).Columns("SETTLEMENT AMOUNT").Level = 0
            e.Layout.Bands(1).Columns("OVERRIDE AMOUNT").Level = 1

            'e.Layout.Bands(1).Columns("FILLER4").Level = 1
            e.Layout.Bands(1).Columns("PAID AMOUNT").Level = 0
            e.Layout.Bands(1).Columns("DATE RECEIVED").Level = 1
            e.Layout.Bands(1).Columns("WORKSHOP DATE").Level = 0
            e.Layout.Bands(1).Columns("WORKSHOP RESULT").Level = 1
            e.Layout.Bands(1).Columns("ADMIN HEARING DATE").Level = 0
            e.Layout.Bands(1).Columns("ADMIN HEARING RESULT").Level = 1
            e.Layout.Bands(1).Columns("SHOW CAUSE HEARING DATE").Level = 0
            e.Layout.Bands(1).Columns("SHOW CAUSE HEARING RESULT").Level = 1
            'e.Layout.Bands(1).Columns("COMMISSION HEARING RESULT").Level = 1
            'e.Layout.Bands(1).Columns("COMMISSION HEARING DATE").Level = 0

            e.Layout.Bands(1).Columns("AGREED ORDER #").Level = 0
            e.Layout.Bands(1).Columns("ADMINISTRATIVE ORDER #").Level = 1
            ' e.Layout.Bands(1).Columns("FILLER5").Level = 1
            e.Layout.Bands(1).Columns("LETTER GENERATED").Level = 0
            e.Layout.Bands(1).Columns("RED TAG DATE").Level = 0
            e.Layout.Bands(1).Columns("LETTER PRINTED").Level = 1
            e.Layout.Bands(1).Columns("COMMENTS").Level = 2
            e.Layout.Bands(1).Columns("COMMENTS").CellDisplayStyle = CellDisplayStyle.FullEditorDisplay
            e.Layout.Bands(1).Columns("COMMENTS").VertScrollBar = True

            e.Layout.Bands(1).Columns("WORKSHOP RESULT").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
            e.Layout.Bands(1).Columns("ADMIN HEARING RESULT").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
            e.Layout.Bands(1).Columns("SHOW CAUSE HEARING RESULT").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
            e.Layout.Bands(1).Columns("COMMISSION HEARING RESULT").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
            e.Layout.Override.TipStyleCell = TipStyle.Hide

            ' populate the whole column as the table is the same for each row
            ' WorkShopResult
            If e.Layout.Bands(1).Columns("WORKSHOP RESULT").ValueList Is Nothing Then
                vListWorkShopResult = New Infragistics.Win.ValueList
                For Each row As DataRow In pOCE.GetWorkshopResults.Tables(0).Rows
                    vListWorkShopResult.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                Next
                e.Layout.Bands(1).Columns("WORKSHOP RESULT").ValueList = vListWorkShopResult
            End If
            ' SHOW CAUSE HEARING RESULT
            If e.Layout.Bands(1).Columns("SHOW CAUSE HEARING RESULT").ValueList Is Nothing Then
                vListShowCauseResult = New Infragistics.Win.ValueList
                For Each row As DataRow In pOCE.GetShowCauseHearingResults.Tables(0).Rows
                    vListShowCauseResult.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                Next
                e.Layout.Bands(1).Columns("SHOW CAUSE HEARING RESULT").ValueList = vListShowCauseResult
            End If
            ' COMMISSION HEARING RESULT
            If e.Layout.Bands(1).Columns("COMMISSION HEARING RESULT").ValueList Is Nothing Then
                vListCommissionResult = New Infragistics.Win.ValueList
                For Each row As DataRow In pOCE.GetCommissionHearingResults.Tables(0).Rows
                    vListCommissionResult.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                Next
                e.Layout.Bands(1).Columns("COMMISSION HEARING RESULT").ValueList = vListCommissionResult
            End If

            ' ADMIN HEARING RESULT
            If e.Layout.Bands(1).Columns("ADMIN HEARING RESULT").ValueList Is Nothing Then
                vListAdminHearingResult = New Infragistics.Win.ValueList
                For Each row As DataRow In pOCE.GetAdminHearingResults.Tables(0).Rows
                    vListAdminHearingResult.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                Next
                e.Layout.Bands(1).Columns("ADMIN HEARING RESULT").ValueList = vListAdminHearingResult
            End If

            ' Will be handled when owner row is expanded
            'For Each ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow In e.Layout.Grid.Rows ' status
            '    If Not ugRow.ChildBands Is Nothing Then
            '        For Each ugChildRow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugRow.ChildBands(0).Rows ' owner
            '            SetugRowComboValue(ugChildRow)
            '            If Not ugChildRow.Cells("POLICY AMOUNT").Value Is DBNull.Value Then
            '                If ugChildRow.Cells("POLICY AMOUNT").Value < 0 Then
            '                    ugChildRow.Cells("POLICY AMOUNT").Value = DBNull.Value
            '                End If
            '            End If
            '            If Not ugChildRow.Cells("SETTLEMENT AMOUNT").Value Is DBNull.Value Then
            '                If ugChildRow.Cells("SETTLEMENT AMOUNT").Value < 0 Then
            '                    ugChildRow.Cells("SETTLEMENT AMOUNT").Value = DBNull.Value
            '                End If
            '            End If
            '            If Not ugChildRow.Cells("OVERRIDE AMOUNT").Value Is DBNull.Value Then
            '                If ugChildRow.Cells("OVERRIDE AMOUNT").Value < 0 Then
            '                    ugChildRow.Cells("OVERRIDE AMOUNT").Value = DBNull.Value
            '                End If
            '            End If
            '            If Not ugChildRow.Cells("PAID AMOUNT").Value Is DBNull.Value Then
            '                If ugChildRow.Cells("PAID AMOUNT").Value < 0 Then
            '                    ugChildRow.Cells("PAID AMOUNT").Value = DBNull.Value
            '                End If
            '            End If
            '            If Not ugChildRow.ChildBands Is Nothing Then
            '                For Each ugGrandChildRow In ugChildRow.ChildBands(0).Rows ' citation
            '                    If Not ugGrandChildRow.ChildBands Is Nothing Then
            '                        For Each ugGreatGrandChildRow In ugGrandChildRow.ChildBands(0).Rows
            '                            If ugGreatGrandChildRow.Cells("RESCINDED").Value = True Then
            '                                ugGreatGrandChildRow.Cells("RESCINDED").Activation = Activation.NoEdit
            '                                ugGreatGrandChildRow.Cells("RESCINDED").Appearance.BackColor = Color.White
            '                            End If
            '                        Next
            '                    End If
            '                Next
            '            End If
            '        Next
            '    End If
            'Next
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub CreateToolTipMessage(ByVal cell As UltraGridCell)

        If EnforcementToolTip Is Nothing Then
            EnforcementToolTip = New ToolTip
        End If

        EnforcementToolTip.SetToolTip(ugEnforcement, cell.Text)

    End Sub

    Private Sub ugEnforcement_MouseLeaveElement(ByVal sender As Object, ByVal e As Infragistics.Win.UIElementEventArgs) Handles ugEnforcement.MouseLeaveElement
        ' if we are not leaving a cell, then don't anything
        If Not e.Element.GetType().Equals(GetType(CellUIElement)) Then
            Exit Sub
        End If


        ' destroy the tooltip
        If Not EnforcementToolTip Is Nothing Then
            EnforcementToolTip.SetToolTip(Me, String.Empty)
            EnforcementToolTip.Dispose()
            EnforcementToolTip = Nothing
        End If
    End Sub

    Private Sub ugEnforcement_MouseEnterElement(ByVal sender As Object, ByVal e As Infragistics.Win.UIElementEventArgs) Handles ugEnforcement.MouseEnterElement

        ' if we are not entering a cell, then don't anything
        If Not e.Element.GetType().Equals(GetType(CellUIElement)) Then
            Exit Sub
        End If

        ' find the cell that the cursor is over, if any
        Dim cell As UltraGridCell = e.Element.GetContext(GetType(UltraGridCell))
        If Not cell Is Nothing Then
            CreateToolTipMessage(cell)
        End If

    End Sub


    Private Sub ugEnforcement_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugEnforcement.CellChange
        If bolLoading Then Exit Sub
        Try
            bolLoading = True
            If e.Cell.Row.Band.Index = 0 Then ' status
                If "SELECTED".Equals(e.Cell.Column.Key) Then
                    For Each ugChildRow In e.Cell.Row.ChildBands(0).Rows
                        ugChildRow.Cells("SELECTED").Value = e.Cell.Text
                        For Each ugGrandChildRow In ugChildRow.ChildBands(0).Rows
                            ugGrandChildRow.Cells("SELECTED").Value = e.Cell.Text
                        Next
                    Next
                End If
                e.Cell.Value = e.Cell.Text
            ElseIf e.Cell.Row.Band.Index = 1 Then ' owner (OCE)
                If "SELECTED".Equals(e.Cell.Column.Key) Then
                    For Each ugGrandChildRow In e.Cell.Row.ChildBands(0).Rows
                        ugGrandChildRow.Cells("SELECTED").Value = e.Cell.Text
                    Next
                    e.Cell.Value = e.Cell.Text
                ElseIf "ENSITE ID".Equals(e.Cell.Column.Key) Then
                    bolOCEModified = True
                    e.Cell.Value = e.Cell.EditorResolved.Value
                ElseIf "PAID AMOUNT".Equals(e.Cell.Column.Key) Then
                    bolOCEModified = True
                    e.Cell.Value = e.Cell.EditorResolved.Value
                ElseIf "WORKSHOP DATE".Equals(e.Cell.Column.Key) Then
                    bolOCEModified = True
                    If e.Cell.Text = "__/__/____" Then
                        e.Cell.Value = DBNull.Value
                    Else
                        e.Cell.Value = e.Cell.Text
                    End If
                ElseIf "SHOW CAUSE HEARING DATE".Equals(e.Cell.Column.Key) Then
                    bolOCEModified = True
                    If e.Cell.Text = "__/__/____" Then
                        e.Cell.Value = DBNull.Value
                    Else
                        e.Cell.Value = e.Cell.Text
                    End If
                ElseIf "ADMIN HEARING DATE".Equals(e.Cell.Column.Key) Then
                    bolOCEModified = True
                    If e.Cell.Text = "__/__/____" Then
                        e.Cell.Value = DBNull.Value
                    Else
                        e.Cell.Value = e.Cell.Text
                    End If
                ElseIf "COMMISSION HEARING RESULT".Equals(e.Cell.Column.Key) Then
                    bolOCEModified = True
                    e.Cell.Value = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
                ElseIf "ADMIN HEARING RESULT".Equals(e.Cell.Column.Key) Then
                    bolOCEModified = True
                    e.Cell.Value = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)

                ElseIf "OVERRIDE DUE DATE".Equals(e.Cell.Column.Key) Then
                    bolOCEModified = True
                    bolOCEDueDateModified = True
                    If e.Cell.Text = "__/__/____" Then
                        e.Cell.Value = DBNull.Value
                    Else
                        e.Cell.Value = e.Cell.Text
                    End If
                ElseIf "OVERRIDE AMOUNT".Equals(e.Cell.Column.Key) Then
                    bolOCEModified = True
                    e.Cell.Value = e.Cell.EditorResolved.Value
                ElseIf "DATE RECEIVED".Equals(e.Cell.Column.Key) Then
                    bolOCEModified = True
                    If e.Cell.Text = "__/__/____" Then
                        e.Cell.Value = DBNull.Value
                    Else
                        e.Cell.Value = e.Cell.Text
                    End If
                ElseIf "WORKSHOP RESULT".Equals(e.Cell.Column.Key) Then
                    bolOCEModified = True
                    e.Cell.Value = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
                ElseIf "SHOW CAUSE HEARING RESULT".Equals(e.Cell.Column.Key) Then
                    bolOCEModified = True
                    e.Cell.Value = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
                ElseIf "COMMENTS".Equals(e.Cell.Column.Key) Then
                    bolOCEModified = True
                    e.Cell.Value = e.Cell.Text

                ElseIf "COMMISSION HEARING DATE".Equals(e.Cell.Column.Key) Then
                    bolOCEModified = True
                    If e.Cell.Text = "__/__/____" Then
                        e.Cell.Value = DBNull.Value
                    Else
                        e.Cell.Value = e.Cell.Text
                    End If
                ElseIf "LETTER PRINTED".Equals(e.Cell.Column.Key) Then
                    If e.Cell.Row.Cells("LETTER GENERATED").Value Is DBNull.Value Then
                        MsgBox("Cannot mark Letter as Printed before Generating the Letter")
                        e.Cell.CancelUpdate()
                    Else
                        bolOCEModified = True
                        e.Cell.Value = e.Cell.Text
                    End If
                ElseIf "Red Tag Date".Equals(e.Cell.Column.Key) Then
                    bolOCEModified = True
                    If e.Cell.Text = "__/__/____" Then
                        e.Cell.Value = DBNull.Value
                    Else
                        e.Cell.Value = e.Cell.Text
                    End If
                ElseIf "AGREED ORDER #".Equals(e.Cell.Column.Key) Then
                    bolOCEModified = True
                    e.Cell.Value = e.Cell.EditorResolved.Value
                ElseIf "ADMINISTRATIVE ORDER #".Equals(e.Cell.Column.Key) Then
                    bolOCEModified = True
                    e.Cell.Value = e.Cell.EditorResolved.Value
                End If
            ElseIf e.Cell.Row.Band.Index = 2 Then ' citation
                If "SELECTED".Equals(e.Cell.Column.Key) Then
                    Dim citationCnt As Integer = e.Cell.Row.ParentRow.ChildBands(0).Rows.Count
                    Dim selectedCount As Integer = 0
                    For Each ugGrandChildRow In e.Cell.Row.ParentRow.ChildBands(0).Rows
                        If (ugGrandChildRow.Cells("SELECTED").Value = True Or _
                                ugGrandChildRow.Cells("SELECTED").Text.ToUpper = "TRUE") And _
                                ugGrandChildRow.Cells("SELECTED").Text.ToUpper <> "FALSE" Then
                            selectedCount += 1
                        Else
                            Exit For
                        End If
                    Next
                    If selectedCount = citationCnt Then
                        e.Cell.Row.ParentRow.Cells("SELECTED").Value = True
                    Else
                        e.Cell.Row.ParentRow.Cells("SELECTED").Value = False
                    End If
                    e.Cell.Value = e.Cell.Text
                ElseIf "DUE".Equals(e.Cell.Column.Key) Then
                    If e.Cell.Text = "__/__/____" Then
                        e.Cell.Value = DBNull.Value
                    Else
                        e.Cell.Value = e.Cell.Text
                    End If
                    bolCitationModified = True
                ElseIf "RECEIVED".Equals(e.Cell.Column.Key) Then
                    If e.Cell.Text = "__/__/____" Then
                        e.Cell.Value = DBNull.Value
                    Else
                        e.Cell.Value = e.Cell.Text
                    End If
                    bolCitationModified = True
                    ' if this is the only row, no need to update other rows
                    If e.Cell.Row.ParentRow.ChildBands(0).Rows.Count > 1 Then
                        bolCitationDateModified = True
                    End If
                End If
            ElseIf e.Cell.Row.Band.Index = 3 Then ' discrep
                If "RESCINDED".Equals(e.Cell.Column.Key) Then
                    If e.Cell.Text = False Then
                        e.Cell.CancelUpdate()
                    Else
                        e.Cell.Value = e.Cell.Text
                        e.Cell.Activation = Activation.NoEdit
                        e.Cell.Appearance.BackColor = Color.White
                        bolDiscrepModified = True
                    End If
                ElseIf "RECEIVED".Equals(e.Cell.Column.Key) Then
                    If e.Cell.Text = "__/__/____" Then
                        e.Cell.Value = DBNull.Value
                    Else
                        e.Cell.Value = e.Cell.Text
                    End If
                    bolDiscrepModified = True
                    ' if this is the only row, no need to update other rows
                    If e.Cell.Row.ParentRow.ChildBands(0).Rows.Count > 1 Then
                        bolDiscrepDateModified = True
                    End If
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            bolLoading = False
        End Try
    End Sub
    Private Sub ugEnforcement_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles ugEnforcement.BeforeRowUpdate
        If bolLoading Then Exit Sub
        Try

            Dim skipQuestion As Integer = 0

            If Not ugEnforcement.ActiveRow Is Nothing Then
                If ugEnforcement.ActiveRow.Band.Index = 1 And bolOCEModified Then
                    bolOCEModified = True

                    Dim pendingletter As Integer = 0



                    If Not SaveOCE(ugEnforcement.ActiveRow) Then
                        MsgBox("There was an error saving OCE - " + ugEnforcement.ActiveRow.Cells("OWNERNAME").Text)
                        bolOCEModified = False
                        e.Cancel = True
                        Exit Sub

                    End If
                    bolOCEModified = False

                ElseIf ugEnforcement.ActiveRow.Band.Index = 2 And bolCitationModified Then
                    bolCitationModified = False
                    If Not SaveOCECitation(ugEnforcement.ActiveRow, skipQuestion) Then
                        MsgBox("There was an error saving OCE Citation - " + ugEnforcement.ActiveRow.ParentRow.Cells("OWNERNAME").Text + "-" + ugEnforcement.ActiveRow.Cells("FACILITY").Text)
                        e.Cancel = True
                        Exit Sub
                    ElseIf bolCitationDateModified Then
                        If e.Row.Cells("RECEIVED").Value Is DBNull.Value Then
                            ModifyOCECitationReceivedDate(e.Row.ParentRow, CDate("01/01/0001"))
                        Else
                            ModifyOCECitationReceivedDate(e.Row.ParentRow, e.Row.Cells("RECEIVED").Value)
                        End If
                    End If
                ElseIf ugEnforcement.ActiveRow.Band.Index = 3 And bolDiscrepModified Then
                    bolDiscrepModified = False
                    If Not SaveOCEDiscrep(ugEnforcement.ActiveRow, skipQuestion) Then

                        MsgBox("There was an error saving OCE Citation Discrep - " + ugEnforcement.ActiveRow.Cells("DISCREP TEXT").Text + "-" + ugEnforcement.ActiveRow.ParentRow.ParentRow.Cells("OWNERNAME").Text + "-" + ugEnforcement.ActiveRow.ParentRow.Cells("FACILITY").Text)
                        e.Cancel = True
                        Exit Sub
                    ElseIf bolDiscrepDateModified Then
                        ModifyOCEDiscrepReceivedDate(e.Row.ParentRow, e.Row.Cells("RECEIVED").Value)
                    End If
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugEnforcement_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugEnforcement.Leave
        If bolLoading Then Exit Sub
        Try
            If Not ugEnforcement.ActiveRow Is Nothing Then
                Dim ea As New Infragistics.Win.UltraWinGrid.CancelableRowEventArgs(ugEnforcement.ActiveRow)
                ugEnforcement_BeforeRowUpdate(sender, ea)
            End If
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugEnforcement_BeforeRowExpanded(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles ugEnforcement.BeforeRowExpanded
        If bolLoading Then Exit Sub
        Dim bolContinue As Boolean = False
        Try
            If e.Row.Band.Index = 0 Then ' status
                'If e.Row.Tag Is Nothing Then
                '    ' add oce's to status
                '    Try
                '        Me.Cursor = Cursors.WaitCursor
                '        bolLoading = True
                '        Dim ds As DataSet = pOCE.GetEnforcements(, False, True, e.Row.Cells("OCE_STATUS").Value)

                '        If e.Row.Cells("OCE_STATUS").Value <> ugEnforcement.ActiveRow.Cells("OCE_STATUS").Value Then
                '            ugEnforcement.ActiveRow = e.Row
                '        End If

                '        For Each dr As DataRow In ds.Tables(1).Rows
                '            ugEnforcement.DisplayLayout.Bands(1).AddNew()
                '            For Each col As DataColumn In ds.Tables(1).Columns
                '                ugEnforcement.ActiveRow.Cells(col.ColumnName).Value = dr(col.ColumnName)
                '            Next
                '            ugEnforcement.ActiveRow.Expanded = True
                '            ugEnforcement.ActiveRow.Expanded = False
                '            For Each drCit As DataRow In ds.Tables(2).Select("OCE_ID = " + dr("OCE_ID").ToString)
                '                ugEnforcement.DisplayLayout.Bands(2).AddNew()
                '                ugEnforcement.ActiveRow.ParentRow.Expanded = False
                '                For Each colCit As DataColumn In ds.Tables(2).Columns
                '                    ugEnforcement.ActiveRow.Cells(colCit.ColumnName).Value = drCit(colCit.ColumnName)
                '                Next
                '                For Each drDisc As DataRow In ds.Tables(3).Select("INS_CIT_ID = " + drCit("INS_CIT_ID").ToString)
                '                    ugEnforcement.DisplayLayout.Bands(3).AddNew()
                '                    ugEnforcement.ActiveRow.ParentRow.ParentRow.Expanded = False
                '                    For Each colDisc As DataColumn In ds.Tables(3).Columns
                '                        ugEnforcement.ActiveRow.Cells(colDisc.ColumnName).Value = drDisc(colDisc.ColumnName)
                '                    Next
                '                Next
                '            Next
                '        Next
                '        ugEnforcement.ActiveRow = e.Row
                '        e.Row.Tag = "0"
                '    Catch ex As Exception
                '        Throw ex
                '    Finally
                '        bolLoading = False
                '    End Try
                'End If
                If Not e.Row.ChildBands Is Nothing Then
                    If e.Row.Tag Is Nothing Then
                        bolContinue = True
                    ElseIf e.Row.Tag = "0" Then
                        bolContinue = True
                    End If
                    If bolContinue Then
                        Me.Cursor = Cursors.WaitCursor
                        e.Row.Tag = "1"
                        For Each ugChildRow As Infragistics.Win.UltraWinGrid.UltraGridRow In e.Row.ChildBands(0).Rows ' owner
                            SetugRowComboValue(ugChildRow)
                            If Not ugChildRow.Cells("POLICY AMOUNT").Value Is DBNull.Value Then
                                If ugChildRow.Cells("POLICY AMOUNT").Value < 0 Then
                                    ugChildRow.Cells("POLICY AMOUNT").Value = DBNull.Value
                                End If
                            End If
                            If Not ugChildRow.Cells("SETTLEMENT AMOUNT").Value Is DBNull.Value Then
                                If ugChildRow.Cells("SETTLEMENT AMOUNT").Value < 0 Then
                                    ugChildRow.Cells("SETTLEMENT AMOUNT").Value = DBNull.Value
                                End If
                            End If
                            If Not ugChildRow.Cells("OVERRIDE AMOUNT").Value Is DBNull.Value Then
                                If ugChildRow.Cells("OVERRIDE AMOUNT").Value < 0 Then
                                    ugChildRow.Cells("OVERRIDE AMOUNT").Value = DBNull.Value
                                End If
                            End If
                            If Not ugChildRow.Cells("PAID AMOUNT").Value Is DBNull.Value Then
                                If ugChildRow.Cells("PAID AMOUNT").Value < 0 Then
                                    ugChildRow.Cells("PAID AMOUNT").Value = DBNull.Value
                                End If
                            End If
                            If Not ugChildRow.ChildBands Is Nothing Then
                                For Each ugGrandChildRow In ugChildRow.ChildBands(0).Rows ' citation
                                    If Not ugGrandChildRow.ChildBands Is Nothing Then
                                        For Each ugGreatGrandChildRow In ugGrandChildRow.ChildBands(0).Rows
                                            If ugGreatGrandChildRow.Cells("RESCINDED").Value = True Then
                                                ugGreatGrandChildRow.Cells("RESCINDED").Activation = Activation.NoEdit
                                                ugGreatGrandChildRow.Cells("RESCINDED").Appearance.BackColor = Color.White
                                            End If
                                        Next
                                    End If
                                Next
                            End If
                        Next
                    End If
                End If
            End If
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            If Me.Cursor.Equals(Cursors.WaitCursor) Then
                Me.Cursor = Cursors.Default
            End If
        End Try
    End Sub
    'Private Sub ugEnforcement_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ugEnforcement.KeyDown
    '    If bolLoading Then Exit Sub
    '    Try
    '        bolLoading = True
    '        If e.KeyCode = Keys.Delete Then
    '            If Not ugEnforcement.ActiveCell Is Nothing Then

    '                If ugEnforcement.ActiveCell.Row.Band.Index = 1 Then ' owner (OCE)
    '                    If "COMMISSION HEARING RESULT".Equals(ugEnforcement.ActiveCell.Column.Key) Then
    '                        ugEnforcement.ActiveCell.Value = 0
    '                        SetugRowComboValue(ugEnforcement.ActiveCell.Row)
    '                        bolOCEModified = True
    '                        btnEnforceRefresh.Focus()
    '                        ugEnforcement.Focus()
    '                    ElseIf "WORKSHOP RESULT".Equals(ugEnforcement.ActiveCell.Column.Key) Then
    '                        ugEnforcement.ActiveCell.Value = 0
    '                        SetugRowComboValue(ugEnforcement.ActiveCell.Row)
    '                        bolOCEModified = True
    '                        btnEnforceRefresh.Focus()
    '                        ugEnforcement.Focus()
    '                    ElseIf "SHOW CAUSE HEARING RESULT".Equals(ugEnforcement.ActiveCell.Column.Key) Then
    '                        ugEnforcement.ActiveCell.Value = 0
    '                        SetugRowComboValue(ugEnforcement.ActiveCell.Row)
    '                        bolOCEModified = True
    '                        btnEnforceRefresh.Focus()
    '                        ugEnforcement.Focus()
    '                    End If
    '                End If

    '            End If
    '        End If
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    Finally
    '        bolLoading = False
    '    End Try
    'End Sub
    Private Sub ugEnforcement_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ugEnforcement.MouseDown
        If Not e.Button = MouseButtons.Right Then Exit Sub

        Dim point As New System.Drawing.Point(e.X, e.Y)

        '   Get the UIElement
        Dim objUIElement As Infragistics.Win.UIElement
        objUIElement = ugEnforcement.DisplayLayout.UIElement.ElementFromPoint(point)
        If objUIElement Is Nothing Then Exit Sub

        '   See if we are over a cell
        Dim objCell As Infragistics.Win.UltraWinGrid.UltraGridCell
        objCell = objUIElement.GetContext(GetType(Infragistics.Win.UltraWinGrid.UltraGridCell))
        If Not objCell Is Nothing Then
            If objCell.Column.Key = "OVERRIDE AMOUNT" Then
                If objCell.Row.Cells("OCE_ID").Value > 0 Then
                    Dim tt As New Infragistics.Win.ToolTip(Me.ugEnforcement)
                    tt.AutoPopDelay = 15000 ' set lenth tip is visible to 15 seconds
                    tt.ToolTipText = pOCE.GetCAEOverrideAmountHistory(objCell.Row.Cells("OCE_ID").Value)
                    tt.SetMargin(1, 1, 1, 1)
                    tt.Show()
                    'MsgBox(pOCE.GetCAEOverrideAmountHistory(objCell.Row.Cells("OCE_ID").Value), MsgBoxStyle.OKOnly, objCell.Row.Cells("OWNERNAME").Text + " Override Amount History")
                End If
            End If
        End If
    End Sub
#End Region

#Region "Assigned Inspection"
    Private Sub btnAssignedInspectionsAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAssignedInspectionsAdd.Click
        Try
            If Not MusterContainer.AppUser.HasAccess(UIUtilsGen.ModuleID.CAE, MusterContainer.AppUser.UserKey, UIUtilsGen.EntityTypes.Inspection) Then
                MessageBox.Show("You do not have rights to save Inspection.", "ACCESS DENIED", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            Dim frmAssignedInspection As New CandEAssignedInspection(0)
            frmAssignedInspection.CallingForm = Me
            Me.Tag = "0"
            frmAssignedInspection.ShowDialog()
            If Me.Tag = "1" Then
                SetupTabs()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnEditAssignedInspection_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEditAssignedInspection.Click
        Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            If Not MusterContainer.AppUser.HasAccess(UIUtilsGen.ModuleID.CAE, MusterContainer.AppUser.UserKey, UIUtilsGen.EntityTypes.Inspection) Then
                MessageBox.Show("You do not have rights to save Inspection.", "ACCESS DENIED", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            ugRow = getSelectedRow()
            If ugRow Is Nothing Then
                MsgBox("Please select an Assigned Inspection to Edit")
                Exit Sub
            End If
            If ugRow.Cells("SUBMITTED").Value Is DBNull.Value And _
                Not ugRow.Cells("COMPLETED").Value Is DBNull.Value Then
                MsgBox("Cannot edit Assigned Inspection - Completed but not submitted to C & E")
                Exit Sub
            Else
                Dim frmAssignedInspection As New CandEAssignedInspection(ugRow.Cells("INSPECTION_ID").Value)
                frmAssignedInspection.CallingForm = Me
                Me.Tag = "0"
                frmAssignedInspection.ShowDialog()
                If Me.Tag = "1" Then
                    'refresh grid
                    SetupTabs()
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnAssignedInspectionsAccept_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAssignedInspectionsAccept.Click
        Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            If Not MusterContainer.AppUser.HasAccess(UIUtilsGen.ModuleID.CAE, MusterContainer.AppUser.UserKey, UIUtilsGen.EntityTypes.Inspection) Then
                MessageBox.Show("You do not have rights to save Inspection.", "ACCESS DENIED", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            ugRow = getSelectedRow()
            If ugRow Is Nothing Then
                MsgBox("Please select an Assigned Inspection")
                Exit Sub
            End If
            If ugRow.Cells("SUBMITTED").Value Is DBNull.Value Then
                MsgBox("Cannot accept Assigned Inspection - not submitted to C & E")
                Exit Sub
            Else
                pInspection.Retrieve(ugRow.Cells("INSPECTION_ID").Value)
                pInspection.InspectionAccepted = True
                pInspection.ModifiedBy = MusterContainer.AppUser.ID
                pInspection.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal, True)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                ' removing the row from the grid
                ugRow.Delete(False)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnAssignedInspExpandCollapseAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAssignedInspExpandCollapseAll.Click
        Try
            If btnAssignedInspExpandCollapseAll.Text = "Expand All" Then
                ExpandAll(True, ugAssignedInspections, btnAssignedInspExpandCollapseAll)
            Else
                ExpandAll(False, ugAssignedInspections, btnAssignedInspExpandCollapseAll)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnAssignedInspectionsDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAssignedInspectionsDelete.Click
        Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            If Not MusterContainer.AppUser.HasAccess(UIUtilsGen.ModuleID.CAE, MusterContainer.AppUser.UserKey, UIUtilsGen.EntityTypes.Inspection) Then
                MessageBox.Show("You do not have rights to save Inspection.", "ACCESS DENIED", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            ugRow = getSelectedRow()
            If ugRow Is Nothing Then
                MsgBox("Please select an Assigned Inspection to Delete")
                Exit Sub
            End If
            If ugRow.Cells("SUBMITTED").Value Is DBNull.Value And _
                ugRow.Cells("COMPLETED").Value Is DBNull.Value Then
                Dim oInspection As New MUSTER.BusinessLogic.pInspection
                oInspection.Retrieve(ugRow.Cells("INSPECTION_ID").Value)
                oInspection.Deleted = True
                oInspection.ModifiedBy = MusterContainer.AppUser.ID

                oInspection.Save(UIUtilsGen.ModuleID.CAE, MusterContainer.AppUser.UserKey, returnVal, True, , MusterContainer.AppUser.ID, False)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If
                ugRow.Delete(False)
            Else
                MsgBox("Cannot delete Assigned Inspection - Either Completed or Submitted to C & E")
                Exit Sub
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugAssignedInspections_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugAssignedInspections.InitializeLayout
        Try
            e.Layout.Bands(0).Columns("INSPECTION_TYPE_ID").Hidden = True
            e.Layout.Bands(0).Columns("INSPECTION_ID").Hidden = True
            e.Layout.Bands(0).Columns("STAFF_ID").Hidden = True
            e.Layout.Bands(0).Columns("OWNER PHONE").Hidden = True
            e.Layout.Bands(0).Columns("ADDRESS_LINE_ONE").Hidden = True
            e.Layout.Bands(0).Columns("COUNTY").Hidden = True
            e.Layout.Bands(0).Columns("CITY").Hidden = True

            e.Layout.Bands(0).Override.CellClickAction = CellClickAction.RowSelect

            e.Layout.Bands(0).Columns("ASSIGNED DATE").CellActivation = Activation.NoEdit
            e.Layout.Bands(0).Columns("COMPLETED").CellActivation = Activation.NoEdit
            e.Layout.Bands(0).Columns("SUBMITTED").CellActivation = Activation.NoEdit
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnAssignedRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAssignedRefresh.Click
        SetupTabs()
    End Sub
#End Region

#Region "Licensees"
    Private Sub SetupLicenseeRowComboValue(ByRef ug As Infragistics.Win.UltraWinGrid.UltraGridRow)
        Try
            If ug.Cells("WorkShop" + vbCrLf + "Date").Value Is DBNull.Value Then
                ug.Cells("WorkShop" + vbCrLf + "Result").Appearance.BackColor = Color.RosyBrown
                ug.Cells("WorkShop" + vbCrLf + "Result").Activation = Activation.NoEdit

                ug.Cells("WorkShop" + vbCrLf + "Date").Appearance.BackColor = Color.RosyBrown
                ug.Cells("WorkShop" + vbCrLf + "Date").Activation = Activation.NoEdit
            Else
                ug.Cells("WorkShop" + vbCrLf + "Result").Appearance.BackColor = Color.Yellow
                ug.Cells("WorkShop" + vbCrLf + "Result").Activation = Activation.AllowEdit

                ug.Cells("WorkShop" + vbCrLf + "Date").Appearance.BackColor = Color.Yellow
                ug.Cells("WorkShop" + vbCrLf + "Date").Activation = Activation.AllowEdit
            End If

            If ug.Cells("Show Cause" + vbCrLf + "Hearing Date").Value Is DBNull.Value Then
                ug.Cells("Show Cause" + vbCrLf + "Hearing Results").Appearance.BackColor = Color.RosyBrown
                ug.Cells("Show Cause" + vbCrLf + "Hearing Results").Activation = Activation.NoEdit

                ug.Cells("Show Cause" + vbCrLf + "Hearing Date").Appearance.BackColor = Color.RosyBrown
                ug.Cells("Show Cause" + vbCrLf + "Hearing Date").Activation = Activation.NoEdit
            Else
                ug.Cells("Show Cause" + vbCrLf + "Hearing Results").Appearance.BackColor = Color.Yellow
                ug.Cells("Show Cause" + vbCrLf + "Hearing Results").Activation = Activation.AllowEdit

                ug.Cells("Show Cause" + vbCrLf + "Hearing Date").Appearance.BackColor = Color.Yellow
                ug.Cells("Show Cause" + vbCrLf + "Hearing Date").Activation = Activation.AllowEdit
            End If

            If ug.Cells("Commission" + vbCrLf + "Hearing Date").Value Is DBNull.Value Then
                ug.Cells("Commission" + vbCrLf + "Hearing Results").Appearance.BackColor = Color.RosyBrown
                ug.Cells("Commission" + vbCrLf + "Hearing Results").Activation = Activation.NoEdit

                ug.Cells("Commission" + vbCrLf + "Hearing Date").Appearance.BackColor = Color.RosyBrown
                ug.Cells("Commission" + vbCrLf + "Hearing Date").Activation = Activation.NoEdit
            Else
                ug.Cells("Commission" + vbCrLf + "Hearing Results").Appearance.BackColor = Color.Yellow
                ug.Cells("Commission" + vbCrLf + "Hearing Results").Activation = Activation.AllowEdit

                ug.Cells("Commission" + vbCrLf + "Hearing Date").Appearance.BackColor = Color.Yellow
                ug.Cells("Commission" + vbCrLf + "Hearing Date").Activation = Activation.AllowEdit
            End If

            ' SHOW CAUSE HEARING RESULT
            If vListShowCauseHearingResult.FindByDataValue(ug.Cells("Show Cause" + vbCrLf + "Hearing Results").Value) Is Nothing Then
                ug.Cells("Show Cause" + vbCrLf + "Hearing Results").Value = DBNull.Value
            End If
            ' COMMISSION HEARING RESULT
            If vListCommissionHearingResult.FindByDataValue(ug.Cells("Commission" + vbCrLf + "Hearing Results").Value) Is Nothing Then
                ug.Cells("Commission" + vbCrLf + "Hearing Results").Value = DBNull.Value
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub btnLicenseeAddLCE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLicenseeAddLCE.Click
        Try
            If Not MusterContainer.AppUser.HasAccess(UIUtilsGen.ModuleID.CAE, MusterContainer.AppUser.UserKey, UIUtilsGen.EntityTypes.CAELicenseeCompliantEvent) Then
                MessageBox.Show("You do not have rights to save Licensee Compliance Event.", "ACCESS DENIED", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            pLCE.Retrieve(0)
            objLCE = New LicenseeComplianceEvent(pLCE, "ADD")
            objLCE.ShowDialog()
            'ugLicensees.DataSource = Nothing
            'ugLicensees.DataSource = pLCE.EntityTable
            SetupTabs()
            'ugLicensees.Rows.ExpandAll(True)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnLicenseeEditLCE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLicenseeEditLCE.Click
        Try
            If Not MusterContainer.AppUser.HasAccess(UIUtilsGen.ModuleID.CAE, MusterContainer.AppUser.UserKey, UIUtilsGen.EntityTypes.CAELicenseeCompliantEvent) Then
                MessageBox.Show("You do not have rights to save Licensee Compliance Event.", "ACCESS DENIED", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            If ugLicensees.ActiveRow.Band.Index = 1 Then
                If ugLicensees.ActiveRow.Cells("STATUS").Value = 995 Then
                    pLCE.Retrieve(ugLicensees.ActiveRow.Cells("LCE_ID").Value)
                    objLCE = New LicenseeComplianceEvent(pLCE)
                    objLCE.ShowDialog()
                    'ugLicensees.DataSource = Nothing
                    'ugLicensees.DataSource = pLCE.EntityTable
                    'ugLicensees.Rows.ExpandAll(True)
                    SetupTabs()
                Else
                    MsgBox("Only LCEs with Status=New can be  edited")
                End If
            Else
                MsgBox("Please select a LCE(LCE row on the grid) to EDIT")
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnLicenseeDeleteLCE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLicenseeDeleteLCE.Click
        Try
            If Not MusterContainer.AppUser.HasAccess(UIUtilsGen.ModuleID.CAE, MusterContainer.AppUser.UserKey, UIUtilsGen.EntityTypes.CAELicenseeCompliantEvent) Then
                MessageBox.Show("You do not have rights to save Licensee Compliance Event.", "ACCESS DENIED", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            If ugLicensees.ActiveRow.Band.Index = 1 Then
                Dim result As DialogResult
                result = MsgBox("Do you want to delete the selected LCE", MsgBoxStyle.YesNo)
                If result = DialogResult.Yes Then
                    pLCE.Retrieve(ugLicensees.ActiveRow.Cells("LCE_ID").Value)
                    If pLCE.Status = "NEW" Then
                        pLCE.Deleted = True
                        pLCE.ModifiedBy = MusterContainer.AppUser.ID
                        pLCE.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal)
                        If Not UIUtilsGen.HasRights(returnVal) Then
                            Exit Sub
                        End If

                        pLCE.Remove(ugLicensees.ActiveRow.Cells("LCE_ID").Value)
                        'ugLicensees.DataSource = Nothing
                        'ugLicensees.DataSource = pLCE.EntityTable
                        'ugLicensees.Rows.ExpandAll(True)
                        SetupTabs()
                    Else
                        MsgBox("Only manually LCE with Status = New may be deleted")
                    End If
                End If
            Else
                MsgBox("Please select a LCE (row on the grid) to delete")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnLicenseeProcessEscalations_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLicenseeProcessEscalations.Click
        Dim drow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim bol As Boolean
        Dim str As String
        Dim strCheckStatus As String = String.Empty

        Try
            If Not MusterContainer.AppUser.HasAccess(UIUtilsGen.ModuleID.CAE, MusterContainer.AppUser.UserKey, UIUtilsGen.EntityTypes.CAELicenseeCompliantEvent) Then
                MessageBox.Show("You do not have rights to save Licensee Compliance Event.", "ACCESS DENIED", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            For Each drow In ugLicensees.Rows
                For Each childrow As Infragistics.Win.UltraWinGrid.UltraGridRow In drow.ChildBands("StatustoLicensee".ToUpper).Rows
                    If childrow.Cells("Selected").Value Then
                        If strCheckStatus <> String.Empty And strCheckStatus <> childrow.Cells("LCE_Status").Value.ToString Then
                            MsgBox("Please select LCEs of same Status")
                            Exit Sub
                        End If
                        strCheckStatus = childrow.Cells("LCE_Status").Value
                        bol = True
                    End If
                Next
            Next
            If Not bol Then
                MsgBox("Please select the LCE to be processed")
                Exit Sub
            End If
            For Each drow In ugLicensees.Rows
                For Each childrow As Infragistics.Win.UltraWinGrid.UltraGridRow In drow.ChildBands("StatustoLicensee".ToUpper).Rows
                    If childrow.Cells("Selected").Value Then
                        pLCE.LCEInfo = pLCE.Retrieve(childrow.Cells("LCE_ID").Value)
                        str = pLCE.EscalationLogic(pLCE.LCEInfo)
                        If str <> String.Empty Then
                            objEscalationDates = New UserDate(str.Split("|")(0), pLCE, CType(str.Split("|")(1), Integer))
                            objEscalationDates.ShowDialog()
                            If objEscalationDates.bolEscalationCancelled = True Then
                                objEscalationDates.bolEscalationCancelled = False
                                Exit Sub
                            End If
                        End If

                        pLCE.OverrideAmount = -1.0
                        pLCE.OverrideDueDate = CDate("01/01/0001")
                        If pLCE.ID <= 0 Then
                            pLCE.CreatedBy = MusterContainer.AppUser.ID
                        Else
                            pLCE.ModifiedBy = MusterContainer.AppUser.ID
                        End If
                        pLCE.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal)
                        If Not UIUtilsGen.HasRights(returnVal) Then
                            Exit Sub
                        End If
                    End If
                Next
            Next
            'ugLicensees.DataSource = Nothing
            'ugLicensees.DataSource = pLCE.EntityTable()
            'ugLicensees.Rows.ExpandAll(True)
            SetupTabs()
            MsgBox("Processing(escalating) the selected LCE's is done successfully")
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnLicenseeViewEnforcementHistory_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLicenseeViewEnforcementHistory.Click
        Dim drow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim bol As Boolean
        Dim str As String
        Dim strCheckStatus As String = String.Empty

        Try
            For Each drow In ugLicensees.Rows
                For Each childrow As Infragistics.Win.UltraWinGrid.UltraGridRow In drow.ChildBands("StatustoLicensee".ToUpper).Rows
                    If childrow.Cells("Selected").Value Then
                        If strCheckStatus <> String.Empty And strCheckStatus <> childrow.Cells("LCE_Status").Value.ToString Then
                            MsgBox("Please select LCEs of same Status")
                            Exit Sub
                        End If
                        strCheckStatus = childrow.Cells("LCE_Status").Value
                        bol = True
                    End If
                Next
            Next
            If Not bol Then
                MsgBox("Please select the LCE to be View Enforcement History")
                Exit Sub
            End If

            For Each drow In ugLicensees.Rows
                For Each childrow As Infragistics.Win.UltraWinGrid.UltraGridRow In drow.ChildBands("StatustoLicensee".ToUpper).Rows
                    If childrow.Cells("Selected").Value Then
                        frmEnforcementHistory = New EnforcementHistory(False, , , , pLCE, childrow.Cells("Licensee_id").Value)
                        frmEnforcementHistory.ShowDialog()
                        SetupTabs()
                        Exit Sub
                    End If
                Next
            Next

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnLicenseeRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLicenseeRefresh.Click
        Try
            SetupTabs()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnLicenseeProcessRescissions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLicenseeProcessRescissions.Click
        Dim drow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim result As DialogResult
        Dim bol As Boolean = False
        Try
            If Not MusterContainer.AppUser.HasAccess(UIUtilsGen.ModuleID.CAE, MusterContainer.AppUser.UserKey, UIUtilsGen.EntityTypes.CAELicenseeCompliantEvent) Then
                MessageBox.Show("You do not have rights to save Licensee Compliance Event.", "ACCESS DENIED", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            For Each drow In ugLicensees.Rows
                For Each childrow As Infragistics.Win.UltraWinGrid.UltraGridRow In drow.ChildBands("StatustoLicensee".ToUpper).Rows
                    If childrow.Cells("Selected").Value Then
                        bol = True
                    End If
                Next
            Next
            If Not bol Then
                MsgBox("Please select the LCE to be processed")
                Exit Sub
            End If
            result = MsgBox("Do you want to process rescissions for the selected LCE's", MsgBoxStyle.YesNo)
            If result = DialogResult.No Then
                Exit Sub
            Else
                For Each drow In ugLicensees.Rows
                    For Each childrow As Infragistics.Win.UltraWinGrid.UltraGridRow In drow.ChildBands("StatustoLicensee".ToUpper).Rows
                        If childrow.Cells("Selected").Value Then
                            pLCE.Retrieve(childrow.Cells("LCE_ID").Value)
                            pLCE.PolicyAmount = -1.0
                            pLCE.OverrideAmount = -1.0
                            pLCE.ShowCauseDate = CDate("01/01/0001")
                            pLCE.CommissionDate = CDate("01/01/0001")
                            pLCE.Rescinded = True
                            pLCE.PendingLetter = 1251 ' NFA Rescind letter"
                            pLCE.PendingLetterTemplateNum = 1302
                            If pLCE.ID <= 0 Then
                                pLCE.CreatedBy = MusterContainer.AppUser.ID
                            Else
                                pLCE.ModifiedBy = MusterContainer.AppUser.ID
                            End If
                            pLCE.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal)
                            If Not UIUtilsGen.HasRights(returnVal) Then
                                Exit Sub
                            End If
                        End If
                    Next
                Next
                'ugLicensees.DataSource = Nothing
                'ugLicensees.DataSource = pLCE.EntityTable()
                'ugLicensees.Rows.ExpandAll(True)
                SetupTabs()
                MsgBox("Rescinding the selected LCEs are completed successfully")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnLicenseeGenerateLetter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLicenseeGenerateLetter.Click
        Dim drow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim result As DialogResult
        Dim bol As Boolean = False
        Try
            If Not MusterContainer.AppUser.HasAccess(UIUtilsGen.ModuleID.CAE, MusterContainer.AppUser.UserKey, UIUtilsGen.EntityTypes.CAELicenseeCompliantEvent) Then
                MessageBox.Show("You do not have rights to save Licensee Compliance Event.", "ACCESS DENIED", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            For Each drow In ugLicensees.Rows
                For Each childrow As Infragistics.Win.UltraWinGrid.UltraGridRow In drow.ChildBands("StatustoLicensee".ToUpper).Rows
                    If childrow.Cells("Selected").Value Then
                        bol = True
                    End If
                Next
            Next
            If Not bol Then
                MsgBox("Please select the LCE to be processed")
                Exit Sub
            End If
            Dim letters As New Reg_Letters
            Dim strErr As String = "Invalid / No Pending Letter for the following:"
            Dim bolCreateLetter As Boolean = False
            For Each drow In ugLicensees.Rows
                For Each childrow As Infragistics.Win.UltraWinGrid.UltraGridRow In drow.ChildBands("StatustoLicensee".ToUpper).Rows
                    bolCreateLetter = False
                    If childrow.Cells("Selected").Value Then
                        If childrow.Cells("PENDING_LETTER_TEMPLATE_NUM").Value Is DBNull.Value Then
                            strErr += pLCE.LicenseeName + " - Facility: " + pLCE.FacilityID.ToString + vbCrLf
                        ElseIf childrow.Cells("PENDING_LETTER_TEMPLATE_NUM").Value = 0 Then
                            strErr += pLCE.LicenseeName + " - Facility: " + pLCE.FacilityID.ToString + vbCrLf
                        Else
                            If pOCE Is Nothing Then
                                pOCE = New MUSTER.BusinessLogic.pOwnerComplianceEvent
                            End If

                            Dim documentID As Integer = 0
                            If letters.GenerateCAELCELetter(childrow, strErr, pOCE.GetLetterGeneratedDate(childrow.Cells("LCE_ID").Value, UIUtilsGen.EntityTypes.CAELicenseeCompliantEvent, , False), documentID) Then
                                pLCE.Retrieve(childrow.Cells("LCE_ID").Value)

                                If pLCE.PendingLetterTemplateNum = UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.LCELetterTemplateNum.NOV) Then
                                    pOCE.SaveLetterGeneratedDate(0, pLCE.ID, UIUtilsGen.EntityTypes.CAELicenseeCompliantEvent, UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.LCELetterTemplateNum.NOV), Today.Date, False, documentID, UIUtilsGen.ModuleID.CAE, MusterContainer.AppUser.UserKey, returnVal)
                                    If Not UIUtilsGen.HasRights(returnVal) Then
                                        Exit Sub
                                    End If
                                End If
                                pLCE.LetterGenerated = Now.Today
                                pLCE.PendingLetter = 0
                                pLCE.PendingLetterTemplateNum = 0
                                pLCE.LetterPrinted = False
                                If pLCE.ID <= 0 Then
                                    pLCE.CreatedBy = MusterContainer.AppUser.ID
                                Else
                                    pLCE.ModifiedBy = MusterContainer.AppUser.ID
                                End If
                                pLCE.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal)
                                If Not UIUtilsGen.HasRights(returnVal) Then
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                Next
            Next
            SetupTabs()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnLicenseesExpandCollapseAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLicenseesExpandCollapseAll.Click
        Try
            If btnLicenseesExpandCollapseAll.Text = "Expand All" Then
                ExpandAll(True, ugLicensees, btnLicenseesExpandCollapseAll)
            Else
                ExpandAll(False, ugLicensees, btnLicenseesExpandCollapseAll)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugLicensees_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugLicensees.CellChange
        Dim dtNull As Date = CDate("01/01/0001")
        Try
            If e.Cell.Row.Band.Index = 1 Then
                If pLCE.ID <> CType(e.Cell.Row.Cells("LCE_id").Value, Integer) Then
                    pLCE.LCEInfo = pLCE.LCECollection.Item(e.Cell.Row.Cells("LCE_id").Value.ToString)
                End If
                If ("Override" + vbCrLf + "Due Date").ToUpper.Equals(e.Cell.Column.Key.ToUpper) Then
                    If e.Cell.Text = "__/__/____" Then
                        e.Cell.Value = DBNull.Value
                        pLCE.OverrideDueDate = dtNull
                    Else
                        e.Cell.Value = e.Cell.Text
                        pLCE.OverrideDueDate = e.Cell.Text
                    End If
                ElseIf ("Date" + vbCrLf + "Received").ToUpper.Equals(e.Cell.Column.Key.ToUpper) Then
                    If e.Cell.Text = "__/__/____" Then
                        e.Cell.Value = DBNull.Value
                        pLCE.DateReceived = dtNull
                    Else
                        e.Cell.Value = e.Cell.Text
                        pLCE.DateReceived = e.Cell.Text
                    End If
                ElseIf ("Show Cause" + vbCrLf + "Hearing Date").ToUpper.Equals(e.Cell.Column.Key.ToUpper) Then
                    If e.Cell.Text = "__/__/____" Then
                        e.Cell.Value = DBNull.Value
                        pLCE.ShowCauseDate = dtNull
                    Else
                        e.Cell.Value = e.Cell.Text
                        pLCE.ShowCauseDate = e.Cell.Value
                    End If
                ElseIf ("Show Cause" + vbCrLf + "Hearing Results").ToUpper.Equals(e.Cell.Column.Key.ToUpper) Then
                    e.Cell.Value = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
                    pLCE.ShowCauseResults = e.Cell.Value
                ElseIf ("Commission" + vbCrLf + "Hearing Results").ToUpper.Equals(e.Cell.Column.Key.ToUpper) Then
                    e.Cell.Value = e.Cell.ValueListResolved.GetValue(e.Cell.ValueListResolved.SelectedItemIndex)
                    pLCE.CommissionResults = e.Cell.Value
                ElseIf ("Paid" + vbCrLf + "Amount").ToUpper.Equals(e.Cell.Column.Key.ToUpper) Then
                    e.Cell.Value = e.Cell.EditorResolved.Value
                    pLCE.PaidAmount = e.Cell.Value
                ElseIf ("Override" + vbCrLf + "Amount").ToUpper.Equals(e.Cell.Column.Key.ToUpper) Then
                    e.Cell.Value = e.Cell.EditorResolved.Value
                    pLCE.OverrideAmount = e.Cell.Value
                ElseIf ("Letter" + vbCrLf + "Printed").ToUpper.Equals(e.Cell.Column.Key.ToUpper) Then
                    e.Cell.Value = e.Cell.Text
                    pLCE.LetterPrinted = e.Cell.Value
                End If
            ElseIf e.Cell.Row.Band.Index = 2 Then
                If pLCE.ID <> CType(e.Cell.Row.ParentRow.Cells("LCE_id").Value, Integer) Then
                    pLCE.LCEInfo = pLCE.LCECollection.Item(e.Cell.Row.ParentRow.Cells("LCE_id").Value.ToString)
                End If
                If ("Due").ToUpper.Equals(e.Cell.Column.Key.ToUpper) Then
                    If e.Cell.Text = "__/__/____" Then
                        e.Cell.Value = DBNull.Value
                        pLCE.CitationDueDate = dtNull
                    Else
                        e.Cell.Value = e.Cell.Text
                        pLCE.CitationDueDate = e.Cell.Text
                    End If
                ElseIf ("Received").ToUpper.Equals(e.Cell.Column.Key.ToUpper) Then
                    If e.Cell.Text = "__/__/____" Then
                        e.Cell.Value = DBNull.Value
                        pLCE.CitationReceivedDate = dtNull
                    Else
                        e.Cell.Value = e.Cell.Text
                        pLCE.CitationReceivedDate = e.Cell.Text
                    End If
                End If
            End If

            'Code to propogate the checkmarks from status to LCE's
            Dim rowCol As Infragistics.Win.UltraWinGrid.RowsCollection
            'Essentially, during the CellChanged event, the cell’s Text property contains the value that has just been entered. 
            'This value can be examined to determine the state of the Checkbox.
            If ugLicensees.ActiveRow.Band.Index = 0 Then
                rowCol = ugLicensees.ActiveRow.ChildBands("StatustoLicensee".ToUpper).Rows
                If e.Cell.Text.ToUpper = "TRUE" Then
                    For Each dr As Infragistics.Win.UltraWinGrid.UltraGridRow In rowCol
                        dr.Cells("Selected").Value = True
                    Next
                Else
                    For Each dr As Infragistics.Win.UltraWinGrid.UltraGridRow In rowCol
                        dr.Cells("Selected").Value = False
                    Next
                End If
            End If
            If ugLicensees.ActiveRow.Band.Index = 1 And e.Cell.Text.ToUpper = "TRUE" Then
                e.Cell.Value = True
                e.Cell.Row.Selected = True
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugLicensees_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles ugLicensees.BeforeRowUpdate
        Try
            If e.Row.Band.Index = 1 Then
                pLCE.Retrieve(e.Row.Cells("LCE_ID").Value)
            ElseIf e.Row.Band.Index = 2 Then
                pLCE.Retrieve(e.Row.ParentRow.Cells("LCE_ID").Value)
            Else
                Exit Sub
            End If
            If pLCE.ID <= 0 Then
                pLCE.CreatedBy = MusterContainer.AppUser.ID
            Else
                pLCE.ModifiedBy = MusterContainer.AppUser.ID
            End If
            pLCE.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                e.Cancel = True
            End If

            ' updating escalation field
            If e.Row.Band.Index = 1 Then
                e.Row.Cells("Escalation").Value = pLCE.GetEscalation(pLCE.LCEInfo)
            ElseIf e.Row.Band.Index = 2 Then
                e.Row.ParentRow.Cells("Escalation").Value = pLCE.GetEscalation(pLCE.LCEInfo)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    'Private Sub ugLicensees_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugLicensees.AfterCellUpdate
    '    Try
    '        If ("Show Cause" + vbCrLf + "Hearing Date").ToUpper.Equals(e.Cell.Column.Key.ToUpper) Then
    '            If e.Cell.Text <> String.Empty And e.Cell.Text <> "__/__/____" Then
    '                UIUtilsGen.FillDateobjectValues(pLCE.ShowCauseDate, e.Cell.Text)
    '            Else
    '                e.Cell.Value = pLCE.ShowCauseDate
    '            End If
    '        End If
    '        If ("Commission" + vbCrLf + "Hearing Date").ToUpper.Equals(e.Cell.Column.Key.ToUpper) Then
    '            If e.Cell.Text <> String.Empty And e.Cell.Text <> "__/__/____" Then
    '                UIUtilsGen.FillDateobjectValues(pLCE.CommissionDate, e.Cell.Text)
    '            Else
    '                e.Cell.Value = pLCE.CommissionDate
    '            End If
    '        End If
    '        If ("Override" + vbCrLf + "Amount").ToUpper.Equals(e.Cell.Column.Key.ToUpper) Then
    '            pLCE.OverrideAmount = e.Cell.Value
    '        End If
    '        If ("Paid" + vbCrLf + "Amount").ToUpper.Equals(e.Cell.Column.Key.ToUpper) Then
    '            pLCE.PaidAmount = e.Cell.Value
    '        End If
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub

    Private Sub ugLicensees_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugLicensees.InitializeLayout
        Try
            e.Layout.Bands(1).Columns("Policy" + vbCrLf + "Amount").MaskInput = "$nnn,nnn,nnn.nn"
            e.Layout.Bands(1).Columns("Policy" + vbCrLf + "Amount").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
            e.Layout.Bands(1).Columns("Policy" + vbCrLf + "Amount").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw

            e.Layout.Bands(1).Columns("Settlement" + vbCrLf + "Amount").MaskInput = "$nnn,nnn,nnn.nn"
            e.Layout.Bands(1).Columns("Settlement" + vbCrLf + "Amount").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
            e.Layout.Bands(1).Columns("Settlement" + vbCrLf + "Amount").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw

            e.Layout.Bands(1).Columns("Override" + vbCrLf + "Amount").MaskInput = "$nnn,nnn,nnn.nn"
            e.Layout.Bands(1).Columns("Override" + vbCrLf + "Amount").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
            e.Layout.Bands(1).Columns("Override" + vbCrLf + "Amount").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw

            e.Layout.Bands(1).Columns("Paid" + vbCrLf + "Amount").MaskInput = "$nnn,nnn,nnn.nn"
            e.Layout.Bands(1).Columns("Paid" + vbCrLf + "Amount").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
            e.Layout.Bands(1).Columns("Paid" + vbCrLf + "Amount").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw

            'e.Layout.Bands(0).Columns.Add("--")
            'e.Layout.Bands(0).Columns.Add("Status")

            If e.Layout.Bands(1).Groups.Count = 0 Then
                e.Layout.Bands(1).Groups.Add("Row1")
                e.Layout.Bands(1).Groups.Add("Row2")
            End If

            e.Layout.Bands(1).GroupHeadersVisible = False

            'After you have bound your grid to a DataSource you should create an unbound column that will be used as your CheckBox column. In the InitializeLayout event add the following code to create an unbound column:
            e.Layout.Bands(0).Columns.Add("Selected").DataType = GetType(Boolean)
            e.Layout.Bands(0).Columns("Selected").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
            e.Layout.Bands(0).Columns("Selected").Header.VisiblePosition = 0

            e.Layout.Bands(1).Columns.Add("Selected").DataType = GetType(Boolean)
            e.Layout.Bands(1).Columns("Selected").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
            e.Layout.Bands(1).Columns("Selected").Header.VisiblePosition = 0

            e.Layout.Bands(1).Override.RowAppearance.BackColor = Color.RosyBrown
            e.Layout.Bands(2).Override.RowAppearance.BackColor = Color.Khaki

            ' setting editable cells color to yellow
            e.Layout.Bands(1).Columns("Override" + vbCrLf + "Due Date").CellAppearance.BackColor = Color.Yellow
            e.Layout.Bands(1).Columns("Override" + vbCrLf + "Amount").CellAppearance.BackColor = Color.Yellow
            e.Layout.Bands(1).Columns("Paid" + vbCrLf + "Amount").CellAppearance.BackColor = Color.Yellow
            e.Layout.Bands(1).Columns("Date" + vbCrLf + "Received").CellAppearance.BackColor = Color.Yellow
            e.Layout.Bands(1).Columns("WorkShop" + vbCrLf + "Date").CellAppearance.BackColor = Color.Yellow
            e.Layout.Bands(1).Columns("WorkShop" + vbCrLf + "Result").CellAppearance.BackColor = Color.Yellow
            e.Layout.Bands(1).Columns("Show Cause" + vbCrLf + "Hearing Date").CellAppearance.BackColor = Color.Yellow
            e.Layout.Bands(1).Columns("Show Cause" + vbCrLf + "Hearing Results").CellAppearance.BackColor = Color.Yellow
            e.Layout.Bands(1).Columns("Commission" + vbCrLf + "Hearing Results").CellAppearance.BackColor = Color.Yellow
            e.Layout.Bands(1).Columns("Commission" + vbCrLf + "Hearing Date").CellAppearance.BackColor = Color.Yellow
            e.Layout.Bands(1).Columns("Letter" + vbCrLf + "Printed").CellAppearance.BackColor = Color.Yellow

            e.Layout.Bands(2).Columns("Due").CellAppearance.BackColor = Color.Yellow
            e.Layout.Bands(2).Columns("Received").CellAppearance.BackColor = Color.Yellow

            'Me.ugLicensees.SupportThemes = True
            'For Each col In Me.ugLicensees.DisplayLayout.Bands(0).Columns
            '    col.Header.Appearance.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent
            '    col.Header.Appearance.BackColor = Color.DarkGray
            'Next
            'For Each col In Me.ugLicensees.DisplayLayout.Bands(1).Columns
            '    col.Header.Appearance.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent
            '    col.Header.Appearance.BackColor = Color.DarkGray
            'Next
            'For Each col In Me.ugLicensees.DisplayLayout.Bands(2).Columns
            '    col.Header.Appearance.ThemedElementAlpha = Infragistics.Win.Alpha.Transparent
            '    col.Header.Appearance.BackColor = Color.DarkGray
            'Next

            e.Layout.Bands(1).Columns("Selected").Group = e.Layout.Bands(1).Groups("Row1")
            e.Layout.Bands(1).Columns("FILLER5").Group = e.Layout.Bands(1).Groups("Row1")
            e.Layout.Bands(1).Columns("Licensee").Group = e.Layout.Bands(1).Groups("Row1")
            e.Layout.Bands(1).Columns("FILLER1").Group = e.Layout.Bands(1).Groups("Row1")
            e.Layout.Bands(1).Columns("Rescinded").Group = e.Layout.Bands(1).Groups("Row1")
            e.Layout.Bands(1).Columns("FILLER2").Group = e.Layout.Bands(1).Groups("Row1")
            e.Layout.Bands(1).Columns("LCE" + vbCrLf + "Date").Group = e.Layout.Bands(1).Groups("Row1")
            e.Layout.Bands(1).Columns("Last" + vbCrLf + "Process Date").Group = e.Layout.Bands(1).Groups("Row1")
            e.Layout.Bands(1).Columns("Next" + vbCrLf + "Due Date").Group = e.Layout.Bands(1).Groups("Row1")
            e.Layout.Bands(1).Columns("Override" + vbCrLf + "Due Date").Group = e.Layout.Bands(1).Groups("Row1")
            e.Layout.Bands(1).Columns("LCE_Status").Group = e.Layout.Bands(1).Groups("Row1")
            e.Layout.Bands(1).Columns("Escalation").Group = e.Layout.Bands(1).Groups("Row1")
            e.Layout.Bands(1).Columns("Policy" + vbCrLf + "Amount").Group = e.Layout.Bands(1).Groups("Row1")
            e.Layout.Bands(1).Columns("Override" + vbCrLf + "Amount").Group = e.Layout.Bands(1).Groups("Row1")
            e.Layout.Bands(1).Columns("Settlement" + vbCrLf + "Amount").Group = e.Layout.Bands(1).Groups("Row1")
            e.Layout.Bands(1).Columns("FILLER3").Group = e.Layout.Bands(1).Groups("Row1")
            e.Layout.Bands(1).Columns("Paid" + vbCrLf + "Amount").Group = e.Layout.Bands(1).Groups("Row1")
            e.Layout.Bands(1).Columns("Date" + vbCrLf + "Received").Group = e.Layout.Bands(1).Groups("Row1")
            e.Layout.Bands(1).Columns("WorkShop" + vbCrLf + "Date").Group = e.Layout.Bands(1).Groups("Row2")
            e.Layout.Bands(1).Columns("WorkShop" + vbCrLf + "Result").Group = e.Layout.Bands(1).Groups("Row2")
            e.Layout.Bands(1).Columns("Show Cause" + vbCrLf + "Hearing Date").Group = e.Layout.Bands(1).Groups("Row2")
            e.Layout.Bands(1).Columns("Show Cause" + vbCrLf + "Hearing Results").Group = e.Layout.Bands(1).Groups("Row2")
            e.Layout.Bands(1).Columns("Commission" + vbCrLf + "Hearing Date").Group = e.Layout.Bands(1).Groups("Row2")
            e.Layout.Bands(1).Columns("Commission" + vbCrLf + "Hearing Results").Group = e.Layout.Bands(1).Groups("Row2")
            e.Layout.Bands(1).Columns("Pending" + vbCrLf + "Letter").Group = e.Layout.Bands(1).Groups("Row2")
            e.Layout.Bands(1).Columns("Letter" + vbCrLf + "Printed").Group = e.Layout.Bands(1).Groups("Row2")
            e.Layout.Bands(1).Columns("Letter" + vbCrLf + "Generated").Group = e.Layout.Bands(1).Groups("Row2")
            e.Layout.Bands(1).Columns("FILLER4").Group = e.Layout.Bands(1).Groups("Row2")

            e.Layout.Bands(1).LevelCount = 2
            e.Layout.Bands(1).Columns("Selected").Level = 0
            e.Layout.Bands(1).Columns("FILLER5").Level = 1
            e.Layout.Bands(1).Columns("Licensee").Level = 0
            e.Layout.Bands(1).Columns("FILLER1").Level = 1
            e.Layout.Bands(1).Columns("Rescinded").Level = 0
            e.Layout.Bands(1).Columns("FILLER2").Level = 1
            e.Layout.Bands(1).Columns("LCE" + vbCrLf + "Date").Level = 0
            e.Layout.Bands(1).Columns("Last" + vbCrLf + "Process Date").Level = 1
            e.Layout.Bands(1).Columns("Next" + vbCrLf + "Due Date").Level = 0
            e.Layout.Bands(1).Columns("Override" + vbCrLf + "Due Date").Level = 1
            e.Layout.Bands(1).Columns("LCE_Status").Level = 0
            e.Layout.Bands(1).Columns("Escalation").Level = 1
            e.Layout.Bands(1).Columns("Policy" + vbCrLf + "Amount").Level = 0
            e.Layout.Bands(1).Columns("Override" + vbCrLf + "Amount").Level = 1
            e.Layout.Bands(1).Columns("Settlement" + vbCrLf + "Amount").Level = 0
            e.Layout.Bands(1).Columns("FILLER3").Level = 1
            e.Layout.Bands(1).Columns("Paid" + vbCrLf + "Amount").Level = 0
            e.Layout.Bands(1).Columns("Date" + vbCrLf + "Received").Level = 1
            e.Layout.Bands(1).Columns("WorkShop" + vbCrLf + "Date").Level = 0
            e.Layout.Bands(1).Columns("WorkShop" + vbCrLf + "Result").Level = 1
            e.Layout.Bands(1).Columns("Show Cause" + vbCrLf + "Hearing Date").Level = 0
            e.Layout.Bands(1).Columns("Show Cause" + vbCrLf + "Hearing Results").Level = 1
            e.Layout.Bands(1).Columns("Commission" + vbCrLf + "Hearing Results").Level = 1
            e.Layout.Bands(1).Columns("Commission" + vbCrLf + "Hearing Date").Level = 0
            e.Layout.Bands(1).Columns("Pending" + vbCrLf + "Letter").Level = 0
            e.Layout.Bands(1).Columns("Letter" + vbCrLf + "Printed").Level = 1
            e.Layout.Bands(1).Columns("Letter" + vbCrLf + "Generated").Level = 0
            e.Layout.Bands(1).Columns("FILLER4").Level = 1

            e.Layout.Bands(1).Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Free
            e.Layout.Bands(1).Override.CellMultiLine = Infragistics.Win.DefaultableBoolean.False
            e.Layout.Bands(1).Override.DefaultColWidth = 80
            e.Layout.Bands(1).ColHeaderLines = 2
            e.Layout.Bands(1).Override.RowSelectorWidth = 2
            e.Layout.Bands(1).Columns("Licensee").CellMultiLine = Infragistics.Win.DefaultableBoolean.False

            e.Layout.Bands(1).Columns("FILLER1").Header.Caption = ""
            e.Layout.Bands(1).Columns("FILLER1").CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled
            'e.Layout.Bands(1).Columns("FILLER1").CellAppearance.BackColorDisabled = Color.LightGoldenrodYellow

            e.Layout.Bands(1).Columns("FILLER2").Header.Caption = ""
            e.Layout.Bands(1).Columns("FILLER2").CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled
            'e.Layout.Bands(1).Columns("FILLER2").CellAppearance.BackColorDisabled = Color.LightGoldenrodYellow

            e.Layout.Bands(1).Columns("FILLER3").Header.Caption = ""
            e.Layout.Bands(1).Columns("FILLER3").CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled
            'e.Layout.Bands(1).Columns("FILLER3").CellAppearance.BackColorDisabled = Color.LightGoldenrodYellow

            e.Layout.Bands(1).Columns("FILLER4").Header.Caption = ""
            e.Layout.Bands(1).Columns("FILLER4").CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled
            'e.Layout.Bands(1).Columns("FILLER4").CellAppearance.BackColorDisabled = Color.LightGoldenrodYellow

            e.Layout.Bands(1).Columns("FILLER5").Header.Caption = ""
            e.Layout.Bands(1).Columns("FILLER5").CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled
            'e.Layout.Bands(1).Columns("FILLER5").CellAppearance.BackColorDisabled = Color.LightGoldenrodYellow

            e.Layout.Bands(1).Columns("LCE_ID").Hidden = True
            e.Layout.Bands(1).Columns("WorkShop" + vbCrLf + "Date").Hidden = True
            e.Layout.Bands(1).Columns("WorkShop" + vbCrLf + "Result").Hidden = True
            e.Layout.Bands(1).Columns("PENDING_LETTER_TEMPLATE_NUM").Hidden = True

            'For Each dr As Infragistics.Win.UltraWinGrid.UltraGridRow In ugGrid.Rows
            '    If Not dr.ChildBands Is Nothing Then

            '        ' Loop throgh each of the child bands.
            '        Dim childBand As Infragistics.Win.UltraWinGrid.UltraGridChildBand = Nothing
            '        For Each childBand In dr.ChildBands
            '            For Each dr1 As Infragistics.Win.UltraWinGrid.UltraGridRow In childBand.Rows
            '                If dr1.Cells("LCE" + vbCrLf + "Date").Text = "01/01/0001" Then
            '                    dr1.Cells("LCE" + vbCrLf + "Date").Value = DBNull.Value
            '                End If
            '            Next
            '        Next
            '    End If
            'Next

            e.Layout.Bands(0).Columns("LCE_Status").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            e.Layout.Bands(1).Columns("LCE_Status").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("Licensee").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("Rescinded").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("LCE" + vbCrLf + "DATE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("Last" + vbCrLf + "Process Date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("Next" + vbCrLf + "Due Date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("Escalation").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("Policy" + vbCrLf + "Amount").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("Settlement" + vbCrLf + "Amount").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("WorkShop" + vbCrLf + "Date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("WorkShop" + vbCrLf + "Result").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("Pending" + vbCrLf + "Letter").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("Letter" + vbCrLf + "Generated").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            'e.Layout.Bands(1).Columns("Letter" + vbCrLf + "Printed").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            e.Layout.Bands(2).Columns("FACILITY_ID").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(2).Columns("Facility").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(2).Columns("Citation").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(2).Columns("Citation Text").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            e.Layout.Bands(2).Columns("Citation Text").Width = 400
            e.Layout.Bands(2).Columns("Facility").Width = 200
            e.Layout.Bands(1).Columns("Licensee").Width = 150
            e.Layout.Bands(1).Columns("FILLER1").Width = 150

            ' Set the Style to DropDownList.
            e.Layout.Bands(1).Columns("Show Cause" + vbCrLf + "Hearing Results").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
            e.Layout.Bands(1).Columns("Commission" + vbCrLf + "Hearing Results").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList

            '' populate columns
            'If e.Layout.ValueLists.All.Length = 0 Then
            '    e.Layout.ValueLists.Add("ATTENDED")
            '    e.Layout.ValueLists.Add("NO SHOW")
            'End If

            ' populate the whole column as the table is the same for each row
            ' Show cause hearing results -property type id is 129 
            If e.Layout.Bands(1).Columns("Show Cause" + vbCrLf + "Hearing Results").ValueList Is Nothing Then
                vListShowCauseHearingResult = New Infragistics.Win.ValueList
                For Each row As DataRow In pLCE.getDropDownValues(129).Rows
                    vListShowCauseHearingResult.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                Next
                e.Layout.Bands(1).Columns("Show Cause" + vbCrLf + "Hearing Results").ValueList = vListShowCauseHearingResult
            End If

            If e.Layout.Bands(1).Columns("Commission" + vbCrLf + "Hearing Results").ValueList Is Nothing Then
                vListCommissionHearingResult = New Infragistics.Win.ValueList
                For Each row As DataRow In pLCE.getDropDownValues(130).Rows
                    vListCommissionHearingResult.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                Next
                e.Layout.Bands(1).Columns("Commission" + vbCrLf + "Hearing Results").ValueList = vListCommissionHearingResult
            End If

            For Each ugRow In ugLicensees.Rows
                If Not ugRow.ChildBands Is Nothing Then
                    If Not ugRow.ChildBands(0).Rows Is Nothing Then
                        For Each ugChildRow In ugRow.ChildBands(0).Rows
                            SetupLicenseeRowComboValue(ugChildRow)
                        Next
                    End If
                End If
            Next
            'If e.Layout.Bands(1).Columns("LCE_Status").ValueList Is Nothing Then
            '    Dim vListShowCauseHearingResult As New Infragistics.Win.ValueList
            '    For Each row As DataRow In pLCE.getDropDownValues(128).Rows
            '        vListShowCauseHearingResult.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
            '    Next
            '    e.Layout.Bands(1).Columns("LCE_Status").ValueList = vListShowCauseHearingResult
            'End If

            'If e.Layout.Bands(1).Columns("Pending" + vbCrLf + "Letter").ValueList Is Nothing Then
            '    Dim vListShowCauseHearingResult As New Infragistics.Win.ValueList
            '    For Each row As DataRow In pLCE.getDropDownValues(131).Rows
            '        vListShowCauseHearingResult.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
            '    Next
            '    e.Layout.Bands(1).Columns("Pending" + vbCrLf + "Letter").ValueList = vListShowCauseHearingResult
            'End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

        '' Set the scroll style to immediate so the rows get scrolled immediately
        '' when the vertical scrollbar thumb is dragged.
        ''
        'e.Layout.ScrollStyle = Infragistics.Win.UltraWinGrid.ScrollStyle.Immediate

        '' ScrollBounds of ScrollToFill will prevent the user from scrolling the
        '' grid further down once the last row becomes fully visible.
        ''
        'e.Layout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill

        'With e.Layout.Override
        '    .RowAppearance.BackColorAlpha = Infragistics.Win.Alpha.Transparent

        '    'use the same appearance for alternate rows
        '    .RowAlternateAppearance = .RowAppearance
        '    .CellAppearance.BackColorAlpha = Infragistics.Win.Alpha.UseAlphaLevel
        '    .CellAppearance.AlphaLevel = 192

        '    .HeaderAppearance.AlphaLevel = 192
        '    .HeaderAppearance.BackColorAlpha = Infragistics.Win.Alpha.UseAlphaLevel
        'End With
        'change the rowConnector style
        'ugLicensees.DisplayLayout.RowConnectorStyle = Infragistics.Win.UltraWinGrid.RowConnectorStyle.Solid
    End Sub
#End Region

    Sub SetUpComboBox()

        Dim oinspectorowner As New BusinessLogic.pInspectorOwnerAssignment

        cmbManager.DisplayMember = "Description"
        cmbManager.ValueMember = "STAFF_ID"
        cmbManager.DataSource = oinspectorowner.getCNEManagers

        oinspectorowner = Nothing

    End Sub

    Sub ChangeSetUp(ByVal sender As Object, ByVal e As EventArgs)
        SetupTabs()
    End Sub

#Region "Form Events"
    Private Sub CandEManagement_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try

            SetUpComboBox()

            cmbManager.SelectedIndex = 0

            For Each item As DataRowView In cmbManager.Items
                If item(0) = MusterContainer.AppUser.UserKey Then
                    cmbManager.SelectedItem = item
                    Exit For
                End If
            Next

            RemoveHandler cmbManager.SelectedIndexChanged, AddressOf ChangeSetUp
            AddHandler cmbManager.SelectedIndexChanged, AddressOf ChangeSetUp

            SetupTabs()


            Dim OAdminUsers As List(Of String) = pFCE.GetAdminUSersList()
            If OAdminUsers IsNot Nothing And OAdminUsers.Count > 0 Then
                If OAdminUsers.Contains(MusterContainer.AppUser.ID.ToUpper) Then
                    btnManualEsc.Enabled = True
                Else
                    btnManualEsc.Enabled = False
                End If
            Else
                btnManualEsc.Enabled = False
            End If

            ''' Commented Code by Susheel on 19th Feb 2016
            'If Not (MusterContainer.AppUser.ID.ToUpper = "ADMIN" OrElse MusterContainer.AppUser.ID.ToUpper = "MPIGFORD" OrElse MusterContainer.AppUser.ID.ToUpper = "LORIN" OrElse MusterContainer.AppUser.ID.ToUpper = "LYNN" OrElse MusterContainer.AppUser.ID.ToUpper = "LANDO" OrElse MusterContainer.AppUser.ID.ToUpper = "SLATER" OrElse MusterContainer.AppUser.ID.ToUpper = "BRANDON" OrElse MusterContainer.AppUser.ID.ToUpper = "JRHEMANN") Then
            '    btnManualEsc.Enabled = False
            'End If
            ''' End

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub


    Private Sub brnAgreedOrder_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles brnAgreedOrder.Click
        If bolLoading Then Exit Sub
        Dim bolLetterGenerated As Boolean
        Dim bolNFARescindLetterGenerated As Boolean = False
        Dim thisDate As Date
        Try

            thisDate = New Date(Now.Year, Now.Month, Now.Day, Now.Hour, Now.Minute, Now.Second)

            If Not MusterContainer.AppUser.HasAccess(UIUtilsGen.ModuleID.CAE, MusterContainer.AppUser.UserKey, UIUtilsGen.EntityTypes.CAEOwnerComplianceEvent) Then
                MessageBox.Show("You do not have rights to save Owner Compliance Event.", "ACCESS DENIED", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End If
            bolLoading = True
            bolLetterGenerated = False
            For Each ugRow In ugEnforcement.Rows ' status

                For Each ugChildRow In ugRow.ChildBands(0).Rows ' owner
                    If ugChildRow.Cells("SELECTED").Value = True Or _
                        ugChildRow.Cells("SELECTED").Text = True Then

                        ' Generate Letter
                        lastUGRowOwner = ugChildRow.Cells("OWNER_ID").Value

                        If Not GenerateOCELetter(ugChildRow, True) Then
                            Exit Sub
                        End If
                    End If
                Next
            Next


            If Not bolLetterGenerated Then
                MsgBox("No Letter(s) Generated")

            ElseIf bolNFARescindLetterGenerated Then
                SetupTabs()
            Else

                If Not ugChildRow Is Nothing Then
                    btnEnforceRefresh.PerformClick()
                End If

            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            bolLoading = False
            thisDate = Nothing
        End Try
    End Sub

    Private Sub CandEManagement_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Try
            MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "CandEManagement")
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub CandEManagement_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.TextChanged
        Try
            MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "WindowName", Me.Text, "CandEManagement")
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "External Events"
    Private Sub objLCE_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles objLCE.Closing
        Dim result As DialogResult
        Try
            If pLCE.colIsDirty Then
                result = MsgBox("There are unsaved changes. Do you want to save them now", MsgBoxStyle.YesNoCancel)
                If result = DialogResult.Yes Then
                    If objLCE.ValidateData() Then
                        'pLCE.Save()
                        objLCE.btnSave.PerformClick() '.btnSave_Click(sender, New EventArgs)
                        SetupTabs()
                        'e.Cancel = True
                    Else
                        e.Cancel = True
                    End If
                ElseIf result = DialogResult.No Then
                    SetupTabs()
                    Exit Sub
                ElseIf result = DialogResult.Cancel Then
                    e.Cancel = True
                    Exit Sub
                End If
            End If
            'ugLicensees.DataSource = Nothing
            'ugLicensees.DataSource = pLCE.EntityTable
            'ugLicensees.Rows.ExpandAll(True)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

#End Region

#Region "UI Control Events"
    Private Sub tabCntrlCandE_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabCntrlCandE.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            SetupTabs()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region


End Class
