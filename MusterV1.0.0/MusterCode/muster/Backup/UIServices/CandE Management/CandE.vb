Imports Infragistics.Win.UltraWinGrid

Public Class CandE
    Inherits System.Windows.Forms.Form

#Region "User Defined Variables"
    'Private frmCAEMgmt As New CandEManagement
    Private WithEvents SF As ShowFlags
    Public MyGuid As New System.Guid
    Private bolLoading As Boolean = True
    Private bolFrmActivated As Boolean = False
    Private pOwn As MUSTER.BusinessLogic.pOwner
    Private pFCE As New MUSTER.BusinessLogic.pFacilityComplianceEvent
    Private pOCE As New MUSTER.BusinessLogic.pOwnerComplianceEvent
    Private WithEvents frmEnforcementHistory As EnforcementHistory
    Private frmInspecHistory As InspectionHistory

    ' variables for OCE Grid
    Private vListWorkShopResult, vListShowCauseResult, vListCommissionResult As Infragistics.Win.ValueList

    Dim rp As New Remove_Pencil

    Private ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
    Dim ugChildRow As Infragistics.Win.UltraWinGrid.UltraGridRow
    Dim ugGrandChildRow As Infragistics.Win.UltraWinGrid.UltraGridRow

    'Variables for Facility form
    Private nFacilityOwnerID As Integer
    Public strFacilityIdTags As String
    Private strFacilityAddress As String
    Private bolDblClick As Boolean = False
    Private nIncrement As Integer = 0

    Public mContainer As MusterContainer

    'contacts
    Private pConStruct As MUSTER.BusinessLogic.pContactStruct
    Private WithEvents oCompanySearch As CompanySearch
    Private WithEvents objCntSearch As ContactSearch
    Public nOwnerID, nFacilityID As Integer
    Private returnVal As String = String.Empty
    Private dsContacts As DataSet
    Private strFilterString As String = String.Empty
    'Private oEntity As New MUSTER.BusinessLogic.pEntity
    'Public strFacilityIdTags As String
    ''Public MyGuid As New System.Guid
    'Public bolNewPersona As Boolean = False
    'Private bolDisplayErrmessage As Boolean = True
    'Private nErrMessage As Integer = 0
    'Private bolValidateSuccess As Boolean = True
    'Private oAddressInfo As MUSTER.Info.AddressInfo
    'Private WithEvents AddressForm As Address
    'Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long
    'Private strActivetbPage As String
    'Private WithEvents ContactFrm As Contacts
    'Private WithEvents objCntSearch As ContactSearch
    'Private dsContacts As DataSet
    'Private result As DialogResult
    'Private bolFrmActivated As Boolean = False
#End Region

#Region " Windows Form Designer generated code "
    Public Sub New(ByRef oOwner As MUSTER.BusinessLogic.pOwner, Optional ByVal OwnerID As Int64 = 0, Optional ByVal FacilityID As Int64 = 0)
        MyBase.New()

        MyGuid = System.Guid.NewGuid
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        'Add any initialization after the InitializeComponent() call

        Cursor.Current = Cursors.AppStarting
        bolLoading = True
        pOwn = oOwner
        pConStruct = New MUSTER.BusinessLogic.pContactStruct
        '
        'Need to tell the AppUser that we've instantiated another CAE form...
        '
        MusterContainer.AppUser.LogEntry("C " + "&" + " E", MyGuid.ToString)
        '
        ' The following line enables all forms to detect the visible form in the MDI container
        '
        MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "C " + "&" + " E")

        UIUtilsGen.PopulateOwnerType(Me.cmbOwnerType, pOwn)
        UIUtilsGen.PopulateFacilityType(Me.cmbFacilityType, pOwn.Facilities)
        UIUtilsGen.PopulateFacilityMethod(Me.cmbFacilityMethod, pOwn.Facilities)
        UIUtilsGen.PopulateFacilityDatum(Me.cmbFacilityDatum, pOwn.Facilities)
        UIUtilsGen.PopulateFacilityLocationType(Me.cmbFacilityLocationType, pOwn.Facilities)

        If OwnerID > 0 Then
            PopulateOwner(Integer.Parse(OwnerID))
            If FacilityID <= 0 Then
                tbCntrlCandE.SelectedTab = tbPageOwnerDetail
                Me.Text = "C " + "&" + " E - Owner Detail - " + OwnerID.ToString + "(" & txtOwnerName.Text & ")"
            End If
            nOwnerID = OwnerID
            CommentsMaintenance(, , True)
        End If

        If FacilityID > 0 Then
            tbCntrlCandE.SelectedTab = tbPageFacilityDetail
            PopulateFacility(Integer.Parse(FacilityID))
            nFacilityID = FacilityID
        End If

        bolLoading = False
        Cursor.Current = Cursors.Default
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        '
        ' Remove any values from the shared collection for this screen
        '
        MusterContainer.AppSemaphores.Remove(MyGuid.ToString)
        '
        ' Log the disposal of the form (exit from Registration form)
        '
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
    Friend WithEvents pnlTop As System.Windows.Forms.Panel
    Friend WithEvents tbPageOwnerDetail As System.Windows.Forms.TabPage
    Public WithEvents pnlOwnerDetail As System.Windows.Forms.Panel
    Public WithEvents chkOwnerAgencyInterest As System.Windows.Forms.CheckBox
    Public WithEvents lblOwnerActiveOrNot As System.Windows.Forms.Label
    Friend WithEvents LinkLblCAPSignup As System.Windows.Forms.LinkLabel
    Public WithEvents lblCAPParticipationLevel As System.Windows.Forms.Label
    Public WithEvents mskTxtOwnerFax As AxMSMask.AxMaskEdBox
    Public WithEvents mskTxtOwnerPhone2 As AxMSMask.AxMaskEdBox
    Public WithEvents mskTxtOwnerPhone As AxMSMask.AxMaskEdBox
    Friend WithEvents lblOwnerEmail As System.Windows.Forms.Label
    Public WithEvents txtOwnerEmail As System.Windows.Forms.TextBox
    Friend WithEvents lblFax As System.Windows.Forms.Label
    Public WithEvents txtOwnerAddress As System.Windows.Forms.TextBox
    Friend WithEvents lblOwnerAddress As System.Windows.Forms.Label
    Public WithEvents txtOwnerName As System.Windows.Forms.TextBox
    Friend WithEvents lblOwnerName As System.Windows.Forms.Label
    Friend WithEvents lblOwnerStatus As System.Windows.Forms.Label
    Friend WithEvents lblOwnerCapParticipant As System.Windows.Forms.Label
    Friend WithEvents lblPhone2 As System.Windows.Forms.Label
    Friend WithEvents lblOwnerType As System.Windows.Forms.Label
    Public WithEvents txtOwnerAIID As System.Windows.Forms.TextBox
    Friend WithEvents lblOwnerAIID As System.Windows.Forms.Label
    Public WithEvents lblOwnerIDValue As System.Windows.Forms.Label
    Friend WithEvents lblOwnerID As System.Windows.Forms.Label
    Friend WithEvents lblOwnerPhone As System.Windows.Forms.Label
    Public WithEvents cmbOwnerType As System.Windows.Forms.ComboBox
    Friend WithEvents tbPageFacilityDetail As System.Windows.Forms.TabPage
    Friend WithEvents pnlFacilityBottom As System.Windows.Forms.Panel
    Friend WithEvents tbCntrlFacility As System.Windows.Forms.TabControl
    Friend WithEvents pnlFacilityLustButton As System.Windows.Forms.Panel
    Public WithEvents lblTotalNoOfLUSTEventsValue As System.Windows.Forms.Label
    Friend WithEvents lblTotalNoOfLUSTEvents As System.Windows.Forms.Label
    Friend WithEvents pnl_FacilityDetail As System.Windows.Forms.Panel
    Friend WithEvents btnFacComments As System.Windows.Forms.Button
    Friend WithEvents btnFacFlags As System.Windows.Forms.Button
    Public WithEvents dtPickUpcomingInstallDateValue As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblUpcomingInstallDate As System.Windows.Forms.Label
    Public WithEvents chkUpcomingInstall As System.Windows.Forms.CheckBox
    Friend WithEvents lnkLblNextFac As System.Windows.Forms.LinkLabel
    Public WithEvents lblCAPStatusValue As System.Windows.Forms.Label
    Friend WithEvents lblCAPStatus As System.Windows.Forms.Label
    Public WithEvents txtFuelBrand As System.Windows.Forms.TextBox
    Friend WithEvents ll As System.Windows.Forms.Label
    Public WithEvents dtFacilityPowerOff As System.Windows.Forms.DateTimePicker
    Friend WithEvents lnkLblPrevFacility As System.Windows.Forms.LinkLabel
    Friend WithEvents lblLUSTSite As System.Windows.Forms.Label
    Public WithEvents chkLUSTSite As System.Windows.Forms.CheckBox
    Friend WithEvents lblPowerOff As System.Windows.Forms.Label
    Friend WithEvents lblCAPCandidate As System.Windows.Forms.Label
    Public WithEvents chkCAPCandidate As System.Windows.Forms.CheckBox
    Friend WithEvents lblFacilityLocationType As System.Windows.Forms.Label
    Public WithEvents cmbFacilityLocationType As System.Windows.Forms.ComboBox
    Friend WithEvents lblFacilityMethod As System.Windows.Forms.Label
    Public WithEvents cmbFacilityMethod As System.Windows.Forms.ComboBox
    Friend WithEvents lblFacilityDatum As System.Windows.Forms.Label
    Public WithEvents cmbFacilityDatum As System.Windows.Forms.ComboBox
    Public WithEvents cmbFacilityType As System.Windows.Forms.ComboBox
    Public WithEvents txtFacilityLatSec As System.Windows.Forms.TextBox
    Public WithEvents txtFacilityLongSec As System.Windows.Forms.TextBox
    Public WithEvents txtFacilityLatMin As System.Windows.Forms.TextBox
    Public WithEvents txtFacilityLongMin As System.Windows.Forms.TextBox
    Friend WithEvents lblFacilityLongMin As System.Windows.Forms.Label
    Friend WithEvents lblFacilityLongSec As System.Windows.Forms.Label
    Friend WithEvents lblFacilityLatMin As System.Windows.Forms.Label
    Friend WithEvents lblFacilityLatSec As System.Windows.Forms.Label
    Friend WithEvents lblFacilityLongDegree As System.Windows.Forms.Label
    Public WithEvents mskTxtFacilityFax As AxMSMask.AxMaskEdBox
    Public WithEvents mskTxtFacilityPhone As AxMSMask.AxMaskEdBox
    Public WithEvents txtFacilityAddress As System.Windows.Forms.TextBox
    Friend WithEvents lblFacilitySIC As System.Windows.Forms.Label
    Friend WithEvents lblFacilityFax As System.Windows.Forms.Label
    Public WithEvents dtPickFacilityRecvd As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblDateReceived As System.Windows.Forms.Label
    Public WithEvents chkSignatureofNF As System.Windows.Forms.CheckBox
    Friend WithEvents lblFacilitySigOnNF As System.Windows.Forms.Label
    Friend WithEvents lblFacilityFuelBrand As System.Windows.Forms.Label
    Public WithEvents lblFacilityStatusValue As System.Windows.Forms.Label
    Friend WithEvents lblFacilityStatus As System.Windows.Forms.Label
    Public WithEvents txtFacilityLongDegree As System.Windows.Forms.TextBox
    Public WithEvents txtFacilityLatDegree As System.Windows.Forms.TextBox
    Friend WithEvents lblFacilityLongitude As System.Windows.Forms.Label
    Friend WithEvents lblFacilityLatitude As System.Windows.Forms.Label
    Friend WithEvents lblFacilityType As System.Windows.Forms.Label
    Public WithEvents txtFacilityAIID As System.Windows.Forms.TextBox
    Friend WithEvents lblfacilityAIID As System.Windows.Forms.Label
    Public WithEvents lblFacilityIDValue As System.Windows.Forms.Label
    Friend WithEvents lblFacilityID As System.Windows.Forms.Label
    Friend WithEvents lblFacilityPhone As System.Windows.Forms.Label
    Public WithEvents txtFacilityName As System.Windows.Forms.TextBox
    Friend WithEvents lblFacilityAddress As System.Windows.Forms.Label
    Friend WithEvents lblFacilityName As System.Windows.Forms.Label
    Friend WithEvents lblFacilityLatDegree As System.Windows.Forms.Label
    Friend WithEvents tbPageSummary As System.Windows.Forms.TabPage
    Friend WithEvents Panel12 As System.Windows.Forms.Panel
    Friend WithEvents tbCntrlCandE As System.Windows.Forms.TabControl
    Friend WithEvents tbPageFCEs As System.Windows.Forms.TabPage
    Friend WithEvents tbPageOwnerCitations As System.Windows.Forms.TabPage
    Friend WithEvents tbPageFacilityCitations As System.Windows.Forms.TabPage
    Friend WithEvents tbCtrlOwner As System.Windows.Forms.TabControl
    Friend WithEvents tbPageOwnerFacilities As System.Windows.Forms.TabPage
    Public WithEvents ugFacilityList As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlOwnerFacilityBottom As System.Windows.Forms.Panel
    Public WithEvents lblNoOfFacilitiesValue As System.Windows.Forms.Label
    Friend WithEvents lblNoOfFacilities As System.Windows.Forms.Label
    Friend WithEvents tbPageOCEs As System.Windows.Forms.TabPage
    Friend WithEvents pnlOwnerBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlOwnerOCEsBottom As System.Windows.Forms.Panel
    Public WithEvents ugOCEs As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlOwnerButtons As System.Windows.Forms.Panel
    Friend WithEvents btnOwnerFlag As System.Windows.Forms.Button
    Friend WithEvents btnOwnerComment As System.Windows.Forms.Button
    Public WithEvents ugFCEs As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlOwnerCitationsHead As System.Windows.Forms.Panel
    Friend WithEvents btnExpandOwnerCitationAll As System.Windows.Forms.Button
    Friend WithEvents ugOCE2 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlFacilityCitationHead As System.Windows.Forms.Panel
    Friend WithEvents ugFCE2 As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnFacilityCancel As System.Windows.Forms.Button
    Friend WithEvents btnFacilitySave As System.Windows.Forms.Button
    Public WithEvents lblOwnerLastEditedOn As System.Windows.Forms.Label
    Public WithEvents lblOwnerLastEditedBy As System.Windows.Forms.Label
    Friend WithEvents btnExpandFacilityCitationAll As System.Windows.Forms.Button
    Friend WithEvents btnEnforceViewEnforceHistory As System.Windows.Forms.Button
    Friend WithEvents btnEnforceViewEnforceHistory1 As System.Windows.Forms.Button
    Public WithEvents lblNoOfOCEsValue As System.Windows.Forms.Label
    Friend WithEvents lblNoOfOCEs As System.Windows.Forms.Label
    Friend WithEvents tbPageOwnerDocuments As System.Windows.Forms.TabPage
    Friend WithEvents tbPageFacilityDocuments As System.Windows.Forms.TabPage
    Friend WithEvents UCOwnerDocuments As MUSTER.DocumentViewControl
    Friend WithEvents UCFacilityDocuments As MUSTER.DocumentViewControl
    Friend WithEvents btnInspHistory As System.Windows.Forms.Button
    Friend WithEvents btnInspHistory1 As System.Windows.Forms.Button
    Friend WithEvents btnCAEOwnerLabels As System.Windows.Forms.Button
    Friend WithEvents btnCAEOwnerEnvelopes As System.Windows.Forms.Button
    Friend WithEvents btnCAEFacLabels As System.Windows.Forms.Button
    Friend WithEvents btnCAEFacEnvelopes As System.Windows.Forms.Button
    Friend WithEvents pnlOwnerSummaryHeader As System.Windows.Forms.Panel
    Friend WithEvents pnlOwnerSummaryDetails As System.Windows.Forms.Panel
    Public WithEvents UCOwnerSummary As MUSTER.OwnerSummary
    Public WithEvents txtFacilitySIC As System.Windows.Forms.Label
    Friend WithEvents tbPageOwnerContactList As System.Windows.Forms.TabPage
    Friend WithEvents pnlOwnerContactHeader As System.Windows.Forms.Panel
    Friend WithEvents chkOwnerShowActiveOnly As System.Windows.Forms.CheckBox
    Friend WithEvents chkOwnerShowRelatedContacts As System.Windows.Forms.CheckBox
    Friend WithEvents chkOwnerShowContactsforAllModules As System.Windows.Forms.CheckBox
    Friend WithEvents lblOwnerContacts As System.Windows.Forms.Label
    Friend WithEvents pnlOwnerContactButtons As System.Windows.Forms.Panel
    Friend WithEvents btnOwnerModifyContact As System.Windows.Forms.Button
    Friend WithEvents btnOwnerDeleteContact As System.Windows.Forms.Button
    Friend WithEvents btnOwnerAssociateContact As System.Windows.Forms.Button
    Friend WithEvents btnOwnerAddSearchContact As System.Windows.Forms.Button
    Friend WithEvents pnlOwnerContactContainer As System.Windows.Forms.Panel
    Friend WithEvents ugOwnerContacts As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents tbPageFacilityContactList As System.Windows.Forms.TabPage
    Friend WithEvents pnlFacilityContactHeader As System.Windows.Forms.Panel
    Friend WithEvents chkFacilityShowActiveContactOnly As System.Windows.Forms.CheckBox
    Friend WithEvents chkFacilityShowRelatedContacts As System.Windows.Forms.CheckBox
    Friend WithEvents chkFacilityShowContactsforAllModule As System.Windows.Forms.CheckBox
    Friend WithEvents lblFacilityContacts As System.Windows.Forms.Label
    Friend WithEvents pnlFacilityContactBottom As System.Windows.Forms.Panel
    Friend WithEvents btnFacilityModifyContact As System.Windows.Forms.Button
    Friend WithEvents btnFacilityDeleteContact As System.Windows.Forms.Button
    Friend WithEvents btnFacilityAssociateContact As System.Windows.Forms.Button
    Friend WithEvents btnFacilityAddSearchContact As System.Windows.Forms.Button
    Friend WithEvents pnlFacilityContactContainer As System.Windows.Forms.Panel
    Friend WithEvents ugFacilityContacts As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents lblInViolationValue As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Public WithEvents txtDesOp As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents lblDesMa As System.Windows.Forms.Label
    Friend WithEvents txtDesMa As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(CandE))
        Me.pnlTop = New System.Windows.Forms.Panel
        Me.lblOwnerLastEditedOn = New System.Windows.Forms.Label
        Me.lblOwnerLastEditedBy = New System.Windows.Forms.Label
        Me.tbCntrlCandE = New System.Windows.Forms.TabControl
        Me.tbPageOwnerDetail = New System.Windows.Forms.TabPage
        Me.pnlOwnerBottom = New System.Windows.Forms.Panel
        Me.tbCtrlOwner = New System.Windows.Forms.TabControl
        Me.tbPageOwnerFacilities = New System.Windows.Forms.TabPage
        Me.ugFacilityList = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlOwnerFacilityBottom = New System.Windows.Forms.Panel
        Me.lblNoOfFacilitiesValue = New System.Windows.Forms.Label
        Me.lblNoOfFacilities = New System.Windows.Forms.Label
        Me.tbPageOCEs = New System.Windows.Forms.TabPage
        Me.ugOCEs = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlOwnerOCEsBottom = New System.Windows.Forms.Panel
        Me.lblNoOfOCEsValue = New System.Windows.Forms.Label
        Me.lblNoOfOCEs = New System.Windows.Forms.Label
        Me.tbPageOwnerDocuments = New System.Windows.Forms.TabPage
        Me.UCOwnerDocuments = New MUSTER.DocumentViewControl
        Me.tbPageOwnerContactList = New System.Windows.Forms.TabPage
        Me.pnlOwnerContactContainer = New System.Windows.Forms.Panel
        Me.ugOwnerContacts = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.Label1 = New System.Windows.Forms.Label
        Me.pnlOwnerContactButtons = New System.Windows.Forms.Panel
        Me.btnOwnerModifyContact = New System.Windows.Forms.Button
        Me.btnOwnerDeleteContact = New System.Windows.Forms.Button
        Me.btnOwnerAssociateContact = New System.Windows.Forms.Button
        Me.btnOwnerAddSearchContact = New System.Windows.Forms.Button
        Me.pnlOwnerContactHeader = New System.Windows.Forms.Panel
        Me.chkOwnerShowActiveOnly = New System.Windows.Forms.CheckBox
        Me.chkOwnerShowRelatedContacts = New System.Windows.Forms.CheckBox
        Me.chkOwnerShowContactsforAllModules = New System.Windows.Forms.CheckBox
        Me.lblOwnerContacts = New System.Windows.Forms.Label
        Me.pnlOwnerDetail = New System.Windows.Forms.Panel
        Me.pnlOwnerButtons = New System.Windows.Forms.Panel
        Me.btnEnforceViewEnforceHistory = New System.Windows.Forms.Button
        Me.btnOwnerFlag = New System.Windows.Forms.Button
        Me.btnOwnerComment = New System.Windows.Forms.Button
        Me.chkOwnerAgencyInterest = New System.Windows.Forms.CheckBox
        Me.lblOwnerActiveOrNot = New System.Windows.Forms.Label
        Me.LinkLblCAPSignup = New System.Windows.Forms.LinkLabel
        Me.lblCAPParticipationLevel = New System.Windows.Forms.Label
        Me.mskTxtOwnerFax = New AxMSMask.AxMaskEdBox
        Me.mskTxtOwnerPhone2 = New AxMSMask.AxMaskEdBox
        Me.mskTxtOwnerPhone = New AxMSMask.AxMaskEdBox
        Me.lblOwnerEmail = New System.Windows.Forms.Label
        Me.txtOwnerEmail = New System.Windows.Forms.TextBox
        Me.lblFax = New System.Windows.Forms.Label
        Me.txtOwnerAddress = New System.Windows.Forms.TextBox
        Me.lblOwnerAddress = New System.Windows.Forms.Label
        Me.txtOwnerName = New System.Windows.Forms.TextBox
        Me.lblOwnerName = New System.Windows.Forms.Label
        Me.lblOwnerStatus = New System.Windows.Forms.Label
        Me.lblOwnerCapParticipant = New System.Windows.Forms.Label
        Me.lblPhone2 = New System.Windows.Forms.Label
        Me.lblOwnerType = New System.Windows.Forms.Label
        Me.txtOwnerAIID = New System.Windows.Forms.TextBox
        Me.lblOwnerAIID = New System.Windows.Forms.Label
        Me.lblOwnerIDValue = New System.Windows.Forms.Label
        Me.lblOwnerID = New System.Windows.Forms.Label
        Me.lblOwnerPhone = New System.Windows.Forms.Label
        Me.cmbOwnerType = New System.Windows.Forms.ComboBox
        Me.btnCAEOwnerLabels = New System.Windows.Forms.Button
        Me.btnCAEOwnerEnvelopes = New System.Windows.Forms.Button
        Me.tbPageFacilityDetail = New System.Windows.Forms.TabPage
        Me.pnlFacilityBottom = New System.Windows.Forms.Panel
        Me.tbCntrlFacility = New System.Windows.Forms.TabControl
        Me.tbPageFCEs = New System.Windows.Forms.TabPage
        Me.ugFCEs = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlFacilityLustButton = New System.Windows.Forms.Panel
        Me.lblTotalNoOfLUSTEventsValue = New System.Windows.Forms.Label
        Me.lblTotalNoOfLUSTEvents = New System.Windows.Forms.Label
        Me.tbPageFacilityDocuments = New System.Windows.Forms.TabPage
        Me.UCFacilityDocuments = New MUSTER.DocumentViewControl
        Me.tbPageFacilityContactList = New System.Windows.Forms.TabPage
        Me.pnlFacilityContactContainer = New System.Windows.Forms.Panel
        Me.ugFacilityContacts = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.Label2 = New System.Windows.Forms.Label
        Me.pnlFacilityContactBottom = New System.Windows.Forms.Panel
        Me.btnFacilityModifyContact = New System.Windows.Forms.Button
        Me.btnFacilityDeleteContact = New System.Windows.Forms.Button
        Me.btnFacilityAssociateContact = New System.Windows.Forms.Button
        Me.btnFacilityAddSearchContact = New System.Windows.Forms.Button
        Me.pnlFacilityContactHeader = New System.Windows.Forms.Panel
        Me.chkFacilityShowActiveContactOnly = New System.Windows.Forms.CheckBox
        Me.chkFacilityShowRelatedContacts = New System.Windows.Forms.CheckBox
        Me.chkFacilityShowContactsforAllModule = New System.Windows.Forms.CheckBox
        Me.lblFacilityContacts = New System.Windows.Forms.Label
        Me.pnl_FacilityDetail = New System.Windows.Forms.Panel
        Me.txtDesMa = New System.Windows.Forms.TextBox
        Me.lblDesMa = New System.Windows.Forms.Label
        Me.txtDesOp = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.btnCAEFacLabels = New System.Windows.Forms.Button
        Me.btnCAEFacEnvelopes = New System.Windows.Forms.Button
        Me.btnFacComments = New System.Windows.Forms.Button
        Me.btnFacFlags = New System.Windows.Forms.Button
        Me.btnFacilityCancel = New System.Windows.Forms.Button
        Me.btnFacilitySave = New System.Windows.Forms.Button
        Me.dtPickUpcomingInstallDateValue = New System.Windows.Forms.DateTimePicker
        Me.lblUpcomingInstallDate = New System.Windows.Forms.Label
        Me.chkUpcomingInstall = New System.Windows.Forms.CheckBox
        Me.lnkLblNextFac = New System.Windows.Forms.LinkLabel
        Me.lblCAPStatusValue = New System.Windows.Forms.Label
        Me.lblCAPStatus = New System.Windows.Forms.Label
        Me.txtFuelBrand = New System.Windows.Forms.TextBox
        Me.ll = New System.Windows.Forms.Label
        Me.dtFacilityPowerOff = New System.Windows.Forms.DateTimePicker
        Me.lnkLblPrevFacility = New System.Windows.Forms.LinkLabel
        Me.lblLUSTSite = New System.Windows.Forms.Label
        Me.chkLUSTSite = New System.Windows.Forms.CheckBox
        Me.lblPowerOff = New System.Windows.Forms.Label
        Me.lblCAPCandidate = New System.Windows.Forms.Label
        Me.chkCAPCandidate = New System.Windows.Forms.CheckBox
        Me.lblFacilityLocationType = New System.Windows.Forms.Label
        Me.cmbFacilityLocationType = New System.Windows.Forms.ComboBox
        Me.lblFacilityMethod = New System.Windows.Forms.Label
        Me.cmbFacilityMethod = New System.Windows.Forms.ComboBox
        Me.lblFacilityDatum = New System.Windows.Forms.Label
        Me.cmbFacilityDatum = New System.Windows.Forms.ComboBox
        Me.cmbFacilityType = New System.Windows.Forms.ComboBox
        Me.txtFacilityLatSec = New System.Windows.Forms.TextBox
        Me.txtFacilityLongSec = New System.Windows.Forms.TextBox
        Me.txtFacilityLatMin = New System.Windows.Forms.TextBox
        Me.txtFacilityLongMin = New System.Windows.Forms.TextBox
        Me.lblFacilityLongMin = New System.Windows.Forms.Label
        Me.lblFacilityLongSec = New System.Windows.Forms.Label
        Me.lblFacilityLatMin = New System.Windows.Forms.Label
        Me.lblFacilityLatSec = New System.Windows.Forms.Label
        Me.lblFacilityLongDegree = New System.Windows.Forms.Label
        Me.mskTxtFacilityFax = New AxMSMask.AxMaskEdBox
        Me.mskTxtFacilityPhone = New AxMSMask.AxMaskEdBox
        Me.txtFacilityAddress = New System.Windows.Forms.TextBox
        Me.lblFacilitySIC = New System.Windows.Forms.Label
        Me.lblFacilityFax = New System.Windows.Forms.Label
        Me.dtPickFacilityRecvd = New System.Windows.Forms.DateTimePicker
        Me.lblDateReceived = New System.Windows.Forms.Label
        Me.chkSignatureofNF = New System.Windows.Forms.CheckBox
        Me.lblFacilitySigOnNF = New System.Windows.Forms.Label
        Me.lblFacilityFuelBrand = New System.Windows.Forms.Label
        Me.lblFacilityStatusValue = New System.Windows.Forms.Label
        Me.lblFacilityStatus = New System.Windows.Forms.Label
        Me.txtFacilityLongDegree = New System.Windows.Forms.TextBox
        Me.txtFacilityLatDegree = New System.Windows.Forms.TextBox
        Me.lblFacilityLongitude = New System.Windows.Forms.Label
        Me.lblFacilityLatitude = New System.Windows.Forms.Label
        Me.lblFacilityType = New System.Windows.Forms.Label
        Me.txtFacilityAIID = New System.Windows.Forms.TextBox
        Me.lblfacilityAIID = New System.Windows.Forms.Label
        Me.lblFacilityIDValue = New System.Windows.Forms.Label
        Me.lblFacilityID = New System.Windows.Forms.Label
        Me.lblFacilityPhone = New System.Windows.Forms.Label
        Me.txtFacilityName = New System.Windows.Forms.TextBox
        Me.lblFacilityAddress = New System.Windows.Forms.Label
        Me.lblFacilityName = New System.Windows.Forms.Label
        Me.lblFacilityLatDegree = New System.Windows.Forms.Label
        Me.btnInspHistory = New System.Windows.Forms.Button
        Me.txtFacilitySIC = New System.Windows.Forms.Label
        Me.lblInViolationValue = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.tbPageOwnerCitations = New System.Windows.Forms.TabPage
        Me.ugOCE2 = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlOwnerCitationsHead = New System.Windows.Forms.Panel
        Me.btnEnforceViewEnforceHistory1 = New System.Windows.Forms.Button
        Me.btnExpandOwnerCitationAll = New System.Windows.Forms.Button
        Me.tbPageFacilityCitations = New System.Windows.Forms.TabPage
        Me.ugFCE2 = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlFacilityCitationHead = New System.Windows.Forms.Panel
        Me.btnExpandFacilityCitationAll = New System.Windows.Forms.Button
        Me.btnInspHistory1 = New System.Windows.Forms.Button
        Me.tbPageSummary = New System.Windows.Forms.TabPage
        Me.pnlOwnerSummaryDetails = New System.Windows.Forms.Panel
        Me.UCOwnerSummary = New MUSTER.OwnerSummary
        Me.pnlOwnerSummaryHeader = New System.Windows.Forms.Panel
        Me.Panel12 = New System.Windows.Forms.Panel
        Me.pnlTop.SuspendLayout()
        Me.tbCntrlCandE.SuspendLayout()
        Me.tbPageOwnerDetail.SuspendLayout()
        Me.pnlOwnerBottom.SuspendLayout()
        Me.tbCtrlOwner.SuspendLayout()
        Me.tbPageOwnerFacilities.SuspendLayout()
        CType(Me.ugFacilityList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlOwnerFacilityBottom.SuspendLayout()
        Me.tbPageOCEs.SuspendLayout()
        CType(Me.ugOCEs, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlOwnerOCEsBottom.SuspendLayout()
        Me.tbPageOwnerDocuments.SuspendLayout()
        Me.tbPageOwnerContactList.SuspendLayout()
        Me.pnlOwnerContactContainer.SuspendLayout()
        CType(Me.ugOwnerContacts, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlOwnerContactButtons.SuspendLayout()
        Me.pnlOwnerContactHeader.SuspendLayout()
        Me.pnlOwnerDetail.SuspendLayout()
        Me.pnlOwnerButtons.SuspendLayout()
        CType(Me.mskTxtOwnerFax, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtOwnerPhone2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtOwnerPhone, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbPageFacilityDetail.SuspendLayout()
        Me.pnlFacilityBottom.SuspendLayout()
        Me.tbCntrlFacility.SuspendLayout()
        Me.tbPageFCEs.SuspendLayout()
        CType(Me.ugFCEs, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlFacilityLustButton.SuspendLayout()
        Me.tbPageFacilityDocuments.SuspendLayout()
        Me.tbPageFacilityContactList.SuspendLayout()
        Me.pnlFacilityContactContainer.SuspendLayout()
        CType(Me.ugFacilityContacts, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlFacilityContactBottom.SuspendLayout()
        Me.pnlFacilityContactHeader.SuspendLayout()
        Me.pnl_FacilityDetail.SuspendLayout()
        CType(Me.mskTxtFacilityFax, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtFacilityPhone, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbPageOwnerCitations.SuspendLayout()
        CType(Me.ugOCE2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlOwnerCitationsHead.SuspendLayout()
        Me.tbPageFacilityCitations.SuspendLayout()
        CType(Me.ugFCE2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlFacilityCitationHead.SuspendLayout()
        Me.tbPageSummary.SuspendLayout()
        Me.pnlOwnerSummaryDetails.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlTop
        '
        Me.pnlTop.BackColor = System.Drawing.SystemColors.Control
        Me.pnlTop.Controls.Add(Me.lblOwnerLastEditedOn)
        Me.pnlTop.Controls.Add(Me.lblOwnerLastEditedBy)
        Me.pnlTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlTop.Name = "pnlTop"
        Me.pnlTop.Size = New System.Drawing.Size(1028, 24)
        Me.pnlTop.TabIndex = 2
        '
        'lblOwnerLastEditedOn
        '
        Me.lblOwnerLastEditedOn.Location = New System.Drawing.Point(704, 5)
        Me.lblOwnerLastEditedOn.Name = "lblOwnerLastEditedOn"
        Me.lblOwnerLastEditedOn.Size = New System.Drawing.Size(168, 16)
        Me.lblOwnerLastEditedOn.TabIndex = 1014
        Me.lblOwnerLastEditedOn.Text = "Last Edited On :"
        '
        'lblOwnerLastEditedBy
        '
        Me.lblOwnerLastEditedBy.Location = New System.Drawing.Point(488, 5)
        Me.lblOwnerLastEditedBy.Name = "lblOwnerLastEditedBy"
        Me.lblOwnerLastEditedBy.Size = New System.Drawing.Size(208, 16)
        Me.lblOwnerLastEditedBy.TabIndex = 1013
        Me.lblOwnerLastEditedBy.Text = "Last Edited By :"
        '
        'tbCntrlCandE
        '
        Me.tbCntrlCandE.Controls.Add(Me.tbPageOwnerDetail)
        Me.tbCntrlCandE.Controls.Add(Me.tbPageFacilityDetail)
        Me.tbCntrlCandE.Controls.Add(Me.tbPageOwnerCitations)
        Me.tbCntrlCandE.Controls.Add(Me.tbPageFacilityCitations)
        Me.tbCntrlCandE.Controls.Add(Me.tbPageSummary)
        Me.tbCntrlCandE.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbCntrlCandE.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbCntrlCandE.ItemSize = New System.Drawing.Size(64, 18)
        Me.tbCntrlCandE.Location = New System.Drawing.Point(0, 24)
        Me.tbCntrlCandE.Multiline = True
        Me.tbCntrlCandE.Name = "tbCntrlCandE"
        Me.tbCntrlCandE.SelectedIndex = 0
        Me.tbCntrlCandE.ShowToolTips = True
        Me.tbCntrlCandE.Size = New System.Drawing.Size(1028, 670)
        Me.tbCntrlCandE.TabIndex = 3
        '
        'tbPageOwnerDetail
        '
        Me.tbPageOwnerDetail.Controls.Add(Me.pnlOwnerBottom)
        Me.tbPageOwnerDetail.Controls.Add(Me.pnlOwnerDetail)
        Me.tbPageOwnerDetail.Location = New System.Drawing.Point(4, 22)
        Me.tbPageOwnerDetail.Name = "tbPageOwnerDetail"
        Me.tbPageOwnerDetail.Size = New System.Drawing.Size(1020, 644)
        Me.tbPageOwnerDetail.TabIndex = 7
        Me.tbPageOwnerDetail.Text = "Owner Details"
        '
        'pnlOwnerBottom
        '
        Me.pnlOwnerBottom.AutoScroll = True
        Me.pnlOwnerBottom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlOwnerBottom.Controls.Add(Me.tbCtrlOwner)
        Me.pnlOwnerBottom.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlOwnerBottom.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlOwnerBottom.Location = New System.Drawing.Point(0, 200)
        Me.pnlOwnerBottom.Name = "pnlOwnerBottom"
        Me.pnlOwnerBottom.Size = New System.Drawing.Size(1020, 444)
        Me.pnlOwnerBottom.TabIndex = 44
        '
        'tbCtrlOwner
        '
        Me.tbCtrlOwner.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
        Me.tbCtrlOwner.Controls.Add(Me.tbPageOwnerFacilities)
        Me.tbCtrlOwner.Controls.Add(Me.tbPageOCEs)
        Me.tbCtrlOwner.Controls.Add(Me.tbPageOwnerDocuments)
        Me.tbCtrlOwner.Controls.Add(Me.tbPageOwnerContactList)
        Me.tbCtrlOwner.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbCtrlOwner.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbCtrlOwner.ItemSize = New System.Drawing.Size(60, 23)
        Me.tbCtrlOwner.Location = New System.Drawing.Point(0, 0)
        Me.tbCtrlOwner.Name = "tbCtrlOwner"
        Me.tbCtrlOwner.SelectedIndex = 0
        Me.tbCtrlOwner.Size = New System.Drawing.Size(1018, 442)
        Me.tbCtrlOwner.TabIndex = 9
        '
        'tbPageOwnerFacilities
        '
        Me.tbPageOwnerFacilities.AutoScroll = True
        Me.tbPageOwnerFacilities.BackColor = System.Drawing.SystemColors.Control
        Me.tbPageOwnerFacilities.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.tbPageOwnerFacilities.Controls.Add(Me.ugFacilityList)
        Me.tbPageOwnerFacilities.Controls.Add(Me.pnlOwnerFacilityBottom)
        Me.tbPageOwnerFacilities.Location = New System.Drawing.Point(4, 27)
        Me.tbPageOwnerFacilities.Name = "tbPageOwnerFacilities"
        Me.tbPageOwnerFacilities.Size = New System.Drawing.Size(1010, 411)
        Me.tbPageOwnerFacilities.TabIndex = 0
        Me.tbPageOwnerFacilities.Text = "Facilities"
        '
        'ugFacilityList
        '
        Me.ugFacilityList.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugFacilityList.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle
        Me.ugFacilityList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugFacilityList.Location = New System.Drawing.Point(0, 0)
        Me.ugFacilityList.Name = "ugFacilityList"
        Me.ugFacilityList.Size = New System.Drawing.Size(1006, 383)
        Me.ugFacilityList.TabIndex = 20
        '
        'pnlOwnerFacilityBottom
        '
        Me.pnlOwnerFacilityBottom.Controls.Add(Me.lblNoOfFacilitiesValue)
        Me.pnlOwnerFacilityBottom.Controls.Add(Me.lblNoOfFacilities)
        Me.pnlOwnerFacilityBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlOwnerFacilityBottom.Location = New System.Drawing.Point(0, 383)
        Me.pnlOwnerFacilityBottom.Name = "pnlOwnerFacilityBottom"
        Me.pnlOwnerFacilityBottom.Size = New System.Drawing.Size(1006, 24)
        Me.pnlOwnerFacilityBottom.TabIndex = 19
        '
        'lblNoOfFacilitiesValue
        '
        Me.lblNoOfFacilitiesValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNoOfFacilitiesValue.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblNoOfFacilitiesValue.Location = New System.Drawing.Point(100, 0)
        Me.lblNoOfFacilitiesValue.Name = "lblNoOfFacilitiesValue"
        Me.lblNoOfFacilitiesValue.Size = New System.Drawing.Size(56, 24)
        Me.lblNoOfFacilitiesValue.TabIndex = 7
        Me.lblNoOfFacilitiesValue.Text = "0"
        '
        'lblNoOfFacilities
        '
        Me.lblNoOfFacilities.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNoOfFacilities.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblNoOfFacilities.Location = New System.Drawing.Point(0, 0)
        Me.lblNoOfFacilities.Name = "lblNoOfFacilities"
        Me.lblNoOfFacilities.Size = New System.Drawing.Size(100, 24)
        Me.lblNoOfFacilities.TabIndex = 6
        Me.lblNoOfFacilities.Text = "No of Facilities:"
        '
        'tbPageOCEs
        '
        Me.tbPageOCEs.AutoScroll = True
        Me.tbPageOCEs.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.tbPageOCEs.Controls.Add(Me.ugOCEs)
        Me.tbPageOCEs.Controls.Add(Me.pnlOwnerOCEsBottom)
        Me.tbPageOCEs.Location = New System.Drawing.Point(4, 27)
        Me.tbPageOCEs.Name = "tbPageOCEs"
        Me.tbPageOCEs.Size = New System.Drawing.Size(1010, 411)
        Me.tbPageOCEs.TabIndex = 1
        Me.tbPageOCEs.Text = "OCEs"
        '
        'ugOCEs
        '
        Me.ugOCEs.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugOCEs.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle
        Me.ugOCEs.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugOCEs.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugOCEs.Location = New System.Drawing.Point(0, 0)
        Me.ugOCEs.Name = "ugOCEs"
        Me.ugOCEs.Size = New System.Drawing.Size(1006, 383)
        Me.ugOCEs.TabIndex = 21
        '
        'pnlOwnerOCEsBottom
        '
        Me.pnlOwnerOCEsBottom.Controls.Add(Me.lblNoOfOCEsValue)
        Me.pnlOwnerOCEsBottom.Controls.Add(Me.lblNoOfOCEs)
        Me.pnlOwnerOCEsBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlOwnerOCEsBottom.Location = New System.Drawing.Point(0, 383)
        Me.pnlOwnerOCEsBottom.Name = "pnlOwnerOCEsBottom"
        Me.pnlOwnerOCEsBottom.Size = New System.Drawing.Size(1006, 24)
        Me.pnlOwnerOCEsBottom.TabIndex = 20
        '
        'lblNoOfOCEsValue
        '
        Me.lblNoOfOCEsValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNoOfOCEsValue.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblNoOfOCEsValue.Location = New System.Drawing.Point(80, 0)
        Me.lblNoOfOCEsValue.Name = "lblNoOfOCEsValue"
        Me.lblNoOfOCEsValue.Size = New System.Drawing.Size(56, 24)
        Me.lblNoOfOCEsValue.TabIndex = 7
        Me.lblNoOfOCEsValue.Text = "0"
        '
        'lblNoOfOCEs
        '
        Me.lblNoOfOCEs.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNoOfOCEs.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblNoOfOCEs.Location = New System.Drawing.Point(0, 0)
        Me.lblNoOfOCEs.Name = "lblNoOfOCEs"
        Me.lblNoOfOCEs.Size = New System.Drawing.Size(80, 24)
        Me.lblNoOfOCEs.TabIndex = 6
        Me.lblNoOfOCEs.Text = "No of OCEs:"
        '
        'tbPageOwnerDocuments
        '
        Me.tbPageOwnerDocuments.AutoScroll = True
        Me.tbPageOwnerDocuments.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.tbPageOwnerDocuments.Controls.Add(Me.UCOwnerDocuments)
        Me.tbPageOwnerDocuments.Location = New System.Drawing.Point(4, 27)
        Me.tbPageOwnerDocuments.Name = "tbPageOwnerDocuments"
        Me.tbPageOwnerDocuments.Size = New System.Drawing.Size(1010, 411)
        Me.tbPageOwnerDocuments.TabIndex = 2
        Me.tbPageOwnerDocuments.Text = "Documents"
        '
        'UCOwnerDocuments
        '
        Me.UCOwnerDocuments.AutoScroll = True
        Me.UCOwnerDocuments.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UCOwnerDocuments.Location = New System.Drawing.Point(0, 0)
        Me.UCOwnerDocuments.Name = "UCOwnerDocuments"
        Me.UCOwnerDocuments.Size = New System.Drawing.Size(1006, 407)
        Me.UCOwnerDocuments.TabIndex = 3
        '
        'tbPageOwnerContactList
        '
        Me.tbPageOwnerContactList.Controls.Add(Me.pnlOwnerContactContainer)
        Me.tbPageOwnerContactList.Controls.Add(Me.pnlOwnerContactButtons)
        Me.tbPageOwnerContactList.Controls.Add(Me.pnlOwnerContactHeader)
        Me.tbPageOwnerContactList.Location = New System.Drawing.Point(4, 27)
        Me.tbPageOwnerContactList.Name = "tbPageOwnerContactList"
        Me.tbPageOwnerContactList.Size = New System.Drawing.Size(1010, 411)
        Me.tbPageOwnerContactList.TabIndex = 3
        Me.tbPageOwnerContactList.Text = "Contacts"
        '
        'pnlOwnerContactContainer
        '
        Me.pnlOwnerContactContainer.Controls.Add(Me.ugOwnerContacts)
        Me.pnlOwnerContactContainer.Controls.Add(Me.Label1)
        Me.pnlOwnerContactContainer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlOwnerContactContainer.Location = New System.Drawing.Point(0, 25)
        Me.pnlOwnerContactContainer.Name = "pnlOwnerContactContainer"
        Me.pnlOwnerContactContainer.Size = New System.Drawing.Size(1010, 356)
        Me.pnlOwnerContactContainer.TabIndex = 4
        '
        'ugOwnerContacts
        '
        Me.ugOwnerContacts.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugOwnerContacts.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugOwnerContacts.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        Me.ugOwnerContacts.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugOwnerContacts.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugOwnerContacts.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugOwnerContacts.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugOwnerContacts.Location = New System.Drawing.Point(0, 0)
        Me.ugOwnerContacts.Name = "ugOwnerContacts"
        Me.ugOwnerContacts.Size = New System.Drawing.Size(1010, 356)
        Me.ugOwnerContacts.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(792, 56)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(7, 23)
        Me.Label1.TabIndex = 2
        '
        'pnlOwnerContactButtons
        '
        Me.pnlOwnerContactButtons.Controls.Add(Me.btnOwnerModifyContact)
        Me.pnlOwnerContactButtons.Controls.Add(Me.btnOwnerDeleteContact)
        Me.pnlOwnerContactButtons.Controls.Add(Me.btnOwnerAssociateContact)
        Me.pnlOwnerContactButtons.Controls.Add(Me.btnOwnerAddSearchContact)
        Me.pnlOwnerContactButtons.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlOwnerContactButtons.DockPadding.All = 3
        Me.pnlOwnerContactButtons.Location = New System.Drawing.Point(0, 381)
        Me.pnlOwnerContactButtons.Name = "pnlOwnerContactButtons"
        Me.pnlOwnerContactButtons.Size = New System.Drawing.Size(1010, 30)
        Me.pnlOwnerContactButtons.TabIndex = 3
        '
        'btnOwnerModifyContact
        '
        Me.btnOwnerModifyContact.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOwnerModifyContact.Location = New System.Drawing.Point(240, 5)
        Me.btnOwnerModifyContact.Name = "btnOwnerModifyContact"
        Me.btnOwnerModifyContact.Size = New System.Drawing.Size(235, 23)
        Me.btnOwnerModifyContact.TabIndex = 1
        Me.btnOwnerModifyContact.Text = "Modify Contact"
        '
        'btnOwnerDeleteContact
        '
        Me.btnOwnerDeleteContact.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOwnerDeleteContact.Location = New System.Drawing.Point(472, 5)
        Me.btnOwnerDeleteContact.Name = "btnOwnerDeleteContact"
        Me.btnOwnerDeleteContact.Size = New System.Drawing.Size(235, 23)
        Me.btnOwnerDeleteContact.TabIndex = 2
        Me.btnOwnerDeleteContact.Text = "Disassociate Contact"
        '
        'btnOwnerAssociateContact
        '
        Me.btnOwnerAssociateContact.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOwnerAssociateContact.Location = New System.Drawing.Point(704, 5)
        Me.btnOwnerAssociateContact.Name = "btnOwnerAssociateContact"
        Me.btnOwnerAssociateContact.Size = New System.Drawing.Size(235, 23)
        Me.btnOwnerAssociateContact.TabIndex = 3
        Me.btnOwnerAssociateContact.Text = "Associate Contact from Different Module"
        '
        'btnOwnerAddSearchContact
        '
        Me.btnOwnerAddSearchContact.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOwnerAddSearchContact.Location = New System.Drawing.Point(8, 5)
        Me.btnOwnerAddSearchContact.Name = "btnOwnerAddSearchContact"
        Me.btnOwnerAddSearchContact.Size = New System.Drawing.Size(235, 23)
        Me.btnOwnerAddSearchContact.TabIndex = 0
        Me.btnOwnerAddSearchContact.Text = "Add/Search Contact to Associate"
        '
        'pnlOwnerContactHeader
        '
        Me.pnlOwnerContactHeader.Controls.Add(Me.chkOwnerShowActiveOnly)
        Me.pnlOwnerContactHeader.Controls.Add(Me.chkOwnerShowRelatedContacts)
        Me.pnlOwnerContactHeader.Controls.Add(Me.chkOwnerShowContactsforAllModules)
        Me.pnlOwnerContactHeader.Controls.Add(Me.lblOwnerContacts)
        Me.pnlOwnerContactHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlOwnerContactHeader.DockPadding.All = 3
        Me.pnlOwnerContactHeader.Location = New System.Drawing.Point(0, 0)
        Me.pnlOwnerContactHeader.Name = "pnlOwnerContactHeader"
        Me.pnlOwnerContactHeader.Size = New System.Drawing.Size(1010, 25)
        Me.pnlOwnerContactHeader.TabIndex = 1
        '
        'chkOwnerShowActiveOnly
        '
        Me.chkOwnerShowActiveOnly.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkOwnerShowActiveOnly.Location = New System.Drawing.Point(635, 5)
        Me.chkOwnerShowActiveOnly.Name = "chkOwnerShowActiveOnly"
        Me.chkOwnerShowActiveOnly.Size = New System.Drawing.Size(144, 16)
        Me.chkOwnerShowActiveOnly.TabIndex = 2
        Me.chkOwnerShowActiveOnly.Tag = "646"
        Me.chkOwnerShowActiveOnly.Text = "Show Active Only"
        '
        'chkOwnerShowRelatedContacts
        '
        Me.chkOwnerShowRelatedContacts.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkOwnerShowRelatedContacts.Location = New System.Drawing.Point(467, 5)
        Me.chkOwnerShowRelatedContacts.Name = "chkOwnerShowRelatedContacts"
        Me.chkOwnerShowRelatedContacts.Size = New System.Drawing.Size(160, 16)
        Me.chkOwnerShowRelatedContacts.TabIndex = 1
        Me.chkOwnerShowRelatedContacts.Tag = "645"
        Me.chkOwnerShowRelatedContacts.Text = "Show Related Contacts"
        '
        'chkOwnerShowContactsforAllModules
        '
        Me.chkOwnerShowContactsforAllModules.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkOwnerShowContactsforAllModules.Location = New System.Drawing.Point(251, 5)
        Me.chkOwnerShowContactsforAllModules.Name = "chkOwnerShowContactsforAllModules"
        Me.chkOwnerShowContactsforAllModules.Size = New System.Drawing.Size(200, 16)
        Me.chkOwnerShowContactsforAllModules.TabIndex = 0
        Me.chkOwnerShowContactsforAllModules.Tag = "644"
        Me.chkOwnerShowContactsforAllModules.Text = "Show Contacts for All Modules"
        '
        'lblOwnerContacts
        '
        Me.lblOwnerContacts.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerContacts.Location = New System.Drawing.Point(8, 5)
        Me.lblOwnerContacts.Name = "lblOwnerContacts"
        Me.lblOwnerContacts.Size = New System.Drawing.Size(104, 16)
        Me.lblOwnerContacts.TabIndex = 139
        Me.lblOwnerContacts.Text = "Owner Contacts"
        '
        'pnlOwnerDetail
        '
        Me.pnlOwnerDetail.BackColor = System.Drawing.SystemColors.Control
        Me.pnlOwnerDetail.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlOwnerDetail.Controls.Add(Me.pnlOwnerButtons)
        Me.pnlOwnerDetail.Controls.Add(Me.chkOwnerAgencyInterest)
        Me.pnlOwnerDetail.Controls.Add(Me.lblOwnerActiveOrNot)
        Me.pnlOwnerDetail.Controls.Add(Me.LinkLblCAPSignup)
        Me.pnlOwnerDetail.Controls.Add(Me.lblCAPParticipationLevel)
        Me.pnlOwnerDetail.Controls.Add(Me.mskTxtOwnerFax)
        Me.pnlOwnerDetail.Controls.Add(Me.mskTxtOwnerPhone2)
        Me.pnlOwnerDetail.Controls.Add(Me.mskTxtOwnerPhone)
        Me.pnlOwnerDetail.Controls.Add(Me.lblOwnerEmail)
        Me.pnlOwnerDetail.Controls.Add(Me.txtOwnerEmail)
        Me.pnlOwnerDetail.Controls.Add(Me.lblFax)
        Me.pnlOwnerDetail.Controls.Add(Me.txtOwnerAddress)
        Me.pnlOwnerDetail.Controls.Add(Me.lblOwnerAddress)
        Me.pnlOwnerDetail.Controls.Add(Me.txtOwnerName)
        Me.pnlOwnerDetail.Controls.Add(Me.lblOwnerName)
        Me.pnlOwnerDetail.Controls.Add(Me.lblOwnerStatus)
        Me.pnlOwnerDetail.Controls.Add(Me.lblOwnerCapParticipant)
        Me.pnlOwnerDetail.Controls.Add(Me.lblPhone2)
        Me.pnlOwnerDetail.Controls.Add(Me.lblOwnerType)
        Me.pnlOwnerDetail.Controls.Add(Me.txtOwnerAIID)
        Me.pnlOwnerDetail.Controls.Add(Me.lblOwnerAIID)
        Me.pnlOwnerDetail.Controls.Add(Me.lblOwnerIDValue)
        Me.pnlOwnerDetail.Controls.Add(Me.lblOwnerID)
        Me.pnlOwnerDetail.Controls.Add(Me.lblOwnerPhone)
        Me.pnlOwnerDetail.Controls.Add(Me.cmbOwnerType)
        Me.pnlOwnerDetail.Controls.Add(Me.btnCAEOwnerLabels)
        Me.pnlOwnerDetail.Controls.Add(Me.btnCAEOwnerEnvelopes)
        Me.pnlOwnerDetail.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlOwnerDetail.Location = New System.Drawing.Point(0, 0)
        Me.pnlOwnerDetail.Name = "pnlOwnerDetail"
        Me.pnlOwnerDetail.Size = New System.Drawing.Size(1020, 200)
        Me.pnlOwnerDetail.TabIndex = 0
        '
        'pnlOwnerButtons
        '
        Me.pnlOwnerButtons.Controls.Add(Me.btnEnforceViewEnforceHistory)
        Me.pnlOwnerButtons.Controls.Add(Me.btnOwnerFlag)
        Me.pnlOwnerButtons.Controls.Add(Me.btnOwnerComment)
        Me.pnlOwnerButtons.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.pnlOwnerButtons.Location = New System.Drawing.Point(376, 134)
        Me.pnlOwnerButtons.Name = "pnlOwnerButtons"
        Me.pnlOwnerButtons.Size = New System.Drawing.Size(320, 40)
        Me.pnlOwnerButtons.TabIndex = 1011
        '
        'btnEnforceViewEnforceHistory
        '
        Me.btnEnforceViewEnforceHistory.Location = New System.Drawing.Point(184, 4)
        Me.btnEnforceViewEnforceHistory.Name = "btnEnforceViewEnforceHistory"
        Me.btnEnforceViewEnforceHistory.Size = New System.Drawing.Size(132, 32)
        Me.btnEnforceViewEnforceHistory.TabIndex = 49
        Me.btnEnforceViewEnforceHistory.Text = "View Enforcement History"
        '
        'btnOwnerFlag
        '
        Me.btnOwnerFlag.Location = New System.Drawing.Point(8, 7)
        Me.btnOwnerFlag.Name = "btnOwnerFlag"
        Me.btnOwnerFlag.TabIndex = 48
        Me.btnOwnerFlag.Text = "Flags"
        '
        'btnOwnerComment
        '
        Me.btnOwnerComment.Location = New System.Drawing.Point(96, 7)
        Me.btnOwnerComment.Name = "btnOwnerComment"
        Me.btnOwnerComment.Size = New System.Drawing.Size(80, 23)
        Me.btnOwnerComment.TabIndex = 47
        Me.btnOwnerComment.Text = "Comments"
        '
        'chkOwnerAgencyInterest
        '
        Me.chkOwnerAgencyInterest.Enabled = False
        Me.chkOwnerAgencyInterest.Location = New System.Drawing.Point(544, 25)
        Me.chkOwnerAgencyInterest.Name = "chkOwnerAgencyInterest"
        Me.chkOwnerAgencyInterest.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkOwnerAgencyInterest.Size = New System.Drawing.Size(112, 24)
        Me.chkOwnerAgencyInterest.TabIndex = 7
        Me.chkOwnerAgencyInterest.Text = "Agency Interest   "
        Me.chkOwnerAgencyInterest.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblOwnerActiveOrNot
        '
        Me.lblOwnerActiveOrNot.BackColor = System.Drawing.SystemColors.Control
        Me.lblOwnerActiveOrNot.Enabled = False
        Me.lblOwnerActiveOrNot.Location = New System.Drawing.Point(424, 8)
        Me.lblOwnerActiveOrNot.Name = "lblOwnerActiveOrNot"
        Me.lblOwnerActiveOrNot.Size = New System.Drawing.Size(112, 16)
        Me.lblOwnerActiveOrNot.TabIndex = 1006
        '
        'LinkLblCAPSignup
        '
        Me.LinkLblCAPSignup.Enabled = False
        Me.LinkLblCAPSignup.Location = New System.Drawing.Point(544, 74)
        Me.LinkLblCAPSignup.Name = "LinkLblCAPSignup"
        Me.LinkLblCAPSignup.Size = New System.Drawing.Size(152, 16)
        Me.LinkLblCAPSignup.TabIndex = 1005
        Me.LinkLblCAPSignup.TabStop = True
        Me.LinkLblCAPSignup.Text = "CAP Signup/Maintenance"
        '
        'lblCAPParticipationLevel
        '
        Me.lblCAPParticipationLevel.Location = New System.Drawing.Point(672, 48)
        Me.lblCAPParticipationLevel.Name = "lblCAPParticipationLevel"
        Me.lblCAPParticipationLevel.Size = New System.Drawing.Size(264, 20)
        Me.lblCAPParticipationLevel.TabIndex = 1004
        Me.lblCAPParticipationLevel.Text = "NONE - 0/0 (Compliant/Candidate)"
        Me.lblCAPParticipationLevel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'mskTxtOwnerFax
        '
        Me.mskTxtOwnerFax.ContainingControl = Me
        Me.mskTxtOwnerFax.Location = New System.Drawing.Point(424, 96)
        Me.mskTxtOwnerFax.Name = "mskTxtOwnerFax"
        Me.mskTxtOwnerFax.OcxState = CType(resources.GetObject("mskTxtOwnerFax.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtOwnerFax.Size = New System.Drawing.Size(96, 23)
        Me.mskTxtOwnerFax.TabIndex = 6
        '
        'mskTxtOwnerPhone2
        '
        Me.mskTxtOwnerPhone2.ContainingControl = Me
        Me.mskTxtOwnerPhone2.Location = New System.Drawing.Point(424, 72)
        Me.mskTxtOwnerPhone2.Name = "mskTxtOwnerPhone2"
        Me.mskTxtOwnerPhone2.OcxState = CType(resources.GetObject("mskTxtOwnerPhone2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtOwnerPhone2.Size = New System.Drawing.Size(96, 23)
        Me.mskTxtOwnerPhone2.TabIndex = 5
        '
        'mskTxtOwnerPhone
        '
        Me.mskTxtOwnerPhone.ContainingControl = Me
        Me.mskTxtOwnerPhone.Location = New System.Drawing.Point(424, 48)
        Me.mskTxtOwnerPhone.Name = "mskTxtOwnerPhone"
        Me.mskTxtOwnerPhone.OcxState = CType(resources.GetObject("mskTxtOwnerPhone.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtOwnerPhone.Size = New System.Drawing.Size(96, 23)
        Me.mskTxtOwnerPhone.TabIndex = 4
        '
        'lblOwnerEmail
        '
        Me.lblOwnerEmail.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerEmail.Location = New System.Drawing.Point(544, 99)
        Me.lblOwnerEmail.Name = "lblOwnerEmail"
        Me.lblOwnerEmail.Size = New System.Drawing.Size(40, 23)
        Me.lblOwnerEmail.TabIndex = 11
        Me.lblOwnerEmail.Text = "Email"
        '
        'txtOwnerEmail
        '
        Me.txtOwnerEmail.AcceptsTab = True
        Me.txtOwnerEmail.AutoSize = False
        Me.txtOwnerEmail.Enabled = False
        Me.txtOwnerEmail.Location = New System.Drawing.Point(590, 96)
        Me.txtOwnerEmail.Name = "txtOwnerEmail"
        Me.txtOwnerEmail.Size = New System.Drawing.Size(200, 21)
        Me.txtOwnerEmail.TabIndex = 8
        Me.txtOwnerEmail.Text = ""
        Me.txtOwnerEmail.WordWrap = False
        '
        'lblFax
        '
        Me.lblFax.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFax.Location = New System.Drawing.Point(346, 96)
        Me.lblFax.Name = "lblFax"
        Me.lblFax.Size = New System.Drawing.Size(48, 23)
        Me.lblFax.TabIndex = 44
        Me.lblFax.Text = "Fax"
        '
        'txtOwnerAddress
        '
        Me.txtOwnerAddress.Location = New System.Drawing.Point(80, 56)
        Me.txtOwnerAddress.Multiline = True
        Me.txtOwnerAddress.Name = "txtOwnerAddress"
        Me.txtOwnerAddress.ReadOnly = True
        Me.txtOwnerAddress.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
        Me.txtOwnerAddress.Size = New System.Drawing.Size(248, 103)
        Me.txtOwnerAddress.TabIndex = 1
        Me.txtOwnerAddress.Text = ""
        Me.txtOwnerAddress.WordWrap = False
        '
        'lblOwnerAddress
        '
        Me.lblOwnerAddress.Location = New System.Drawing.Point(7, 56)
        Me.lblOwnerAddress.Name = "lblOwnerAddress"
        Me.lblOwnerAddress.Size = New System.Drawing.Size(72, 16)
        Me.lblOwnerAddress.TabIndex = 88
        Me.lblOwnerAddress.Text = "Address"
        '
        'txtOwnerName
        '
        Me.txtOwnerName.Location = New System.Drawing.Point(80, 32)
        Me.txtOwnerName.Name = "txtOwnerName"
        Me.txtOwnerName.ReadOnly = True
        Me.txtOwnerName.Size = New System.Drawing.Size(248, 21)
        Me.txtOwnerName.TabIndex = 0
        Me.txtOwnerName.Text = ""
        '
        'lblOwnerName
        '
        Me.lblOwnerName.Location = New System.Drawing.Point(8, 32)
        Me.lblOwnerName.Name = "lblOwnerName"
        Me.lblOwnerName.Size = New System.Drawing.Size(88, 23)
        Me.lblOwnerName.TabIndex = 86
        Me.lblOwnerName.Text = "Name"
        '
        'lblOwnerStatus
        '
        Me.lblOwnerStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerStatus.Location = New System.Drawing.Point(346, 8)
        Me.lblOwnerStatus.Name = "lblOwnerStatus"
        Me.lblOwnerStatus.Size = New System.Drawing.Size(78, 23)
        Me.lblOwnerStatus.TabIndex = 84
        Me.lblOwnerStatus.Text = "Owner Status"
        '
        'lblOwnerCapParticipant
        '
        Me.lblOwnerCapParticipant.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerCapParticipant.Location = New System.Drawing.Point(544, 52)
        Me.lblOwnerCapParticipant.Name = "lblOwnerCapParticipant"
        Me.lblOwnerCapParticipant.Size = New System.Drawing.Size(128, 23)
        Me.lblOwnerCapParticipant.TabIndex = 52
        Me.lblOwnerCapParticipant.Text = "CAP Participation Level"
        '
        'lblPhone2
        '
        Me.lblPhone2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPhone2.Location = New System.Drawing.Point(346, 72)
        Me.lblPhone2.Name = "lblPhone2"
        Me.lblPhone2.Size = New System.Drawing.Size(56, 23)
        Me.lblPhone2.TabIndex = 45
        Me.lblPhone2.Text = "Phone 2"
        '
        'lblOwnerType
        '
        Me.lblOwnerType.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerType.Location = New System.Drawing.Point(8, 160)
        Me.lblOwnerType.Name = "lblOwnerType"
        Me.lblOwnerType.Size = New System.Drawing.Size(72, 23)
        Me.lblOwnerType.TabIndex = 40
        Me.lblOwnerType.Text = "Owner Type:"
        '
        'txtOwnerAIID
        '
        Me.txtOwnerAIID.AcceptsTab = True
        Me.txtOwnerAIID.AutoSize = False
        Me.txtOwnerAIID.Enabled = False
        Me.txtOwnerAIID.Location = New System.Drawing.Point(424, 24)
        Me.txtOwnerAIID.Name = "txtOwnerAIID"
        Me.txtOwnerAIID.Size = New System.Drawing.Size(96, 21)
        Me.txtOwnerAIID.TabIndex = 3
        Me.txtOwnerAIID.Text = ""
        Me.txtOwnerAIID.WordWrap = False
        '
        'lblOwnerAIID
        '
        Me.lblOwnerAIID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerAIID.Location = New System.Drawing.Point(346, 30)
        Me.lblOwnerAIID.Name = "lblOwnerAIID"
        Me.lblOwnerAIID.Size = New System.Drawing.Size(72, 16)
        Me.lblOwnerAIID.TabIndex = 38
        Me.lblOwnerAIID.Text = "Ensite ID"
        '
        'lblOwnerIDValue
        '
        Me.lblOwnerIDValue.Enabled = False
        Me.lblOwnerIDValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerIDValue.Location = New System.Drawing.Point(86, 8)
        Me.lblOwnerIDValue.Name = "lblOwnerIDValue"
        Me.lblOwnerIDValue.Size = New System.Drawing.Size(96, 23)
        Me.lblOwnerIDValue.TabIndex = 0
        '
        'lblOwnerID
        '
        Me.lblOwnerID.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerID.Location = New System.Drawing.Point(8, 8)
        Me.lblOwnerID.Name = "lblOwnerID"
        Me.lblOwnerID.Size = New System.Drawing.Size(64, 23)
        Me.lblOwnerID.TabIndex = 36
        Me.lblOwnerID.Text = "Owner ID"
        '
        'lblOwnerPhone
        '
        Me.lblOwnerPhone.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerPhone.Location = New System.Drawing.Point(346, 48)
        Me.lblOwnerPhone.Name = "lblOwnerPhone"
        Me.lblOwnerPhone.Size = New System.Drawing.Size(56, 23)
        Me.lblOwnerPhone.TabIndex = 32
        Me.lblOwnerPhone.Text = "Phone"
        '
        'cmbOwnerType
        '
        Me.cmbOwnerType.DisplayMember = "1"
        Me.cmbOwnerType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbOwnerType.DropDownWidth = 200
        Me.cmbOwnerType.Enabled = False
        Me.cmbOwnerType.ItemHeight = 15
        Me.cmbOwnerType.Location = New System.Drawing.Point(80, 160)
        Me.cmbOwnerType.Name = "cmbOwnerType"
        Me.cmbOwnerType.Size = New System.Drawing.Size(248, 23)
        Me.cmbOwnerType.TabIndex = 2
        Me.cmbOwnerType.ValueMember = "1"
        '
        'btnCAEOwnerLabels
        '
        Me.btnCAEOwnerLabels.Location = New System.Drawing.Point(5, 116)
        Me.btnCAEOwnerLabels.Name = "btnCAEOwnerLabels"
        Me.btnCAEOwnerLabels.Size = New System.Drawing.Size(70, 23)
        Me.btnCAEOwnerLabels.TabIndex = 1068
        Me.btnCAEOwnerLabels.Text = "Labels"
        '
        'btnCAEOwnerEnvelopes
        '
        Me.btnCAEOwnerEnvelopes.Location = New System.Drawing.Point(5, 88)
        Me.btnCAEOwnerEnvelopes.Name = "btnCAEOwnerEnvelopes"
        Me.btnCAEOwnerEnvelopes.Size = New System.Drawing.Size(70, 23)
        Me.btnCAEOwnerEnvelopes.TabIndex = 1067
        Me.btnCAEOwnerEnvelopes.Text = "Envelopes"
        '
        'tbPageFacilityDetail
        '
        Me.tbPageFacilityDetail.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.tbPageFacilityDetail.Controls.Add(Me.pnlFacilityBottom)
        Me.tbPageFacilityDetail.Controls.Add(Me.pnl_FacilityDetail)
        Me.tbPageFacilityDetail.Location = New System.Drawing.Point(4, 22)
        Me.tbPageFacilityDetail.Name = "tbPageFacilityDetail"
        Me.tbPageFacilityDetail.Size = New System.Drawing.Size(1020, 644)
        Me.tbPageFacilityDetail.TabIndex = 8
        Me.tbPageFacilityDetail.Text = "Facility Details"
        Me.tbPageFacilityDetail.Visible = False
        '
        'pnlFacilityBottom
        '
        Me.pnlFacilityBottom.AutoScroll = True
        Me.pnlFacilityBottom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlFacilityBottom.Controls.Add(Me.tbCntrlFacility)
        Me.pnlFacilityBottom.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlFacilityBottom.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlFacilityBottom.Location = New System.Drawing.Point(0, 304)
        Me.pnlFacilityBottom.Name = "pnlFacilityBottom"
        Me.pnlFacilityBottom.Size = New System.Drawing.Size(1016, 336)
        Me.pnlFacilityBottom.TabIndex = 3
        '
        'tbCntrlFacility
        '
        Me.tbCntrlFacility.Controls.Add(Me.tbPageFCEs)
        Me.tbCntrlFacility.Controls.Add(Me.tbPageFacilityDocuments)
        Me.tbCntrlFacility.Controls.Add(Me.tbPageFacilityContactList)
        Me.tbCntrlFacility.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbCntrlFacility.Location = New System.Drawing.Point(0, 0)
        Me.tbCntrlFacility.Name = "tbCntrlFacility"
        Me.tbCntrlFacility.SelectedIndex = 0
        Me.tbCntrlFacility.Size = New System.Drawing.Size(1014, 334)
        Me.tbCntrlFacility.TabIndex = 31
        '
        'tbPageFCEs
        '
        Me.tbPageFCEs.AutoScroll = True
        Me.tbPageFCEs.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.tbPageFCEs.Controls.Add(Me.ugFCEs)
        Me.tbPageFCEs.Controls.Add(Me.pnlFacilityLustButton)
        Me.tbPageFCEs.Location = New System.Drawing.Point(4, 24)
        Me.tbPageFCEs.Name = "tbPageFCEs"
        Me.tbPageFCEs.Size = New System.Drawing.Size(1006, 306)
        Me.tbPageFCEs.TabIndex = 0
        Me.tbPageFCEs.Text = "FCEs"
        '
        'ugFCEs
        '
        Me.ugFCEs.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugFCEs.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugFCEs.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        Me.ugFCEs.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        Me.ugFCEs.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugFCEs.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugFCEs.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugFCEs.Location = New System.Drawing.Point(0, 0)
        Me.ugFCEs.Name = "ugFCEs"
        Me.ugFCEs.Size = New System.Drawing.Size(1002, 278)
        Me.ugFCEs.TabIndex = 30
        '
        'pnlFacilityLustButton
        '
        Me.pnlFacilityLustButton.Controls.Add(Me.lblTotalNoOfLUSTEventsValue)
        Me.pnlFacilityLustButton.Controls.Add(Me.lblTotalNoOfLUSTEvents)
        Me.pnlFacilityLustButton.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlFacilityLustButton.Location = New System.Drawing.Point(0, 278)
        Me.pnlFacilityLustButton.Name = "pnlFacilityLustButton"
        Me.pnlFacilityLustButton.Size = New System.Drawing.Size(1002, 24)
        Me.pnlFacilityLustButton.TabIndex = 97
        '
        'lblTotalNoOfLUSTEventsValue
        '
        Me.lblTotalNoOfLUSTEventsValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalNoOfLUSTEventsValue.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblTotalNoOfLUSTEventsValue.Location = New System.Drawing.Point(192, 0)
        Me.lblTotalNoOfLUSTEventsValue.Name = "lblTotalNoOfLUSTEventsValue"
        Me.lblTotalNoOfLUSTEventsValue.Size = New System.Drawing.Size(48, 24)
        Me.lblTotalNoOfLUSTEventsValue.TabIndex = 5
        Me.lblTotalNoOfLUSTEventsValue.Text = "0"
        '
        'lblTotalNoOfLUSTEvents
        '
        Me.lblTotalNoOfLUSTEvents.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalNoOfLUSTEvents.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblTotalNoOfLUSTEvents.Location = New System.Drawing.Point(0, 0)
        Me.lblTotalNoOfLUSTEvents.Name = "lblTotalNoOfLUSTEvents"
        Me.lblTotalNoOfLUSTEvents.Size = New System.Drawing.Size(192, 24)
        Me.lblTotalNoOfLUSTEvents.TabIndex = 4
        Me.lblTotalNoOfLUSTEvents.Text = "Number of FCEs at this Location:"
        '
        'tbPageFacilityDocuments
        '
        Me.tbPageFacilityDocuments.AutoScroll = True
        Me.tbPageFacilityDocuments.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.tbPageFacilityDocuments.Controls.Add(Me.UCFacilityDocuments)
        Me.tbPageFacilityDocuments.Location = New System.Drawing.Point(4, 24)
        Me.tbPageFacilityDocuments.Name = "tbPageFacilityDocuments"
        Me.tbPageFacilityDocuments.Size = New System.Drawing.Size(1006, 306)
        Me.tbPageFacilityDocuments.TabIndex = 1
        Me.tbPageFacilityDocuments.Text = "Documents"
        '
        'UCFacilityDocuments
        '
        Me.UCFacilityDocuments.AutoScroll = True
        Me.UCFacilityDocuments.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UCFacilityDocuments.Location = New System.Drawing.Point(0, 0)
        Me.UCFacilityDocuments.Name = "UCFacilityDocuments"
        Me.UCFacilityDocuments.Size = New System.Drawing.Size(1002, 302)
        Me.UCFacilityDocuments.TabIndex = 3
        '
        'tbPageFacilityContactList
        '
        Me.tbPageFacilityContactList.Controls.Add(Me.pnlFacilityContactContainer)
        Me.tbPageFacilityContactList.Controls.Add(Me.pnlFacilityContactBottom)
        Me.tbPageFacilityContactList.Controls.Add(Me.pnlFacilityContactHeader)
        Me.tbPageFacilityContactList.Location = New System.Drawing.Point(4, 24)
        Me.tbPageFacilityContactList.Name = "tbPageFacilityContactList"
        Me.tbPageFacilityContactList.Size = New System.Drawing.Size(1006, 306)
        Me.tbPageFacilityContactList.TabIndex = 2
        Me.tbPageFacilityContactList.Text = "Contacts"
        '
        'pnlFacilityContactContainer
        '
        Me.pnlFacilityContactContainer.Controls.Add(Me.ugFacilityContacts)
        Me.pnlFacilityContactContainer.Controls.Add(Me.Label2)
        Me.pnlFacilityContactContainer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlFacilityContactContainer.Location = New System.Drawing.Point(0, 25)
        Me.pnlFacilityContactContainer.Name = "pnlFacilityContactContainer"
        Me.pnlFacilityContactContainer.Size = New System.Drawing.Size(1006, 251)
        Me.pnlFacilityContactContainer.TabIndex = 4
        '
        'ugFacilityContacts
        '
        Me.ugFacilityContacts.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugFacilityContacts.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugFacilityContacts.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        Me.ugFacilityContacts.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugFacilityContacts.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugFacilityContacts.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugFacilityContacts.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugFacilityContacts.Location = New System.Drawing.Point(0, 0)
        Me.ugFacilityContacts.Name = "ugFacilityContacts"
        Me.ugFacilityContacts.Size = New System.Drawing.Size(1006, 251)
        Me.ugFacilityContacts.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(792, 56)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(7, 23)
        Me.Label2.TabIndex = 2
        '
        'pnlFacilityContactBottom
        '
        Me.pnlFacilityContactBottom.Controls.Add(Me.btnFacilityModifyContact)
        Me.pnlFacilityContactBottom.Controls.Add(Me.btnFacilityDeleteContact)
        Me.pnlFacilityContactBottom.Controls.Add(Me.btnFacilityAssociateContact)
        Me.pnlFacilityContactBottom.Controls.Add(Me.btnFacilityAddSearchContact)
        Me.pnlFacilityContactBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlFacilityContactBottom.DockPadding.All = 3
        Me.pnlFacilityContactBottom.Location = New System.Drawing.Point(0, 276)
        Me.pnlFacilityContactBottom.Name = "pnlFacilityContactBottom"
        Me.pnlFacilityContactBottom.Size = New System.Drawing.Size(1006, 30)
        Me.pnlFacilityContactBottom.TabIndex = 3
        '
        'btnFacilityModifyContact
        '
        Me.btnFacilityModifyContact.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFacilityModifyContact.Location = New System.Drawing.Point(240, 4)
        Me.btnFacilityModifyContact.Name = "btnFacilityModifyContact"
        Me.btnFacilityModifyContact.Size = New System.Drawing.Size(235, 23)
        Me.btnFacilityModifyContact.TabIndex = 1
        Me.btnFacilityModifyContact.Text = "Modify Contact"
        '
        'btnFacilityDeleteContact
        '
        Me.btnFacilityDeleteContact.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFacilityDeleteContact.Location = New System.Drawing.Point(472, 4)
        Me.btnFacilityDeleteContact.Name = "btnFacilityDeleteContact"
        Me.btnFacilityDeleteContact.Size = New System.Drawing.Size(235, 23)
        Me.btnFacilityDeleteContact.TabIndex = 2
        Me.btnFacilityDeleteContact.Text = "Disassociate Contact"
        '
        'btnFacilityAssociateContact
        '
        Me.btnFacilityAssociateContact.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFacilityAssociateContact.Location = New System.Drawing.Point(704, 4)
        Me.btnFacilityAssociateContact.Name = "btnFacilityAssociateContact"
        Me.btnFacilityAssociateContact.Size = New System.Drawing.Size(235, 23)
        Me.btnFacilityAssociateContact.TabIndex = 3
        Me.btnFacilityAssociateContact.Text = "Associate Contact from Different Module"
        '
        'btnFacilityAddSearchContact
        '
        Me.btnFacilityAddSearchContact.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFacilityAddSearchContact.Location = New System.Drawing.Point(8, 4)
        Me.btnFacilityAddSearchContact.Name = "btnFacilityAddSearchContact"
        Me.btnFacilityAddSearchContact.Size = New System.Drawing.Size(235, 23)
        Me.btnFacilityAddSearchContact.TabIndex = 0
        Me.btnFacilityAddSearchContact.Text = "Add/Search Contact to Associate"
        '
        'pnlFacilityContactHeader
        '
        Me.pnlFacilityContactHeader.Controls.Add(Me.chkFacilityShowActiveContactOnly)
        Me.pnlFacilityContactHeader.Controls.Add(Me.chkFacilityShowRelatedContacts)
        Me.pnlFacilityContactHeader.Controls.Add(Me.chkFacilityShowContactsforAllModule)
        Me.pnlFacilityContactHeader.Controls.Add(Me.lblFacilityContacts)
        Me.pnlFacilityContactHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlFacilityContactHeader.DockPadding.All = 3
        Me.pnlFacilityContactHeader.Location = New System.Drawing.Point(0, 0)
        Me.pnlFacilityContactHeader.Name = "pnlFacilityContactHeader"
        Me.pnlFacilityContactHeader.Size = New System.Drawing.Size(1006, 25)
        Me.pnlFacilityContactHeader.TabIndex = 2
        '
        'chkFacilityShowActiveContactOnly
        '
        Me.chkFacilityShowActiveContactOnly.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkFacilityShowActiveContactOnly.Location = New System.Drawing.Point(635, 7)
        Me.chkFacilityShowActiveContactOnly.Name = "chkFacilityShowActiveContactOnly"
        Me.chkFacilityShowActiveContactOnly.Size = New System.Drawing.Size(144, 16)
        Me.chkFacilityShowActiveContactOnly.TabIndex = 2
        Me.chkFacilityShowActiveContactOnly.Tag = "646"
        Me.chkFacilityShowActiveContactOnly.Text = "Show Active Only"
        '
        'chkFacilityShowRelatedContacts
        '
        Me.chkFacilityShowRelatedContacts.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkFacilityShowRelatedContacts.Location = New System.Drawing.Point(467, 7)
        Me.chkFacilityShowRelatedContacts.Name = "chkFacilityShowRelatedContacts"
        Me.chkFacilityShowRelatedContacts.Size = New System.Drawing.Size(160, 16)
        Me.chkFacilityShowRelatedContacts.TabIndex = 1
        Me.chkFacilityShowRelatedContacts.Tag = "645"
        Me.chkFacilityShowRelatedContacts.Text = "Show Related Contacts"
        '
        'chkFacilityShowContactsforAllModule
        '
        Me.chkFacilityShowContactsforAllModule.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkFacilityShowContactsforAllModule.Location = New System.Drawing.Point(251, 7)
        Me.chkFacilityShowContactsforAllModule.Name = "chkFacilityShowContactsforAllModule"
        Me.chkFacilityShowContactsforAllModule.Size = New System.Drawing.Size(200, 16)
        Me.chkFacilityShowContactsforAllModule.TabIndex = 0
        Me.chkFacilityShowContactsforAllModule.Tag = "644"
        Me.chkFacilityShowContactsforAllModule.Text = "Show Contacts for All Modules"
        '
        'lblFacilityContacts
        '
        Me.lblFacilityContacts.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityContacts.Location = New System.Drawing.Point(8, 7)
        Me.lblFacilityContacts.Name = "lblFacilityContacts"
        Me.lblFacilityContacts.Size = New System.Drawing.Size(104, 16)
        Me.lblFacilityContacts.TabIndex = 139
        Me.lblFacilityContacts.Text = "Facility Contacts"
        '
        'pnl_FacilityDetail
        '
        Me.pnl_FacilityDetail.BackColor = System.Drawing.SystemColors.Control
        Me.pnl_FacilityDetail.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnl_FacilityDetail.Controls.Add(Me.txtDesMa)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblDesMa)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtDesOp)
        Me.pnl_FacilityDetail.Controls.Add(Me.Label3)
        Me.pnl_FacilityDetail.Controls.Add(Me.btnCAEFacLabels)
        Me.pnl_FacilityDetail.Controls.Add(Me.btnCAEFacEnvelopes)
        Me.pnl_FacilityDetail.Controls.Add(Me.btnFacComments)
        Me.pnl_FacilityDetail.Controls.Add(Me.btnFacFlags)
        Me.pnl_FacilityDetail.Controls.Add(Me.btnFacilityCancel)
        Me.pnl_FacilityDetail.Controls.Add(Me.btnFacilitySave)
        Me.pnl_FacilityDetail.Controls.Add(Me.dtPickUpcomingInstallDateValue)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblUpcomingInstallDate)
        Me.pnl_FacilityDetail.Controls.Add(Me.chkUpcomingInstall)
        Me.pnl_FacilityDetail.Controls.Add(Me.lnkLblNextFac)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblCAPStatusValue)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblCAPStatus)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFuelBrand)
        Me.pnl_FacilityDetail.Controls.Add(Me.ll)
        Me.pnl_FacilityDetail.Controls.Add(Me.dtFacilityPowerOff)
        Me.pnl_FacilityDetail.Controls.Add(Me.lnkLblPrevFacility)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblLUSTSite)
        Me.pnl_FacilityDetail.Controls.Add(Me.chkLUSTSite)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblPowerOff)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblCAPCandidate)
        Me.pnl_FacilityDetail.Controls.Add(Me.chkCAPCandidate)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityLocationType)
        Me.pnl_FacilityDetail.Controls.Add(Me.cmbFacilityLocationType)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityMethod)
        Me.pnl_FacilityDetail.Controls.Add(Me.cmbFacilityMethod)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityDatum)
        Me.pnl_FacilityDetail.Controls.Add(Me.cmbFacilityDatum)
        Me.pnl_FacilityDetail.Controls.Add(Me.cmbFacilityType)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilityLatSec)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilityLongSec)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilityLatMin)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilityLongMin)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityLongMin)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityLongSec)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityLatMin)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityLatSec)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityLongDegree)
        Me.pnl_FacilityDetail.Controls.Add(Me.mskTxtFacilityFax)
        Me.pnl_FacilityDetail.Controls.Add(Me.mskTxtFacilityPhone)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilityAddress)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilitySIC)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityFax)
        Me.pnl_FacilityDetail.Controls.Add(Me.dtPickFacilityRecvd)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblDateReceived)
        Me.pnl_FacilityDetail.Controls.Add(Me.chkSignatureofNF)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilitySigOnNF)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityFuelBrand)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityStatusValue)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityStatus)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilityLongDegree)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilityLatDegree)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityLongitude)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityLatitude)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityType)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilityAIID)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblfacilityAIID)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityIDValue)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityID)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityPhone)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilityName)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityAddress)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityName)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityLatDegree)
        Me.pnl_FacilityDetail.Controls.Add(Me.btnInspHistory)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilitySIC)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblInViolationValue)
        Me.pnl_FacilityDetail.Controls.Add(Me.Label4)
        Me.pnl_FacilityDetail.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnl_FacilityDetail.Location = New System.Drawing.Point(0, 0)
        Me.pnl_FacilityDetail.Name = "pnl_FacilityDetail"
        Me.pnl_FacilityDetail.Size = New System.Drawing.Size(1016, 304)
        Me.pnl_FacilityDetail.TabIndex = 1
        '
        'txtDesMa
        '
        Me.txtDesMa.Enabled = False
        Me.txtDesMa.Location = New System.Drawing.Point(688, 202)
        Me.txtDesMa.Name = "txtDesMa"
        Me.txtDesMa.Size = New System.Drawing.Size(176, 21)
        Me.txtDesMa.TabIndex = 1072
        Me.txtDesMa.Text = ""
        '
        'lblDesMa
        '
        Me.lblDesMa.Location = New System.Drawing.Point(600, 200)
        Me.lblDesMa.Name = "lblDesMa"
        Me.lblDesMa.Size = New System.Drawing.Size(72, 23)
        Me.lblDesMa.TabIndex = 1071
        Me.lblDesMa.Text = "Des. Magr.:"
        '
        'txtDesOp
        '
        Me.txtDesOp.AcceptsTab = True
        Me.txtDesOp.AutoSize = False
        Me.txtDesOp.Enabled = False
        Me.txtDesOp.Location = New System.Drawing.Point(688, 176)
        Me.txtDesOp.Name = "txtDesOp"
        Me.txtDesOp.Size = New System.Drawing.Size(344, 21)
        Me.txtDesOp.TabIndex = 1069
        Me.txtDesOp.Text = ""
        Me.txtDesOp.WordWrap = False
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(600, 176)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(88, 23)
        Me.Label3.TabIndex = 1070
        Me.Label3.Text = "Des. Oper. :"
        '
        'btnCAEFacLabels
        '
        Me.btnCAEFacLabels.Location = New System.Drawing.Point(8, 119)
        Me.btnCAEFacLabels.Name = "btnCAEFacLabels"
        Me.btnCAEFacLabels.Size = New System.Drawing.Size(70, 23)
        Me.btnCAEFacLabels.TabIndex = 1068
        Me.btnCAEFacLabels.Text = "Labels"
        '
        'btnCAEFacEnvelopes
        '
        Me.btnCAEFacEnvelopes.Location = New System.Drawing.Point(8, 89)
        Me.btnCAEFacEnvelopes.Name = "btnCAEFacEnvelopes"
        Me.btnCAEFacEnvelopes.Size = New System.Drawing.Size(70, 23)
        Me.btnCAEFacEnvelopes.TabIndex = 1067
        Me.btnCAEFacEnvelopes.Text = "Envelopes"
        '
        'btnFacComments
        '
        Me.btnFacComments.Location = New System.Drawing.Point(864, 272)
        Me.btnFacComments.Name = "btnFacComments"
        Me.btnFacComments.Size = New System.Drawing.Size(80, 23)
        Me.btnFacComments.TabIndex = 1050
        Me.btnFacComments.Text = "Comments"
        '
        'btnFacFlags
        '
        Me.btnFacFlags.Location = New System.Drawing.Point(784, 272)
        Me.btnFacFlags.Name = "btnFacFlags"
        Me.btnFacFlags.TabIndex = 1051
        Me.btnFacFlags.Text = "Flags"
        '
        'btnFacilityCancel
        '
        Me.btnFacilityCancel.Enabled = False
        Me.btnFacilityCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFacilityCancel.Location = New System.Drawing.Point(696, 272)
        Me.btnFacilityCancel.Name = "btnFacilityCancel"
        Me.btnFacilityCancel.TabIndex = 1046
        Me.btnFacilityCancel.Text = "Cancel"
        '
        'btnFacilitySave
        '
        Me.btnFacilitySave.Enabled = False
        Me.btnFacilitySave.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFacilitySave.Location = New System.Drawing.Point(592, 272)
        Me.btnFacilitySave.Name = "btnFacilitySave"
        Me.btnFacilitySave.Size = New System.Drawing.Size(96, 23)
        Me.btnFacilitySave.TabIndex = 1045
        Me.btnFacilitySave.Text = "Save Facility"
        Me.btnFacilitySave.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'dtPickUpcomingInstallDateValue
        '
        Me.dtPickUpcomingInstallDateValue.Checked = False
        Me.dtPickUpcomingInstallDateValue.Enabled = False
        Me.dtPickUpcomingInstallDateValue.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickUpcomingInstallDateValue.Location = New System.Drawing.Point(480, 200)
        Me.dtPickUpcomingInstallDateValue.Name = "dtPickUpcomingInstallDateValue"
        Me.dtPickUpcomingInstallDateValue.Size = New System.Drawing.Size(88, 21)
        Me.dtPickUpcomingInstallDateValue.TabIndex = 12
        '
        'lblUpcomingInstallDate
        '
        Me.lblUpcomingInstallDate.Location = New System.Drawing.Point(320, 200)
        Me.lblUpcomingInstallDate.Name = "lblUpcomingInstallDate"
        Me.lblUpcomingInstallDate.Size = New System.Drawing.Size(160, 23)
        Me.lblUpcomingInstallDate.TabIndex = 1044
        Me.lblUpcomingInstallDate.Text = "Upcoming Installation Date"
        Me.lblUpcomingInstallDate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkUpcomingInstall
        '
        Me.chkUpcomingInstall.Enabled = False
        Me.chkUpcomingInstall.Location = New System.Drawing.Point(312, 176)
        Me.chkUpcomingInstall.Name = "chkUpcomingInstall"
        Me.chkUpcomingInstall.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkUpcomingInstall.Size = New System.Drawing.Size(152, 23)
        Me.chkUpcomingInstall.TabIndex = 11
        Me.chkUpcomingInstall.Text = "Upcoming Installation "
        '
        'lnkLblNextFac
        '
        Me.lnkLblNextFac.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lnkLblNextFac.Location = New System.Drawing.Point(256, 8)
        Me.lnkLblNextFac.Name = "lnkLblNextFac"
        Me.lnkLblNextFac.Size = New System.Drawing.Size(56, 16)
        Me.lnkLblNextFac.TabIndex = 26
        Me.lnkLblNextFac.TabStop = True
        Me.lnkLblNextFac.Text = "Next>>"
        '
        'lblCAPStatusValue
        '
        Me.lblCAPStatusValue.BackColor = System.Drawing.SystemColors.Control
        Me.lblCAPStatusValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblCAPStatusValue.Enabled = False
        Me.lblCAPStatusValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCAPStatusValue.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblCAPStatusValue.Location = New System.Drawing.Point(424, 85)
        Me.lblCAPStatusValue.Name = "lblCAPStatusValue"
        Me.lblCAPStatusValue.Size = New System.Drawing.Size(120, 16)
        Me.lblCAPStatusValue.TabIndex = 1038
        '
        'lblCAPStatus
        '
        Me.lblCAPStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCAPStatus.Location = New System.Drawing.Point(320, 78)
        Me.lblCAPStatus.Name = "lblCAPStatus"
        Me.lblCAPStatus.Size = New System.Drawing.Size(72, 23)
        Me.lblCAPStatus.TabIndex = 1037
        Me.lblCAPStatus.Text = "CAP Status:"
        Me.lblCAPStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtFuelBrand
        '
        Me.txtFuelBrand.Location = New System.Drawing.Point(424, 152)
        Me.txtFuelBrand.Name = "txtFuelBrand"
        Me.txtFuelBrand.ReadOnly = True
        Me.txtFuelBrand.Size = New System.Drawing.Size(120, 21)
        Me.txtFuelBrand.TabIndex = 10
        Me.txtFuelBrand.Text = ""
        '
        'll
        '
        Me.ll.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ll.Location = New System.Drawing.Point(1128, 56)
        Me.ll.Name = "ll"
        Me.ll.Size = New System.Drawing.Size(24, 23)
        Me.ll.TabIndex = 1035
        '
        'dtFacilityPowerOff
        '
        Me.dtFacilityPowerOff.Checked = False
        Me.dtFacilityPowerOff.Enabled = False
        Me.dtFacilityPowerOff.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtFacilityPowerOff.Location = New System.Drawing.Point(960, 64)
        Me.dtFacilityPowerOff.Name = "dtFacilityPowerOff"
        Me.dtFacilityPowerOff.ShowCheckBox = True
        Me.dtFacilityPowerOff.Size = New System.Drawing.Size(104, 21)
        Me.dtFacilityPowerOff.TabIndex = 9
        Me.dtFacilityPowerOff.Visible = False
        '
        'lnkLblPrevFacility
        '
        Me.lnkLblPrevFacility.AutoSize = True
        Me.lnkLblPrevFacility.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lnkLblPrevFacility.Location = New System.Drawing.Point(176, 8)
        Me.lnkLblPrevFacility.Name = "lnkLblPrevFacility"
        Me.lnkLblPrevFacility.Size = New System.Drawing.Size(70, 17)
        Me.lnkLblPrevFacility.TabIndex = 25
        Me.lnkLblPrevFacility.TabStop = True
        Me.lnkLblPrevFacility.Text = "<< Previous"
        '
        'lblLUSTSite
        '
        Me.lblLUSTSite.Location = New System.Drawing.Point(320, 55)
        Me.lblLUSTSite.Name = "lblLUSTSite"
        Me.lblLUSTSite.Size = New System.Drawing.Size(104, 23)
        Me.lblLUSTSite.TabIndex = 1030
        Me.lblLUSTSite.Text = "Active LUST Site"
        Me.lblLUSTSite.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkLUSTSite
        '
        Me.chkLUSTSite.Enabled = False
        Me.chkLUSTSite.Location = New System.Drawing.Point(424, 62)
        Me.chkLUSTSite.Name = "chkLUSTSite"
        Me.chkLUSTSite.Size = New System.Drawing.Size(16, 16)
        Me.chkLUSTSite.TabIndex = 6
        Me.chkLUSTSite.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPowerOff
        '
        Me.lblPowerOff.Location = New System.Drawing.Point(888, 64)
        Me.lblPowerOff.Name = "lblPowerOff"
        Me.lblPowerOff.Size = New System.Drawing.Size(72, 23)
        Me.lblPowerOff.TabIndex = 1028
        Me.lblPowerOff.Text = "Power Off"
        Me.lblPowerOff.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblPowerOff.Visible = False
        '
        'lblCAPCandidate
        '
        Me.lblCAPCandidate.Location = New System.Drawing.Point(320, 101)
        Me.lblCAPCandidate.Name = "lblCAPCandidate"
        Me.lblCAPCandidate.Size = New System.Drawing.Size(96, 23)
        Me.lblCAPCandidate.TabIndex = 1026
        Me.lblCAPCandidate.Text = "CAP Candidate"
        Me.lblCAPCandidate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkCAPCandidate
        '
        Me.chkCAPCandidate.Enabled = False
        Me.chkCAPCandidate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkCAPCandidate.Location = New System.Drawing.Point(424, 108)
        Me.chkCAPCandidate.Name = "chkCAPCandidate"
        Me.chkCAPCandidate.Size = New System.Drawing.Size(16, 16)
        Me.chkCAPCandidate.TabIndex = 7
        Me.chkCAPCandidate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblFacilityLocationType
        '
        Me.lblFacilityLocationType.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLocationType.Location = New System.Drawing.Point(600, 152)
        Me.lblFacilityLocationType.Name = "lblFacilityLocationType"
        Me.lblFacilityLocationType.Size = New System.Drawing.Size(80, 23)
        Me.lblFacilityLocationType.TabIndex = 1024
        Me.lblFacilityLocationType.Text = "Type:"
        '
        'cmbFacilityLocationType
        '
        Me.cmbFacilityLocationType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbFacilityLocationType.DropDownWidth = 250
        Me.cmbFacilityLocationType.Enabled = False
        Me.cmbFacilityLocationType.ItemHeight = 15
        Me.cmbFacilityLocationType.Location = New System.Drawing.Point(688, 150)
        Me.cmbFacilityLocationType.Name = "cmbFacilityLocationType"
        Me.cmbFacilityLocationType.Size = New System.Drawing.Size(176, 23)
        Me.cmbFacilityLocationType.TabIndex = 23
        '
        'lblFacilityMethod
        '
        Me.lblFacilityMethod.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityMethod.Location = New System.Drawing.Point(600, 126)
        Me.lblFacilityMethod.Name = "lblFacilityMethod"
        Me.lblFacilityMethod.Size = New System.Drawing.Size(80, 23)
        Me.lblFacilityMethod.TabIndex = 1022
        Me.lblFacilityMethod.Text = "Method:"
        '
        'cmbFacilityMethod
        '
        Me.cmbFacilityMethod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbFacilityMethod.DropDownWidth = 350
        Me.cmbFacilityMethod.Enabled = False
        Me.cmbFacilityMethod.ItemHeight = 15
        Me.cmbFacilityMethod.Location = New System.Drawing.Point(688, 126)
        Me.cmbFacilityMethod.Name = "cmbFacilityMethod"
        Me.cmbFacilityMethod.Size = New System.Drawing.Size(176, 23)
        Me.cmbFacilityMethod.TabIndex = 22
        '
        'lblFacilityDatum
        '
        Me.lblFacilityDatum.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityDatum.Location = New System.Drawing.Point(600, 104)
        Me.lblFacilityDatum.Name = "lblFacilityDatum"
        Me.lblFacilityDatum.Size = New System.Drawing.Size(80, 23)
        Me.lblFacilityDatum.TabIndex = 1020
        Me.lblFacilityDatum.Text = "Datum:"
        '
        'cmbFacilityDatum
        '
        Me.cmbFacilityDatum.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbFacilityDatum.DropDownWidth = 250
        Me.cmbFacilityDatum.Enabled = False
        Me.cmbFacilityDatum.ItemHeight = 15
        Me.cmbFacilityDatum.Location = New System.Drawing.Point(688, 104)
        Me.cmbFacilityDatum.Name = "cmbFacilityDatum"
        Me.cmbFacilityDatum.Size = New System.Drawing.Size(176, 23)
        Me.cmbFacilityDatum.TabIndex = 21
        '
        'cmbFacilityType
        '
        Me.cmbFacilityType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbFacilityType.DropDownWidth = 180
        Me.cmbFacilityType.Enabled = False
        Me.cmbFacilityType.ItemHeight = 15
        Me.cmbFacilityType.Location = New System.Drawing.Point(688, 32)
        Me.cmbFacilityType.Name = "cmbFacilityType"
        Me.cmbFacilityType.Size = New System.Drawing.Size(176, 23)
        Me.cmbFacilityType.TabIndex = 14
        '
        'txtFacilityLatSec
        '
        Me.txtFacilityLatSec.AcceptsTab = True
        Me.txtFacilityLatSec.AutoSize = False
        Me.txtFacilityLatSec.Enabled = False
        Me.txtFacilityLatSec.Location = New System.Drawing.Point(768, 56)
        Me.txtFacilityLatSec.MaxLength = 5
        Me.txtFacilityLatSec.Name = "txtFacilityLatSec"
        Me.txtFacilityLatSec.Size = New System.Drawing.Size(37, 21)
        Me.txtFacilityLatSec.TabIndex = 17
        Me.txtFacilityLatSec.Text = ""
        Me.txtFacilityLatSec.WordWrap = False
        '
        'txtFacilityLongSec
        '
        Me.txtFacilityLongSec.AcceptsTab = True
        Me.txtFacilityLongSec.AutoSize = False
        Me.txtFacilityLongSec.Enabled = False
        Me.txtFacilityLongSec.Location = New System.Drawing.Point(768, 80)
        Me.txtFacilityLongSec.MaxLength = 5
        Me.txtFacilityLongSec.Name = "txtFacilityLongSec"
        Me.txtFacilityLongSec.Size = New System.Drawing.Size(38, 21)
        Me.txtFacilityLongSec.TabIndex = 20
        Me.txtFacilityLongSec.Text = ""
        Me.txtFacilityLongSec.WordWrap = False
        '
        'txtFacilityLatMin
        '
        Me.txtFacilityLatMin.AcceptsTab = True
        Me.txtFacilityLatMin.AutoSize = False
        Me.txtFacilityLatMin.Enabled = False
        Me.txtFacilityLatMin.Location = New System.Drawing.Point(736, 56)
        Me.txtFacilityLatMin.MaxLength = 2
        Me.txtFacilityLatMin.Name = "txtFacilityLatMin"
        Me.txtFacilityLatMin.Size = New System.Drawing.Size(28, 21)
        Me.txtFacilityLatMin.TabIndex = 16
        Me.txtFacilityLatMin.Text = ""
        Me.txtFacilityLatMin.WordWrap = False
        '
        'txtFacilityLongMin
        '
        Me.txtFacilityLongMin.AcceptsTab = True
        Me.txtFacilityLongMin.AutoSize = False
        Me.txtFacilityLongMin.Enabled = False
        Me.txtFacilityLongMin.Location = New System.Drawing.Point(736, 80)
        Me.txtFacilityLongMin.MaxLength = 2
        Me.txtFacilityLongMin.Name = "txtFacilityLongMin"
        Me.txtFacilityLongMin.Size = New System.Drawing.Size(28, 21)
        Me.txtFacilityLongMin.TabIndex = 19
        Me.txtFacilityLongMin.Text = ""
        Me.txtFacilityLongMin.WordWrap = False
        '
        'lblFacilityLongMin
        '
        Me.lblFacilityLongMin.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLongMin.Location = New System.Drawing.Point(760, 80)
        Me.lblFacilityLongMin.Name = "lblFacilityLongMin"
        Me.lblFacilityLongMin.Size = New System.Drawing.Size(8, 23)
        Me.lblFacilityLongMin.TabIndex = 1018
        Me.lblFacilityLongMin.Text = "'"
        '
        'lblFacilityLongSec
        '
        Me.lblFacilityLongSec.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLongSec.Location = New System.Drawing.Point(808, 80)
        Me.lblFacilityLongSec.Name = "lblFacilityLongSec"
        Me.lblFacilityLongSec.Size = New System.Drawing.Size(10, 16)
        Me.lblFacilityLongSec.TabIndex = 1017
        Me.lblFacilityLongSec.Text = """"
        '
        'lblFacilityLatMin
        '
        Me.lblFacilityLatMin.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLatMin.Location = New System.Drawing.Point(760, 56)
        Me.lblFacilityLatMin.Name = "lblFacilityLatMin"
        Me.lblFacilityLatMin.Size = New System.Drawing.Size(8, 23)
        Me.lblFacilityLatMin.TabIndex = 1016
        Me.lblFacilityLatMin.Text = "'"
        '
        'lblFacilityLatSec
        '
        Me.lblFacilityLatSec.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLatSec.Location = New System.Drawing.Point(808, 56)
        Me.lblFacilityLatSec.Name = "lblFacilityLatSec"
        Me.lblFacilityLatSec.Size = New System.Drawing.Size(10, 16)
        Me.lblFacilityLatSec.TabIndex = 1015
        Me.lblFacilityLatSec.Text = """"
        '
        'lblFacilityLongDegree
        '
        Me.lblFacilityLongDegree.Font = New System.Drawing.Font("Symbol", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLongDegree.Location = New System.Drawing.Point(720, 72)
        Me.lblFacilityLongDegree.Name = "lblFacilityLongDegree"
        Me.lblFacilityLongDegree.Size = New System.Drawing.Size(16, 16)
        Me.lblFacilityLongDegree.TabIndex = 1010
        Me.lblFacilityLongDegree.Text = "o"
        '
        'mskTxtFacilityFax
        '
        Me.mskTxtFacilityFax.ContainingControl = Me
        Me.mskTxtFacilityFax.Location = New System.Drawing.Point(88, 186)
        Me.mskTxtFacilityFax.Name = "mskTxtFacilityFax"
        Me.mskTxtFacilityFax.OcxState = CType(resources.GetObject("mskTxtFacilityFax.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtFacilityFax.Size = New System.Drawing.Size(96, 23)
        Me.mskTxtFacilityFax.TabIndex = 3
        '
        'mskTxtFacilityPhone
        '
        Me.mskTxtFacilityPhone.ContainingControl = Me
        Me.mskTxtFacilityPhone.Location = New System.Drawing.Point(88, 162)
        Me.mskTxtFacilityPhone.Name = "mskTxtFacilityPhone"
        Me.mskTxtFacilityPhone.OcxState = CType(resources.GetObject("mskTxtFacilityPhone.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtFacilityPhone.Size = New System.Drawing.Size(96, 23)
        Me.mskTxtFacilityPhone.TabIndex = 2
        '
        'txtFacilityAddress
        '
        Me.txtFacilityAddress.Location = New System.Drawing.Point(88, 56)
        Me.txtFacilityAddress.Multiline = True
        Me.txtFacilityAddress.Name = "txtFacilityAddress"
        Me.txtFacilityAddress.ReadOnly = True
        Me.txtFacilityAddress.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
        Me.txtFacilityAddress.Size = New System.Drawing.Size(224, 104)
        Me.txtFacilityAddress.TabIndex = 1
        Me.txtFacilityAddress.Text = ""
        Me.txtFacilityAddress.WordWrap = False
        '
        'lblFacilitySIC
        '
        Me.lblFacilitySIC.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilitySIC.Location = New System.Drawing.Point(320, 124)
        Me.lblFacilitySIC.Name = "lblFacilitySIC"
        Me.lblFacilitySIC.Size = New System.Drawing.Size(80, 23)
        Me.lblFacilitySIC.TabIndex = 150
        Me.lblFacilitySIC.Text = "SIC:"
        Me.lblFacilitySIC.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblFacilityFax
        '
        Me.lblFacilityFax.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityFax.Location = New System.Drawing.Point(8, 186)
        Me.lblFacilityFax.Name = "lblFacilityFax"
        Me.lblFacilityFax.Size = New System.Drawing.Size(72, 23)
        Me.lblFacilityFax.TabIndex = 147
        Me.lblFacilityFax.Text = "Fax:"
        '
        'dtPickFacilityRecvd
        '
        Me.dtPickFacilityRecvd.Checked = False
        Me.dtPickFacilityRecvd.Enabled = False
        Me.dtPickFacilityRecvd.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickFacilityRecvd.Location = New System.Drawing.Point(424, 11)
        Me.dtPickFacilityRecvd.Name = "dtPickFacilityRecvd"
        Me.dtPickFacilityRecvd.ShowCheckBox = True
        Me.dtPickFacilityRecvd.Size = New System.Drawing.Size(104, 21)
        Me.dtPickFacilityRecvd.TabIndex = 5
        '
        'lblDateReceived
        '
        Me.lblDateReceived.Location = New System.Drawing.Point(320, 9)
        Me.lblDateReceived.Name = "lblDateReceived"
        Me.lblDateReceived.TabIndex = 145
        Me.lblDateReceived.Text = "Date Received:"
        Me.lblDateReceived.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkSignatureofNF
        '
        Me.chkSignatureofNF.Enabled = False
        Me.chkSignatureofNF.Location = New System.Drawing.Point(136, 224)
        Me.chkSignatureofNF.Name = "chkSignatureofNF"
        Me.chkSignatureofNF.Size = New System.Drawing.Size(16, 16)
        Me.chkSignatureofNF.TabIndex = 4
        Me.chkSignatureofNF.Text = "CheckBox5"
        '
        'lblFacilitySigOnNF
        '
        Me.lblFacilitySigOnNF.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilitySigOnNF.Location = New System.Drawing.Point(8, 224)
        Me.lblFacilitySigOnNF.Name = "lblFacilitySigOnNF"
        Me.lblFacilitySigOnNF.Size = New System.Drawing.Size(120, 23)
        Me.lblFacilitySigOnNF.TabIndex = 127
        Me.lblFacilitySigOnNF.Text = "Signature Received:"
        '
        'lblFacilityFuelBrand
        '
        Me.lblFacilityFuelBrand.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityFuelBrand.Location = New System.Drawing.Point(320, 152)
        Me.lblFacilityFuelBrand.Name = "lblFacilityFuelBrand"
        Me.lblFacilityFuelBrand.Size = New System.Drawing.Size(96, 23)
        Me.lblFacilityFuelBrand.TabIndex = 125
        Me.lblFacilityFuelBrand.Text = "Fuel Brand:"
        Me.lblFacilityFuelBrand.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblFacilityStatusValue
        '
        Me.lblFacilityStatusValue.BackColor = System.Drawing.Color.Transparent
        Me.lblFacilityStatusValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblFacilityStatusValue.Enabled = False
        Me.lblFacilityStatusValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityStatusValue.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblFacilityStatusValue.Location = New System.Drawing.Point(424, 40)
        Me.lblFacilityStatusValue.Name = "lblFacilityStatusValue"
        Me.lblFacilityStatusValue.Size = New System.Drawing.Size(120, 16)
        Me.lblFacilityStatusValue.TabIndex = 124
        '
        'lblFacilityStatus
        '
        Me.lblFacilityStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityStatus.Location = New System.Drawing.Point(320, 32)
        Me.lblFacilityStatus.Name = "lblFacilityStatus"
        Me.lblFacilityStatus.Size = New System.Drawing.Size(88, 23)
        Me.lblFacilityStatus.TabIndex = 123
        Me.lblFacilityStatus.Text = "Facility Status:"
        Me.lblFacilityStatus.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtFacilityLongDegree
        '
        Me.txtFacilityLongDegree.AcceptsTab = True
        Me.txtFacilityLongDegree.AutoSize = False
        Me.txtFacilityLongDegree.Enabled = False
        Me.txtFacilityLongDegree.Location = New System.Drawing.Point(688, 80)
        Me.txtFacilityLongDegree.MaxLength = 3
        Me.txtFacilityLongDegree.Name = "txtFacilityLongDegree"
        Me.txtFacilityLongDegree.Size = New System.Drawing.Size(32, 21)
        Me.txtFacilityLongDegree.TabIndex = 18
        Me.txtFacilityLongDegree.Text = ""
        Me.txtFacilityLongDegree.WordWrap = False
        '
        'txtFacilityLatDegree
        '
        Me.txtFacilityLatDegree.AcceptsTab = True
        Me.txtFacilityLatDegree.AutoSize = False
        Me.txtFacilityLatDegree.Enabled = False
        Me.txtFacilityLatDegree.Location = New System.Drawing.Point(688, 56)
        Me.txtFacilityLatDegree.MaxLength = 3
        Me.txtFacilityLatDegree.Name = "txtFacilityLatDegree"
        Me.txtFacilityLatDegree.Size = New System.Drawing.Size(32, 21)
        Me.txtFacilityLatDegree.TabIndex = 15
        Me.txtFacilityLatDegree.Text = ""
        Me.txtFacilityLatDegree.WordWrap = False
        '
        'lblFacilityLongitude
        '
        Me.lblFacilityLongitude.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLongitude.Location = New System.Drawing.Point(600, 80)
        Me.lblFacilityLongitude.Name = "lblFacilityLongitude"
        Me.lblFacilityLongitude.Size = New System.Drawing.Size(96, 23)
        Me.lblFacilityLongitude.TabIndex = 111
        Me.lblFacilityLongitude.Text = "Longitude:"
        '
        'lblFacilityLatitude
        '
        Me.lblFacilityLatitude.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLatitude.Location = New System.Drawing.Point(600, 56)
        Me.lblFacilityLatitude.Name = "lblFacilityLatitude"
        Me.lblFacilityLatitude.Size = New System.Drawing.Size(96, 23)
        Me.lblFacilityLatitude.TabIndex = 110
        Me.lblFacilityLatitude.Text = "Latitude:"
        '
        'lblFacilityType
        '
        Me.lblFacilityType.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityType.Location = New System.Drawing.Point(600, 32)
        Me.lblFacilityType.Name = "lblFacilityType"
        Me.lblFacilityType.Size = New System.Drawing.Size(80, 23)
        Me.lblFacilityType.TabIndex = 106
        Me.lblFacilityType.Text = "Facility Type:"
        '
        'txtFacilityAIID
        '
        Me.txtFacilityAIID.AcceptsTab = True
        Me.txtFacilityAIID.AutoSize = False
        Me.txtFacilityAIID.Enabled = False
        Me.txtFacilityAIID.Location = New System.Drawing.Point(688, 8)
        Me.txtFacilityAIID.Name = "txtFacilityAIID"
        Me.txtFacilityAIID.ReadOnly = True
        Me.txtFacilityAIID.Size = New System.Drawing.Size(136, 21)
        Me.txtFacilityAIID.TabIndex = 13
        Me.txtFacilityAIID.Text = ""
        Me.txtFacilityAIID.WordWrap = False
        '
        'lblfacilityAIID
        '
        Me.lblfacilityAIID.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblfacilityAIID.Location = New System.Drawing.Point(600, 8)
        Me.lblfacilityAIID.Name = "lblfacilityAIID"
        Me.lblfacilityAIID.Size = New System.Drawing.Size(80, 23)
        Me.lblfacilityAIID.TabIndex = 104
        Me.lblfacilityAIID.Text = "Facility AIID:"
        '
        'lblFacilityIDValue
        '
        Me.lblFacilityIDValue.Enabled = False
        Me.lblFacilityIDValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityIDValue.Location = New System.Drawing.Point(88, 8)
        Me.lblFacilityIDValue.Name = "lblFacilityIDValue"
        Me.lblFacilityIDValue.Size = New System.Drawing.Size(96, 23)
        Me.lblFacilityIDValue.TabIndex = 103
        '
        'lblFacilityID
        '
        Me.lblFacilityID.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityID.Location = New System.Drawing.Point(8, 8)
        Me.lblFacilityID.Name = "lblFacilityID"
        Me.lblFacilityID.Size = New System.Drawing.Size(96, 23)
        Me.lblFacilityID.TabIndex = 102
        Me.lblFacilityID.Text = "Facility ID:"
        '
        'lblFacilityPhone
        '
        Me.lblFacilityPhone.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityPhone.Location = New System.Drawing.Point(8, 162)
        Me.lblFacilityPhone.Name = "lblFacilityPhone"
        Me.lblFacilityPhone.Size = New System.Drawing.Size(72, 23)
        Me.lblFacilityPhone.TabIndex = 98
        Me.lblFacilityPhone.Text = "Phone:"
        '
        'txtFacilityName
        '
        Me.txtFacilityName.AcceptsTab = True
        Me.txtFacilityName.AutoSize = False
        Me.txtFacilityName.Location = New System.Drawing.Point(88, 32)
        Me.txtFacilityName.Name = "txtFacilityName"
        Me.txtFacilityName.ReadOnly = True
        Me.txtFacilityName.Size = New System.Drawing.Size(224, 21)
        Me.txtFacilityName.TabIndex = 0
        Me.txtFacilityName.Text = ""
        Me.txtFacilityName.WordWrap = False
        '
        'lblFacilityAddress
        '
        Me.lblFacilityAddress.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityAddress.Location = New System.Drawing.Point(8, 56)
        Me.lblFacilityAddress.Name = "lblFacilityAddress"
        Me.lblFacilityAddress.Size = New System.Drawing.Size(96, 23)
        Me.lblFacilityAddress.TabIndex = 90
        Me.lblFacilityAddress.Text = "Address:"
        '
        'lblFacilityName
        '
        Me.lblFacilityName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityName.Location = New System.Drawing.Point(8, 32)
        Me.lblFacilityName.Name = "lblFacilityName"
        Me.lblFacilityName.Size = New System.Drawing.Size(96, 23)
        Me.lblFacilityName.TabIndex = 89
        Me.lblFacilityName.Text = "Facility Name:"
        '
        'lblFacilityLatDegree
        '
        Me.lblFacilityLatDegree.Font = New System.Drawing.Font("Symbol", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLatDegree.Location = New System.Drawing.Point(720, 48)
        Me.lblFacilityLatDegree.Name = "lblFacilityLatDegree"
        Me.lblFacilityLatDegree.Size = New System.Drawing.Size(16, 16)
        Me.lblFacilityLatDegree.TabIndex = 1009
        Me.lblFacilityLatDegree.Text = "o"
        '
        'btnInspHistory
        '
        Me.btnInspHistory.Location = New System.Drawing.Point(864, 224)
        Me.btnInspHistory.Name = "btnInspHistory"
        Me.btnInspHistory.Size = New System.Drawing.Size(80, 35)
        Me.btnInspHistory.TabIndex = 1050
        Me.btnInspHistory.Text = "Inspection History"
        Me.btnInspHistory.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'txtFacilitySIC
        '
        Me.txtFacilitySIC.BackColor = System.Drawing.SystemColors.Control
        Me.txtFacilitySIC.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFacilitySIC.Enabled = False
        Me.txtFacilitySIC.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFacilitySIC.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.txtFacilitySIC.Location = New System.Drawing.Point(424, 128)
        Me.txtFacilitySIC.Name = "txtFacilitySIC"
        Me.txtFacilitySIC.Size = New System.Drawing.Size(120, 16)
        Me.txtFacilitySIC.TabIndex = 1038
        '
        'lblInViolationValue
        '
        Me.lblInViolationValue.BackColor = System.Drawing.SystemColors.Control
        Me.lblInViolationValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblInViolationValue.Enabled = False
        Me.lblInViolationValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInViolationValue.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblInViolationValue.Location = New System.Drawing.Point(416, 240)
        Me.lblInViolationValue.Name = "lblInViolationValue"
        Me.lblInViolationValue.Size = New System.Drawing.Size(152, 17)
        Me.lblInViolationValue.TabIndex = 1038
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(320, 232)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(88, 23)
        Me.Label4.TabIndex = 1037
        Me.Label4.Text = "In Violation:"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'tbPageOwnerCitations
        '
        Me.tbPageOwnerCitations.Controls.Add(Me.ugOCE2)
        Me.tbPageOwnerCitations.Controls.Add(Me.pnlOwnerCitationsHead)
        Me.tbPageOwnerCitations.Location = New System.Drawing.Point(4, 22)
        Me.tbPageOwnerCitations.Name = "tbPageOwnerCitations"
        Me.tbPageOwnerCitations.Size = New System.Drawing.Size(1020, 644)
        Me.tbPageOwnerCitations.TabIndex = 9
        Me.tbPageOwnerCitations.Text = "Owner Citations"
        '
        'ugOCE2
        '
        Me.ugOCE2.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugOCE2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugOCE2.Location = New System.Drawing.Point(0, 32)
        Me.ugOCE2.Name = "ugOCE2"
        Me.ugOCE2.Size = New System.Drawing.Size(1020, 612)
        Me.ugOCE2.TabIndex = 12
        '
        'pnlOwnerCitationsHead
        '
        Me.pnlOwnerCitationsHead.Controls.Add(Me.btnEnforceViewEnforceHistory1)
        Me.pnlOwnerCitationsHead.Controls.Add(Me.btnExpandOwnerCitationAll)
        Me.pnlOwnerCitationsHead.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlOwnerCitationsHead.Location = New System.Drawing.Point(0, 0)
        Me.pnlOwnerCitationsHead.Name = "pnlOwnerCitationsHead"
        Me.pnlOwnerCitationsHead.Size = New System.Drawing.Size(1020, 32)
        Me.pnlOwnerCitationsHead.TabIndex = 2
        '
        'btnEnforceViewEnforceHistory1
        '
        Me.btnEnforceViewEnforceHistory1.Location = New System.Drawing.Point(112, 5)
        Me.btnEnforceViewEnforceHistory1.Name = "btnEnforceViewEnforceHistory1"
        Me.btnEnforceViewEnforceHistory1.Size = New System.Drawing.Size(160, 22)
        Me.btnEnforceViewEnforceHistory1.TabIndex = 148
        Me.btnEnforceViewEnforceHistory1.Text = "View Enforcement History"
        '
        'btnExpandOwnerCitationAll
        '
        Me.btnExpandOwnerCitationAll.Location = New System.Drawing.Point(7, 5)
        Me.btnExpandOwnerCitationAll.Name = "btnExpandOwnerCitationAll"
        Me.btnExpandOwnerCitationAll.Size = New System.Drawing.Size(96, 22)
        Me.btnExpandOwnerCitationAll.TabIndex = 147
        Me.btnExpandOwnerCitationAll.Text = "Expand All"
        Me.btnExpandOwnerCitationAll.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'tbPageFacilityCitations
        '
        Me.tbPageFacilityCitations.Controls.Add(Me.ugFCE2)
        Me.tbPageFacilityCitations.Controls.Add(Me.pnlFacilityCitationHead)
        Me.tbPageFacilityCitations.Location = New System.Drawing.Point(4, 22)
        Me.tbPageFacilityCitations.Name = "tbPageFacilityCitations"
        Me.tbPageFacilityCitations.Size = New System.Drawing.Size(1020, 644)
        Me.tbPageFacilityCitations.TabIndex = 10
        Me.tbPageFacilityCitations.Text = "Facility Citations"
        '
        'ugFCE2
        '
        Me.ugFCE2.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugFCE2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugFCE2.Location = New System.Drawing.Point(0, 32)
        Me.ugFCE2.Name = "ugFCE2"
        Me.ugFCE2.Size = New System.Drawing.Size(1020, 612)
        Me.ugFCE2.TabIndex = 13
        '
        'pnlFacilityCitationHead
        '
        Me.pnlFacilityCitationHead.Controls.Add(Me.btnExpandFacilityCitationAll)
        Me.pnlFacilityCitationHead.Controls.Add(Me.btnInspHistory1)
        Me.pnlFacilityCitationHead.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlFacilityCitationHead.Location = New System.Drawing.Point(0, 0)
        Me.pnlFacilityCitationHead.Name = "pnlFacilityCitationHead"
        Me.pnlFacilityCitationHead.Size = New System.Drawing.Size(1020, 32)
        Me.pnlFacilityCitationHead.TabIndex = 3
        '
        'btnExpandFacilityCitationAll
        '
        Me.btnExpandFacilityCitationAll.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExpandFacilityCitationAll.Location = New System.Drawing.Point(7, 5)
        Me.btnExpandFacilityCitationAll.Name = "btnExpandFacilityCitationAll"
        Me.btnExpandFacilityCitationAll.Size = New System.Drawing.Size(96, 22)
        Me.btnExpandFacilityCitationAll.TabIndex = 147
        Me.btnExpandFacilityCitationAll.Text = "Expand All"
        Me.btnExpandFacilityCitationAll.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'btnInspHistory1
        '
        Me.btnInspHistory1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnInspHistory1.Location = New System.Drawing.Point(110, 5)
        Me.btnInspHistory1.Name = "btnInspHistory1"
        Me.btnInspHistory1.Size = New System.Drawing.Size(114, 22)
        Me.btnInspHistory1.TabIndex = 147
        Me.btnInspHistory1.Text = "Inspection History"
        Me.btnInspHistory1.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'tbPageSummary
        '
        Me.tbPageSummary.AutoScroll = True
        Me.tbPageSummary.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.tbPageSummary.Controls.Add(Me.pnlOwnerSummaryDetails)
        Me.tbPageSummary.Controls.Add(Me.pnlOwnerSummaryHeader)
        Me.tbPageSummary.Controls.Add(Me.Panel12)
        Me.tbPageSummary.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbPageSummary.Location = New System.Drawing.Point(4, 22)
        Me.tbPageSummary.Name = "tbPageSummary"
        Me.tbPageSummary.Size = New System.Drawing.Size(1020, 644)
        Me.tbPageSummary.TabIndex = 0
        Me.tbPageSummary.Text = "Owner Summary"
        Me.tbPageSummary.Visible = False
        '
        'pnlOwnerSummaryDetails
        '
        Me.pnlOwnerSummaryDetails.Controls.Add(Me.UCOwnerSummary)
        Me.pnlOwnerSummaryDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlOwnerSummaryDetails.Location = New System.Drawing.Point(0, 16)
        Me.pnlOwnerSummaryDetails.Name = "pnlOwnerSummaryDetails"
        Me.pnlOwnerSummaryDetails.Size = New System.Drawing.Size(1008, 624)
        Me.pnlOwnerSummaryDetails.TabIndex = 8
        '
        'UCOwnerSummary
        '
        Me.UCOwnerSummary.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UCOwnerSummary.Location = New System.Drawing.Point(0, 0)
        Me.UCOwnerSummary.Name = "UCOwnerSummary"
        Me.UCOwnerSummary.Size = New System.Drawing.Size(1008, 624)
        Me.UCOwnerSummary.TabIndex = 0
        '
        'pnlOwnerSummaryHeader
        '
        Me.pnlOwnerSummaryHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlOwnerSummaryHeader.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlOwnerSummaryHeader.Location = New System.Drawing.Point(0, 0)
        Me.pnlOwnerSummaryHeader.Name = "pnlOwnerSummaryHeader"
        Me.pnlOwnerSummaryHeader.Size = New System.Drawing.Size(1008, 16)
        Me.pnlOwnerSummaryHeader.TabIndex = 7
        '
        'Panel12
        '
        Me.Panel12.Dock = System.Windows.Forms.DockStyle.Right
        Me.Panel12.DockPadding.Left = 10
        Me.Panel12.Location = New System.Drawing.Point(1008, 0)
        Me.Panel12.Name = "Panel12"
        Me.Panel12.Size = New System.Drawing.Size(8, 640)
        Me.Panel12.TabIndex = 6
        '
        'CandE
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1028, 694)
        Me.Controls.Add(Me.tbCntrlCandE)
        Me.Controls.Add(Me.pnlTop)
        Me.Name = "CandE"
        Me.Text = "C & E"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.pnlTop.ResumeLayout(False)
        Me.tbCntrlCandE.ResumeLayout(False)
        Me.tbPageOwnerDetail.ResumeLayout(False)
        Me.pnlOwnerBottom.ResumeLayout(False)
        Me.tbCtrlOwner.ResumeLayout(False)
        Me.tbPageOwnerFacilities.ResumeLayout(False)
        CType(Me.ugFacilityList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlOwnerFacilityBottom.ResumeLayout(False)
        Me.tbPageOCEs.ResumeLayout(False)
        CType(Me.ugOCEs, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlOwnerOCEsBottom.ResumeLayout(False)
        Me.tbPageOwnerDocuments.ResumeLayout(False)
        Me.tbPageOwnerContactList.ResumeLayout(False)
        Me.pnlOwnerContactContainer.ResumeLayout(False)
        CType(Me.ugOwnerContacts, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlOwnerContactButtons.ResumeLayout(False)
        Me.pnlOwnerContactHeader.ResumeLayout(False)
        Me.pnlOwnerDetail.ResumeLayout(False)
        Me.pnlOwnerButtons.ResumeLayout(False)
        CType(Me.mskTxtOwnerFax, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtOwnerPhone2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtOwnerPhone, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbPageFacilityDetail.ResumeLayout(False)
        Me.pnlFacilityBottom.ResumeLayout(False)
        Me.tbCntrlFacility.ResumeLayout(False)
        Me.tbPageFCEs.ResumeLayout(False)
        CType(Me.ugFCEs, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlFacilityLustButton.ResumeLayout(False)
        Me.tbPageFacilityDocuments.ResumeLayout(False)
        Me.tbPageFacilityContactList.ResumeLayout(False)
        Me.pnlFacilityContactContainer.ResumeLayout(False)
        CType(Me.ugFacilityContacts, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlFacilityContactBottom.ResumeLayout(False)
        Me.pnlFacilityContactHeader.ResumeLayout(False)
        Me.pnl_FacilityDetail.ResumeLayout(False)
        CType(Me.mskTxtFacilityFax, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtFacilityPhone, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbPageOwnerCitations.ResumeLayout(False)
        CType(Me.ugOCE2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlOwnerCitationsHead.ResumeLayout(False)
        Me.tbPageFacilityCitations.ResumeLayout(False)
        CType(Me.ugFCE2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlFacilityCitationHead.ResumeLayout(False)
        Me.tbPageSummary.ResumeLayout(False)
        Me.pnlOwnerSummaryDetails.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub txtFacilityAddress_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFacilityAddress.DoubleClick
        Try

            Dim addressForm As Address
            If (pOwn.Facilities.ID = 0 Or pOwn.Facilities.ID <> nFacilityID) And nFacilityID > 0 Then
                UIUtilsGen.PopulateFacilityInfo(Me, pOwn.OwnerInfo, pOwn.Facilities, nFacilityID)
            End If
            Address.EditAddress(addressForm, pOwn.Facilities.ID, pOwn.Facilities.FacilityAddresses, "Facility", UIUtilsGen.ModuleID.CAE, txtFacilityAddress, UIUtilsGen.EntityTypes.Facility, True)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtOwnerAddress_DblClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOwnerAddress.DoubleClick
        Try

            Dim addressForm As Address
            Address.EditAddress(addressForm, pOwn.ID, pOwn.Addresses, "Owner", UIUtilsGen.ModuleID.CAE, txtOwnerAddress, UIUtilsGen.EntityTypes.Owner)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub




    Private Sub ExpandAll(ByVal bol As Boolean, ByRef ug As Infragistics.Win.UltraWinGrid.UltraGrid, Optional ByRef btn As Button = Nothing)
        If bol Then
            If Not btn Is Nothing Then btn.Text = "Collapse All"
            ug.Rows.ExpandAll(True)
        Else
            If Not btn Is Nothing Then btn.Text = "Expand All"
            ug.Rows.CollapseAll(True)
        End If
    End Sub
    Private Sub SetupTabs()
        Try
            Select Case tbCntrlCandE.SelectedTab.Name
                Case tbPageOwnerDetail.Name
                    PopulateOwner(lblOwnerIDValue.Text)
                    UCOwnerDocuments.EntityID = Integer.Parse(lblOwnerIDValue.Text)
                    Me.Text = "C " + "&" + " E - Owner Detail - " + lblOwnerIDValue.Text + "(" & txtOwnerName.Text & ")"
                    MC.lblFacilityInfo.Text = String.Empty
                    MC.lblFacilityID.Text = String.Empty
                Case tbPageFacilityDetail.Name
                    If Not ugFacilityList Is Nothing Then
                        If ugFacilityList.Rows.Count = 0 Then
                            MsgBox("There are no facilities for Owner")
                            tbCntrlCandE.SelectedTab = tbPageOwnerDetail
                        Else
                            If lblFacilityIDValue.Text = String.Empty Then
                                If ugFacilityList.ActiveRow Is Nothing Then
                                    lblFacilityIDValue.Text = ugFacilityList.Rows(0).Cells("FACILITYID").Value.ToString
                                Else
                                    lblFacilityIDValue.Text = ugFacilityList.ActiveRow.Cells("FACILITYID").Value.ToString
                                End If
                            End If
                            PopulateFacility(lblFacilityIDValue.Text)
                            nFacilityID = lblFacilityIDValue.Text
                            Me.Text = "C " + "&" + " E - Facility Detail - " + lblFacilityIDValue.Text + "(" & IIf(txtFacilityName.Text = String.Empty, "New", txtFacilityName.Text) & ")"
                        End If
                    Else
                        MsgBox("There are no facilities for Owner")
                        tbCntrlCandE.SelectedTab = tbPageOwnerDetail
                    End If
                Case tbPageOwnerCitations.Name
                    PopulateOCE(lblOwnerIDValue.Text)
                    Me.Text = "C " + "&" + " E - Owner Citations - " + lblOwnerIDValue.Text + "(" & txtOwnerName.Text & ")"
                    If bolFrmActivated Then
                        MC.FlagsChanged(lblOwnerIDValue.Text, UIUtilsGen.EntityTypes.Owner, "C " + "&" + " E", Me.Text)
                    End If
                Case tbPageFacilityCitations.Name
                    If Not ugFacilityList Is Nothing Then
                        If ugFacilityList.Rows.Count = 0 Then
                            MsgBox("There are no facilities for Owner")
                            tbCntrlCandE.SelectedTab = tbPageOwnerDetail
                        Else
                            If lblFacilityIDValue.Text = String.Empty Then
                                PopulateFacility(ugFacilityList.Rows(0).Cells("FACILITYID").Value)
                            Else
                                PopulateFCE(lblFacilityIDValue.Text)
                            End If
                            Me.Text = "C " + "&" + " E - Facility Citations - " + lblFacilityIDValue.Text + "(" & IIf(txtFacilityName.Text = String.Empty, "New", txtFacilityName.Text) & ")"
                            If bolFrmActivated Then
                                MC.FlagsChanged(lblFacilityIDValue.Text, UIUtilsGen.EntityTypes.Facility, "C " + "&" + " E", Me.Text)
                            End If
                        End If
                    Else
                        MsgBox("There are no facilities for Owner")
                        tbCntrlCandE.SelectedTab = tbPageOwnerDetail
                    End If
                Case tbPageSummary.Name
                    Me.Text = "C " + "&" + " E - Owner Summary (" & txtOwnerName.Text & ")"
                    UIUtilsGen.PopulateOwnerSummary(pOwn, Me)
                Case Else
                    Exit Sub
            End Select
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub tbCntrlCandE_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbCntrlCandE.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            SetupTabs()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

#Region "General"
    Private Sub ShowError(ByVal ex As Exception)
        Dim MyErr As New ErrorReport(ex)
        MyErr.ShowDialog()
    End Sub
#End Region


#Region "Owner / Owner Citations Tab"
    Friend Sub PopulateOwner(ByVal ownerID As Integer)
        Try
            UIUtilsGen.PopulateOwnerInfo(ownerID, pOwn, Me)
            PopulateOwnerFacilities(ownerID)
            If tbCtrlOwner.SelectedTab.Name = tbPageOCEs.Name Then
                PopulateOCE(ownerID)
            ElseIf tbCtrlOwner.SelectedTab.Name = tbPageOwnerDocuments.Name Then
                UCOwnerDocuments.LoadDocumentsGrid(ownerID, UIUtilsGen.EntityTypes.Owner, UIUtilsGen.ModuleID.CAE)
            ElseIf tbCtrlOwner.SelectedTab.Name = tbPageOwnerContactList.Name Then
                LoadContacts(ugOwnerContacts, ownerID, UIUtilsGen.EntityTypes.Owner)
            End If
            If bolFrmActivated Then
                MC.FlagsChanged(ownerID, UIUtilsGen.EntityTypes.Owner, "C " + "&" + " E", Me.Text)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Friend Sub PopulateOCE(ByVal ownerID As Integer)
        Try
            Dim ds As DataSet = pOCE.GetEnforcements(ownerID, False)
            ugOCEs.DataSource = ds
            ugOCEs.DrawFilter = rp
            ExpandAll(False, ugOCEs, )
            ugOCE2.DataSource = ds
            ugOCE2.DrawFilter = rp
            ExpandAll(False, ugOCE2, btnExpandOwnerCitationAll)
            lblNoOfOCEsValue.Text = ugOCEs.Rows.Count.ToString


            RemoveHandler ugOCEs.DoubleClick, AddressOf JumpToCNEManagement
            AddHandler ugOCEs.DoubleClick, AddressOf JumpToCNEManagement

            RemoveHandler ugOCE2.DoubleClick, AddressOf JumpToCNEManagement
            AddHandler ugOCE2.DoubleClick, AddressOf JumpToCNEManagement



        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub



    Friend Sub PopulateOwnerFacilities(ByVal ownerID As Integer)
        Try
            ugFacilityList.DataSource = pOwn.GetFacilitiesCAESummaryTable
            ugFacilityList.DrawFilter = rp
            lblNoOfFacilitiesValue.Text = ugFacilityList.Rows.Count.ToString
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ViewEnforcementHistory()
        Try
            frmEnforcementHistory = New EnforcementHistory(True, pOwn.ID, txtOwnerName.Text.Trim, pOCE)
            frmEnforcementHistory.ShowDialog()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub SetugRowComboValue(ByRef ug As Infragistics.Win.UltraWinGrid.UltraGridRow)
        Try
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
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub SetupOCE(ByRef ug As Infragistics.Win.UltraWinGrid.UltraGrid)
        Try
            ug.DisplayLayout.UseFixedHeaders = True
            ug.DisplayLayout.Override.FixedHeaderIndicator = Infragistics.Win.UltraWinGrid.FixedHeaderIndicator.None

            ' cannot have fixed header if row header is split into two rows
            'ug.DisplayLayout.Bands(0).Columns("SELECTED").Header.Fixed = True
            'ug.DisplayLayout.Bands(0).Columns("OWNERNAME").Header.Fixed = True
            'ug.DisplayLayout.Bands(0).Columns("ENSITE ID").Header.Fixed = True
            'ug.DisplayLayout.Bands(0).Columns("FILLER1").Header.Fixed = True

            ug.DisplayLayout.Bands(1).Columns("SELECTED").Header.Fixed = True
            ug.DisplayLayout.Bands(1).Columns("OWNER_ID").Header.Fixed = True
            ug.DisplayLayout.Bands(1).Columns("FACILITY_ID").Header.Fixed = True
            ug.DisplayLayout.Bands(1).Columns("FACILITY").Header.Fixed = True

            ug.DisplayLayout.Bands(0).Columns("ENSITE ID").MaskInput = "nnnnnnnnn"
            ug.DisplayLayout.Bands(0).Columns("ENSITE ID").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
            ug.DisplayLayout.Bands(0).Columns("ENSITE ID").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw

            ug.DisplayLayout.Bands(0).Columns("POLICY AMOUNT").MaskInput = "$nnn,nnn,nnn.nn"
            ug.DisplayLayout.Bands(0).Columns("POLICY AMOUNT").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
            ug.DisplayLayout.Bands(0).Columns("POLICY AMOUNT").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw

            ug.DisplayLayout.Bands(0).Columns("OVERRIDE AMOUNT").MaskInput = "$nnn,nnn,nnn.nn"
            ug.DisplayLayout.Bands(0).Columns("OVERRIDE AMOUNT").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
            ug.DisplayLayout.Bands(0).Columns("OVERRIDE AMOUNT").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw

            ug.DisplayLayout.Bands(0).Columns("SETTLEMENT AMOUNT").MaskInput = "$nnn,nnn,nnn.nn"
            ug.DisplayLayout.Bands(0).Columns("SETTLEMENT AMOUNT").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
            ug.DisplayLayout.Bands(0).Columns("SETTLEMENT AMOUNT").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw

            ug.DisplayLayout.Bands(0).Columns("PAID AMOUNT").MaskInput = "$nnn,nnn,nnn.nn"
            ug.DisplayLayout.Bands(0).Columns("PAID AMOUNT").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
            ug.DisplayLayout.Bands(0).Columns("PAID AMOUNT").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw

            ug.DisplayLayout.Bands(0).Columns("AGREED ORDER #").MaskInput = "&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            '     ug.DisplayLayout.Bands(0).Columns("AGREED ORDER #").MaskInput = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"
            ug.DisplayLayout.Bands(0).Columns("AGREED ORDER #").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
            ug.DisplayLayout.Bands(0).Columns("AGREED ORDER #").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw

            ug.DisplayLayout.Bands(0).Columns("ADMINISTRATIVE ORDER #").MaskInput = "&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&"
            ug.DisplayLayout.Bands(0).Columns("ADMINISTRATIVE ORDER #").MaskDisplayMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.IncludeLiterals
            ug.DisplayLayout.Bands(0).Columns("ADMINISTRATIVE ORDER #").MaskDataMode = Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw

            ug.DisplayLayout.Bands(0).Override.RowAppearance.BackColor = Color.RosyBrown
            ug.DisplayLayout.Bands(0).Override.RowAlternateAppearance.BackColor = Color.PeachPuff
            ug.DisplayLayout.Bands(1).Override.RowAppearance.BackColor = Color.Khaki

            ug.DisplayLayout.Bands(0).Columns("OCE DATE").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Descending

            ug.DisplayLayout.Bands(0).Columns("SELECTED").Hidden = True
            ug.DisplayLayout.Bands(0).Columns("OCE_ID").Hidden = True
            ug.DisplayLayout.Bands(0).Columns("OCE_STATUS").Hidden = True
            ug.DisplayLayout.Bands(0).Columns("OCE_ID").Hidden = True
            ug.DisplayLayout.Bands(0).Columns("OWNER_ID").Hidden = True
            ug.DisplayLayout.Bands(0).Columns("OCE_PATH").Hidden = True
            ug.DisplayLayout.Bands(0).Columns("CITATION").Hidden = True
            ug.DisplayLayout.Bands(0).Columns("CITATION_DUE_DATE").Hidden = True
            ug.DisplayLayout.Bands(0).Columns("WORKSHOP_REQUIRED").Hidden = True
            ug.DisplayLayout.Bands(0).Columns("PENDING_LETTER").Hidden = True
            ug.DisplayLayout.Bands(0).Columns("CREATED_BY").Hidden = True
            ug.DisplayLayout.Bands(0).Columns("DATE_CREATED").Hidden = True
            ug.DisplayLayout.Bands(0).Columns("LAST_EDITED_BY").Hidden = True
            ug.DisplayLayout.Bands(0).Columns("DATE_LAST_EDITED").Hidden = True
            ug.DisplayLayout.Bands(0).Columns("DELETED").Hidden = True
            ug.DisplayLayout.Bands(0).Columns("FILLER1").Hidden = True
            ug.DisplayLayout.Bands(0).Columns("PENDING_LETTER_TEMPLATE_NUM").Hidden = True
            ug.DisplayLayout.Bands(0).Columns("ADDRESS_LINE_ONE").Hidden = True
            ug.DisplayLayout.Bands(0).Columns("ADDRESS_TWO").Hidden = True
            ug.DisplayLayout.Bands(0).Columns("CITY").Hidden = True
            ug.DisplayLayout.Bands(0).Columns("STATE").Hidden = True
            ug.DisplayLayout.Bands(0).Columns("ZIP").Hidden = True
            ug.DisplayLayout.Bands(0).Columns("ORGANIZATION_ID").Hidden = True
            ug.DisplayLayout.Bands(0).Columns("PERSON_ID").Hidden = True

            ug.DisplayLayout.Bands(1).Columns("SELECTED").Hidden = True
            ug.DisplayLayout.Bands(1).Columns("INS_CIT_ID").Hidden = True
            ug.DisplayLayout.Bands(1).Columns("INSPECTION_ID").Hidden = True
            ug.DisplayLayout.Bands(1).Columns("QUESTION_ID").Hidden = True
            ug.DisplayLayout.Bands(1).Columns("OCE_ID").Hidden = True
            ug.DisplayLayout.Bands(1).Columns("FCE_ID").Hidden = True
            ug.DisplayLayout.Bands(1).Columns("CITATION_ID").Hidden = True
            ug.DisplayLayout.Bands(1).Columns("NFA_DATE").Hidden = True
            ug.DisplayLayout.Bands(1).Columns("DELETED").Hidden = True
            ug.DisplayLayout.Bands(1).Columns("SMALL").Hidden = True
            ug.DisplayLayout.Bands(1).Columns("MEDIUM").Hidden = True
            ug.DisplayLayout.Bands(1).Columns("LARGE").Hidden = True
            ug.DisplayLayout.Bands(1).Columns("CITATION").Hidden = True
            ug.DisplayLayout.Bands(1).Columns("ADDRESS_LINE_ONE").Hidden = True
            ug.DisplayLayout.Bands(1).Columns("ADDRESS_TWO").Hidden = True
            ug.DisplayLayout.Bands(1).Columns("CITY").Hidden = True
            ug.DisplayLayout.Bands(1).Columns("STATE").Hidden = True
            ug.DisplayLayout.Bands(1).Columns("ZIP").Hidden = True
            'ug.DisplayLayout.Bands(1).Columns("COUNTY").Hidden = True
            ug.DisplayLayout.Bands(1).Columns("CorrectiveAction").Hidden = True
            ug.DisplayLayout.Bands(1).Columns("INSPECTEDON").Hidden = True
            ug.DisplayLayout.Bands(1).Columns("OCE_ID").Hidden = True
            ug.DisplayLayout.Bands(1).Columns("STATELONG").Hidden = True
            ug.DisplayLayout.Bands(1).Columns("COUNTY").Hidden = True
            ug.DisplayLayout.Bands(1).Columns("CITATIONTEXT_FOR_LETTER").Hidden = True
            ug.DisplayLayout.Bands(1).Columns("CCAT").Hidden = True



            ug.DisplayLayout.Bands(2).Columns("INS_CIT_ID").Hidden = True
            ug.DisplayLayout.Bands(2).Columns("INSPECTION_ID").Hidden = True
            ug.DisplayLayout.Bands(2).Columns("CITATION_ID").Hidden = True
            ug.DisplayLayout.Bands(2).Columns("QUESTION_ID").Hidden = True
            ug.DisplayLayout.Bands(2).Columns("INS_DESCREP_ID").Hidden = True
            ug.DisplayLayout.Bands(2).Columns("CorrectiveAction").Hidden = True
            ug.DisplayLayout.Bands(2).Columns("CitationText").Hidden = True

            ug.DisplayLayout.Override.CellClickAction = CellClickAction.RowSelect

            'ug.DisplayLayout.Bands(0).Columns("POLICY AMOUNT").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            'ug.DisplayLayout.Bands(0).Columns("OVERRIDE AMOUNT").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            ug.DisplayLayout.Bands(0).Columns("DATE RECEIVED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ug.DisplayLayout.Bands(0).Columns("WORKSHOP DATE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ug.DisplayLayout.Bands(0).Columns("WORKSHOP RESULT").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ug.DisplayLayout.Bands(0).Columns("SHOW CAUSE HEARING DATE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ug.DisplayLayout.Bands(0).Columns("SHOW CAUSE HEARING RESULT").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ug.DisplayLayout.Bands(0).Columns("COMMISSION HEARING RESULT").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ug.DisplayLayout.Bands(0).Columns("COMMISSION HEARING DATE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ug.DisplayLayout.Bands(0).Columns("LETTER PRINTED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ug.DisplayLayout.Bands(0).Columns("OVERRIDE DUE DATE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ug.DisplayLayout.Bands(0).Columns("OCE DATE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ug.DisplayLayout.Bands(0).Columns("NEXT DUE DATE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ug.DisplayLayout.Bands(0).Columns("LAST PROCESS DATE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ug.DisplayLayout.Bands(0).Columns("LETTER GENERATED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            ug.DisplayLayout.Bands(0).Columns("FILLER1").Header.Caption = ""
            'ug.DisplayLayout.Bands(0).Columns("FILLER2").Header.Caption = ""
            ug.DisplayLayout.Bands(0).Columns("FILLER3").Header.Caption = ""
            ug.DisplayLayout.Bands(0).Columns("FILLER4").Header.Caption = ""
            ug.DisplayLayout.Bands(0).Columns("FILLER5").Header.Caption = ""

            ug.DisplayLayout.Bands(0).Columns("OCE DATE").Header.Caption = "OCE" + vbCrLf + "DATE"
            ug.DisplayLayout.Bands(0).Columns("LAST PROCESS DATE").Header.Caption = "LAST" + vbCrLf + "PROCESS DATE"
            ug.DisplayLayout.Bands(0).Columns("NEXT DUE DATE").Header.Caption = "NEXT" + vbCrLf + "DUE DATE"
            ug.DisplayLayout.Bands(0).Columns("OVERRIDE DUE DATE").Header.Caption = "OVERRIDE" + vbCrLf + "DUE DATE"
            ug.DisplayLayout.Bands(0).Columns("POLICY AMOUNT").Header.Caption = "POLICY" + vbCrLf + "AMOUNT"
            ug.DisplayLayout.Bands(0).Columns("OVERRIDE AMOUNT").Header.Caption = "OVERRIDE" + vbCrLf + "AMOUNT"
            ug.DisplayLayout.Bands(0).Columns("SETTLEMENT AMOUNT").Header.Caption = "SETTLEMENT" + vbCrLf + "AMOUNT"

            ug.DisplayLayout.Bands(0).Columns("PAID AMOUNT").Header.Caption = "PAID" + vbCrLf + "AMOUNT"
            ug.DisplayLayout.Bands(0).Columns("DATE RECEIVED").Header.Caption = "DATE" + vbCrLf + "RECEIVED"
            ug.DisplayLayout.Bands(0).Columns("WORKSHOP DATE").Header.Caption = "WORKSHOP" + vbCrLf + "DATE"
            ug.DisplayLayout.Bands(0).Columns("WORKSHOP RESULT").Header.Caption = "WORKSHOP" + vbCrLf + "RESULT"
            ug.DisplayLayout.Bands(0).Columns("SHOW CAUSE HEARING DATE").Header.Caption = "SHOW CAUSE" + vbCrLf + "HEARING DATE"
            ug.DisplayLayout.Bands(0).Columns("SHOW CAUSE HEARING RESULT").Header.Caption = "SHOW CAUSE" + vbCrLf + "HEARING RESULT"
            ug.DisplayLayout.Bands(0).Columns("COMMISSION HEARING RESULT").Header.Caption = "COMMISSION" + vbCrLf + "HEARING RESULT"
            ug.DisplayLayout.Bands(0).Columns("COMMISSION HEARING DATE").Header.Caption = "COMMISSION" + vbCrLf + "HEARING DATE"
            ug.DisplayLayout.Bands(0).Columns("AGREED ORDER #").Header.Caption = "AGREED" + vbCrLf + "ORDER #"
            ug.DisplayLayout.Bands(0).Columns("ADMINISTRATIVE ORDER #").Header.Caption = "ADMINISTRATIVE" + vbCrLf + "ORDER #"
            ug.DisplayLayout.Bands(0).Columns("PENDING LETTER").Header.Caption = "PENDING" + vbCrLf + "LETTER"
            ug.DisplayLayout.Bands(0).Columns("LETTER PRINTED").Header.Caption = "LETTER" + vbCrLf + "PRINTED"
            ug.DisplayLayout.Bands(0).Columns("LETTER GENERATED").Header.Caption = "LETTER" + vbCrLf + "GENERATED"

            ug.DisplayLayout.Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Free
            ug.DisplayLayout.Bands(0).Override.DefaultColWidth = 100
            ug.DisplayLayout.Bands(0).ColHeaderLines = 2
            ug.DisplayLayout.Bands(0).Override.RowSelectorWidth = 1
            ug.DisplayLayout.Bands(0).Override.RowSizing = RowSizing.AutoFree
            ug.DisplayLayout.Bands(0).Columns("OWNERNAME").CellMultiLine = Infragistics.Win.DefaultableBoolean.False
            ug.DisplayLayout.Bands(0).Override.DefaultColWidth = 100
            ug.DisplayLayout.Bands(0).ColHeaderLines = 2
            ug.DisplayLayout.Bands(0).Override.RowSelectorWidth = 2

            ' the below lines cause the mask input not to work
            'ug.DisplayLayout.Bands(0).Override.CellMultiLine = Infragistics.Win.DefaultableBoolean.True
            'ug.DisplayLayout.Bands(0).Columns("PENDING LETTER").RowLayoutColumnInfo.SpanY = 2

            If ug.DisplayLayout.Bands(0).Groups.Count = 0 Then
                ug.DisplayLayout.Bands(0).Groups.Add("ROW1")
                'ug.displaylayout.Bands(0).Groups.Add("ROW2")
            End If

            ug.DisplayLayout.Bands(0).GroupHeadersVisible = False

            ug.DisplayLayout.Bands(0).Columns("SELECTED").Group = ug.DisplayLayout.Bands(0).Groups("ROW1")
            ug.DisplayLayout.Bands(0).Columns("FILLER1").Group = ug.DisplayLayout.Bands(0).Groups("ROW1")
            ug.DisplayLayout.Bands(0).Columns("OWNERNAME").Group = ug.DisplayLayout.Bands(0).Groups("ROW1")
            ug.DisplayLayout.Bands(0).Columns("ENSITE ID").Group = ug.DisplayLayout.Bands(0).Groups("ROW1")
            ug.DisplayLayout.Bands(0).Columns("RESCINDED").Group = ug.DisplayLayout.Bands(0).Groups("ROW1")
            ug.DisplayLayout.Bands(0).Columns("FILLER3").Group = ug.DisplayLayout.Bands(0).Groups("ROW1")
            ug.DisplayLayout.Bands(0).Columns("OCE DATE").Group = ug.DisplayLayout.Bands(0).Groups("ROW1")
            ug.DisplayLayout.Bands(0).Columns("LAST PROCESS DATE").Group = ug.DisplayLayout.Bands(0).Groups("ROW1")
            ug.DisplayLayout.Bands(0).Columns("NEXT DUE DATE").Group = ug.DisplayLayout.Bands(0).Groups("ROW1")
            ug.DisplayLayout.Bands(0).Columns("OVERRIDE DUE DATE").Group = ug.DisplayLayout.Bands(0).Groups("ROW1")
            ug.DisplayLayout.Bands(0).Columns("STATUS").Group = ug.DisplayLayout.Bands(0).Groups("ROW1")
            ug.DisplayLayout.Bands(0).Columns("ESCALATION").Group = ug.DisplayLayout.Bands(0).Groups("ROW1")
            ug.DisplayLayout.Bands(0).Columns("POLICY AMOUNT").Group = ug.DisplayLayout.Bands(0).Groups("ROW1")
            ug.DisplayLayout.Bands(0).Columns("OVERRIDE AMOUNT").Group = ug.DisplayLayout.Bands(0).Groups("ROW1")
            ug.DisplayLayout.Bands(0).Columns("SETTLEMENT AMOUNT").Group = ug.DisplayLayout.Bands(0).Groups("ROW1")
            ug.DisplayLayout.Bands(0).Columns("FILLER4").Group = ug.DisplayLayout.Bands(0).Groups("ROW1")

            ug.DisplayLayout.Bands(0).Columns("PAID AMOUNT").Group = ug.DisplayLayout.Bands(0).Groups("ROW1")
            ug.DisplayLayout.Bands(0).Columns("DATE RECEIVED").Group = ug.DisplayLayout.Bands(0).Groups("ROW1")
            ug.DisplayLayout.Bands(0).Columns("WORKSHOP DATE").Group = ug.DisplayLayout.Bands(0).Groups("ROW1")
            ug.DisplayLayout.Bands(0).Columns("WORKSHOP RESULT").Group = ug.DisplayLayout.Bands(0).Groups("ROW1")
            ug.DisplayLayout.Bands(0).Columns("SHOW CAUSE HEARING DATE").Group = ug.DisplayLayout.Bands(0).Groups("ROW1")
            ug.DisplayLayout.Bands(0).Columns("SHOW CAUSE HEARING RESULT").Group = ug.DisplayLayout.Bands(0).Groups("ROW1")
            ug.DisplayLayout.Bands(0).Columns("COMMISSION HEARING DATE").Group = ug.DisplayLayout.Bands(0).Groups("ROW1")
            ug.DisplayLayout.Bands(0).Columns("COMMISSION HEARING RESULT").Group = ug.DisplayLayout.Bands(0).Groups("ROW1")
            ug.DisplayLayout.Bands(0).Columns("AGREED ORDER #").Group = ug.DisplayLayout.Bands(0).Groups("ROW1")
            ug.DisplayLayout.Bands(0).Columns("ADMINISTRATIVE ORDER #").Group = ug.DisplayLayout.Bands(0).Groups("ROW1")
            ug.DisplayLayout.Bands(0).Columns("PENDING LETTER").Group = ug.DisplayLayout.Bands(0).Groups("ROW1")
            ug.DisplayLayout.Bands(0).Columns("LETTER PRINTED").Group = ug.DisplayLayout.Bands(0).Groups("ROW1")
            ug.DisplayLayout.Bands(0).Columns("LETTER GENERATED").Group = ug.DisplayLayout.Bands(0).Groups("ROW1")
            ug.DisplayLayout.Bands(0).Columns("FILLER5").Group = ug.DisplayLayout.Bands(0).Groups("ROW1")

            ug.DisplayLayout.Bands(0).LevelCount = 2
            ug.DisplayLayout.Bands(0).Columns("SELECTED").Level = 0
            ug.DisplayLayout.Bands(0).Columns("FILLER1").Level = 1
            ug.DisplayLayout.Bands(0).Columns("OWNERNAME").Level = 0
            ug.DisplayLayout.Bands(0).Columns("ENSITE ID").Level = 1
            ug.DisplayLayout.Bands(0).Columns("RESCINDED").Level = 0
            ug.DisplayLayout.Bands(0).Columns("FILLER3").Level = 1
            ug.DisplayLayout.Bands(0).Columns("OCE DATE").Level = 0
            ug.DisplayLayout.Bands(0).Columns("LAST PROCESS DATE").Level = 1
            ug.DisplayLayout.Bands(0).Columns("NEXT DUE DATE").Level = 0
            ug.DisplayLayout.Bands(0).Columns("OVERRIDE DUE DATE").Level = 1
            ug.DisplayLayout.Bands(0).Columns("STATUS").Level = 0
            ug.DisplayLayout.Bands(0).Columns("ESCALATION").Level = 1
            ug.DisplayLayout.Bands(0).Columns("POLICY AMOUNT").Level = 0
            ug.DisplayLayout.Bands(0).Columns("OVERRIDE AMOUNT").Level = 1
            ug.DisplayLayout.Bands(0).Columns("SETTLEMENT AMOUNT").Level = 0
            ug.DisplayLayout.Bands(0).Columns("FILLER4").Level = 1
            ug.DisplayLayout.Bands(0).Columns("PAID AMOUNT").Level = 0
            ug.DisplayLayout.Bands(0).Columns("DATE RECEIVED").Level = 1
            ug.DisplayLayout.Bands(0).Columns("WORKSHOP DATE").Level = 0
            ug.DisplayLayout.Bands(0).Columns("WORKSHOP RESULT").Level = 1
            ug.DisplayLayout.Bands(0).Columns("SHOW CAUSE HEARING DATE").Level = 0
            ug.DisplayLayout.Bands(0).Columns("SHOW CAUSE HEARING RESULT").Level = 1
            ug.DisplayLayout.Bands(0).Columns("COMMISSION HEARING RESULT").Level = 1
            ug.DisplayLayout.Bands(0).Columns("COMMISSION HEARING DATE").Level = 0
            ug.DisplayLayout.Bands(0).Columns("AGREED ORDER #").Level = 0
            ug.DisplayLayout.Bands(0).Columns("ADMINISTRATIVE ORDER #").Level = 1
            ug.DisplayLayout.Bands(0).Columns("PENDING LETTER").Level = 0
            ug.DisplayLayout.Bands(0).Columns("LETTER PRINTED").Level = 1
            ug.DisplayLayout.Bands(0).Columns("LETTER GENERATED").Level = 0
            ug.DisplayLayout.Bands(0).Columns("FILLER5").Level = 1

            ug.DisplayLayout.Bands(0).Columns("WORKSHOP RESULT").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
            ug.DisplayLayout.Bands(0).Columns("SHOW CAUSE HEARING RESULT").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList
            ug.DisplayLayout.Bands(0).Columns("COMMISSION HEARING RESULT").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownList

            If ug.DisplayLayout.Bands.Count > 1 Then
                ug.DisplayLayout.Bands(1).Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True

                ug.DisplayLayout.Bands(1).Columns("FACILITY").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                ug.DisplayLayout.Bands(1).Columns("DUE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                ug.DisplayLayout.Bands(1).Columns("RECEIVED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

                If ug.DisplayLayout.Bands.Count > 2 Then
                    ug.DisplayLayout.Bands(2).Columns("DISCREP TEXT").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                    ug.DisplayLayout.Bands(2).Columns("RESCINDED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                    ug.DisplayLayout.Bands(2).Columns("RECEIVED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                    ug.DisplayLayout.Bands(2).Columns("DISCREP TEXT").Width = 400
                End If
            End If

            ' populate the whole column as the table is the same for each row
            ' WorkShopResult
            If ug.DisplayLayout.Bands(0).Columns("WORKSHOP RESULT").ValueList Is Nothing Then
                vListWorkShopResult = New Infragistics.Win.ValueList
                For Each row As DataRow In pOCE.GetWorkshopResults.Tables(0).Rows
                    vListWorkShopResult.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                Next
                ug.DisplayLayout.Bands(0).Columns("WORKSHOP RESULT").ValueList = vListWorkShopResult
            End If
            ' SHOW CAUSE HEARING RESULT
            If ug.DisplayLayout.Bands(0).Columns("SHOW CAUSE HEARING RESULT").ValueList Is Nothing Then
                vListShowCauseResult = New Infragistics.Win.ValueList
                For Each row As DataRow In pOCE.GetShowCauseHearingResults.Tables(0).Rows
                    vListShowCauseResult.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                Next
                ug.DisplayLayout.Bands(0).Columns("SHOW CAUSE HEARING RESULT").ValueList = vListShowCauseResult
            End If
            ' COMMISSION HEARING RESULT
            If ug.DisplayLayout.Bands(0).Columns("COMMISSION HEARING RESULT").ValueList Is Nothing Then
                vListCommissionResult = New Infragistics.Win.ValueList
                For Each row As DataRow In pOCE.GetCommissionHearingResults.Tables(0).Rows
                    vListCommissionResult.ValueListItems.Add(row.Item("PROPERTY_ID"), row.Item("PROPERTY_NAME").ToString)
                Next
                ug.DisplayLayout.Bands(0).Columns("COMMISSION HEARING RESULT").ValueList = vListCommissionResult
            End If

            For Each ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow In ug.DisplayLayout.Grid.Rows ' status
                SetugRowComboValue(ugRow)
                If Not ugRow.Cells("POLICY AMOUNT").Value Is DBNull.Value Then
                    If ugRow.Cells("POLICY AMOUNT").Value < 0 Then
                        ugRow.Cells("POLICY AMOUNT").Value = DBNull.Value
                    End If
                End If
                If Not ugRow.Cells("SETTLEMENT AMOUNT").Value Is DBNull.Value Then
                    If ugRow.Cells("SETTLEMENT AMOUNT").Value < 0 Then
                        ugRow.Cells("SETTLEMENT AMOUNT").Value = DBNull.Value
                    End If
                End If
                If Not ugRow.Cells("OVERRIDE AMOUNT").Value Is DBNull.Value Then
                    If ugRow.Cells("OVERRIDE AMOUNT").Value < 0 Then
                        ugRow.Cells("OVERRIDE AMOUNT").Value = DBNull.Value
                    End If
                End If
                If Not ugRow.Cells("PAID AMOUNT").Value Is DBNull.Value Then
                    If ugRow.Cells("PAID AMOUNT").Value < 0 Then
                        ugRow.Cells("PAID AMOUNT").Value = DBNull.Value
                    End If
                End If

                If Not ugRow.Cells("AGREED ORDER #").Value Is DBNull.Value Then
                    If ugRow.Cells("AGREED ORDER #").Value = String.Empty Then
                        ugRow.Cells("AGREED ORDER #").Value = DBNull.Value
                    End If
                End If
                If Not ugRow.Cells("ADMINISTRATIVE ORDER #").Value Is DBNull.Value Then
                    If ugRow.Cells("ADMINISTRATIVE ORDER #").Value = String.Empty Then
                        ugRow.Cells("ADMINISTRATIVE ORDER #").Value = DBNull.Value
                    End If
                End If

                '' WorkShopResult
                'If vListWorkShopResult.FindByDataValue(ugRow.Cells("WORKSHOP RESULT").Value) Is Nothing Then
                '    ugRow.Cells("WORKSHOP RESULT").Value = DBNull.Value
                'End If
                '' SHOW CAUSE HEARING RESULT
                'If vListShowCauseResult.FindByDataValue(ugRow.Cells("SHOW CAUSE HEARING RESULT").Value) Is Nothing Then
                '    ugRow.Cells("SHOW CAUSE HEARING RESULT").Value = DBNull.Value
                'End If
                '' COMMISSION HEARING RESULT
                'If vListCommissionResult.FindByDataValue(ugRow.Cells("COMMISSION HEARING RESULT").Value) Is Nothing Then
                '    ugRow.Cells("COMMISSION HEARING RESULT").Value = DBNull.Value
                'End If
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub btnExpandOwnerCitationAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExpandOwnerCitationAll.Click
        Try
            If btnExpandOwnerCitationAll.Text = "Expand All" Then
                ExpandAll(True, ugOCE2, btnExpandOwnerCitationAll)
            Else
                ExpandAll(False, ugOCE2, btnExpandOwnerCitationAll)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugFacilityList_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugFacilityList.InitializeLayout
        Try
            e.Layout.Bands(0).Columns("OWNER_ID").Hidden = True
            e.Layout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
            e.Layout.Bands(0).Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
            e.Layout.Bands(0).Columns("Last Inspected").CellActivation = Activation.NoEdit
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugFacilityList_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugFacilityList.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            If Not ugFacilityList.ActiveRow Is Nothing Then
                lblFacilityIDValue.Text = ugFacilityList.ActiveRow.Cells("FACILITYID").Value.ToString
                nFacilityID = CInt(ugFacilityList.ActiveRow.Cells("FACILITYID").Value)
                tbCntrlCandE.SelectedTab = tbPageFacilityDetail
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugOCEs_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugOCEs.InitializeLayout
        Try
            SetupOCE(ugOCEs)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugOCE2_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugOCE2.InitializeLayout
        Try
            SetupOCE(ugOCE2)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub tbCtrlOwner_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbCtrlOwner.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            If tbCtrlOwner.SelectedTab.Name = tbPageOwnerFacilities.Name Then
                PopulateOwnerFacilities(nOwnerID)
            ElseIf tbCtrlOwner.SelectedTab.Name = tbPageOCEs.Name Then
                PopulateOCE(nOwnerID)
            ElseIf tbCtrlOwner.SelectedTab.Name = tbPageOwnerDocuments.Name Then
                UCOwnerDocuments.LoadDocumentsGrid(nOwnerID, UIUtilsGen.EntityTypes.Owner, UIUtilsGen.ModuleID.CAE)
            ElseIf tbCtrlOwner.SelectedTab.Name = tbPageOwnerContactList.Name Then
                LoadContacts(ugOwnerContacts, nOwnerID, UIUtilsGen.EntityTypes.Owner)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnOwnerFlag_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOwnerFlag.Click
        FlagMaintenance(sender, e)
    End Sub
    Private Sub btnOwnerComment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOwnerComment.Click
        CommentsMaintenance(sender, e)
    End Sub

    Private Sub btnEnforceViewEnforceHistory_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEnforceViewEnforceHistory.Click
        ViewEnforcementHistory()
    End Sub
    Private Sub btnEnforceViewEnforceHistory1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEnforceViewEnforceHistory1.Click
        ViewEnforcementHistory()
    End Sub

#End Region

#Region "Facility / Facility Citations Tab"
    Friend Sub PopulateFacility(ByVal facID As Integer)
        Try
            lblInViolationValue.Text = String.Empty
            txtDesOp.Text = pOwn.Facilities.DesignatedOperator
            UIUtilsGen.PopulateFacilityInfo(Me, pOwn.OwnerInfo, pOwn.Facilities, facID)
            txtDesMa.Text = pOwn.Facilities.DesignatedManager
            For Each ugRow In ugFacilityList.Rows
                If ugRow.Cells("FACILITYID").Value = facID Then
                    lblInViolationValue.Text = ugRow.Cells("In Violation").Text
                End If
            Next
            If lblInViolationValue.Text.ToUpper = "YES" Then
                lblInViolationValue.BackColor = Color.Red
            Else
                lblInViolationValue.BackColor = System.Drawing.SystemColors.Control
            End If
            PopulateFCE(facID)
            If tbCntrlFacility.SelectedTab.Name = tbPageFacilityDocuments.Name Then
                UCFacilityDocuments.LoadDocumentsGrid(Integer.Parse(Me.lblFacilityIDValue.Text), UIUtilsGen.EntityTypes.Facility, UIUtilsGen.ModuleID.CAE)
            ElseIf tbCntrlFacility.SelectedTab.Name = tbPageFacilityContactList.Name Then
                LoadContacts(ugFacilityContacts, nFacilityID, UIUtilsGen.EntityTypes.Facility)
            End If
            If bolFrmActivated Then
                MC.FlagsChanged(facID, UIUtilsGen.EntityTypes.Facility, "C " + "&" + " E", Me.Text)
                'MC.FlagsChanged(lblFacilityIDValue.Text, UIUtilsGen.EntityTypes.Facility, "C " + "&" + " E", Me.Text)
            End If
            CommentsMaintenance(, , True)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Friend Sub PopulateFCE(ByVal facID As Integer)
        Try
            Dim ds As DataSet = pFCE.GetCompliances(facID, False)
            ugFCEs.DataSource = ds
            ugFCEs.DrawFilter = rp
            ExpandAll(False, ugFCEs, )
            ugFCE2.DataSource = ds
            ugFCE2.DrawFilter = rp
            ExpandAll(True, ugFCE2, btnExpandFacilityCitationAll)
            lblTotalNoOfLUSTEventsValue.Text = ugFCEs.Rows.Count.ToString


        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ViewInspectionHistory()
        Dim dsViewHistory As DataSet
        Dim oInspection As New MUSTER.BusinessLogic.pInspection
        Try
            dsViewHistory = oInspection.RetrieveInspectionHistory(pOwn.Facilities.ID)
            If dsViewHistory.Tables(0).Rows.Count = 0 Then
                MsgBox("Facility has no Inspection history")
                Exit Sub
            Else
                frmInspecHistory = New InspectionHistory(oInspection, dsViewHistory, pOwn.Facilities.Name)
                frmInspecHistory.ShowDialog()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Function GetPrevNextFacility(ByVal facID As Integer, ByVal getNext As Boolean) As Integer
        Try
            Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
            Dim retVal As Integer = facID
            Dim sl As New SortedList

            For Each ugRow In ugFacilityList.Rows
                sl.Add(ugRow.Cells("FacilityID").Value, ugRow.Cells("FacilityID").Value)
            Next
            Return GetPrevNext(sl, getNext, facID)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function
    Private Function GetPrevNext(ByVal sl As SortedList, ByVal getNext As Boolean, ByVal key As Integer) As Integer
        Try
            Dim retVal As Integer
            Dim index As Integer = sl.IndexOfKey(key)

            If getNext Then
                If sl.Count = 1 Then
                    index = -1
                ElseIf index = sl.Count - 1 Then
                    index = -1
                End If
                retVal = sl.GetByIndex(index + 1)
            Else
                If sl.Count = 1 Then
                    index = 1
                ElseIf index = 0 Then
                    index = sl.Count
                End If
                retVal = sl.GetByIndex(index - 1)
            End If
            Return retVal
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Sub SetupFCE(ByRef ug As Infragistics.Win.UltraWinGrid.UltraGrid)
        Try
            ug.DisplayLayout.Bands(0).Override.RowAppearance.BackColor = Color.RosyBrown
            ug.DisplayLayout.Bands(0).Override.RowAlternateAppearance.BackColor = Color.PeachPuff
            ug.DisplayLayout.Bands(1).Override.RowAppearance.BackColor = Color.Khaki

            ug.DisplayLayout.Bands(0).Columns("FCE_ID").Hidden = True
            ug.DisplayLayout.Bands(0).Columns("OWNER_ID").Hidden = True
            ug.DisplayLayout.Bands(0).Columns("INSPECTION_ID").Hidden = True
            ug.DisplayLayout.Bands(0).Columns("SELECTED").Hidden = True

            ug.DisplayLayout.Bands(0).Columns("DATE GENERATED").Header.Caption = "FCE GENERATED DATE"

            ug.DisplayLayout.Bands(0).Columns("DATE GENERATED").CellActivation = Activation.NoEdit
            ug.DisplayLayout.Bands(0).Columns("INSPECTED ON").CellActivation = Activation.NoEdit

            ug.DisplayLayout.Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Free

            ug.DisplayLayout.Override.CellClickAction = CellClickAction.RowSelect

            ug.DisplayLayout.Bands(0).Columns("INSPECTED ON").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Descending

            If ug.DisplayLayout.Bands.Count > 1 Then
                ug.DisplayLayout.Bands(1).Columns("INSPECTION_ID").Hidden = True
                ug.DisplayLayout.Bands(1).Columns("FCE_ID").Hidden = True
                ug.DisplayLayout.Bands(1).Columns("FACILITY_ID").Hidden = True
                ug.DisplayLayout.Bands(1).Columns("INS_CIT_ID").Hidden = True
                ug.DisplayLayout.Bands(1).Columns("QUESTION_ID").Hidden = True
                ug.DisplayLayout.Bands(1).Columns("CITATION_ID").Hidden = True
                ug.DisplayLayout.Bands(1).Columns("SMALL").Hidden = True
                ug.DisplayLayout.Bands(1).Columns("MEDIUM").Hidden = True
                ug.DisplayLayout.Bands(1).Columns("LARGE").Hidden = True

                If ug.DisplayLayout.Bands.Count > 2 Then
                    ug.DisplayLayout.Bands(2).Columns("INS_CIT_ID").Hidden = True
                    ug.DisplayLayout.Bands(2).Columns("INSPECTION_ID").Hidden = True
                    ug.DisplayLayout.Bands(2).Columns("CITATION_ID").Hidden = True
                    ug.DisplayLayout.Bands(2).Columns("QUESTION_ID").Hidden = True

                    ug.DisplayLayout.Bands(2).Columns("DISCREP TEXT").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                    ug.DisplayLayout.Bands(2).Columns("DISCREP TEXT").Width = 400
                End If

            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub btnInspHistory_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInspHistory.Click
        ViewInspectionHistory()
    End Sub
    Private Sub btnInspHistory1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInspHistory1.Click
        ViewInspectionHistory()
    End Sub

    Private Sub btnExpandFacilityCitationAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExpandFacilityCitationAll.Click
        Try
            If btnExpandFacilityCitationAll.Text = "Expand All" Then
                ExpandAll(True, ugFCE2, btnExpandFacilityCitationAll)
            Else
                ExpandAll(False, ugFCE2, btnExpandFacilityCitationAll)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugFCEs_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugFCEs.InitializeLayout
        Try
            SetupFCE(ugFCEs)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugFCE2_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugFCE2.InitializeLayout
        Try
            SetupFCE(ugFCE2)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnFacFlags_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFacFlags.Click
        FlagMaintenance(sender, e)
    End Sub
    Private Sub btnFacComments_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFacComments.Click
        CommentsMaintenance(sender, e)
    End Sub

    Private Sub lnkLblNextFac_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkLblNextFac.LinkClicked
        Try
            If IsNumeric(lblFacilityIDValue.Text) Then
                lblFacilityIDValue.Text = GetPrevNextFacility(lblFacilityIDValue.Text, True).ToString
                SetupTabs()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub lnkLblPrevFacility_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkLblPrevFacility.LinkClicked
        Try
            If IsNumeric(lblFacilityIDValue.Text) Then
                lblFacilityIDValue.Text = GetPrevNextFacility(lblFacilityIDValue.Text, False).ToString
                SetupTabs()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub tbCntrlFacility_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbCntrlFacility.Click
        If tbCntrlFacility.SelectedTab.Name = tbPageFacilityDocuments.Name Then
            UCFacilityDocuments.LoadDocumentsGrid(nFacilityID, UIUtilsGen.EntityTypes.Facility, UIUtilsGen.ModuleID.CAE)
        ElseIf tbCntrlFacility.SelectedTab.Name = tbPageFacilityContactList.Name Then
            LoadContacts(ugFacilityContacts, nFacilityID, UIUtilsGen.EntityTypes.Facility)
        ElseIf tbCntrlFacility.SelectedTab.Name = tbPageFCEs.Name Then
            PopulateFCE(nFacilityID)
        End If
    End Sub
#End Region

#Region "Flags"
    Private Sub SF_FlagAdded(ByVal entityID As Integer, ByVal entityType As Integer, ByVal [Module] As String, ByVal ParentFormText As String) Handles SF.FlagAdded
        MC.FlagsChanged(entityID, entityType, [Module], ParentFormText)
    End Sub
    Private Sub SF_RefreshCalendar() Handles SF.RefreshCalendar
        MC.RefreshCalendarInfo()
        MC.LoadDueToMeCalendar()
        MC.LoadToDoCalendar()
    End Sub
    Private Sub FlagMaintenance(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim ownID, facID As Integer
            Select Case Me.tbCntrlCandE.SelectedTab.Name
                Case tbPageOwnerDetail.Name
                    SF = New ShowFlags(pOwn.ID, UIUtilsGen.EntityTypes.Owner, "C " + "&" + " E")
                Case tbPageFacilityDetail.Name
                    SF = New ShowFlags(pOwn.Facilities.ID, UIUtilsGen.EntityTypes.Facility, "C " + "&" + " E")
                Case Else
                    Exit Sub
            End Select
            SF.ShowDialog()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "Comments"
    Private Sub CommentsMaintenance(Optional ByVal sender As System.Object = Nothing, Optional ByVal e As System.EventArgs = Nothing, Optional ByVal bolSetCounts As Boolean = False, Optional ByVal resetBtnColor As Boolean = False)
        Dim SC As ShowComments
        Dim nEntityType As Integer = 0
        Dim nEntityID As Integer = 0
        Dim strEntityName As String = String.Empty
        Dim oComments As MUSTER.BusinessLogic.pComments
        Dim nCommentsCount As Integer = 0
        Try
            Select Case Me.tbCntrlCandE.SelectedTab.Name
                Case tbPageOwnerDetail.Name
                    strEntityName = "Owner : " + CStr(pOwn.ID) + " " + Me.txtOwnerName.Text
                    oComments = pOwn.Comments
                    nEntityID = pOwn.ID
                    nEntityType = UIUtilsGen.EntityTypes.Owner
                Case tbPageFacilityDetail.Name
                    strEntityName = "Facility : " + CStr(pOwn.Facilities.ID) + " " + pOwn.Facilities.Name
                    oComments = pOwn.Facilities.Comments
                    nEntityID = pOwn.Facilities.ID
                    nEntityType = UIUtilsGen.EntityTypes.Facility
                Case Else
                    Exit Sub
            End Select
            If Not resetBtnColor Then
                SC = New ShowComments(nEntityID, nEntityType, IIf(bolSetCounts, "", "C " + "&" + " E"), strEntityName, oComments, Me.Text)
                If bolSetCounts Then
                    nCommentsCount = SC.GetCounts()
                Else
                    SC.ShowDialog()
                    nCommentsCount = IIf(SC.nCommentsCount <= 0, SC.GetCounts(), SC.nCommentsCount)
                End If
            End If
            If nEntityType = UIUtilsGen.EntityTypes.Owner Then
                If nCommentsCount > 0 Then
                    btnOwnerComment.BackColor = Drawing.Color.FromName(UIUtilsGen.CommentsButton_HasCmts)
                Else
                    btnOwnerComment.BackColor = Drawing.Color.FromName(UIUtilsGen.CommentsButton_NoCmts)
                End If
            ElseIf nEntityType = UIUtilsGen.EntityTypes.Facility Then
                If nCommentsCount > 0 Then
                    btnFacComments.BackColor = Drawing.Color.FromName(UIUtilsGen.CommentsButton_HasCmts)
                Else
                    btnFacComments.BackColor = Drawing.Color.FromName(UIUtilsGen.CommentsButton_NoCmts)
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "Contacts"
#Region "Owner Contacts"
    Private Sub ugOwnerContacts_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugOwnerContacts.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            ModifyContact(ugOwnerContacts)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnOwnerAddSearchContact_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOwnerAddSearchContact.Click
        Try
            pConStruct = New MUSTER.BusinessLogic.pContactStruct
            objCntSearch = New ContactSearch(nOwnerID, UIUtilsGen.EntityTypes.Owner, "C " + "&" + " E", pConStruct)
            'objCntSearch.Show()
            objCntSearch.ShowDialog()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnOwnerModifyContact_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOwnerModifyContact.Click
        Try
            ModifyContact(ugOwnerContacts)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnOwnerAssociateContact_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOwnerAssociateContact.Click
        Try
            AssociateContact(ugOwnerContacts, nOwnerID, UIUtilsGen.EntityTypes.Owner)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnOwnerDeleteContact_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOwnerDeleteContact.Click
        Try
            DeleteContact(ugOwnerContacts, nOwnerID)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub chkOwnerShowActiveOnly_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkOwnerShowActiveOnly.CheckedChanged
        'Dim dsContactsLocal As DataSet
        Try
            'If Not dsContacts Is Nothing Then
            SetOwnerFilter()
            'End If


        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub chkOwnerShowContactsforAllModules_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkOwnerShowContactsforAllModules.CheckedChanged
        'Dim dsContactsLocal As DataSet
        Try


            'If Not dsContacts Is Nothing Then
            SetOwnerFilter()
            'End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub chkOwnerShowRelatedContacts_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkOwnerShowRelatedContacts.CheckedChanged
        Try
            ' If Not dsContacts Is Nothing Then
            SetOwnerFilter()
            'End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub SetOwnerFilter()
        Dim dsContactsLocal As DataSet
        Dim bolActive As Boolean = False
        Dim strEntities As String = String.Empty
        Dim nModuleID As Integer = 0
        Dim nEntityID As Integer = 0
        Dim nEntityType As Integer = 0
        Dim nRelatedEntityType As Integer = 0
        Try
            'strFilterString = String.Empty
            If chkOwnerShowActiveOnly.Checked Then
                bolActive = True
            Else
                bolActive = False
            End If

            nEntityType = 9
            If chkOwnerShowContactsforAllModules.Checked Then
                ' User has the ability to view the contacts associated for the entity in other modules
                nEntityID = pOwn.ID
                nModuleID = 0

            Else
                nEntityID = pOwn.ID
                nModuleID = 613
            End If

            If chkOwnerShowRelatedContacts.Checked And strFacilityIdTags <> String.Empty Then
                strEntities = strFacilityIdTags
                nRelatedEntityType = 6
            Else
                strEntities = String.Empty
            End If

            UIUtilsGen.LoadContacts(ugOwnerContacts, nEntityID, nEntityType, pConStruct, nModuleID, strEntities, bolActive, String.Empty, nRelatedEntityType)

        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
#End Region
#Region "Facility Contacts"

    Private Sub ugFacilityContacts_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugFacilityContacts.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            ModifyContact(ugFacilityContacts)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnFacilityAddSearchContact_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFacilityAddSearchContact.Click
        Try
            pConStruct = New MUSTER.BusinessLogic.pContactStruct
            objCntSearch = New ContactSearch(nFacilityID, UIUtilsGen.EntityTypes.Facility, "C " + "&" + " E", pConStruct)
            objCntSearch.ShowDialog()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnFacilityModifyContact_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFacilityModifyContact.Click
        Try
            ModifyContact(ugFacilityContacts)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnFacilityAssociateContact_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFacilityAssociateContact.Click
        Try
            AssociateContact(ugFacilityContacts, nFacilityID, UIUtilsGen.EntityTypes.Facility)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub btnFacilityDeleteContact_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFacilityDeleteContact.Click
        Try
            DeleteContact(ugFacilityContacts, nFacilityID)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub chkFacilityShowActiveContactOnly_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkFacilityShowActiveContactOnly.CheckedChanged
        Try
            SetFacilityFilter()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub chkFacilityShowContactsforAllModule_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkFacilityShowContactsforAllModule.CheckedChanged
        Try
            SetFacilityFilter()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub chkFacilityShowRelatedContacts_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkFacilityShowRelatedContacts.CheckedChanged
        Try
            SetFacilityFilter()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub SetFilter()
        Try
            strFilterString = String.Empty
            Dim strEntityID As String
            If tbCntrlCandE.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL" Then
                strEntityID = pOwn.ID.ToString
            Else
                strEntityID = pOwn.Facility.ID.ToString
            End If

            If chkOwnerShowActiveOnly.Checked Then
                strFilterString = "(ACTIVE = 1"
            Else
                strFilterString = "("
            End If

            If chkOwnerShowContactsforAllModules.Checked Then

                ' User has the ability to view the contacts associated for the entity in other modules
                If strFilterString = "(" Then
                    strFilterString += "ENTITYID = " + strEntityID
                Else
                    strFilterString += "AND ENTITYID = " + strEntityID
                End If
            Else
                If strFilterString = "(" Then
                    strFilterString += " MODULEID = 613 And ENTITYID = " + strEntityID
                Else
                    strFilterString += " AND MODULEID = 613 And ENTITYID = " + strEntityID
                End If
            End If
            If chkOwnerShowRelatedContacts.Checked Then
                strFilterString += " OR " + IIf(Not strFacilityIdTags = String.Empty, " ENTITYID in (" + strFacilityIdTags + "))", "")
            Else
                strFilterString += ")"
            End If

            dsContacts.Tables(0).DefaultView.RowFilter = strFilterString
            UIUtilsGen.LoadContacts(Me.ugOwnerContacts, 0, 0, pConStruct, 0, , , , , dsContacts.Tables(0).DefaultView)

        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub SetFacilityFilter()
        Dim dsContactsLocal As DataSet
        Dim bolActive As Boolean = False
        Dim strEntities As String = String.Empty
        Dim nModuleID As Integer = 0
        Dim nEntityID As Integer = 0
        Dim nEntityType As Integer = 0
        Dim nRelatedEntityType As Integer = 0
        Dim strEntityAssocIDs As String = String.Empty
        Try

            'strFilterString = String.Empty
            If chkFacilityShowActiveContactOnly.Checked Then
                bolActive = True
            Else
                bolActive = False
            End If
            nEntityType = 6

            If chkFacilityShowContactsforAllModule.Checked Then
                'User has the ability to view the contacts associated for the entity in other modules
                Dim strFilterForAllModules As String = pConStruct.GetContactsForAllModules(pOwn.Facilities.ID.ToString)
                strEntityAssocIDs = strFilterForAllModules
                nModuleID = 0
                nEntityID = pOwn.Facilities.ID.ToString

                nModuleID = 0
            Else
                nEntityID = pOwn.Facility.ID.ToString
                nModuleID = 613
            End If

            If chkFacilityShowRelatedContacts.Checked Then
                strEntities = strFacilityIdTags
                nRelatedEntityType = 6
            Else
                strEntities = String.Empty
            End If
            UIUtilsGen.LoadContacts(ugFacilityContacts, nEntityID, nEntityType, pConStruct, nModuleID, strEntities, bolActive, strEntityAssocIDs, nRelatedEntityType)

        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
#End Region
#Region "Common Functions"
    Private Sub LoadContacts(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal EntityID As Integer, ByVal EntityType As Integer)
        Try

            UIUtilsGen.LoadContacts(ugGrid, EntityID, EntityType, pConStruct, 613)

            If tbCntrlCandE.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL" Then
                Me.chkOwnerShowActiveOnly.Checked = False
                Me.chkOwnerShowActiveOnly.Checked = True
            ElseIf tbCntrlCandE.SelectedTab.Name.ToUpper = "TBPAGEFACILITYDETAIL" Then
                Me.chkFacilityShowActiveContactOnly.Checked = False
                Me.chkFacilityShowActiveContactOnly.Checked = True
            End If

        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Function ModifyContact(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid)
        Try

            If UIUtilsGen.ModifyContact(ugGrid, 613, pConStruct) Then
                Contact_ContactAdded()
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function

    Private Function AssociateContact(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal nEntityID As Integer, ByVal nEntityType As Integer)
        Try

            If UIUtilsGen.AssociateContact(ugGrid, nEntityID, nEntityType, 613, pConStruct) Then
                Contact_ContactAdded()
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function

    Private Function DeleteContact(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal nEntityID As Integer)
        Try

            If UIUtilsGen.DeleteContact(ugGrid, nEntityID, 613, pConStruct) Then
                Contact_ContactAdded()
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function

#End Region
#Region "Close Events"
    Private Sub Search_ContactAdded() Handles objCntSearch.ContactAdded
        If tbCntrlCandE.SelectedTab.Name = tbPageOwnerDetail.Name Then
            LoadContacts(ugOwnerContacts, nOwnerID, UIUtilsGen.EntityTypes.Owner)
            SetOwnerFilter()
        Else
            LoadContacts(ugOwnerContacts, nFacilityID, UIUtilsGen.EntityTypes.Facility)
            SetFacilityFilter()
        End If
    End Sub
    Private Sub Contact_ContactAdded()
        Try
            If tbCntrlCandE.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL".ToUpper Then
                LoadContacts(ugOwnerContacts, nOwnerID, UIUtilsGen.EntityTypes.Owner)
            Else
                LoadContacts(ugOwnerContacts, nFacilityID, UIUtilsGen.EntityTypes.Facility)
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub objCntSearch_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles objCntSearch.Closing
        If Not objCntSearch Is Nothing Then
            objCntSearch = Nothing
        End If
    End Sub
#End Region
#End Region

#Region "Form Events"

    Private Sub JumpToCNEManagement(ByVal sender As Object, ByVal e As EventArgs)

        If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
            Me.Cursor = Cursors.Default
            Exit Sub
        End If

        With DirectCast(sender, Infragistics.Win.UltraWinGrid.UltraGrid)

            If .ActiveRow.Band.Index <= 1 Then

                Dim keyValue As Integer
                Dim fieldName As String

                If .ActiveRow.Band.Index = 0 Then
                    fieldName = "OWNER_ID"
                    keyValue = Convert.ToInt32(.ActiveRow.Cells("OWNER_ID").Value)
                Else
                    fieldName = "FACILITY_ID"
                    keyValue = Convert.ToInt32(.ActiveRow.Cells("FACILITY_ID").Value)
                End If


                'initialize new CNE management Form
                Dim CNEM As New CandEManagement

                'sets the command and faciliy ID
                GridMaster.GlobalInstance.ActivateBot(keyValue, CNEM.GetType, GridMaster.OPENCNERECORDBYID, fieldName)

                'Opens CNE management and opens Enforcements Tab 

                CNEM.MdiParent = Me.MdiParent

                CNEM.tabCntrlCandE.SelectedTab = CNEM.tbPageEnforcement

                CNEM.Show()

            End If

        End With

    End Sub

    Private Sub CandE_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        ' The following line enables all forms to detect the visible form in the MDI container
        '
        MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "C " + "&" + " E")
        bolFrmActivated = True
        If lblOwnerIDValue.Text <> String.Empty Then ' And lblFacilityIDValue.Text = String.Empty Then
            pOwn.Retrieve(Me.lblOwnerIDValue.Text, "SELF")
        End If
    End Sub
    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)
        Try
            UIUtilsGen.RemoveOwner(pOwn, Me)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub CandE_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.TextChanged
        Try
            MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "WindowName", Me.Text, "C " + "&" + " E")
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "Envelopes and Labels"
    Private Sub btnCAEOwnerEnvelopes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCAEOwnerEnvelopes.Click
        Dim arrAddress(4) As String
        Try

            arrAddress(0) = pOwn.Addresses.AddressLine1
            arrAddress(1) = pOwn.Addresses.AddressLine2
            arrAddress(2) = pOwn.Addresses.City
            arrAddress(3) = pOwn.Addresses.State
            arrAddress(4) = pOwn.Addresses.Zip
            'strAddress = pOwn.Addresses.AddressLine1 + "," + pOwn.Addresses.AddressLine2 + "," + pOwn.Addresses.City + "," + pOwn.Addresses.State + "," + pOwn.Addresses.Zip
            If pOwn.Addresses.AddressId > 0 And pOwn.ID > 0 Then

                Dim strContactName As String
                Dim dsContactsLocal = pConStruct.GetFilteredContacts(pOwn.ID, 612)

                If dsContactsLocal.Tables(0).Rows.Count > 0 Then
                    For Each contactRow As DataRow In dsContactsLocal.Tables(0).Rows
                        If contactRow("Type") = "Registration Representative" Then
                            strContactName = contactRow("CONTACT_name")
                            Exit For
                        End If

                    Next

                End If

                UIUtilsGen.CreateEnvelopes(Me.txtOwnerName.Text, arrAddress, "CAE", pOwn.ID, strContactName)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCAEOwnerLabels_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCAEOwnerLabels.Click
        Dim arrAddress(4) As String
        Try
            arrAddress(0) = pOwn.Addresses.AddressLine1
            arrAddress(1) = pOwn.Addresses.AddressLine2
            arrAddress(2) = pOwn.Addresses.City
            arrAddress(3) = pOwn.Addresses.State
            arrAddress(4) = pOwn.Addresses.Zip
            'strAddress = pOwn.Addresses.AddressLine1 + "," + pOwn.Addresses.AddressLine2 + "," + pOwn.Addresses.City + "," + pOwn.Addresses.State + "," + pOwn.Addresses.Zip
            If pOwn.Addresses.AddressId > 0 And pOwn.ID > 0 Then
                UIUtilsGen.CreateLabels(Me.txtOwnerName.Text, arrAddress, "CAE", pOwn.ID)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCAEFacEnvelopes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCAEFacEnvelopes.Click
        Dim arrAddress(4) As String

        Try
            arrAddress(0) = pOwn.Facilities.FacilityAddresses.AddressLine1
            arrAddress(1) = pOwn.Facilities.FacilityAddresses.AddressLine2
            arrAddress(2) = pOwn.Facilities.FacilityAddresses.City
            arrAddress(3) = pOwn.Facilities.FacilityAddresses.State
            arrAddress(4) = pOwn.Facilities.FacilityAddresses.Zip
            'strAddress = pOwn.Facilities.FacilityAddresses.AddressLine1 + "," + pOwn.Facilities.FacilityAddresses.AddressLine2 + "," + pOwn.Facilities.FacilityAddresses.City + "," + pOwn.Facilities.FacilityAddresses.State + "," + pOwn.Facilities.FacilityAddresses.Zip
            If pOwn.Facilities.FacilityAddresses.AddressId > 0 And pOwn.ID > 0 Then
                UIUtilsGen.CreateEnvelopes(Me.txtFacilityName.Text, arrAddress, "CAE", pOwn.Facilities.ID)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCAEFacLabels_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCAEFacLabels.Click
        Dim arrAddress(4) As String
        Try
            arrAddress(0) = pOwn.Facilities.FacilityAddresses.AddressLine1
            arrAddress(1) = pOwn.Facilities.FacilityAddresses.AddressLine2
            arrAddress(2) = pOwn.Facilities.FacilityAddresses.City
            arrAddress(3) = pOwn.Facilities.FacilityAddresses.State
            arrAddress(4) = pOwn.Facilities.FacilityAddresses.Zip
            'strAddress = pOwn.Facilities.FacilityAddresses.AddressLine1 + "," + pOwn.Facilities.FacilityAddresses.AddressLine2 + "," + pOwn.Facilities.FacilityAddresses.City + "," + pOwn.Facilities.FacilityAddresses.State + "," + pOwn.Facilities.FacilityAddresses.Zip
            If pOwn.Facilities.FacilityAddresses.AddressId > 0 And pOwn.ID > 0 Then
                UIUtilsGen.CreateLabels(Me.txtFacilityName.Text, arrAddress, "CAE", pOwn.Facilities.ID)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

    Public Property FormLoading() As Boolean
        Get
            Return bolLoading
        End Get
        Set(ByVal Value As Boolean)
            bolLoading = Value
        End Set
    End Property
    Public ReadOnly Property MC() As MusterContainer
        Get
            mContainer = Me.MdiParent
            If mContainer Is Nothing Then
                mContainer = New MusterContainer
            End If
            Return mContainer
        End Get
    End Property

End Class
