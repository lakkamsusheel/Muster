Public Class Financial
    Inherits BaseModuleScreen



    ' Upgrade & fixes
    ''' 1.1        Thomas Franey  added Activity Type 50 to condition on line  6029


#Region "User Defined Variables"

    Public strFacilityIdTags As String
    Public nFacilityID As Integer
    'Public Shared strTotalPaidDocTag As String
    Public MyGuid As New System.Guid
    Private bolLoading As Boolean = False
    Public bolNewPersona As Boolean = False
    Private bolDisplayErrmessage As Boolean = True
    Private bolValidateSuccess As Boolean = True
    Private bolFrmActivated As Boolean = False

    Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long

    Private frmCommitment As Commitment
    Private frmAdjustment As Adjustment
    Private frmIncompApplication As IncompleteApplication
    Private frmInvoice As Invoice

    Private dsContacts As DataSet
    Private result As DialogResult

    Private oFinancial As New MUSTER.BusinessLogic.pFinancial
    Private oTechnical As New MUSTER.BusinessLogic.pLustEvent
    Private oAddressInfo As MUSTER.Info.AddressInfo
    Private oCommitment As New MUSTER.BusinessLogic.pFinancialCommitment

    Private WithEvents pOwn As MUSTER.BusinessLogic.pOwner
    Private WithEvents SF As ShowFlags
    Private WithEvents objCntSearch As ContactSearch
    Private nAdjustmentBand As Int16
    Private strFinancialEventIdTags As String
    Private nLastEventID As Int64
    Private pConStruct As New MUSTER.BusinessLogic.pContactStruct
    Dim returnVal As String = String.Empty
    Private strVendorEnvelopLabelAddress As String = String.Empty
    Friend bolFromTechnical As Boolean = False
    Friend nCurrentEventID As Integer = -1
    Dim arrVendorAddress(4) As String
    Dim strFilterString As String = String.Empty

    Private nReimbursementID As Integer = 0
    Private nCommitmentID As Integer = 0
    Private bolEnableDeleteEvent = False

    Public Property FinancialEventGrid() As Infragistics.Win.UltraWinGrid.UltraGrid

        Get
            Return ugFinancialGrid
        End Get

        Set(ByVal Value As Infragistics.Win.UltraWinGrid.UltraGrid)
            ugFinancialGrid = Value
        End Set
    End Property


#End Region

#Region " Windows Form Designer generated code "

    Public Sub New(ByRef oOwner As MUSTER.BusinessLogic.pOwner, ByVal OwnerID As Int64, ByVal FacilityID As Int64)
        MyBase.New()
        MyGuid = System.Guid.NewGuid
        bolLoading = True
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        pOwn = oOwner
        'Add any initialization after the InitializeComponent() call
        MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "Financial")
        Try

            InitControls()
            If OwnerID > 0 Then
                PopulateOwnerInfo(Integer.Parse(OwnerID))
                If FacilityID <= 0 Then
                    tbCntrlFinancial.SelectedTab = tbPageOwnerDetail
                    Me.Text = "Financial - Owner Detail (" & txtOwnerName.Text & ")"
                End If
            End If

            If FacilityID > 0 Then
                tbCntrlFinancial.SelectedTab = tbPageFacilityDetail
                PopulateFacilityInfo(Integer.Parse(FacilityID))
            End If

        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Public Sub New(ByRef oOwner As MUSTER.BusinessLogic.pOwner, ByVal OwnerID As Int64, ByVal FacilityID As Int64, ByVal FinancialEventID As Int64, Optional ByVal fromTechnical As Boolean = False)
        MyBase.New()
        MyGuid = System.Guid.NewGuid
        bolLoading = True
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        pOwn = oOwner
        'Add any initialization after the InitializeComponent() call
        MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "Financial")
        Try
            InitControls()
            bolFromTechnical = fromTechnical
            PopulateOwnerInfo(Integer.Parse(OwnerID))
            PopulateFacilityInfo(Integer.Parse(FacilityID))
            LoadFinancialData(FinancialEventID)

        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Public Sub New(Optional ByVal OwnerID As Int64 = 0, Optional ByVal FacilityID As Int64 = 0)
        MyBase.New()
        bolLoading = True
        MyGuid = System.Guid.NewGuid
        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "Financial")
        Try
            InitControls()
            If OwnerID > 0 Then
                PopulateOwnerInfo(Integer.Parse(OwnerID))
                If FacilityID <= 0 Then
                    tbCntrlFinancial.SelectedTab = tbPageOwnerDetail
                    Me.Text = "Financial - Owner Detail (" & txtOwnerName.Text & ")"
                End If
            End If

            If FacilityID > 0 Then
                tbCntrlFinancial.SelectedTab = tbPageFacilityDetail
                PopulateFacilityInfo(Integer.Parse(FacilityID))
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub New(ByRef oOwner As MUSTER.BusinessLogic.pOwner)
        MyBase.New()
        bolLoading = True
        MyGuid = System.Guid.NewGuid
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        pOwn = oOwner

        'Add any initialization after the InitializeComponent() call

        MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "Financial")
        'Ends here 
        Try

            InitControls()
            PopulateOwnerInfo(pOwn.ID)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        MusterContainer.AppSemaphores.Remove(MyGuid.ToString)
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
    Friend WithEvents tbCntrlFinancial As System.Windows.Forms.TabControl
    Friend WithEvents tbPageOwnerDetail As System.Windows.Forms.TabPage
    Friend WithEvents pnlOwnerBottom As System.Windows.Forms.Panel
    Friend WithEvents tbCtrlOwner As System.Windows.Forms.TabControl
    Friend WithEvents tbPageOwnerFacilities As System.Windows.Forms.TabPage
    Public WithEvents ugFacilityList As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlOwnerFacilityBottom As System.Windows.Forms.Panel
    Public WithEvents lblNoOfFacilitiesValue As System.Windows.Forms.Label
    Friend WithEvents lblNoOfFacilities As System.Windows.Forms.Label
    Friend WithEvents tbPageOwnerContactList As System.Windows.Forms.TabPage
    Friend WithEvents pnlOwnerContactButtons As System.Windows.Forms.Panel
    Friend WithEvents btnOwnerModifyContact As System.Windows.Forms.Button
    Friend WithEvents btnOwnerDeleteContact As System.Windows.Forms.Button
    Friend WithEvents btnOwnerAssociateContact As System.Windows.Forms.Button
    Friend WithEvents btnOwnerAddSearchContact As System.Windows.Forms.Button
    Friend WithEvents pnlOwnerContactContainer As System.Windows.Forms.Panel
    Friend WithEvents ugOwnerContacts As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents pnlOwnerContactHeader As System.Windows.Forms.Panel
    Friend WithEvents chkOwnerShowActiveOnly As System.Windows.Forms.CheckBox
    Friend WithEvents chkOwnerShowRelatedContacts As System.Windows.Forms.CheckBox
    Friend WithEvents chkOwnerShowContactsforAllModules As System.Windows.Forms.CheckBox
    Friend WithEvents lblOwnerContacts As System.Windows.Forms.Label
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
    Friend WithEvents tbPageFinancialEvent As System.Windows.Forms.TabPage
    Friend WithEvents tbPageSummary As System.Windows.Forms.TabPage
    Friend WithEvents Panel12 As System.Windows.Forms.Panel
    Friend WithEvents pnl_FacilityDetail As System.Windows.Forms.Panel
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
    Public WithEvents lblDateTransfered As System.Windows.Forms.Label
    Friend WithEvents lblLUSTSite As System.Windows.Forms.Label
    Public WithEvents chkLUSTSite As System.Windows.Forms.CheckBox
    Friend WithEvents lblPowerOff As System.Windows.Forms.Label
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
    Friend WithEvents txtFacilityFax As System.Windows.Forms.TextBox
    Friend WithEvents lblFacilityFax As System.Windows.Forms.Label
    Public WithEvents dtPickFacilityRecvd As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblDateReceived As System.Windows.Forms.Label
    Friend WithEvents txtFuelBrandcmb As System.Windows.Forms.ComboBox
    Friend WithEvents btnFacilityChangeCancel As System.Windows.Forms.Button
    Public WithEvents txtDueByNF As System.Windows.Forms.TextBox
    Friend WithEvents lblFacilitySigNFDue As System.Windows.Forms.Label
    Public WithEvents chkSignatureofNF As System.Windows.Forms.CheckBox
    Friend WithEvents lblPotentialOwner As System.Windows.Forms.Label
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
    Friend WithEvents txtfacilityPhone As System.Windows.Forms.TextBox
    Friend WithEvents lblFacilityPhone As System.Windows.Forms.Label
    Public WithEvents txtFacilityName As System.Windows.Forms.TextBox
    Friend WithEvents lblFacilityAddress As System.Windows.Forms.Label
    Friend WithEvents lblFacilityName As System.Windows.Forms.Label
    Friend WithEvents txtFacilityZip As System.Windows.Forms.TextBox
    Friend WithEvents lblFacilityLatDegree As System.Windows.Forms.Label
    Friend WithEvents pnlFacilityBottom As System.Windows.Forms.Panel
    Friend WithEvents ugFinancialGrid As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlFacilityFinancialButton As System.Windows.Forms.Panel
    Public WithEvents lblTotalNoOfFinancialEventsValue As System.Windows.Forms.Label
    Friend WithEvents lblTotalNoOfFinancialEvents As System.Windows.Forms.Label
    Friend WithEvents tbCtrlFacFinancialEvt As System.Windows.Forms.TabControl
    Friend WithEvents tbPageFacFinancialEvents As System.Windows.Forms.TabPage
    Friend WithEvents btnAddFinancialEvt As System.Windows.Forms.Button
    Friend WithEvents btnFacFEExpand As System.Windows.Forms.Button
    Friend WithEvents btnFacFECollapse As System.Windows.Forms.Button
    Friend WithEvents pnlFinancialHeader As System.Windows.Forms.Panel
    Friend WithEvents tbCtrlFinancialEvtDetails As System.Windows.Forms.TabControl
    Friend WithEvents tbPageFinancialEvtDetails As System.Windows.Forms.TabPage
    Friend WithEvents pnlFinEvtsDetails As System.Windows.Forms.Panel
    Friend WithEvents pnlFinEvtsBottom As System.Windows.Forms.Panel
    Friend WithEvents btnSaveEvent As System.Windows.Forms.Button
    Friend WithEvents btnDeleteEvent As System.Windows.Forms.Button
    Friend WithEvents btnCancelEvent As System.Windows.Forms.Button
    Friend WithEvents btnFinancialFlags As System.Windows.Forms.Button
    Friend WithEvents btnFinancialComments As System.Windows.Forms.Button
    Friend WithEvents lblFinancialID As System.Windows.Forms.Label
    Friend WithEvents lblFinancialCountVal As System.Windows.Forms.Label
    Friend WithEvents lblFinancialIDValue As System.Windows.Forms.Label
    Friend WithEvents lblFinancialStatus As System.Windows.Forms.Label
    Friend WithEvents cmbFinancialStatus As System.Windows.Forms.ComboBox
    Friend WithEvents PnlEvtInfo As System.Windows.Forms.Panel
    Friend WithEvents lblEvtInfoDisplay As System.Windows.Forms.Label
    Friend WithEvents lblEvtInfoHead As System.Windows.Forms.Label
    Friend WithEvents pnlEvtInfoDetails As System.Windows.Forms.Panel
    Friend WithEvents pnlCommitments As System.Windows.Forms.Panel
    Friend WithEvents lblCommitmentsHead As System.Windows.Forms.Label
    Friend WithEvents lblCommitmentsDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlCommitmentsDetails As System.Windows.Forms.Panel
    Friend WithEvents pnlPayments As System.Windows.Forms.Panel
    Friend WithEvents lblPaymentsHead As System.Windows.Forms.Label
    Friend WithEvents lblPaymentsDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlPaymentsDetails As System.Windows.Forms.Panel
    Friend WithEvents cmbTechEvent As System.Windows.Forms.ComboBox
    Friend WithEvents lblTechStartDate As System.Windows.Forms.Label
    Friend WithEvents lblPM As System.Windows.Forms.Label
    Friend WithEvents lblPMValue As System.Windows.Forms.Label
    Friend WithEvents lblTechStartDateValue As System.Windows.Forms.Label
    Friend WithEvents lblMGPTFStatus As System.Windows.Forms.Label
    Friend WithEvents lblMGPTFStatusValue As System.Windows.Forms.Label
    Friend WithEvents lblTechStatus As System.Windows.Forms.Label
    Friend WithEvents lblTechStatusValue As System.Windows.Forms.Label
    Friend WithEvents lblEngineeringFirm As System.Windows.Forms.Label
    Friend WithEvents lblEngineeringFirmValue As System.Windows.Forms.Label
    Friend WithEvents lblVendor As System.Windows.Forms.Label
    Friend WithEvents lblVendorNo As System.Windows.Forms.Label
    Friend WithEvents txtVendorNo As System.Windows.Forms.TextBox
    Friend WithEvents lblFinancialStartDate As System.Windows.Forms.Label
    Friend WithEvents lblVendorAddress As System.Windows.Forms.Label
    Friend WithEvents txtVendorAddress As System.Windows.Forms.TextBox
    Friend WithEvents pnlContacts As System.Windows.Forms.Panel
    Friend WithEvents lblContactsHead As System.Windows.Forms.Label
    Friend WithEvents lblContactsDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlContactDetails As System.Windows.Forms.Panel
    Friend WithEvents ugCommitments As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnCommitmentsExpand As System.Windows.Forms.Button
    Friend WithEvents btnCommitmentsCollapse As System.Windows.Forms.Button
    Friend WithEvents chkShowCommitments As System.Windows.Forms.CheckBox
    Friend WithEvents btnAddCommitment As System.Windows.Forms.Button
    Friend WithEvents btnModViewCommitment As System.Windows.Forms.Button
    Friend WithEvents btnDeleteCommitment As System.Windows.Forms.Button
    Friend WithEvents btnGenerateApprovalForm As System.Windows.Forms.Button
    Friend WithEvents PnlCommitmentButtons As System.Windows.Forms.Panel
    Friend WithEvents btnViewApprovalForm As System.Windows.Forms.Button
    Friend WithEvents btnAddAdjustment As System.Windows.Forms.Button
    Friend WithEvents btnDeleteAdjustment As System.Windows.Forms.Button
    Friend WithEvents pnlCommitmentsTotals As System.Windows.Forms.Panel
    Friend WithEvents lblTotals As System.Windows.Forms.Label
    Friend WithEvents lblCommitmentValue As System.Windows.Forms.Label
    Friend WithEvents lblAdjustmentValue As System.Windows.Forms.Label
    Friend WithEvents lblPaymentValue As System.Windows.Forms.Label
    Friend WithEvents lblBalanceValue As System.Windows.Forms.Label
    Friend WithEvents btnPaymentExpand As System.Windows.Forms.Button
    Friend WithEvents btnCollapse As System.Windows.Forms.Button
    Friend WithEvents ugPayments As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlPaymentTotals As System.Windows.Forms.Panel
    Friend WithEvents lblPaymentTotals As System.Windows.Forms.Label
    Friend WithEvents lblPaymentRequested As System.Windows.Forms.Label
    Friend WithEvents lblpaymentInv As System.Windows.Forms.Label
    Friend WithEvents lblPaid As System.Windows.Forms.Label
    Friend WithEvents btnAddRequest As System.Windows.Forms.Button
    Friend WithEvents btnIncompleteApplication As System.Windows.Forms.Button
    Friend WithEvents btnDeleteRequest As System.Windows.Forms.Button
    Friend WithEvents btnAddInvoice As System.Windows.Forms.Button
    Friend WithEvents btnGenerateNoticeOfReim As System.Windows.Forms.Button
    Friend WithEvents btnModifyViewInvoice As System.Windows.Forms.Button
    Friend WithEvents btnDeleteInvoice As System.Windows.Forms.Button
    Friend WithEvents btnViewNoticeOfReim As System.Windows.Forms.Button
    Friend WithEvents btnModViewAdjustment As System.Windows.Forms.Button
    Friend WithEvents btnGoToTechnical As System.Windows.Forms.Button
    Friend WithEvents pnlOwnerButtons As System.Windows.Forms.Panel
    Friend WithEvents btnOwnerFlag As System.Windows.Forms.Button
    Friend WithEvents btnOwnerComment As System.Windows.Forms.Button
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents btnFacComments As System.Windows.Forms.Button
    Friend WithEvents btnFacFlags As System.Windows.Forms.Button
    Friend WithEvents lblTechEventTitle As System.Windows.Forms.Label
    Friend WithEvents lblTechEvent As System.Windows.Forms.Label
    Friend WithEvents pnlPaymentButtons As System.Windows.Forms.Panel
    Friend WithEvents dtFinancialStart As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnVendorPack As System.Windows.Forms.Button
    Friend WithEvents pnlFinancialContactHeader As System.Windows.Forms.Panel
    Friend WithEvents pnlFinancialContactButtons As System.Windows.Forms.Panel
    Friend WithEvents pnlFinancialContactDetails As System.Windows.Forms.Panel
    Friend WithEvents chkFinancialShowRelatedContacts As System.Windows.Forms.CheckBox
    Friend WithEvents chkFinancialShowContactsForAllModules As System.Windows.Forms.CheckBox
    Friend WithEvents lblFinancialContacts As System.Windows.Forms.Label
    Friend WithEvents chkFinancialShowActive As System.Windows.Forms.CheckBox
    Friend WithEvents ugFinancialContacts As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnFinancialContactModify As System.Windows.Forms.Button
    Friend WithEvents btnFinancialContactDelete As System.Windows.Forms.Button
    Friend WithEvents btnFinancialContactAssociate As System.Windows.Forms.Button
    Friend WithEvents btnFinancialContactAddorSearch As System.Windows.Forms.Button
    Friend WithEvents txtVendor As System.Windows.Forms.TextBox
    Public WithEvents lblOwnerLastEditedOn As System.Windows.Forms.Label
    Public WithEvents lblOwnerLastEditedBy As System.Windows.Forms.Label
    Friend WithEvents txtVendorNumberValue As System.Windows.Forms.TextBox
    Friend WithEvents lblVendorNumber As System.Windows.Forms.Label
    Friend WithEvents tbPageOwnerDocuments As System.Windows.Forms.TabPage
    Friend WithEvents tbPageFacilityDocuments As System.Windows.Forms.TabPage
    Friend WithEvents UCFacilityDocuments As MUSTER.DocumentViewControl
    Friend WithEvents UCOwnerDocuments As MUSTER.DocumentViewControl
    Friend WithEvents pnlOwnerSummaryHeader As System.Windows.Forms.Panel
    Friend WithEvents pnlOwnerSummaryDetails As System.Windows.Forms.Panel
    Public WithEvents UCOwnerSummary As MUSTER.OwnerSummary
    Friend WithEvents btnFinOwnerLabels As System.Windows.Forms.Button
    Friend WithEvents btnFinOwnerEnvelopes As System.Windows.Forms.Button
    Friend WithEvents btnFinFacLabels As System.Windows.Forms.Button
    Friend WithEvents btnFinFacEnvelopes As System.Windows.Forms.Button
    Friend WithEvents btnFinLabels As System.Windows.Forms.Button
    Friend WithEvents btnFinEnvelopes As System.Windows.Forms.Button
    Public WithEvents txtFacilitySIC As System.Windows.Forms.Label
    Public WithEvents dtFinancialClosedDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblFinancialClosedDate As System.Windows.Forms.Label
    Friend WithEvents btnPlanning As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Financial))
        Me.pnlTop = New System.Windows.Forms.Panel
        Me.lblOwnerLastEditedOn = New System.Windows.Forms.Label
        Me.lblOwnerLastEditedBy = New System.Windows.Forms.Label
        Me.btnGoToTechnical = New System.Windows.Forms.Button
        Me.tbCntrlFinancial = New System.Windows.Forms.TabControl
        Me.tbPageOwnerDetail = New System.Windows.Forms.TabPage
        Me.pnlOwnerBottom = New System.Windows.Forms.Panel
        Me.tbCtrlOwner = New System.Windows.Forms.TabControl
        Me.tbPageOwnerFacilities = New System.Windows.Forms.TabPage
        Me.ugFacilityList = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlOwnerFacilityBottom = New System.Windows.Forms.Panel
        Me.lblNoOfFacilitiesValue = New System.Windows.Forms.Label
        Me.lblNoOfFacilities = New System.Windows.Forms.Label
        Me.tbPageOwnerContactList = New System.Windows.Forms.TabPage
        Me.pnlOwnerContactContainer = New System.Windows.Forms.Panel
        Me.ugOwnerContacts = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.Label1 = New System.Windows.Forms.Label
        Me.pnlOwnerContactHeader = New System.Windows.Forms.Panel
        Me.chkOwnerShowActiveOnly = New System.Windows.Forms.CheckBox
        Me.chkOwnerShowRelatedContacts = New System.Windows.Forms.CheckBox
        Me.chkOwnerShowContactsforAllModules = New System.Windows.Forms.CheckBox
        Me.lblOwnerContacts = New System.Windows.Forms.Label
        Me.pnlOwnerContactButtons = New System.Windows.Forms.Panel
        Me.btnOwnerModifyContact = New System.Windows.Forms.Button
        Me.btnOwnerDeleteContact = New System.Windows.Forms.Button
        Me.btnOwnerAssociateContact = New System.Windows.Forms.Button
        Me.btnOwnerAddSearchContact = New System.Windows.Forms.Button
        Me.tbPageOwnerDocuments = New System.Windows.Forms.TabPage
        Me.UCOwnerDocuments = New MUSTER.DocumentViewControl
        Me.pnlOwnerDetail = New System.Windows.Forms.Panel
        Me.btnFinOwnerLabels = New System.Windows.Forms.Button
        Me.btnFinOwnerEnvelopes = New System.Windows.Forms.Button
        Me.pnlOwnerButtons = New System.Windows.Forms.Panel
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
        Me.tbPageFacilityDetail = New System.Windows.Forms.TabPage
        Me.pnlFacilityBottom = New System.Windows.Forms.Panel
        Me.tbCtrlFacFinancialEvt = New System.Windows.Forms.TabControl
        Me.tbPageFacFinancialEvents = New System.Windows.Forms.TabPage
        Me.btnFacFECollapse = New System.Windows.Forms.Button
        Me.btnFacFEExpand = New System.Windows.Forms.Button
        Me.btnAddFinancialEvt = New System.Windows.Forms.Button
        Me.ugFinancialGrid = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlFacilityFinancialButton = New System.Windows.Forms.Panel
        Me.lblTotalNoOfFinancialEventsValue = New System.Windows.Forms.Label
        Me.lblTotalNoOfFinancialEvents = New System.Windows.Forms.Label
        Me.tbPageFacilityDocuments = New System.Windows.Forms.TabPage
        Me.UCFacilityDocuments = New MUSTER.DocumentViewControl
        Me.pnl_FacilityDetail = New System.Windows.Forms.Panel
        Me.btnFinFacLabels = New System.Windows.Forms.Button
        Me.btnFinFacEnvelopes = New System.Windows.Forms.Button
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.btnFacComments = New System.Windows.Forms.Button
        Me.btnFacFlags = New System.Windows.Forms.Button
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
        Me.lblDateTransfered = New System.Windows.Forms.Label
        Me.lblLUSTSite = New System.Windows.Forms.Label
        Me.chkLUSTSite = New System.Windows.Forms.CheckBox
        Me.lblPowerOff = New System.Windows.Forms.Label
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
        Me.txtFacilityFax = New System.Windows.Forms.TextBox
        Me.lblFacilityFax = New System.Windows.Forms.Label
        Me.dtPickFacilityRecvd = New System.Windows.Forms.DateTimePicker
        Me.lblDateReceived = New System.Windows.Forms.Label
        Me.txtFuelBrandcmb = New System.Windows.Forms.ComboBox
        Me.btnFacilityChangeCancel = New System.Windows.Forms.Button
        Me.txtDueByNF = New System.Windows.Forms.TextBox
        Me.lblFacilitySigNFDue = New System.Windows.Forms.Label
        Me.chkSignatureofNF = New System.Windows.Forms.CheckBox
        Me.lblPotentialOwner = New System.Windows.Forms.Label
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
        Me.txtfacilityPhone = New System.Windows.Forms.TextBox
        Me.lblFacilityPhone = New System.Windows.Forms.Label
        Me.txtFacilityName = New System.Windows.Forms.TextBox
        Me.lblFacilityAddress = New System.Windows.Forms.Label
        Me.lblFacilityName = New System.Windows.Forms.Label
        Me.txtFacilityZip = New System.Windows.Forms.TextBox
        Me.lblFacilityLatDegree = New System.Windows.Forms.Label
        Me.txtFacilitySIC = New System.Windows.Forms.Label
        Me.tbPageFinancialEvent = New System.Windows.Forms.TabPage
        Me.tbCtrlFinancialEvtDetails = New System.Windows.Forms.TabControl
        Me.tbPageFinancialEvtDetails = New System.Windows.Forms.TabPage
        Me.pnlFinEvtsDetails = New System.Windows.Forms.Panel
        Me.pnlContactDetails = New System.Windows.Forms.Panel
        Me.pnlFinancialContactDetails = New System.Windows.Forms.Panel
        Me.ugFinancialContacts = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlFinancialContactButtons = New System.Windows.Forms.Panel
        Me.btnFinancialContactModify = New System.Windows.Forms.Button
        Me.btnFinancialContactDelete = New System.Windows.Forms.Button
        Me.btnFinancialContactAssociate = New System.Windows.Forms.Button
        Me.btnFinancialContactAddorSearch = New System.Windows.Forms.Button
        Me.pnlFinancialContactHeader = New System.Windows.Forms.Panel
        Me.chkFinancialShowActive = New System.Windows.Forms.CheckBox
        Me.chkFinancialShowRelatedContacts = New System.Windows.Forms.CheckBox
        Me.chkFinancialShowContactsForAllModules = New System.Windows.Forms.CheckBox
        Me.lblFinancialContacts = New System.Windows.Forms.Label
        Me.pnlContacts = New System.Windows.Forms.Panel
        Me.lblContactsHead = New System.Windows.Forms.Label
        Me.lblContactsDisplay = New System.Windows.Forms.Label
        Me.pnlPaymentsDetails = New System.Windows.Forms.Panel
        Me.pnlPaymentButtons = New System.Windows.Forms.Panel
        Me.btnViewNoticeOfReim = New System.Windows.Forms.Button
        Me.btnDeleteInvoice = New System.Windows.Forms.Button
        Me.btnModifyViewInvoice = New System.Windows.Forms.Button
        Me.btnGenerateNoticeOfReim = New System.Windows.Forms.Button
        Me.btnAddInvoice = New System.Windows.Forms.Button
        Me.btnDeleteRequest = New System.Windows.Forms.Button
        Me.btnIncompleteApplication = New System.Windows.Forms.Button
        Me.btnAddRequest = New System.Windows.Forms.Button
        Me.pnlPaymentTotals = New System.Windows.Forms.Panel
        Me.lblPaid = New System.Windows.Forms.Label
        Me.lblpaymentInv = New System.Windows.Forms.Label
        Me.lblPaymentRequested = New System.Windows.Forms.Label
        Me.lblPaymentTotals = New System.Windows.Forms.Label
        Me.ugPayments = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.btnCollapse = New System.Windows.Forms.Button
        Me.btnPaymentExpand = New System.Windows.Forms.Button
        Me.pnlPayments = New System.Windows.Forms.Panel
        Me.lblPaymentsHead = New System.Windows.Forms.Label
        Me.lblPaymentsDisplay = New System.Windows.Forms.Label
        Me.pnlCommitmentsDetails = New System.Windows.Forms.Panel
        Me.pnlCommitmentsTotals = New System.Windows.Forms.Panel
        Me.lblBalanceValue = New System.Windows.Forms.Label
        Me.lblPaymentValue = New System.Windows.Forms.Label
        Me.lblAdjustmentValue = New System.Windows.Forms.Label
        Me.lblCommitmentValue = New System.Windows.Forms.Label
        Me.lblTotals = New System.Windows.Forms.Label
        Me.PnlCommitmentButtons = New System.Windows.Forms.Panel
        Me.btnDeleteAdjustment = New System.Windows.Forms.Button
        Me.btnModViewAdjustment = New System.Windows.Forms.Button
        Me.btnAddAdjustment = New System.Windows.Forms.Button
        Me.btnViewApprovalForm = New System.Windows.Forms.Button
        Me.btnGenerateApprovalForm = New System.Windows.Forms.Button
        Me.btnDeleteCommitment = New System.Windows.Forms.Button
        Me.btnModViewCommitment = New System.Windows.Forms.Button
        Me.btnAddCommitment = New System.Windows.Forms.Button
        Me.chkShowCommitments = New System.Windows.Forms.CheckBox
        Me.btnCommitmentsCollapse = New System.Windows.Forms.Button
        Me.btnCommitmentsExpand = New System.Windows.Forms.Button
        Me.ugCommitments = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlCommitments = New System.Windows.Forms.Panel
        Me.lblCommitmentsHead = New System.Windows.Forms.Label
        Me.lblCommitmentsDisplay = New System.Windows.Forms.Label
        Me.pnlEvtInfoDetails = New System.Windows.Forms.Panel
        Me.btnFinLabels = New System.Windows.Forms.Button
        Me.btnFinEnvelopes = New System.Windows.Forms.Button
        Me.lblVendorNumber = New System.Windows.Forms.Label
        Me.txtVendorNumberValue = New System.Windows.Forms.TextBox
        Me.txtVendor = New System.Windows.Forms.TextBox
        Me.btnVendorPack = New System.Windows.Forms.Button
        Me.txtVendorAddress = New System.Windows.Forms.TextBox
        Me.lblVendorAddress = New System.Windows.Forms.Label
        Me.dtFinancialStart = New System.Windows.Forms.DateTimePicker
        Me.lblFinancialStartDate = New System.Windows.Forms.Label
        Me.txtVendorNo = New System.Windows.Forms.TextBox
        Me.lblVendorNo = New System.Windows.Forms.Label
        Me.lblVendor = New System.Windows.Forms.Label
        Me.lblEngineeringFirmValue = New System.Windows.Forms.Label
        Me.lblEngineeringFirm = New System.Windows.Forms.Label
        Me.lblTechStatusValue = New System.Windows.Forms.Label
        Me.lblTechStatus = New System.Windows.Forms.Label
        Me.lblMGPTFStatusValue = New System.Windows.Forms.Label
        Me.lblMGPTFStatus = New System.Windows.Forms.Label
        Me.lblTechStartDateValue = New System.Windows.Forms.Label
        Me.lblPMValue = New System.Windows.Forms.Label
        Me.lblPM = New System.Windows.Forms.Label
        Me.lblTechStartDate = New System.Windows.Forms.Label
        Me.lblTechEventTitle = New System.Windows.Forms.Label
        Me.cmbTechEvent = New System.Windows.Forms.ComboBox
        Me.lblTechEvent = New System.Windows.Forms.Label
        Me.PnlEvtInfo = New System.Windows.Forms.Panel
        Me.lblEvtInfoHead = New System.Windows.Forms.Label
        Me.lblEvtInfoDisplay = New System.Windows.Forms.Label
        Me.pnlFinEvtsBottom = New System.Windows.Forms.Panel
        Me.btnPlanning = New System.Windows.Forms.Button
        Me.btnFinancialComments = New System.Windows.Forms.Button
        Me.btnFinancialFlags = New System.Windows.Forms.Button
        Me.btnCancelEvent = New System.Windows.Forms.Button
        Me.btnDeleteEvent = New System.Windows.Forms.Button
        Me.btnSaveEvent = New System.Windows.Forms.Button
        Me.pnlFinancialHeader = New System.Windows.Forms.Panel
        Me.lblFinancialClosedDate = New System.Windows.Forms.Label
        Me.dtFinancialClosedDate = New System.Windows.Forms.DateTimePicker
        Me.cmbFinancialStatus = New System.Windows.Forms.ComboBox
        Me.lblFinancialStatus = New System.Windows.Forms.Label
        Me.lblFinancialID = New System.Windows.Forms.Label
        Me.lblFinancialCountVal = New System.Windows.Forms.Label
        Me.lblFinancialIDValue = New System.Windows.Forms.Label
        Me.tbPageSummary = New System.Windows.Forms.TabPage
        Me.pnlOwnerSummaryDetails = New System.Windows.Forms.Panel
        Me.UCOwnerSummary = New MUSTER.OwnerSummary
        Me.Panel12 = New System.Windows.Forms.Panel
        Me.pnlOwnerSummaryHeader = New System.Windows.Forms.Panel
        Me.pnlTop.SuspendLayout()
        Me.tbCntrlFinancial.SuspendLayout()
        Me.tbPageOwnerDetail.SuspendLayout()
        Me.pnlOwnerBottom.SuspendLayout()
        Me.tbCtrlOwner.SuspendLayout()
        Me.tbPageOwnerFacilities.SuspendLayout()
        CType(Me.ugFacilityList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlOwnerFacilityBottom.SuspendLayout()
        Me.tbPageOwnerContactList.SuspendLayout()
        Me.pnlOwnerContactContainer.SuspendLayout()
        CType(Me.ugOwnerContacts, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlOwnerContactHeader.SuspendLayout()
        Me.pnlOwnerContactButtons.SuspendLayout()
        Me.tbPageOwnerDocuments.SuspendLayout()
        Me.pnlOwnerDetail.SuspendLayout()
        Me.pnlOwnerButtons.SuspendLayout()
        CType(Me.mskTxtOwnerFax, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtOwnerPhone2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtOwnerPhone, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbPageFacilityDetail.SuspendLayout()
        Me.pnlFacilityBottom.SuspendLayout()
        Me.tbCtrlFacFinancialEvt.SuspendLayout()
        Me.tbPageFacFinancialEvents.SuspendLayout()
        CType(Me.ugFinancialGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlFacilityFinancialButton.SuspendLayout()
        Me.tbPageFacilityDocuments.SuspendLayout()
        Me.pnl_FacilityDetail.SuspendLayout()
        Me.Panel2.SuspendLayout()
        CType(Me.mskTxtFacilityFax, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtFacilityPhone, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbPageFinancialEvent.SuspendLayout()
        Me.tbCtrlFinancialEvtDetails.SuspendLayout()
        Me.tbPageFinancialEvtDetails.SuspendLayout()
        Me.pnlFinEvtsDetails.SuspendLayout()
        Me.pnlContactDetails.SuspendLayout()
        Me.pnlFinancialContactDetails.SuspendLayout()
        CType(Me.ugFinancialContacts, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlFinancialContactButtons.SuspendLayout()
        Me.pnlFinancialContactHeader.SuspendLayout()
        Me.pnlContacts.SuspendLayout()
        Me.pnlPaymentsDetails.SuspendLayout()
        Me.pnlPaymentButtons.SuspendLayout()
        Me.pnlPaymentTotals.SuspendLayout()
        CType(Me.ugPayments, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlPayments.SuspendLayout()
        Me.pnlCommitmentsDetails.SuspendLayout()
        Me.pnlCommitmentsTotals.SuspendLayout()
        Me.PnlCommitmentButtons.SuspendLayout()
        CType(Me.ugCommitments, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlCommitments.SuspendLayout()
        Me.pnlEvtInfoDetails.SuspendLayout()
        Me.PnlEvtInfo.SuspendLayout()
        Me.pnlFinEvtsBottom.SuspendLayout()
        Me.pnlFinancialHeader.SuspendLayout()
        Me.tbPageSummary.SuspendLayout()
        Me.pnlOwnerSummaryDetails.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlTop
        '
        Me.pnlTop.BackColor = System.Drawing.SystemColors.Control
        Me.pnlTop.Controls.Add(Me.lblOwnerLastEditedOn)
        Me.pnlTop.Controls.Add(Me.lblOwnerLastEditedBy)
        Me.pnlTop.Controls.Add(Me.btnGoToTechnical)
        Me.pnlTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlTop.Name = "pnlTop"
        Me.pnlTop.Size = New System.Drawing.Size(1016, 24)
        Me.pnlTop.TabIndex = 2
        '
        'lblOwnerLastEditedOn
        '
        Me.lblOwnerLastEditedOn.Location = New System.Drawing.Point(688, 5)
        Me.lblOwnerLastEditedOn.Name = "lblOwnerLastEditedOn"
        Me.lblOwnerLastEditedOn.Size = New System.Drawing.Size(168, 16)
        Me.lblOwnerLastEditedOn.TabIndex = 1014
        Me.lblOwnerLastEditedOn.Text = "Last Edited On :"
        '
        'lblOwnerLastEditedBy
        '
        Me.lblOwnerLastEditedBy.Location = New System.Drawing.Point(472, 4)
        Me.lblOwnerLastEditedBy.Name = "lblOwnerLastEditedBy"
        Me.lblOwnerLastEditedBy.Size = New System.Drawing.Size(208, 16)
        Me.lblOwnerLastEditedBy.TabIndex = 1013
        Me.lblOwnerLastEditedBy.Text = "Last Edited By :"
        '
        'btnGoToTechnical
        '
        Me.btnGoToTechnical.Location = New System.Drawing.Point(864, 0)
        Me.btnGoToTechnical.Name = "btnGoToTechnical"
        Me.btnGoToTechnical.Size = New System.Drawing.Size(104, 23)
        Me.btnGoToTechnical.TabIndex = 2
        Me.btnGoToTechnical.Text = "Go To Technical"
        Me.btnGoToTechnical.Visible = False
        '
        'tbCntrlFinancial
        '
        Me.tbCntrlFinancial.Controls.Add(Me.tbPageOwnerDetail)
        Me.tbCntrlFinancial.Controls.Add(Me.tbPageFacilityDetail)
        Me.tbCntrlFinancial.Controls.Add(Me.tbPageFinancialEvent)
        Me.tbCntrlFinancial.Controls.Add(Me.tbPageSummary)
        Me.tbCntrlFinancial.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbCntrlFinancial.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbCntrlFinancial.ItemSize = New System.Drawing.Size(64, 18)
        Me.tbCntrlFinancial.Location = New System.Drawing.Point(0, 24)
        Me.tbCntrlFinancial.Multiline = True
        Me.tbCntrlFinancial.Name = "tbCntrlFinancial"
        Me.tbCntrlFinancial.SelectedIndex = 0
        Me.tbCntrlFinancial.ShowToolTips = True
        Me.tbCntrlFinancial.Size = New System.Drawing.Size(1016, 670)
        Me.tbCntrlFinancial.TabIndex = 3
        '
        'tbPageOwnerDetail
        '
        Me.tbPageOwnerDetail.Controls.Add(Me.pnlOwnerBottom)
        Me.tbPageOwnerDetail.Controls.Add(Me.pnlOwnerDetail)
        Me.tbPageOwnerDetail.Location = New System.Drawing.Point(4, 22)
        Me.tbPageOwnerDetail.Name = "tbPageOwnerDetail"
        Me.tbPageOwnerDetail.Size = New System.Drawing.Size(1008, 644)
        Me.tbPageOwnerDetail.TabIndex = 7
        Me.tbPageOwnerDetail.Text = "Owner Details"
        '
        'pnlOwnerBottom
        '
        Me.pnlOwnerBottom.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlOwnerBottom.Controls.Add(Me.tbCtrlOwner)
        Me.pnlOwnerBottom.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlOwnerBottom.Location = New System.Drawing.Point(0, 200)
        Me.pnlOwnerBottom.Name = "pnlOwnerBottom"
        Me.pnlOwnerBottom.Size = New System.Drawing.Size(1008, 444)
        Me.pnlOwnerBottom.TabIndex = 44
        '
        'tbCtrlOwner
        '
        Me.tbCtrlOwner.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
        Me.tbCtrlOwner.Controls.Add(Me.tbPageOwnerFacilities)
        Me.tbCtrlOwner.Controls.Add(Me.tbPageOwnerContactList)
        Me.tbCtrlOwner.Controls.Add(Me.tbPageOwnerDocuments)
        Me.tbCtrlOwner.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbCtrlOwner.Location = New System.Drawing.Point(0, 0)
        Me.tbCtrlOwner.Name = "tbCtrlOwner"
        Me.tbCtrlOwner.SelectedIndex = 0
        Me.tbCtrlOwner.Size = New System.Drawing.Size(1006, 442)
        Me.tbCtrlOwner.TabIndex = 9
        '
        'tbPageOwnerFacilities
        '
        Me.tbPageOwnerFacilities.BackColor = System.Drawing.SystemColors.Control
        Me.tbPageOwnerFacilities.Controls.Add(Me.ugFacilityList)
        Me.tbPageOwnerFacilities.Controls.Add(Me.pnlOwnerFacilityBottom)
        Me.tbPageOwnerFacilities.Location = New System.Drawing.Point(4, 27)
        Me.tbPageOwnerFacilities.Name = "tbPageOwnerFacilities"
        Me.tbPageOwnerFacilities.Size = New System.Drawing.Size(998, 411)
        Me.tbPageOwnerFacilities.TabIndex = 0
        Me.tbPageOwnerFacilities.Text = "Facilities"
        '
        'ugFacilityList
        '
        Me.ugFacilityList.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugFacilityList.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle
        Me.ugFacilityList.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugFacilityList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugFacilityList.Location = New System.Drawing.Point(0, 0)
        Me.ugFacilityList.Name = "ugFacilityList"
        Me.ugFacilityList.Size = New System.Drawing.Size(998, 387)
        Me.ugFacilityList.TabIndex = 20
        '
        'pnlOwnerFacilityBottom
        '
        Me.pnlOwnerFacilityBottom.Controls.Add(Me.lblNoOfFacilitiesValue)
        Me.pnlOwnerFacilityBottom.Controls.Add(Me.lblNoOfFacilities)
        Me.pnlOwnerFacilityBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlOwnerFacilityBottom.Location = New System.Drawing.Point(0, 387)
        Me.pnlOwnerFacilityBottom.Name = "pnlOwnerFacilityBottom"
        Me.pnlOwnerFacilityBottom.Size = New System.Drawing.Size(998, 24)
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
        'tbPageOwnerContactList
        '
        Me.tbPageOwnerContactList.Controls.Add(Me.pnlOwnerContactContainer)
        Me.tbPageOwnerContactList.Controls.Add(Me.pnlOwnerContactHeader)
        Me.tbPageOwnerContactList.Controls.Add(Me.pnlOwnerContactButtons)
        Me.tbPageOwnerContactList.Location = New System.Drawing.Point(4, 27)
        Me.tbPageOwnerContactList.Name = "tbPageOwnerContactList"
        Me.tbPageOwnerContactList.Size = New System.Drawing.Size(998, 411)
        Me.tbPageOwnerContactList.TabIndex = 1
        Me.tbPageOwnerContactList.Text = "Contacts"
        Me.tbPageOwnerContactList.Visible = False
        '
        'pnlOwnerContactContainer
        '
        Me.pnlOwnerContactContainer.Controls.Add(Me.ugOwnerContacts)
        Me.pnlOwnerContactContainer.Controls.Add(Me.Label1)
        Me.pnlOwnerContactContainer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlOwnerContactContainer.Location = New System.Drawing.Point(0, 25)
        Me.pnlOwnerContactContainer.Name = "pnlOwnerContactContainer"
        Me.pnlOwnerContactContainer.Size = New System.Drawing.Size(998, 356)
        Me.pnlOwnerContactContainer.TabIndex = 2
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
        Me.ugOwnerContacts.Size = New System.Drawing.Size(998, 356)
        Me.ugOwnerContacts.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(792, 56)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(7, 23)
        Me.Label1.TabIndex = 2
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
        Me.pnlOwnerContactHeader.Size = New System.Drawing.Size(998, 25)
        Me.pnlOwnerContactHeader.TabIndex = 0
        '
        'chkOwnerShowActiveOnly
        '
        Me.chkOwnerShowActiveOnly.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkOwnerShowActiveOnly.Location = New System.Drawing.Point(635, 6)
        Me.chkOwnerShowActiveOnly.Name = "chkOwnerShowActiveOnly"
        Me.chkOwnerShowActiveOnly.Size = New System.Drawing.Size(144, 16)
        Me.chkOwnerShowActiveOnly.TabIndex = 2
        Me.chkOwnerShowActiveOnly.Tag = "646"
        Me.chkOwnerShowActiveOnly.Text = "Show Active Only"
        '
        'chkOwnerShowRelatedContacts
        '
        Me.chkOwnerShowRelatedContacts.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkOwnerShowRelatedContacts.Location = New System.Drawing.Point(467, 6)
        Me.chkOwnerShowRelatedContacts.Name = "chkOwnerShowRelatedContacts"
        Me.chkOwnerShowRelatedContacts.Size = New System.Drawing.Size(160, 16)
        Me.chkOwnerShowRelatedContacts.TabIndex = 1
        Me.chkOwnerShowRelatedContacts.Tag = "645"
        Me.chkOwnerShowRelatedContacts.Text = "Show Related Contacts"
        '
        'chkOwnerShowContactsforAllModules
        '
        Me.chkOwnerShowContactsforAllModules.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkOwnerShowContactsforAllModules.Location = New System.Drawing.Point(251, 6)
        Me.chkOwnerShowContactsforAllModules.Name = "chkOwnerShowContactsforAllModules"
        Me.chkOwnerShowContactsforAllModules.Size = New System.Drawing.Size(200, 16)
        Me.chkOwnerShowContactsforAllModules.TabIndex = 0
        Me.chkOwnerShowContactsforAllModules.Tag = "644"
        Me.chkOwnerShowContactsforAllModules.Text = "Show Contacts for All Modules"
        '
        'lblOwnerContacts
        '
        Me.lblOwnerContacts.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerContacts.Location = New System.Drawing.Point(8, 6)
        Me.lblOwnerContacts.Name = "lblOwnerContacts"
        Me.lblOwnerContacts.Size = New System.Drawing.Size(104, 16)
        Me.lblOwnerContacts.TabIndex = 139
        Me.lblOwnerContacts.Text = "Owner Contacts"
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
        Me.pnlOwnerContactButtons.Size = New System.Drawing.Size(998, 30)
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
        'tbPageOwnerDocuments
        '
        Me.tbPageOwnerDocuments.Controls.Add(Me.UCOwnerDocuments)
        Me.tbPageOwnerDocuments.Location = New System.Drawing.Point(4, 27)
        Me.tbPageOwnerDocuments.Name = "tbPageOwnerDocuments"
        Me.tbPageOwnerDocuments.Size = New System.Drawing.Size(998, 411)
        Me.tbPageOwnerDocuments.TabIndex = 2
        Me.tbPageOwnerDocuments.Text = "Documents"
        '
        'UCOwnerDocuments
        '
        Me.UCOwnerDocuments.AutoScroll = True
        Me.UCOwnerDocuments.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UCOwnerDocuments.Location = New System.Drawing.Point(0, 0)
        Me.UCOwnerDocuments.Name = "UCOwnerDocuments"
        Me.UCOwnerDocuments.Size = New System.Drawing.Size(998, 411)
        Me.UCOwnerDocuments.TabIndex = 3
        '
        'pnlOwnerDetail
        '
        Me.pnlOwnerDetail.BackColor = System.Drawing.SystemColors.Control
        Me.pnlOwnerDetail.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlOwnerDetail.Controls.Add(Me.btnFinOwnerLabels)
        Me.pnlOwnerDetail.Controls.Add(Me.btnFinOwnerEnvelopes)
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
        Me.pnlOwnerDetail.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlOwnerDetail.Location = New System.Drawing.Point(0, 0)
        Me.pnlOwnerDetail.Name = "pnlOwnerDetail"
        Me.pnlOwnerDetail.Size = New System.Drawing.Size(1008, 200)
        Me.pnlOwnerDetail.TabIndex = 0
        '
        'btnFinOwnerLabels
        '
        Me.btnFinOwnerLabels.Location = New System.Drawing.Point(4, 119)
        Me.btnFinOwnerLabels.Name = "btnFinOwnerLabels"
        Me.btnFinOwnerLabels.Size = New System.Drawing.Size(70, 23)
        Me.btnFinOwnerLabels.TabIndex = 1064
        Me.btnFinOwnerLabels.Text = "Labels"
        '
        'btnFinOwnerEnvelopes
        '
        Me.btnFinOwnerEnvelopes.Location = New System.Drawing.Point(4, 88)
        Me.btnFinOwnerEnvelopes.Name = "btnFinOwnerEnvelopes"
        Me.btnFinOwnerEnvelopes.Size = New System.Drawing.Size(70, 23)
        Me.btnFinOwnerEnvelopes.TabIndex = 1063
        Me.btnFinOwnerEnvelopes.Text = "Envelopes"
        '
        'pnlOwnerButtons
        '
        Me.pnlOwnerButtons.Controls.Add(Me.btnOwnerFlag)
        Me.pnlOwnerButtons.Controls.Add(Me.btnOwnerComment)
        Me.pnlOwnerButtons.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.pnlOwnerButtons.Location = New System.Drawing.Point(400, 144)
        Me.pnlOwnerButtons.Name = "pnlOwnerButtons"
        Me.pnlOwnerButtons.Size = New System.Drawing.Size(176, 37)
        Me.pnlOwnerButtons.TabIndex = 1007
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
        Me.btnOwnerComment.Location = New System.Drawing.Point(88, 7)
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
        Me.lblCAPParticipationLevel.Size = New System.Drawing.Size(264, 16)
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
        'tbPageFacilityDetail
        '
        Me.tbPageFacilityDetail.Controls.Add(Me.pnlFacilityBottom)
        Me.tbPageFacilityDetail.Controls.Add(Me.pnl_FacilityDetail)
        Me.tbPageFacilityDetail.Location = New System.Drawing.Point(4, 22)
        Me.tbPageFacilityDetail.Name = "tbPageFacilityDetail"
        Me.tbPageFacilityDetail.Size = New System.Drawing.Size(1008, 644)
        Me.tbPageFacilityDetail.TabIndex = 8
        Me.tbPageFacilityDetail.Text = "Facility Details"
        Me.tbPageFacilityDetail.Visible = False
        '
        'pnlFacilityBottom
        '
        Me.pnlFacilityBottom.Controls.Add(Me.tbCtrlFacFinancialEvt)
        Me.pnlFacilityBottom.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlFacilityBottom.Location = New System.Drawing.Point(0, 256)
        Me.pnlFacilityBottom.Name = "pnlFacilityBottom"
        Me.pnlFacilityBottom.Size = New System.Drawing.Size(1008, 388)
        Me.pnlFacilityBottom.TabIndex = 3
        '
        'tbCtrlFacFinancialEvt
        '
        Me.tbCtrlFacFinancialEvt.Controls.Add(Me.tbPageFacFinancialEvents)
        Me.tbCtrlFacFinancialEvt.Controls.Add(Me.tbPageFacilityDocuments)
        Me.tbCtrlFacFinancialEvt.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbCtrlFacFinancialEvt.Location = New System.Drawing.Point(0, 0)
        Me.tbCtrlFacFinancialEvt.Name = "tbCtrlFacFinancialEvt"
        Me.tbCtrlFacFinancialEvt.SelectedIndex = 0
        Me.tbCtrlFacFinancialEvt.Size = New System.Drawing.Size(1008, 388)
        Me.tbCtrlFacFinancialEvt.TabIndex = 0
        '
        'tbPageFacFinancialEvents
        '
        Me.tbPageFacFinancialEvents.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.tbPageFacFinancialEvents.Controls.Add(Me.btnFacFECollapse)
        Me.tbPageFacFinancialEvents.Controls.Add(Me.btnFacFEExpand)
        Me.tbPageFacFinancialEvents.Controls.Add(Me.btnAddFinancialEvt)
        Me.tbPageFacFinancialEvents.Controls.Add(Me.ugFinancialGrid)
        Me.tbPageFacFinancialEvents.Controls.Add(Me.pnlFacilityFinancialButton)
        Me.tbPageFacFinancialEvents.Location = New System.Drawing.Point(4, 24)
        Me.tbPageFacFinancialEvents.Name = "tbPageFacFinancialEvents"
        Me.tbPageFacFinancialEvents.Size = New System.Drawing.Size(1000, 360)
        Me.tbPageFacFinancialEvents.TabIndex = 0
        Me.tbPageFacFinancialEvents.Text = "Financial Events"
        '
        'btnFacFECollapse
        '
        Me.btnFacFECollapse.Location = New System.Drawing.Point(193, 0)
        Me.btnFacFECollapse.Name = "btnFacFECollapse"
        Me.btnFacFECollapse.Size = New System.Drawing.Size(78, 23)
        Me.btnFacFECollapse.TabIndex = 101
        Me.btnFacFECollapse.Text = "Collapse All"
        '
        'btnFacFEExpand
        '
        Me.btnFacFEExpand.Location = New System.Drawing.Point(122, 0)
        Me.btnFacFEExpand.Name = "btnFacFEExpand"
        Me.btnFacFEExpand.Size = New System.Drawing.Size(72, 23)
        Me.btnFacFEExpand.TabIndex = 100
        Me.btnFacFEExpand.Text = "Expand All"
        '
        'btnAddFinancialEvt
        '
        Me.btnAddFinancialEvt.Location = New System.Drawing.Point(0, 0)
        Me.btnAddFinancialEvt.Name = "btnAddFinancialEvt"
        Me.btnAddFinancialEvt.Size = New System.Drawing.Size(122, 23)
        Me.btnAddFinancialEvt.TabIndex = 99
        Me.btnAddFinancialEvt.Text = "Add Financial Event"
        '
        'ugFinancialGrid
        '
        Me.ugFinancialGrid.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugFinancialGrid.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugFinancialGrid.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugFinancialGrid.Location = New System.Drawing.Point(0, 0)
        Me.ugFinancialGrid.Name = "ugFinancialGrid"
        Me.ugFinancialGrid.Size = New System.Drawing.Size(996, 332)
        Me.ugFinancialGrid.TabIndex = 0
        Me.ugFinancialGrid.Text = "Financial Events"
        '
        'pnlFacilityFinancialButton
        '
        Me.pnlFacilityFinancialButton.Controls.Add(Me.lblTotalNoOfFinancialEventsValue)
        Me.pnlFacilityFinancialButton.Controls.Add(Me.lblTotalNoOfFinancialEvents)
        Me.pnlFacilityFinancialButton.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlFacilityFinancialButton.Location = New System.Drawing.Point(0, 332)
        Me.pnlFacilityFinancialButton.Name = "pnlFacilityFinancialButton"
        Me.pnlFacilityFinancialButton.Size = New System.Drawing.Size(996, 24)
        Me.pnlFacilityFinancialButton.TabIndex = 98
        '
        'lblTotalNoOfFinancialEventsValue
        '
        Me.lblTotalNoOfFinancialEventsValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalNoOfFinancialEventsValue.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblTotalNoOfFinancialEventsValue.Location = New System.Drawing.Point(168, 0)
        Me.lblTotalNoOfFinancialEventsValue.Name = "lblTotalNoOfFinancialEventsValue"
        Me.lblTotalNoOfFinancialEventsValue.Size = New System.Drawing.Size(48, 24)
        Me.lblTotalNoOfFinancialEventsValue.TabIndex = 5
        Me.lblTotalNoOfFinancialEventsValue.Text = "0"
        '
        'lblTotalNoOfFinancialEvents
        '
        Me.lblTotalNoOfFinancialEvents.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalNoOfFinancialEvents.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblTotalNoOfFinancialEvents.Location = New System.Drawing.Point(0, 0)
        Me.lblTotalNoOfFinancialEvents.Name = "lblTotalNoOfFinancialEvents"
        Me.lblTotalNoOfFinancialEvents.Size = New System.Drawing.Size(168, 24)
        Me.lblTotalNoOfFinancialEvents.TabIndex = 4
        Me.lblTotalNoOfFinancialEvents.Text = "Number of Financial Events:"
        '
        'tbPageFacilityDocuments
        '
        Me.tbPageFacilityDocuments.Controls.Add(Me.UCFacilityDocuments)
        Me.tbPageFacilityDocuments.Location = New System.Drawing.Point(4, 24)
        Me.tbPageFacilityDocuments.Name = "tbPageFacilityDocuments"
        Me.tbPageFacilityDocuments.Size = New System.Drawing.Size(1000, 360)
        Me.tbPageFacilityDocuments.TabIndex = 1
        Me.tbPageFacilityDocuments.Text = "Documents"
        '
        'UCFacilityDocuments
        '
        Me.UCFacilityDocuments.AutoScroll = True
        Me.UCFacilityDocuments.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UCFacilityDocuments.Location = New System.Drawing.Point(0, 0)
        Me.UCFacilityDocuments.Name = "UCFacilityDocuments"
        Me.UCFacilityDocuments.Size = New System.Drawing.Size(1000, 360)
        Me.UCFacilityDocuments.TabIndex = 3
        '
        'pnl_FacilityDetail
        '
        Me.pnl_FacilityDetail.BackColor = System.Drawing.SystemColors.Control
        Me.pnl_FacilityDetail.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnl_FacilityDetail.Controls.Add(Me.btnFinFacLabels)
        Me.pnl_FacilityDetail.Controls.Add(Me.btnFinFacEnvelopes)
        Me.pnl_FacilityDetail.Controls.Add(Me.Panel2)
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
        Me.pnl_FacilityDetail.Controls.Add(Me.lblDateTransfered)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblLUSTSite)
        Me.pnl_FacilityDetail.Controls.Add(Me.chkLUSTSite)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblPowerOff)
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
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilityFax)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityFax)
        Me.pnl_FacilityDetail.Controls.Add(Me.dtPickFacilityRecvd)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblDateReceived)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFuelBrandcmb)
        Me.pnl_FacilityDetail.Controls.Add(Me.btnFacilityChangeCancel)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtDueByNF)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilitySigNFDue)
        Me.pnl_FacilityDetail.Controls.Add(Me.chkSignatureofNF)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblPotentialOwner)
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
        Me.pnl_FacilityDetail.Controls.Add(Me.txtfacilityPhone)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityPhone)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilityName)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityAddress)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityName)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilityZip)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityLatDegree)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilitySIC)
        Me.pnl_FacilityDetail.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnl_FacilityDetail.Location = New System.Drawing.Point(0, 0)
        Me.pnl_FacilityDetail.Name = "pnl_FacilityDetail"
        Me.pnl_FacilityDetail.Size = New System.Drawing.Size(1008, 256)
        Me.pnl_FacilityDetail.TabIndex = 2
        '
        'btnFinFacLabels
        '
        Me.btnFinFacLabels.Location = New System.Drawing.Point(8, 120)
        Me.btnFinFacLabels.Name = "btnFinFacLabels"
        Me.btnFinFacLabels.Size = New System.Drawing.Size(72, 23)
        Me.btnFinFacLabels.TabIndex = 1066
        Me.btnFinFacLabels.Text = "Labels"
        '
        'btnFinFacEnvelopes
        '
        Me.btnFinFacEnvelopes.Location = New System.Drawing.Point(8, 88)
        Me.btnFinFacEnvelopes.Name = "btnFinFacEnvelopes"
        Me.btnFinFacEnvelopes.Size = New System.Drawing.Size(72, 23)
        Me.btnFinFacEnvelopes.TabIndex = 1065
        Me.btnFinFacEnvelopes.Text = "Envelopes"
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.btnFacComments)
        Me.Panel2.Controls.Add(Me.btnFacFlags)
        Me.Panel2.Location = New System.Drawing.Point(608, 192)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(178, 32)
        Me.Panel2.TabIndex = 1045
        '
        'btnFacComments
        '
        Me.btnFacComments.Location = New System.Drawing.Point(88, 5)
        Me.btnFacComments.Name = "btnFacComments"
        Me.btnFacComments.Size = New System.Drawing.Size(80, 23)
        Me.btnFacComments.TabIndex = 1039
        Me.btnFacComments.Text = "Comments"
        '
        'btnFacFlags
        '
        Me.btnFacFlags.Location = New System.Drawing.Point(8, 5)
        Me.btnFacFlags.Name = "btnFacFlags"
        Me.btnFacFlags.TabIndex = 1040
        Me.btnFacFlags.Text = "Flags"
        '
        'dtPickUpcomingInstallDateValue
        '
        Me.dtPickUpcomingInstallDateValue.Checked = False
        Me.dtPickUpcomingInstallDateValue.Enabled = False
        Me.dtPickUpcomingInstallDateValue.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickUpcomingInstallDateValue.Location = New System.Drawing.Point(491, 198)
        Me.dtPickUpcomingInstallDateValue.Name = "dtPickUpcomingInstallDateValue"
        Me.dtPickUpcomingInstallDateValue.ShowCheckBox = True
        Me.dtPickUpcomingInstallDateValue.Size = New System.Drawing.Size(101, 21)
        Me.dtPickUpcomingInstallDateValue.TabIndex = 12
        '
        'lblUpcomingInstallDate
        '
        Me.lblUpcomingInstallDate.Location = New System.Drawing.Point(336, 198)
        Me.lblUpcomingInstallDate.Name = "lblUpcomingInstallDate"
        Me.lblUpcomingInstallDate.Size = New System.Drawing.Size(160, 23)
        Me.lblUpcomingInstallDate.TabIndex = 1044
        Me.lblUpcomingInstallDate.Text = "Upcoming Installation Date"
        '
        'chkUpcomingInstall
        '
        Me.chkUpcomingInstall.Location = New System.Drawing.Point(332, 171)
        Me.chkUpcomingInstall.Name = "chkUpcomingInstall"
        Me.chkUpcomingInstall.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkUpcomingInstall.Size = New System.Drawing.Size(144, 24)
        Me.chkUpcomingInstall.TabIndex = 11
        Me.chkUpcomingInstall.Text = "Upcoming Installation "
        '
        'lnkLblNextFac
        '
        Me.lnkLblNextFac.Location = New System.Drawing.Point(701, 229)
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
        Me.lblCAPStatusValue.Location = New System.Drawing.Point(424, 75)
        Me.lblCAPStatusValue.Name = "lblCAPStatusValue"
        Me.lblCAPStatusValue.Size = New System.Drawing.Size(120, 16)
        Me.lblCAPStatusValue.TabIndex = 1038
        '
        'lblCAPStatus
        '
        Me.lblCAPStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCAPStatus.Location = New System.Drawing.Point(336, 75)
        Me.lblCAPStatus.Name = "lblCAPStatus"
        Me.lblCAPStatus.Size = New System.Drawing.Size(72, 16)
        Me.lblCAPStatus.TabIndex = 1037
        Me.lblCAPStatus.Text = "CAP Status:"
        '
        'txtFuelBrand
        '
        Me.txtFuelBrand.Location = New System.Drawing.Point(424, 150)
        Me.txtFuelBrand.Name = "txtFuelBrand"
        Me.txtFuelBrand.Size = New System.Drawing.Size(72, 21)
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
        Me.dtFacilityPowerOff.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtFacilityPowerOff.Location = New System.Drawing.Point(424, 224)
        Me.dtFacilityPowerOff.Name = "dtFacilityPowerOff"
        Me.dtFacilityPowerOff.ShowCheckBox = True
        Me.dtFacilityPowerOff.Size = New System.Drawing.Size(104, 21)
        Me.dtFacilityPowerOff.TabIndex = 9
        Me.dtFacilityPowerOff.Visible = False
        '
        'lnkLblPrevFacility
        '
        Me.lnkLblPrevFacility.AutoSize = True
        Me.lnkLblPrevFacility.Location = New System.Drawing.Point(621, 229)
        Me.lnkLblPrevFacility.Name = "lnkLblPrevFacility"
        Me.lnkLblPrevFacility.Size = New System.Drawing.Size(70, 17)
        Me.lnkLblPrevFacility.TabIndex = 25
        Me.lnkLblPrevFacility.TabStop = True
        Me.lnkLblPrevFacility.Text = "<< Previous"
        '
        'lblDateTransfered
        '
        Me.lblDateTransfered.Location = New System.Drawing.Point(896, 152)
        Me.lblDateTransfered.Name = "lblDateTransfered"
        Me.lblDateTransfered.Size = New System.Drawing.Size(32, 16)
        Me.lblDateTransfered.TabIndex = 1034
        Me.lblDateTransfered.Visible = False
        '
        'lblLUSTSite
        '
        Me.lblLUSTSite.Location = New System.Drawing.Point(536, 64)
        Me.lblLUSTSite.Name = "lblLUSTSite"
        Me.lblLUSTSite.Size = New System.Drawing.Size(24, 16)
        Me.lblLUSTSite.TabIndex = 1030
        Me.lblLUSTSite.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblLUSTSite.Visible = False
        '
        'chkLUSTSite
        '
        Me.chkLUSTSite.CheckAlign = System.Drawing.ContentAlignment.TopLeft
        Me.chkLUSTSite.Location = New System.Drawing.Point(328, 52)
        Me.chkLUSTSite.Name = "chkLUSTSite"
        Me.chkLUSTSite.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkLUSTSite.Size = New System.Drawing.Size(120, 16)
        Me.chkLUSTSite.TabIndex = 6
        Me.chkLUSTSite.Text = "Active LUST Site"
        Me.chkLUSTSite.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'lblPowerOff
        '
        Me.lblPowerOff.Location = New System.Drawing.Point(325, 224)
        Me.lblPowerOff.Name = "lblPowerOff"
        Me.lblPowerOff.Size = New System.Drawing.Size(72, 16)
        Me.lblPowerOff.TabIndex = 1028
        Me.lblPowerOff.Text = "Power Off"
        Me.lblPowerOff.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblPowerOff.Visible = False
        '
        'chkCAPCandidate
        '
        Me.chkCAPCandidate.CheckAlign = System.Drawing.ContentAlignment.TopLeft
        Me.chkCAPCandidate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkCAPCandidate.Location = New System.Drawing.Point(328, 96)
        Me.chkCAPCandidate.Name = "chkCAPCandidate"
        Me.chkCAPCandidate.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkCAPCandidate.Size = New System.Drawing.Size(120, 16)
        Me.chkCAPCandidate.TabIndex = 7
        Me.chkCAPCandidate.Text = "CAP Candidate"
        Me.chkCAPCandidate.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'lblFacilityLocationType
        '
        Me.lblFacilityLocationType.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLocationType.Location = New System.Drawing.Point(600, 150)
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
        Me.cmbFacilityLocationType.Size = New System.Drawing.Size(120, 23)
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
        Me.cmbFacilityMethod.Size = New System.Drawing.Size(120, 23)
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
        Me.cmbFacilityDatum.Size = New System.Drawing.Size(120, 23)
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
        Me.cmbFacilityType.Size = New System.Drawing.Size(136, 23)
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
        Me.lblFacilitySIC.Location = New System.Drawing.Point(338, 120)
        Me.lblFacilitySIC.Name = "lblFacilitySIC"
        Me.lblFacilitySIC.Size = New System.Drawing.Size(80, 23)
        Me.lblFacilitySIC.TabIndex = 150
        Me.lblFacilitySIC.Text = "SIC:"
        '
        'txtFacilityFax
        '
        Me.txtFacilityFax.AcceptsTab = True
        Me.txtFacilityFax.AutoSize = False
        Me.txtFacilityFax.Location = New System.Drawing.Point(928, 208)
        Me.txtFacilityFax.Name = "txtFacilityFax"
        Me.txtFacilityFax.Size = New System.Drawing.Size(104, 21)
        Me.txtFacilityFax.TabIndex = 148
        Me.txtFacilityFax.Text = ""
        Me.txtFacilityFax.Visible = False
        Me.txtFacilityFax.WordWrap = False
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
        Me.dtPickFacilityRecvd.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickFacilityRecvd.Location = New System.Drawing.Point(424, 9)
        Me.dtPickFacilityRecvd.Name = "dtPickFacilityRecvd"
        Me.dtPickFacilityRecvd.ShowCheckBox = True
        Me.dtPickFacilityRecvd.Size = New System.Drawing.Size(104, 21)
        Me.dtPickFacilityRecvd.TabIndex = 5
        '
        'lblDateReceived
        '
        Me.lblDateReceived.Location = New System.Drawing.Point(336, 9)
        Me.lblDateReceived.Name = "lblDateReceived"
        Me.lblDateReceived.TabIndex = 145
        Me.lblDateReceived.Text = "Date Received:"
        '
        'txtFuelBrandcmb
        '
        Me.txtFuelBrandcmb.DropDownStyle = System.Windows.Forms.ComboBoxStyle.Simple
        Me.txtFuelBrandcmb.ItemHeight = 15
        Me.txtFuelBrandcmb.Items.AddRange(New Object() {"Shell", "Exon", "Texaco", "Mobil"})
        Me.txtFuelBrandcmb.Location = New System.Drawing.Point(872, 64)
        Me.txtFuelBrandcmb.Name = "txtFuelBrandcmb"
        Me.txtFuelBrandcmb.Size = New System.Drawing.Size(96, 23)
        Me.txtFuelBrandcmb.TabIndex = 23
        Me.txtFuelBrandcmb.Text = "Remove after cure period"
        Me.txtFuelBrandcmb.Visible = False
        '
        'btnFacilityChangeCancel
        '
        Me.btnFacilityChangeCancel.Enabled = False
        Me.btnFacilityChangeCancel.Location = New System.Drawing.Point(880, 176)
        Me.btnFacilityChangeCancel.Name = "btnFacilityChangeCancel"
        Me.btnFacilityChangeCancel.TabIndex = 30
        Me.btnFacilityChangeCancel.Text = "Cancel"
        Me.btnFacilityChangeCancel.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnFacilityChangeCancel.Visible = False
        '
        'txtDueByNF
        '
        Me.txtDueByNF.AcceptsTab = True
        Me.txtDueByNF.AutoSize = False
        Me.txtDueByNF.Enabled = False
        Me.txtDueByNF.Location = New System.Drawing.Point(936, 176)
        Me.txtDueByNF.Name = "txtDueByNF"
        Me.txtDueByNF.Size = New System.Drawing.Size(64, 21)
        Me.txtDueByNF.TabIndex = 136
        Me.txtDueByNF.Text = ""
        Me.txtDueByNF.Visible = False
        Me.txtDueByNF.WordWrap = False
        '
        'lblFacilitySigNFDue
        '
        Me.lblFacilitySigNFDue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilitySigNFDue.ForeColor = System.Drawing.Color.Red
        Me.lblFacilitySigNFDue.Location = New System.Drawing.Point(856, 160)
        Me.lblFacilitySigNFDue.Name = "lblFacilitySigNFDue"
        Me.lblFacilitySigNFDue.Size = New System.Drawing.Size(40, 23)
        Me.lblFacilitySigNFDue.TabIndex = 134
        Me.lblFacilitySigNFDue.Text = "Due By:"
        Me.lblFacilitySigNFDue.Visible = False
        '
        'chkSignatureofNF
        '
        Me.chkSignatureofNF.CheckAlign = System.Drawing.ContentAlignment.TopLeft
        Me.chkSignatureofNF.Location = New System.Drawing.Point(8, 217)
        Me.chkSignatureofNF.Name = "chkSignatureofNF"
        Me.chkSignatureofNF.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkSignatureofNF.Size = New System.Drawing.Size(144, 16)
        Me.chkSignatureofNF.TabIndex = 4
        Me.chkSignatureofNF.Text = ": Signature Received"
        Me.chkSignatureofNF.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'lblPotentialOwner
        '
        Me.lblPotentialOwner.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPotentialOwner.ForeColor = System.Drawing.Color.Red
        Me.lblPotentialOwner.Location = New System.Drawing.Point(912, 208)
        Me.lblPotentialOwner.Name = "lblPotentialOwner"
        Me.lblPotentialOwner.Size = New System.Drawing.Size(128, 23)
        Me.lblPotentialOwner.TabIndex = 129
        Me.lblPotentialOwner.Text = "Potential Owner:"
        Me.lblPotentialOwner.Visible = False
        '
        'lblFacilitySigOnNF
        '
        Me.lblFacilitySigOnNF.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilitySigOnNF.Location = New System.Drawing.Point(8, 216)
        Me.lblFacilitySigOnNF.Name = "lblFacilitySigOnNF"
        Me.lblFacilitySigOnNF.Size = New System.Drawing.Size(120, 23)
        Me.lblFacilitySigOnNF.TabIndex = 127
        Me.lblFacilitySigOnNF.Text = "Signature Received:"
        '
        'lblFacilityFuelBrand
        '
        Me.lblFacilityFuelBrand.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityFuelBrand.Location = New System.Drawing.Point(337, 150)
        Me.lblFacilityFuelBrand.Name = "lblFacilityFuelBrand"
        Me.lblFacilityFuelBrand.Size = New System.Drawing.Size(96, 23)
        Me.lblFacilityFuelBrand.TabIndex = 125
        Me.lblFacilityFuelBrand.Text = "Fuel Brand:"
        '
        'lblFacilityStatusValue
        '
        Me.lblFacilityStatusValue.BackColor = System.Drawing.SystemColors.Control
        Me.lblFacilityStatusValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblFacilityStatusValue.Enabled = False
        Me.lblFacilityStatusValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityStatusValue.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblFacilityStatusValue.Location = New System.Drawing.Point(424, 33)
        Me.lblFacilityStatusValue.Name = "lblFacilityStatusValue"
        Me.lblFacilityStatusValue.Size = New System.Drawing.Size(120, 16)
        Me.lblFacilityStatusValue.TabIndex = 124
        '
        'lblFacilityStatus
        '
        Me.lblFacilityStatus.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityStatus.Location = New System.Drawing.Point(336, 33)
        Me.lblFacilityStatus.Name = "lblFacilityStatus"
        Me.lblFacilityStatus.Size = New System.Drawing.Size(88, 23)
        Me.lblFacilityStatus.TabIndex = 123
        Me.lblFacilityStatus.Text = "Facility Status:"
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
        Me.lblFacilityLongitude.Location = New System.Drawing.Point(592, 80)
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
        'txtfacilityPhone
        '
        Me.txtfacilityPhone.AcceptsTab = True
        Me.txtfacilityPhone.AutoSize = False
        Me.txtfacilityPhone.Location = New System.Drawing.Point(928, 184)
        Me.txtfacilityPhone.Name = "txtfacilityPhone"
        Me.txtfacilityPhone.Size = New System.Drawing.Size(104, 21)
        Me.txtfacilityPhone.TabIndex = 99
        Me.txtfacilityPhone.Text = ""
        Me.txtfacilityPhone.Visible = False
        Me.txtfacilityPhone.WordWrap = False
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
        Me.txtFacilityName.Enabled = False
        Me.txtFacilityName.Location = New System.Drawing.Point(88, 32)
        Me.txtFacilityName.Name = "txtFacilityName"
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
        'txtFacilityZip
        '
        Me.txtFacilityZip.Location = New System.Drawing.Point(896, 32)
        Me.txtFacilityZip.Name = "txtFacilityZip"
        Me.txtFacilityZip.Size = New System.Drawing.Size(96, 21)
        Me.txtFacilityZip.TabIndex = 11
        Me.txtFacilityZip.Text = ""
        Me.txtFacilityZip.Visible = False
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
        'txtFacilitySIC
        '
        Me.txtFacilitySIC.BackColor = System.Drawing.SystemColors.Control
        Me.txtFacilitySIC.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFacilitySIC.Enabled = False
        Me.txtFacilitySIC.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFacilitySIC.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.txtFacilitySIC.Location = New System.Drawing.Point(424, 120)
        Me.txtFacilitySIC.Name = "txtFacilitySIC"
        Me.txtFacilitySIC.Size = New System.Drawing.Size(120, 16)
        Me.txtFacilitySIC.TabIndex = 1038
        '
        'tbPageFinancialEvent
        '
        Me.tbPageFinancialEvent.Controls.Add(Me.tbCtrlFinancialEvtDetails)
        Me.tbPageFinancialEvent.Location = New System.Drawing.Point(4, 22)
        Me.tbPageFinancialEvent.Name = "tbPageFinancialEvent"
        Me.tbPageFinancialEvent.Size = New System.Drawing.Size(1008, 644)
        Me.tbPageFinancialEvent.TabIndex = 12
        Me.tbPageFinancialEvent.Text = "Financial Events"
        Me.tbPageFinancialEvent.Visible = False
        '
        'tbCtrlFinancialEvtDetails
        '
        Me.tbCtrlFinancialEvtDetails.Controls.Add(Me.tbPageFinancialEvtDetails)
        Me.tbCtrlFinancialEvtDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbCtrlFinancialEvtDetails.Location = New System.Drawing.Point(0, 0)
        Me.tbCtrlFinancialEvtDetails.Name = "tbCtrlFinancialEvtDetails"
        Me.tbCtrlFinancialEvtDetails.SelectedIndex = 0
        Me.tbCtrlFinancialEvtDetails.Size = New System.Drawing.Size(1008, 644)
        Me.tbCtrlFinancialEvtDetails.TabIndex = 2
        '
        'tbPageFinancialEvtDetails
        '
        Me.tbPageFinancialEvtDetails.Controls.Add(Me.pnlFinEvtsDetails)
        Me.tbPageFinancialEvtDetails.Controls.Add(Me.pnlFinEvtsBottom)
        Me.tbPageFinancialEvtDetails.Controls.Add(Me.pnlFinancialHeader)
        Me.tbPageFinancialEvtDetails.Location = New System.Drawing.Point(4, 24)
        Me.tbPageFinancialEvtDetails.Name = "tbPageFinancialEvtDetails"
        Me.tbPageFinancialEvtDetails.Size = New System.Drawing.Size(1000, 616)
        Me.tbPageFinancialEvtDetails.TabIndex = 0
        Me.tbPageFinancialEvtDetails.Text = "Financial Event Details"
        '
        'pnlFinEvtsDetails
        '
        Me.pnlFinEvtsDetails.AutoScroll = True
        Me.pnlFinEvtsDetails.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlFinEvtsDetails.Controls.Add(Me.pnlContactDetails)
        Me.pnlFinEvtsDetails.Controls.Add(Me.pnlContacts)
        Me.pnlFinEvtsDetails.Controls.Add(Me.pnlPaymentsDetails)
        Me.pnlFinEvtsDetails.Controls.Add(Me.pnlPayments)
        Me.pnlFinEvtsDetails.Controls.Add(Me.pnlCommitmentsDetails)
        Me.pnlFinEvtsDetails.Controls.Add(Me.pnlCommitments)
        Me.pnlFinEvtsDetails.Controls.Add(Me.pnlEvtInfoDetails)
        Me.pnlFinEvtsDetails.Controls.Add(Me.PnlEvtInfo)
        Me.pnlFinEvtsDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlFinEvtsDetails.Location = New System.Drawing.Point(0, 32)
        Me.pnlFinEvtsDetails.Name = "pnlFinEvtsDetails"
        Me.pnlFinEvtsDetails.Size = New System.Drawing.Size(1000, 544)
        Me.pnlFinEvtsDetails.TabIndex = 1
        '
        'pnlContactDetails
        '
        Me.pnlContactDetails.Controls.Add(Me.pnlFinancialContactDetails)
        Me.pnlContactDetails.Controls.Add(Me.pnlFinancialContactButtons)
        Me.pnlContactDetails.Controls.Add(Me.pnlFinancialContactHeader)
        Me.pnlContactDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlContactDetails.Location = New System.Drawing.Point(0, 1042)
        Me.pnlContactDetails.Name = "pnlContactDetails"
        Me.pnlContactDetails.Size = New System.Drawing.Size(980, 296)
        Me.pnlContactDetails.TabIndex = 30
        '
        'pnlFinancialContactDetails
        '
        Me.pnlFinancialContactDetails.Controls.Add(Me.ugFinancialContacts)
        Me.pnlFinancialContactDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlFinancialContactDetails.Location = New System.Drawing.Point(0, 27)
        Me.pnlFinancialContactDetails.Name = "pnlFinancialContactDetails"
        Me.pnlFinancialContactDetails.Size = New System.Drawing.Size(980, 229)
        Me.pnlFinancialContactDetails.TabIndex = 2
        '
        'ugFinancialContacts
        '
        Me.ugFinancialContacts.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugFinancialContacts.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugFinancialContacts.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        Me.ugFinancialContacts.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugFinancialContacts.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugFinancialContacts.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugFinancialContacts.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugFinancialContacts.Location = New System.Drawing.Point(0, 0)
        Me.ugFinancialContacts.Name = "ugFinancialContacts"
        Me.ugFinancialContacts.Size = New System.Drawing.Size(980, 229)
        Me.ugFinancialContacts.TabIndex = 1
        '
        'pnlFinancialContactButtons
        '
        Me.pnlFinancialContactButtons.Controls.Add(Me.btnFinancialContactModify)
        Me.pnlFinancialContactButtons.Controls.Add(Me.btnFinancialContactDelete)
        Me.pnlFinancialContactButtons.Controls.Add(Me.btnFinancialContactAssociate)
        Me.pnlFinancialContactButtons.Controls.Add(Me.btnFinancialContactAddorSearch)
        Me.pnlFinancialContactButtons.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlFinancialContactButtons.Location = New System.Drawing.Point(0, 256)
        Me.pnlFinancialContactButtons.Name = "pnlFinancialContactButtons"
        Me.pnlFinancialContactButtons.Size = New System.Drawing.Size(980, 40)
        Me.pnlFinancialContactButtons.TabIndex = 1
        '
        'btnFinancialContactModify
        '
        Me.btnFinancialContactModify.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFinancialContactModify.Location = New System.Drawing.Point(240, 6)
        Me.btnFinancialContactModify.Name = "btnFinancialContactModify"
        Me.btnFinancialContactModify.Size = New System.Drawing.Size(235, 26)
        Me.btnFinancialContactModify.TabIndex = 5
        Me.btnFinancialContactModify.Text = "Modify Contact"
        '
        'btnFinancialContactDelete
        '
        Me.btnFinancialContactDelete.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFinancialContactDelete.Location = New System.Drawing.Point(472, 6)
        Me.btnFinancialContactDelete.Name = "btnFinancialContactDelete"
        Me.btnFinancialContactDelete.Size = New System.Drawing.Size(235, 26)
        Me.btnFinancialContactDelete.TabIndex = 6
        Me.btnFinancialContactDelete.Text = "Disassociate Contact"
        '
        'btnFinancialContactAssociate
        '
        Me.btnFinancialContactAssociate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFinancialContactAssociate.Location = New System.Drawing.Point(704, 6)
        Me.btnFinancialContactAssociate.Name = "btnFinancialContactAssociate"
        Me.btnFinancialContactAssociate.Size = New System.Drawing.Size(235, 26)
        Me.btnFinancialContactAssociate.TabIndex = 7
        Me.btnFinancialContactAssociate.Text = "Associate Contact from Different Module"
        '
        'btnFinancialContactAddorSearch
        '
        Me.btnFinancialContactAddorSearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFinancialContactAddorSearch.Location = New System.Drawing.Point(8, 6)
        Me.btnFinancialContactAddorSearch.Name = "btnFinancialContactAddorSearch"
        Me.btnFinancialContactAddorSearch.Size = New System.Drawing.Size(235, 26)
        Me.btnFinancialContactAddorSearch.TabIndex = 4
        Me.btnFinancialContactAddorSearch.Text = "Add/Search Contact to Associate"
        '
        'pnlFinancialContactHeader
        '
        Me.pnlFinancialContactHeader.Controls.Add(Me.chkFinancialShowActive)
        Me.pnlFinancialContactHeader.Controls.Add(Me.chkFinancialShowRelatedContacts)
        Me.pnlFinancialContactHeader.Controls.Add(Me.chkFinancialShowContactsForAllModules)
        Me.pnlFinancialContactHeader.Controls.Add(Me.lblFinancialContacts)
        Me.pnlFinancialContactHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlFinancialContactHeader.Location = New System.Drawing.Point(0, 0)
        Me.pnlFinancialContactHeader.Name = "pnlFinancialContactHeader"
        Me.pnlFinancialContactHeader.Size = New System.Drawing.Size(980, 27)
        Me.pnlFinancialContactHeader.TabIndex = 0
        '
        'chkFinancialShowActive
        '
        Me.chkFinancialShowActive.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkFinancialShowActive.Location = New System.Drawing.Point(776, 6)
        Me.chkFinancialShowActive.Name = "chkFinancialShowActive"
        Me.chkFinancialShowActive.Size = New System.Drawing.Size(144, 16)
        Me.chkFinancialShowActive.TabIndex = 143
        Me.chkFinancialShowActive.Tag = "646"
        Me.chkFinancialShowActive.Text = "Show Active Only"
        '
        'chkFinancialShowRelatedContacts
        '
        Me.chkFinancialShowRelatedContacts.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkFinancialShowRelatedContacts.Location = New System.Drawing.Point(616, 6)
        Me.chkFinancialShowRelatedContacts.Name = "chkFinancialShowRelatedContacts"
        Me.chkFinancialShowRelatedContacts.Size = New System.Drawing.Size(160, 16)
        Me.chkFinancialShowRelatedContacts.TabIndex = 141
        Me.chkFinancialShowRelatedContacts.Tag = "645"
        Me.chkFinancialShowRelatedContacts.Text = "Show Related Contacts"
        '
        'chkFinancialShowContactsForAllModules
        '
        Me.chkFinancialShowContactsForAllModules.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkFinancialShowContactsForAllModules.Location = New System.Drawing.Point(416, 6)
        Me.chkFinancialShowContactsForAllModules.Name = "chkFinancialShowContactsForAllModules"
        Me.chkFinancialShowContactsForAllModules.Size = New System.Drawing.Size(200, 16)
        Me.chkFinancialShowContactsForAllModules.TabIndex = 140
        Me.chkFinancialShowContactsForAllModules.Tag = "644"
        Me.chkFinancialShowContactsForAllModules.Text = "Show Contacts for All Modules"
        '
        'lblFinancialContacts
        '
        Me.lblFinancialContacts.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFinancialContacts.Location = New System.Drawing.Point(8, 3)
        Me.lblFinancialContacts.Name = "lblFinancialContacts"
        Me.lblFinancialContacts.Size = New System.Drawing.Size(112, 16)
        Me.lblFinancialContacts.TabIndex = 142
        Me.lblFinancialContacts.Text = "Financial Contacts"
        '
        'pnlContacts
        '
        Me.pnlContacts.Controls.Add(Me.lblContactsHead)
        Me.pnlContacts.Controls.Add(Me.lblContactsDisplay)
        Me.pnlContacts.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlContacts.Location = New System.Drawing.Point(0, 1018)
        Me.pnlContacts.Name = "pnlContacts"
        Me.pnlContacts.Size = New System.Drawing.Size(980, 24)
        Me.pnlContacts.TabIndex = 7
        '
        'lblContactsHead
        '
        Me.lblContactsHead.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblContactsHead.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblContactsHead.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblContactsHead.Location = New System.Drawing.Point(16, 0)
        Me.lblContactsHead.Name = "lblContactsHead"
        Me.lblContactsHead.Size = New System.Drawing.Size(964, 24)
        Me.lblContactsHead.TabIndex = 5
        Me.lblContactsHead.Text = "Contacts"
        Me.lblContactsHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblContactsDisplay
        '
        Me.lblContactsDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblContactsDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblContactsDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblContactsDisplay.Name = "lblContactsDisplay"
        Me.lblContactsDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblContactsDisplay.TabIndex = 4
        Me.lblContactsDisplay.Text = "-"
        Me.lblContactsDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlPaymentsDetails
        '
        Me.pnlPaymentsDetails.Controls.Add(Me.pnlPaymentButtons)
        Me.pnlPaymentsDetails.Controls.Add(Me.pnlPaymentTotals)
        Me.pnlPaymentsDetails.Controls.Add(Me.ugPayments)
        Me.pnlPaymentsDetails.Controls.Add(Me.btnCollapse)
        Me.pnlPaymentsDetails.Controls.Add(Me.btnPaymentExpand)
        Me.pnlPaymentsDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlPaymentsDetails.Location = New System.Drawing.Point(0, 618)
        Me.pnlPaymentsDetails.Name = "pnlPaymentsDetails"
        Me.pnlPaymentsDetails.Size = New System.Drawing.Size(980, 400)
        Me.pnlPaymentsDetails.TabIndex = 18
        '
        'pnlPaymentButtons
        '
        Me.pnlPaymentButtons.Controls.Add(Me.btnViewNoticeOfReim)
        Me.pnlPaymentButtons.Controls.Add(Me.btnDeleteInvoice)
        Me.pnlPaymentButtons.Controls.Add(Me.btnModifyViewInvoice)
        Me.pnlPaymentButtons.Controls.Add(Me.btnGenerateNoticeOfReim)
        Me.pnlPaymentButtons.Controls.Add(Me.btnAddInvoice)
        Me.pnlPaymentButtons.Controls.Add(Me.btnDeleteRequest)
        Me.pnlPaymentButtons.Controls.Add(Me.btnIncompleteApplication)
        Me.pnlPaymentButtons.Controls.Add(Me.btnAddRequest)
        Me.pnlPaymentButtons.Location = New System.Drawing.Point(16, 328)
        Me.pnlPaymentButtons.Name = "pnlPaymentButtons"
        Me.pnlPaymentButtons.Size = New System.Drawing.Size(936, 68)
        Me.pnlPaymentButtons.TabIndex = 5
        '
        'btnViewNoticeOfReim
        '
        Me.btnViewNoticeOfReim.Location = New System.Drawing.Point(720, 8)
        Me.btnViewNoticeOfReim.Name = "btnViewNoticeOfReim"
        Me.btnViewNoticeOfReim.Size = New System.Drawing.Size(160, 23)
        Me.btnViewNoticeOfReim.TabIndex = 27
        Me.btnViewNoticeOfReim.Text = "View Notice"
        '
        'btnDeleteInvoice
        '
        Me.btnDeleteInvoice.Location = New System.Drawing.Point(554, 40)
        Me.btnDeleteInvoice.Name = "btnDeleteInvoice"
        Me.btnDeleteInvoice.Size = New System.Drawing.Size(160, 23)
        Me.btnDeleteInvoice.TabIndex = 29
        Me.btnDeleteInvoice.Text = "Delete Invoice"
        '
        'btnModifyViewInvoice
        '
        Me.btnModifyViewInvoice.Location = New System.Drawing.Point(388, 40)
        Me.btnModifyViewInvoice.Name = "btnModifyViewInvoice"
        Me.btnModifyViewInvoice.Size = New System.Drawing.Size(160, 23)
        Me.btnModifyViewInvoice.TabIndex = 28
        Me.btnModifyViewInvoice.Text = "Modify/View Invoice"
        '
        'btnGenerateNoticeOfReim
        '
        Me.btnGenerateNoticeOfReim.Location = New System.Drawing.Point(554, 8)
        Me.btnGenerateNoticeOfReim.Name = "btnGenerateNoticeOfReim"
        Me.btnGenerateNoticeOfReim.Size = New System.Drawing.Size(160, 23)
        Me.btnGenerateNoticeOfReim.TabIndex = 26
        Me.btnGenerateNoticeOfReim.Text = "Generate Notice"
        '
        'btnAddInvoice
        '
        Me.btnAddInvoice.Location = New System.Drawing.Point(222, 40)
        Me.btnAddInvoice.Name = "btnAddInvoice"
        Me.btnAddInvoice.Size = New System.Drawing.Size(160, 23)
        Me.btnAddInvoice.TabIndex = 25
        Me.btnAddInvoice.Text = "Add Invoice"
        '
        'btnDeleteRequest
        '
        Me.btnDeleteRequest.Location = New System.Drawing.Point(388, 8)
        Me.btnDeleteRequest.Name = "btnDeleteRequest"
        Me.btnDeleteRequest.Size = New System.Drawing.Size(160, 23)
        Me.btnDeleteRequest.TabIndex = 24
        Me.btnDeleteRequest.Text = "Delete Request"
        '
        'btnIncompleteApplication
        '
        Me.btnIncompleteApplication.Location = New System.Drawing.Point(222, 8)
        Me.btnIncompleteApplication.Name = "btnIncompleteApplication"
        Me.btnIncompleteApplication.Size = New System.Drawing.Size(160, 23)
        Me.btnIncompleteApplication.TabIndex = 23
        Me.btnIncompleteApplication.Text = "Modify Request"
        '
        'btnAddRequest
        '
        Me.btnAddRequest.Location = New System.Drawing.Point(56, 8)
        Me.btnAddRequest.Name = "btnAddRequest"
        Me.btnAddRequest.Size = New System.Drawing.Size(160, 23)
        Me.btnAddRequest.TabIndex = 22
        Me.btnAddRequest.Text = "Add Request"
        '
        'pnlPaymentTotals
        '
        Me.pnlPaymentTotals.Controls.Add(Me.lblPaid)
        Me.pnlPaymentTotals.Controls.Add(Me.lblpaymentInv)
        Me.pnlPaymentTotals.Controls.Add(Me.lblPaymentRequested)
        Me.pnlPaymentTotals.Controls.Add(Me.lblPaymentTotals)
        Me.pnlPaymentTotals.Location = New System.Drawing.Point(80, 296)
        Me.pnlPaymentTotals.Name = "pnlPaymentTotals"
        Me.pnlPaymentTotals.Size = New System.Drawing.Size(416, 32)
        Me.pnlPaymentTotals.TabIndex = 4
        '
        'lblPaid
        '
        Me.lblPaid.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblPaid.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPaid.Location = New System.Drawing.Point(272, 5)
        Me.lblPaid.Name = "lblPaid"
        Me.lblPaid.TabIndex = 3
        Me.lblPaid.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblpaymentInv
        '
        Me.lblpaymentInv.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblpaymentInv.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblpaymentInv.Location = New System.Drawing.Point(176, 5)
        Me.lblpaymentInv.Name = "lblpaymentInv"
        Me.lblpaymentInv.TabIndex = 2
        Me.lblpaymentInv.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPaymentRequested
        '
        Me.lblPaymentRequested.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblPaymentRequested.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPaymentRequested.Location = New System.Drawing.Point(78, 5)
        Me.lblPaymentRequested.Name = "lblPaymentRequested"
        Me.lblPaymentRequested.TabIndex = 1
        Me.lblPaymentRequested.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPaymentTotals
        '
        Me.lblPaymentTotals.Location = New System.Drawing.Point(8, 8)
        Me.lblPaymentTotals.Name = "lblPaymentTotals"
        Me.lblPaymentTotals.Size = New System.Drawing.Size(40, 17)
        Me.lblPaymentTotals.TabIndex = 0
        Me.lblPaymentTotals.Text = "Totals:"
        '
        'ugPayments
        '
        Me.ugPayments.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugPayments.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugPayments.Location = New System.Drawing.Point(16, 40)
        Me.ugPayments.Name = "ugPayments"
        Me.ugPayments.Size = New System.Drawing.Size(936, 250)
        Me.ugPayments.TabIndex = 21
        '
        'btnCollapse
        '
        Me.btnCollapse.Location = New System.Drawing.Point(95, 8)
        Me.btnCollapse.Name = "btnCollapse"
        Me.btnCollapse.Size = New System.Drawing.Size(78, 23)
        Me.btnCollapse.TabIndex = 20
        Me.btnCollapse.Text = "Collapse All"
        '
        'btnPaymentExpand
        '
        Me.btnPaymentExpand.Location = New System.Drawing.Point(16, 8)
        Me.btnPaymentExpand.Name = "btnPaymentExpand"
        Me.btnPaymentExpand.TabIndex = 19
        Me.btnPaymentExpand.Text = "Expand All"
        '
        'pnlPayments
        '
        Me.pnlPayments.Controls.Add(Me.lblPaymentsHead)
        Me.pnlPayments.Controls.Add(Me.lblPaymentsDisplay)
        Me.pnlPayments.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlPayments.Location = New System.Drawing.Point(0, 594)
        Me.pnlPayments.Name = "pnlPayments"
        Me.pnlPayments.Size = New System.Drawing.Size(980, 24)
        Me.pnlPayments.TabIndex = 5
        '
        'lblPaymentsHead
        '
        Me.lblPaymentsHead.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblPaymentsHead.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblPaymentsHead.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblPaymentsHead.Location = New System.Drawing.Point(16, 0)
        Me.lblPaymentsHead.Name = "lblPaymentsHead"
        Me.lblPaymentsHead.Size = New System.Drawing.Size(964, 24)
        Me.lblPaymentsHead.TabIndex = 3
        Me.lblPaymentsHead.Text = "Payments"
        Me.lblPaymentsHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblPaymentsDisplay
        '
        Me.lblPaymentsDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPaymentsDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblPaymentsDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblPaymentsDisplay.Name = "lblPaymentsDisplay"
        Me.lblPaymentsDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblPaymentsDisplay.TabIndex = 2
        Me.lblPaymentsDisplay.Text = "-"
        Me.lblPaymentsDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlCommitmentsDetails
        '
        Me.pnlCommitmentsDetails.Controls.Add(Me.pnlCommitmentsTotals)
        Me.pnlCommitmentsDetails.Controls.Add(Me.PnlCommitmentButtons)
        Me.pnlCommitmentsDetails.Controls.Add(Me.chkShowCommitments)
        Me.pnlCommitmentsDetails.Controls.Add(Me.btnCommitmentsCollapse)
        Me.pnlCommitmentsDetails.Controls.Add(Me.btnCommitmentsExpand)
        Me.pnlCommitmentsDetails.Controls.Add(Me.ugCommitments)
        Me.pnlCommitmentsDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlCommitmentsDetails.Location = New System.Drawing.Point(0, 194)
        Me.pnlCommitmentsDetails.Name = "pnlCommitmentsDetails"
        Me.pnlCommitmentsDetails.Size = New System.Drawing.Size(980, 400)
        Me.pnlCommitmentsDetails.TabIndex = 5
        '
        'pnlCommitmentsTotals
        '
        Me.pnlCommitmentsTotals.Controls.Add(Me.lblBalanceValue)
        Me.pnlCommitmentsTotals.Controls.Add(Me.lblPaymentValue)
        Me.pnlCommitmentsTotals.Controls.Add(Me.lblAdjustmentValue)
        Me.pnlCommitmentsTotals.Controls.Add(Me.lblCommitmentValue)
        Me.pnlCommitmentsTotals.Controls.Add(Me.lblTotals)
        Me.pnlCommitmentsTotals.Location = New System.Drawing.Point(410, 296)
        Me.pnlCommitmentsTotals.Name = "pnlCommitmentsTotals"
        Me.pnlCommitmentsTotals.Size = New System.Drawing.Size(488, 32)
        Me.pnlCommitmentsTotals.TabIndex = 5
        '
        'lblBalanceValue
        '
        Me.lblBalanceValue.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblBalanceValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblBalanceValue.Location = New System.Drawing.Point(353, 6)
        Me.lblBalanceValue.Name = "lblBalanceValue"
        Me.lblBalanceValue.TabIndex = 4
        Me.lblBalanceValue.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPaymentValue
        '
        Me.lblPaymentValue.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblPaymentValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPaymentValue.Location = New System.Drawing.Point(254, 6)
        Me.lblPaymentValue.Name = "lblPaymentValue"
        Me.lblPaymentValue.TabIndex = 3
        Me.lblPaymentValue.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblAdjustmentValue
        '
        Me.lblAdjustmentValue.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblAdjustmentValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblAdjustmentValue.Location = New System.Drawing.Point(155, 6)
        Me.lblAdjustmentValue.Name = "lblAdjustmentValue"
        Me.lblAdjustmentValue.TabIndex = 2
        Me.lblAdjustmentValue.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCommitmentValue
        '
        Me.lblCommitmentValue.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblCommitmentValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblCommitmentValue.Location = New System.Drawing.Point(56, 6)
        Me.lblCommitmentValue.Name = "lblCommitmentValue"
        Me.lblCommitmentValue.TabIndex = 1
        Me.lblCommitmentValue.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblTotals
        '
        Me.lblTotals.Location = New System.Drawing.Point(8, 8)
        Me.lblTotals.Name = "lblTotals"
        Me.lblTotals.Size = New System.Drawing.Size(48, 17)
        Me.lblTotals.TabIndex = 0
        Me.lblTotals.Text = "Totals:"
        '
        'PnlCommitmentButtons
        '
        Me.PnlCommitmentButtons.Controls.Add(Me.btnDeleteAdjustment)
        Me.PnlCommitmentButtons.Controls.Add(Me.btnModViewAdjustment)
        Me.PnlCommitmentButtons.Controls.Add(Me.btnAddAdjustment)
        Me.PnlCommitmentButtons.Controls.Add(Me.btnViewApprovalForm)
        Me.PnlCommitmentButtons.Controls.Add(Me.btnGenerateApprovalForm)
        Me.PnlCommitmentButtons.Controls.Add(Me.btnDeleteCommitment)
        Me.PnlCommitmentButtons.Controls.Add(Me.btnModViewCommitment)
        Me.PnlCommitmentButtons.Controls.Add(Me.btnAddCommitment)
        Me.PnlCommitmentButtons.Location = New System.Drawing.Point(8, 328)
        Me.PnlCommitmentButtons.Name = "PnlCommitmentButtons"
        Me.PnlCommitmentButtons.Size = New System.Drawing.Size(936, 67)
        Me.PnlCommitmentButtons.TabIndex = 4
        '
        'btnDeleteAdjustment
        '
        Me.btnDeleteAdjustment.Location = New System.Drawing.Point(554, 40)
        Me.btnDeleteAdjustment.Name = "btnDeleteAdjustment"
        Me.btnDeleteAdjustment.Size = New System.Drawing.Size(160, 23)
        Me.btnDeleteAdjustment.TabIndex = 17
        Me.btnDeleteAdjustment.Text = "Delete Adjustment"
        '
        'btnModViewAdjustment
        '
        Me.btnModViewAdjustment.Location = New System.Drawing.Point(388, 40)
        Me.btnModViewAdjustment.Name = "btnModViewAdjustment"
        Me.btnModViewAdjustment.Size = New System.Drawing.Size(160, 23)
        Me.btnModViewAdjustment.TabIndex = 16
        Me.btnModViewAdjustment.Text = "Modify/View Adjustment"
        '
        'btnAddAdjustment
        '
        Me.btnAddAdjustment.Location = New System.Drawing.Point(222, 40)
        Me.btnAddAdjustment.Name = "btnAddAdjustment"
        Me.btnAddAdjustment.Size = New System.Drawing.Size(160, 23)
        Me.btnAddAdjustment.TabIndex = 15
        Me.btnAddAdjustment.Text = "Add Adjustment"
        '
        'btnViewApprovalForm
        '
        Me.btnViewApprovalForm.Location = New System.Drawing.Point(720, 8)
        Me.btnViewApprovalForm.Name = "btnViewApprovalForm"
        Me.btnViewApprovalForm.Size = New System.Drawing.Size(160, 23)
        Me.btnViewApprovalForm.TabIndex = 14
        Me.btnViewApprovalForm.Text = "View Approval Form"
        '
        'btnGenerateApprovalForm
        '
        Me.btnGenerateApprovalForm.Location = New System.Drawing.Point(554, 8)
        Me.btnGenerateApprovalForm.Name = "btnGenerateApprovalForm"
        Me.btnGenerateApprovalForm.Size = New System.Drawing.Size(160, 23)
        Me.btnGenerateApprovalForm.TabIndex = 13
        Me.btnGenerateApprovalForm.Text = "Generate Approval Form"
        '
        'btnDeleteCommitment
        '
        Me.btnDeleteCommitment.Location = New System.Drawing.Point(388, 8)
        Me.btnDeleteCommitment.Name = "btnDeleteCommitment"
        Me.btnDeleteCommitment.Size = New System.Drawing.Size(160, 23)
        Me.btnDeleteCommitment.TabIndex = 12
        Me.btnDeleteCommitment.Text = "Delete Commitment"
        '
        'btnModViewCommitment
        '
        Me.btnModViewCommitment.Location = New System.Drawing.Point(222, 8)
        Me.btnModViewCommitment.Name = "btnModViewCommitment"
        Me.btnModViewCommitment.Size = New System.Drawing.Size(160, 23)
        Me.btnModViewCommitment.TabIndex = 11
        Me.btnModViewCommitment.Text = "Modify/View Commitment"
        '
        'btnAddCommitment
        '
        Me.btnAddCommitment.Location = New System.Drawing.Point(56, 8)
        Me.btnAddCommitment.Name = "btnAddCommitment"
        Me.btnAddCommitment.Size = New System.Drawing.Size(160, 23)
        Me.btnAddCommitment.TabIndex = 10
        Me.btnAddCommitment.Text = "Add Commitment"
        '
        'chkShowCommitments
        '
        Me.chkShowCommitments.Location = New System.Drawing.Point(512, 8)
        Me.chkShowCommitments.Name = "chkShowCommitments"
        Me.chkShowCommitments.Size = New System.Drawing.Size(200, 24)
        Me.chkShowCommitments.TabIndex = 8
        Me.chkShowCommitments.Text = "Show Open Commitments Only"
        '
        'btnCommitmentsCollapse
        '
        Me.btnCommitmentsCollapse.Location = New System.Drawing.Point(93, 8)
        Me.btnCommitmentsCollapse.Name = "btnCommitmentsCollapse"
        Me.btnCommitmentsCollapse.Size = New System.Drawing.Size(78, 23)
        Me.btnCommitmentsCollapse.TabIndex = 7
        Me.btnCommitmentsCollapse.Text = "Collapse All"
        '
        'btnCommitmentsExpand
        '
        Me.btnCommitmentsExpand.Location = New System.Drawing.Point(16, 8)
        Me.btnCommitmentsExpand.Name = "btnCommitmentsExpand"
        Me.btnCommitmentsExpand.TabIndex = 6
        Me.btnCommitmentsExpand.Text = "Expand All"
        '
        'ugCommitments
        '
        Me.ugCommitments.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugCommitments.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugCommitments.Location = New System.Drawing.Point(16, 32)
        Me.ugCommitments.Name = "ugCommitments"
        Me.ugCommitments.Size = New System.Drawing.Size(936, 264)
        Me.ugCommitments.TabIndex = 9
        '
        'pnlCommitments
        '
        Me.pnlCommitments.Controls.Add(Me.lblCommitmentsHead)
        Me.pnlCommitments.Controls.Add(Me.lblCommitmentsDisplay)
        Me.pnlCommitments.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlCommitments.Location = New System.Drawing.Point(0, 170)
        Me.pnlCommitments.Name = "pnlCommitments"
        Me.pnlCommitments.Size = New System.Drawing.Size(980, 24)
        Me.pnlCommitments.TabIndex = 3
        '
        'lblCommitmentsHead
        '
        Me.lblCommitmentsHead.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblCommitmentsHead.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblCommitmentsHead.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblCommitmentsHead.Location = New System.Drawing.Point(16, 0)
        Me.lblCommitmentsHead.Name = "lblCommitmentsHead"
        Me.lblCommitmentsHead.Size = New System.Drawing.Size(964, 24)
        Me.lblCommitmentsHead.TabIndex = 3
        Me.lblCommitmentsHead.Text = "Commitments"
        Me.lblCommitmentsHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblCommitmentsDisplay
        '
        Me.lblCommitmentsDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblCommitmentsDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblCommitmentsDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblCommitmentsDisplay.Name = "lblCommitmentsDisplay"
        Me.lblCommitmentsDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblCommitmentsDisplay.TabIndex = 2
        Me.lblCommitmentsDisplay.Text = "-"
        Me.lblCommitmentsDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlEvtInfoDetails
        '
        Me.pnlEvtInfoDetails.Controls.Add(Me.btnFinLabels)
        Me.pnlEvtInfoDetails.Controls.Add(Me.btnFinEnvelopes)
        Me.pnlEvtInfoDetails.Controls.Add(Me.lblVendorNumber)
        Me.pnlEvtInfoDetails.Controls.Add(Me.txtVendorNumberValue)
        Me.pnlEvtInfoDetails.Controls.Add(Me.txtVendor)
        Me.pnlEvtInfoDetails.Controls.Add(Me.btnVendorPack)
        Me.pnlEvtInfoDetails.Controls.Add(Me.txtVendorAddress)
        Me.pnlEvtInfoDetails.Controls.Add(Me.lblVendorAddress)
        Me.pnlEvtInfoDetails.Controls.Add(Me.dtFinancialStart)
        Me.pnlEvtInfoDetails.Controls.Add(Me.lblFinancialStartDate)
        Me.pnlEvtInfoDetails.Controls.Add(Me.txtVendorNo)
        Me.pnlEvtInfoDetails.Controls.Add(Me.lblVendorNo)
        Me.pnlEvtInfoDetails.Controls.Add(Me.lblVendor)
        Me.pnlEvtInfoDetails.Controls.Add(Me.lblEngineeringFirmValue)
        Me.pnlEvtInfoDetails.Controls.Add(Me.lblEngineeringFirm)
        Me.pnlEvtInfoDetails.Controls.Add(Me.lblTechStatusValue)
        Me.pnlEvtInfoDetails.Controls.Add(Me.lblTechStatus)
        Me.pnlEvtInfoDetails.Controls.Add(Me.lblMGPTFStatusValue)
        Me.pnlEvtInfoDetails.Controls.Add(Me.lblMGPTFStatus)
        Me.pnlEvtInfoDetails.Controls.Add(Me.lblTechStartDateValue)
        Me.pnlEvtInfoDetails.Controls.Add(Me.lblPMValue)
        Me.pnlEvtInfoDetails.Controls.Add(Me.lblPM)
        Me.pnlEvtInfoDetails.Controls.Add(Me.lblTechStartDate)
        Me.pnlEvtInfoDetails.Controls.Add(Me.lblTechEventTitle)
        Me.pnlEvtInfoDetails.Controls.Add(Me.cmbTechEvent)
        Me.pnlEvtInfoDetails.Controls.Add(Me.lblTechEvent)
        Me.pnlEvtInfoDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlEvtInfoDetails.Location = New System.Drawing.Point(0, 24)
        Me.pnlEvtInfoDetails.Name = "pnlEvtInfoDetails"
        Me.pnlEvtInfoDetails.Size = New System.Drawing.Size(980, 146)
        Me.pnlEvtInfoDetails.TabIndex = 1
        '
        'btnFinLabels
        '
        Me.btnFinLabels.Location = New System.Drawing.Point(35, 120)
        Me.btnFinLabels.Name = "btnFinLabels"
        Me.btnFinLabels.Size = New System.Drawing.Size(72, 20)
        Me.btnFinLabels.TabIndex = 1068
        Me.btnFinLabels.Text = "Labels"
        '
        'btnFinEnvelopes
        '
        Me.btnFinEnvelopes.Location = New System.Drawing.Point(35, 97)
        Me.btnFinEnvelopes.Name = "btnFinEnvelopes"
        Me.btnFinEnvelopes.Size = New System.Drawing.Size(72, 20)
        Me.btnFinEnvelopes.TabIndex = 1067
        Me.btnFinEnvelopes.Text = "Envelopes"
        '
        'lblVendorNumber
        '
        Me.lblVendorNumber.Location = New System.Drawing.Point(8, 48)
        Me.lblVendorNumber.Name = "lblVendorNumber"
        Me.lblVendorNumber.Size = New System.Drawing.Size(100, 16)
        Me.lblVendorNumber.TabIndex = 26
        Me.lblVendorNumber.Text = "Vendor Number:"
        '
        'txtVendorNumberValue
        '
        Me.txtVendorNumberValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtVendorNumberValue.Location = New System.Drawing.Point(112, 48)
        Me.txtVendorNumberValue.Name = "txtVendorNumberValue"
        Me.txtVendorNumberValue.ReadOnly = True
        Me.txtVendorNumberValue.TabIndex = 25
        Me.txtVendorNumberValue.Text = ""
        '
        'txtVendor
        '
        Me.txtVendor.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtVendor.Location = New System.Drawing.Point(400, 48)
        Me.txtVendor.Name = "txtVendor"
        Me.txtVendor.ReadOnly = True
        Me.txtVendor.Size = New System.Drawing.Size(256, 21)
        Me.txtVendor.TabIndex = 24
        Me.txtVendor.Text = ""
        '
        'btnVendorPack
        '
        Me.btnVendorPack.Location = New System.Drawing.Point(680, 80)
        Me.btnVendorPack.Name = "btnVendorPack"
        Me.btnVendorPack.Size = New System.Drawing.Size(144, 23)
        Me.btnVendorPack.TabIndex = 23
        Me.btnVendorPack.Text = "Generate Vendor Pack"
        '
        'txtVendorAddress
        '
        Me.txtVendorAddress.Location = New System.Drawing.Point(112, 80)
        Me.txtVendorAddress.Multiline = True
        Me.txtVendorAddress.Name = "txtVendorAddress"
        Me.txtVendorAddress.ReadOnly = True
        Me.txtVendorAddress.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtVendorAddress.Size = New System.Drawing.Size(200, 60)
        Me.txtVendorAddress.TabIndex = 20
        Me.txtVendorAddress.Text = ""
        '
        'lblVendorAddress
        '
        Me.lblVendorAddress.Location = New System.Drawing.Point(57, 78)
        Me.lblVendorAddress.Name = "lblVendorAddress"
        Me.lblVendorAddress.Size = New System.Drawing.Size(54, 16)
        Me.lblVendorAddress.TabIndex = 19
        Me.lblVendorAddress.Text = "Address:"
        '
        'dtFinancialStart
        '
        Me.dtFinancialStart.Checked = False
        Me.dtFinancialStart.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtFinancialStart.Location = New System.Drawing.Point(736, 48)
        Me.dtFinancialStart.Name = "dtFinancialStart"
        Me.dtFinancialStart.Size = New System.Drawing.Size(88, 21)
        Me.dtFinancialStart.TabIndex = 4
        '
        'lblFinancialStartDate
        '
        Me.lblFinancialStartDate.Location = New System.Drawing.Point(664, 48)
        Me.lblFinancialStartDate.Name = "lblFinancialStartDate"
        Me.lblFinancialStartDate.Size = New System.Drawing.Size(64, 32)
        Me.lblFinancialStartDate.TabIndex = 17
        Me.lblFinancialStartDate.Text = "Financial Start Date:"
        Me.lblFinancialStartDate.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtVendorNo
        '
        Me.txtVendorNo.Location = New System.Drawing.Point(880, 72)
        Me.txtVendorNo.Name = "txtVendorNo"
        Me.txtVendorNo.Size = New System.Drawing.Size(88, 21)
        Me.txtVendorNo.TabIndex = 16
        Me.txtVendorNo.Text = ""
        Me.txtVendorNo.Visible = False
        '
        'lblVendorNo
        '
        Me.lblVendorNo.Location = New System.Drawing.Point(888, 112)
        Me.lblVendorNo.Name = "lblVendorNo"
        Me.lblVendorNo.Size = New System.Drawing.Size(16, 23)
        Me.lblVendorNo.TabIndex = 15
        Me.lblVendorNo.Text = "#:"
        Me.lblVendorNo.Visible = False
        '
        'lblVendor
        '
        Me.lblVendor.Location = New System.Drawing.Point(344, 48)
        Me.lblVendor.Name = "lblVendor"
        Me.lblVendor.Size = New System.Drawing.Size(48, 17)
        Me.lblVendor.TabIndex = 13
        Me.lblVendor.Text = "Vendor:"
        '
        'lblEngineeringFirmValue
        '
        Me.lblEngineeringFirmValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblEngineeringFirmValue.Location = New System.Drawing.Point(400, 80)
        Me.lblEngineeringFirmValue.Name = "lblEngineeringFirmValue"
        Me.lblEngineeringFirmValue.Size = New System.Drawing.Size(256, 56)
        Me.lblEngineeringFirmValue.TabIndex = 12
        '
        'lblEngineeringFirm
        '
        Me.lblEngineeringFirm.Location = New System.Drawing.Point(320, 80)
        Me.lblEngineeringFirm.Name = "lblEngineeringFirm"
        Me.lblEngineeringFirm.Size = New System.Drawing.Size(71, 32)
        Me.lblEngineeringFirm.TabIndex = 11
        Me.lblEngineeringFirm.Text = "Engineering Firm:"
        '
        'lblTechStatusValue
        '
        Me.lblTechStatusValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTechStatusValue.Location = New System.Drawing.Point(736, 16)
        Me.lblTechStatusValue.Name = "lblTechStatusValue"
        Me.lblTechStatusValue.Size = New System.Drawing.Size(88, 23)
        Me.lblTechStatusValue.TabIndex = 10
        '
        'lblTechStatus
        '
        Me.lblTechStatus.Location = New System.Drawing.Point(667, 8)
        Me.lblTechStatus.Name = "lblTechStatus"
        Me.lblTechStatus.Size = New System.Drawing.Size(63, 32)
        Me.lblTechStatus.TabIndex = 9
        Me.lblTechStatus.Text = "Technical Status:"
        Me.lblTechStatus.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblMGPTFStatusValue
        '
        Me.lblMGPTFStatusValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblMGPTFStatusValue.Location = New System.Drawing.Point(555, 16)
        Me.lblMGPTFStatusValue.Name = "lblMGPTFStatusValue"
        Me.lblMGPTFStatusValue.TabIndex = 8
        '
        'lblMGPTFStatus
        '
        Me.lblMGPTFStatus.Location = New System.Drawing.Point(496, 8)
        Me.lblMGPTFStatus.Name = "lblMGPTFStatus"
        Me.lblMGPTFStatus.Size = New System.Drawing.Size(56, 32)
        Me.lblMGPTFStatus.TabIndex = 7
        Me.lblMGPTFStatus.Text = "MGPTF Status:"
        Me.lblMGPTFStatus.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTechStartDateValue
        '
        Me.lblTechStartDateValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTechStartDateValue.Location = New System.Drawing.Point(248, 16)
        Me.lblTechStartDateValue.Name = "lblTechStartDateValue"
        Me.lblTechStartDateValue.Size = New System.Drawing.Size(88, 23)
        Me.lblTechStartDateValue.TabIndex = 6
        '
        'lblPMValue
        '
        Me.lblPMValue.BackColor = System.Drawing.SystemColors.Control
        Me.lblPMValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPMValue.Location = New System.Drawing.Point(384, 16)
        Me.lblPMValue.Name = "lblPMValue"
        Me.lblPMValue.Size = New System.Drawing.Size(104, 23)
        Me.lblPMValue.TabIndex = 5
        '
        'lblPM
        '
        Me.lblPM.Location = New System.Drawing.Point(352, 16)
        Me.lblPM.Name = "lblPM"
        Me.lblPM.Size = New System.Drawing.Size(32, 17)
        Me.lblPM.TabIndex = 4
        Me.lblPM.Text = "PM:"
        '
        'lblTechStartDate
        '
        Me.lblTechStartDate.Location = New System.Drawing.Point(176, 8)
        Me.lblTechStartDate.Name = "lblTechStartDate"
        Me.lblTechStartDate.Size = New System.Drawing.Size(66, 32)
        Me.lblTechStartDate.TabIndex = 2
        Me.lblTechStartDate.Text = "Technical Start Date:"
        Me.lblTechStartDate.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTechEventTitle
        '
        Me.lblTechEventTitle.Location = New System.Drawing.Point(11, 16)
        Me.lblTechEventTitle.Name = "lblTechEventTitle"
        Me.lblTechEventTitle.Size = New System.Drawing.Size(100, 17)
        Me.lblTechEventTitle.TabIndex = 0
        Me.lblTechEventTitle.Text = "Technical Event:"
        '
        'cmbTechEvent
        '
        Me.cmbTechEvent.Location = New System.Drawing.Point(112, 16)
        Me.cmbTechEvent.Name = "cmbTechEvent"
        Me.cmbTechEvent.Size = New System.Drawing.Size(64, 23)
        Me.cmbTechEvent.TabIndex = 2
        '
        'lblTechEvent
        '
        Me.lblTechEvent.BackColor = System.Drawing.SystemColors.Window
        Me.lblTechEvent.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTechEvent.Location = New System.Drawing.Point(112, 16)
        Me.lblTechEvent.Name = "lblTechEvent"
        Me.lblTechEvent.Size = New System.Drawing.Size(64, 23)
        Me.lblTechEvent.TabIndex = 22
        '
        'PnlEvtInfo
        '
        Me.PnlEvtInfo.Controls.Add(Me.lblEvtInfoHead)
        Me.PnlEvtInfo.Controls.Add(Me.lblEvtInfoDisplay)
        Me.PnlEvtInfo.Dock = System.Windows.Forms.DockStyle.Top
        Me.PnlEvtInfo.Location = New System.Drawing.Point(0, 0)
        Me.PnlEvtInfo.Name = "PnlEvtInfo"
        Me.PnlEvtInfo.Size = New System.Drawing.Size(980, 24)
        Me.PnlEvtInfo.TabIndex = 1
        '
        'lblEvtInfoHead
        '
        Me.lblEvtInfoHead.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblEvtInfoHead.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblEvtInfoHead.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblEvtInfoHead.Location = New System.Drawing.Point(16, 0)
        Me.lblEvtInfoHead.Name = "lblEvtInfoHead"
        Me.lblEvtInfoHead.Size = New System.Drawing.Size(964, 24)
        Me.lblEvtInfoHead.TabIndex = 2
        Me.lblEvtInfoHead.Text = "Event Info"
        Me.lblEvtInfoHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblEvtInfoDisplay
        '
        Me.lblEvtInfoDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblEvtInfoDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblEvtInfoDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblEvtInfoDisplay.Name = "lblEvtInfoDisplay"
        Me.lblEvtInfoDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblEvtInfoDisplay.TabIndex = 0
        Me.lblEvtInfoDisplay.Text = "-"
        Me.lblEvtInfoDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlFinEvtsBottom
        '
        Me.pnlFinEvtsBottom.Controls.Add(Me.btnPlanning)
        Me.pnlFinEvtsBottom.Controls.Add(Me.btnFinancialComments)
        Me.pnlFinEvtsBottom.Controls.Add(Me.btnFinancialFlags)
        Me.pnlFinEvtsBottom.Controls.Add(Me.btnCancelEvent)
        Me.pnlFinEvtsBottom.Controls.Add(Me.btnDeleteEvent)
        Me.pnlFinEvtsBottom.Controls.Add(Me.btnSaveEvent)
        Me.pnlFinEvtsBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlFinEvtsBottom.Location = New System.Drawing.Point(0, 576)
        Me.pnlFinEvtsBottom.Name = "pnlFinEvtsBottom"
        Me.pnlFinEvtsBottom.Size = New System.Drawing.Size(1000, 40)
        Me.pnlFinEvtsBottom.TabIndex = 32
        '
        'btnPlanning
        '
        Me.btnPlanning.Enabled = False
        Me.btnPlanning.Location = New System.Drawing.Point(848, 8)
        Me.btnPlanning.Name = "btnPlanning"
        Me.btnPlanning.Size = New System.Drawing.Size(144, 23)
        Me.btnPlanning.TabIndex = 38
        Me.btnPlanning.Text = "Activity Planning"
        '
        'btnFinancialComments
        '
        Me.btnFinancialComments.Location = New System.Drawing.Point(672, 8)
        Me.btnFinancialComments.Name = "btnFinancialComments"
        Me.btnFinancialComments.Size = New System.Drawing.Size(160, 23)
        Me.btnFinancialComments.TabIndex = 37
        Me.btnFinancialComments.Text = "Comments"
        '
        'btnFinancialFlags
        '
        Me.btnFinancialFlags.Location = New System.Drawing.Point(504, 8)
        Me.btnFinancialFlags.Name = "btnFinancialFlags"
        Me.btnFinancialFlags.Size = New System.Drawing.Size(160, 23)
        Me.btnFinancialFlags.TabIndex = 36
        Me.btnFinancialFlags.Text = "Flags"
        '
        'btnCancelEvent
        '
        Me.btnCancelEvent.Location = New System.Drawing.Point(176, 8)
        Me.btnCancelEvent.Name = "btnCancelEvent"
        Me.btnCancelEvent.Size = New System.Drawing.Size(160, 23)
        Me.btnCancelEvent.TabIndex = 34
        Me.btnCancelEvent.Text = "Cancel Changes"
        '
        'btnDeleteEvent
        '
        Me.btnDeleteEvent.Enabled = False
        Me.btnDeleteEvent.Location = New System.Drawing.Point(344, 8)
        Me.btnDeleteEvent.Name = "btnDeleteEvent"
        Me.btnDeleteEvent.Size = New System.Drawing.Size(160, 23)
        Me.btnDeleteEvent.TabIndex = 35
        Me.btnDeleteEvent.Text = "Delete Event"
        '
        'btnSaveEvent
        '
        Me.btnSaveEvent.Location = New System.Drawing.Point(8, 8)
        Me.btnSaveEvent.Name = "btnSaveEvent"
        Me.btnSaveEvent.Size = New System.Drawing.Size(160, 23)
        Me.btnSaveEvent.TabIndex = 33
        Me.btnSaveEvent.Text = "Save Event"
        '
        'pnlFinancialHeader
        '
        Me.pnlFinancialHeader.Controls.Add(Me.lblFinancialClosedDate)
        Me.pnlFinancialHeader.Controls.Add(Me.dtFinancialClosedDate)
        Me.pnlFinancialHeader.Controls.Add(Me.cmbFinancialStatus)
        Me.pnlFinancialHeader.Controls.Add(Me.lblFinancialStatus)
        Me.pnlFinancialHeader.Controls.Add(Me.lblFinancialID)
        Me.pnlFinancialHeader.Controls.Add(Me.lblFinancialCountVal)
        Me.pnlFinancialHeader.Controls.Add(Me.lblFinancialIDValue)
        Me.pnlFinancialHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlFinancialHeader.Location = New System.Drawing.Point(0, 0)
        Me.pnlFinancialHeader.Name = "pnlFinancialHeader"
        Me.pnlFinancialHeader.Size = New System.Drawing.Size(1000, 32)
        Me.pnlFinancialHeader.TabIndex = 0
        '
        'lblFinancialClosedDate
        '
        Me.lblFinancialClosedDate.Location = New System.Drawing.Point(450, 8)
        Me.lblFinancialClosedDate.Name = "lblFinancialClosedDate"
        Me.lblFinancialClosedDate.Size = New System.Drawing.Size(100, 17)
        Me.lblFinancialClosedDate.TabIndex = 214
        Me.lblFinancialClosedDate.Text = "Closed Date: "
        '
        'dtFinancialClosedDate
        '
        Me.dtFinancialClosedDate.Checked = False
        Me.dtFinancialClosedDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtFinancialClosedDate.Location = New System.Drawing.Point(560, 8)
        Me.dtFinancialClosedDate.Name = "dtFinancialClosedDate"
        Me.dtFinancialClosedDate.ShowCheckBox = True
        Me.dtFinancialClosedDate.Size = New System.Drawing.Size(104, 21)
        Me.dtFinancialClosedDate.TabIndex = 213
        '
        'cmbFinancialStatus
        '
        Me.cmbFinancialStatus.Location = New System.Drawing.Point(263, 5)
        Me.cmbFinancialStatus.Name = "cmbFinancialStatus"
        Me.cmbFinancialStatus.Size = New System.Drawing.Size(121, 23)
        Me.cmbFinancialStatus.TabIndex = 0
        '
        'lblFinancialStatus
        '
        Me.lblFinancialStatus.Location = New System.Drawing.Point(152, 8)
        Me.lblFinancialStatus.Name = "lblFinancialStatus"
        Me.lblFinancialStatus.Size = New System.Drawing.Size(100, 17)
        Me.lblFinancialStatus.TabIndex = 212
        Me.lblFinancialStatus.Text = "Financial Status:"
        '
        'lblFinancialID
        '
        Me.lblFinancialID.AutoSize = True
        Me.lblFinancialID.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.lblFinancialID.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFinancialID.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblFinancialID.Location = New System.Drawing.Point(8, 8)
        Me.lblFinancialID.Name = "lblFinancialID"
        Me.lblFinancialID.Size = New System.Drawing.Size(55, 17)
        Me.lblFinancialID.TabIndex = 209
        Me.lblFinancialID.Text = "Event #: "
        '
        'lblFinancialCountVal
        '
        Me.lblFinancialCountVal.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.lblFinancialCountVal.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFinancialCountVal.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblFinancialCountVal.Location = New System.Drawing.Point(96, 8)
        Me.lblFinancialCountVal.Name = "lblFinancialCountVal"
        Me.lblFinancialCountVal.Size = New System.Drawing.Size(39, 17)
        Me.lblFinancialCountVal.TabIndex = 211
        Me.lblFinancialCountVal.Text = "of ???"
        Me.lblFinancialCountVal.Visible = False
        '
        'lblFinancialIDValue
        '
        Me.lblFinancialIDValue.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.lblFinancialIDValue.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblFinancialIDValue.Location = New System.Drawing.Point(61, 8)
        Me.lblFinancialIDValue.Name = "lblFinancialIDValue"
        Me.lblFinancialIDValue.Size = New System.Drawing.Size(35, 17)
        Me.lblFinancialIDValue.TabIndex = 210
        Me.lblFinancialIDValue.Text = "00"
        Me.lblFinancialIDValue.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'tbPageSummary
        '
        Me.tbPageSummary.AutoScroll = True
        Me.tbPageSummary.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.tbPageSummary.Controls.Add(Me.pnlOwnerSummaryDetails)
        Me.tbPageSummary.Controls.Add(Me.Panel12)
        Me.tbPageSummary.Controls.Add(Me.pnlOwnerSummaryHeader)
        Me.tbPageSummary.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbPageSummary.Location = New System.Drawing.Point(4, 22)
        Me.tbPageSummary.Name = "tbPageSummary"
        Me.tbPageSummary.Size = New System.Drawing.Size(1008, 644)
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
        Me.pnlOwnerSummaryDetails.Size = New System.Drawing.Size(996, 624)
        Me.pnlOwnerSummaryDetails.TabIndex = 7
        '
        'UCOwnerSummary
        '
        Me.UCOwnerSummary.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UCOwnerSummary.Location = New System.Drawing.Point(0, 0)
        Me.UCOwnerSummary.Name = "UCOwnerSummary"
        Me.UCOwnerSummary.Size = New System.Drawing.Size(996, 624)
        Me.UCOwnerSummary.TabIndex = 0
        '
        'Panel12
        '
        Me.Panel12.Dock = System.Windows.Forms.DockStyle.Right
        Me.Panel12.DockPadding.Left = 10
        Me.Panel12.Location = New System.Drawing.Point(996, 16)
        Me.Panel12.Name = "Panel12"
        Me.Panel12.Size = New System.Drawing.Size(8, 624)
        Me.Panel12.TabIndex = 6
        '
        'pnlOwnerSummaryHeader
        '
        Me.pnlOwnerSummaryHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlOwnerSummaryHeader.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlOwnerSummaryHeader.Location = New System.Drawing.Point(0, 0)
        Me.pnlOwnerSummaryHeader.Name = "pnlOwnerSummaryHeader"
        Me.pnlOwnerSummaryHeader.Size = New System.Drawing.Size(1004, 16)
        Me.pnlOwnerSummaryHeader.TabIndex = 2
        '
        'Financial
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1016, 694)
        Me.Controls.Add(Me.tbCntrlFinancial)
        Me.Controls.Add(Me.pnlTop)
        Me.Name = "Financial"
        Me.Text = "Financial"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.pnlTop.ResumeLayout(False)
        Me.tbCntrlFinancial.ResumeLayout(False)
        Me.tbPageOwnerDetail.ResumeLayout(False)
        Me.pnlOwnerBottom.ResumeLayout(False)
        Me.tbCtrlOwner.ResumeLayout(False)
        Me.tbPageOwnerFacilities.ResumeLayout(False)
        CType(Me.ugFacilityList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlOwnerFacilityBottom.ResumeLayout(False)
        Me.tbPageOwnerContactList.ResumeLayout(False)
        Me.pnlOwnerContactContainer.ResumeLayout(False)
        CType(Me.ugOwnerContacts, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlOwnerContactHeader.ResumeLayout(False)
        Me.pnlOwnerContactButtons.ResumeLayout(False)
        Me.tbPageOwnerDocuments.ResumeLayout(False)
        Me.pnlOwnerDetail.ResumeLayout(False)
        Me.pnlOwnerButtons.ResumeLayout(False)
        CType(Me.mskTxtOwnerFax, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtOwnerPhone2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtOwnerPhone, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbPageFacilityDetail.ResumeLayout(False)
        Me.pnlFacilityBottom.ResumeLayout(False)
        Me.tbCtrlFacFinancialEvt.ResumeLayout(False)
        Me.tbPageFacFinancialEvents.ResumeLayout(False)
        CType(Me.ugFinancialGrid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlFacilityFinancialButton.ResumeLayout(False)
        Me.tbPageFacilityDocuments.ResumeLayout(False)
        Me.pnl_FacilityDetail.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        CType(Me.mskTxtFacilityFax, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtFacilityPhone, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbPageFinancialEvent.ResumeLayout(False)
        Me.tbCtrlFinancialEvtDetails.ResumeLayout(False)
        Me.tbPageFinancialEvtDetails.ResumeLayout(False)
        Me.pnlFinEvtsDetails.ResumeLayout(False)
        Me.pnlContactDetails.ResumeLayout(False)
        Me.pnlFinancialContactDetails.ResumeLayout(False)
        CType(Me.ugFinancialContacts, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlFinancialContactButtons.ResumeLayout(False)
        Me.pnlFinancialContactHeader.ResumeLayout(False)
        Me.pnlContacts.ResumeLayout(False)
        Me.pnlPaymentsDetails.ResumeLayout(False)
        Me.pnlPaymentButtons.ResumeLayout(False)
        Me.pnlPaymentTotals.ResumeLayout(False)
        CType(Me.ugPayments, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlPayments.ResumeLayout(False)
        Me.pnlCommitmentsDetails.ResumeLayout(False)
        Me.pnlCommitmentsTotals.ResumeLayout(False)
        Me.PnlCommitmentButtons.ResumeLayout(False)
        CType(Me.ugCommitments, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlCommitments.ResumeLayout(False)
        Me.pnlEvtInfoDetails.ResumeLayout(False)
        Me.PnlEvtInfo.ResumeLayout(False)
        Me.pnlFinEvtsBottom.ResumeLayout(False)
        Me.pnlFinancialHeader.ResumeLayout(False)
        Me.tbPageSummary.ResumeLayout(False)
        Me.pnlOwnerSummaryDetails.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Intialization"

    Private Sub InitControls()
        Try
            UIUtilsGen.PopulateOwnerType(cmbOwnerType, pOwn)
            If pOwn.ID <> 0 Then
                UIUtilsGen.PopulateFacilityType(Me.cmbFacilityType, pOwn.Facilities)
                UIUtilsGen.PopulateFacilityDatum(Me.cmbFacilityDatum, pOwn.Facilities)
                UIUtilsGen.PopulateFacilityMethod(Me.cmbFacilityMethod, pOwn.Facilities)
                UIUtilsGen.PopulateFacilityLocationType(Me.cmbFacilityLocationType, pOwn.Facilities)
            End If
            'btnFacilitySave.Enabled = False
            'btnFacilityCancel.Enabled = False
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Property FormLoading() As Boolean
        Get
            Return bolLoading
        End Get
        Set(ByVal Value As Boolean)
            bolLoading = Value
        End Set
    End Property

    Private Sub Financial_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.TextChanged
        Try
            MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "WindowName", Me.Text, "Financial")
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub Financial_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        'Dim MyFrm As MusterContainer
        Try
            MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "Financial")
            'MyFrm = Me.MdiParent
            'If lblOwnerIDValue.Text <> String.Empty And lblFacilityIDValue.Text = String.Empty Then
            '    MyFrm.FlagsChanged(Me.lblOwnerIDValue.Text, 0, 0, 0, "Financial", Me.Text)
            'ElseIf lblOwnerIDValue.Text <> String.Empty And lblFacilityIDValue.Text <> String.Empty Then
            '    MyFrm.FlagsChanged(Me.lblOwnerIDValue.Text, Me.lblFacilityIDValue.Text, 0, 0, "Financial", Me.Text)
            'End If
            If lblOwnerIDValue.Text <> String.Empty Then ' And lblFacilityIDValue.Text = String.Empty Then
                pOwn.Retrieve(Me.lblOwnerIDValue.Text, "SELF")
            End If
            If lblFinancialIDValue.Text <> "00" And lblFinancialIDValue.Text <> "New" Then
                LoadCommitmentsGrid()
                LoadPaymentsGrid()
            End If
            bolFrmActivated = True
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

#End Region

#Region "Event Handlers for Expanding/Collapsing the different Sections of the Lust Event form"
    Private Sub ShowHideControl(ByVal ObjControl As Control)
        Try
            If ObjControl.Visible Then
                ObjControl.Visible = False
            Else
                ObjControl.Visible = True
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub lblEvtInfoDisplay_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblEvtInfoDisplay.Click, lblEvtInfoHead.Click
        Try
            If lblEvtInfoDisplay.Text = "+" Then
                lblEvtInfoDisplay.Text = "-"
            Else
                lblEvtInfoDisplay.Text = "+"
            End If
            ShowHideControl(Me.pnlEvtInfoDetails)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub lblCommitmentsDisplay_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblCommitmentsDisplay.Click, lblCommitmentsHead.Click
        Try
            If lblCommitmentsDisplay.Text = "+" Then
                lblCommitmentsDisplay.Text = "-"
            Else
                lblCommitmentsDisplay.Text = "+"
            End If
            ShowHideControl(Me.pnlCommitmentsDetails)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub lblPaymentsDisplay_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblPaymentsDisplay.Click, lblPaymentsHead.Click
        Try
            If lblPaymentsDisplay.Text = "+" Then
                lblPaymentsDisplay.Text = "-"
            Else
                lblPaymentsDisplay.Text = "+"
            End If
            ShowHideControl(Me.pnlPaymentsDetails)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub lblContactsDisplay_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblContactsDisplay.Click, lblContactsHead.Click
        Try
            If lblContactsDisplay.Text = "+" Then
                lblContactsDisplay.Text = "-"
            Else
                lblContactsDisplay.Text = "+"
            End If
            ShowHideControl(Me.pnlContactDetails)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "Tab Operations"

    Private Sub tbCntrlFinancial_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbCntrlFinancial.Click
        Dim MyFrm As MusterContainer
        Try
            Select Case tbCntrlFinancial.SelectedTab.Name.ToUpper
                Case "TBPAGEFACILITYDETAIL"
                    If Me.ugFacilityList.Rows.Count <> 0 Then
                        If Me.lblFacilityIDValue.Text = String.Empty Then
                            PopulateFacilityInfo(Integer.Parse(ugFacilityList.ActiveRow.Cells("FACILITYID").Text))
                        Else
                            PopulateFacilityInfo(Integer.Parse(Me.lblFacilityIDValue.Text))
                        End If
                        Me.Tag = Me.lblFacilityIDValue.Text
                        Me.lblFacilityIDValue.Focus()
                        MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "FacilityDetails", Me.lblFacilityIDValue.Text, "Financial")
                    End If
                    If ugFacilityList.Rows.Count <= 0 And Me.lblOwnerIDValue.Text <> String.Empty Then
                        Dim msgResult As MsgBoxResult
                        msgResult = MsgBox("No facilities found for owner" + lblOwnerIDValue.Text)
                        Exit Select
                    End If

                    Me.Text = "Financial - Facility Detail - (" & IIf(txtFacilityName.Text = String.Empty, "New", txtFacilityName.Text) & ")"
                    nCurrentEventID = -1
                Case "TBPAGEOWNERDETAIL"
                    Me.Text = "Financial - Owner Detail (" & txtOwnerName.Text & ")"
                    If lblOwnerIDValue.Text <> String.Empty Then

                        UIUtilsGen.PopulateOwnerFacilities(pOwn, Me, Integer.Parse(Me.lblOwnerIDValue.Text))



                    End If

                    If pOwn.ID > 0 And Not tbCtrlOwner.TabPages.Contains(tbPageOwnerContactList) Then
                        tbCtrlOwner.TabPages.Add(tbPageOwnerContactList)
                        lblOwnerContacts.Text = "Owner Contacts"
                    End If
                    'LoadContacts(ugOwnerContacts, pOwn.ID, 9)
                    MyFrm = Me.MdiParent
                    MyFrm.FlagsChanged(Me.lblOwnerIDValue.Text, UIUtilsGen.EntityTypes.Owner, "Financial", Me.Text)
                    tbCtrlOwner.SelectedTab = tbPageOwnerFacilities
                    nCurrentEventID = -1
                Case "TBCTRLFINANCIALEVENT"
                    'zxc - Add Financial Event # Here
                    Me.Text = "Financial Event - "
                    tbPageFinancialEvent.Enabled = True

                    If ugFinancialGrid.Rows.Count <> 0 Then
                        If (nLastEventID <> ugFinancialGrid.ActiveRow.Cells("FIN_EVENT_ID").Text) And nLastEventID >= 0 Then
                            LoadFinancialData(ugFinancialGrid.ActiveRow.Cells("FIN_EVENT_ID").Text)
                        End If
                    Else
                        tbPageFinancialEvent.Enabled = False
                        nCurrentEventID = -1
                    End If
                Case "TBPAGEFINANCIALEVENT"
                    If nLastEventID > 0 Then
                        LoadFinancialData(nLastEventID)
                    Else
                        If cmbTechEvent.Items.Count <= 0 Then
                            tbPageFinancialEvent.Enabled = False
                            nCurrentEventID = -1
                        End If
                    End If
                Case "TBPAGESUMMARY"
                    Me.Text = "Financial - Owner Summary (" & txtOwnerName.Text & ")"
                    UIUtilsGen.PopulateOwnerSummary(pOwn, Me)
                    nCurrentEventID = -1
            End Select

            Me.CausesValidation = True
            Me.CausesValidation = False

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub tbCntrlFinancial_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbCntrlFinancial.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Dim nFacilityID As Integer
        Try
            Select Case tbCntrlFinancial.SelectedTab.Name.ToUpper
                Case "TBPAGEOWNERDETAIL"
                    Me.Text = "Financial - Owner Detail (" & txtOwnerName.Text & ")"
                    Me.PopulateOwnerInfo(pOwn.ID)
                    tbCtrlOwner.SelectedTab = tbPageOwnerFacilities
                    nCurrentEventID = -1
                Case "TBPAGEFACILITYDETAIL"
                    Me.Text = "Financial - Facility Detail - (" & IIf(txtFacilityName.Text = String.Empty, "New", txtFacilityName.Text) & ")"
                    If ugFacilityList.Rows.Count > 0 And Not pOwn.Facilities.ID > 0 Then
                        If ugFacilityList.ActiveRow Is Nothing Then
                            ugFacilityList.ActiveRow = ugFacilityList.Rows(0)
                        End If
                        nFacilityID = ugFacilityList.ActiveRow.Cells("FacilityID").Value
                        Me.PopulateFacilityInfo(nFacilityID)
                    Else
                        Me.PopulateFacilityInfo(pOwn.Facilities.ID)
                    End If
                    nCurrentEventID = -1
                Case "TBPAGEFINANCIALEVENT"
                    'zxc - Add Financial Event # Here
                    Me.Text = "Financial Event - "
                    tbPageFinancialEvent.Enabled = True

                    If ugFinancialGrid.Rows.Count <> 0 Then
                        If Not ugFinancialGrid.ActiveRow Is Nothing Then
                            If (nLastEventID <> ugFinancialGrid.ActiveRow.Cells("FIN_EVENT_ID").Text) And nLastEventID >= 0 Then
                                LoadFinancialData(ugFinancialGrid.ActiveRow.Cells("FIN_EVENT_ID").Text)
                            End If
                        End If
                    End If

            End Select

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

#End Region

#Region "Form Events"


    Private Sub ugFinancialContact_Changed(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugFinancialContacts.CellChange

        Me.oFinancial.IsDirty = True

        btnSaveEvent.Enabled = True

    End Sub


    Private Sub Financial_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim cntrl As Control
        Dim uiUts As New UIUtilsGen
        Try

            For Each cntrl In Me.Controls
                uiUts.ClearComboBox(cntrl)
                'uiUts.RetainCurrentDateValue(cntrl)
            Next
            bolLoading = False
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)
        Try

            'If pOwn.colIsDirty() Then
            '    Dim Results As Long = MsgBox("There are unsaved changes.  Do you wish to save the changes?", MsgBoxStyle.Question & MsgBoxStyle.YesNoCancel, "Data Changed")
            '    If Results = MsgBoxResult.Yes Then

            '        Dim success As Boolean = False
            '        pOwn.ModifiedBy = MusterContainer.AppUser.ID
            '        success = pOwn.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
            '        If Not UIUtilsGen.HasRights(returnVal) Then
            '            e.Cancel = True
            '            Exit Sub
            '        End If

            '        If Not success Then
            '            e.Cancel = True
            '            bolValidateSuccess = True
            '            bolDisplayErrmessage = True
            '            Exit Sub
            '        End If
            '    ElseIf Results = MsgBoxResult.Cancel Then
            '        e.Cancel = True
            '        Exit Sub
            '    End If
            'End If
            'if any other forms are using the owner, leave alone. else remove from collection
            UIUtilsGen.RemoveOwner(pOwn, Me)

            If oFinancial.IsDirty And btnSaveEvent.Enabled = True And btnSaveEvent.Visible = True Then
                If _container.DirtyIgnored = MsgBoxResult.Yes OrElse (_container.DirtyIgnored = -1 AndAlso MsgBox("Do you wish to save changes?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes) Then
                    If oFinancial.ID <= 0 Then
                        oFinancial.CreatedBy = MusterContainer.AppUser.ID
                    Else
                        oFinancial.ModifiedBy = MusterContainer.AppUser.ID
                    End If
                    oFinancial.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If
                End If
            End If


        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Owner Operations"

#Region "UI Support Routines"
    Friend Sub PopulateOwnerInfo(ByVal OwnerID As Integer)
        Dim MyFrm As MusterContainer
        Try
            UIUtilsGen.PopulateOwnerInfo(OwnerID, pOwn, Me)
            If Not tbCtrlOwner.TabPages.Contains(tbPageOwnerContactList) Then
                tbCtrlOwner.TabPages.Add(tbPageOwnerContactList)
                lblOwnerContacts.Text = "Owner Contacts"
            End If
            Select Case tbCtrlOwner.SelectedTab.Name
                Case tbPageOwnerContactList.Name
                    LoadContacts(ugOwnerContacts, OwnerID, UIUtilsGen.EntityTypes.Owner)
                Case tbPageOwnerDocuments.Name
                    UCOwnerDocuments.LoadDocumentsGrid(OwnerID, UIUtilsGen.EntityTypes.Owner, UIUtilsGen.ModuleID.Financial)
            End Select
            'LoadContacts(ugOwnerContacts, OwnerID, 9)
            If bolFrmActivated Then
                MyFrm = Me.MdiParent
                'oEntity.GetEntity("Owner")
                If lblOwnerIDValue.Text <> String.Empty Then
                    MyFrm.FlagsChanged(Me.lblOwnerIDValue.Text, UIUtilsGen.EntityTypes.Owner, "Financial", Me.Text)
                End If
            End If
            CommentsMaintenance(, , True)
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
#End Region

#Region "UI Control Events"
    Private Sub ugFacilityList_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugFacilityList.DoubleClick
        If Me.Cursor Is Cursors.WaitCursor Then
            Exit Sub
        End If
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            Me.Cursor = Cursors.WaitCursor
            tbCntrlFinancial.SelectedTab = Me.tbPageFacilityDetail
            PopulateFacilityInfo(Integer.Parse(ugFacilityList.ActiveRow.Cells.Item("FacilityID").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw)))
            Me.Tag = Me.lblFacilityIDValue.Text

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub
    Private Sub ugFacilityList_AfterSortChange(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.BandEventArgs) Handles ugFacilityList.AfterSortChange
        Try
            If ugFacilityList.Rows.Count > 0 Then
                ugFacilityList.ActiveRow = ugFacilityList.Rows(0)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnOwnerComment_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOwnerComment.Click
        Try
            CommentsMaintenance(sender, e)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#End Region

#Region "Facility Operations"

#Region "UI Support Routines"
    Friend Sub PopulateFacilityInfo(Optional ByVal FacilityID As Integer = 0)
        Dim MyFrm As MusterContainer
        Try
            UIUtilsGen.PopulateFacilityInfo(Me, pOwn.OwnerInfo, pOwn.Facilities, FacilityID)
            nFacilityID = FacilityID
            Me.Text = "Financial Events - Facility #" & lblFacilityIDValue.Text & " (" & txtFacilityName.Text & ")"
            If FacilityID > 0 And Not tbCtrlFacFinancialEvt.TabPages.Contains(tbPageOwnerContactList) Then
                tbCtrlFacFinancialEvt.TabPages.Add(tbPageOwnerContactList)
                lblOwnerContacts.Text = "Facility Contacts"
            End If
            Select Case tbCtrlFacFinancialEvt.SelectedTab.Name
                Case tbPageOwnerContactList.Name
                    If Me.lblFacilityIDValue.Text <> String.Empty Then
                        LoadContacts(ugOwnerContacts, FacilityID, UIUtilsGen.EntityTypes.Facility)
                    End If
                Case tbPageFacilityDocuments.Name
                    UCFacilityDocuments.LoadDocumentsGrid(FacilityID, UIUtilsGen.EntityTypes.Facility, UIUtilsGen.ModuleID.Financial)
            End Select
            'LoadContacts(ugOwnerContacts, FacilityID, 6)
            If bolFrmActivated Then
                MyFrm = Me.MdiParent
                'oEntity.GetEntity("Facility")
                MyFrm.FlagsChanged(Me.lblFacilityIDValue.Text, UIUtilsGen.EntityTypes.Facility, "Financial", Me.Text)
            End If

            tbCtrlFacFinancialEvt.SelectedTab = tbPageFacFinancialEvents
            CommentsMaintenance(, , True)

            ' 2186
            Dim nTotalOpenEvents As Integer = 0
            Dim nOpenEventID As Integer = 0
            For Each ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugFinancialGrid.Rows
                If ugrow.Cells("Financial_event_status").Text.IndexOf("active") > -1 Then
                    nTotalOpenEvents += 1
                    nOpenEventID = ugrow.Cells("FIN_EVENT_ID").Value
                End If
            Next
            If nTotalOpenEvents = 1 Then
                LoadFinancialData(nOpenEventID)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub GetFinancialEventsForFacility()
        Dim drRow As DataRow
        Dim rowcount As Integer = 0
        Dim str As String = String.Empty
        Try
            strFinancialEventIdTags = String.Empty
            ugFinancialGrid.DataSource = pOwn.Facilities.FinancialEventDataset
            For Each drRow In pOwn.Facilities.FinancialEventDataset.Tables(0).Rows
                If rowcount < pOwn.Facilities.FinancialEventDataset.Tables(0).Rows.Count - 1 Then
                    str = ","
                Else
                    str = ""
                End If
                strFinancialEventIdTags += drRow("FIN_EVENT_ID").ToString + str
                rowcount += 1
            Next
            ugFinancialGrid.Rows.ExpandAll(True)
            ugFinancialGrid.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
            ugFinancialGrid.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
            ugFinancialGrid.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
            If ugFinancialGrid.Rows.Count > 0 Then
                ugFinancialGrid.DisplayLayout.Bands(0).Columns("FIN_EVENT_ID").Hidden = True
                ugFinancialGrid.DisplayLayout.Bands(0).Columns("Event_Date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                ugFinancialGrid.DisplayLayout.Bands(0).Columns("CommitmentTotal").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                ugFinancialGrid.DisplayLayout.Bands(0).Columns("RequestedTotal").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                ugFinancialGrid.DisplayLayout.Bands(0).Columns("TotalPaid").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                'ugFinancialGrid.DisplayLayout.Bands(0).Columns("EVENT_ID").Hidden = True
                'ugFinancialGrid.DisplayLayout.Bands(0).Columns("Priority").Width = 50
            End If

        Catch ex As Exception
            Throw ex
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
#End Region

#Region "UI Control Events"

    Private Sub lnkLblNextFac_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkLblNextFac.LinkClicked
        Try
            If IsNumeric(lblFacilityIDValue.Text) Then
                PopulateFacilityInfo(GetPrevNextFacility(lblFacilityIDValue.Text, True))
                MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "FacilityDetails", Me.lblFacilityIDValue.Text, "Financial")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub lnkLblPrevFacility_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkLblPrevFacility.LinkClicked
        Try
            If IsNumeric(lblFacilityIDValue.Text) Then
                PopulateFacilityInfo(GetPrevNextFacility(lblFacilityIDValue.Text, False))
                MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "FacilityDetails", Me.lblFacilityIDValue.Text, "Financial")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnFacComments_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFacComments.Click
        Try
            CommentsMaintenance(sender, e)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

#End Region

#End Region

#Region "Financial Event Operations"

#Region "UI Control Routines"

    Private Sub cmbFinancialStatus_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbFinancialStatus.SelectedIndexChanged
        If bolLoading = True Then Exit Sub

        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow

        If cmbFinancialStatus.Text = "Closed" Then
            For Each ugrow In ugCommitments.Rows    '2.a.v above
                If ugrow.Cells("Balance").Value > 0 Then
                    bolLoading = True
                    cmbFinancialStatus.SelectedValue = 1027
                    MsgBox("The Status May Not Be Changed To Closed If Any Commitments Have A Non-Zero Balance.")
                    bolLoading = False
                    Exit Sub
                End If
            Next
        End If
        If oFinancial.Status <> cmbFinancialStatus.SelectedValue Then
            oFinancial.Status = cmbFinancialStatus.SelectedValue
        End If

    End Sub

    Private Sub cmbTechEvent_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbTechEvent.SelectedIndexChanged
        If bolLoading = True Then Exit Sub
        lblTechEvent.Text = Me.cmbTechEvent.Text
        oFinancial.TecEventID = cmbTechEvent.SelectedValue

        LoadTechInfo()
    End Sub

    Private Sub btnAddFinancialEvt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddFinancialEvt.Click

        nLastEventID = -1
        nCurrentEventID = -1
        LoadFinancialData(0)


        bolLoading = True
        PopulateTechEvent()
        lblTechEvent.Text = ""
        If cmbTechEvent.Items.Count > 0 Then
            LoadTechInfo()
            oFinancial.TecEventID = cmbTechEvent.SelectedValue
        End If
        bolLoading = False
        pnlContacts.Visible = False
        pnlContactDetails.Visible = False

        If cmbTechEvent.Items.Count <= 0 Then
            MsgBox("There are no eligible Tech Events for Facility #" & lblFacilityIDValue.Text)
            Me.tbPageFinancialEvent.Enabled = False
        Else
            cmbTechEvent.SelectedIndex = 0
            lblTechEvent.Text = cmbTechEvent.Text
        End If
        Me.Text = "Financial - Add Financial Event (" & pOwn.Facilities.ID & ")"
    End Sub

    Private Sub btnCancelEvent_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancelEvent.Click
        If lblTechEvent.Text = "" Then
            bolLoading = True
            PopulateTechEvent()
            lblTechEvent.Text = ""
            If cmbTechEvent.Items.Count > 0 Then
                LoadTechInfo()
                oFinancial.TecEventID = cmbTechEvent.SelectedValue
            End If
            bolLoading = False
            nLastEventID = -1
            nCurrentEventID = -1
        Else
            oFinancial.Reset()
            LoadFinancialData(oFinancial.ID)
        End If
        'If lblTechEvent.Text = "" Then
        '    bolLoading = True
        '    PopulateTechEvent()
        '    lblTechEvent.Text = ""
        '    If cmbTechEvent.Items.Count > 0 Then
        '        LoadTechInfo()
        '        oFinancial.TecEventID = cmbTechEvent.SelectedValue
        '    End If
        '    bolLoading = False
        '    nLastEventID = -1
        'Else
        '    If ugFinancialGrid.Rows.Count > 0 Then
        '        LoadFinancialData(ugFinancialGrid.ActiveRow.Cells("FIN_EVENT_ID").Text)
        '    Else
        '        'tbCntrlFinancial.SelectedTab = Me.tbPageFinancialEvent
        '        tbCntrlFinancial.SelectedTab = tbPageFacilityDetail
        '    End If
        'End If
    End Sub


    Public Overrides Sub PerformClick()

        If Not ugFinancialGrid.ActiveRow Is Nothing Then
            LoadFinancialData(ugFinancialGrid.ActiveRow.Cells("FIN_EVENT_ID").Text)
        End If
    End Sub

    Public Sub ugFinancialGrid_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugFinancialGrid.DoubleClick
        MyBase.DoubleClickEventOnGrid(sender, e)
    End Sub

    Private Sub btnSaveEvent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveEvent.Click
        Try

            If oFinancial.ID <= 0 Then
                oFinancial.CreatedBy = MusterContainer.AppUser.ID
            Else
                oFinancial.ModifiedBy = MusterContainer.AppUser.ID
            End If
            oFinancial.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            LoadFinancialData(oFinancial.ID)
            MsgBox("Financial Event Successfully Saved")
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Save Financial Event" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub dtFinancialStart_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtFinancialStart.ValueChanged
        If bolLoading = True Then Exit Sub
        oFinancial.StartDate = dtFinancialStart.Value.Date
    End Sub

#End Region

#Region "UI Support Routines"
    Private Sub frmClosing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        If sender.GetType.Name.IndexOf("IncompleteApplication") >= 0 Then
        ElseIf sender.GetType.Name.IndexOf("Commitment") >= 0 Then
        ElseIf sender.GetType.Name.IndexOf("Adjustment") >= 0 Then
        ElseIf sender.GetType.Name.IndexOf("Invoice") >= 0 Then
        End If
    End Sub

    Private Sub frmClosed(ByVal sender As Object, ByVal e As System.EventArgs)
        If sender.GetType.Name.IndexOf("IncompleteApplication") >= 0 Then
            frmIncompApplication = Nothing
        ElseIf sender.GetType.Name.IndexOf("Commitment") >= 0 Then
            frmCommitment = Nothing
        ElseIf sender.GetType.Name.IndexOf("Adjustment") >= 0 Then
            frmAdjustment = Nothing
        ElseIf sender.GetType.Name.IndexOf("Invoice") >= 0 Then
            frmInvoice = Nothing
        End If
    End Sub

    Private Sub PopulateFinancialStatus()
        Try
            Dim dtLustEventStatus As DataTable = oFinancial.PopulateFinancialStatus
            If Not IsNothing(dtLustEventStatus) Then
                cmbFinancialStatus.DataSource = dtLustEventStatus
                cmbFinancialStatus.DisplayMember = "PROPERTY_NAME"
                cmbFinancialStatus.ValueMember = "PROPERTY_ID"
            Else
                cmbFinancialStatus.DataSource = Nothing
            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Populate Financial Status" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub PopulateTechEvent()
        Try
            Dim dtTechEvent As DataTable = oFinancial.PopulateEligibleTecEvents(lblFacilityIDValue.Text)
            If Not IsNothing(dtTechEvent) Then
                cmbTechEvent.DataSource = dtTechEvent
                cmbTechEvent.DisplayMember = "EVENT_SEQUENCE"
                cmbTechEvent.ValueMember = "EVENT_ID"
            Else
                cmbTechEvent.DataSource = Nothing
            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Populate Leak Priority" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub



#End Region

#Region "Commitments"

#Region "UI Support Routines"
    Private Sub CreateModifyCommitment(ByVal CommitmentID)
        Dim MyFrm As MusterContainer
        Try
            If IsNothing(frmCommitment) Then
                frmCommitment = New Commitment
                frmCommitment.FinancialEventID = oFinancial.ID
                frmCommitment.FinancialCommitmentID = CommitmentID
                frmCommitment.ugCommitRow = ugCommitments.ActiveRow
                nCommitmentID = CommitmentID
                frmCommitment.CallingForm = Me

                AddHandler frmCommitment.Closing, AddressOf frmClosing
                AddHandler frmCommitment.Closed, AddressOf frmClosed
            End If
            frmCommitment.ShowDialog()

            If nCommitmentID <= 0 Then
                If Not Me.Tag Is Nothing Then
                    If Me.Tag.ToString.StartsWith("C") Then
                        If IsNumeric(Me.Tag.ToString.TrimStart("C")) Then
                            nCommitmentID = Me.Tag.ToString.TrimStart("C")
                        End If
                    End If
                End If
            End If

            LoadCommitmentsGrid()

            MyFrm = MdiParent
            MyFrm.RefreshCalendarInfo()
            MyFrm.LoadDueToMeCalendar()
            MyFrm.LoadToDoCalendar()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "UI Control Events"

    Private Sub btnAddCommitment_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddCommitment.Click
        Try
            CreateModifyCommitment(0)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnModViewCommitment_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnModViewCommitment.Click
        Try
            If Not ugCommitments.ActiveRow Is Nothing Then
                If ugCommitments.ActiveRow.Band.Index = 0 Then
                    CreateModifyCommitment(ugCommitments.ActiveRow.Cells("CommitmentID").Value)
                Else
                    MsgBox("Please Select A Commitment To Modify")
                End If
            Else
                MsgBox("Please Select A Commitment To Modify")
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#End Region

#Region "Adjustment"
#Region "UI Support Routines"
    Private Sub CreateModifyAdjustment(ByVal CommitmentID As Int64, ByVal AdjustmentID As Int64)
        Try
            If IsNothing(frmAdjustment) Then

                Dim contactID As Integer = 0

                If oCommitment.ReimburseERAC AndAlso TypeOf ugFinancialContacts.DataSource Is DataSet AndAlso DirectCast(ugFinancialContacts.DataSource, DataSet).Tables.Count > 0 Then
                    Dim dr As DataRow() = DirectCast(ugFinancialContacts.DataSource, DataSet).Tables(0).Select("Type='Engineer/ERAC Representative'")

                    If dr.GetUpperBound(0) >= 0 AndAlso dr(0).Item("EntityAssocID") > 0 Then
                        contactID = dr(0).Item("ContactID")
                    Else
                        For Each row As DataRow In dr
                            If row.Item("EntityAssocID") > 0 Then
                                contactID = dr(0).Item("ContactID")
                            End If
                        Next


                    End If

                End If

                frmAdjustment = New Adjustment(contactID)
                frmAdjustment.FinancialEventID = oFinancial.ID
                frmAdjustment.FinancialCommitmentID = CommitmentID
                frmAdjustment.AdjustmentID = AdjustmentID
                frmAdjustment.Balance = 0
                If ugCommitments.ActiveRow.ParentRow Is Nothing Then
                    frmAdjustment.ugCommitRow = ugCommitments.ActiveRow
                Else
                    frmAdjustment.ugCommitRow = ugCommitments.ActiveRow.ParentRow
                End If
                AddHandler frmAdjustment.Closing, AddressOf frmClosing
                AddHandler frmAdjustment.Closed, AddressOf frmClosed
            End If

            nCommitmentID = CommitmentID

            frmAdjustment.ShowDialog()
            LoadCommitmentsGrid()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
#Region "UI Control Events"
    Private Sub btnAddAdjustment_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddAdjustment.Click
        Try
            If ugCommitments.ActiveRow.Band.Index = 0 Then
                CreateModifyAdjustment(ugCommitments.ActiveRow.Cells("CommitmentID").Value, 0)
            Else
                MsgBox("Please Select A Commitment To Add Adjustment To")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnModViewAdjustment_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnModViewAdjustment.Click
        Try
            If Not Me.ugCommitments.ActiveRow Is Nothing Then
                If ugCommitments.ActiveRow.Band.Index = nAdjustmentBand And nAdjustmentBand > 0 Then
                    CreateModifyAdjustment(ugCommitments.ActiveRow.Cells("CommitmentID").Value, ugCommitments.ActiveRow.Cells("ChildID").Value)
                Else
                    MsgBox("Please Select An Adjustment To Modify")
                End If
            Else
                MsgBox("Please Select An Adjustment To Modify")
            End If


        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnDeleteAdjustment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteAdjustment.Click

        Dim oFinancialAdjustment As New MUSTER.BusinessLogic.pFinancialCommitAdjustment
        Dim nTotalBalance As Double = 0.0
        Dim nTotlaChangeOrderAdjust As Double = 0.0
        Dim nTotalUnencumbranceAdjust As Double = 0.0
        Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim ChildBand As Infragistics.Win.UltraWinGrid.UltraGridChildBand
        Dim Childrow As Infragistics.Win.UltraWinGrid.UltraGridRow

        Try
            If ugCommitments.ActiveRow.Band.Index = nAdjustmentBand And nAdjustmentBand > 0 Then
                For Each ugRow In ugCommitments.Rows
                    If Not ugRow.ChildBands Is Nothing Then
                        For Each ChildBand In ugRow.ChildBands
                            If Not ChildBand.Rows Is Nothing Then
                                If ChildBand.Rows.Count > 0 Then

                                    For Each Childrow In ChildBand.Rows
                                        'If Not Childrow.Cells.Contains("Vendor_Inv_Number") Then
                                        'If Childrow.Cells.Contains("Adjust_Type") Then
                                        If Childrow.Cells.Exists("Adjust_Type") Then
                                            If Childrow.Cells("ChildID").Value <> ugCommitments.ActiveRow.Cells("ChildID").Value Then
                                                If UCase(Childrow.Cells("Adjust_Type").Value) = UCase("change order") Then
                                                    If Not Childrow.Cells("Adjust_Amount").Value Is System.DBNull.Value Then
                                                        nTotlaChangeOrderAdjust += CDbl(Childrow.Cells("Adjust_Amount").Value)
                                                    End If
                                                ElseIf UCase(Childrow.Cells("Adjust_Type").Value) = UCase("unencumberance") Then
                                                    If Not Childrow.Cells("Adjust_Amount").Value Is System.DBNull.Value Then
                                                        nTotalUnencumbranceAdjust += CDbl(Childrow.Cells("Adjust_Amount").Value)
                                                    End If
                                                End If
                                            End If
                                        End If
                                    Next
                                End If
                            End If
                        Next
                    End If

                Next
                If ugCommitments.ActiveRow.HasParent Then
                    ' nTotalBalance = IIf(strTotalBalance <> String.Empty, CSng(strTotalBalance), 0) + nTotlaChangeOrderAdjust - nTotalUnencumbranceAdjust - IIf(strTotalPayment <> String.Empty, CSng(strTotalPayment), 0)
                    nTotalBalance = IIf(lblCommitmentValue.Text <> String.Empty, CDbl(lblCommitmentValue.Text), 0) + nTotlaChangeOrderAdjust - nTotalUnencumbranceAdjust - IIf(lblPaymentValue.Text <> String.Empty, CDbl(lblPaymentValue.Text), 0)
                End If
                'nTotalBalance = CSng(ugCommitments.ActiveRow.Cells("Commitment").Value) + nTotlaChangeOrderAdjust - nTotalUnencumbranceAdjust - CSng(ugCommitments.ActiveRow.Cells("Payment").Value)
                If nTotalBalance < 0 Then
                    MsgBox("Total Balance is less than 0.  Adjustment cannot be deleted.")
                    Exit Sub
                End If
                'If ugCommitments.ActiveRow.Band.Index = nAdjustmentBand And nAdjustmentBand > 0 Then
                If MsgBox("Are you sure you want to delete this Adjustment?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    oFinancialAdjustment.Retrieve(ugCommitments.ActiveRow.Cells("ChildID").Value)
                    oFinancialAdjustment.Deleted = True
                    oFinancialAdjustment.ModifiedBy = MusterContainer.AppUser.ID()

                    nCommitmentID = oFinancialAdjustment.CommitmentID

                    oFinancialAdjustment.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)
                    MsgBox("Adjustment Deleted", MsgBoxStyle.OKOnly, "Adjustment Deleted")
                    LoadCommitmentsGrid()
                    LoadPaymentsGrid()
                End If
            Else
                MsgBox("Please Select An Adjustment To Delete")
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

#End Region
#End Region

#Region "Payments"
#Region "UI Support Routines"
    Private Sub CreateModifyInvoice(ByVal InvoiceID As Int64)
        Try
            If ugPayments.ActiveRow.Band.Index = 0 Then
                If ugPayments.Selected.Rows.Count > 0 Then
                    If ugPayments.ActiveRow.Cells("Incomplete").Value = True Then
                        MsgBox("You cannot add an Invoice for an Incomplete Request.")
                        Exit Sub
                    End If
                End If
            End If

            If IsNothing(frmInvoice) Then

                Dim contactID As Integer = 0


                If ugPayments.ActiveRow Is Nothing Then
                    ugPayments.ActiveRow = ugPayments.Rows(0).ChildBands(0).Rows(0)
                End If

                Dim nNum As Integer = -1
                If ugPayments.ActiveRow.HasParent Then
                    nNum = ugPayments.ActiveRow.ParentRow.Cells("Reimbursement_ID").Value()
                Else
                    nNum = ugPayments.ActiveRow.Cells("Reimbursement_ID").Value()
                End If

                Dim oRem As New BusinessLogic.pFinancialReimbursement
                oRem.Retrieve(nNum)
                oCommitment.Retrieve(oRem.CommitmentID)
                oRem = Nothing


                If oCommitment.ReimburseERAC AndAlso TypeOf ugFinancialContacts.DataSource Is DataSet AndAlso DirectCast(ugFinancialContacts.DataSource, DataSet).Tables.Count > 0 Then
                    Dim dr As DataRow() = DirectCast(ugFinancialContacts.DataSource, DataSet).Tables(0).Select("Type='Engineer/ERAC Representative'")

                    If dr.GetUpperBound(0) >= 0 AndAlso dr(0).Item("EntityAssocID") > 0 Then
                        contactID = dr(0).Item("ContactID")
                    Else
                        For Each row As DataRow In dr
                            If row.Item("EntityAssocID") > 0 Then
                                contactID = dr(0).Item("ContactID")
                            End If
                        Next


                    End If

                End If

                frmInvoice = New Invoice(contactID)

                AddHandler frmInvoice.Closing, AddressOf frmClosing
                AddHandler frmInvoice.Closed, AddressOf frmClosed
            End If
            frmInvoice.FinancialEventID = oFinancial.ID
            frmInvoice.FinancialInvoiceID = InvoiceID
            If ugPayments.Selected.Rows.Count > 0 Then
                If InvoiceID <= 0 Then
                    If ugPayments.ActiveRow.Band.Index = 0 Then
                        nReimbursementID = ugPayments.ActiveRow.Cells("Reimbursement_ID").Value
                    Else
                        nReimbursementID = ugPayments.ActiveRow.ParentRow.Cells("Reimbursement_ID").Value
                    End If
                    frmInvoice.SelectedRequestId = nReimbursementID ' ugPayments.ActiveRow.Cells("Reimbursement_ID").Value
                Else
                    nReimbursementID = ugPayments.ActiveRow.ParentRow.Cells("Reimbursement_ID").Value
                End If
            End If
            frmInvoice.ShowDialog()
            LoadCommitmentsGrid()
            LoadPaymentsGrid()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
#Region "UI Control Events"
    Private Sub btnIncompleteApplication_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnIncompleteApplication.Click
        Try
            ModifyUpdateReimbursmentRequest()


        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ModifyUpdateReimbursmentRequest()
        Try
            If Not (ugPayments.Rows.Count > 0) Then
                Exit Sub
            End If
            If ugPayments.ActiveRow.Band.Index = 0 Then
                If IsNothing(frmIncompApplication) Then
                    frmIncompApplication = New IncompleteApplication
                    AddHandler frmIncompApplication.Closing, AddressOf frmClosing
                    AddHandler frmIncompApplication.Closed, AddressOf frmClosed
                End If
                frmIncompApplication.FinancialEventID = oFinancial.ID
                nReimbursementID = ugPayments.ActiveRow.Cells("Reimbursement_ID").Value
                frmIncompApplication.FinancialReimbursementID = nReimbursementID
                frmIncompApplication.CallingForm = Me
                frmIncompApplication.ShowDialog()
                LoadPaymentsGrid()
            Else
                MsgBox("Please Select A Payment Request From The Grid.")
                Exit Sub
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnAddRequest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddRequest.Click
        Try

            If IsNothing(frmIncompApplication) Then
                frmIncompApplication = New IncompleteApplication
                AddHandler frmIncompApplication.Closing, AddressOf frmClosing
                AddHandler frmIncompApplication.Closed, AddressOf frmClosed
            End If
            frmIncompApplication.FinancialEventID = oFinancial.ID
            frmIncompApplication.FinancialReimbursementID = 0
            'frmIncompApplication.FinancialCommitmentID = 0
            frmIncompApplication.CallingForm = Me
            frmIncompApplication.ShowDialog()
            If Not Me.Tag Is Nothing Then
                If Me.Tag.ToString.StartsWith("R") Then
                    If IsNumeric(Me.Tag.ToString.TrimStart("R")) Then
                        nReimbursementID = Me.Tag.ToString.TrimStart("R")
                    End If
                End If
            End If
            LoadPaymentsGrid()

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnModifyViewInvoice_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnModifyViewInvoice.Click
        Try
            If Not ugPayments.ActiveRow Is Nothing Then
                If ugPayments.ActiveRow.Band.Index = 1 Then

                    CreateModifyInvoice(ugPayments.ActiveRow.Cells("Invoices_ID").Value)
                Else
                    MsgBox("Please Select An Invoice To Modify/View.")
                    Exit Sub
                End If
            Else
                MsgBox("Please Select An Invoice To Modify/View.")
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnAddInvoice_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddInvoice.Click
        Try
            CreateModifyInvoice(0)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#End Region
#End Region

#Region "Contacts"
#Region "Owner Contacts"

    Private Sub tbCtrlOwner_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbCtrlOwner.Click
        If bolLoading Then Exit Sub
        Try
            Select Case tbCtrlOwner.SelectedTab.Name
                Case tbPageOwnerFacilities.Name
                    'If tbCtrlOwner.Contains(Me.tbPrevFacs) Then
                    '    Me.tbCtrlOwner.TabPages.RemoveAt(1)
                    'End If
                Case tbPageOwnerContactList.Name
                    If Me.lblOwnerIDValue.Text <> String.Empty Then
                        LoadContacts(ugOwnerContacts, Integer.Parse(Me.lblOwnerIDValue.Text), UIUtilsGen.EntityTypes.Owner)
                    End If
                    'Dim dsOwnerContacts As DataSet
                    'dsOwnerContacts = MusterContainer.pConStruct.GetAll()
                    'dsOwnerContacts.Tables(0).DefaultView.RowFilter = "MODULEID = 891 And ENTITYID = " + pOwn.ID.ToString
                    'dsContacts.Tables(0).DefaultView.Sort = "CONTACT_NAME ASC"
                    'ugOwnerContacts.DataSource = dsOwnerContacts.Tables(0).DefaultView
                    'ugOwnerContacts.DisplayLayout.Bands(0).Columns("Parent_Contact").Hidden = True
                    'ugOwnerContacts.DisplayLayout.Bands(0).Columns("Child_Contact").Hidden = True
                    'ugOwnerContacts.DisplayLayout.Bands(0).Columns("ContactAssocID").Hidden = True
                    'ugOwnerContacts.DisplayLayout.Bands(0).Columns("ContactID").Hidden = True
                    'ugOwnerContacts.DisplayLayout.Bands(0).Columns("ModuleID").Hidden = True
                    'ugOwnerContacts.DisplayLayout.Bands(0).Columns("ContactType").Hidden = True
                Case tbPageOwnerDocuments.Name
                    UCOwnerDocuments.LoadDocumentsGrid(Integer.Parse(Me.lblOwnerIDValue.Text), UIUtilsGen.EntityTypes.Owner, UIUtilsGen.ModuleID.Financial)
            End Select
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnOwnerAddSearchContact_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOwnerAddSearchContact.Click
        Try
            pConStruct = New MUSTER.BusinessLogic.pContactStruct
            If tbCntrlFinancial.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL" Then
                objCntSearch = New ContactSearch(pOwn.ID, 9, "Financial", pConStruct)
            Else
                objCntSearch = New ContactSearch(pOwn.Facility.ID, 6, "Financial", pConStruct)
            End If
            'objCntSearch.Show()
            objCntSearch.ShowDialog()

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnOwnerModifyContact_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOwnerModifyContact.Click
        Try
            ModifyContact(ugOwnerContacts)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnOwnerAssociateContact_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOwnerAssociateContact.Click
        Try
            If tbCntrlFinancial.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL" Then
                AssociateContact(ugOwnerContacts, pOwn.ID, 9)
            Else
                AssociateContact(ugOwnerContacts, pOwn.Facility.ID, 6)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnOwnerDeleteContact_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOwnerDeleteContact.Click
        Try
            If tbCntrlFinancial.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL" Then
                DeleteContact(ugOwnerContacts, pOwn.ID)
            Else
                DeleteContact(ugOwnerContacts, pOwn.Facility.ID)
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()

        End Try
    End Sub

    Private Sub btnOwnerContactClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Me.Close()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()

        End Try
    End Sub

    Private Sub chkOwnerShowActiveOnly_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkOwnerShowActiveOnly.CheckedChanged
        Try
            ''If Not dsContacts Is Nothing Then
            SetFilter()
            'End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub chkOwnerShowContactsforAllModules_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkOwnerShowContactsforAllModules.CheckedChanged
        Try
            'If Not dsContacts Is Nothing Then
            SetFilter()
            'End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub chkOwnerShowRelatedContacts_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkOwnerShowRelatedContacts.CheckedChanged
        Try
            'If Not dsContacts Is Nothing Then
            SetFilter()
            'End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Sub SetFilter()
        Dim strMode As String = String.Empty
        Dim dsContactsLocal As DataSet
        Dim bolActive As Boolean = False
        Dim strEntities As String = String.Empty
        Dim nModuleID As Integer = 0
        Dim nEntityID As Integer = 0
        Dim strEntityAssocIDs As String = String.Empty
        Dim nEntityType As Integer = 0
        Dim nRelatedEntityType As Integer = 0
        Try
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            If tbCntrlFinancial.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL" Then
                nEntityID = pOwn.ID
                nEntityType = 9
                strMode = "OWNER"
            ElseIf tbCntrlFinancial.SelectedTab.Name.ToUpper = "TBPAGEFACILITYDETAIL" Then
                nEntityID = pOwn.Facility.ID
                nEntityType = 6
            End If

            If chkOwnerShowActiveOnly.Checked Then
                bolActive = True
            Else
                bolActive = False
            End If

            If chkOwnerShowContactsforAllModules.Checked Then
                'User has the ability to view the contacts associated for the entity in other modules
                If strMode <> "OWNER" Then
                    Dim strFilterForAllModules As String = pConStruct.GetContactsForAllModules(pOwn.Facility.ID.ToString)
                    strEntityAssocIDs = strFilterForAllModules
                    nModuleID = 0
                End If
            Else
                nModuleID = 616
            End If
            If chkOwnerShowRelatedContacts.Checked Then
                strEntities = strFacilityIdTags
                nRelatedEntityType = 6
            End If

            UIUtilsGen.LoadContacts(ugOwnerContacts, nEntityID, nEntityType, pConStruct, nModuleID, strEntities, bolActive, strEntityAssocIDs, nRelatedEntityType)


            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Dim strMode As String = String.Empty
            'Try
            '    strFilterString = String.Empty
            '    Dim strEntityID As String
            '    If tbCntrlFinancial.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL" Then
            '        strEntityID = pOwn.ID.ToString
            '        strMode = "OWNER"
            '    Else
            '        strEntityID = pOwn.Facility.ID.ToString
            '    End If

            '    If chkOwnerShowActiveOnly.Checked Then
            '        strFilterString = "ACTIVE = 1"
            '    Else
            '        strFilterString = ""
            '    End If

            '    If chkOwnerShowContactsforAllModules.Checked Then
            '        'strFilterString += ""
            '        If strMode <> "OWNER" Then
            '            Dim strFilterForAllModules As String = pConStruct.GetContactsForAllModules(pOwn.Facility.ID.ToString)
            '            If strFilterString = String.Empty Then
            '                'strFilterString += "ENTITYID = " + strEntityID + " OR " + IIf(Not strFilterForAllModules = String.Empty, " entityassocid in (" + strFilterForAllModules + ")", "")
            '                strFilterString += "ENTITYID = " + strEntityID + IIf(Not strFilterForAllModules = String.Empty, " OR " + " entityassocid in (" + strFilterForAllModules + ")", "")
            '            Else
            '                'strFilterString += "AND ENTITYID = " + strEntityID + " OR " + IIf(Not strFilterForAllModules = String.Empty, " entityassocid in (" + strFilterForAllModules + ")", "")
            '                strFilterString += "AND ENTITYID = " + strEntityID + IIf(Not strFilterForAllModules = String.Empty, " OR " + " entityassocid in (" + strFilterForAllModules + ")", "")
            '            End If
            '        Else
            '            If strFilterString = String.Empty Then
            '                strFilterString += "ENTITYID = " + strEntityID
            '            Else
            '                strFilterString += "AND ENTITYID = " + strEntityID
            '            End If

            '        End If
            '    Else
            '        If strFilterString = String.Empty Then
            '            strFilterString = " MODULEID = 616 And ENTITYID = " + strEntityID
            '        Else
            '            strFilterString += " AND MODULEID = 616 And ENTITYID = " + strEntityID
            '        End If
            '    End If

            '    If chkOwnerShowRelatedContacts.Checked Then
            '        If strFilterString = String.Empty Then
            '            strFilterString = " (ENTITYID = " + strEntityID + IIf(Not strFacilityIdTags = String.Empty, " OR ENTITYID in (" + strFacilityIdTags + ")", "") + ")"
            '        Else
            '            strFilterString += " AND (ENTITYID = " + strEntityID + IIf(Not strFacilityIdTags = String.Empty, " OR ENTITYID in (" + strFacilityIdTags + ")", "") + ")"
            '        End If
            '    Else
            '        strFilterString += ""
            '    End If

            '    dsContacts.Tables(0).DefaultView.RowFilter = strFilterString
            '    ugOwnerContacts.DataSource = dsContacts.Tables(0).DefaultView

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "FinancialEvent Contacts"
    Private Sub chkFinancialShowContactsForAllModules_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFinancialShowContactsForAllModules.CheckedChanged
        Try
            'If Not dsContacts Is Nothing Then
            FinancialSetFilter()
            'End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub chkFinancialShowRelatedContacts_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFinancialShowRelatedContacts.CheckedChanged
        Try
            'If Not dsContacts Is Nothing Then
            FinancialSetFilter()
            'End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub chkFinancialShowActive_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFinancialShowActive.CheckedChanged
        Try
            'If Not dsContacts Is Nothing Then
            FinancialSetFilter()
            'End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnFinancialContactAddorSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFinancialContactAddorSearch.Click
        Try
            pConStruct = New MUSTER.BusinessLogic.pContactStruct
            objCntSearch = New ContactSearch(oFinancial.ID, 32, "Financial", pConStruct)
            'objCntSearch.Show()
            objCntSearch.ShowDialog()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnFinancialContactModify_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFinancialContactModify.Click
        Try
            ModifyContact(ugFinancialContacts)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnFinancialContactDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFinancialContactDelete.Click
        Try
            DeleteContact(ugFinancialContacts, oFinancial.ID)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnFinancialContactAssociate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFinancialContactAssociate.Click
        Try
            AssociateContact(ugFinancialContacts, oFinancial.ID, 32)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()

        End Try
    End Sub

    Private Sub FinancialSetFilter()
        Dim strMode As String = String.Empty
        Dim dsContactsLocal As DataSet
        Dim bolActive As Boolean = False
        Dim strEntities As String = String.Empty
        Dim nModuleID As Integer = 0
        Dim nEntityID As Integer = 0
        Dim strEntityAssocIDs As String = String.Empty
        Dim nEntityType As Integer = 0
        Dim nRelatedEntityType As Integer = 0
        Try
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If chkFinancialShowActive.Checked Then
                bolActive = True
            Else
                bolActive = False
            End If
            If chkFinancialShowContactsForAllModules.Checked Then
                Dim strFilterForAllModules As String = pConStruct.GetContactsForAllModules(pOwn.Facility.ID.ToString)
                nEntityID = pOwn.Facilities.ID

                nEntityType = 6
                strEntityAssocIDs = String.Format("{0}FOR:{1}", strFilterForAllModules, oFinancial.ID)
                nModuleID = 0

            Else
                nEntityType = 32
                nEntityID = oFinancial.ID
                nModuleID = 616

            End If

            If chkFinancialShowRelatedContacts.Checked Then
                strEntities = strFinancialEventIdTags
                nRelatedEntityType = 32
                'strFilterString += " OR " + IIf(Not strFinancialEventIdTags = String.Empty, " ENTITYID in (" + strFinancialEventIdTags + "))", "")
            End If

            UIUtilsGen.LoadContacts(ugFinancialContacts, nEntityID, nEntityType, pConStruct, nModuleID, strEntities, bolActive, strEntityAssocIDs, nRelatedEntityType)

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'strFilterString = String.Empty
            'Dim strEntityID As String

            'strEntityID = oFinancial.ID.ToString

            'If chkFinancialShowActive.Checked Then
            '    strFilterString = "(ACTIVE = 1"
            'Else
            '    strFilterString = "("
            'End If

            'If chkFinancialShowContactsForAllModules.Checked Then
            '    'If strFilterString = "(" Then
            '    '    strFilterString += "ENTITYID = " + pOwn.Facility.ID.ToString + " OR ENTITYID = " + oFinancial.ID.ToString
            '    'Else
            '    '    strFilterString += " AND ENTITYID = " + pOwn.Facility.ID.ToString + " OR ENTITYID = " + oFinancial.ID.ToString
            '    'End If
            '    Dim strFilterForAllModules As String = pConStruct.GetContactsForAllModules(pOwn.Facility.ID.ToString)
            '    If strFilterString = "(" Then
            '        strFilterString += "ENTITYID = " + pOwn.Facilities.ID.ToString + " OR " + IIf(Not strFilterForAllModules = String.Empty, " entityassocid in (" + strFilterForAllModules + ")", "")
            '    Else
            '        strFilterString += "AND ENTITYID = " + pOwn.Facilities.ID.ToString + " OR " + IIf(Not strFilterForAllModules = String.Empty, " entityassocid in (" + strFilterForAllModules + ")", "")
            '    End If
            'Else
            '    If strFilterString = "(" Then
            '        strFilterString += " MODULEID = 616 And ENTITYID = " + strEntityID
            '    Else
            '        strFilterString += " AND MODULEID = 616 And ENTITYID = " + strEntityID
            '    End If
            'End If

            'If chkFinancialShowRelatedContacts.Checked Then
            '    strFilterString += " OR " + IIf(Not strFinancialEventIdTags = String.Empty, " ENTITYID in (" + strFinancialEventIdTags + "))", "")
            'Else
            '    strFilterString += ")"
            'End If

            'dsContacts.Tables(0).DefaultView.RowFilter = strFilterString
            'ugFinancialContacts.DataSource = dsContacts.Tables(0).DefaultView

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "Common Functions"
    Private Sub LoadContacts(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal EntityID As Integer, ByVal EntityType As Integer)
        
        Try

            Dim nVendorID As Integer

            UIUtilsGen.LoadContacts(ugGrid, EntityID, EntityType, pConStruct, 616)

            If tbCntrlFinancial.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL" Or tbCntrlFinancial.SelectedTab.Name.ToUpper = "TBPAGEFACILITYDETAIL" Then

                Me.chkOwnerShowActiveOnly.Checked = False
                Me.chkOwnerShowActiveOnly.Checked = True
            ElseIf tbCntrlFinancial.SelectedTab.Name.ToUpper = "TBPAGEFINANCIALEVENT" Then
                Me.chkFinancialShowActive.Checked = False
                Me.chkFinancialShowActive.Checked = True
            End If

            If EntityType = 32 Then
                nVendorID = oFinancial.VendorID
                txtVendor.Text = String.Empty
                txtVendorAddress.Text = String.Empty
                Me.txtVendorNumberValue.Text = ""
                Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
                For Each ugrow In ugGrid.Rows
                    If ugrow.Cells("LetterContactType").Value = 1185 Then
                        If ugrow.Cells("IsPerson").Value = True And Not (ugrow.Cells("Assoc Company").Value Is System.DBNull.Value Or ugrow.Cells("Assoc Company").Value = String.Empty) Then
                            Dim dtCompanyAddress As DataTable
                            oFinancial.VendorID = ugrow.Cells("ContactID").Value
                            txtVendor.Text = ugrow.Cells("Assoc Company").Text
                            dtCompanyAddress = pConStruct.ContactDatum.GetCompanyAddress(CInt(ugrow.Cells("Child_Contact").Value))
                            LoadVendorAddress(, dtCompanyAddress)
                            pConStruct.ContactDatum.Retrieve(ugrow.Cells("child_contact").Value, False)
                            If pConStruct.ContactDatum.contactDatumInfo.VendorNumber = "0" Then
                                Me.txtVendorNumberValue.Text = ""
                            Else
                                Me.txtVendorNumberValue.Text = pConStruct.ContactDatum.contactDatumInfo.VendorNumber
                            End If
                        Else
                            oFinancial.VendorID = ugrow.Cells("ContactID").Value
                            txtVendor.Text = ugrow.Cells("Contact_Name").Text
                            LoadVendorAddress(ugrow)
                            Me.txtVendorNumberValue.Text = ugrow.Cells("Vendor_Number").Value
                            Exit For
                        End If
                    End If
                Next
                If txtVendor.Text = String.Empty Then
                    oFinancial.VendorID = 0
                End If
                If nVendorID <> oFinancial.VendorID Then
                    If oFinancial.ID <= 0 Then
                        oFinancial.CreatedBy = MusterContainer.AppUser.ID
                    Else
                        oFinancial.ModifiedBy = MusterContainer.AppUser.ID
                    End If
                    oFinancial.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal, , False, bolFromTechnical)
                    bolFromTechnical = False
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If
                End If
            End If
            bolFromTechnical = False
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Function ModifyContact(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid)
        Try

            If UIUtilsGen.ModifyContact(ugGrid, 616, pConStruct) Then
                Me.Contact_ContactAdded()
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function

    Private Function AssociateContact(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal nEntityID As Integer, ByVal nEntityType As Integer)
        Try
            If UIUtilsGen.AssociateContact(ugGrid, nEntityID, nEntityType, 616, pConStruct) Then
                Me.Contact_ContactAdded()
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function

    Private Function DeleteContact(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal nEntityID As Integer)
        Try

            If UIUtilsGen.DeleteContact(ugGrid, nEntityID, 616, pConStruct) Then
                Me.Contact_ContactAdded()
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function

#End Region

#Region "Close Events"
    Private Sub Search_ContactAdded() Handles objCntSearch.ContactAdded
        If tbCntrlFinancial.SelectedTab.Name.ToUpper = "TBPAGEFINANCIALEVENT" Then
            LoadContacts(ugFinancialContacts, oFinancial.ID, 32)
            chkFinancialShowContactsForAllModules.Checked = False
        Else
            If tbCntrlFinancial.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL" Then
                LoadContacts(ugOwnerContacts, pOwn.ID, 9)
            Else
                LoadContacts(ugOwnerContacts, pOwn.Facility.ID, 6)
            End If
        End If
    End Sub

    Private Sub Contact_ContactAdded()
        If tbCntrlFinancial.SelectedTab.Name.ToUpper = "TBPAGEFINANCIALEVENT" Then
            LoadContacts(ugFinancialContacts, oFinancial.ID, 32)
            chkFinancialShowContactsForAllModules.Checked = False
        Else
            If tbCntrlFinancial.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL" Then
                LoadContacts(ugOwnerContacts, pOwn.ID, 9)
            Else
                LoadContacts(ugOwnerContacts, pOwn.Facility.ID, 6)
            End If
        End If
    End Sub

    
    Private Sub objCntSearch_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles objCntSearch.Closing
        If Not objCntSearch Is Nothing Then
            objCntSearch = Nothing
        End If
    End Sub
#End Region


#End Region

#Region "Comments"
    Private Sub CommentsMaintenance(Optional ByVal sender As System.Object = Nothing, Optional ByVal e As System.EventArgs = Nothing, Optional ByVal bolSetCounts As Boolean = False, Optional ByVal resetBtnColor As Boolean = False)
        Dim SC As ShowComments
        Dim nEntityType As Integer = 0
        Dim nEntityID As Integer = 0
        Dim strEntityName As String = String.Empty
        Dim oComments As MUSTER.BusinessLogic.pComments
        Dim bolEnableShowAllModules As Boolean = True
        Dim nCommentsCount As Integer = 0
        Try
            Select Case tbCntrlFinancial.SelectedTab.Name
                Case tbPageOwnerDetail.Name
                    nEntityType = UIUtilsGen.EntityTypes.Owner
                    strEntityName = "Owner : " + CStr(pOwn.ID) + " " + Me.txtOwnerName.Text
                    oComments = pOwn.Comments
                    nEntityID = pOwn.ID
                Case tbPageFacilityDetail.Name
                    strEntityName = "Facility : " + CStr(pOwn.Facilities.ID) + " " + pOwn.Facilities.Name
                    oComments = pOwn.Facilities.Comments
                    nEntityID = pOwn.Facilities.ID
                    nEntityType = UIUtilsGen.EntityTypes.Facility
                Case tbPageFinancialEvent.Name
                    strEntityName = "Facility : " + CStr(pOwn.Facilities.ID) + " " + pOwn.Facilities.Name + ", Financial Event : " + CStr(oFinancial.Sequence)
                    oComments = New MUSTER.BusinessLogic.pComments
                    nEntityID = oFinancial.ID
                    nEntityType = UIUtilsGen.EntityTypes.FinancialEvent
                    bolEnableShowAllModules = False
                Case Else
                    Exit Sub
            End Select
            If Not resetBtnColor Then
                SC = New ShowComments(nEntityID, nEntityType, IIf(bolSetCounts, "", "Financial"), strEntityName, oComments, Me.Text, , bolEnableShowAllModules)
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
            ElseIf nEntityType = UIUtilsGen.EntityTypes.FinancialEvent Then
                If nCommentsCount > 0 Then
                    btnFinancialComments.BackColor = Drawing.Color.FromName(UIUtilsGen.CommentsButton_HasCmts)
                Else
                    btnFinancialComments.BackColor = Drawing.Color.FromName(UIUtilsGen.CommentsButton_NoCmts)
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "Flags"

#Region "UI Support Routines"
    Private Sub SF_FlagAdded(ByVal entityID As Integer, ByVal entityType As Integer, ByVal [Module] As String, ByVal ParentFormText As String) Handles SF.FlagAdded
        ' New event declared to allow registration to trigger
        '   container to check flag status.  This is bogus, but
        '   the current design which permits something other than
        '   MusterContainer to create a registration object (i.e.
        '   OwnerSearchResults creates it, so MusterContainer
        '   doesn't know about it) precludes use of events to notify
        '   MusterContainer of an request.  Therefore, the event FlagsChanged
        '   cannot be fired here and then caught by the MusterContainer.
        '   
        Dim MyFrm As MusterContainer
        MyFrm = Me.MdiParent
        'oEntity.GetEntity("Owner")
        If Not MyFrm Is Nothing Then
            Select Case Me.tbCntrlFinancial.SelectedTab.Name
                Case tbPageFinancialEvent.Name
                    MyFrm.FlagsChanged(lblFacilityIDValue.Text, UIUtilsGen.EntityTypes.Facility, [Module], ParentFormText, entityID, entityType)
                Case Else
                    MyFrm.FlagsChanged(entityID, entityType, [Module], ParentFormText)
            End Select
        End If
    End Sub

    Private Sub SF_RefreshCalendar() Handles SF.RefreshCalendar
        Dim mc As MusterContainer = Me.MdiParent
        mc.RefreshCalendarInfo()
        mc.LoadDueToMeCalendar()
        mc.LoadToDoCalendar()
    End Sub

    Private Sub FlagMaintenance(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Select Case Me.tbCntrlFinancial.SelectedTab.Name
                Case tbPageOwnerDetail.Name
                    SF = New ShowFlags(pOwn.ID, UIUtilsGen.EntityTypes.Owner, "Financial")
                Case tbPageFacilityDetail.Name
                    SF = New ShowFlags(pOwn.Facilities.ID, UIUtilsGen.EntityTypes.Facility, "Financial")
                Case tbPageFinancialEvent.Name
                    SF = New ShowFlags(oFinancial.ID, UIUtilsGen.EntityTypes.FinancialEvent, "Financial", , , CStr(oFinancial.Sequence))
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

#Region "UI Control Events"
    Private Sub btnOwnerFlag_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOwnerFlag.Click
        FlagMaintenance(sender, e)
        'Dim MyFrm As MusterContainer
        'MyFrm = Me.MdiParent
        'oEntity.GetEntity("Owner")
        'MyFrm.FlagsChanged(Me.lblOwnerIDValue.Text, 0, 0, 0, "Financial", Me.Text)
    End Sub
    Private Sub btnFacFlags_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFacFlags.Click
        FlagMaintenance(sender, e)
        'Dim MyFrm As MusterContainer
        'MyFrm = Me.MdiParent
        'oEntity.GetEntity("Facility")
        'MyFrm.FlagsChanged(Me.lblOwnerIDValue.Text, Me.lblFacilityIDValue.Text, 0, 0, "Financial", Me.Text)
    End Sub
    Private Sub btnFinancialFlags_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFinancialFlags.Click
        FlagMaintenance(sender, e)
    End Sub
#End Region

#End Region

    Public Sub LoadFinancialData(ByVal nFinancialID As Int64)
        bolEnableDeleteEvent = True
        tbPageFinancialEvent.Enabled = True
        bolLoading = True
        tbCntrlFinancial.SelectedTab = Me.tbPageFinancialEvent
        oFinancial.Retrieve(nFinancialID)
        nLastEventID = nFinancialID
        nCurrentEventID = nFinancialID

        PopulateFinancialStatus()

        If nFinancialID > 0 Then
            'btnDeleteEvent.Enabled = True
            btnFinancialComments.Enabled = True
            btnFinancialFlags.Enabled = True
            PnlCommitmentButtons.Enabled = True
            pnlPaymentButtons.Enabled = True
            lblTechEvent.Visible = True
            cmbTechEvent.Visible = False
            lblTechEvent.Text = oFinancial.TecEventIDDesc
            lblFinancialIDValue.Text = oFinancial.Sequence
            cmbFinancialStatus.SelectedValue = oFinancial.Status
            UIUtilsGen.SetDatePickerValue(dtFinancialStart, oFinancial.StartDate)
            UIUtilsGen.SetDatePickerValue(dtFinancialClosedDate, oFinancial.ClosedDate)
            pnlContacts.Visible = True
            pnlContactDetails.Visible = True
            lblOwnerLastEditedBy.Text = "Last Edited By : " & IIf(oFinancial.ModifiedBy = String.Empty, oFinancial.CreatedBy.ToString, oFinancial.ModifiedBy.ToString)
            lblOwnerLastEditedOn.Text = "Last Edited On : " & IIf(oFinancial.ModifiedOn = CDate("01/01/0001"), oFinancial.CreatedOn.ToString, oFinancial.ModifiedOn.ToString)
            Me.Text = "Financial Events - Facility #" & lblFacilityIDValue.Text & " (" & txtFacilityName.Text & ")"
            If oFinancial.Status = 1028 Then
                btnAddCommitment.Enabled = False
            Else
                btnAddCommitment.Enabled = True
            End If
            CommentsMaintenance(, , True)
        Else
            btnDeleteEvent.Enabled = False
            btnFinancialComments.Enabled = False
            btnFinancialFlags.Enabled = False
            PnlCommitmentButtons.Enabled = False
            pnlPaymentButtons.Enabled = False
            lblTechEvent.Visible = False
            cmbTechEvent.Visible = True
            lblFinancialIDValue.Text = "New"
            cmbFinancialStatus.SelectedIndex = 1
            dtFinancialStart.Value = Now.Date


            oFinancial.Status = cmbFinancialStatus.SelectedValue
            oFinancial.StartDate = dtFinancialStart.Value.Date
            CommentsMaintenance(, , True, True)
        End If
        LoadTechInfo()
        'If oFinancial.VendorID > 0 Then
        '    'cmbVendor.SelectedValue = oFinancial.VendorID
        '    'txtVendor.Text = oFinancial.VendorID
        '    LoadVendorAddress()
        'End If
        LoadDropDowns()


        LoadCommitmentsGrid()
        LoadPaymentsGrid()
        'Load Contacts
        LoadContacts(ugFinancialContacts, oFinancial.ID, 32)
        chkFinancialShowContactsForAllModules.Checked = False
        chkShowCommitments.Checked = False
        'If nFinancialID > 0 And Not tbCtrlFinancialContacts.TabPages.Contains(tbPageOwnerContactList) Then
        '    tbCtrlFinancialContacts.TabPages.Add(tbPageOwnerContactList)
        'End If
        'LoadContacts(ugOwnerContacts, nFinancialID, )
        btnDeleteEvent.Enabled = bolEnableDeleteEvent
        Dim MyFrm As MusterContainer = Me.MdiParent
        If Not MyFrm Is Nothing Then
            MyFrm.FlagsChanged(Me.lblFacilityIDValue.Text, UIUtilsGen.EntityTypes.Facility, "Financial", Me.Text, nFinancialID, UIUtilsGen.EntityTypes.FinancialEvent)
        End If
        bolLoading = False
    End Sub

    Private Sub LoadTechInfo()
        Dim oCompany As New MUSTER.BusinessLogic.pCompany
        Dim oCompanyAddress As New MUSTER.BusinessLogic.pComAddress

        ugCommitments.Width = pnlCommitmentsDetails.Width - 50
        ugPayments.Width = pnlPaymentsDetails.Width - 50
        Try

            If (cmbTechEvent.Text <> "" Or lblTechEvent.Text <> "") And lblFacilityIDValue.Text <> "" Then
                If lblTechEvent.Text <> "" Then
                    oTechnical.Retrieve(lblFacilityIDValue.Text, lblTechEvent.Text)
                Else
                    oTechnical.Retrieve(lblFacilityIDValue.Text, cmbTechEvent.Text)
                End If

                lblTechStartDateValue.Text = oTechnical.Started.ToShortDateString
                lblPMValue.Text = oTechnical.PMDesc
                lblMGPTFStatusValue.Text = oTechnical.MGPTFStatusDesc
                lblTechStatusValue.Text = oTechnical.TechnicalStatusDesc

                oCompany.Retrieve(oTechnical.ERAC)

                If oCompany.ID <> 0 AndAlso oCompany.COMPANY_NAME.Length > 0 Then
                    lblEngineeringFirmValue.Text = oCompany.COMPANY_NAME & IIf(oCompany.PRO_ENGIN = String.Empty, "", vbCrLf & oCompany.PRO_ENGIN)

                    If oCompany.PRO_ENGIN_ADD_ID > 0 Then
                        oCompanyAddress.Retrieve(oCompany.PRO_ENGIN_ADD_ID)
                        lblEngineeringFirmValue.Text &= vbCrLf & oCompanyAddress.AddressLine1 & IIf(oCompanyAddress.AddressLine2 = String.Empty, "", vbCrLf & oCompanyAddress.AddressLine2) & vbCrLf & oCompanyAddress.City & ", " & vbCrLf & oCompanyAddress.State & " " & vbCrLf & oCompanyAddress.Zip
                    End If

                Else
                    lblEngineeringFirmValue.Text = String.Empty
                End If
            End If

            Dim futurePlans As New BusinessLogic.pTecFinancialActivityPlanner
            futurePlans.Retrieve(Me.oTechnical.ID)

            If futurePlans.ID <> 0 Then
                btnPlanning.Enabled = True
            End If

            futurePlans = Nothing

        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub LoadDropDowns()
        Try
            'bolLoading = True
            'Dim dtVendorList As DataTable = oFinancial.PopulateVendorList


            ''cmbVendor.DataSource = dtVendorList
            ''cmbVendor.DisplayMember = "VendorName"
            ''cmbVendor.ValueMember = "ContactID"
            ''cmbVendor.SelectedIndex = -1

            'txtVendorAddress.Text = ""
            'bolLoading = False
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Load DropDowns " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()

        End Try


    End Sub

    Private Sub LoadCommitmentsGrid()
        Dim dsLocal As DataSet
        Dim tmpBand As Int16
        Dim dtTotals As DataTable

        dtTotals = oFinancial.CommitmentTotalsDatatable(0, chkShowCommitments.Checked, False)
        dsLocal = oFinancial.CommitmentGridDataset(chkShowCommitments.Checked)
        If dsLocal.Tables.Count > 0 Then
            dsLocal.Tables(0).DefaultView.Sort = "ApprovedDate desc"
            ugCommitments.DataSource = dsLocal.Tables(0).DefaultView
        Else
            ugCommitments.DataSource = dsLocal
        End If
        'ugCommitments.DataSource = dsLocal
        ugCommitments.Rows.CollapseAll(True)
        ugCommitments.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        ugCommitments.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        ugCommitments.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        If dsLocal.Tables(0).Rows.Count > 0 Then ' Does Commitment Table have rows
            ugCommitments.DisplayLayout.Bands(0).Override.MaxSelectedRows = 1

            ugCommitments.DisplayLayout.Bands(0).Columns("FIN_EVENT_ID").Hidden = True
            'ugCommitments.DisplayLayout.Bands(0).Columns("CommitmentID").Hidden = True
            ugCommitments.DisplayLayout.Bands(0).Columns("Document_Location").Hidden = True
            '
            ugCommitments.DisplayLayout.Bands(0).Columns("ApprovedDate").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugCommitments.DisplayLayout.Bands(0).Columns("PONumber").Header.Caption = "PO#"
            ugCommitments.DisplayLayout.Bands(0).Columns("ApprovedDate").Header.Caption = "Approved"
            ugCommitments.DisplayLayout.Bands(0).Columns("Funding_Type").Header.Caption = "Funding Type"
            ugCommitments.DisplayLayout.Bands(0).Columns("ThirdPartyPayment").Header.Caption = "3rd"


            ugCommitments.DisplayLayout.Bands(0).Columns("PONumber").Width = 100
            ugCommitments.DisplayLayout.Bands(0).Columns("ApprovedDate").Width = 75
            ugCommitments.DisplayLayout.Bands(0).Columns("Task").Width = 124
            ugCommitments.DisplayLayout.Bands(0).Columns("Funding_Type").Width = 100
            ugCommitments.DisplayLayout.Bands(0).Columns("ThirdPartyPayment").Width = 35
            ugCommitments.DisplayLayout.Bands(0).Columns("Commitment").Width = 100
            ugCommitments.DisplayLayout.Bands(0).Columns("Adjustment").Width = 100
            ugCommitments.DisplayLayout.Bands(0).Columns("Payment").Width = 100
            ugCommitments.DisplayLayout.Bands(0).Columns("Balance").Width = 100

            ugCommitments.DisplayLayout.Bands(0).Columns("PONumber").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            ugCommitments.DisplayLayout.Bands(0).Columns("ApprovedDate").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugCommitments.DisplayLayout.Bands(0).Columns("Task").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            ugCommitments.DisplayLayout.Bands(0).Columns("Funding_Type").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
            ugCommitments.DisplayLayout.Bands(0).Columns("ThirdPartyPayment").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            ugCommitments.DisplayLayout.Bands(0).Columns("Commitment").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugCommitments.DisplayLayout.Bands(0).Columns("Adjustment").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugCommitments.DisplayLayout.Bands(0).Columns("Payment").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugCommitments.DisplayLayout.Bands(0).Columns("Balance").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugCommitments.DisplayLayout.Bands(0).Columns("Comments").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left

            ugCommitments.DisplayLayout.Bands(0).Columns("PONumber").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            ugCommitments.DisplayLayout.Bands(0).Columns("ApprovedDate").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            ugCommitments.DisplayLayout.Bands(0).Columns("Task").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            ugCommitments.DisplayLayout.Bands(0).Columns("Funding_Type").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            ugCommitments.DisplayLayout.Bands(0).Columns("ThirdPartyPayment").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            ugCommitments.DisplayLayout.Bands(0).Columns("Commitment").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            ugCommitments.DisplayLayout.Bands(0).Columns("Adjustment").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            ugCommitments.DisplayLayout.Bands(0).Columns("Payment").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            ugCommitments.DisplayLayout.Bands(0).Columns("Balance").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            ugCommitments.DisplayLayout.Bands(0).Columns("Comments").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
        End If
        tmpBand = 0
        pnlCommitmentsTotals.Left = 410
        If dsLocal.Tables(1).Rows.Count > 0 Then ' Does Invoice Table have rows
            tmpBand += 1
            ugCommitments.DisplayLayout.Bands(tmpBand).Override.MaxSelectedRows = 1

            ugCommitments.DisplayLayout.Bands(tmpBand).Columns("FIN_EVENT_ID").Hidden = True
            ugCommitments.DisplayLayout.Bands(tmpBand).Columns("CommitmentID").Hidden = True
            ugCommitments.DisplayLayout.Bands(tmpBand).Columns("ChildID").Hidden = True
            ugCommitments.DisplayLayout.Bands(tmpBand).Columns("Requested_amount").Hidden = True
            ugCommitments.DisplayLayout.Bands(tmpBand).Columns("Reimbursement_id").Hidden = True

            ugCommitments.DisplayLayout.Bands(tmpBand).Columns("Payment_Date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugCommitments.DisplayLayout.Bands(tmpBand).Columns("Payment_Date").Header.Caption = "Date"
            ugCommitments.DisplayLayout.Bands(tmpBand).Columns("Paid").Header.Caption = "Paid"
            ugCommitments.DisplayLayout.Bands(tmpBand).Columns("Final").Header.Caption = "Final"
            ugCommitments.DisplayLayout.Bands(tmpBand).Columns("PONumber").Header.Caption = "PO #"
            ugCommitments.DisplayLayout.Bands(tmpBand).Columns("Vendor_Inv_Number").Header.Caption = "Invoice #"

            ugCommitments.DisplayLayout.Bands(tmpBand).Columns("Paid").ColSpan = 2
            ugCommitments.DisplayLayout.Bands(tmpBand).Columns("Paid").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugCommitments.DisplayLayout.Bands(tmpBand).Columns("Comment").ColSpan = 4

            pnlCommitmentsTotals.Left = 435
        End If

        If dsLocal.Tables(2).Rows.Count > 0 Then ' Does Adjustment Table have rows
            tmpBand += 1
            ugCommitments.DisplayLayout.Bands(tmpBand).Columns("Fin_Event_ID").Hidden = True
            ugCommitments.DisplayLayout.Bands(tmpBand).Columns("COMMITMENTID").Hidden = True
            ugCommitments.DisplayLayout.Bands(tmpBand).Columns("ChildID").Hidden = True

            ugCommitments.DisplayLayout.Bands(tmpBand).Columns("Adjust_Date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            ugCommitments.DisplayLayout.Bands(tmpBand).Columns("Adjust_Date").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugCommitments.DisplayLayout.Bands(tmpBand).Columns("Adjust_Amount").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            ugCommitments.DisplayLayout.Bands(tmpBand).Columns("Adjust_Date").Header.Caption = "Date"
            ugCommitments.DisplayLayout.Bands(tmpBand).Columns("Adjust_Type").Header.Caption = "Type"
            ugCommitments.DisplayLayout.Bands(tmpBand).Columns("Adjust_Amount").Header.Caption = "Amount"
            ugCommitments.DisplayLayout.Bands(tmpBand).Columns("Comments").Header.Caption = "Comments"
            ugCommitments.DisplayLayout.Bands(tmpBand).Columns("Director_App_Req").Header.Caption = "Dir. App. Reqd."
            ugCommitments.DisplayLayout.Bands(tmpBand).Columns("Fin_App_Req").Header.Caption = "Fin. App. Reqd."
            ugCommitments.DisplayLayout.Bands(tmpBand).Columns("Approved").Header.Caption = "Approved"

            ugCommitments.DisplayLayout.Bands(tmpBand).Columns("Fin_App_Req").ColSpan = 2
            ugCommitments.DisplayLayout.Bands(tmpBand).Columns("Comments").ColSpan = 3

            pnlCommitmentsTotals.Left = 435
        End If
        nAdjustmentBand = tmpBand

        If nCommitmentID > 0 Then
            For Each ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugCommitments.Rows
                If ugRow.Cells("CommitmentID").Value = nCommitmentID Then
                    ugRow.Activate()
                    ugRow.Selected = True
                    ugRow.Expanded = True
                    Exit For
                End If
            Next
        End If


        If IsNothing(dtTotals) Then
            lblAdjustmentValue.Text = ""
            lblBalanceValue.Text = ""
            lblCommitmentValue.Text = ""
            lblPaymentValue.Text = ""
        Else
            lblAdjustmentValue.Text = dtTotals.Rows(0)("EventAdjustmentTotal").ToString
            lblBalanceValue.Text = dtTotals.Rows(0)("EventBalanceTotal").ToString
            lblCommitmentValue.Text = dtTotals.Rows(0)("EventCommitmentTotal").ToString
            lblPaymentValue.Text = dtTotals.Rows(0)("EventPaymentTotal").ToString
            'strTotalPaidDocTag = lblPaymentValue.Text
        End If

        If Me.ugCommitments.Rows.Count > 0 AndAlso Me.ugCommitments.ActiveRow Is Nothing Then
            ugCommitments.ActiveRow = Me.ugCommitments.Rows(0)
            ugCommitments.Rows(0).Selected = True
        End If

    End Sub

    Private Sub ugCommitments_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugCommitments.DoubleClick
        'ugCommitments
        If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
            Me.Cursor = Cursors.Default
            Exit Sub
        End If

        Select Case ugCommitments.ActiveRow.Band.Index
            Case 0 ' Commitment
                CreateModifyCommitment(ugCommitments.ActiveRow.Cells("CommitmentID").Value)
            Case 1 ' Invoice or Adjustment
                If nAdjustmentBand = 2 Then
                    ' Invoice
                    'ModifyInvoice
                Else
                    ' Adjustment
                    CreateModifyAdjustment(ugCommitments.ActiveRow.Cells("CommitmentID").Value, ugCommitments.ActiveRow.Cells("ChildID").Value)
                End If
            Case 2 'Adjustment
                CreateModifyAdjustment(ugCommitments.ActiveRow.Cells("CommitmentID").Value, ugCommitments.ActiveRow.Cells("ChildID").Value)
        End Select


    End Sub

    Private Sub LoadPaymentsGrid()
        Dim dsLocal As DataSet
        Dim tmpBand As Int16
        Dim dtTotals As DataTable

        dtTotals = oFinancial.PaymentTotalsDatatable(0, chkShowCommitments.Checked)
        dsLocal = oFinancial.PaymentGridDataset(chkShowCommitments.Checked)
        'ugPayments.DataSource = dsLocal
        If dsLocal.Tables.Count > 0 Then
            dsLocal.Tables(0).DefaultView.Sort = "RECEIVED_DATE DESC"
            ugPayments.DataSource = dsLocal.Tables(0).DefaultView
        Else
            ugPayments.DataSource = dsLocal
        End If

        ugPayments.Rows.CollapseAll(True)
        ugPayments.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        ugPayments.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        If dsLocal.Tables(0).Rows.Count > 0 Then ' Does Payment Table have rows
            bolEnableDeleteEvent = False

            ugPayments.DisplayLayout.Bands(0).Override.MaxSelectedRows = 1

            ugPayments.DisplayLayout.Bands(0).Columns("FIN_EVENT_ID").Hidden = True
            ugPayments.DisplayLayout.Bands(0).Columns("Reimbursement_ID").Hidden = True
            ugPayments.DisplayLayout.Bands(0).Columns("COMMITMENT_ID").Hidden = True
            ugPayments.DisplayLayout.Bands(0).Columns("Document_Location").Hidden = True
            ugPayments.DisplayLayout.Bands(0).Columns("rawRequestedAmount").Hidden = True
            ugPayments.DisplayLayout.Bands(0).Columns("rawPaidAmount").Hidden = True

            ugPayments.DisplayLayout.Bands(0).Columns("Received_date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugPayments.DisplayLayout.Bands(0).Columns("Received_date").Header.Caption = "Received"
            ugPayments.DisplayLayout.Bands(0).Columns("Received_date").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            ugPayments.DisplayLayout.Bands(0).Columns("Requested_Amount").Header.Caption = "Requested"
            ugPayments.DisplayLayout.Bands(0).Columns("Requested_Amount").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            ugPayments.DisplayLayout.Bands(0).Columns("Requested_Invoiced").Header.Caption = "Requested(Inv)"
            ugPayments.DisplayLayout.Bands(0).Columns("Requested_Invoiced").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            ugPayments.DisplayLayout.Bands(0).Columns("Paid").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            ugPayments.DisplayLayout.Bands(0).Columns("Payment_Number").Header.Caption = "Pmt#"
            ugPayments.DisplayLayout.Bands(0).Columns("Payment_Number").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            ugPayments.DisplayLayout.Bands(0).Columns("Payment_Date").Header.Caption = "Pmt Date"
            ugPayments.DisplayLayout.Bands(0).Columns("Payment_Date").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugPayments.DisplayLayout.Bands(0).Columns("Payment_Date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            ugPayments.DisplayLayout.Bands(0).Columns("ApprovalRequired").Header.Caption = "App Reqd"
            ugPayments.DisplayLayout.Bands(0).Columns("ApprovalRequired").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            ugPayments.DisplayLayout.Bands(0).Columns("Approved").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            ugPayments.DisplayLayout.Bands(0).Columns("On_Hold").Header.Caption = "On Hold"
            ugPayments.DisplayLayout.Bands(0).Columns("On_Hold").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            ugPayments.DisplayLayout.Bands(0).Columns("Incomplete").Header.Caption = "Incomplete"
            ugPayments.DisplayLayout.Bands(0).Columns("Incomplete").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center


            ugPayments.DisplayLayout.Bands(0).Columns("Received_date").Width = 100
            ugPayments.DisplayLayout.Bands(0).Columns("Requested_Amount").Width = 100
            ugPayments.DisplayLayout.Bands(0).Columns("Requested_Invoiced").Width = 100
            ugPayments.DisplayLayout.Bands(0).Columns("Paid").Width = 100

        End If
        'tmpBand = 0
        'pnlCommitmentsTotals.Left = 440
        If dsLocal.Tables(1).Rows.Count > 0 Then ' Does Invoice Table have rows
            bolEnableDeleteEvent = False
            ugPayments.DisplayLayout.Bands(1).Override.MaxSelectedRows = 1

            ugPayments.DisplayLayout.Bands(1).Columns("Invoices_ID").Hidden = True
            ugPayments.DisplayLayout.Bands(1).Columns("COMMITMENTID").Hidden = True
            ugPayments.DisplayLayout.Bands(1).Columns("Fin_Event_ID").Hidden = True
            ugPayments.DisplayLayout.Bands(1).Columns("Reimbursement_ID").Hidden = True

            ugPayments.DisplayLayout.Bands(1).Columns("Vendor_Inv_Number").Header.Caption = "Invoice #"
            ugPayments.DisplayLayout.Bands(1).Columns("Vendor_Inv_Number").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left

            ugPayments.DisplayLayout.Bands(1).Columns("PONumber").Header.Caption = "PO #"
            ugPayments.DisplayLayout.Bands(1).Columns("PONumber").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left

            ugPayments.DisplayLayout.Bands(1).Columns("ActivityDescShort").Header.Caption = "Tasks"
            ugPayments.DisplayLayout.Bands(1).Columns("ActivityDescShort").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left

            ugPayments.DisplayLayout.Bands(1).Columns("On_Hold").Header.Caption = "On Hold"
            ugPayments.DisplayLayout.Bands(1).Columns("On_Hold").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            ugPayments.DisplayLayout.Bands(1).Columns("ApprovalRequired").Header.Caption = "App Reqd"
            ugPayments.DisplayLayout.Bands(1).Columns("ApprovalRequired").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            ugPayments.DisplayLayout.Bands(1).Columns("Approved").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            ugPayments.DisplayLayout.Bands(1).Columns("Final").Header.Caption = "Final"
            ugPayments.DisplayLayout.Bands(1).Columns("Final").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center

            ugPayments.DisplayLayout.Bands(1).Columns("Invoiced_Amount").Header.Caption = "Invoiced"
            ugPayments.DisplayLayout.Bands(1).Columns("Invoiced_Amount").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            ugPayments.DisplayLayout.Bands(1).Columns("Paid").Header.Caption = "Paid"
            ugPayments.DisplayLayout.Bands(1).Columns("Paid").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            ugPayments.DisplayLayout.Bands(1).Columns("Comment").Header.Caption = "Comment"
            ugPayments.DisplayLayout.Bands(1).Columns("Comment").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left



            '    ugPayments.DisplayLayout.Bands(tmpBand).Columns("Adjust_Date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            '    ugPayments.DisplayLayout.Bands(tmpBand).Columns("Adjust_Date").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            '    ugPayments.DisplayLayout.Bands(tmpBand).Columns("Adjust_Amount").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right


            '    ugPayments.DisplayLayout.Bands(tmpBand).Columns("Fin_App_Req").ColSpan = 2
            '    ugPayments.DisplayLayout.Bands(tmpBand).Columns("Comments").ColSpan = 3

            '    pnlCommitmentsTotals.Left = 465
        End If
        'nAdjustmentBand = tmpBand

        If nReimbursementID > 0 Then
            For Each ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugPayments.Rows
                If ugRow.Cells("Reimbursement_ID").Value = nReimbursementID Then
                    ugRow.Activate()
                    ugRow.Selected = True
                    ugRow.Expanded = True
                    Exit For
                End If
            Next
        End If

        If IsNothing(dtTotals) Then
            lblPaymentRequested.Text = ""
            lblpaymentInv.Text = ""
            lblPaid.Text = ""
        Else
            lblPaymentRequested.Text = dtTotals.Rows(0)("Requested_Amount").ToString
            lblpaymentInv.Text = dtTotals.Rows(0)("Requested_Invoiced").ToString
            lblPaid.Text = dtTotals.Rows(0)("Paid").ToString

        End If


    End Sub

    Private Sub pnlCommitmentsDetails_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles pnlCommitmentsDetails.Resize
        ugCommitments.Width = pnlCommitmentsDetails.Width - 50
    End Sub

    Private Sub pnlPaymentsDetails_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles pnlPaymentsDetails.Resize
        ugPayments.Width = pnlPaymentsDetails.Width - 50
    End Sub

    Private Sub btnDeleteCommitment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteCommitment.Click
        Dim MyFrm As MusterContainer
        'P1 02/26/06 next line only
        Dim dsLocal As DataSet
        Dim drRow As DataRow
        Dim oTechDox As New MUSTER.BusinessLogic.pLustEventDocument
        Try
            MyFrm = MdiParent
            If ugCommitments.ActiveRow.Band.Index = 0 Then
                If ugCommitments.ActiveRow.HasChild Then
                    MsgBox("Commitments with Adjustments or Invoices cannot be deleted.")
                    Exit Sub
                End If

                If ugCommitments.ActiveRow.Cells("PONumber").Value > "" Then
                    MsgBox("Commitments with PO Numbers cannot be deleted.")
                    Exit Sub
                End If
                If CDbl(lblBalanceValue.Text) - CDbl(ugCommitments.ActiveRow.Cells("Commitment").Value) < 0 Then
                    MsgBox("Total Balance Will Be Less Than 0.  Commitments cannot be deleted.")
                    Exit Sub
                End If
                If MsgBox("Are you sure you want to delete this Commitment?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    oCommitment.Retrieve(ugCommitments.ActiveRow.Cells("CommitmentID").Value)
                    oCommitment.Deleted = True
                    oCommitment.ModifiedBy = MusterContainer.AppUser.ID
                    oCommitment.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If

                    nCommitmentID = 0

                    LoadCommitmentsGrid()
                    dsLocal = oFinancial.PopulateCommitmentTecDocList(oFinancial.TecEventID, oCommitment.CommitmentID, False, "MODIFY")
                    If dsLocal.Tables.Count > 0 Then
                        If dsLocal.Tables(0).Rows.Count > 0 Then
                            For Each drRow In dsLocal.Tables(0).Rows
                                oTechDox.Retrieve(drRow.Item("Event_Activity_Document_ID"))
                                oTechDox.DueDate = "01/01/0001"
                                oTechDox.CommitmentID = 0
                                If oTechDox.ID <= 0 Then
                                    oTechDox.CreatedBy = MusterContainer.AppUser.ID
                                Else
                                    oTechDox.ModifiedBy = MusterContainer.AppUser.ID
                                End If
                                oTechDox.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)
                            Next
                        End If
                    End If

                    MyFrm.RefreshCalendarInfo()
                    MyFrm.LoadDueToMeCalendar()
                    MyFrm.LoadToDoCalendar()
                    MsgBox("Commitment Deleted.")
                End If
            Else
                MsgBox("No Commitment Selected.")
            End If
        Catch ex As Exception

        End Try


    End Sub

    Private Sub ugPayments_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugPayments.DoubleClick
        If Me.Cursor Is Cursors.WaitCursor Then
            Exit Sub
        End If
        If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        If ugPayments.ActiveRow.Band.Index = 0 Then
            ModifyUpdateReimbursmentRequest()
        Else
            CreateModifyInvoice(ugPayments.ActiveRow.Cells("Invoices_ID").Value)
        End If
    End Sub

    Private Sub btnDeleteRequest_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteRequest.Click
        Dim oReimbursement As New MUSTER.BusinessLogic.pFinancialReimbursement

        If ugPayments.ActiveRow.Band.Index = 0 Then
            'oReimbursement.Retrieve(ugPayments.ActiveRow.Cells("Reimbursement_ID").Value)
            If Not ugPayments.ActiveRow.ChildBands Is Nothing Then
                If Not ugPayments.ActiveRow.ChildBands(0).Rows.Count = 0 Then
                    MsgBox("This Reimbursement Request has an Invoice associated with it, and therefore cannot be deleted.")
                    Exit Sub
                End If
            End If

            'If Not ugPayments.ActiveRow.ChildBands(0).Rows.Count = 0 Then
            '    MsgBox("This Reimbursement Request has an Invoice associated with it, and therefore cannot be deleted.")
            '    'If oReimbursement.CommitmentID > 0 Then
            '    'MsgBox("This Reimbursement Request has an Invoice associated with it, and therefore cannot be deleted.")
            'Else
            If MsgBox("Are you sure you want to delete this Reimbursement Request?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                oReimbursement.Retrieve(ugPayments.ActiveRow.Cells("Reimbursement_ID").Value)
                oReimbursement.Deleted = True
                oReimbursement.ModifiedBy = MusterContainer.AppUser.ID
                oReimbursement.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If
                nReimbursementID = 0
                LoadPaymentsGrid()
            End If
            'End If
        End If
    End Sub

    Private Sub btnDeleteInvoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteInvoice.Click
        Dim oInvoice As New MUSTER.BusinessLogic.pFinancialInvoice

        If ugPayments.ActiveRow.Band.Index <> 0 Then
            oInvoice.Retrieve(ugPayments.ActiveRow.Cells("Invoices_ID").Value)
            If oInvoice.PaidAmount > 0 Then
                MsgBox("This Invoice has been paid, and therefore cannot be deleted.")
            Else
                If MsgBox("Are you sure you want to delete this Invoice?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                    oInvoice.Deleted = True
                    oInvoice.ModifiedBy = MusterContainer.AppUser.ID
                    oInvoice.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If

                    nReimbursementID = ugPayments.ActiveRow.ParentRow.Cells("Reimbursement_ID").Value

                    LoadPaymentsGrid()
                    LoadCommitmentsGrid()
                    MsgBox("Invoice Deleted", MsgBoxStyle.OKOnly, "Invoice Deleted")
                End If
            End If
        End If
    End Sub

    'Private Sub cmbVendor_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbVendor.SelectedIndexChanged
    '    LoadVendorAddress()

    'End Sub

    Private Sub LoadVendorAddress(Optional ByVal ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow = Nothing, Optional ByVal dtCompanyAddress As DataTable = Nothing)
        'Dim uRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim oContact As New MUSTER.BusinessLogic.pContactDatum
        Dim Address As String = String.Empty
        Dim strVendorEnvelopeAddress As String

        'If bolLoading Then Exit Sub
        Try
            'oContact.Retrieve(cmbVendor.SelectedValue)

            'Me.txtVendorAddress.ReadOnly = True
            'Address = oContact.AddressLine1
            'If oContact.AddressLine2 > String.Empty Then
            '    Address &= vbCrLf & oContact.AddressLine2
            'End If
            'Address &= vbCrLf & oContact.City & ", " & oContact.State & " " & oContact.ZipCode
            'txtVendorAddress.Text = Address

            If Not ugRow Is Nothing Then
                Address = ugRow.Cells("Address_one").Value + vbCrLf + IIf(ugRow.Cells("Address_two").Text = String.Empty, "", ugRow.Cells("Address_two").Value + vbCrLf) + IIf(ugRow.Cells("City").Text = String.Empty, "", ugRow.Cells("City").Value + ", ") + IIf(ugRow.Cells("State").Text = String.Empty, "", ugRow.Cells("State").Value + " ") + ugRow.Cells("Zip").Value
                strVendorEnvelopeAddress = ugRow.Cells("Address_one").Value + "," + IIf(ugRow.Cells("Address_two").Text = String.Empty, "", ugRow.Cells("Address_two").Value) + "," + IIf(ugRow.Cells("City").Text = String.Empty, "", ugRow.Cells("City").Value) + "," + IIf(ugRow.Cells("State").Text = String.Empty, "", ugRow.Cells("State").Value) + "," + ugRow.Cells("Zip").Value
                If ugRow.Cells("Address_one").Value Is DBNull.Value Then
                    arrVendorAddress(0) = ""
                Else
                    arrVendorAddress(0) = ugRow.Cells("Address_one").Value
                End If

                arrVendorAddress(1) = IIf(ugRow.Cells("Address_two").Text = String.Empty, "", ugRow.Cells("Address_two").Value)
                arrVendorAddress(2) = IIf(ugRow.Cells("City").Text = String.Empty, "", ugRow.Cells("City").Value)
                arrVendorAddress(3) = IIf(ugRow.Cells("State").Text = String.Empty, "", ugRow.Cells("State").Value)
                arrVendorAddress(4) = IIf(ugRow.Cells("Zip").Text = String.Empty, "", ugRow.Cells("Zip").Value)
            ElseIf dtCompanyAddress.Rows.Count > 0 Then
                Address = IIf(dtCompanyAddress.Rows(0).Item("Address_one") Is DBNull.Value, "", dtCompanyAddress.Rows(0).Item("Address_one")) + vbCrLf + IIf(dtCompanyAddress.Rows(0).Item("Address_two") Is DBNull.Value, "", dtCompanyAddress.Rows(0).Item("Address_two") + vbCrLf) + IIf(dtCompanyAddress.Rows(0).Item("City") = String.Empty, "", dtCompanyAddress.Rows(0).Item("City") + ", ") + IIf(dtCompanyAddress.Rows(0).Item("State") = String.Empty, "", dtCompanyAddress.Rows(0).Item("State") + " ") + dtCompanyAddress.Rows(0).Item("ZipCode")
                strVendorEnvelopeAddress = IIf(dtCompanyAddress.Rows(0).Item("Address_one") Is DBNull.Value, "", dtCompanyAddress.Rows(0).Item("Address_one")) + "," + IIf(dtCompanyAddress.Rows(0).Item("Address_two") Is DBNull.Value, "", dtCompanyAddress.Rows(0).Item("Address_two")) + "," + IIf(dtCompanyAddress.Rows(0).Item("City") = String.Empty, "", dtCompanyAddress.Rows(0).Item("City")) + "," + IIf(dtCompanyAddress.Rows(0).Item("State") = String.Empty, "", dtCompanyAddress.Rows(0).Item("State")) + "," + dtCompanyAddress.Rows(0).Item("ZipCode")

                arrVendorAddress(0) = dtCompanyAddress.Rows(0).Item("Address_one")
                arrVendorAddress(1) = IIf(dtCompanyAddress.Rows(0).Item("Address_two") Is DBNull.Value, "", dtCompanyAddress.Rows(0).Item("Address_two"))
                arrVendorAddress(2) = IIf(dtCompanyAddress.Rows(0).Item("City") Is DBNull.Value, "", dtCompanyAddress.Rows(0).Item("City"))
                arrVendorAddress(3) = IIf(dtCompanyAddress.Rows(0).Item("State") Is DBNull.Value, "", dtCompanyAddress.Rows(0).Item("State"))
                arrVendorAddress(4) = dtCompanyAddress.Rows(0).Item("ZipCode")
            Else
                For Each ugRow In ugFinancialContacts.Rows
                    If ugRow.Cells("LetterContactType").Value = 1185 Then
                        Address = ugRow.Cells("Address_one").Value + vbCrLf + IIf(ugRow.Cells("Address_two").Text = String.Empty, "", ugRow.Cells("Address_two").Value + vbCrLf) + IIf(ugRow.Cells("City").Text = String.Empty, "", ugRow.Cells("City").Value + ", ") + IIf(ugRow.Cells("State").Text = String.Empty, "", ugRow.Cells("State").Value + " ") + ugRow.Cells("Zip").Value
                        strVendorEnvelopeAddress = ugRow.Cells("Address_one").Value + "," + IIf(ugRow.Cells("Address_two").Text = String.Empty, "", ugRow.Cells("Address_two").Value) + "," + IIf(ugRow.Cells("City").Text = String.Empty, "", ugRow.Cells("City").Value) + "," + IIf(ugRow.Cells("State").Text = String.Empty, "", ugRow.Cells("State").Value) + "," + ugRow.Cells("Zip").Value
                        arrVendorAddress(0) = ugRow.Cells("Address_one").Value
                        arrVendorAddress(1) = IIf(ugRow.Cells("Address_two").Text = String.Empty, "", ugRow.Cells("Address_two").Value)
                        arrVendorAddress(2) = IIf(ugRow.Cells("City").Text = String.Empty, "", ugRow.Cells("City").Value)
                        arrVendorAddress(3) = IIf(ugRow.Cells("State").Text = String.Empty, "", ugRow.Cells("State").Value)
                        arrVendorAddress(4) = ugRow.Cells("Zip").Value
                        oFinancial.VendorID = ugRow.Cells("ContactID").Value
                        Exit For
                    End If
                Next
            End If

            txtVendorAddress.Text = Address
            strVendorEnvelopLabelAddress = strVendorEnvelopeAddress
            If txtVendor.Text = String.Empty Then
                oFinancial.VendorID = 0
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub btnGenerateApprovalForm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGenerateApprovalForm.Click
        Dim oCommitment As New MUSTER.BusinessLogic.pFinancialCommitment
        Dim oLetter As New Reg_Letters
        Dim strLongName As String
        Dim strShortName As String
        Dim strTemplate As String

        Try
            'Generating an Approval Form
            '1.	The user generates the following documents:
            '            a.Cover(Letter)
            '            b.Approval(Form)
            '            c.Commitment(Memorandum)
            '   for a Commitment by selecting an existing Commitment and indicating his desire to generate an Approval Form.
            '2.	The system responds by generating the documents as defined in the Global DDD, Appendix B, Letters and Reports, Letter Mechanism section.  
            '   Special(considerations)
            '   a.	If the Activity is Remediation System Operation, Maintenance, and Monitoring, the Cover Letter format will be that of the RSOM Letter.   
            '       If the Activity is Remediation System Operation, Maintenance, and Monitoring 2 through 7, the Cover Letter’s format will be that of the RSOM 2-7 Letter.  
            '       For all other Activities the Cover Letter’s format will be that of the standard Cover Letter.
            '   b.	If the selected Commitment’s Activity is Immediate Response Action Contractor Services, the Approval Form’s format will be that of the IRAC Letter.  
            '       For all other Activities the Approval Form’s format will be that of the standard Approval form.  
            '   c.	The Standard Form will display the same dollar amounts and quantity information (cost formats) described in the Happy Scenario 2a..  
            '   d.	If the selected Commitment’s Activity is Remediation Services - Excavation and Disposal of Contaminated Soils, the Approval Form will contain 
            '       'Remediation Contractor: <<Contractor Name>>’ above the cost format section.

            If ugCommitments.ActiveRow.Band.Index = 0 Then
                If Not IsDBNull(ugCommitments.ActiveRow.Cells("Document_Location").Value) Then
                    If MsgBox("An Approval Has Already Been Generated For The Selected Commitment.  Do You Want To Create A New Approval Form?", MsgBoxStyle.YesNo, "Reprint Approval Form?") = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If


                If ugCommitments.ActiveRow.Band.Index = 0 Then
                    'If Not IsDBNull(ugCommitments.ActiveRow.Cells("Document_Location").Value) Then
                    '    If MsgBox("An Approval Has Already Been Generated For The Selected Commitment.  Do You Want To Create A New Approval Form?", MsgBoxStyle.YesNo, "Reprint Approval Form?") = MsgBoxResult.No Then
                    '        Exit Sub
                    '    End If
                    nCommitmentID = ugCommitments.ActiveRow.Cells("CommitmentID").Value
                    oCommitment.Retrieve(ugCommitments.ActiveRow.Cells("CommitmentID").Value)

                    Dim act As New BusinessLogic.pFinancialActivity
                    act.Retrieve(oCommitment.ActivityType)

                    Dim contactID As Integer = 0

                    If oCommitment.ReimburseERAC AndAlso TypeOf ugFinancialContacts.DataSource Is DataSet AndAlso DirectCast(ugFinancialContacts.DataSource, DataSet).Tables.Count > 0 Then
                        Dim dr As DataRow() = DirectCast(ugFinancialContacts.DataSource, DataSet).Tables(0).Select("Type='Engineer/ERAC Representative'")

                        If dr.GetUpperBound(0) >= 0 AndAlso dr(0).Item("EntityAssocID") > 0 Then
                            contactID = dr(0).Item("ContactID")
                        Else
                            For Each row As DataRow In dr
                                If row.Item("EntityAssocID") > 0 Then
                                    contactID = dr(0).Item("ContactID")
                                End If
                            Next

                        End If

                    End If

                    ' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
                    'Cover Letter

                    Dim standardCoverLetter As String = "FundingApprovalCoverLetterTemplate.doc"
                    Dim standardNoticeLetter As String = "MGPTFApprovalFormTemplateNotRSOM.doc"

                    'set template
                    If act.CoverTemplate = String.Empty Then

                        strTemplate = standardCoverLetter
                    Else

                        strTemplate = act.CoverTemplate
                    End If


                    'For now, use to set long and short name
                    If oCommitment.ActivityType = 34 Then 'the Cover Letter format will be that of the RSOM Letter.  
                        strLongName = "Financial Approval RSOM Cover Letter"
                        strShortName = "RSOMLtr"
                        '  strTemplate = "RSOMYr1CoverLetterTemplate.doc"

                        'by: thomas Franey: added Activity Type 50 to condition on line  6029
                    ElseIf oCommitment.ActivityType = 35 _
                    Or oCommitment.ActivityType = 36 _
                    Or oCommitment.ActivityType = 37 _
                    Or oCommitment.ActivityType = 38 _
                    Or oCommitment.ActivityType = 39 _
                    Or oCommitment.ActivityType = 50 _
                    Or oCommitment.ActivityType = 40 Then                    'the Cover Letter’s format will be that of the RSOM 2-7 Letter
                        strLongName = "Financial Approval RSOM 2-7 Cover Letter"
                        strShortName = "RSOM2-7Ltr"
                        ' strTemplate = "RSOMYr2CoverLetterTemplate.doc"

                    Else 'the Cover Letter’s format will be that of the standard Cover Letter.
                        strLongName = "Financial Approval Cover Letter"
                        strShortName = "AppCoverLtr"
                        ' strTemplate = "FundingApprovalCoverLetterTemplate.doc"

                    End If

                    oLetter.GenerateFinancialLetter(lblFacilityIDValue.Text, strLongName, strShortName, strLongName, strTemplate, pOwn, oFinancial.TecEventID, oCommitment.Fin_Event_ID, oCommitment.CommitmentID, 0, 0, 0, lblOwnerIDValue.Text, oFinancial.ID, oFinancial.Sequence, UIUtilsGen.EntityTypes.FinancialEvent, contactID)
                    UIUtilsGen.Delay(, 2)

                    ' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
                    'Approval Form



                    'set template
                    If act.NoticeTemplate = String.Empty Then

                        strTemplate = standardNoticeLetter
                    Else

                        strTemplate = act.NoticeTemplate
                    End If



                    'set long and short na me for notices   
                    If oCommitment.ActivityType = 19 Then 'the Approval Form’s format will be that of the IRAC Letter. 
                        strLongName = "Financial IRAC Letter"
                        strShortName = "IRACLtr"
                        'strTemplate = "IRACLetterTemplate.doc"
                    Else 'the Approval Form’s format will be that of the standard Approval form.
                        strLongName = "Financial Approval Form"
                        strShortName = "ApprovalForm"
                        'If oCommitment.ActivityType = 34 Then
                        '   strTemplate = "MGPTFApprovalFormTemplateRSOMYr1.doc"
                        'ElseIf oCommitment.ActivityType = 35 _
                        '       Or oCommitment.ActivityType = 36 _
                        '      Or oCommitment.ActivityType = 37 _
                        '     Or oCommitment.ActivityType = 38 _
                        '    Or oCommitment.ActivityType = 39 _
                        '   Or oCommitment.ActivityType = 50 _
                        '  Or oCommitment.ActivityType = 40 Then
                        '  strTemplate = "MGPTFApprovalFormTemplateRSOMYr2-7.doc"
                        'Else
                        '  strTemplate = "MGPTFApprovalFormTemplateNotRSOM.doc"
                        ' End If
                    End If

                    oLetter.GenerateFinancialLetter(lblFacilityIDValue.Text, strLongName, strShortName, strLongName, strTemplate, pOwn, oFinancial.TecEventID, oCommitment.Fin_Event_ID, oCommitment.CommitmentID, 0, 0, 0, lblOwnerIDValue.Text, oFinancial.ID, oFinancial.Sequence, UIUtilsGen.EntityTypes.FinancialEvent, contactID)
                    UIUtilsGen.Delay(, 2)




                    ' +-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
                    'Commitment Memo

                    strLongName = "Financial Commitment Memo"
                    strShortName = "CommitmentMemo"
                    strTemplate = "CommitmentMemoTemplate.doc"
                    oLetter.GenerateFinancialLetter(lblFacilityIDValue.Text, strLongName, strShortName, strLongName, strTemplate, pOwn, oFinancial.TecEventID, oCommitment.Fin_Event_ID, oCommitment.CommitmentID, 0, 0, 0, lblOwnerIDValue.Text, oFinancial.ID, oFinancial.Sequence, UIUtilsGen.EntityTypes.FinancialEvent, contactID)
                    UIUtilsGen.Delay(, 1)
                    'End If
                    LoadCommitmentsGrid()
                Else
                    MsgBox("Please Select A Commitment From The Grid.")
                End If

            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnGenerateNoticeOfReim_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGenerateNoticeOfReim.Click
        Dim Childrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim oReimbursement As New MUSTER.BusinessLogic.pFinancialReimbursement
        Dim oCommittment As New MUSTER.BusinessLogic.pFinancialCommitment
        Dim nPaymentNumber As Int16
        Dim bolNewPaymentNumber As Boolean
        Dim oLetter As New Reg_Letters
        Dim strLongName As String
        Dim strShortName As String
        Dim strTemplate As String
        Dim MyFrm As MusterContainer

        Try
            If Not (ugPayments.Rows.Count > 0) Then
                Exit Sub
            End If


            'Generating a Notice of Reimbursement
            '   1.	The user generates a Notice of Reimbursement by selecting an existing Reimbursement Request and initiating an action to generate the Notice of Reimbursement.
            '   2.	The system responds by:
            '       a.	Disallowing the generation if any of the following conditions are True:
            '           i.	    The Reimbursement’s Requested Amount does not equal the sum total of the Invoiced Amount of all of the Reimbursement Request’s Invoices.
            '           ii.	    The Reimbursement Request has one or more Invoices that require approval and are not approved.  Requires approval and approved are defined as:
            '                   1.	Requires Approval: Invoice Final = True and Invoice contains one or more Technical Reports which contains a Due Date.
            '                   2.	Approved: Invoice Requires Approval = True and all of Invoice’s Technical Reports containing Due Dates also contain Approved dates.
            '           iii.	The Reimbursement Request has one or more Invoices whose corresponding Commitment does not contain a PO#.
            '           iv. 	The Reimbursement Request has one or more Invoices that are On Hold.
            '           v.	    The Reimbursement Request has one or more Invoices whose corresponding Commitment has one or more unapproved Change Order Adjustment.  
            '                   An unapproved Change Order Adjustment is an Adjustment with Type = Change Order that requires approval (Director’s Approval Required = True 
            '                   or Financial Approval Required = True) but Approved = False.
            '           vi.	    The Notice of Reimbursement contains Incomplete Application Reasons.  If this is true, the Incomplete Application Reasons UI will be displayed.
            '           vii.	The Reimbursement Request has one or more Invoices whose paid amount is null.  Note: Amount can be 0 but not null.
            '       b.	If allowed per above and a Notice of Reimbursement has not already been generated for the selected Reimbursement Request, generating a Payment Number.  
            '           The Payment Number will be a sequential number starting at 1 and will be unique to the Financial Event.  
            '       c.	If allowed per above and a Notice of Reimbursement has not already been generated for the selected Reimbursement Request, generating the Reimbursement 
            '           Memo (format dependent upon the Reimbursement Request’s Invoice’s Commitment Funding Type) and the Notice of Reimbursement.  
            '       d.	If a Notice of Reimbursement has already been generated for the selected Reimbursement Request, allowing the user to regenerate a Reimbursement Memo 
            '           and Notice of Reimbursement.  A new Payment Number will not be generated.  
            '       e.	If allowed per above, setting the Closed Date of all Documents in the Technical module corresponding to Technical Reports marked Paid in the 
            '           Reimbursement(Request) 's Invoice(s) to the current date.  Also, closing the activity (setting the Activity Closed Date = current date) that is 
            '           associated to the documents closed if activity has a technical completed date and the associated activity has no open documents.  Additionally if an 
            '           activity closed in this step is “Plugging Monitoring Well”, a To Do Calendar Entry will be created for the Technical Event’s Project Manager on the 
            '           current date to remind the PM to close the Technical Event.



            If ugPayments.ActiveRow.Band.Index = 0 Then

                '+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
                '2.a above
                If IsDBNull(ugPayments.ActiveRow.Cells("Requested_Invoiced").Value) Then
                    MsgBox("Invoiced Amount must be greater than zero.")
                    Exit Sub
                End If
                If IsDBNull(ugPayments.ActiveRow.Cells("Requested_Amount").Value) Then
                    MsgBox("Requested Amount must be greater than zero.")
                    Exit Sub
                End If
                If ugPayments.ActiveRow.Cells("Requested_Amount").Value <> ugPayments.ActiveRow.Cells("Requested_Invoiced").Value Then ' 2.a.i above
                    MsgBox("Reimbursement`s Requested Amount does not equal the sum total of the Invoiced Amount")
                    Exit Sub
                End If

                If ugPayments.ActiveRow.Cells("ApprovalRequired").Value = True And ugPayments.ActiveRow.Cells("Approved").Value = False Then ' 2.a.ii above
                    MsgBox("Reimbursement Request Has One Or More Invoices That Require Approval Which Are Not Approved")
                    Exit Sub
                End If

                For Each Childrow In ugPayments.ActiveRow.ChildBands(0).Rows '2.a.iii above
                    If Not IsDBNull(Childrow.Cells("PONumber").Value) Then
                        If Not (Childrow.Cells("PONumber").Value > String.Empty) Then
                            MsgBox("Reimbursement Request Has One Or More Invoices Whose Corresponding Commitment Does Not Contain A PO#.")
                            Exit Sub
                        End If
                    End If
                Next

                For Each Childrow In ugPayments.ActiveRow.ChildBands(0).Rows '2.a.iv above
                    If Childrow.Cells("On_Hold").Value = True Then
                        MsgBox("Reimbursement Request Has One Or More Invoices That Are On Hold")
                        Exit Sub
                    End If
                Next

                oReimbursement.Retrieve(ugPayments.ActiveRow.Cells("reimbursement_ID").Value)

                oCommittment.Retrieve(oReimbursement.CommitmentID)

                For Each ugrow In ugCommitments.Rows    '2.a.v above
                    If ugrow.Cells("CommitmentID").Value = oReimbursement.CommitmentID Then
                        If ugrow.ChildBands.Count > 1 Then
                            For Each Childrow In ugrow.ChildBands(1).Rows
                                If Childrow.Cells("Adjust_Type").Text.ToUpper = ("change order").ToUpper Then
                                    If (Childrow.Cells("Director_App_Req").Value = True Or Childrow.Cells("Fin_App_Req").Value = True) And Childrow.Cells("Approved").Value <> True Then
                                        MsgBox("The Reimbursement Request`s Corresponding Commitment Has One Or More Unapproved Change Order Adjustment(s)")
                                        Exit Sub
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next

                oReimbursement.Retrieve(ugPayments.ActiveRow.Cells("Reimbursement_ID").Value) '2.a.vi above
                If oReimbursement.Incomplete Then
                    MsgBox("The Notice of Reimbursement Contains Incomplete Application Reasons.")
                    ModifyUpdateReimbursmentRequest()
                    Exit Sub
                End If


                '+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
                '2.b above

                nPaymentNumber = oReimbursement.PaymentNumber

                If nPaymentNumber = 0 Then
                    '2.a.v above
                    ' #2758 - need to get payment numbers from rows not shown on grid if show open commitment checkbox is checked
                    'For Each ugrow In ugPayments.Rows
                    '    If ugrow.Cells("Payment_Number").Value > nPaymentNumber Then
                    '        nPaymentNumber = ugrow.Cells("Payment_Number").Value
                    '    End If
                    'Next
                    nPaymentNumber = oFinancial.GetMaxPaymentNumber(oFinancial.ID)
                    'added on April 10, 2008 by Hua Cao - Issue Tracker #3140; If payment amount is 0, then payment number should be 0.
                    If ugPayments.ActiveRow.Cells("Paid").Value > 0 Then
                        nPaymentNumber += 1
                    Else
                        nPaymentNumber = 0
                    End If
                    oReimbursement.PaymentNumber = nPaymentNumber
                    oReimbursement.PaymentDate = Now.Date
                    If oReimbursement.id <= 0 Then
                        oReimbursement.CreatedBy = MusterContainer.AppUser.ID
                    Else
                        oReimbursement.ModifiedBy = MusterContainer.AppUser.ID
                    End If
                    oReimbursement.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If

                    LoadPaymentsGrid()
                    LoadCommitmentsGrid()
                    bolNewPaymentNumber = True
                Else
                    bolNewPaymentNumber = False
                End If

                '+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
                '2.c & 2.d above

                strLongName = "Reimbursement Memo"

                strShortName = "ReimbursementMemo"


                Select Case oCommittment.FundingType
                    Case 1077       'FED
                        strTemplate = "FEDReimbursementMemoTemplate.doc"
                    Case 1655    'ARRA
                        strTemplate = "ARRAReimbursementMemoTemplate.doc"
                    Case 1078   'SDN 
                        strTemplate = "SDNReimbursementMemoTemplate.doc"
                    Case 1079   'SDR
                        strTemplate = "SDRReimbursementMemoTemplate.doc"
                    Case 1574         'KATRINA
                        strTemplate = "KATReimbursementMemoTemplate.doc"
                    Case Else          'STFS
                        strTemplate = "STFSReimbursementMemoTemplate.doc"


                End Select

                Dim contactID As Integer = 0

                If oCommittment.ReimburseERAC AndAlso TypeOf ugFinancialContacts.DataSource Is DataSet AndAlso DirectCast(ugFinancialContacts.DataSource, DataSet).Tables.Count > 0 Then
                    Dim dr As DataRow() = DirectCast(ugFinancialContacts.DataSource, DataSet).Tables(0).Select("Type='Engineer/ERAC Representative'")

                    If dr.GetUpperBound(0) >= 0 AndAlso dr(0).Item("EntityAssocID") > 0 Then
                        contactID = dr(0).Item("ContactID")
                    Else
                        For Each row As DataRow In dr
                            If row.Item("EntityAssocID") > 0 Then
                                contactID = dr(0).Item("ContactID")
                            End If
                        Next


                    End If

                End If


                oLetter.GenerateFinancialLetter(lblFacilityIDValue.Text, strLongName, strShortName, strLongName, strTemplate, pOwn, oFinancial.TecEventID, oReimbursement.FinancialEventID, oReimbursement.CommitmentID, oReimbursement.id, 0, 0, lblOwnerIDValue.Text, oFinancial.ID, oFinancial.Sequence, UIUtilsGen.EntityTypes.FinancialEvent, contactID)
                UIUtilsGen.Delay(, 2)

                strLongName = "Notice of Reimbursement"
                strShortName = "ReimbursementNotice"
                strTemplate = "NoticeofReimbursementTemplate.doc"

                oLetter.GenerateFinancialLetter(lblFacilityIDValue.Text, strLongName, strShortName, strLongName, strTemplate, pOwn, oFinancial.TecEventID, oReimbursement.FinancialEventID, oReimbursement.CommitmentID, oReimbursement.id, 0, 0, lblOwnerIDValue.Text, oFinancial.ID, oFinancial.Sequence, UIUtilsGen.EntityTypes.FinancialEvent, contactID)
                UIUtilsGen.Delay(, 1)
                '+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-
                '2.e above

                oReimbursement.ProcessReimbursementNotification(oReimbursement.id, CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                MyFrm = MdiParent
                MyFrm.RefreshCalendarInfo()
                MyFrm.LoadDueToMeCalendar()
                MyFrm.LoadToDoCalendar()
            Else
                MsgBox("Please Select A Payment Request From The Grid.")
                Exit Sub
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnViewApprovalForm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnViewApprovalForm.Click
        Dim dtApprovalLetters As DataTable
        Dim drRow As DataRow
        Try
            If ugCommitments.ActiveRow.Band.Index = 0 Then
                If Not IsDBNull(ugCommitments.ActiveRow.Cells("Document_Location").Value) Then
                    nCommitmentID = ugCommitments.ActiveRow.Cells("CommitmentID").Value
                    dtApprovalLetters = oCommitment.GetApprovalDocuments(Integer.Parse(ugCommitments.ActiveRow.Cells("CommitmentID").Value))
                    If dtApprovalLetters.Rows.Count > 0 Then
                        For Each drRow In dtApprovalLetters.Rows
                            LetterGenerator.ViewDocument(drRow.Item("DocName"))
                        Next
                    End If
                Else
                    MsgBox("There Is No Approval Form Associated With The Selected Commitment.")
                End If

            Else
                MsgBox("Please Select A Commitment From The Grid.")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Sub btnViewNoticeOfReim_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnViewNoticeOfReim.Click
        Dim dtNoticeLetters As DataTable
        Dim drRow As DataRow
        Dim oReimbursement As New MUSTER.BusinessLogic.pFinancialReimbursement
        Try
            If ugPayments.ActiveRow.Band.Index = 0 Then
                If Not IsDBNull(ugPayments.ActiveRow.Cells("Document_Location").Value) Then
                    dtNoticeLetters = oReimbursement.GetNoticeDocuments(Integer.Parse(ugPayments.ActiveRow.Cells("Reimbursement_ID").Value))
                    'ViewDocument(ugPayments.ActiveRow.Cells("Document_Location").Value)
                    If dtNoticeLetters.Rows.Count > 0 Then
                        For Each drRow In dtNoticeLetters.Rows
                            LetterGenerator.ViewDocument(drRow.Item("DocName"))
                        Next
                    End If
                Else
                    MsgBox("There Is No Reimbursement Notice Associated With The Selected Reimbursement Request.")
                End If
            Else
                MsgBox("Please Select A Payment Request From The Grid.")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub



    Private Sub btnVendorPack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnVendorPack.Click
        Dim lclDOC_PATH As String = MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_Templates).ProfileValue & "\"

        If lclDOC_PATH = "\" Then
            MsgBox("Document Path Unspecified. Please give the path before generating the letter.")
        End If

        LetterGenerator.ViewDocument(lclDOC_PATH & "Financial\" & "VendorPackTemplate.doc")

    End Sub

    Private Sub btnFinancialComments_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFinancialComments.Click
        CommentsMaintenance(sender, e)
    End Sub


    Private Sub btnCommitmentsExpand_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCommitmentsExpand.Click
        ugCommitments.Rows.ExpandAll(True)
    End Sub

    Private Sub btnCommitmentsCollapse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCommitmentsCollapse.Click
        ugCommitments.Rows.CollapseAll(True)
    End Sub

    Private Sub btnPaymentExpand_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPaymentExpand.Click
        ugPayments.Rows.ExpandAll(True)
    End Sub

    Private Sub btnCollapse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCollapse.Click
        ugPayments.Rows.CollapseAll(True)
    End Sub


    Private Sub btnDeleteEvent_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteEvent.Click

        Try
            If MsgBox("Do you wish to delete this Financial Event?", MsgBoxStyle.YesNo, "Confirm Delete") = MsgBoxResult.Yes Then
                oFinancial.Deleted = True
                oFinancial.ModifiedBy = MusterContainer.AppUser.ID
                oFinancial.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If
                MsgBox("Financial Event Successfully Deleted")
                GetFinancialEventsForFacility()
                tbPageFacilityDetail.Enabled = True
                tbCntrlFinancial.SelectedTab = tbPageFacilityDetail
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Sub chkShowCommitments_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowCommitments.CheckedChanged

        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim ugPayment As Infragistics.Win.UltraWinGrid.UltraGridRow

        Try
            'If chkShowCommitments.Checked Then
            '    For Each ugPayment In ugPayments.Rows
            '        ugPayment.Hidden = False
            '    Next
            '    For Each ugrow In ugCommitments.Rows    '2.a.v above
            '        'bolExit = False
            '        If ugrow.Cells("Balance").Value = 0 Then
            '            ugrow.Hidden = True

            '            For Each ugPayment In ugPayments.Rows
            '                If Not ugPayment.Cells("commitment_ID").Value Is System.DBNull.Value And ugPayment.Cells("commitment_ID").Value = ugrow.Cells("commitmentID").Value Then
            '                    If Not ugPayment.Cells("Paid").Value Is System.DBNull.Value Then
            '                        If ((ugPayment.Cells("Paid").Value = ugrow.Cells("Commitment").Value)) Or ((CDec(ugPayment.Cells("Paid").Value) - CDec(ugrow.Cells("Adjustment").Value)) = CDec(ugrow.Cells("Commitment").Value)) Then
            '                            ugPayment.Hidden = True
            '                            If lblPaymentRequested.Text <> String.Empty Then
            '                                lblPaymentRequested.Text = CDec(lblPaymentRequested.Text) - CDec(ugPayment.Cells("Requested_Amount").Value)
            '                            End If
            '                            If lblpaymentInv.Text <> String.Empty Then
            '                                lblpaymentInv.Text = CDec(lblpaymentInv.Text) - CDec(ugPayment.Cells("Requested_Invoiced").Value)
            '                            End If
            '                            If lblPaid.Text <> String.Empty Then
            '                                lblPaid.Text = CDec(lblPaid.Text) - CDec(ugPayment.Cells("Paid").Value)
            '                            End If

            '                        End If
            '                    End If
            '                    Exit For
            '                End If
            '            Next
            '            'If Not ugrow.ChildBands Is Nothing Then
            '            '    For Each ChildBand In ugrow.ChildBands
            '            '        If Not ChildBand.Rows Is Nothing Then
            '            '            If ChildBand.Rows.Count > 0 Then

            '            '                For Each Childrow In ChildBand.Rows
            '            '                    If Not Childrow.Cells("commitmentID") Is System.DBNull.Value And Childrow.Cells("commitmentID").Value = ugrow.Cells("commitmentID").Value Then
            '            '                        If Not ugPayment.Cells("Paid").Value Is System.DBNull.Value Then
            '            '                            If (ugPayment.Cells("Paid").Value = ugrow.Cells("Commitment").Value) Then
            '            '                                ugPayment.Hidden = True

            '            '                            End If
            '            '                        End If
            '            '                        bolExit = True
            '            '                        Exit For
            '            '                    End If
            '            '                Next
            '            '            End If
            '            '        End If
            '            '        If bolExit Then
            '            '            Exit For
            '            '        End If
            '            '    Next
            '            'End If
            '            'If bolExit Then
            '            '    Exit For
            '            'End If
            '            If lblAdjustmentValue.Text <> String.Empty Then
            '                lblAdjustmentValue.Text = CDec(lblAdjustmentValue.Text) - CDec(ugrow.Cells("Adjustment").Value)
            '            End If
            '            If lblBalanceValue.Text <> String.Empty Then
            '                lblBalanceValue.Text = CDec(lblBalanceValue.Text) - CDec(ugrow.Cells("Balance").Value)
            '            End If
            '            If lblCommitmentValue.Text <> String.Empty Then
            '                lblCommitmentValue.Text = CDec(lblCommitmentValue.Text) - CDec(ugrow.Cells("Commitment").Value)
            '            End If
            '            If lblPaymentValue.Text <> String.Empty Then
            '                lblPaymentValue.Text = CDec(lblPaymentValue.Text) - CDec(ugrow.Cells("payment").Value)
            '            End If

            '        Else
            '            ugrow.Hidden = False
            '        End If
            '    Next
            'Else
            '    LoadCommitmentsGrid()
            '    LoadPaymentsGrid()
            'End If
            LoadCommitmentsGrid()
            LoadPaymentsGrid()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub tbCtrlFacFinancialEvt_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbCtrlFacFinancialEvt.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            Select Case tbCtrlFacFinancialEvt.SelectedTab.Name
                Case tbPageOwnerContactList.Name
                    If Me.lblFacilityIDValue.Text <> String.Empty Then
                        LoadContacts(ugOwnerContacts, Integer.Parse(Me.lblFacilityIDValue.Text), UIUtilsGen.EntityTypes.Facility)
                    End If
                Case tbPageFacilityDocuments.Name
                    UCFacilityDocuments.LoadDocumentsGrid(Integer.Parse(Me.lblFacilityIDValue.Text), UIUtilsGen.EntityTypes.Facility, UIUtilsGen.ModuleID.Financial)
            End Select
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugOwnerContacts_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugOwnerContacts.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
                ModifyContact(ugOwnerContacts)

            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugFinancialContacts_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugFinancialContacts.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            ModifyContact(ugFinancialContacts)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub


#Region "Envelopes and Labels"
    Private Sub btnFinEnvelopes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFinEnvelopes.Click
        Try
            If txtVendorAddress.Text <> String.Empty Then
                UIUtilsGen.CreateEnvelopes(txtVendor.Text, arrVendorAddress, "FIN", oFinancial.ID)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnFinLabels_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFinLabels.Click
        Try
            If txtVendorAddress.Text <> String.Empty Then
                UIUtilsGen.CreateLabels(txtVendor.Text, arrVendorAddress, "FIN", oFinancial.ID)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnFinFacLabels_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFinFacLabels.Click
        Dim strAddress As String
        Dim arrAddress(4) As String
        Try
            arrAddress(0) = pOwn.Facilities.FacilityAddresses.AddressLine1
            arrAddress(1) = pOwn.Facilities.FacilityAddresses.AddressLine2
            arrAddress(2) = pOwn.Facilities.FacilityAddresses.City
            arrAddress(3) = pOwn.Facilities.FacilityAddresses.State
            arrAddress(4) = pOwn.Facilities.FacilityAddresses.Zip
            'strAddress = pOwn.Facilities.FacilityAddresses.AddressLine1 + "," + pOwn.Facilities.FacilityAddresses.AddressLine2 + "," + pOwn.Facilities.FacilityAddresses.City + "," + pOwn.Facilities.FacilityAddresses.State + "," + pOwn.Facilities.FacilityAddresses.Zip
            If pOwn.Facilities.FacilityAddresses.AddressId > 0 And pOwn.ID > 0 Then
                UIUtilsGen.CreateLabels(Me.txtFacilityName.Text, arrAddress, "FIN", pOwn.Facilities.ID)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnFinOwnerLabels_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFinOwnerLabels.Click
        Dim strAddress As String
        Dim arrAddress(4) As String
        Try
            arrAddress(0) = pOwn.Addresses.AddressLine1
            arrAddress(1) = pOwn.Addresses.AddressLine2
            arrAddress(2) = pOwn.Addresses.City
            arrAddress(3) = pOwn.Addresses.State
            arrAddress(4) = pOwn.Addresses.Zip
            'strAddress = pOwn.Addresses.AddressLine1 + "," + pOwn.Addresses.AddressLine2 + "," + pOwn.Addresses.City + "," + pOwn.Addresses.State + "," + pOwn.Addresses.Zip
            If pOwn.Addresses.AddressId > 0 And pOwn.ID > 0 Then
                UIUtilsGen.CreateLabels(Me.txtOwnerName.Text, arrAddress, "FIN", pOwn.ID)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnFinOwnerEnvelopes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFinOwnerEnvelopes.Click
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

                If Not dsContactsLocal Is Nothing AndAlso dsContactsLocal.Tables.Count > 0 AndAlso dsContactsLocal.Tables(0).Rows.Count > 0 Then
                    For Each contactRow As DataRow In dsContactsLocal.Tables(0).Rows
                        If contactRow("Type") = "Registration Representative" Then
                            strContactName = contactRow("CONTACT_name")
                            Exit For
                        End If

                    Next
                Else

                    If Not pOwn.Persona.LastName.Length > 0 Then
                        strContactName = String.Format("{0}{1}{2}{3}", pOwn.Persona.FirstName, IIf(pOwn.Persona.MiddleName.Length > 0, " ", ""), pOwn.Persona.MiddleName, pOwn.Persona.LastName)
                    Else
                        strContactName = pOwn.Persona.Company
                    End If

                End If


                UIUtilsGen.CreateEnvelopes(Me.txtOwnerName.Text, arrAddress, "FIN", pOwn.ID, strContactName)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnFinFacEnvelopes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFinFacEnvelopes.Click
        Dim arrAddress(4) As String
        Try
            arrAddress(0) = pOwn.Facilities.FacilityAddresses.AddressLine1
            arrAddress(1) = pOwn.Facilities.FacilityAddresses.AddressLine2
            arrAddress(2) = pOwn.Facilities.FacilityAddresses.City
            arrAddress(3) = pOwn.Facilities.FacilityAddresses.State
            arrAddress(4) = pOwn.Facilities.FacilityAddresses.Zip
            'strAddress = pOwn.Facilities.FacilityAddresses.AddressLine1 + "," + pOwn.Facilities.FacilityAddresses.AddressLine2 + "," + pOwn.Facilities.FacilityAddresses.City + "," + pOwn.Facilities.FacilityAddresses.State + "," + pOwn.Facilities.FacilityAddresses.Zip
            If pOwn.Facilities.FacilityAddresses.AddressId > 0 And pOwn.ID > 0 Then
                UIUtilsGen.CreateEnvelopes(Me.txtFacilityName.Text, arrAddress, "FIN", pOwn.Facilities.ID)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

    Private Sub dtFinancialClosedDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtFinancialClosedDate.ValueChanged
        If bolLoading = True Then Exit Sub
        UIUtilsGen.ToggleDateFormat(dtFinancialClosedDate)
        oFinancial.ClosedDate = UIUtilsGen.GetDatePickerValue(dtFinancialClosedDate)
        oFinancial.ClosedDate = oFinancial.ClosedDate.Date
    End Sub

    Private Sub btnActPlanning_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPlanning.Click
        Dim evtActPlanning As New ActivityPlanning(616, Me.oTechnical.FacilityID, Me.pOwn.Facility.Name, oTechnical.ID, oTechnical.EVENTSEQUENCE)

        evtActPlanning.ShowDialog()

        evtActPlanning.Dispose()

    End Sub



    Private Sub txtFacilityAddress_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFacilityAddress.DoubleClick
        Try

            Dim addressForm As Address
            If (pOwn.Facilities.ID = 0 Or pOwn.Facilities.ID <> nFacilityID) And nFacilityID > 0 Then
                UIUtilsGen.PopulateFacilityInfo(Me, pOwn.OwnerInfo, pOwn.Facilities, nFacilityID)
                '  UIUtilsGen.PopulateFacilityInfo(Me, oOwner.OwnerInfo, oOwner.Facilities, strFacilityIdTags)
            End If
            Address.EditAddress(addressForm, pOwn.Facilities.ID, pOwn.Facilities.FacilityAddresses, "Facility", UIUtilsGen.ModuleID.Financial, txtFacilityAddress, UIUtilsGen.EntityTypes.Facility, True)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtOwnerAddress_DblClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOwnerAddress.DoubleClick
        Try

            Dim addressForm As Address
            Address.EditAddress(addressForm, pOwn.ID, pOwn.Addresses, "Owner", UIUtilsGen.ModuleID.Financial, txtOwnerAddress, UIUtilsGen.EntityTypes.Owner)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub

End Class
