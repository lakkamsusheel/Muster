Public Class Fees
    Inherits System.Windows.Forms.Form










    'changes made
    '''''     By TMF        2/18/2009                   Line 3232, 3233, 3239. 
    '                      Remarked lines 3232 and 3239 and inderted line to set active row to nothing to allow
    '                      system to keep facility id in scope







#Region "User Defined Variables"
    Private WithEvents pOwn As MUSTER.BusinessLogic.pOwner
    Private bolFormLoaded As Boolean = False
    Public strFacilityIdTags As String
    Public MyGuid As New System.Guid
    Private bolLoading As Boolean = False
    Public bolNewPersona As Boolean = False
    Private bolDisplayErrmessage As Boolean = True
    Private bolValidateSuccess As Boolean = True
    Private oAddressInfo As MUSTER.Info.AddressInfo
    Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long
    Private frmLateFeeWaiver As LateFeeWaiverRequest
    Private frmGenCreditMemo As GenerateCreditMemo
    Private frmReallocateOverage As ReallocateOwnerOverage
    Private frmGenDebitMemo As GenerateDebitMemo
    Private frmOverpayment As OverpaymentReason
    Private frmGenRefund As GenerateRefund
    Private WithEvents SF As ShowFlags
    Private bolFrmActivated As Boolean = False
    Private oManualInvoice As New MUSTER.BusinessLogic.pFeeInvoice
    Private dtManualInvoice As New System.Data.DataTable("Invoices")
    Private bolRowSelfUpdating As Boolean
    Dim oFeesBasis As New MUSTER.BusinessLogic.pFeeBasis
    Dim returnVal As String = String.Empty
    Dim ValidManualInvoice As Boolean
    Dim nFacilityID As Integer = 0

    Public mContainer As MusterContainer
    'contacts
    Private pConStruct As MUSTER.BusinessLogic.pContactStruct
    Private WithEvents oCompanySearch As CompanySearch
    Private WithEvents objCntSearch As ContactSearch
    Public nOwnerID As Integer
    Private dsContacts As DataSet
    Private strFilterString As String = String.Empty

#End Region
#Region " Windows Form Designer generated code "
    Public Sub New(ByRef oOwner As MUSTER.BusinessLogic.pOwner, ByVal OwnerID As Int64, ByVal FacilityID As Int64)
        MyBase.New()
        bolLoading = True
        MyGuid = System.Guid.NewGuid
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        pOwn = oOwner
        pConStruct = New MUSTER.BusinessLogic.pContactStruct
        'Add any initialization after the InitializeComponent() call
        MusterContainer.AppUser.LogEntry(Me.Text, MyGuid.ToString)
        MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "WindowName", Me.Text)
        MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "Fees")
        Try
            InitControls()
            If OwnerID > 0 Then
                PopulateOwnerInfo(Integer.Parse(OwnerID))
                If FacilityID <= 0 Then
                    tbCntrlFees.SelectedTab = tbPageOwnerDetail
                    Me.Text = "Fees - Owner Detail (" & txtOwnerName.Text & ")"
                End If
                nOwnerID = OwnerID
            Else
                LoadugFacilityList()
            End If

            If FacilityID > 0 Then
                tbCntrlFees.SelectedTab = tbPageFacilityDetail
                PopulateFacilityInfo(Integer.Parse(FacilityID))
                'ProcessFacTransactions()
                nFacilityID = FacilityID
            Else
                lblFacilityIDValue.Text = ""
            End If
            'Add any initialization after the InitializeComponent() call
            pnlInvoices.Dock = DockStyle.Fill
            pnlManualInvoice.Dock = DockStyle.Fill
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
        MusterContainer.AppUser.LogEntry(Me.Text, MyGuid.ToString)
        MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "WindowName", Me.Text)
        MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "Fees")
        Try
            InitControls()
            If OwnerID > 0 Then
                PopulateOwnerInfo(Integer.Parse(OwnerID))
                If FacilityID <= 0 Then
                    tbCntrlFees.SelectedTab = tbPageOwnerDetail
                    Me.Text = "Fees - Owner Detail (" & txtOwnerName.Text & ")"
                End If
            Else
                LoadugFacilityList()
            End If

            If FacilityID > 0 Then
                tbCntrlFees.SelectedTab = tbPageFacilityDetail
                PopulateFacilityInfo(Integer.Parse(FacilityID))
                'ProcessFacTransactions()
            Else
                lblFacilityIDValue.Text = ""
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
        MusterContainer.AppUser.LogEntry(Me.Text, MyGuid.ToString)
        MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "WindowName", Me.Text)
        MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "Fees")
        'Ends here 
        Try

            InitControls()
            PopulateOwnerInfo(pOwn.ID)
            ' LoadugFacilityList()

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
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
    Friend WithEvents lblOwnerID As System.Windows.Forms.Label
    Friend WithEvents lblOwnerType As System.Windows.Forms.Label
    Friend WithEvents lblName As System.Windows.Forms.Label
    Friend WithEvents lblAddress As System.Windows.Forms.Label
    Friend WithEvents lblOwnerStatus As System.Windows.Forms.Label
    Friend WithEvents lblEnsiteID As System.Windows.Forms.Label
    Friend WithEvents lblPhone As System.Windows.Forms.Label
    Friend WithEvents lblPhone2 As System.Windows.Forms.Label
    Public WithEvents mskTxtOwnerFax As AxMSMask.AxMaskEdBox
    Public WithEvents mskTxtOwnerPhone2 As AxMSMask.AxMaskEdBox
    Public WithEvents mskTxtOwnerPhone As AxMSMask.AxMaskEdBox
    Friend WithEvents lblFax As System.Windows.Forms.Label
    Friend WithEvents lblOwnerEmail As System.Windows.Forms.Label
    Public WithEvents txtOwnerEmail As System.Windows.Forms.TextBox
    Friend WithEvents pnlOwnerBottom As System.Windows.Forms.Panel
    Friend WithEvents lblOwnerFeeTotals As System.Windows.Forms.Label
    Friend WithEvents lblBalance As System.Windows.Forms.Label
    Friend WithEvents txtBalance As System.Windows.Forms.TextBox
    Friend WithEvents lblOverage As System.Windows.Forms.Label
    Friend WithEvents txtOverage As System.Windows.Forms.TextBox
    Friend WithEvents txtFinal As System.Windows.Forms.TextBox
    Friend WithEvents lblFinal As System.Windows.Forms.Label
    Friend WithEvents lblBankruptcy As System.Windows.Forms.Label
    Friend WithEvents lblCapter As System.Windows.Forms.Label
    Friend WithEvents lblDate As System.Windows.Forms.Label
    Friend WithEvents txtChapter As System.Windows.Forms.TextBox
    Friend WithEvents txtDate As System.Windows.Forms.TextBox
    Friend WithEvents tbPageFacilities As System.Windows.Forms.TabPage
    Friend WithEvents tbPageInvoices As System.Windows.Forms.TabPage
    Friend WithEvents tbPageReceipts As System.Windows.Forms.TabPage
    Friend WithEvents tbPageRefunds As System.Windows.Forms.TabPage
    Friend WithEvents pnlOwnerFacilityBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlFacDetails As System.Windows.Forms.Panel
    Friend WithEvents ugFacilityList As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlInvoices As System.Windows.Forms.Panel
    Friend WithEvents pnlInvoicesTop As System.Windows.Forms.Panel
    Friend WithEvents rdBtnAllInvoices As System.Windows.Forms.RadioButton
    Friend WithEvents rdBtnLateInvoices As System.Windows.Forms.RadioButton
    Friend WithEvents rdBtnOutstandingBalance As System.Windows.Forms.RadioButton
    Friend WithEvents pnlInvoicesBottom As System.Windows.Forms.Panel
    Public WithEvents lblTotalAdjustments As System.Windows.Forms.Label
    Public WithEvents lblTotalCredits As System.Windows.Forms.Label
    Public WithEvents lblTotalPayments As System.Windows.Forms.Label
    Public WithEvents lblTotalCharges As System.Windows.Forms.Label
    Public WithEvents lblInvoices As System.Windows.Forms.Label
    Public WithEvents lblTotalInvoices As System.Windows.Forms.Label
    Friend WithEvents lblOwnerTotalsfor As System.Windows.Forms.Label
    Friend WithEvents pnlInvoicesControls As System.Windows.Forms.Panel
    Friend WithEvents lblTotalOverage As System.Windows.Forms.Label
    Friend WithEvents pnlInvoicesDetails As System.Windows.Forms.Panel
    Friend WithEvents ugInvoices As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlManualInvoice As System.Windows.Forms.Panel
    Friend WithEvents pnlManualInvoiceBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlManualInvoiceDetails As System.Windows.Forms.Panel
    Friend WithEvents lblTotalTanks As System.Windows.Forms.Label
    Friend WithEvents lblTotalTanksValue As System.Windows.Forms.Label
    Friend WithEvents lblInvoiceAmount As System.Windows.Forms.Label
    Friend WithEvents btnIssueInvoice As System.Windows.Forms.Button
    Friend WithEvents ugManualInvoices As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnManualInvoiceCancel As System.Windows.Forms.Button
    Friend WithEvents btnRecordInvoiceNoDate As System.Windows.Forms.Button
    Friend WithEvents btnReqLateFeeWaiver As System.Windows.Forms.Button
    Friend WithEvents btnReallocateOverage As System.Windows.Forms.Button
    Friend WithEvents btnGenCreditMemo As System.Windows.Forms.Button
    Friend WithEvents btnGenerateInvoice As System.Windows.Forms.Button
    Friend WithEvents btnGenerateDebitMemo As System.Windows.Forms.Button
    Friend WithEvents pnlOwnerTotalsContainer As System.Windows.Forms.Panel
    Public WithEvents lblCurrentAdjustments As System.Windows.Forms.Label
    Public WithEvents lblCurrentCredits As System.Windows.Forms.Label
    Public WithEvents lblCurrentPayments As System.Windows.Forms.Label
    Public WithEvents lblTotalDue As System.Windows.Forms.Label
    Public WithEvents lblLatePenalty As System.Windows.Forms.Label
    Public WithEvents lblCurrentFees As System.Windows.Forms.Label
    Public WithEvents lblPriorBalance As System.Windows.Forms.Label
    Friend WithEvents lblNoOfFacilities As System.Windows.Forms.Label
    Friend WithEvents pnlReceiptsBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlReceiptsDetails As System.Windows.Forms.Panel
    Friend WithEvents btnOverpaymentReason As System.Windows.Forms.Button
    Friend WithEvents btnAdjustPayment As System.Windows.Forms.Button
    Friend WithEvents pnlRefundsBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlRefundsDetails As System.Windows.Forms.Panel
    Friend WithEvents btnGenerateRefund As System.Windows.Forms.Button
    Friend WithEvents btnEnterWarrant As System.Windows.Forms.Button
    Friend WithEvents btnChangeWarrant As System.Windows.Forms.Button
    Friend WithEvents ugRefunds As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnl_FacilityDetail As System.Windows.Forms.Panel
    Public WithEvents txtDueByNF As System.Windows.Forms.TextBox
    Friend WithEvents lblFacilitySigNFDue As System.Windows.Forms.Label
    Public WithEvents lblDateTransfered As System.Windows.Forms.Label
    Friend WithEvents ll As System.Windows.Forms.Label
    Friend WithEvents lblLUSTSite As System.Windows.Forms.Label
    Public WithEvents chkLUSTSite As System.Windows.Forms.CheckBox
    Friend WithEvents lblCAPCandidate As System.Windows.Forms.Label
    Public WithEvents chkCAPCandidate As System.Windows.Forms.CheckBox
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
    Public WithEvents txtFacilityAddress As System.Windows.Forms.TextBox
    Friend WithEvents lblFacilitySIC As System.Windows.Forms.Label
    Public WithEvents dtPickFacilityRecvd As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblDateReceived As System.Windows.Forms.Label
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
    Public WithEvents txtFacilityName As System.Windows.Forms.TextBox
    Friend WithEvents lblFacilityAddress As System.Windows.Forms.Label
    Friend WithEvents lblFacilityName As System.Windows.Forms.Label
    Friend WithEvents lblFacilityLatDegree As System.Windows.Forms.Label
    Friend WithEvents lnkLblNextFac As System.Windows.Forms.LinkLabel
    Friend WithEvents lnkLblPrevFacility As System.Windows.Forms.LinkLabel
    Friend WithEvents tbCtrlOwnerTransactions As System.Windows.Forms.TabControl
    Friend WithEvents pnlFacilityBottom As System.Windows.Forms.Panel
    Friend WithEvents tbCtrlFacTransactions As System.Windows.Forms.TabControl
    Friend WithEvents tbPageByTransaction As System.Windows.Forms.TabPage
    Friend WithEvents tbPageByInvoice As System.Windows.Forms.TabPage
    Friend WithEvents tbPageOldAccessHis As System.Windows.Forms.TabPage
    Friend WithEvents pnlByInvoiceDetails As System.Windows.Forms.Panel
    Friend WithEvents ugByInvoice As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlByInvoiceTop As System.Windows.Forms.Panel
    Friend WithEvents rdBtnCurrentOutstanding As System.Windows.Forms.RadioButton
    Friend WithEvents rdBtnAllYears As System.Windows.Forms.RadioButton
    Friend WithEvents pnlByInvoiceBottom As System.Windows.Forms.Panel
    Friend WithEvents btnFacGenerateCreditMemo As System.Windows.Forms.Button
    Friend WithEvents btnFacGenerateInvoice As System.Windows.Forms.Button
    Friend WithEvents pnlByTransacTop As System.Windows.Forms.Panel
    Friend WithEvents pnlByTransacDetails As System.Windows.Forms.Panel
    Friend WithEvents rdBtnCurrentFiscalYear As System.Windows.Forms.RadioButton
    Friend WithEvents rdBtnByTransAllYears As System.Windows.Forms.RadioButton
    Friend WithEvents ugByTransaction As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents ugOldAccessHistory As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents tbCntrlFees As System.Windows.Forms.TabControl
    Friend WithEvents tbPageOwnerDetail As System.Windows.Forms.TabPage
    Friend WithEvents tbPageFacilityDetail As System.Windows.Forms.TabPage
    Public WithEvents lblOwnerIDValue As System.Windows.Forms.Label
    Public WithEvents cmbOwnerType As System.Windows.Forms.ComboBox
    Public WithEvents pnlOwnerDetails As System.Windows.Forms.Panel
    Public WithEvents txtOwnerName As System.Windows.Forms.TextBox
    Public WithEvents txtOwnerAddress As System.Windows.Forms.TextBox
    Public WithEvents lblOwnerActiveOrNot As System.Windows.Forms.Label
    Public WithEvents txtOwnerAIID As System.Windows.Forms.TextBox
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents btnFacComments As System.Windows.Forms.Button
    Friend WithEvents btnFacFlags As System.Windows.Forms.Button
    Friend WithEvents pnlOwnerButtons As System.Windows.Forms.Panel
    Friend WithEvents btnOwnerFlag As System.Windows.Forms.Button
    Friend WithEvents btnOwnerComment As System.Windows.Forms.Button
    Public WithEvents lblFacToDateBalance As System.Windows.Forms.Label
    Public WithEvents lblTotalBalance As System.Windows.Forms.Label
    Public WithEvents lblTotalOverageValue As System.Windows.Forms.Label
    Public WithEvents lblTotalLegal As System.Windows.Forms.Label
    Friend WithEvents pnlOwnerReceiptTop As System.Windows.Forms.Panel
    Friend WithEvents rdAllFY As System.Windows.Forms.RadioButton
    Friend WithEvents rdCurrentFY As System.Windows.Forms.RadioButton
    Friend WithEvents pnlOwnerReceiptFill As System.Windows.Forms.Panel
    Friend WithEvents ugReceipts As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnManualInvoiceTop As System.Windows.Forms.Panel
    Friend WithEvents lblInvoiceAmountValue As System.Windows.Forms.Label
    Friend WithEvents cmbFeeType As System.Windows.Forms.ComboBox
    Friend WithEvents lblFeeType As System.Windows.Forms.Label
    Friend WithEvents lblMIOwnerID As System.Windows.Forms.Label
    Friend WithEvents lblMIOwnerIDValue As System.Windows.Forms.Label
    Friend WithEvents lblMIOwnerName As System.Windows.Forms.Label
    Friend WithEvents lblMIOwnerNameValue As System.Windows.Forms.Label
    Friend WithEvents rdbtnInvoiceCurrentFY As System.Windows.Forms.RadioButton
    Friend WithEvents pnlFeesHeader As System.Windows.Forms.Panel
    Public WithEvents lblOwnerLastEditedOn As System.Windows.Forms.Label
    Public WithEvents lblOwnerLastEditedBy As System.Windows.Forms.Label
    Public WithEvents lblLegal As System.Windows.Forms.Label
    Friend WithEvents tbPageOwnerDocuments As System.Windows.Forms.TabPage
    Friend WithEvents tbPageFacilityDocuments As System.Windows.Forms.TabPage
    Friend WithEvents UCFacilityDocuments As MUSTER.DocumentViewControl
    Friend WithEvents UCOwnerDocuments As MUSTER.DocumentViewControl
    Friend WithEvents btnFeesOwnerLabels As System.Windows.Forms.Button
    Friend WithEvents btnFeesOwnerEnvelopes As System.Windows.Forms.Button
    Friend WithEvents btnFeesFacLabels As System.Windows.Forms.Button
    Friend WithEvents btnFeesFacEnvelopes As System.Windows.Forms.Button
    Friend WithEvents txtBP2KOwnerID As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Public WithEvents mskTxtFacilityFax As AxMSMask.AxMaskEdBox
    Public WithEvents mskTxtFacilityPhone As AxMSMask.AxMaskEdBox
    Friend WithEvents lblFacilityFax As System.Windows.Forms.Label
    Friend WithEvents lblFacilityPhone As System.Windows.Forms.Label
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
    Friend WithEvents Label3 As System.Windows.Forms.Label
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
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnFacEditCreditMemo As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents tbPageOverages As System.Windows.Forms.TabPage
    Friend WithEvents ugOverage As Infragistics.Win.UltraWinGrid.UltraGrid
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Fees))
        Me.tbCntrlFees = New System.Windows.Forms.TabControl
        Me.tbPageOwnerDetail = New System.Windows.Forms.TabPage
        Me.pnlOwnerBottom = New System.Windows.Forms.Panel
        Me.tbCtrlOwnerTransactions = New System.Windows.Forms.TabControl
        Me.tbPageFacilities = New System.Windows.Forms.TabPage
        Me.pnlFacDetails = New System.Windows.Forms.Panel
        Me.ugFacilityList = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlOwnerFacilityBottom = New System.Windows.Forms.Panel
        Me.pnlOwnerTotalsContainer = New System.Windows.Forms.Panel
        Me.lblLegal = New System.Windows.Forms.Label
        Me.lblFacToDateBalance = New System.Windows.Forms.Label
        Me.lblCurrentAdjustments = New System.Windows.Forms.Label
        Me.lblCurrentCredits = New System.Windows.Forms.Label
        Me.lblCurrentPayments = New System.Windows.Forms.Label
        Me.lblTotalDue = New System.Windows.Forms.Label
        Me.lblLatePenalty = New System.Windows.Forms.Label
        Me.lblCurrentFees = New System.Windows.Forms.Label
        Me.lblPriorBalance = New System.Windows.Forms.Label
        Me.lblNoOfFacilities = New System.Windows.Forms.Label
        Me.tbPageRefunds = New System.Windows.Forms.TabPage
        Me.pnlRefundsDetails = New System.Windows.Forms.Panel
        Me.ugRefunds = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlRefundsBottom = New System.Windows.Forms.Panel
        Me.btnChangeWarrant = New System.Windows.Forms.Button
        Me.btnEnterWarrant = New System.Windows.Forms.Button
        Me.btnGenerateRefund = New System.Windows.Forms.Button
        Me.tbPageReceipts = New System.Windows.Forms.TabPage
        Me.pnlReceiptsDetails = New System.Windows.Forms.Panel
        Me.pnlOwnerReceiptFill = New System.Windows.Forms.Panel
        Me.ugReceipts = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlOwnerReceiptTop = New System.Windows.Forms.Panel
        Me.rdAllFY = New System.Windows.Forms.RadioButton
        Me.rdCurrentFY = New System.Windows.Forms.RadioButton
        Me.pnlReceiptsBottom = New System.Windows.Forms.Panel
        Me.btnAdjustPayment = New System.Windows.Forms.Button
        Me.btnOverpaymentReason = New System.Windows.Forms.Button
        Me.tbPageOverages = New System.Windows.Forms.TabPage
        Me.ugOverage = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.tbPageOwnerDocuments = New System.Windows.Forms.TabPage
        Me.UCOwnerDocuments = New MUSTER.DocumentViewControl
        Me.tbPageOwnerContactList = New System.Windows.Forms.TabPage
        Me.pnlOwnerContactContainer = New System.Windows.Forms.Panel
        Me.ugOwnerContacts = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.Label3 = New System.Windows.Forms.Label
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
        Me.tbPageInvoices = New System.Windows.Forms.TabPage
        Me.pnlManualInvoice = New System.Windows.Forms.Panel
        Me.pnlManualInvoiceDetails = New System.Windows.Forms.Panel
        Me.ugManualInvoices = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlManualInvoiceBottom = New System.Windows.Forms.Panel
        Me.btnManualInvoiceCancel = New System.Windows.Forms.Button
        Me.btnIssueInvoice = New System.Windows.Forms.Button
        Me.pnManualInvoiceTop = New System.Windows.Forms.Panel
        Me.lblMIOwnerNameValue = New System.Windows.Forms.Label
        Me.lblMIOwnerName = New System.Windows.Forms.Label
        Me.lblMIOwnerIDValue = New System.Windows.Forms.Label
        Me.lblMIOwnerID = New System.Windows.Forms.Label
        Me.lblFeeType = New System.Windows.Forms.Label
        Me.cmbFeeType = New System.Windows.Forms.ComboBox
        Me.lblInvoiceAmountValue = New System.Windows.Forms.Label
        Me.lblInvoiceAmount = New System.Windows.Forms.Label
        Me.lblTotalTanksValue = New System.Windows.Forms.Label
        Me.lblTotalTanks = New System.Windows.Forms.Label
        Me.pnlInvoices = New System.Windows.Forms.Panel
        Me.pnlInvoicesDetails = New System.Windows.Forms.Panel
        Me.ugInvoices = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlInvoicesControls = New System.Windows.Forms.Panel
        Me.btnGenerateDebitMemo = New System.Windows.Forms.Button
        Me.btnRecordInvoiceNoDate = New System.Windows.Forms.Button
        Me.btnReqLateFeeWaiver = New System.Windows.Forms.Button
        Me.btnReallocateOverage = New System.Windows.Forms.Button
        Me.btnGenCreditMemo = New System.Windows.Forms.Button
        Me.btnGenerateInvoice = New System.Windows.Forms.Button
        Me.lblTotalOverage = New System.Windows.Forms.Label
        Me.pnlInvoicesBottom = New System.Windows.Forms.Panel
        Me.lblTotalOverageValue = New System.Windows.Forms.Label
        Me.lblTotalLegal = New System.Windows.Forms.Label
        Me.lblTotalBalance = New System.Windows.Forms.Label
        Me.lblTotalAdjustments = New System.Windows.Forms.Label
        Me.lblTotalCredits = New System.Windows.Forms.Label
        Me.lblTotalPayments = New System.Windows.Forms.Label
        Me.lblTotalCharges = New System.Windows.Forms.Label
        Me.lblInvoices = New System.Windows.Forms.Label
        Me.lblTotalInvoices = New System.Windows.Forms.Label
        Me.lblOwnerTotalsfor = New System.Windows.Forms.Label
        Me.pnlInvoicesTop = New System.Windows.Forms.Panel
        Me.rdbtnInvoiceCurrentFY = New System.Windows.Forms.RadioButton
        Me.rdBtnAllInvoices = New System.Windows.Forms.RadioButton
        Me.rdBtnLateInvoices = New System.Windows.Forms.RadioButton
        Me.rdBtnOutstandingBalance = New System.Windows.Forms.RadioButton
        Me.pnlOwnerDetails = New System.Windows.Forms.Panel
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtBP2KOwnerID = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.btnFeesOwnerLabels = New System.Windows.Forms.Button
        Me.btnFeesOwnerEnvelopes = New System.Windows.Forms.Button
        Me.pnlOwnerButtons = New System.Windows.Forms.Panel
        Me.btnOwnerFlag = New System.Windows.Forms.Button
        Me.btnOwnerComment = New System.Windows.Forms.Button
        Me.txtDate = New System.Windows.Forms.TextBox
        Me.txtChapter = New System.Windows.Forms.TextBox
        Me.lblDate = New System.Windows.Forms.Label
        Me.lblCapter = New System.Windows.Forms.Label
        Me.lblBankruptcy = New System.Windows.Forms.Label
        Me.lblFinal = New System.Windows.Forms.Label
        Me.txtFinal = New System.Windows.Forms.TextBox
        Me.txtOverage = New System.Windows.Forms.TextBox
        Me.lblOverage = New System.Windows.Forms.Label
        Me.txtBalance = New System.Windows.Forms.TextBox
        Me.lblBalance = New System.Windows.Forms.Label
        Me.lblOwnerFeeTotals = New System.Windows.Forms.Label
        Me.lblOwnerEmail = New System.Windows.Forms.Label
        Me.txtOwnerEmail = New System.Windows.Forms.TextBox
        Me.mskTxtOwnerFax = New AxMSMask.AxMaskEdBox
        Me.mskTxtOwnerPhone2 = New AxMSMask.AxMaskEdBox
        Me.mskTxtOwnerPhone = New AxMSMask.AxMaskEdBox
        Me.lblFax = New System.Windows.Forms.Label
        Me.lblOwnerActiveOrNot = New System.Windows.Forms.Label
        Me.txtOwnerAIID = New System.Windows.Forms.TextBox
        Me.lblPhone2 = New System.Windows.Forms.Label
        Me.lblPhone = New System.Windows.Forms.Label
        Me.lblEnsiteID = New System.Windows.Forms.Label
        Me.lblOwnerStatus = New System.Windows.Forms.Label
        Me.txtOwnerAddress = New System.Windows.Forms.TextBox
        Me.lblAddress = New System.Windows.Forms.Label
        Me.txtOwnerName = New System.Windows.Forms.TextBox
        Me.lblName = New System.Windows.Forms.Label
        Me.cmbOwnerType = New System.Windows.Forms.ComboBox
        Me.lblOwnerType = New System.Windows.Forms.Label
        Me.lblOwnerIDValue = New System.Windows.Forms.Label
        Me.lblOwnerID = New System.Windows.Forms.Label
        Me.tbPageFacilityDetail = New System.Windows.Forms.TabPage
        Me.pnlFacilityBottom = New System.Windows.Forms.Panel
        Me.tbCtrlFacTransactions = New System.Windows.Forms.TabControl
        Me.tbPageByTransaction = New System.Windows.Forms.TabPage
        Me.pnlByTransacDetails = New System.Windows.Forms.Panel
        Me.ugByTransaction = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlByTransacTop = New System.Windows.Forms.Panel
        Me.rdBtnCurrentFiscalYear = New System.Windows.Forms.RadioButton
        Me.rdBtnByTransAllYears = New System.Windows.Forms.RadioButton
        Me.tbPageByInvoice = New System.Windows.Forms.TabPage
        Me.pnlByInvoiceDetails = New System.Windows.Forms.Panel
        Me.ugByInvoice = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlByInvoiceTop = New System.Windows.Forms.Panel
        Me.rdBtnCurrentOutstanding = New System.Windows.Forms.RadioButton
        Me.rdBtnAllYears = New System.Windows.Forms.RadioButton
        Me.pnlByInvoiceBottom = New System.Windows.Forms.Panel
        Me.btnFacGenerateCreditMemo = New System.Windows.Forms.Button
        Me.btnFacGenerateInvoice = New System.Windows.Forms.Button
        Me.btnFacEditCreditMemo = New System.Windows.Forms.Button
        Me.tbPageOldAccessHis = New System.Windows.Forms.TabPage
        Me.ugOldAccessHistory = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.tbPageFacilityDocuments = New System.Windows.Forms.TabPage
        Me.UCFacilityDocuments = New MUSTER.DocumentViewControl
        Me.tbPageFacilityContactList = New System.Windows.Forms.TabPage
        Me.pnlFacilityContactContainer = New System.Windows.Forms.Panel
        Me.ugFacilityContacts = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.Label4 = New System.Windows.Forms.Label
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
        Me.mskTxtFacilityFax = New AxMSMask.AxMaskEdBox
        Me.mskTxtFacilityPhone = New AxMSMask.AxMaskEdBox
        Me.lblFacilityFax = New System.Windows.Forms.Label
        Me.lblFacilityPhone = New System.Windows.Forms.Label
        Me.btnFeesFacLabels = New System.Windows.Forms.Button
        Me.btnFeesFacEnvelopes = New System.Windows.Forms.Button
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.btnFacComments = New System.Windows.Forms.Button
        Me.btnFacFlags = New System.Windows.Forms.Button
        Me.lnkLblNextFac = New System.Windows.Forms.LinkLabel
        Me.lnkLblPrevFacility = New System.Windows.Forms.LinkLabel
        Me.txtDueByNF = New System.Windows.Forms.TextBox
        Me.lblFacilitySigNFDue = New System.Windows.Forms.Label
        Me.lblDateTransfered = New System.Windows.Forms.Label
        Me.ll = New System.Windows.Forms.Label
        Me.lblLUSTSite = New System.Windows.Forms.Label
        Me.chkLUSTSite = New System.Windows.Forms.CheckBox
        Me.lblCAPCandidate = New System.Windows.Forms.Label
        Me.chkCAPCandidate = New System.Windows.Forms.CheckBox
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
        Me.txtFacilityAddress = New System.Windows.Forms.TextBox
        Me.lblFacilitySIC = New System.Windows.Forms.Label
        Me.dtPickFacilityRecvd = New System.Windows.Forms.DateTimePicker
        Me.lblDateReceived = New System.Windows.Forms.Label
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
        Me.txtFacilityName = New System.Windows.Forms.TextBox
        Me.lblFacilityAddress = New System.Windows.Forms.Label
        Me.lblFacilityName = New System.Windows.Forms.Label
        Me.lblFacilityLatDegree = New System.Windows.Forms.Label
        Me.txtFacilitySIC = New System.Windows.Forms.Label
        Me.pnlFeesHeader = New System.Windows.Forms.Panel
        Me.lblOwnerLastEditedOn = New System.Windows.Forms.Label
        Me.lblOwnerLastEditedBy = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.tbCntrlFees.SuspendLayout()
        Me.tbPageOwnerDetail.SuspendLayout()
        Me.pnlOwnerBottom.SuspendLayout()
        Me.tbCtrlOwnerTransactions.SuspendLayout()
        Me.tbPageFacilities.SuspendLayout()
        Me.pnlFacDetails.SuspendLayout()
        CType(Me.ugFacilityList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlOwnerFacilityBottom.SuspendLayout()
        Me.pnlOwnerTotalsContainer.SuspendLayout()
        Me.tbPageRefunds.SuspendLayout()
        Me.pnlRefundsDetails.SuspendLayout()
        CType(Me.ugRefunds, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlRefundsBottom.SuspendLayout()
        Me.tbPageReceipts.SuspendLayout()
        Me.pnlReceiptsDetails.SuspendLayout()
        Me.pnlOwnerReceiptFill.SuspendLayout()
        CType(Me.ugReceipts, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlOwnerReceiptTop.SuspendLayout()
        Me.pnlReceiptsBottom.SuspendLayout()
        Me.tbPageOverages.SuspendLayout()
        CType(Me.ugOverage, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbPageOwnerDocuments.SuspendLayout()
        Me.tbPageOwnerContactList.SuspendLayout()
        Me.pnlOwnerContactContainer.SuspendLayout()
        CType(Me.ugOwnerContacts, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlOwnerContactButtons.SuspendLayout()
        Me.pnlOwnerContactHeader.SuspendLayout()
        Me.tbPageInvoices.SuspendLayout()
        Me.pnlManualInvoice.SuspendLayout()
        Me.pnlManualInvoiceDetails.SuspendLayout()
        CType(Me.ugManualInvoices, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlManualInvoiceBottom.SuspendLayout()
        Me.pnManualInvoiceTop.SuspendLayout()
        Me.pnlInvoices.SuspendLayout()
        Me.pnlInvoicesDetails.SuspendLayout()
        CType(Me.ugInvoices, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlInvoicesControls.SuspendLayout()
        Me.pnlInvoicesBottom.SuspendLayout()
        Me.pnlInvoicesTop.SuspendLayout()
        Me.pnlOwnerDetails.SuspendLayout()
        Me.pnlOwnerButtons.SuspendLayout()
        CType(Me.mskTxtOwnerFax, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtOwnerPhone2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtOwnerPhone, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbPageFacilityDetail.SuspendLayout()
        Me.pnlFacilityBottom.SuspendLayout()
        Me.tbCtrlFacTransactions.SuspendLayout()
        Me.tbPageByTransaction.SuspendLayout()
        Me.pnlByTransacDetails.SuspendLayout()
        CType(Me.ugByTransaction, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlByTransacTop.SuspendLayout()
        Me.tbPageByInvoice.SuspendLayout()
        Me.pnlByInvoiceDetails.SuspendLayout()
        CType(Me.ugByInvoice, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlByInvoiceTop.SuspendLayout()
        Me.pnlByInvoiceBottom.SuspendLayout()
        Me.tbPageOldAccessHis.SuspendLayout()
        CType(Me.ugOldAccessHistory, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbPageFacilityDocuments.SuspendLayout()
        Me.tbPageFacilityContactList.SuspendLayout()
        Me.pnlFacilityContactContainer.SuspendLayout()
        CType(Me.ugFacilityContacts, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlFacilityContactBottom.SuspendLayout()
        Me.pnlFacilityContactHeader.SuspendLayout()
        Me.pnl_FacilityDetail.SuspendLayout()
        CType(Me.mskTxtFacilityFax, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtFacilityPhone, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.pnlFeesHeader.SuspendLayout()
        Me.SuspendLayout()
        '
        'tbCntrlFees
        '
        Me.tbCntrlFees.Controls.Add(Me.tbPageOwnerDetail)
        Me.tbCntrlFees.Controls.Add(Me.tbPageFacilityDetail)
        Me.tbCntrlFees.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbCntrlFees.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbCntrlFees.Location = New System.Drawing.Point(0, 24)
        Me.tbCntrlFees.Multiline = True
        Me.tbCntrlFees.Name = "tbCntrlFees"
        Me.tbCntrlFees.SelectedIndex = 0
        Me.tbCntrlFees.ShowToolTips = True
        Me.tbCntrlFees.Size = New System.Drawing.Size(992, 670)
        Me.tbCntrlFees.TabIndex = 0
        '
        'tbPageOwnerDetail
        '
        Me.tbPageOwnerDetail.Controls.Add(Me.pnlOwnerBottom)
        Me.tbPageOwnerDetail.Controls.Add(Me.pnlOwnerDetails)
        Me.tbPageOwnerDetail.Location = New System.Drawing.Point(4, 24)
        Me.tbPageOwnerDetail.Name = "tbPageOwnerDetail"
        Me.tbPageOwnerDetail.Size = New System.Drawing.Size(984, 642)
        Me.tbPageOwnerDetail.TabIndex = 0
        Me.tbPageOwnerDetail.Text = "Owner Details"
        '
        'pnlOwnerBottom
        '
        Me.pnlOwnerBottom.Controls.Add(Me.tbCtrlOwnerTransactions)
        Me.pnlOwnerBottom.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlOwnerBottom.Location = New System.Drawing.Point(0, 184)
        Me.pnlOwnerBottom.Name = "pnlOwnerBottom"
        Me.pnlOwnerBottom.Size = New System.Drawing.Size(984, 458)
        Me.pnlOwnerBottom.TabIndex = 1
        '
        'tbCtrlOwnerTransactions
        '
        Me.tbCtrlOwnerTransactions.Controls.Add(Me.tbPageFacilities)
        Me.tbCtrlOwnerTransactions.Controls.Add(Me.tbPageRefunds)
        Me.tbCtrlOwnerTransactions.Controls.Add(Me.tbPageReceipts)
        Me.tbCtrlOwnerTransactions.Controls.Add(Me.tbPageOverages)
        Me.tbCtrlOwnerTransactions.Controls.Add(Me.tbPageOwnerDocuments)
        Me.tbCtrlOwnerTransactions.Controls.Add(Me.tbPageOwnerContactList)
        Me.tbCtrlOwnerTransactions.Controls.Add(Me.tbPageInvoices)
        Me.tbCtrlOwnerTransactions.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbCtrlOwnerTransactions.Location = New System.Drawing.Point(0, 0)
        Me.tbCtrlOwnerTransactions.Name = "tbCtrlOwnerTransactions"
        Me.tbCtrlOwnerTransactions.SelectedIndex = 0
        Me.tbCtrlOwnerTransactions.Size = New System.Drawing.Size(984, 458)
        Me.tbCtrlOwnerTransactions.TabIndex = 0
        '
        'tbPageFacilities
        '
        Me.tbPageFacilities.Controls.Add(Me.pnlFacDetails)
        Me.tbPageFacilities.Controls.Add(Me.pnlOwnerFacilityBottom)
        Me.tbPageFacilities.Location = New System.Drawing.Point(4, 24)
        Me.tbPageFacilities.Name = "tbPageFacilities"
        Me.tbPageFacilities.Size = New System.Drawing.Size(976, 430)
        Me.tbPageFacilities.TabIndex = 0
        Me.tbPageFacilities.Text = "Facilities"
        '
        'pnlFacDetails
        '
        Me.pnlFacDetails.Controls.Add(Me.ugFacilityList)
        Me.pnlFacDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlFacDetails.Location = New System.Drawing.Point(0, 0)
        Me.pnlFacDetails.Name = "pnlFacDetails"
        Me.pnlFacDetails.Size = New System.Drawing.Size(976, 406)
        Me.pnlFacDetails.TabIndex = 21
        '
        'ugFacilityList
        '
        Me.ugFacilityList.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugFacilityList.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle
        Me.ugFacilityList.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugFacilityList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugFacilityList.Location = New System.Drawing.Point(0, 0)
        Me.ugFacilityList.Name = "ugFacilityList"
        Me.ugFacilityList.Size = New System.Drawing.Size(976, 406)
        Me.ugFacilityList.TabIndex = 0
        '
        'pnlOwnerFacilityBottom
        '
        Me.pnlOwnerFacilityBottom.Controls.Add(Me.pnlOwnerTotalsContainer)
        Me.pnlOwnerFacilityBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlOwnerFacilityBottom.Location = New System.Drawing.Point(0, 406)
        Me.pnlOwnerFacilityBottom.Name = "pnlOwnerFacilityBottom"
        Me.pnlOwnerFacilityBottom.Size = New System.Drawing.Size(976, 24)
        Me.pnlOwnerFacilityBottom.TabIndex = 20
        '
        'pnlOwnerTotalsContainer
        '
        Me.pnlOwnerTotalsContainer.Controls.Add(Me.lblLegal)
        Me.pnlOwnerTotalsContainer.Controls.Add(Me.lblFacToDateBalance)
        Me.pnlOwnerTotalsContainer.Controls.Add(Me.lblCurrentAdjustments)
        Me.pnlOwnerTotalsContainer.Controls.Add(Me.lblCurrentCredits)
        Me.pnlOwnerTotalsContainer.Controls.Add(Me.lblCurrentPayments)
        Me.pnlOwnerTotalsContainer.Controls.Add(Me.lblTotalDue)
        Me.pnlOwnerTotalsContainer.Controls.Add(Me.lblLatePenalty)
        Me.pnlOwnerTotalsContainer.Controls.Add(Me.lblCurrentFees)
        Me.pnlOwnerTotalsContainer.Controls.Add(Me.lblPriorBalance)
        Me.pnlOwnerTotalsContainer.Controls.Add(Me.lblNoOfFacilities)
        Me.pnlOwnerTotalsContainer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlOwnerTotalsContainer.Location = New System.Drawing.Point(0, 0)
        Me.pnlOwnerTotalsContainer.Name = "pnlOwnerTotalsContainer"
        Me.pnlOwnerTotalsContainer.Size = New System.Drawing.Size(976, 24)
        Me.pnlOwnerTotalsContainer.TabIndex = 2
        '
        'lblLegal
        '
        Me.lblLegal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblLegal.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblLegal.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblLegal.Location = New System.Drawing.Point(900, 0)
        Me.lblLegal.Name = "lblLegal"
        Me.lblLegal.Size = New System.Drawing.Size(75, 24)
        Me.lblLegal.TabIndex = 26
        Me.lblLegal.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblFacToDateBalance
        '
        Me.lblFacToDateBalance.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblFacToDateBalance.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblFacToDateBalance.Location = New System.Drawing.Point(825, 0)
        Me.lblFacToDateBalance.Name = "lblFacToDateBalance"
        Me.lblFacToDateBalance.Size = New System.Drawing.Size(75, 24)
        Me.lblFacToDateBalance.TabIndex = 25
        Me.lblFacToDateBalance.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblCurrentAdjustments
        '
        Me.lblCurrentAdjustments.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCurrentAdjustments.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblCurrentAdjustments.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblCurrentAdjustments.Location = New System.Drawing.Point(750, 0)
        Me.lblCurrentAdjustments.Name = "lblCurrentAdjustments"
        Me.lblCurrentAdjustments.Size = New System.Drawing.Size(75, 24)
        Me.lblCurrentAdjustments.TabIndex = 22
        Me.lblCurrentAdjustments.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblCurrentCredits
        '
        Me.lblCurrentCredits.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCurrentCredits.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblCurrentCredits.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblCurrentCredits.Location = New System.Drawing.Point(675, 0)
        Me.lblCurrentCredits.Name = "lblCurrentCredits"
        Me.lblCurrentCredits.Size = New System.Drawing.Size(75, 24)
        Me.lblCurrentCredits.TabIndex = 21
        Me.lblCurrentCredits.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblCurrentPayments
        '
        Me.lblCurrentPayments.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCurrentPayments.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblCurrentPayments.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblCurrentPayments.Location = New System.Drawing.Point(600, 0)
        Me.lblCurrentPayments.Name = "lblCurrentPayments"
        Me.lblCurrentPayments.Size = New System.Drawing.Size(75, 24)
        Me.lblCurrentPayments.TabIndex = 20
        Me.lblCurrentPayments.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTotalDue
        '
        Me.lblTotalDue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalDue.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblTotalDue.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblTotalDue.Location = New System.Drawing.Point(525, 0)
        Me.lblTotalDue.Name = "lblTotalDue"
        Me.lblTotalDue.Size = New System.Drawing.Size(75, 24)
        Me.lblTotalDue.TabIndex = 19
        Me.lblTotalDue.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblLatePenalty
        '
        Me.lblLatePenalty.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblLatePenalty.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblLatePenalty.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblLatePenalty.Location = New System.Drawing.Point(450, 0)
        Me.lblLatePenalty.Name = "lblLatePenalty"
        Me.lblLatePenalty.Size = New System.Drawing.Size(75, 24)
        Me.lblLatePenalty.TabIndex = 18
        Me.lblLatePenalty.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblCurrentFees
        '
        Me.lblCurrentFees.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCurrentFees.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblCurrentFees.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblCurrentFees.Location = New System.Drawing.Point(375, 0)
        Me.lblCurrentFees.Name = "lblCurrentFees"
        Me.lblCurrentFees.Size = New System.Drawing.Size(75, 24)
        Me.lblCurrentFees.TabIndex = 17
        Me.lblCurrentFees.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblPriorBalance
        '
        Me.lblPriorBalance.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblPriorBalance.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblPriorBalance.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblPriorBalance.Location = New System.Drawing.Point(300, 0)
        Me.lblPriorBalance.Name = "lblPriorBalance"
        Me.lblPriorBalance.Size = New System.Drawing.Size(75, 24)
        Me.lblPriorBalance.TabIndex = 16
        Me.lblPriorBalance.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblNoOfFacilities
        '
        Me.lblNoOfFacilities.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNoOfFacilities.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblNoOfFacilities.Location = New System.Drawing.Point(0, 0)
        Me.lblNoOfFacilities.Name = "lblNoOfFacilities"
        Me.lblNoOfFacilities.Size = New System.Drawing.Size(300, 24)
        Me.lblNoOfFacilities.TabIndex = 15
        Me.lblNoOfFacilities.Text = "Owner Totals for 3 Facilities:"
        Me.lblNoOfFacilities.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'tbPageRefunds
        '
        Me.tbPageRefunds.Controls.Add(Me.pnlRefundsDetails)
        Me.tbPageRefunds.Controls.Add(Me.pnlRefundsBottom)
        Me.tbPageRefunds.Location = New System.Drawing.Point(4, 24)
        Me.tbPageRefunds.Name = "tbPageRefunds"
        Me.tbPageRefunds.Size = New System.Drawing.Size(976, 430)
        Me.tbPageRefunds.TabIndex = 3
        Me.tbPageRefunds.Text = "Refunds"
        '
        'pnlRefundsDetails
        '
        Me.pnlRefundsDetails.Controls.Add(Me.ugRefunds)
        Me.pnlRefundsDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlRefundsDetails.Location = New System.Drawing.Point(0, 0)
        Me.pnlRefundsDetails.Name = "pnlRefundsDetails"
        Me.pnlRefundsDetails.Size = New System.Drawing.Size(976, 390)
        Me.pnlRefundsDetails.TabIndex = 1
        '
        'ugRefunds
        '
        Me.ugRefunds.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugRefunds.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle
        Me.ugRefunds.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugRefunds.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugRefunds.Location = New System.Drawing.Point(0, 0)
        Me.ugRefunds.Name = "ugRefunds"
        Me.ugRefunds.Size = New System.Drawing.Size(976, 390)
        Me.ugRefunds.TabIndex = 0
        '
        'pnlRefundsBottom
        '
        Me.pnlRefundsBottom.Controls.Add(Me.btnChangeWarrant)
        Me.pnlRefundsBottom.Controls.Add(Me.btnEnterWarrant)
        Me.pnlRefundsBottom.Controls.Add(Me.btnGenerateRefund)
        Me.pnlRefundsBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlRefundsBottom.Location = New System.Drawing.Point(0, 390)
        Me.pnlRefundsBottom.Name = "pnlRefundsBottom"
        Me.pnlRefundsBottom.Size = New System.Drawing.Size(976, 40)
        Me.pnlRefundsBottom.TabIndex = 1
        '
        'btnChangeWarrant
        '
        Me.btnChangeWarrant.Location = New System.Drawing.Point(824, 8)
        Me.btnChangeWarrant.Name = "btnChangeWarrant"
        Me.btnChangeWarrant.Size = New System.Drawing.Size(112, 23)
        Me.btnChangeWarrant.TabIndex = 4
        Me.btnChangeWarrant.Text = "Change Warrant"
        Me.btnChangeWarrant.Visible = False
        '
        'btnEnterWarrant
        '
        Me.btnEnterWarrant.Location = New System.Drawing.Point(720, 8)
        Me.btnEnterWarrant.Name = "btnEnterWarrant"
        Me.btnEnterWarrant.Size = New System.Drawing.Size(96, 23)
        Me.btnEnterWarrant.TabIndex = 3
        Me.btnEnterWarrant.Text = "Enter Warrant"
        Me.btnEnterWarrant.Visible = False
        '
        'btnGenerateRefund
        '
        Me.btnGenerateRefund.Location = New System.Drawing.Point(416, 8)
        Me.btnGenerateRefund.Name = "btnGenerateRefund"
        Me.btnGenerateRefund.Size = New System.Drawing.Size(112, 23)
        Me.btnGenerateRefund.TabIndex = 2
        Me.btnGenerateRefund.Text = "Generate Refund"
        '
        'tbPageReceipts
        '
        Me.tbPageReceipts.Controls.Add(Me.pnlReceiptsDetails)
        Me.tbPageReceipts.Controls.Add(Me.pnlReceiptsBottom)
        Me.tbPageReceipts.Location = New System.Drawing.Point(4, 24)
        Me.tbPageReceipts.Name = "tbPageReceipts"
        Me.tbPageReceipts.Size = New System.Drawing.Size(976, 430)
        Me.tbPageReceipts.TabIndex = 2
        Me.tbPageReceipts.Text = "Receipts"
        '
        'pnlReceiptsDetails
        '
        Me.pnlReceiptsDetails.Controls.Add(Me.pnlOwnerReceiptFill)
        Me.pnlReceiptsDetails.Controls.Add(Me.pnlOwnerReceiptTop)
        Me.pnlReceiptsDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlReceiptsDetails.Location = New System.Drawing.Point(0, 0)
        Me.pnlReceiptsDetails.Name = "pnlReceiptsDetails"
        Me.pnlReceiptsDetails.Size = New System.Drawing.Size(976, 390)
        Me.pnlReceiptsDetails.TabIndex = 1
        '
        'pnlOwnerReceiptFill
        '
        Me.pnlOwnerReceiptFill.Controls.Add(Me.ugReceipts)
        Me.pnlOwnerReceiptFill.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlOwnerReceiptFill.Location = New System.Drawing.Point(0, 24)
        Me.pnlOwnerReceiptFill.Name = "pnlOwnerReceiptFill"
        Me.pnlOwnerReceiptFill.Size = New System.Drawing.Size(976, 366)
        Me.pnlOwnerReceiptFill.TabIndex = 7
        '
        'ugReceipts
        '
        Me.ugReceipts.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugReceipts.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle
        Me.ugReceipts.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugReceipts.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugReceipts.Location = New System.Drawing.Point(0, 0)
        Me.ugReceipts.Name = "ugReceipts"
        Me.ugReceipts.Size = New System.Drawing.Size(976, 366)
        Me.ugReceipts.TabIndex = 1
        '
        'pnlOwnerReceiptTop
        '
        Me.pnlOwnerReceiptTop.Controls.Add(Me.rdAllFY)
        Me.pnlOwnerReceiptTop.Controls.Add(Me.rdCurrentFY)
        Me.pnlOwnerReceiptTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlOwnerReceiptTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlOwnerReceiptTop.Name = "pnlOwnerReceiptTop"
        Me.pnlOwnerReceiptTop.Size = New System.Drawing.Size(976, 24)
        Me.pnlOwnerReceiptTop.TabIndex = 5
        '
        'rdAllFY
        '
        Me.rdAllFY.Location = New System.Drawing.Point(192, 2)
        Me.rdAllFY.Name = "rdAllFY"
        Me.rdAllFY.Size = New System.Drawing.Size(144, 20)
        Me.rdAllFY.TabIndex = 5
        Me.rdAllFY.Text = "All Fiscal Years"
        '
        'rdCurrentFY
        '
        Me.rdCurrentFY.Checked = True
        Me.rdCurrentFY.Location = New System.Drawing.Point(8, 2)
        Me.rdCurrentFY.Name = "rdCurrentFY"
        Me.rdCurrentFY.Size = New System.Drawing.Size(176, 20)
        Me.rdCurrentFY.TabIndex = 4
        Me.rdCurrentFY.TabStop = True
        Me.rdCurrentFY.Text = "Current Fiscal Year Only"
        '
        'pnlReceiptsBottom
        '
        Me.pnlReceiptsBottom.Controls.Add(Me.btnAdjustPayment)
        Me.pnlReceiptsBottom.Controls.Add(Me.btnOverpaymentReason)
        Me.pnlReceiptsBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlReceiptsBottom.Location = New System.Drawing.Point(0, 390)
        Me.pnlReceiptsBottom.Name = "pnlReceiptsBottom"
        Me.pnlReceiptsBottom.Size = New System.Drawing.Size(976, 40)
        Me.pnlReceiptsBottom.TabIndex = 1
        '
        'btnAdjustPayment
        '
        Me.btnAdjustPayment.Location = New System.Drawing.Point(632, 8)
        Me.btnAdjustPayment.Name = "btnAdjustPayment"
        Me.btnAdjustPayment.Size = New System.Drawing.Size(104, 23)
        Me.btnAdjustPayment.TabIndex = 3
        Me.btnAdjustPayment.Text = "Adjust Payment"
        Me.btnAdjustPayment.Visible = False
        '
        'btnOverpaymentReason
        '
        Me.btnOverpaymentReason.Location = New System.Drawing.Point(388, 8)
        Me.btnOverpaymentReason.Name = "btnOverpaymentReason"
        Me.btnOverpaymentReason.Size = New System.Drawing.Size(168, 23)
        Me.btnOverpaymentReason.TabIndex = 2
        Me.btnOverpaymentReason.Text = "Overpayment Reason Entry"
        '
        'tbPageOverages
        '
        Me.tbPageOverages.Controls.Add(Me.ugOverage)
        Me.tbPageOverages.Location = New System.Drawing.Point(4, 24)
        Me.tbPageOverages.Name = "tbPageOverages"
        Me.tbPageOverages.Size = New System.Drawing.Size(976, 430)
        Me.tbPageOverages.TabIndex = 6
        Me.tbPageOverages.Text = "Overages"
        '
        'ugOverage
        '
        Me.ugOverage.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugOverage.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugOverage.Location = New System.Drawing.Point(0, 0)
        Me.ugOverage.Name = "ugOverage"
        Me.ugOverage.Size = New System.Drawing.Size(976, 430)
        Me.ugOverage.TabIndex = 0
        Me.ugOverage.Text = "ugOverage"
        '
        'tbPageOwnerDocuments
        '
        Me.tbPageOwnerDocuments.Controls.Add(Me.UCOwnerDocuments)
        Me.tbPageOwnerDocuments.Location = New System.Drawing.Point(4, 24)
        Me.tbPageOwnerDocuments.Name = "tbPageOwnerDocuments"
        Me.tbPageOwnerDocuments.Size = New System.Drawing.Size(976, 430)
        Me.tbPageOwnerDocuments.TabIndex = 4
        Me.tbPageOwnerDocuments.Text = "Documents"
        '
        'UCOwnerDocuments
        '
        Me.UCOwnerDocuments.AutoScroll = True
        Me.UCOwnerDocuments.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UCOwnerDocuments.Location = New System.Drawing.Point(0, 0)
        Me.UCOwnerDocuments.Name = "UCOwnerDocuments"
        Me.UCOwnerDocuments.Size = New System.Drawing.Size(976, 430)
        Me.UCOwnerDocuments.TabIndex = 3
        '
        'tbPageOwnerContactList
        '
        Me.tbPageOwnerContactList.Controls.Add(Me.pnlOwnerContactContainer)
        Me.tbPageOwnerContactList.Controls.Add(Me.pnlOwnerContactButtons)
        Me.tbPageOwnerContactList.Controls.Add(Me.pnlOwnerContactHeader)
        Me.tbPageOwnerContactList.Location = New System.Drawing.Point(4, 24)
        Me.tbPageOwnerContactList.Name = "tbPageOwnerContactList"
        Me.tbPageOwnerContactList.Size = New System.Drawing.Size(976, 430)
        Me.tbPageOwnerContactList.TabIndex = 5
        Me.tbPageOwnerContactList.Text = "Contacts"
        '
        'pnlOwnerContactContainer
        '
        Me.pnlOwnerContactContainer.Controls.Add(Me.ugOwnerContacts)
        Me.pnlOwnerContactContainer.Controls.Add(Me.Label3)
        Me.pnlOwnerContactContainer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlOwnerContactContainer.Location = New System.Drawing.Point(0, 25)
        Me.pnlOwnerContactContainer.Name = "pnlOwnerContactContainer"
        Me.pnlOwnerContactContainer.Size = New System.Drawing.Size(976, 375)
        Me.pnlOwnerContactContainer.TabIndex = 5
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
        Me.ugOwnerContacts.Size = New System.Drawing.Size(976, 375)
        Me.ugOwnerContacts.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(792, 56)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(7, 23)
        Me.Label3.TabIndex = 2
        '
        'pnlOwnerContactButtons
        '
        Me.pnlOwnerContactButtons.Controls.Add(Me.btnOwnerModifyContact)
        Me.pnlOwnerContactButtons.Controls.Add(Me.btnOwnerDeleteContact)
        Me.pnlOwnerContactButtons.Controls.Add(Me.btnOwnerAssociateContact)
        Me.pnlOwnerContactButtons.Controls.Add(Me.btnOwnerAddSearchContact)
        Me.pnlOwnerContactButtons.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlOwnerContactButtons.DockPadding.All = 3
        Me.pnlOwnerContactButtons.Location = New System.Drawing.Point(0, 400)
        Me.pnlOwnerContactButtons.Name = "pnlOwnerContactButtons"
        Me.pnlOwnerContactButtons.Size = New System.Drawing.Size(976, 30)
        Me.pnlOwnerContactButtons.TabIndex = 4
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
        Me.pnlOwnerContactHeader.Size = New System.Drawing.Size(976, 25)
        Me.pnlOwnerContactHeader.TabIndex = 2
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
        'tbPageInvoices
        '
        Me.tbPageInvoices.Controls.Add(Me.pnlManualInvoice)
        Me.tbPageInvoices.Controls.Add(Me.pnlInvoices)
        Me.tbPageInvoices.Location = New System.Drawing.Point(4, 24)
        Me.tbPageInvoices.Name = "tbPageInvoices"
        Me.tbPageInvoices.Size = New System.Drawing.Size(976, 430)
        Me.tbPageInvoices.TabIndex = 1
        Me.tbPageInvoices.Text = "Invoices"
        '
        'pnlManualInvoice
        '
        Me.pnlManualInvoice.Controls.Add(Me.pnlManualInvoiceDetails)
        Me.pnlManualInvoice.Controls.Add(Me.pnlManualInvoiceBottom)
        Me.pnlManualInvoice.Controls.Add(Me.pnManualInvoiceTop)
        Me.pnlManualInvoice.Location = New System.Drawing.Point(0, 0)
        Me.pnlManualInvoice.Name = "pnlManualInvoice"
        Me.pnlManualInvoice.Size = New System.Drawing.Size(976, 224)
        Me.pnlManualInvoice.TabIndex = 4
        Me.pnlManualInvoice.Visible = False
        '
        'pnlManualInvoiceDetails
        '
        Me.pnlManualInvoiceDetails.Controls.Add(Me.ugManualInvoices)
        Me.pnlManualInvoiceDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlManualInvoiceDetails.Location = New System.Drawing.Point(0, 72)
        Me.pnlManualInvoiceDetails.Name = "pnlManualInvoiceDetails"
        Me.pnlManualInvoiceDetails.Size = New System.Drawing.Size(976, 112)
        Me.pnlManualInvoiceDetails.TabIndex = 2
        '
        'ugManualInvoices
        '
        Me.ugManualInvoices.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugManualInvoices.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle
        Me.ugManualInvoices.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugManualInvoices.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugManualInvoices.Location = New System.Drawing.Point(0, 0)
        Me.ugManualInvoices.Name = "ugManualInvoices"
        Me.ugManualInvoices.Size = New System.Drawing.Size(976, 112)
        Me.ugManualInvoices.TabIndex = 0
        '
        'pnlManualInvoiceBottom
        '
        Me.pnlManualInvoiceBottom.Controls.Add(Me.btnManualInvoiceCancel)
        Me.pnlManualInvoiceBottom.Controls.Add(Me.btnIssueInvoice)
        Me.pnlManualInvoiceBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlManualInvoiceBottom.Location = New System.Drawing.Point(0, 184)
        Me.pnlManualInvoiceBottom.Name = "pnlManualInvoiceBottom"
        Me.pnlManualInvoiceBottom.Size = New System.Drawing.Size(976, 40)
        Me.pnlManualInvoiceBottom.TabIndex = 1
        '
        'btnManualInvoiceCancel
        '
        Me.btnManualInvoiceCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnManualInvoiceCancel.Location = New System.Drawing.Point(476, 8)
        Me.btnManualInvoiceCancel.Name = "btnManualInvoiceCancel"
        Me.btnManualInvoiceCancel.Size = New System.Drawing.Size(88, 23)
        Me.btnManualInvoiceCancel.TabIndex = 4
        Me.btnManualInvoiceCancel.Text = "Cancel"
        '
        'btnIssueInvoice
        '
        Me.btnIssueInvoice.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnIssueInvoice.Location = New System.Drawing.Point(348, 8)
        Me.btnIssueInvoice.Name = "btnIssueInvoice"
        Me.btnIssueInvoice.Size = New System.Drawing.Size(112, 23)
        Me.btnIssueInvoice.TabIndex = 3
        Me.btnIssueInvoice.Text = "Issue Invoice"
        '
        'pnManualInvoiceTop
        '
        Me.pnManualInvoiceTop.Controls.Add(Me.lblMIOwnerNameValue)
        Me.pnManualInvoiceTop.Controls.Add(Me.lblMIOwnerName)
        Me.pnManualInvoiceTop.Controls.Add(Me.lblMIOwnerIDValue)
        Me.pnManualInvoiceTop.Controls.Add(Me.lblMIOwnerID)
        Me.pnManualInvoiceTop.Controls.Add(Me.lblFeeType)
        Me.pnManualInvoiceTop.Controls.Add(Me.cmbFeeType)
        Me.pnManualInvoiceTop.Controls.Add(Me.lblInvoiceAmountValue)
        Me.pnManualInvoiceTop.Controls.Add(Me.lblInvoiceAmount)
        Me.pnManualInvoiceTop.Controls.Add(Me.lblTotalTanksValue)
        Me.pnManualInvoiceTop.Controls.Add(Me.lblTotalTanks)
        Me.pnManualInvoiceTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnManualInvoiceTop.Location = New System.Drawing.Point(0, 0)
        Me.pnManualInvoiceTop.Name = "pnManualInvoiceTop"
        Me.pnManualInvoiceTop.Size = New System.Drawing.Size(976, 72)
        Me.pnManualInvoiceTop.TabIndex = 0
        '
        'lblMIOwnerNameValue
        '
        Me.lblMIOwnerNameValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblMIOwnerNameValue.Location = New System.Drawing.Point(112, 32)
        Me.lblMIOwnerNameValue.Name = "lblMIOwnerNameValue"
        Me.lblMIOwnerNameValue.Size = New System.Drawing.Size(272, 23)
        Me.lblMIOwnerNameValue.TabIndex = 9
        '
        'lblMIOwnerName
        '
        Me.lblMIOwnerName.Location = New System.Drawing.Point(8, 32)
        Me.lblMIOwnerName.Name = "lblMIOwnerName"
        Me.lblMIOwnerName.TabIndex = 8
        Me.lblMIOwnerName.Text = "Owner Name"
        '
        'lblMIOwnerIDValue
        '
        Me.lblMIOwnerIDValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblMIOwnerIDValue.Location = New System.Drawing.Point(112, 8)
        Me.lblMIOwnerIDValue.Name = "lblMIOwnerIDValue"
        Me.lblMIOwnerIDValue.TabIndex = 7
        '
        'lblMIOwnerID
        '
        Me.lblMIOwnerID.Location = New System.Drawing.Point(8, 8)
        Me.lblMIOwnerID.Name = "lblMIOwnerID"
        Me.lblMIOwnerID.TabIndex = 6
        Me.lblMIOwnerID.Text = "Owner ID"
        '
        'lblFeeType
        '
        Me.lblFeeType.Location = New System.Drawing.Point(312, 8)
        Me.lblFeeType.Name = "lblFeeType"
        Me.lblFeeType.Size = New System.Drawing.Size(72, 23)
        Me.lblFeeType.TabIndex = 5
        Me.lblFeeType.Text = "Fee Type"
        '
        'cmbFeeType
        '
        Me.cmbFeeType.Items.AddRange(New Object() {"Minor Billing", "Miscellaneous Invoice"})
        Me.cmbFeeType.Location = New System.Drawing.Point(392, 8)
        Me.cmbFeeType.Name = "cmbFeeType"
        Me.cmbFeeType.Size = New System.Drawing.Size(200, 23)
        Me.cmbFeeType.TabIndex = 4
        Me.cmbFeeType.Text = "ComboBox1"
        '
        'lblInvoiceAmountValue
        '
        Me.lblInvoiceAmountValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblInvoiceAmountValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInvoiceAmountValue.Location = New System.Drawing.Point(808, 32)
        Me.lblInvoiceAmountValue.Name = "lblInvoiceAmountValue"
        Me.lblInvoiceAmountValue.Size = New System.Drawing.Size(100, 24)
        Me.lblInvoiceAmountValue.TabIndex = 3
        Me.lblInvoiceAmountValue.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblInvoiceAmount
        '
        Me.lblInvoiceAmount.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblInvoiceAmount.Location = New System.Drawing.Point(704, 32)
        Me.lblInvoiceAmount.Name = "lblInvoiceAmount"
        Me.lblInvoiceAmount.Size = New System.Drawing.Size(96, 24)
        Me.lblInvoiceAmount.TabIndex = 2
        Me.lblInvoiceAmount.Text = "Invoice Amount"
        Me.lblInvoiceAmount.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblTotalTanksValue
        '
        Me.lblTotalTanksValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalTanksValue.Location = New System.Drawing.Point(808, 7)
        Me.lblTotalTanksValue.Name = "lblTotalTanksValue"
        Me.lblTotalTanksValue.Size = New System.Drawing.Size(100, 24)
        Me.lblTotalTanksValue.TabIndex = 1
        Me.lblTotalTanksValue.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblTotalTanks
        '
        Me.lblTotalTanks.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblTotalTanks.Location = New System.Drawing.Point(704, 7)
        Me.lblTotalTanks.Name = "lblTotalTanks"
        Me.lblTotalTanks.Size = New System.Drawing.Size(96, 24)
        Me.lblTotalTanks.TabIndex = 0
        Me.lblTotalTanks.Text = "Total Tanks"
        Me.lblTotalTanks.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlInvoices
        '
        Me.pnlInvoices.Controls.Add(Me.pnlInvoicesDetails)
        Me.pnlInvoices.Controls.Add(Me.pnlInvoicesControls)
        Me.pnlInvoices.Controls.Add(Me.pnlInvoicesBottom)
        Me.pnlInvoices.Controls.Add(Me.pnlInvoicesTop)
        Me.pnlInvoices.Location = New System.Drawing.Point(0, 232)
        Me.pnlInvoices.Name = "pnlInvoices"
        Me.pnlInvoices.Size = New System.Drawing.Size(976, 200)
        Me.pnlInvoices.TabIndex = 3
        '
        'pnlInvoicesDetails
        '
        Me.pnlInvoicesDetails.Controls.Add(Me.ugInvoices)
        Me.pnlInvoicesDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlInvoicesDetails.Location = New System.Drawing.Point(0, 24)
        Me.pnlInvoicesDetails.Name = "pnlInvoicesDetails"
        Me.pnlInvoicesDetails.Size = New System.Drawing.Size(976, 104)
        Me.pnlInvoicesDetails.TabIndex = 4
        '
        'ugInvoices
        '
        Me.ugInvoices.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugInvoices.DisplayLayout.Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Free
        Me.ugInvoices.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle
        Me.ugInvoices.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugInvoices.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugInvoices.Location = New System.Drawing.Point(0, 0)
        Me.ugInvoices.Name = "ugInvoices"
        Me.ugInvoices.Size = New System.Drawing.Size(976, 104)
        Me.ugInvoices.TabIndex = 4
        '
        'pnlInvoicesControls
        '
        Me.pnlInvoicesControls.Controls.Add(Me.btnGenerateDebitMemo)
        Me.pnlInvoicesControls.Controls.Add(Me.btnRecordInvoiceNoDate)
        Me.pnlInvoicesControls.Controls.Add(Me.btnReqLateFeeWaiver)
        Me.pnlInvoicesControls.Controls.Add(Me.btnReallocateOverage)
        Me.pnlInvoicesControls.Controls.Add(Me.btnGenCreditMemo)
        Me.pnlInvoicesControls.Controls.Add(Me.btnGenerateInvoice)
        Me.pnlInvoicesControls.Controls.Add(Me.lblTotalOverage)
        Me.pnlInvoicesControls.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlInvoicesControls.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlInvoicesControls.Location = New System.Drawing.Point(0, 128)
        Me.pnlInvoicesControls.Name = "pnlInvoicesControls"
        Me.pnlInvoicesControls.Size = New System.Drawing.Size(976, 48)
        Me.pnlInvoicesControls.TabIndex = 5
        '
        'btnGenerateDebitMemo
        '
        Me.btnGenerateDebitMemo.Location = New System.Drawing.Point(412, 8)
        Me.btnGenerateDebitMemo.Name = "btnGenerateDebitMemo"
        Me.btnGenerateDebitMemo.Size = New System.Drawing.Size(80, 33)
        Me.btnGenerateDebitMemo.TabIndex = 8
        Me.btnGenerateDebitMemo.Text = "Generate Debit Memo"
        '
        'btnRecordInvoiceNoDate
        '
        Me.btnRecordInvoiceNoDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRecordInvoiceNoDate.Location = New System.Drawing.Point(32, 8)
        Me.btnRecordInvoiceNoDate.Name = "btnRecordInvoiceNoDate"
        Me.btnRecordInvoiceNoDate.Size = New System.Drawing.Size(88, 33)
        Me.btnRecordInvoiceNoDate.TabIndex = 11
        Me.btnRecordInvoiceNoDate.Text = "Record Invoice Number/Date"
        Me.btnRecordInvoiceNoDate.Visible = False
        '
        'btnReqLateFeeWaiver
        '
        Me.btnReqLateFeeWaiver.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnReqLateFeeWaiver.Location = New System.Drawing.Point(580, 8)
        Me.btnReqLateFeeWaiver.Name = "btnReqLateFeeWaiver"
        Me.btnReqLateFeeWaiver.Size = New System.Drawing.Size(88, 33)
        Me.btnReqLateFeeWaiver.TabIndex = 10
        Me.btnReqLateFeeWaiver.Text = "Request Late Fee Waiver"
        '
        'btnReallocateOverage
        '
        Me.btnReallocateOverage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnReallocateOverage.Location = New System.Drawing.Point(500, 8)
        Me.btnReallocateOverage.Name = "btnReallocateOverage"
        Me.btnReallocateOverage.Size = New System.Drawing.Size(72, 33)
        Me.btnReallocateOverage.TabIndex = 9
        Me.btnReallocateOverage.Text = "Reallocate Overage"
        '
        'btnGenCreditMemo
        '
        Me.btnGenCreditMemo.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGenCreditMemo.Location = New System.Drawing.Point(324, 8)
        Me.btnGenCreditMemo.Name = "btnGenCreditMemo"
        Me.btnGenCreditMemo.Size = New System.Drawing.Size(80, 33)
        Me.btnGenCreditMemo.TabIndex = 7
        Me.btnGenCreditMemo.Text = "Generate Credit Memo"
        '
        'btnGenerateInvoice
        '
        Me.btnGenerateInvoice.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGenerateInvoice.Location = New System.Drawing.Point(244, 8)
        Me.btnGenerateInvoice.Name = "btnGenerateInvoice"
        Me.btnGenerateInvoice.Size = New System.Drawing.Size(72, 33)
        Me.btnGenerateInvoice.TabIndex = 6
        Me.btnGenerateInvoice.Text = "Generate Invoice"
        '
        'lblTotalOverage
        '
        Me.lblTotalOverage.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalOverage.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTotalOverage.Location = New System.Drawing.Point(868, 24)
        Me.lblTotalOverage.Name = "lblTotalOverage"
        Me.lblTotalOverage.Size = New System.Drawing.Size(80, 17)
        Me.lblTotalOverage.TabIndex = 5
        Me.lblTotalOverage.Text = "Overage"
        Me.lblTotalOverage.UseMnemonic = False
        Me.lblTotalOverage.Visible = False
        '
        'pnlInvoicesBottom
        '
        Me.pnlInvoicesBottom.Controls.Add(Me.lblTotalOverageValue)
        Me.pnlInvoicesBottom.Controls.Add(Me.lblTotalLegal)
        Me.pnlInvoicesBottom.Controls.Add(Me.lblTotalBalance)
        Me.pnlInvoicesBottom.Controls.Add(Me.lblTotalAdjustments)
        Me.pnlInvoicesBottom.Controls.Add(Me.lblTotalCredits)
        Me.pnlInvoicesBottom.Controls.Add(Me.lblTotalPayments)
        Me.pnlInvoicesBottom.Controls.Add(Me.lblTotalCharges)
        Me.pnlInvoicesBottom.Controls.Add(Me.lblInvoices)
        Me.pnlInvoicesBottom.Controls.Add(Me.lblTotalInvoices)
        Me.pnlInvoicesBottom.Controls.Add(Me.lblOwnerTotalsfor)
        Me.pnlInvoicesBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlInvoicesBottom.Location = New System.Drawing.Point(0, 176)
        Me.pnlInvoicesBottom.Name = "pnlInvoicesBottom"
        Me.pnlInvoicesBottom.Size = New System.Drawing.Size(976, 24)
        Me.pnlInvoicesBottom.TabIndex = 2
        '
        'lblTotalOverageValue
        '
        Me.lblTotalOverageValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTotalOverageValue.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblTotalOverageValue.Location = New System.Drawing.Point(949, 0)
        Me.lblTotalOverageValue.Name = "lblTotalOverageValue"
        Me.lblTotalOverageValue.Size = New System.Drawing.Size(75, 24)
        Me.lblTotalOverageValue.TabIndex = 15
        Me.lblTotalOverageValue.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.lblTotalOverageValue.UseMnemonic = False
        Me.lblTotalOverageValue.Visible = False
        '
        'lblTotalLegal
        '
        Me.lblTotalLegal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalLegal.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblTotalLegal.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblTotalLegal.Location = New System.Drawing.Point(874, 0)
        Me.lblTotalLegal.Name = "lblTotalLegal"
        Me.lblTotalLegal.Size = New System.Drawing.Size(75, 24)
        Me.lblTotalLegal.TabIndex = 16
        Me.lblTotalLegal.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTotalBalance
        '
        Me.lblTotalBalance.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalBalance.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblTotalBalance.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblTotalBalance.Location = New System.Drawing.Point(799, 0)
        Me.lblTotalBalance.Name = "lblTotalBalance"
        Me.lblTotalBalance.Size = New System.Drawing.Size(75, 24)
        Me.lblTotalBalance.TabIndex = 17
        Me.lblTotalBalance.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTotalAdjustments
        '
        Me.lblTotalAdjustments.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalAdjustments.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblTotalAdjustments.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblTotalAdjustments.Location = New System.Drawing.Point(724, 0)
        Me.lblTotalAdjustments.Name = "lblTotalAdjustments"
        Me.lblTotalAdjustments.Size = New System.Drawing.Size(75, 24)
        Me.lblTotalAdjustments.TabIndex = 14
        Me.lblTotalAdjustments.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTotalCredits
        '
        Me.lblTotalCredits.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalCredits.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblTotalCredits.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblTotalCredits.Location = New System.Drawing.Point(649, 0)
        Me.lblTotalCredits.Name = "lblTotalCredits"
        Me.lblTotalCredits.Size = New System.Drawing.Size(75, 24)
        Me.lblTotalCredits.TabIndex = 13
        Me.lblTotalCredits.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTotalPayments
        '
        Me.lblTotalPayments.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalPayments.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblTotalPayments.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblTotalPayments.Location = New System.Drawing.Point(574, 0)
        Me.lblTotalPayments.Name = "lblTotalPayments"
        Me.lblTotalPayments.Size = New System.Drawing.Size(75, 24)
        Me.lblTotalPayments.TabIndex = 12
        Me.lblTotalPayments.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblTotalCharges
        '
        Me.lblTotalCharges.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalCharges.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblTotalCharges.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblTotalCharges.Location = New System.Drawing.Point(499, 0)
        Me.lblTotalCharges.Name = "lblTotalCharges"
        Me.lblTotalCharges.Size = New System.Drawing.Size(75, 24)
        Me.lblTotalCharges.TabIndex = 11
        Me.lblTotalCharges.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblInvoices
        '
        Me.lblInvoices.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblInvoices.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblInvoices.Location = New System.Drawing.Point(415, 0)
        Me.lblInvoices.Name = "lblInvoices"
        Me.lblInvoices.Size = New System.Drawing.Size(84, 24)
        Me.lblInvoices.TabIndex = 10
        Me.lblInvoices.Text = "Invoices"
        '
        'lblTotalInvoices
        '
        Me.lblTotalInvoices.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalInvoices.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblTotalInvoices.Location = New System.Drawing.Point(375, 0)
        Me.lblTotalInvoices.Name = "lblTotalInvoices"
        Me.lblTotalInvoices.Size = New System.Drawing.Size(40, 24)
        Me.lblTotalInvoices.TabIndex = 9
        Me.lblTotalInvoices.Text = "0"
        '
        'lblOwnerTotalsfor
        '
        Me.lblOwnerTotalsfor.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblOwnerTotalsfor.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblOwnerTotalsfor.Location = New System.Drawing.Point(0, 0)
        Me.lblOwnerTotalsfor.Name = "lblOwnerTotalsfor"
        Me.lblOwnerTotalsfor.Size = New System.Drawing.Size(375, 24)
        Me.lblOwnerTotalsfor.TabIndex = 8
        Me.lblOwnerTotalsfor.Text = "Owner Totals for:"
        Me.lblOwnerTotalsfor.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'pnlInvoicesTop
        '
        Me.pnlInvoicesTop.Controls.Add(Me.rdbtnInvoiceCurrentFY)
        Me.pnlInvoicesTop.Controls.Add(Me.rdBtnAllInvoices)
        Me.pnlInvoicesTop.Controls.Add(Me.rdBtnLateInvoices)
        Me.pnlInvoicesTop.Controls.Add(Me.rdBtnOutstandingBalance)
        Me.pnlInvoicesTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlInvoicesTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlInvoicesTop.Name = "pnlInvoicesTop"
        Me.pnlInvoicesTop.Size = New System.Drawing.Size(976, 24)
        Me.pnlInvoicesTop.TabIndex = 0
        '
        'rdbtnInvoiceCurrentFY
        '
        Me.rdbtnInvoiceCurrentFY.Location = New System.Drawing.Point(608, 2)
        Me.rdbtnInvoiceCurrentFY.Name = "rdbtnInvoiceCurrentFY"
        Me.rdbtnInvoiceCurrentFY.Size = New System.Drawing.Size(160, 20)
        Me.rdbtnInvoiceCurrentFY.TabIndex = 4
        Me.rdbtnInvoiceCurrentFY.Text = "Current Fiscal Year Only"
        '
        'rdBtnAllInvoices
        '
        Me.rdBtnAllInvoices.Location = New System.Drawing.Point(432, 2)
        Me.rdBtnAllInvoices.Name = "rdBtnAllInvoices"
        Me.rdBtnAllInvoices.Size = New System.Drawing.Size(126, 20)
        Me.rdBtnAllInvoices.TabIndex = 3
        Me.rdBtnAllInvoices.Text = "All Invoices"
        '
        'rdBtnLateInvoices
        '
        Me.rdBtnLateInvoices.Location = New System.Drawing.Point(232, 2)
        Me.rdBtnLateInvoices.Name = "rdBtnLateInvoices"
        Me.rdBtnLateInvoices.Size = New System.Drawing.Size(126, 20)
        Me.rdBtnLateInvoices.TabIndex = 2
        Me.rdBtnLateInvoices.Text = "Late Invoices Only"
        '
        'rdBtnOutstandingBalance
        '
        Me.rdBtnOutstandingBalance.Checked = True
        Me.rdBtnOutstandingBalance.Location = New System.Drawing.Point(8, 2)
        Me.rdBtnOutstandingBalance.Name = "rdBtnOutstandingBalance"
        Me.rdBtnOutstandingBalance.Size = New System.Drawing.Size(168, 20)
        Me.rdBtnOutstandingBalance.TabIndex = 1
        Me.rdBtnOutstandingBalance.TabStop = True
        Me.rdBtnOutstandingBalance.Text = "Outstanding Balance Only"
        '
        'pnlOwnerDetails
        '
        Me.pnlOwnerDetails.Controls.Add(Me.Label2)
        Me.pnlOwnerDetails.Controls.Add(Me.txtBP2KOwnerID)
        Me.pnlOwnerDetails.Controls.Add(Me.Label1)
        Me.pnlOwnerDetails.Controls.Add(Me.btnFeesOwnerLabels)
        Me.pnlOwnerDetails.Controls.Add(Me.btnFeesOwnerEnvelopes)
        Me.pnlOwnerDetails.Controls.Add(Me.pnlOwnerButtons)
        Me.pnlOwnerDetails.Controls.Add(Me.txtDate)
        Me.pnlOwnerDetails.Controls.Add(Me.txtChapter)
        Me.pnlOwnerDetails.Controls.Add(Me.lblDate)
        Me.pnlOwnerDetails.Controls.Add(Me.lblCapter)
        Me.pnlOwnerDetails.Controls.Add(Me.lblBankruptcy)
        Me.pnlOwnerDetails.Controls.Add(Me.lblFinal)
        Me.pnlOwnerDetails.Controls.Add(Me.txtFinal)
        Me.pnlOwnerDetails.Controls.Add(Me.txtOverage)
        Me.pnlOwnerDetails.Controls.Add(Me.lblOverage)
        Me.pnlOwnerDetails.Controls.Add(Me.txtBalance)
        Me.pnlOwnerDetails.Controls.Add(Me.lblBalance)
        Me.pnlOwnerDetails.Controls.Add(Me.lblOwnerFeeTotals)
        Me.pnlOwnerDetails.Controls.Add(Me.lblOwnerEmail)
        Me.pnlOwnerDetails.Controls.Add(Me.txtOwnerEmail)
        Me.pnlOwnerDetails.Controls.Add(Me.mskTxtOwnerFax)
        Me.pnlOwnerDetails.Controls.Add(Me.mskTxtOwnerPhone2)
        Me.pnlOwnerDetails.Controls.Add(Me.mskTxtOwnerPhone)
        Me.pnlOwnerDetails.Controls.Add(Me.lblFax)
        Me.pnlOwnerDetails.Controls.Add(Me.lblOwnerActiveOrNot)
        Me.pnlOwnerDetails.Controls.Add(Me.txtOwnerAIID)
        Me.pnlOwnerDetails.Controls.Add(Me.lblPhone2)
        Me.pnlOwnerDetails.Controls.Add(Me.lblPhone)
        Me.pnlOwnerDetails.Controls.Add(Me.lblEnsiteID)
        Me.pnlOwnerDetails.Controls.Add(Me.lblOwnerStatus)
        Me.pnlOwnerDetails.Controls.Add(Me.txtOwnerAddress)
        Me.pnlOwnerDetails.Controls.Add(Me.lblAddress)
        Me.pnlOwnerDetails.Controls.Add(Me.txtOwnerName)
        Me.pnlOwnerDetails.Controls.Add(Me.lblName)
        Me.pnlOwnerDetails.Controls.Add(Me.cmbOwnerType)
        Me.pnlOwnerDetails.Controls.Add(Me.lblOwnerType)
        Me.pnlOwnerDetails.Controls.Add(Me.lblOwnerIDValue)
        Me.pnlOwnerDetails.Controls.Add(Me.lblOwnerID)
        Me.pnlOwnerDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlOwnerDetails.Location = New System.Drawing.Point(0, 0)
        Me.pnlOwnerDetails.Name = "pnlOwnerDetails"
        Me.pnlOwnerDetails.Size = New System.Drawing.Size(984, 184)
        Me.pnlOwnerDetails.TabIndex = 0
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(728, 128)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(72, 16)
        Me.Label2.TabIndex = 1071
        Me.Label2.Text = "BP2K"
        '
        'txtBP2KOwnerID
        '
        Me.txtBP2KOwnerID.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtBP2KOwnerID.Location = New System.Drawing.Point(800, 144)
        Me.txtBP2KOwnerID.Name = "txtBP2KOwnerID"
        Me.txtBP2KOwnerID.ReadOnly = True
        Me.txtBP2KOwnerID.TabIndex = 1070
        Me.txtBP2KOwnerID.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(728, 144)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 16)
        Me.Label1.TabIndex = 1069
        Me.Label1.Text = "Owner ID"
        '
        'btnFeesOwnerLabels
        '
        Me.btnFeesOwnerLabels.Location = New System.Drawing.Point(8, 138)
        Me.btnFeesOwnerLabels.Name = "btnFeesOwnerLabels"
        Me.btnFeesOwnerLabels.Size = New System.Drawing.Size(70, 23)
        Me.btnFeesOwnerLabels.TabIndex = 1068
        Me.btnFeesOwnerLabels.Text = "Labels"
        '
        'btnFeesOwnerEnvelopes
        '
        Me.btnFeesOwnerEnvelopes.Location = New System.Drawing.Point(8, 110)
        Me.btnFeesOwnerEnvelopes.Name = "btnFeesOwnerEnvelopes"
        Me.btnFeesOwnerEnvelopes.Size = New System.Drawing.Size(70, 23)
        Me.btnFeesOwnerEnvelopes.TabIndex = 1067
        Me.btnFeesOwnerEnvelopes.Text = "Envelopes"
        '
        'pnlOwnerButtons
        '
        Me.pnlOwnerButtons.Controls.Add(Me.btnOwnerFlag)
        Me.pnlOwnerButtons.Controls.Add(Me.btnOwnerComment)
        Me.pnlOwnerButtons.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.pnlOwnerButtons.Location = New System.Drawing.Point(424, 152)
        Me.pnlOwnerButtons.Name = "pnlOwnerButtons"
        Me.pnlOwnerButtons.Size = New System.Drawing.Size(176, 31)
        Me.pnlOwnerButtons.TabIndex = 63
        '
        'btnOwnerFlag
        '
        Me.btnOwnerFlag.Location = New System.Drawing.Point(10, 3)
        Me.btnOwnerFlag.Name = "btnOwnerFlag"
        Me.btnOwnerFlag.TabIndex = 48
        Me.btnOwnerFlag.Text = "Flags"
        '
        'btnOwnerComment
        '
        Me.btnOwnerComment.Location = New System.Drawing.Point(90, 3)
        Me.btnOwnerComment.Name = "btnOwnerComment"
        Me.btnOwnerComment.Size = New System.Drawing.Size(80, 23)
        Me.btnOwnerComment.TabIndex = 47
        Me.btnOwnerComment.Text = "Comments"
        '
        'txtDate
        '
        Me.txtDate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtDate.Location = New System.Drawing.Point(800, 56)
        Me.txtDate.Name = "txtDate"
        Me.txtDate.ReadOnly = True
        Me.txtDate.TabIndex = 62
        Me.txtDate.Text = ""
        '
        'txtChapter
        '
        Me.txtChapter.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtChapter.Location = New System.Drawing.Point(800, 32)
        Me.txtChapter.Name = "txtChapter"
        Me.txtChapter.ReadOnly = True
        Me.txtChapter.TabIndex = 61
        Me.txtChapter.Text = ""
        '
        'lblDate
        '
        Me.lblDate.Location = New System.Drawing.Point(744, 56)
        Me.lblDate.Name = "lblDate"
        Me.lblDate.Size = New System.Drawing.Size(54, 17)
        Me.lblDate.TabIndex = 60
        Me.lblDate.Text = "Date:"
        '
        'lblCapter
        '
        Me.lblCapter.Location = New System.Drawing.Point(744, 32)
        Me.lblCapter.Name = "lblCapter"
        Me.lblCapter.Size = New System.Drawing.Size(54, 17)
        Me.lblCapter.TabIndex = 59
        Me.lblCapter.Text = "Chapter:"
        '
        'lblBankruptcy
        '
        Me.lblBankruptcy.Location = New System.Drawing.Point(728, 8)
        Me.lblBankruptcy.Name = "lblBankruptcy"
        Me.lblBankruptcy.Size = New System.Drawing.Size(72, 23)
        Me.lblBankruptcy.TabIndex = 58
        Me.lblBankruptcy.Text = "Bankruptcy"
        '
        'lblFinal
        '
        Me.lblFinal.Location = New System.Drawing.Point(544, 80)
        Me.lblFinal.Name = "lblFinal"
        Me.lblFinal.Size = New System.Drawing.Size(54, 17)
        Me.lblFinal.TabIndex = 57
        Me.lblFinal.Text = "Final:"
        '
        'txtFinal
        '
        Me.txtFinal.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFinal.Location = New System.Drawing.Point(616, 80)
        Me.txtFinal.Name = "txtFinal"
        Me.txtFinal.ReadOnly = True
        Me.txtFinal.TabIndex = 56
        Me.txtFinal.Text = "0.00"
        Me.txtFinal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtOverage
        '
        Me.txtOverage.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtOverage.ForeColor = System.Drawing.Color.Red
        Me.txtOverage.Location = New System.Drawing.Point(616, 56)
        Me.txtOverage.Name = "txtOverage"
        Me.txtOverage.ReadOnly = True
        Me.txtOverage.TabIndex = 55
        Me.txtOverage.Text = "0.00"
        Me.txtOverage.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblOverage
        '
        Me.lblOverage.Location = New System.Drawing.Point(544, 56)
        Me.lblOverage.Name = "lblOverage"
        Me.lblOverage.Size = New System.Drawing.Size(55, 17)
        Me.lblOverage.TabIndex = 54
        Me.lblOverage.Text = "Overage:"
        '
        'txtBalance
        '
        Me.txtBalance.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtBalance.Location = New System.Drawing.Point(616, 32)
        Me.txtBalance.Name = "txtBalance"
        Me.txtBalance.ReadOnly = True
        Me.txtBalance.TabIndex = 53
        Me.txtBalance.Text = "0.00"
        Me.txtBalance.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblBalance
        '
        Me.lblBalance.Location = New System.Drawing.Point(544, 32)
        Me.lblBalance.Name = "lblBalance"
        Me.lblBalance.Size = New System.Drawing.Size(54, 17)
        Me.lblBalance.TabIndex = 52
        Me.lblBalance.Text = "Balance:"
        '
        'lblOwnerFeeTotals
        '
        Me.lblOwnerFeeTotals.Location = New System.Drawing.Point(528, 8)
        Me.lblOwnerFeeTotals.Name = "lblOwnerFeeTotals"
        Me.lblOwnerFeeTotals.Size = New System.Drawing.Size(112, 23)
        Me.lblOwnerFeeTotals.TabIndex = 51
        Me.lblOwnerFeeTotals.Text = "Owner Fee Totals"
        '
        'lblOwnerEmail
        '
        Me.lblOwnerEmail.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerEmail.Location = New System.Drawing.Point(320, 130)
        Me.lblOwnerEmail.Name = "lblOwnerEmail"
        Me.lblOwnerEmail.Size = New System.Drawing.Size(40, 17)
        Me.lblOwnerEmail.TabIndex = 50
        Me.lblOwnerEmail.Text = "Email"
        '
        'txtOwnerEmail
        '
        Me.txtOwnerEmail.AcceptsTab = True
        Me.txtOwnerEmail.AutoSize = False
        Me.txtOwnerEmail.Enabled = False
        Me.txtOwnerEmail.Location = New System.Drawing.Point(408, 128)
        Me.txtOwnerEmail.Name = "txtOwnerEmail"
        Me.txtOwnerEmail.Size = New System.Drawing.Size(192, 21)
        Me.txtOwnerEmail.TabIndex = 49
        Me.txtOwnerEmail.Text = ""
        Me.txtOwnerEmail.WordWrap = False
        '
        'mskTxtOwnerFax
        '
        Me.mskTxtOwnerFax.ContainingControl = Me
        Me.mskTxtOwnerFax.Location = New System.Drawing.Point(408, 104)
        Me.mskTxtOwnerFax.Name = "mskTxtOwnerFax"
        Me.mskTxtOwnerFax.OcxState = CType(resources.GetObject("mskTxtOwnerFax.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtOwnerFax.Size = New System.Drawing.Size(100, 23)
        Me.mskTxtOwnerFax.TabIndex = 47
        '
        'mskTxtOwnerPhone2
        '
        Me.mskTxtOwnerPhone2.ContainingControl = Me
        Me.mskTxtOwnerPhone2.Location = New System.Drawing.Point(408, 80)
        Me.mskTxtOwnerPhone2.Name = "mskTxtOwnerPhone2"
        Me.mskTxtOwnerPhone2.OcxState = CType(resources.GetObject("mskTxtOwnerPhone2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtOwnerPhone2.Size = New System.Drawing.Size(100, 23)
        Me.mskTxtOwnerPhone2.TabIndex = 46
        '
        'mskTxtOwnerPhone
        '
        Me.mskTxtOwnerPhone.ContainingControl = Me
        Me.mskTxtOwnerPhone.Location = New System.Drawing.Point(408, 56)
        Me.mskTxtOwnerPhone.Name = "mskTxtOwnerPhone"
        Me.mskTxtOwnerPhone.OcxState = CType(resources.GetObject("mskTxtOwnerPhone.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtOwnerPhone.Size = New System.Drawing.Size(100, 23)
        Me.mskTxtOwnerPhone.TabIndex = 45
        '
        'lblFax
        '
        Me.lblFax.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFax.Location = New System.Drawing.Point(320, 106)
        Me.lblFax.Name = "lblFax"
        Me.lblFax.Size = New System.Drawing.Size(48, 17)
        Me.lblFax.TabIndex = 48
        Me.lblFax.Text = "Fax"
        '
        'lblOwnerActiveOrNot
        '
        Me.lblOwnerActiveOrNot.Location = New System.Drawing.Point(408, 8)
        Me.lblOwnerActiveOrNot.Name = "lblOwnerActiveOrNot"
        Me.lblOwnerActiveOrNot.TabIndex = 13
        '
        'txtOwnerAIID
        '
        Me.txtOwnerAIID.Enabled = False
        Me.txtOwnerAIID.Location = New System.Drawing.Point(408, 32)
        Me.txtOwnerAIID.Name = "txtOwnerAIID"
        Me.txtOwnerAIID.ReadOnly = True
        Me.txtOwnerAIID.TabIndex = 12
        Me.txtOwnerAIID.Text = ""
        '
        'lblPhone2
        '
        Me.lblPhone2.Location = New System.Drawing.Point(320, 83)
        Me.lblPhone2.Name = "lblPhone2"
        Me.lblPhone2.Size = New System.Drawing.Size(56, 17)
        Me.lblPhone2.TabIndex = 11
        Me.lblPhone2.Text = "Phone 2:"
        '
        'lblPhone
        '
        Me.lblPhone.Location = New System.Drawing.Point(320, 58)
        Me.lblPhone.Name = "lblPhone"
        Me.lblPhone.Size = New System.Drawing.Size(48, 17)
        Me.lblPhone.TabIndex = 10
        Me.lblPhone.Text = "Phone:"
        '
        'lblEnsiteID
        '
        Me.lblEnsiteID.Location = New System.Drawing.Point(320, 34)
        Me.lblEnsiteID.Name = "lblEnsiteID"
        Me.lblEnsiteID.Size = New System.Drawing.Size(64, 17)
        Me.lblEnsiteID.TabIndex = 9
        Me.lblEnsiteID.Text = "Ensite ID:"
        '
        'lblOwnerStatus
        '
        Me.lblOwnerStatus.Location = New System.Drawing.Point(320, 8)
        Me.lblOwnerStatus.Name = "lblOwnerStatus"
        Me.lblOwnerStatus.Size = New System.Drawing.Size(84, 17)
        Me.lblOwnerStatus.TabIndex = 8
        Me.lblOwnerStatus.Text = "Owner Status:"
        '
        'txtOwnerAddress
        '
        Me.txtOwnerAddress.Location = New System.Drawing.Point(83, 80)
        Me.txtOwnerAddress.Multiline = True
        Me.txtOwnerAddress.Name = "txtOwnerAddress"
        Me.txtOwnerAddress.ReadOnly = True
        Me.txtOwnerAddress.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
        Me.txtOwnerAddress.Size = New System.Drawing.Size(221, 96)
        Me.txtOwnerAddress.TabIndex = 7
        Me.txtOwnerAddress.Text = ""
        Me.txtOwnerAddress.WordWrap = False
        '
        'lblAddress
        '
        Me.lblAddress.Location = New System.Drawing.Point(8, 80)
        Me.lblAddress.Name = "lblAddress"
        Me.lblAddress.Size = New System.Drawing.Size(53, 17)
        Me.lblAddress.TabIndex = 6
        Me.lblAddress.Text = "Address:"
        '
        'txtOwnerName
        '
        Me.txtOwnerName.Location = New System.Drawing.Point(83, 56)
        Me.txtOwnerName.Name = "txtOwnerName"
        Me.txtOwnerName.ReadOnly = True
        Me.txtOwnerName.Size = New System.Drawing.Size(221, 21)
        Me.txtOwnerName.TabIndex = 5
        Me.txtOwnerName.Text = ""
        '
        'lblName
        '
        Me.lblName.Location = New System.Drawing.Point(8, 56)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(48, 17)
        Me.lblName.TabIndex = 4
        Me.lblName.Text = "Name:"
        '
        'cmbOwnerType
        '
        Me.cmbOwnerType.Enabled = False
        Me.cmbOwnerType.Location = New System.Drawing.Point(83, 32)
        Me.cmbOwnerType.Name = "cmbOwnerType"
        Me.cmbOwnerType.Size = New System.Drawing.Size(221, 23)
        Me.cmbOwnerType.TabIndex = 3
        '
        'lblOwnerType
        '
        Me.lblOwnerType.Location = New System.Drawing.Point(8, 32)
        Me.lblOwnerType.Name = "lblOwnerType"
        Me.lblOwnerType.Size = New System.Drawing.Size(75, 17)
        Me.lblOwnerType.TabIndex = 2
        Me.lblOwnerType.Text = "Owner Type:"
        '
        'lblOwnerIDValue
        '
        Me.lblOwnerIDValue.Location = New System.Drawing.Point(83, 8)
        Me.lblOwnerIDValue.Name = "lblOwnerIDValue"
        Me.lblOwnerIDValue.TabIndex = 1
        '
        'lblOwnerID
        '
        Me.lblOwnerID.Location = New System.Drawing.Point(8, 8)
        Me.lblOwnerID.Name = "lblOwnerID"
        Me.lblOwnerID.Size = New System.Drawing.Size(64, 17)
        Me.lblOwnerID.TabIndex = 0
        Me.lblOwnerID.Text = "Owner ID:"
        '
        'tbPageFacilityDetail
        '
        Me.tbPageFacilityDetail.Controls.Add(Me.pnlFacilityBottom)
        Me.tbPageFacilityDetail.Controls.Add(Me.pnl_FacilityDetail)
        Me.tbPageFacilityDetail.Location = New System.Drawing.Point(4, 24)
        Me.tbPageFacilityDetail.Name = "tbPageFacilityDetail"
        Me.tbPageFacilityDetail.Size = New System.Drawing.Size(984, 642)
        Me.tbPageFacilityDetail.TabIndex = 1
        Me.tbPageFacilityDetail.Text = "Facility Details"
        '
        'pnlFacilityBottom
        '
        Me.pnlFacilityBottom.Controls.Add(Me.tbCtrlFacTransactions)
        Me.pnlFacilityBottom.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlFacilityBottom.Location = New System.Drawing.Point(0, 184)
        Me.pnlFacilityBottom.Name = "pnlFacilityBottom"
        Me.pnlFacilityBottom.Size = New System.Drawing.Size(984, 458)
        Me.pnlFacilityBottom.TabIndex = 3
        '
        'tbCtrlFacTransactions
        '
        Me.tbCtrlFacTransactions.Controls.Add(Me.tbPageByTransaction)
        Me.tbCtrlFacTransactions.Controls.Add(Me.tbPageByInvoice)
        Me.tbCtrlFacTransactions.Controls.Add(Me.tbPageOldAccessHis)
        Me.tbCtrlFacTransactions.Controls.Add(Me.tbPageFacilityDocuments)
        Me.tbCtrlFacTransactions.Controls.Add(Me.tbPageFacilityContactList)
        Me.tbCtrlFacTransactions.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbCtrlFacTransactions.Location = New System.Drawing.Point(0, 0)
        Me.tbCtrlFacTransactions.Multiline = True
        Me.tbCtrlFacTransactions.Name = "tbCtrlFacTransactions"
        Me.tbCtrlFacTransactions.SelectedIndex = 0
        Me.tbCtrlFacTransactions.ShowToolTips = True
        Me.tbCtrlFacTransactions.Size = New System.Drawing.Size(984, 458)
        Me.tbCtrlFacTransactions.TabIndex = 0
        '
        'tbPageByTransaction
        '
        Me.tbPageByTransaction.Controls.Add(Me.pnlByTransacDetails)
        Me.tbPageByTransaction.Controls.Add(Me.pnlByTransacTop)
        Me.tbPageByTransaction.Location = New System.Drawing.Point(4, 24)
        Me.tbPageByTransaction.Name = "tbPageByTransaction"
        Me.tbPageByTransaction.Size = New System.Drawing.Size(976, 430)
        Me.tbPageByTransaction.TabIndex = 0
        Me.tbPageByTransaction.Text = "By Transaction"
        '
        'pnlByTransacDetails
        '
        Me.pnlByTransacDetails.Controls.Add(Me.ugByTransaction)
        Me.pnlByTransacDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlByTransacDetails.Location = New System.Drawing.Point(0, 24)
        Me.pnlByTransacDetails.Name = "pnlByTransacDetails"
        Me.pnlByTransacDetails.Size = New System.Drawing.Size(976, 406)
        Me.pnlByTransacDetails.TabIndex = 1
        '
        'ugByTransaction
        '
        Me.ugByTransaction.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugByTransaction.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle
        Me.ugByTransaction.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugByTransaction.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugByTransaction.Location = New System.Drawing.Point(0, 0)
        Me.ugByTransaction.Name = "ugByTransaction"
        Me.ugByTransaction.Size = New System.Drawing.Size(976, 406)
        Me.ugByTransaction.TabIndex = 3
        '
        'pnlByTransacTop
        '
        Me.pnlByTransacTop.Controls.Add(Me.rdBtnCurrentFiscalYear)
        Me.pnlByTransacTop.Controls.Add(Me.rdBtnByTransAllYears)
        Me.pnlByTransacTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlByTransacTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlByTransacTop.Name = "pnlByTransacTop"
        Me.pnlByTransacTop.Size = New System.Drawing.Size(976, 24)
        Me.pnlByTransacTop.TabIndex = 0
        '
        'rdBtnCurrentFiscalYear
        '
        Me.rdBtnCurrentFiscalYear.Location = New System.Drawing.Point(8, 3)
        Me.rdBtnCurrentFiscalYear.Name = "rdBtnCurrentFiscalYear"
        Me.rdBtnCurrentFiscalYear.Size = New System.Drawing.Size(148, 18)
        Me.rdBtnCurrentFiscalYear.TabIndex = 1
        Me.rdBtnCurrentFiscalYear.Text = "Current Fiscal Year"
        '
        'rdBtnByTransAllYears
        '
        Me.rdBtnByTransAllYears.Checked = True
        Me.rdBtnByTransAllYears.Location = New System.Drawing.Point(200, 3)
        Me.rdBtnByTransAllYears.Name = "rdBtnByTransAllYears"
        Me.rdBtnByTransAllYears.Size = New System.Drawing.Size(104, 17)
        Me.rdBtnByTransAllYears.TabIndex = 2
        Me.rdBtnByTransAllYears.TabStop = True
        Me.rdBtnByTransAllYears.Text = "All Years"
        '
        'tbPageByInvoice
        '
        Me.tbPageByInvoice.Controls.Add(Me.pnlByInvoiceDetails)
        Me.tbPageByInvoice.Controls.Add(Me.pnlByInvoiceTop)
        Me.tbPageByInvoice.Controls.Add(Me.pnlByInvoiceBottom)
        Me.tbPageByInvoice.Location = New System.Drawing.Point(4, 24)
        Me.tbPageByInvoice.Name = "tbPageByInvoice"
        Me.tbPageByInvoice.Size = New System.Drawing.Size(976, 430)
        Me.tbPageByInvoice.TabIndex = 1
        Me.tbPageByInvoice.Text = "By Invoice"
        '
        'pnlByInvoiceDetails
        '
        Me.pnlByInvoiceDetails.Controls.Add(Me.ugByInvoice)
        Me.pnlByInvoiceDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlByInvoiceDetails.Location = New System.Drawing.Point(0, 24)
        Me.pnlByInvoiceDetails.Name = "pnlByInvoiceDetails"
        Me.pnlByInvoiceDetails.Size = New System.Drawing.Size(976, 366)
        Me.pnlByInvoiceDetails.TabIndex = 5
        '
        'ugByInvoice
        '
        Me.ugByInvoice.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugByInvoice.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle
        Me.ugByInvoice.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugByInvoice.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugByInvoice.Location = New System.Drawing.Point(0, 0)
        Me.ugByInvoice.Name = "ugByInvoice"
        Me.ugByInvoice.Size = New System.Drawing.Size(976, 366)
        Me.ugByInvoice.TabIndex = 3
        '
        'pnlByInvoiceTop
        '
        Me.pnlByInvoiceTop.Controls.Add(Me.rdBtnCurrentOutstanding)
        Me.pnlByInvoiceTop.Controls.Add(Me.rdBtnAllYears)
        Me.pnlByInvoiceTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlByInvoiceTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlByInvoiceTop.Name = "pnlByInvoiceTop"
        Me.pnlByInvoiceTop.Size = New System.Drawing.Size(976, 24)
        Me.pnlByInvoiceTop.TabIndex = 0
        '
        'rdBtnCurrentOutstanding
        '
        Me.rdBtnCurrentOutstanding.Location = New System.Drawing.Point(8, 2)
        Me.rdBtnCurrentOutstanding.Name = "rdBtnCurrentOutstanding"
        Me.rdBtnCurrentOutstanding.Size = New System.Drawing.Size(148, 18)
        Me.rdBtnCurrentOutstanding.TabIndex = 1
        Me.rdBtnCurrentOutstanding.Text = "Current && Outstanding"
        '
        'rdBtnAllYears
        '
        Me.rdBtnAllYears.Checked = True
        Me.rdBtnAllYears.Location = New System.Drawing.Point(200, 2)
        Me.rdBtnAllYears.Name = "rdBtnAllYears"
        Me.rdBtnAllYears.Size = New System.Drawing.Size(104, 17)
        Me.rdBtnAllYears.TabIndex = 2
        Me.rdBtnAllYears.TabStop = True
        Me.rdBtnAllYears.Text = "All Years"
        '
        'pnlByInvoiceBottom
        '
        Me.pnlByInvoiceBottom.Controls.Add(Me.btnFacGenerateCreditMemo)
        Me.pnlByInvoiceBottom.Controls.Add(Me.btnFacGenerateInvoice)
        Me.pnlByInvoiceBottom.Controls.Add(Me.btnFacEditCreditMemo)
        Me.pnlByInvoiceBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlByInvoiceBottom.Location = New System.Drawing.Point(0, 390)
        Me.pnlByInvoiceBottom.Name = "pnlByInvoiceBottom"
        Me.pnlByInvoiceBottom.Size = New System.Drawing.Size(976, 40)
        Me.pnlByInvoiceBottom.TabIndex = 4
        '
        'btnFacGenerateCreditMemo
        '
        Me.btnFacGenerateCreditMemo.Location = New System.Drawing.Point(312, 8)
        Me.btnFacGenerateCreditMemo.Name = "btnFacGenerateCreditMemo"
        Me.btnFacGenerateCreditMemo.Size = New System.Drawing.Size(144, 23)
        Me.btnFacGenerateCreditMemo.TabIndex = 6
        Me.btnFacGenerateCreditMemo.Text = "Generate Credit Memo"
        '
        'btnFacGenerateInvoice
        '
        Me.btnFacGenerateInvoice.Location = New System.Drawing.Point(184, 8)
        Me.btnFacGenerateInvoice.Name = "btnFacGenerateInvoice"
        Me.btnFacGenerateInvoice.Size = New System.Drawing.Size(112, 23)
        Me.btnFacGenerateInvoice.TabIndex = 5
        Me.btnFacGenerateInvoice.Text = "Generate Invoice"
        '
        'btnFacEditCreditMemo
        '
        Me.btnFacEditCreditMemo.Location = New System.Drawing.Point(472, 8)
        Me.btnFacEditCreditMemo.Name = "btnFacEditCreditMemo"
        Me.btnFacEditCreditMemo.Size = New System.Drawing.Size(144, 23)
        Me.btnFacEditCreditMemo.TabIndex = 6
        Me.btnFacEditCreditMemo.Text = "Edit Credit Memo"
        '
        'tbPageOldAccessHis
        '
        Me.tbPageOldAccessHis.Controls.Add(Me.ugOldAccessHistory)
        Me.tbPageOldAccessHis.Location = New System.Drawing.Point(4, 24)
        Me.tbPageOldAccessHis.Name = "tbPageOldAccessHis"
        Me.tbPageOldAccessHis.Size = New System.Drawing.Size(976, 430)
        Me.tbPageOldAccessHis.TabIndex = 2
        Me.tbPageOldAccessHis.Text = "Old Access History"
        '
        'ugOldAccessHistory
        '
        Me.ugOldAccessHistory.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugOldAccessHistory.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle
        Me.ugOldAccessHistory.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugOldAccessHistory.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugOldAccessHistory.Location = New System.Drawing.Point(0, 0)
        Me.ugOldAccessHistory.Name = "ugOldAccessHistory"
        Me.ugOldAccessHistory.Size = New System.Drawing.Size(976, 430)
        Me.ugOldAccessHistory.TabIndex = 0
        '
        'tbPageFacilityDocuments
        '
        Me.tbPageFacilityDocuments.Controls.Add(Me.UCFacilityDocuments)
        Me.tbPageFacilityDocuments.Location = New System.Drawing.Point(4, 24)
        Me.tbPageFacilityDocuments.Name = "tbPageFacilityDocuments"
        Me.tbPageFacilityDocuments.Size = New System.Drawing.Size(976, 430)
        Me.tbPageFacilityDocuments.TabIndex = 3
        Me.tbPageFacilityDocuments.Text = "Documents"
        '
        'UCFacilityDocuments
        '
        Me.UCFacilityDocuments.AutoScroll = True
        Me.UCFacilityDocuments.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UCFacilityDocuments.Location = New System.Drawing.Point(0, 0)
        Me.UCFacilityDocuments.Name = "UCFacilityDocuments"
        Me.UCFacilityDocuments.Size = New System.Drawing.Size(976, 430)
        Me.UCFacilityDocuments.TabIndex = 3
        '
        'tbPageFacilityContactList
        '
        Me.tbPageFacilityContactList.Controls.Add(Me.pnlFacilityContactContainer)
        Me.tbPageFacilityContactList.Controls.Add(Me.pnlFacilityContactBottom)
        Me.tbPageFacilityContactList.Controls.Add(Me.pnlFacilityContactHeader)
        Me.tbPageFacilityContactList.Location = New System.Drawing.Point(4, 24)
        Me.tbPageFacilityContactList.Name = "tbPageFacilityContactList"
        Me.tbPageFacilityContactList.Size = New System.Drawing.Size(976, 430)
        Me.tbPageFacilityContactList.TabIndex = 4
        Me.tbPageFacilityContactList.Text = "Contacts"
        '
        'pnlFacilityContactContainer
        '
        Me.pnlFacilityContactContainer.Controls.Add(Me.ugFacilityContacts)
        Me.pnlFacilityContactContainer.Controls.Add(Me.Label4)
        Me.pnlFacilityContactContainer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlFacilityContactContainer.Location = New System.Drawing.Point(0, 25)
        Me.pnlFacilityContactContainer.Name = "pnlFacilityContactContainer"
        Me.pnlFacilityContactContainer.Size = New System.Drawing.Size(976, 375)
        Me.pnlFacilityContactContainer.TabIndex = 5
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
        Me.ugFacilityContacts.Size = New System.Drawing.Size(976, 375)
        Me.ugFacilityContacts.TabIndex = 0
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(792, 56)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(7, 23)
        Me.Label4.TabIndex = 2
        '
        'pnlFacilityContactBottom
        '
        Me.pnlFacilityContactBottom.Controls.Add(Me.btnFacilityModifyContact)
        Me.pnlFacilityContactBottom.Controls.Add(Me.btnFacilityDeleteContact)
        Me.pnlFacilityContactBottom.Controls.Add(Me.btnFacilityAssociateContact)
        Me.pnlFacilityContactBottom.Controls.Add(Me.btnFacilityAddSearchContact)
        Me.pnlFacilityContactBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlFacilityContactBottom.DockPadding.All = 3
        Me.pnlFacilityContactBottom.Location = New System.Drawing.Point(0, 400)
        Me.pnlFacilityContactBottom.Name = "pnlFacilityContactBottom"
        Me.pnlFacilityContactBottom.Size = New System.Drawing.Size(976, 30)
        Me.pnlFacilityContactBottom.TabIndex = 4
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
        Me.pnlFacilityContactHeader.Size = New System.Drawing.Size(976, 25)
        Me.pnlFacilityContactHeader.TabIndex = 3
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
        Me.pnl_FacilityDetail.Controls.Add(Me.mskTxtFacilityFax)
        Me.pnl_FacilityDetail.Controls.Add(Me.mskTxtFacilityPhone)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityFax)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityPhone)
        Me.pnl_FacilityDetail.Controls.Add(Me.btnFeesFacLabels)
        Me.pnl_FacilityDetail.Controls.Add(Me.btnFeesFacEnvelopes)
        Me.pnl_FacilityDetail.Controls.Add(Me.Panel2)
        Me.pnl_FacilityDetail.Controls.Add(Me.lnkLblNextFac)
        Me.pnl_FacilityDetail.Controls.Add(Me.lnkLblPrevFacility)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtDueByNF)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilitySigNFDue)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblDateTransfered)
        Me.pnl_FacilityDetail.Controls.Add(Me.ll)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblLUSTSite)
        Me.pnl_FacilityDetail.Controls.Add(Me.chkLUSTSite)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblCAPCandidate)
        Me.pnl_FacilityDetail.Controls.Add(Me.chkCAPCandidate)
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
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilityAddress)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilitySIC)
        Me.pnl_FacilityDetail.Controls.Add(Me.dtPickFacilityRecvd)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblDateReceived)
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
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilityName)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityAddress)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityName)
        Me.pnl_FacilityDetail.Controls.Add(Me.lblFacilityLatDegree)
        Me.pnl_FacilityDetail.Controls.Add(Me.txtFacilitySIC)
        Me.pnl_FacilityDetail.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnl_FacilityDetail.Location = New System.Drawing.Point(0, 0)
        Me.pnl_FacilityDetail.Name = "pnl_FacilityDetail"
        Me.pnl_FacilityDetail.Size = New System.Drawing.Size(984, 184)
        Me.pnl_FacilityDetail.TabIndex = 2
        '
        'mskTxtFacilityFax
        '
        Me.mskTxtFacilityFax.ContainingControl = Me
        Me.mskTxtFacilityFax.Location = New System.Drawing.Point(424, 136)
        Me.mskTxtFacilityFax.Name = "mskTxtFacilityFax"
        Me.mskTxtFacilityFax.OcxState = CType(resources.GetObject("mskTxtFacilityFax.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtFacilityFax.Size = New System.Drawing.Size(96, 23)
        Me.mskTxtFacilityFax.TabIndex = 1072
        '
        'mskTxtFacilityPhone
        '
        Me.mskTxtFacilityPhone.ContainingControl = Me
        Me.mskTxtFacilityPhone.Location = New System.Drawing.Point(424, 112)
        Me.mskTxtFacilityPhone.Name = "mskTxtFacilityPhone"
        Me.mskTxtFacilityPhone.OcxState = CType(resources.GetObject("mskTxtFacilityPhone.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtFacilityPhone.Size = New System.Drawing.Size(96, 23)
        Me.mskTxtFacilityPhone.TabIndex = 1071
        '
        'lblFacilityFax
        '
        Me.lblFacilityFax.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityFax.Location = New System.Drawing.Point(320, 136)
        Me.lblFacilityFax.Name = "lblFacilityFax"
        Me.lblFacilityFax.Size = New System.Drawing.Size(72, 23)
        Me.lblFacilityFax.TabIndex = 1074
        Me.lblFacilityFax.Text = "Fax:"
        '
        'lblFacilityPhone
        '
        Me.lblFacilityPhone.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityPhone.Location = New System.Drawing.Point(320, 112)
        Me.lblFacilityPhone.Name = "lblFacilityPhone"
        Me.lblFacilityPhone.Size = New System.Drawing.Size(72, 23)
        Me.lblFacilityPhone.TabIndex = 1073
        Me.lblFacilityPhone.Text = "Phone:"
        '
        'btnFeesFacLabels
        '
        Me.btnFeesFacLabels.Location = New System.Drawing.Point(8, 117)
        Me.btnFeesFacLabels.Name = "btnFeesFacLabels"
        Me.btnFeesFacLabels.Size = New System.Drawing.Size(70, 23)
        Me.btnFeesFacLabels.TabIndex = 1070
        Me.btnFeesFacLabels.Text = "Labels"
        '
        'btnFeesFacEnvelopes
        '
        Me.btnFeesFacEnvelopes.Location = New System.Drawing.Point(8, 89)
        Me.btnFeesFacEnvelopes.Name = "btnFeesFacEnvelopes"
        Me.btnFeesFacEnvelopes.Size = New System.Drawing.Size(70, 23)
        Me.btnFeesFacEnvelopes.TabIndex = 1069
        Me.btnFeesFacEnvelopes.Text = "Envelopes"
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.btnFacComments)
        Me.Panel2.Controls.Add(Me.btnFacFlags)
        Me.Panel2.Location = New System.Drawing.Point(608, 130)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(192, 32)
        Me.Panel2.TabIndex = 1052
        '
        'btnFacComments
        '
        Me.btnFacComments.Location = New System.Drawing.Point(104, 5)
        Me.btnFacComments.Name = "btnFacComments"
        Me.btnFacComments.Size = New System.Drawing.Size(80, 23)
        Me.btnFacComments.TabIndex = 1039
        Me.btnFacComments.Text = "Comments"
        '
        'btnFacFlags
        '
        Me.btnFacFlags.Location = New System.Drawing.Point(24, 5)
        Me.btnFacFlags.Name = "btnFacFlags"
        Me.btnFacFlags.TabIndex = 1040
        Me.btnFacFlags.Text = "Flags"
        '
        'lnkLblNextFac
        '
        Me.lnkLblNextFac.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lnkLblNextFac.Location = New System.Drawing.Point(720, 161)
        Me.lnkLblNextFac.Name = "lnkLblNextFac"
        Me.lnkLblNextFac.Size = New System.Drawing.Size(56, 16)
        Me.lnkLblNextFac.TabIndex = 1051
        Me.lnkLblNextFac.TabStop = True
        Me.lnkLblNextFac.Text = "Next>>"
        '
        'lnkLblPrevFacility
        '
        Me.lnkLblPrevFacility.AutoSize = True
        Me.lnkLblPrevFacility.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lnkLblPrevFacility.Location = New System.Drawing.Point(640, 161)
        Me.lnkLblPrevFacility.Name = "lnkLblPrevFacility"
        Me.lnkLblPrevFacility.Size = New System.Drawing.Size(70, 17)
        Me.lnkLblPrevFacility.TabIndex = 1050
        Me.lnkLblPrevFacility.TabStop = True
        Me.lnkLblPrevFacility.Text = "<< Previous"
        '
        'txtDueByNF
        '
        Me.txtDueByNF.AcceptsTab = True
        Me.txtDueByNF.AutoSize = False
        Me.txtDueByNF.Enabled = False
        Me.txtDueByNF.Location = New System.Drawing.Point(944, 136)
        Me.txtDueByNF.Name = "txtDueByNF"
        Me.txtDueByNF.Size = New System.Drawing.Size(64, 21)
        Me.txtDueByNF.TabIndex = 1049
        Me.txtDueByNF.Text = ""
        Me.txtDueByNF.Visible = False
        Me.txtDueByNF.WordWrap = False
        '
        'lblFacilitySigNFDue
        '
        Me.lblFacilitySigNFDue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilitySigNFDue.ForeColor = System.Drawing.Color.Red
        Me.lblFacilitySigNFDue.Location = New System.Drawing.Point(960, 112)
        Me.lblFacilitySigNFDue.Name = "lblFacilitySigNFDue"
        Me.lblFacilitySigNFDue.Size = New System.Drawing.Size(112, 23)
        Me.lblFacilitySigNFDue.TabIndex = 1048
        Me.lblFacilitySigNFDue.Text = "Due By:"
        Me.lblFacilitySigNFDue.Visible = False
        '
        'lblDateTransfered
        '
        Me.lblDateTransfered.BackColor = System.Drawing.SystemColors.Control
        Me.lblDateTransfered.Enabled = False
        Me.lblDateTransfered.Location = New System.Drawing.Point(936, 64)
        Me.lblDateTransfered.Name = "lblDateTransfered"
        Me.lblDateTransfered.TabIndex = 1047
        '
        'll
        '
        Me.ll.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ll.Location = New System.Drawing.Point(1128, 56)
        Me.ll.Name = "ll"
        Me.ll.Size = New System.Drawing.Size(24, 23)
        Me.ll.TabIndex = 1035
        '
        'lblLUSTSite
        '
        Me.lblLUSTSite.Location = New System.Drawing.Point(320, 56)
        Me.lblLUSTSite.Name = "lblLUSTSite"
        Me.lblLUSTSite.Size = New System.Drawing.Size(104, 23)
        Me.lblLUSTSite.TabIndex = 1030
        Me.lblLUSTSite.Text = "Active LUST Site"
        Me.lblLUSTSite.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkLUSTSite
        '
        Me.chkLUSTSite.Enabled = False
        Me.chkLUSTSite.Location = New System.Drawing.Point(424, 64)
        Me.chkLUSTSite.Name = "chkLUSTSite"
        Me.chkLUSTSite.Size = New System.Drawing.Size(16, 16)
        Me.chkLUSTSite.TabIndex = 6
        Me.chkLUSTSite.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCAPCandidate
        '
        Me.lblCAPCandidate.Location = New System.Drawing.Point(464, 56)
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
        Me.chkCAPCandidate.Location = New System.Drawing.Point(568, 64)
        Me.chkCAPCandidate.Name = "chkCAPCandidate"
        Me.chkCAPCandidate.Size = New System.Drawing.Size(16, 16)
        Me.chkCAPCandidate.TabIndex = 7
        Me.chkCAPCandidate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbFacilityType
        '
        Me.cmbFacilityType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbFacilityType.DropDownWidth = 180
        Me.cmbFacilityType.Enabled = False
        Me.cmbFacilityType.ItemHeight = 15
        Me.cmbFacilityType.Location = New System.Drawing.Point(688, 40)
        Me.cmbFacilityType.Name = "cmbFacilityType"
        Me.cmbFacilityType.Size = New System.Drawing.Size(176, 23)
        Me.cmbFacilityType.TabIndex = 14
        '
        'txtFacilityLatSec
        '
        Me.txtFacilityLatSec.AcceptsTab = True
        Me.txtFacilityLatSec.AutoSize = False
        Me.txtFacilityLatSec.Enabled = False
        Me.txtFacilityLatSec.Location = New System.Drawing.Point(768, 72)
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
        Me.txtFacilityLongSec.Location = New System.Drawing.Point(768, 104)
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
        Me.txtFacilityLatMin.Location = New System.Drawing.Point(736, 72)
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
        Me.txtFacilityLongMin.Location = New System.Drawing.Point(736, 104)
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
        Me.lblFacilityLongSec.Location = New System.Drawing.Point(808, 104)
        Me.lblFacilityLongSec.Name = "lblFacilityLongSec"
        Me.lblFacilityLongSec.Size = New System.Drawing.Size(10, 16)
        Me.lblFacilityLongSec.TabIndex = 1017
        Me.lblFacilityLongSec.Text = """"
        '
        'lblFacilityLatMin
        '
        Me.lblFacilityLatMin.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLatMin.Location = New System.Drawing.Point(888, 48)
        Me.lblFacilityLatMin.Name = "lblFacilityLatMin"
        Me.lblFacilityLatMin.Size = New System.Drawing.Size(8, 23)
        Me.lblFacilityLatMin.TabIndex = 1016
        Me.lblFacilityLatMin.Text = "'"
        Me.lblFacilityLatMin.Visible = False
        '
        'lblFacilityLatSec
        '
        Me.lblFacilityLatSec.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLatSec.Location = New System.Drawing.Point(808, 72)
        Me.lblFacilityLatSec.Name = "lblFacilityLatSec"
        Me.lblFacilityLatSec.Size = New System.Drawing.Size(10, 16)
        Me.lblFacilityLatSec.TabIndex = 1015
        Me.lblFacilityLatSec.Text = """"
        '
        'lblFacilityLongDegree
        '
        Me.lblFacilityLongDegree.Font = New System.Drawing.Font("Symbol", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLongDegree.Location = New System.Drawing.Point(720, 96)
        Me.lblFacilityLongDegree.Name = "lblFacilityLongDegree"
        Me.lblFacilityLongDegree.Size = New System.Drawing.Size(16, 16)
        Me.lblFacilityLongDegree.TabIndex = 1010
        Me.lblFacilityLongDegree.Text = "o"
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
        Me.lblFacilitySIC.Location = New System.Drawing.Point(320, 80)
        Me.lblFacilitySIC.Name = "lblFacilitySIC"
        Me.lblFacilitySIC.Size = New System.Drawing.Size(80, 23)
        Me.lblFacilitySIC.TabIndex = 150
        Me.lblFacilitySIC.Text = "SIC:"
        Me.lblFacilitySIC.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
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
        'lblFacilityStatusValue
        '
        Me.lblFacilityStatusValue.BackColor = System.Drawing.Color.Transparent
        Me.lblFacilityStatusValue.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblFacilityStatusValue.Enabled = False
        Me.lblFacilityStatusValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityStatusValue.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblFacilityStatusValue.Location = New System.Drawing.Point(424, 40)
        Me.lblFacilityStatusValue.Name = "lblFacilityStatusValue"
        Me.lblFacilityStatusValue.Size = New System.Drawing.Size(160, 16)
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
        Me.txtFacilityLongDegree.Location = New System.Drawing.Point(688, 104)
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
        Me.txtFacilityLatDegree.Location = New System.Drawing.Point(688, 72)
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
        Me.lblFacilityLongitude.Location = New System.Drawing.Point(600, 104)
        Me.lblFacilityLongitude.Name = "lblFacilityLongitude"
        Me.lblFacilityLongitude.Size = New System.Drawing.Size(96, 23)
        Me.lblFacilityLongitude.TabIndex = 111
        Me.lblFacilityLongitude.Text = "Longitude:"
        '
        'lblFacilityLatitude
        '
        Me.lblFacilityLatitude.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityLatitude.Location = New System.Drawing.Point(600, 72)
        Me.lblFacilityLatitude.Name = "lblFacilityLatitude"
        Me.lblFacilityLatitude.Size = New System.Drawing.Size(96, 23)
        Me.lblFacilityLatitude.TabIndex = 110
        Me.lblFacilityLatitude.Text = "Latitude:"
        '
        'lblFacilityType
        '
        Me.lblFacilityType.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFacilityType.Location = New System.Drawing.Point(600, 40)
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
        Me.lblFacilityLatDegree.Location = New System.Drawing.Point(720, 64)
        Me.lblFacilityLatDegree.Name = "lblFacilityLatDegree"
        Me.lblFacilityLatDegree.Size = New System.Drawing.Size(16, 16)
        Me.lblFacilityLatDegree.TabIndex = 1009
        Me.lblFacilityLatDegree.Text = "o"
        '
        'txtFacilitySIC
        '
        Me.txtFacilitySIC.BackColor = System.Drawing.Color.Transparent
        Me.txtFacilitySIC.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtFacilitySIC.Enabled = False
        Me.txtFacilitySIC.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFacilitySIC.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.txtFacilitySIC.Location = New System.Drawing.Point(424, 88)
        Me.txtFacilitySIC.Name = "txtFacilitySIC"
        Me.txtFacilitySIC.Size = New System.Drawing.Size(160, 16)
        Me.txtFacilitySIC.TabIndex = 124
        '
        'pnlFeesHeader
        '
        Me.pnlFeesHeader.Controls.Add(Me.lblOwnerLastEditedOn)
        Me.pnlFeesHeader.Controls.Add(Me.lblOwnerLastEditedBy)
        Me.pnlFeesHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlFeesHeader.Location = New System.Drawing.Point(0, 0)
        Me.pnlFeesHeader.Name = "pnlFeesHeader"
        Me.pnlFeesHeader.Size = New System.Drawing.Size(992, 24)
        Me.pnlFeesHeader.TabIndex = 1
        '
        'lblOwnerLastEditedOn
        '
        Me.lblOwnerLastEditedOn.Location = New System.Drawing.Point(672, 5)
        Me.lblOwnerLastEditedOn.Name = "lblOwnerLastEditedOn"
        Me.lblOwnerLastEditedOn.Size = New System.Drawing.Size(168, 16)
        Me.lblOwnerLastEditedOn.TabIndex = 1016
        Me.lblOwnerLastEditedOn.Text = "Last Edited On :"
        '
        'lblOwnerLastEditedBy
        '
        Me.lblOwnerLastEditedBy.Location = New System.Drawing.Point(456, 4)
        Me.lblOwnerLastEditedBy.Name = "lblOwnerLastEditedBy"
        Me.lblOwnerLastEditedBy.Size = New System.Drawing.Size(208, 16)
        Me.lblOwnerLastEditedBy.TabIndex = 1015
        Me.lblOwnerLastEditedBy.Text = "Last Edited By :"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(312, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.TabIndex = 2
        Me.Label5.Text = "Label5"
        Me.Label5.Visible = False
        '
        'Fees
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(992, 694)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.tbCntrlFees)
        Me.Controls.Add(Me.pnlFeesHeader)
        Me.Name = "Fees"
        Me.Text = "Fees"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.tbCntrlFees.ResumeLayout(False)
        Me.tbPageOwnerDetail.ResumeLayout(False)
        Me.pnlOwnerBottom.ResumeLayout(False)
        Me.tbCtrlOwnerTransactions.ResumeLayout(False)
        Me.tbPageFacilities.ResumeLayout(False)
        Me.pnlFacDetails.ResumeLayout(False)
        CType(Me.ugFacilityList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlOwnerFacilityBottom.ResumeLayout(False)
        Me.pnlOwnerTotalsContainer.ResumeLayout(False)
        Me.tbPageRefunds.ResumeLayout(False)
        Me.pnlRefundsDetails.ResumeLayout(False)
        CType(Me.ugRefunds, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlRefundsBottom.ResumeLayout(False)
        Me.tbPageReceipts.ResumeLayout(False)
        Me.pnlReceiptsDetails.ResumeLayout(False)
        Me.pnlOwnerReceiptFill.ResumeLayout(False)
        CType(Me.ugReceipts, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlOwnerReceiptTop.ResumeLayout(False)
        Me.pnlReceiptsBottom.ResumeLayout(False)
        Me.tbPageOverages.ResumeLayout(False)
        CType(Me.ugOverage, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbPageOwnerDocuments.ResumeLayout(False)
        Me.tbPageOwnerContactList.ResumeLayout(False)
        Me.pnlOwnerContactContainer.ResumeLayout(False)
        CType(Me.ugOwnerContacts, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlOwnerContactButtons.ResumeLayout(False)
        Me.pnlOwnerContactHeader.ResumeLayout(False)
        Me.tbPageInvoices.ResumeLayout(False)
        Me.pnlManualInvoice.ResumeLayout(False)
        Me.pnlManualInvoiceDetails.ResumeLayout(False)
        CType(Me.ugManualInvoices, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlManualInvoiceBottom.ResumeLayout(False)
        Me.pnManualInvoiceTop.ResumeLayout(False)
        Me.pnlInvoices.ResumeLayout(False)
        Me.pnlInvoicesDetails.ResumeLayout(False)
        CType(Me.ugInvoices, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlInvoicesControls.ResumeLayout(False)
        Me.pnlInvoicesBottom.ResumeLayout(False)
        Me.pnlInvoicesTop.ResumeLayout(False)
        Me.pnlOwnerDetails.ResumeLayout(False)
        Me.pnlOwnerButtons.ResumeLayout(False)
        CType(Me.mskTxtOwnerFax, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtOwnerPhone2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtOwnerPhone, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbPageFacilityDetail.ResumeLayout(False)
        Me.pnlFacilityBottom.ResumeLayout(False)
        Me.tbCtrlFacTransactions.ResumeLayout(False)
        Me.tbPageByTransaction.ResumeLayout(False)
        Me.pnlByTransacDetails.ResumeLayout(False)
        CType(Me.ugByTransaction, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlByTransacTop.ResumeLayout(False)
        Me.tbPageByInvoice.ResumeLayout(False)
        Me.pnlByInvoiceDetails.ResumeLayout(False)
        CType(Me.ugByInvoice, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlByInvoiceTop.ResumeLayout(False)
        Me.pnlByInvoiceBottom.ResumeLayout(False)
        Me.tbPageOldAccessHis.ResumeLayout(False)
        CType(Me.ugOldAccessHistory, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbPageFacilityDocuments.ResumeLayout(False)
        Me.tbPageFacilityContactList.ResumeLayout(False)
        Me.pnlFacilityContactContainer.ResumeLayout(False)
        CType(Me.ugFacilityContacts, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlFacilityContactBottom.ResumeLayout(False)
        Me.pnlFacilityContactHeader.ResumeLayout(False)
        Me.pnl_FacilityDetail.ResumeLayout(False)
        CType(Me.mskTxtFacilityFax, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtFacilityPhone, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.pnlFeesHeader.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region "Intialization"
    Private Sub InitControls()
        Try
            UIUtilsGen.PopulateOwnerType(cmbOwnerType, pOwn)
            If pOwn.ID <> 0 Then
                UIUtilsGen.PopulateFacilityType(Me.cmbFacilityType, pOwn.Facilities)
            End If
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
    Private Sub Fees_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.TextChanged
        Try
            MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "WindowName", Me.Text, "Fees")
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub Fees_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        'Dim MyFrm As MusterContainer
        Try
            MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "Fees")
            'MyFrm = Me.MdiParent
            'If lblOwnerIDValue.Text <> String.Empty And lblFacilityIDValue.Text = String.Empty Then
            'MyFrm.FlagsChanged(lblOwnerIDValue.Text, 0, 0, 0, "Fees", Me.Text)
            'ElseIf lblOwnerIDValue.Text <> String.Empty And lblFacilityIDValue.Text <> String.Empty Then
            'MyFrm.FlagsChanged(lblOwnerIDValue.Text, lblFacilityIDValue.Text, 0, 0, "Fees", Me.Text)
            'End If
            If lblOwnerIDValue.Text <> String.Empty Then ' And lblFacilityIDValue.Text = String.Empty Then
                pOwn.Retrieve(Me.lblOwnerIDValue.Text, "SELF")
            End If
            RecalcOwnerTotals()
            bolFrmActivated = True
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#Region "General"
    Private Sub ShowError(ByVal ex As Exception)
        Dim MyErr As New ErrorReport(ex)
        MyErr.ShowDialog()
    End Sub
#End Region
#Region "Tab Operations"
    Private Sub tbCntrlFees_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbCntrlFees.Click
        If bolLoading Then Exit Sub
        Dim MyFrm As MusterContainer
        Try
            Select Case tbCntrlFees.SelectedTab.Name.ToUpper
                Case "TBPAGEFACILITYDETAIL"
                    If Me.ugFacilityList.Rows.Count <> 0 Then
                        If Me.lblFacilityIDValue.Text = String.Empty Or lblFacilityIDValue.Text = "-1" Then
                            If ugFacilityList.ActiveRow Is Nothing Then
                                ugFacilityList.ActiveRow = ugFacilityList.Rows(0)
                            End If
                            PopulateFacilityInfo(Integer.Parse(ugFacilityList.ActiveRow.Cells("FACILITY_ID").Text))
                        Else
                            PopulateFacilityInfo(Integer.Parse(Me.lblFacilityIDValue.Text))
                            nFacilityID = Integer.Parse(lblFacilityIDValue.Text)
                        End If
                        Me.Tag = Me.lblFacilityIDValue.Text
                        'ProcessFacTransactions()
                        Me.lblFacilityIDValue.Focus()
                        MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "FacilityDetails", Me.lblFacilityIDValue.Text, "Fees")
                    End If
                    If ugFacilityList.Rows.Count <= 0 And Me.lblOwnerIDValue.Text <> String.Empty Then
                        Dim msgResult As MsgBoxResult
                        msgResult = MsgBox("No facilities found for owner" + lblOwnerIDValue.Text)
                        Exit Select
                    End If

                    Me.Text = "Fees - Facility Detail - (" & IIf(txtFacilityName.Text = String.Empty, "New", txtFacilityName.Text) & ")"
                Case "TBPAGEOWNERDETAIL"
                    Me.Text = "Fees - Owner Detail (" & txtOwnerName.Text & ")"
                    'If lblOwnerIDValue.Text <> String.Empty Then
                    '    UIUtilsGen.PopulateOwnerFacilities(pOwn, Me, Integer.Parse(Me.lblOwnerIDValue.Text))
                    'End If
                    'MyFrm = Me.MdiParent
                    'MyFrm.FlagsChanged(lblOwnerIDValue.Text, uiutilsgen.EntityTypes.Owner, "Fees", Me.Text)
            End Select
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub tbCntrlFees_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbCntrlFees.SelectedIndexChanged
        If bolLoading Then Exit Sub
        ' Dim nFacilityID As Integer
        Try
            Select Case tbCntrlFees.SelectedTab.Name.ToUpper
                Case "TBPAGEOWNERDETAIL"
                    Me.Text = "Fees - Owner Detail (" & txtOwnerName.Text & ")"
                    Me.PopulateOwnerInfo(pOwn.ID)
                    ' clearing the facility info in mustercontainer
                    'Me.lblFacilityIDValue.Text = String.Empty
                    ugFacilityList.ActiveRow = Nothing

                Case "TBPAGEFACILITYDETAIL"
                    Me.Text = "Fees - Facility Detail - (" & IIf(txtFacilityName.Text = String.Empty, "New", txtFacilityName.Text) & ")"
                    If ugFacilityList.Rows.Count > 0 Then
                        If Not ugFacilityList.ActiveRow Is Nothing Then
                            ' ugFacilityList.ActiveRow = ugFacilityList.Rows(0)

                            nFacilityID = ugFacilityList.ActiveRow.Cells("Facility_ID").Value

                        End If

                        'If pOwn.Facilities.ID <= 0 Then
                        'Me.PopulateFacilityInfo(Integer.Parse(ugFacilityList.Rows(0).Cells("Facility_ID").Value))
                        'Else
                        '  Me.PopulateFacilityInfo(pOwn.Facilities.ID)
                        ' End If

                        If nFacilityID <= 0 Then
                            Me.PopulateFacilityInfo(pOwn.Facilities.ID)
                        Else
                            Me.PopulateFacilityInfo(nFacilityID)
                        End If
                    End If
            End Select
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub tbCtrlFacTransactions_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbCtrlFacTransactions.SelectedIndexChanged
        If Not bolFormLoaded Then Exit Sub
        If bolLoading Then Exit Sub
        'ProcessFacTransactions()
        Select Case tbCtrlFacTransactions.SelectedTab.Name
            Case tbPageByTransaction.Name
                LoadugByTransaction()
            Case tbPageByInvoice.Name
                LoadugByInvoice()
            Case tbPageOldAccessHis.Name
                LoadugOldAccessHistory()
            Case tbPageFacilityDocuments.Name
                UCFacilityDocuments.LoadDocumentsGrid(nFacilityID, UIUtilsGen.EntityTypes.Facility, UIUtilsGen.ModuleID.Fees)
            Case tbPageFacilityContactList.Name
                LoadContacts(ugFacilityContacts, nFacilityID, UIUtilsGen.EntityTypes.Facility)
        End Select
    End Sub

    Private Sub tbCtrlOwnerTransactions_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbCtrlOwnerTransactions.SelectedIndexChanged
        If Not bolFormLoaded Then Exit Sub
        If bolLoading Then Exit Sub
        Select Case tbCtrlOwnerTransactions.SelectedTab.Name
            Case tbPageFacilities.Name ' 0  'FacList
                LoadugFacilityList()

            Case tbPageInvoices.Name ' 1  'Invoices
                LoadugInvoices()

            Case tbPageReceipts.Name ' 2  'Receipts
                LoadugReceipts()

            Case tbPageRefunds.Name ' 3  'Refunds
                LoadugRefunds()

            Case tbPageOverages.Name ' 4  'Overages
                LoadugOverages()
            Case tbPageOwnerDocuments.Name
                UCOwnerDocuments.LoadDocumentsGrid(nOwnerID, UIUtilsGen.EntityTypes.Owner, UIUtilsGen.ModuleID.Fees)

            Case tbPageOwnerContactList.Name
                LoadContacts(ugOwnerContacts, nOwnerID, UIUtilsGen.EntityTypes.Owner)
        End Select

    End Sub
#End Region
#Region "Form Events"
    Private Sub Fees_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim cntrl As Control
        Dim uiUts As New UIUtilsGen
        Try
            bolLoading = True
            For Each cntrl In Me.Controls
                uiUts.ClearComboBox(cntrl)
                'uiUts.RetainCurrentDateValue(cntrl)
            Next
            bolLoading = False
            bolFormLoaded = True

            SetupManualInvoiceGrid()
            RecalcOwnerTotals()

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
            '        success = pOwn.Save(CType(UIUtilsGen.ModuleID.Fees, Integer), MusterContainer.AppUser.UserKey, returnVal)
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
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
#Region "UI Support Routines"

    Private Sub txtFacilityAddress_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFacilityAddress.DoubleClick
        Try

            Dim addressForm As Address
            If (pOwn.Facilities.ID = 0 Or pOwn.Facilities.ID <> nFacilityID) And nFacilityID > 0 Then
                UIUtilsGen.PopulateFacilityInfo(Me, pOwn.OwnerInfo, pOwn.Facilities, nFacilityID)

            End If
            Address.EditAddress(addressForm, pOwn.Facilities.ID, pOwn.Facilities.FacilityAddresses, "Facility", UIUtilsGen.ModuleID.Fees, txtFacilityAddress, UIUtilsGen.EntityTypes.Facility, True)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtOwnerAddress_DblClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtOwnerAddress.DoubleClick
        Try

            Dim addressForm As Address
            Address.EditAddress(addressForm, pOwn.ID, pOwn.Addresses, "Owner", UIUtilsGen.ModuleID.Fees, txtOwnerAddress, UIUtilsGen.EntityTypes.Owner)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub



    Private Sub frmClosing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        If sender.GetType.Name.IndexOf("LateFeeWaiverRequest") >= 0 Then
        ElseIf sender.GetType.Name.IndexOf("GenerateCreditMemo") >= 0 Then
        ElseIf sender.GetType.Name.IndexOf("ReallocateOwnerOverage") >= 0 Then
        ElseIf sender.GetType.Name.IndexOf("GenerateDebitMemo") >= 0 Then
        ElseIf sender.GetType.Name.IndexOf("OverpaymentReason") >= 0 Then
        ElseIf sender.GetType.Name.IndexOf("GenerateRefund") >= 0 Then
        End If
    End Sub
    Private Sub frmClosed(ByVal sender As Object, ByVal e As System.EventArgs)
        If sender.GetType.Name.IndexOf("LateFeeWaiverRequest") >= 0 Then
            frmLateFeeWaiver = Nothing
        ElseIf sender.GetType.Name.IndexOf("GenerateCreditMemo") >= 0 Then
            frmGenCreditMemo = Nothing
        ElseIf sender.GetType.Name.IndexOf("ReallocateOwnerOverage") >= 0 Then
            frmReallocateOverage = Nothing
        ElseIf sender.GetType.Name.IndexOf("GenerateDebitMemo") >= 0 Then
            frmGenDebitMemo = Nothing
        ElseIf sender.GetType.Name.IndexOf("OverpaymentReason") >= 0 Then
            frmOverpayment = Nothing
        ElseIf sender.GetType.Name.IndexOf("GenerateRefund") >= 0 Then
            frmGenRefund = Nothing
        End If
    End Sub
#End Region
#Region "Owner Operations"
#Region "UI Support Routines"
    Friend Sub PopulateOwnerInfo(ByVal OwnerID As Integer)
        Dim MyFrm As MusterContainer
        Try
            ' Me.lblFacilityIDValue.Text = String.Empty
            UIUtilsGen.PopulateOwnerInfo(OwnerID, pOwn, Me)
            Select Case tbCtrlOwnerTransactions.SelectedTab.Name
                Case tbPageFacilities.Name
                    LoadugFacilityList()
                Case tbPageInvoices.Name
                    LoadugInvoices()
                Case tbPageReceipts.Name
                    LoadugReceipts()
                Case tbPageRefunds.Name
                    LoadugRefunds()
                Case tbPageOverages.Name
                    LoadugOverages()
                Case tbPageOwnerDocuments.Name
                    UCOwnerDocuments.LoadDocumentsGrid(OwnerID, 9, 892)
                Case tbPageOwnerContactList.Name
                    LoadContacts(ugOwnerContacts, OwnerID, UIUtilsGen.EntityTypes.Owner)
            End Select
            If bolFrmActivated Then
                MyFrm = Me.MdiParent
                'oEntity.GetEntity("Owner")
                If lblOwnerIDValue.Text <> String.Empty Then
                    MyFrm.FlagsChanged(lblOwnerIDValue.Text, UIUtilsGen.EntityTypes.Owner, "Fees", Me.Text)
                End If
            End If
            CommentsMaintenance(, , True)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
#Region "UI Control Events"
    Private Sub ugFacilityList_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ugFacilityList.DoubleClick
        If Me.Cursor Is Cursors.WaitCursor Then
            Exit Sub
        End If
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            Me.Cursor = Cursors.WaitCursor
            tbCntrlFees.SelectedTab = Me.tbPageFacilityDetail
            PopulateFacilityInfo(Integer.Parse(ugFacilityList.ActiveRow.Cells.Item("Facility_ID").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw)))
            nFacilityID = Integer.Parse(ugFacilityList.ActiveRow.Cells.Item("Facility_ID").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw))
            'ProcessFacTransactions()
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
            'ProcessFacTransactions()
            Select Case tbCtrlFacTransactions.SelectedTab.Name
                Case tbPageByTransaction.Name
                    LoadugByTransaction()
                Case tbPageByInvoice.Name
                    LoadugByInvoice()
                Case tbPageOldAccessHis.Name
                    LoadugOldAccessHistory()
                Case tbPageFacilityDocuments.Name
                    UCFacilityDocuments.LoadDocumentsGrid(nFacilityID, UIUtilsGen.EntityTypes.Facility, UIUtilsGen.ModuleID.Fees)
                Case tbPageFacilityContactList.Name
                    LoadContacts(ugFacilityContacts, nFacilityID, UIUtilsGen.EntityTypes.Facility)
            End Select
            If bolFrmActivated Then
                MyFrm = Me.MdiParent
                'oEntity.GetEntity("Facility")
                MyFrm.FlagsChanged(lblFacilityIDValue.Text, UIUtilsGen.EntityTypes.Facility, "Fees", Me.Text)
            End If
            CommentsMaintenance(, , True)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
#Region "UI Control Events"

    Private Sub lnkLblNextFac_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkLblNextFac.LinkClicked
        Try

            PopulateFacilityInfo(GetPrevNextFacility(lblFacilityIDValue.Text, True))
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub lnkLblPrevFacility_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkLblPrevFacility.LinkClicked
        Try

            PopulateFacilityInfo(GetPrevNextFacility(lblFacilityIDValue.Text, False))
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    'Private Sub lnkLblNextFac_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkLblNextFac.LinkClicked
    '    Try
    '        PopulateFacilityInfo(Integer.Parse(pOwn.Facilities.GetNext()))
    '        MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "FacilityDetails", Me.lblFacilityIDValue.Text, "Fees")
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub lnkLblPrevFacility_LinkClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkLblPrevFacility.LinkClicked
    '    Try
    '        PopulateFacilityInfo(Integer.Parse(Integer.Parse(pOwn.Facilities.GetPrevious())))
    '        MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "FacilityDetails", Me.lblFacilityIDValue.Text, "Fees")
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    Private Function GetPrevNextFacility(ByVal facID As Integer, ByVal getNext As Boolean) As Integer
        Try
            Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
            Dim retVal As Integer = facID
            Dim sl As New SortedList

            For Each ugRow In ugFacilityList.Rows
                sl.Add(ugRow.Cells("Facility_ID").Value, ugRow.Cells("Facility_ID").Value)
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
    Private Function GetPrevNext(ByVal sl As SortedList, ByVal getNext As Boolean, ByVal key As String) As String
        Try
            Dim retVal As String
            Dim index As String = sl.IndexOfKey(key)

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
#Region "Invoices"
#Region "UI Support Routines"

#End Region
#Region "UI Control Events"
    Private Sub btnReqLateFeeWaiver_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReqLateFeeWaiver.Click
        Dim nCurrentFY As Int16

        Try
            If GetVisibleRowCount(ugInvoices) = 0 Then
                MsgBox("No Invoices Available For Late Fee Waiver", MsgBoxStyle.OKOnly, "Invalid Invoice")
                Exit Sub
            End If

            'Get Fiscal Year from Fees Basis
            nCurrentFY = oFeesBasis.GetFiscalYearForFee

            If ugInvoices.Rows.Count > 0 Then
                ' #2814
                Dim bolRowisLateFee As Boolean = False
                Dim bolRowisMisc As Boolean = False

                If ugInvoices.ActiveRow.Band.Index = 0 Then
                    If ugInvoices.ActiveRow.Cells("Fee_Type").Text.IndexOf("Late Fee") > -1 Then
                        bolRowisLateFee = True
                    ElseIf ugInvoices.ActiveRow.Cells("Fee_Type").Text.IndexOf("Miscellaneous") > -1 Then
                        MsgBox("You must select a Miscellaneous Invoice Detail Line with description = 'Transfer Invoice - UST-LATE' within the current fiscal year.", MsgBoxStyle.OKOnly, "Bad Row Selected")
                        Exit Sub
                    Else
                        MsgBox("You must select either a Late Fee record OR an Invoice Detail Line with description = 'Transfer Invoice - UST-LATE' for current fiscal year.", MsgBoxStyle.OKOnly, "Bad Row Selected")
                        Exit Sub
                    End If
                ElseIf ugInvoices.ActiveRow.Band.Index = 1 Then
                    If ugInvoices.ActiveRow.ParentRow.Cells("Fee_Type").Text.IndexOf("Late Fee") > -1 Then
                        ugInvoices.ActiveRow = ugInvoices.ActiveRow.ParentRow
                        bolRowisLateFee = True
                    ElseIf ugInvoices.ActiveRow.ParentRow.Cells("Fee_Type").Text.IndexOf("Late Fee") <= -1 Then
                        bolRowisMisc = True
                    Else
                        MsgBox("You must select either a Late Fee record OR an Invoice Detail Line with description = 'Transfer Invoice - UST-LATE' for current fiscal year.", MsgBoxStyle.OKOnly, "Bad Row Selected")
                        Exit Sub
                    End If
                End If

                If bolRowisLateFee Or bolRowisMisc Then
                    If bolRowisLateFee Then
                        If ugInvoices.ActiveRow.Cells("Late_Cert_ID").Value Is DBNull.Value Then
                            MsgBox("Invalid Late Fee Invoice.", MsgBoxStyle.OKOnly, "Bad Row Selected")
                            Exit Sub
                        End If
                        If ugInvoices.ActiveRow.Cells("SFY").Value <> nCurrentFY Then
                            MsgBox("You must select a Late Fee Invoice within the current fiscal year.", MsgBoxStyle.OKOnly, "Bad Fiscal Year")
                            Exit Sub
                        End If
                    Else
                        If ugInvoices.ActiveRow.Cells("LineDescription").Text.IndexOf("Transfer Invoice - UST-LATE") < 0 Then
                            MsgBox("You must select a Miscellaneous Invoice Detail Line with description = 'Transfer Invoice - UST-LATE' within the current fiscal year.", MsgBoxStyle.OKOnly, "Bad Row Selected")
                            Exit Sub
                        End If
                        If ugInvoices.ActiveRow.Cells("SFY").Value <> nCurrentFY Then
                            MsgBox("You must select a Miscellaneous Invoice within the current fiscal year.", MsgBoxStyle.OKOnly, "Bad Fiscal Year")
                            Exit Sub
                        End If
                    End If

                    If IsNothing(frmLateFeeWaiver) Then
                        frmLateFeeWaiver = New LateFeeWaiverRequest
                        AddHandler frmLateFeeWaiver.Closing, AddressOf frmClosing
                        AddHandler frmLateFeeWaiver.Closed, AddressOf frmClosed
                        frmLateFeeWaiver.Text = "Late Waiver Request Maintenance"
                    End If

                    If bolRowisLateFee Then
                        frmLateFeeWaiver.LateCertID = ugInvoices.ActiveRow.Cells("Late_Cert_ID").Value
                    Else
                        frmLateFeeWaiver.LateCertID = 0
                    End If

                    frmLateFeeWaiver.ugRowMisc = IIf(bolRowisMisc, ugInvoices.ActiveRow, Nothing)
                    frmLateFeeWaiver.ShowDialog()
                    LoadugInvoices()
                End If

                'If ugInvoices.ActiveRow.Band.Index <> 0 Then
                '    MsgBox("You Must Select The Invoice Header, Not The Detail Line.", MsgBoxStyle.OKOnly, "Bad Row")
                '    Exit Sub
                'End If
                'If IsDBNull(ugInvoices.ActiveRow.Cells("Late_Cert_ID").Value) Then
                '    MsgBox("You must select a Late Fee Invoice.", MsgBoxStyle.OKOnly, "Select Late Fee Invoice")
                '    Exit Sub
                'End If
                'If ugInvoices.ActiveRow.Cells("SFY").Value <> nCurrentFY Then
                '    MsgBox("You Must Select a Late Fee Invoice Within The Current Fiscal Year.", MsgBoxStyle.OKOnly, "Bad Fiscal Year")
                '    Exit Sub
                'End If
            Else
                'MsgBox("You must select a Late Fee Invoice.", MsgBoxStyle.OKOnly, "Select Row")
                MsgBox("You must select either a Late Fee record OR an Invoice Detail Line with description = 'Transfer Invoice - UST-LATE' for current fiscal year.", MsgBoxStyle.OKOnly, "Bad Row Selected")
                Exit Sub
            End If
            'If IsNothing(frmLateFeeWaiver) Then
            '    frmLateFeeWaiver = New LateFeeWaiverRequest
            '    AddHandler frmLateFeeWaiver.Closing, AddressOf frmClosing
            '    AddHandler frmLateFeeWaiver.Closed, AddressOf frmClosed
            '    frmLateFeeWaiver.Text = "Late Waiver Request Maintenance"
            'End If
            'frmLateFeeWaiver.LateCertID = ugInvoices.ActiveRow.Cells("Late_Cert_ID").Value
            'frmLateFeeWaiver.ShowDialog()
            'LoadugInvoices()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Function GetVisibleRowCount(ByVal ugTable As Infragistics.Win.UltraWinGrid.UltraGrid) As Int32
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim iCount As Int32
        Try

            iCount = 0
            For Each ugrow In ugTable.Rows
                If ugrow.Hidden = False Then
                    iCount += 1
                End If
            Next
            Return iCount
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Function
    Private Sub btnGenCreditMemo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGenCreditMemo.Click
        Dim oInv As New MUSTER.BusinessLogic.pFeeInvoice

        Try
            If GetVisibleRowCount(ugInvoices) = 0 Then
                MsgBox("No Invoices Available For Credit Memo", MsgBoxStyle.OKOnly, "Invalid Invoice")
                Exit Sub
            End If
            If IsNothing(frmGenCreditMemo) Then
                frmGenCreditMemo = New GenerateCreditMemo


                AddHandler frmGenCreditMemo.Closing, AddressOf frmClosing
                AddHandler frmGenCreditMemo.Closed, AddressOf frmClosed
            End If

            oInv.Retrieve(ugInvoices.ActiveRow.Cells("INV_ID").Text)
            If IsDBNull(oInv.WarrantNumber) Then
                MsgBox("Invoice must have an Invoice Number from BP2K before a Credit Memo can be issued.", MsgBoxStyle.OKOnly, "Invalid Invoice Number")
                Exit Sub
            End If
            If oInv.WarrantNumber = "" Then
                MsgBox("Invoice must have an Invoice Number from BP2K before a Credit Memo can be issued.", MsgBoxStyle.OKOnly, "Invalid Invoice Number")
                Exit Sub
            End If
            frmGenCreditMemo.InvoiceID = oInv.ID
            frmGenCreditMemo.OwnerID = pOwn.ID
            frmGenCreditMemo.FacilityID = 0

            frmGenCreditMemo.ShowDialog()

            LoadugInvoices()
            RecalcOwnerTotals()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnReallocateOverage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReallocateOverage.Click
        Try
            If GetVisibleRowCount(ugInvoices) = 0 Then
                MsgBox("No Invoices Available For Reallocating Overage", MsgBoxStyle.OKOnly, "Invalid Invoice")
                Exit Sub
            End If
            If IsNothing(frmReallocateOverage) Then
                frmReallocateOverage = New ReallocateOwnerOverage
                frmReallocateOverage.InvoiceID = ugInvoices.ActiveRow.Cells("INV_ID").Text
                frmReallocateOverage.OwnerID = pOwn.ID

                AddHandler frmReallocateOverage.Closing, AddressOf frmClosing
                AddHandler frmReallocateOverage.Closed, AddressOf frmClosed
            End If
            frmReallocateOverage.ShowDialog()
            LoadugInvoices()
            RecalcOwnerTotals()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnGenerateInvoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGenerateInvoice.Click
        Try
            PrepareManualInvoicePanel()
            pnlManualInvoice.Visible = True
            pnlInvoices.Visible = False
            RecalcOwnerTotals()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnManualInvoiceCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnManualInvoiceCancel.Click
        Try
            pnlManualInvoice.Visible = False
            pnlInvoices.Visible = True
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnGenerateDebitMemo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnGenerateDebitMemo.Click
        Try
            If IsNothing(frmGenDebitMemo) Then
                frmGenDebitMemo = New GenerateDebitMemo
                AddHandler frmGenDebitMemo.Closing, AddressOf frmClosing
                AddHandler frmGenDebitMemo.Closed, AddressOf frmClosed
            End If
            frmGenDebitMemo.OwnerID = pOwn.ID
            frmGenDebitMemo.ShowDialog()
            RecalcOwnerTotals()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#End Region
#Region "Receipts"
#Region "UI Support Routines"

#End Region
#Region "UI Control Events"
    Private Sub btnOverpaymentReason_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOverpaymentReason.Click
        Dim Childrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim PaymentID As Int64
        Try
            If ugReceipts.ActiveRow Is Nothing Then Exit Sub

            If IsNothing(frmOverpayment) Then
                frmOverpayment = New OverpaymentReason
                AddHandler frmOverpayment.Closing, AddressOf frmClosing
                AddHandler frmOverpayment.Closed, AddressOf frmClosed
            End If

            If ugReceipts.ActiveRow.Band.Index = 0 Then
                PaymentID = 0
                For Each Childrow In ugReceipts.ActiveRow.ChildBands(0).Rows()
                    If Childrow.Cells("Facility_ID").Text.Trim.ToUpper = "OVERPAYMENT" Then
                        PaymentID = Childrow.Cells("RECPT_ID").Value
                        Exit For
                    End If
                Next
                If PaymentID = 0 Then
                    MsgBox("Selected receipt group does not contain an overpayment.", MsgBoxStyle.OKOnly, "Incorrect Row")
                    Exit Sub
                End If
            Else
                If ugReceipts.ActiveRow.Cells("Facility_ID").Text.Trim.ToUpper = "OVERPAYMENT" Then
                    PaymentID = ugReceipts.ActiveRow.Cells("RECPT_ID").Value
                Else
                    MsgBox("Selected detail line is not an overpayment.", MsgBoxStyle.OKOnly, "Incorrect Row")
                    Exit Sub
                End If
            End If
            frmOverpayment.PaymentID = PaymentID
            frmOverpayment.OwnerID = pOwn.ID
            frmOverpayment.ShowDialog()

            LoadugReceipts()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#End Region
#Region "Refunds"
#Region "UI Support Routines"

#End Region
#Region "UI Control Events"
    Private Sub btnGenerateRefund_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnGenerateRefund.Click
        Try
            If IsNothing(frmGenRefund) Then
                frmGenRefund = New GenerateRefund
                AddHandler frmGenRefund.Closing, AddressOf frmClosing
                AddHandler frmGenRefund.Closed, AddressOf frmClosed
            End If
            frmGenRefund.OwnerID = pOwn.ID
            frmGenRefund.ShowDialog()
            RecalcOwnerTotals()
            LoadugRefunds()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
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
        ' set the bool to false if comment is from event
        Dim bolEnableShowAllModules As Boolean = True
        Dim nCommentsCount As Integer = 0
        Try
            Select Case Me.tbCntrlFees.SelectedTab.Name
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
                Case Else
                    Exit Sub
            End Select
            If Not resetBtnColor Then
                SC = New ShowComments(nEntityID, nEntityType, IIf(bolSetCounts, "", "Fees"), strEntityName, oComments, Me.Text, , bolEnableShowAllModules)
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
        MyFrm.FlagsChanged(entityID, entityType, [Module], ParentFormText)
    End Sub
    Private Sub SF_RefreshCalendar() Handles SF.RefreshCalendar
        Dim mc As MusterContainer = Me.MdiParent
        mc.RefreshCalendarInfo()
        mc.LoadDueToMeCalendar()
        mc.LoadToDoCalendar()
    End Sub
    Private Sub FlagMaintenance(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Select Case Me.tbCntrlFees.SelectedTab.Name
                Case tbPageOwnerDetail.Name
                    SF = New ShowFlags(pOwn.ID, UIUtilsGen.EntityTypes.Owner, "Fees")
                Case tbPageFacilityDetail.Name
                    SF = New ShowFlags(pOwn.Facilities.ID, UIUtilsGen.EntityTypes.Facility, "Fees")
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
        'MyFrm.FlagsChanged(Me.lblOwnerIDValue.Text, 0, 0, 0, "Fees", Me.Text)
    End Sub
    Private Sub btnFacFlags_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFacFlags.Click
        FlagMaintenance(sender, e)
        'Dim MyFrm As MusterContainer
        'MyFrm = Me.MdiParent
        'oEntity.GetEntity("Facility")
        'MyFrm.FlagsChanged(lblOwnerIDValue.Text, lblFacilityIDValue.Text, 0, 0, "Fees", Me.Text)
    End Sub
#End Region
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
            objCntSearch = New ContactSearch(nOwnerID, UIUtilsGen.EntityTypes.Owner, "FEES", pConStruct)
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
        Try
            SetOwnerFilter()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub chkOwnerShowContactsforAllModules_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkOwnerShowContactsforAllModules.CheckedChanged
        Try
            SetOwnerFilter()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub chkOwnerShowRelatedContacts_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkOwnerShowRelatedContacts.CheckedChanged
        Try
            SetOwnerFilter()
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
                nModuleID = 892
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
            objCntSearch = New ContactSearch(nFacilityID, UIUtilsGen.EntityTypes.Facility, "FEES", pConStruct)
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
            If tbCntrlFees.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL" Then
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
                    strFilterString += " MODULEID = 892 And ENTITYID = " + strEntityID
                Else
                    strFilterString += " AND MODULEID = 892 And ENTITYID = " + strEntityID
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
                nModuleID = 892
            End If

            If chkFacilityShowRelatedContacts.Checked Then
                strEntities = strFacilityIdTags
                nRelatedEntityType = 6
            Else
                strEntities = String.Empty
            End If

            UIUtilsGen.LoadContacts(ugOwnerContacts, nEntityID, nEntityType, pConStruct, nModuleID, strEntities, bolActive, strEntityAssocIDs, nRelatedEntityType)

        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
#End Region
#Region "Common Functions"
    Private Sub LoadContacts(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal EntityID As Integer, ByVal EntityType As Integer)

        Try

            UIUtilsGen.LoadContacts(ugGrid, EntityID, EntityType, pConStruct, 892)

            If tbCntrlFees.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL" Then
                Me.chkOwnerShowActiveOnly.Checked = False
                Me.chkOwnerShowActiveOnly.Checked = True
            ElseIf tbCntrlFees.SelectedTab.Name.ToUpper = "TBPAGEFACILITYDETAIL" Then
                Me.chkFacilityShowActiveContactOnly.Checked = False
                Me.chkFacilityShowActiveContactOnly.Checked = True
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Function ModifyContact(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid)
        Try

            If UIUtilsGen.ModifyContact(ugGrid, 892, pConStruct) Then
                Me.Contact_ContactAdded()
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function

    Private Function AssociateContact(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal nEntityID As Integer, ByVal nEntityType As Integer)
        Try
            If UIUtilsGen.AssociateContact(ugGrid, nEntityID, nEntityType, 892, pConStruct) Then
                Me.Contact_ContactAdded()
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function

    Private Function DeleteContact(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal nEntityID As Integer)
        Try

            If UIUtilsGen.DeleteContact(ugGrid, nEntityID, 892, pConStruct) Then
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
        If tbCntrlFees.SelectedTab.Name = tbPageOwnerDetail.Name Then
            LoadContacts(ugOwnerContacts, nOwnerID, UIUtilsGen.EntityTypes.Owner)
            SetOwnerFilter()
        Else
            LoadContacts(ugOwnerContacts, nFacilityID, UIUtilsGen.EntityTypes.Facility)
            SetFacilityFilter()
        End If
    End Sub
    Private Sub Contact_ContactAdded()
        Try
            If tbCntrlFees.SelectedTab.Name.ToUpper = "TBPAGEOWNERDETAIL".ToUpper Then
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

    Public ReadOnly Property MC() As MusterContainer
        Get
            mContainer = Me.MdiParent
            If mContainer Is Nothing Then
                mContainer = New MusterContainer
            End If
            Return mContainer
        End Get
    End Property

    Private Sub ProcessFacTransactions()
        Select Case tbCtrlFacTransactions.SelectedIndex
            Case 0  'By Transaction
                LoadugByTransaction()

            Case 1  'By Invoice
                LoadugByInvoice()

            Case 2  ' Old Access Info
                LoadugOldAccessHistory()
        End Select
    End Sub


    Private Sub LoadugOldAccessHistory()
        Dim oFeeInvoice As New MUSTER.BusinessLogic.pFeeInvoice
        Dim dsLocal As DataSet
        Dim tmpBand As Int16
        Dim rp As New Remove_Pencil
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim sBalance As Single

        dsLocal = oFeeInvoice.GetFacilityOldAccessInfo_ByFacility(lblFacilityIDValue.Text)
        ugOldAccessHistory.DrawFilter = rp
        ugOldAccessHistory.DataSource = dsLocal
        ugOldAccessHistory.Rows.CollapseAll(True)
        ugOldAccessHistory.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        ugOldAccessHistory.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        ugOldAccessHistory.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        sBalance = 0

        If dsLocal.Tables(0).Rows.Count > 0 Then ' Does Table have rows
            ugOldAccessHistory.Visible = True

            ugOldAccessHistory.DisplayLayout.Bands(0).Columns("Balance").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugOldAccessHistory.DisplayLayout.Bands(0).Columns("Balance").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            ugOldAccessHistory.DisplayLayout.Bands(0).Columns("Debit").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugOldAccessHistory.DisplayLayout.Bands(0).Columns("Debit").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            ugOldAccessHistory.DisplayLayout.Bands(0).Columns("Credit").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugOldAccessHistory.DisplayLayout.Bands(0).Columns("Credit").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            For Each ugrow In ugOldAccessHistory.Rows

                If ugrow.Cells("Debit").Value <> "" Then
                    sBalance += ugrow.Cells("Debit").Value
                End If

                If ugrow.Cells("Credit").Value <> "" Then
                    sBalance -= ugrow.Cells("Credit").Value
                End If

                ugrow.Cells("Balance").Value = FormatNumber(sBalance, 2, TriState.True, TriState.False, TriState.True)

                If ugrow.Cells("Credit").Value <> "" Then
                    ugrow.Cells("Credit").Appearance.ForeColor = System.Drawing.Color.Red
                End If

                If sBalance < 0 Then
                    ugrow.Cells("Balance").Appearance.ForeColor = System.Drawing.Color.Red
                End If


            Next

        Else
            ugOldAccessHistory.Visible = False

        End If

    End Sub
    '
    Private Sub LoadugFacilityList()
        Dim oFeeInvoice As New MUSTER.BusinessLogic.pFeeInvoice
        Dim dsLocal As DataSet
        Dim tmpBand As Int16
        Dim RowCount As Int64
        Dim SumPriorBalance As Single
        Dim SumCurrentFees As Single
        Dim SumLateFees As Single
        Dim SumTotalDue As Single
        Dim SumCurrentPayments As Single
        Dim SumCurrentCredits As Single
        Dim SumCurrentAdjustments As Single
        Dim SumLegal As Single
        Dim SumToDateBalance As Single
        Dim tmpDate As Date

        Dim strTemp As String

        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow

        txtBP2KOwnerID.Text = pOwn.BP2KOwnerID
        txtChapter.Text = pOwn.BankruptChapter
        If pOwn.BankruptDate <> tmpDate Then
            txtDate.Text = pOwn.BankruptDate
        Else
            txtDate.Text = ""
        End If


        If IsNothing(pOwn) Then Exit Sub

        dsLocal = oFeeInvoice.GetFacilitySummaryGrid_ByOwnerID(pOwn.ID)
        ugFacilityList.DataSource = dsLocal
        ugFacilityList.Rows.CollapseAll(True)
        ugFacilityList.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        ugFacilityList.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        'ugFacilityList.DisplayLayout.AutoFitColumns = True
        If dsLocal.Tables(0).Rows.Count > 0 Then ' Does Table have rows
            ugFacilityList.Visible = True
            ugFacilityList.DisplayLayout.Bands(0).Columns("Owner_ID").Hidden = True

            ugFacilityList.DisplayLayout.Bands(0).Columns("Facility_ID").Header.Caption = "Facility ID"
            ugFacilityList.DisplayLayout.Bands(0).Columns("Name").Header.Caption = "Facility Name"
            ugFacilityList.DisplayLayout.Bands(0).Columns("PriorBalance").Header.Caption = "Previous" '"Prior Balance"
            ugFacilityList.DisplayLayout.Bands(0).Columns("CurrentFees").Header.Caption = "Charges" '"Current Fees"
            ugFacilityList.DisplayLayout.Bands(0).Columns("LateFees").Header.Caption = "Late Fees" '"Late Penalty"
            ugFacilityList.DisplayLayout.Bands(0).Columns("TotalDue").Header.Caption = "Total Due"
            ugFacilityList.DisplayLayout.Bands(0).Columns("CurrentPayments").Header.Caption = "Payments" '"Current Payments"
            ugFacilityList.DisplayLayout.Bands(0).Columns("CurrentCredits").Header.Caption = "Credits" '"Current Credits"
            ugFacilityList.DisplayLayout.Bands(0).Columns("CurrentAdjustments").Header.Caption = "Adjustments" '"Current Adjustments"
            ugFacilityList.DisplayLayout.Bands(0).Columns("Legal").Header.Caption = "Legal" '"Current Legal"
            ugFacilityList.DisplayLayout.Bands(0).Columns("ToDateBalance").Header.Caption = "Balance" '"Facility To Date Balance"

            'ugFacilityList.DisplayLayout.Bands(0).Columns("Date_Created").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            ugFacilityList.DisplayLayout.Bands(0).Columns("Facility_ID").Width = 50
            ugFacilityList.DisplayLayout.Bands(0).Columns("Name").Width = 140
            ugFacilityList.DisplayLayout.Bands(0).Columns("City").Width = 90

            ugFacilityList.DisplayLayout.Bands(0).Columns("PriorBalance").Width = 75
            ugFacilityList.DisplayLayout.Bands(0).Columns("CurrentFees").Width = 75
            ugFacilityList.DisplayLayout.Bands(0).Columns("LateFees").Width = 75
            ugFacilityList.DisplayLayout.Bands(0).Columns("TotalDue").Width = 75
            ugFacilityList.DisplayLayout.Bands(0).Columns("CurrentPayments").Width = 75
            ugFacilityList.DisplayLayout.Bands(0).Columns("CurrentCredits").Width = 75
            ugFacilityList.DisplayLayout.Bands(0).Columns("CurrentPayments").Width = 75
            ugFacilityList.DisplayLayout.Bands(0).Columns("CurrentAdjustments").Width = 75
            ugFacilityList.DisplayLayout.Bands(0).Columns("Legal").Width = 75
            ugFacilityList.DisplayLayout.Bands(0).Columns("ToDateBalance").Width = 75

            ugFacilityList.DisplayLayout.Bands(0).Columns("PriorBalance").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugFacilityList.DisplayLayout.Bands(0).Columns("PriorBalance").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            ugFacilityList.DisplayLayout.Bands(0).Columns("CurrentFees").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugFacilityList.DisplayLayout.Bands(0).Columns("CurrentFees").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            ugFacilityList.DisplayLayout.Bands(0).Columns("LateFees").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugFacilityList.DisplayLayout.Bands(0).Columns("LateFees").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            ugFacilityList.DisplayLayout.Bands(0).Columns("TotalDue").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugFacilityList.DisplayLayout.Bands(0).Columns("TotalDue").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            ugFacilityList.DisplayLayout.Bands(0).Columns("CurrentPayments").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugFacilityList.DisplayLayout.Bands(0).Columns("CurrentPayments").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            ugFacilityList.DisplayLayout.Bands(0).Columns("CurrentCredits").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugFacilityList.DisplayLayout.Bands(0).Columns("CurrentCredits").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            ugFacilityList.DisplayLayout.Bands(0).Columns("CurrentAdjustments").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugFacilityList.DisplayLayout.Bands(0).Columns("CurrentAdjustments").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            ugFacilityList.DisplayLayout.Bands(0).Columns("Legal").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugFacilityList.DisplayLayout.Bands(0).Columns("Legal").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            ugFacilityList.DisplayLayout.Bands(0).Columns("ToDateBalance").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugFacilityList.DisplayLayout.Bands(0).Columns("ToDateBalance").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            RowCount = 0
            SumPriorBalance = 0
            SumCurrentFees = 0
            SumLateFees = 0
            SumTotalDue = 0
            SumCurrentPayments = 0
            SumCurrentCredits = 0
            SumCurrentAdjustments = 0
            SumLegal = 0
            SumToDateBalance = 0

            strTemp = ""

            For Each ugrow In ugFacilityList.Rows
                RowCount += 1

                If ugrow.Cells("PriorBalance").Value < 0 Then
                    ugrow.Cells("PriorBalance").Appearance.ForeColor = System.Drawing.Color.Red
                End If

                SumPriorBalance += ugrow.Cells("PriorBalance").Value
                SumCurrentFees += ugrow.Cells("CurrentFees").Value
                SumLateFees += ugrow.Cells("LateFees").Value
                SumTotalDue += ugrow.Cells("TotalDue").Value

                If ugrow.Cells("CurrentPayments").Value > 0 Then
                    ugrow.Cells("CurrentPayments").Appearance.ForeColor = System.Drawing.Color.Red
                End If
                SumCurrentPayments += ugrow.Cells("CurrentPayments").Value

                If ugrow.Cells("CurrentCredits").Value > 0 Then
                    ugrow.Cells("CurrentCredits").Appearance.ForeColor = System.Drawing.Color.Red
                End If
                SumCurrentCredits += ugrow.Cells("CurrentCredits").Value

                If ugrow.Cells("CurrentAdjustments").Value < 0 Then
                    ugrow.Cells("CurrentAdjustments").Appearance.ForeColor = System.Drawing.Color.Red
                End If
                SumCurrentAdjustments += ugrow.Cells("CurrentAdjustments").Value

                If ugrow.Cells("Legal").Value > 0 Then
                    ugrow.Cells("Legal").Appearance.ForeColor = System.Drawing.Color.Red
                End If
                SumLegal += ugrow.Cells("Legal").Value


                If ugrow.Cells("ToDateBalance").Value < 0 Then
                    ugrow.Cells("ToDateBalance").Appearance.ForeColor = System.Drawing.Color.Red
                End If
                SumToDateBalance += ugrow.Cells("ToDateBalance").Value

            Next

            lblNoOfFacilities.Text = "Owner Totals for " & RowCount
            If RowCount = 1 Then
                lblNoOfFacilities.Text += " Facility: "
            Else
                lblNoOfFacilities.Text += " Facilities: "
            End If
            lblPriorBalance.Text = FormatNumber(SumPriorBalance, 2, TriState.True, TriState.False, TriState.True)
            If SumPriorBalance < 0 Then
                lblPriorBalance.ForeColor = System.Drawing.Color.Red
            Else
                lblPriorBalance.ForeColor = System.Drawing.Color.Black
            End If

            lblCurrentFees.Text = FormatNumber(SumCurrentFees, 2, TriState.True, TriState.False, TriState.True)
            lblLatePenalty.Text = FormatNumber(SumLateFees, 2, TriState.True, TriState.False, TriState.True)
            lblTotalDue.Text = FormatNumber(SumTotalDue, 2, TriState.True, TriState.False, TriState.True)

            lblCurrentPayments.Text = FormatNumber(SumCurrentPayments, 2, TriState.True, TriState.False, TriState.True)
            If SumCurrentPayments > 0 Then
                lblCurrentPayments.ForeColor = System.Drawing.Color.Red
            Else
                lblCurrentPayments.ForeColor = System.Drawing.Color.Black
            End If

            lblCurrentCredits.Text = FormatNumber(SumCurrentCredits, 2, TriState.True, TriState.False, TriState.True)
            If SumCurrentCredits > 0 Then
                lblCurrentCredits.ForeColor = System.Drawing.Color.Red
            Else
                lblCurrentCredits.ForeColor = System.Drawing.Color.Black
            End If

            lblCurrentAdjustments.Text = FormatNumber(SumCurrentAdjustments, 2, TriState.True, TriState.False, TriState.True)
            If SumCurrentAdjustments < 0 Then
                lblCurrentAdjustments.ForeColor = System.Drawing.Color.Red
            Else
                lblCurrentAdjustments.ForeColor = System.Drawing.Color.Black
            End If

            lblLegal.Text = FormatNumber(SumLegal, 2, TriState.True, TriState.False, TriState.True)
            If SumLegal > 0 Then
                lblLegal.ForeColor = System.Drawing.Color.Red
            Else
                lblLegal.ForeColor = System.Drawing.Color.Black
            End If

            lblFacToDateBalance.Text = FormatNumber(SumToDateBalance, 2, TriState.True, TriState.False, TriState.True)
            If lblFacToDateBalance.Text < 0 Then
                lblFacToDateBalance.ForeColor = System.Drawing.Color.Red
            Else
                lblFacToDateBalance.ForeColor = System.Drawing.Color.Black
            End If
        Else
            ugFacilityList.Visible = False
        End If

    End Sub


    Private Sub LoadugInvoices()
        Dim oFeeInvoice As New MUSTER.BusinessLogic.pFeeInvoice
        Dim dsLocal As DataSet
        Dim tmpBand As Int16

        If IsNothing(pOwn) Then Exit Sub

        dsLocal = oFeeInvoice.GetFacilitySummaryInvoiceGrid_ByOwnerID(pOwn.ID)
        ugInvoices.DataSource = dsLocal
        ugInvoices.Rows.CollapseAll(True)
        ugInvoices.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        ugInvoices.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        ugInvoices.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        'ugInvoices.DisplayLayout.AutoFitColumns = True
        If dsLocal.Tables(0).Rows.Count > 0 Then ' Does Table have rows
            ugInvoices.Visible = True

            ugInvoices.DisplayLayout.Bands(0).Columns("Owner_ID").Hidden = True
            ugInvoices.DisplayLayout.Bands(0).Columns("INV_ID").Hidden = True
            ugInvoices.DisplayLayout.Bands(0).Columns("SFY").Hidden = True
            ugInvoices.DisplayLayout.Bands(0).Columns("Late_Cert_ID").Hidden = True

            ugInvoices.DisplayLayout.Bands(0).Columns("InvoiceID").Header.Caption = "Invoice Number"
            ugInvoices.DisplayLayout.Bands(0).Columns("Fee_Type").Header.Caption = "Fee Type"
            ugInvoices.DisplayLayout.Bands(0).Columns("InvoiceDate").Header.Caption = "Invoice Date"
            ugInvoices.DisplayLayout.Bands(0).Columns("Due_Date").Header.Caption = "Due Date"
            ugInvoices.DisplayLayout.Bands(0).Columns("Balance").Header.Caption = "Invoice Balance"

            ugInvoices.DisplayLayout.Bands(0).Columns("InvoiceID").Width = 200
            ugInvoices.DisplayLayout.Bands(0).Columns("InvoiceDate").Width = 130
            ugInvoices.DisplayLayout.Bands(0).Columns("Due_Date").Width = 80
            ugInvoices.DisplayLayout.Bands(0).Columns("Fee_Type").Width = 85
            ugInvoices.DisplayLayout.Bands(0).Columns("Facilities").Width = 55
            ugInvoices.DisplayLayout.Bands(0).Columns("Quantity").Width = 55

            ugInvoices.DisplayLayout.Bands(0).Columns("Charges").Width = 75
            ugInvoices.DisplayLayout.Bands(0).Columns("Payments").Width = 75
            ugInvoices.DisplayLayout.Bands(0).Columns("Credits").Width = 75
            ugInvoices.DisplayLayout.Bands(0).Columns("Adjustments").Width = 75
            ugInvoices.DisplayLayout.Bands(0).Columns("Legal").Width = 75
            ugInvoices.DisplayLayout.Bands(0).Columns("Balance").Width = 135

            ugInvoices.DisplayLayout.Bands(0).Columns("AdviceNumber").Width = 100



            ugInvoices.DisplayLayout.Bands(0).Columns("Facilities").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugInvoices.DisplayLayout.Bands(0).Columns("Facilities").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            ugInvoices.DisplayLayout.Bands(0).Columns("Quantity").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugInvoices.DisplayLayout.Bands(0).Columns("Quantity").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            ugInvoices.DisplayLayout.Bands(0).Columns("Charges").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugInvoices.DisplayLayout.Bands(0).Columns("Charges").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            ugInvoices.DisplayLayout.Bands(0).Columns("Payments").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugInvoices.DisplayLayout.Bands(0).Columns("Payments").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            ugInvoices.DisplayLayout.Bands(0).Columns("Credits").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugInvoices.DisplayLayout.Bands(0).Columns("Credits").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            ugInvoices.DisplayLayout.Bands(0).Columns("Adjustments").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugInvoices.DisplayLayout.Bands(0).Columns("Adjustments").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            ugInvoices.DisplayLayout.Bands(0).Columns("Legal").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugInvoices.DisplayLayout.Bands(0).Columns("Legal").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            ugInvoices.DisplayLayout.Bands(0).Columns("Balance").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugInvoices.DisplayLayout.Bands(0).Columns("Balance").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            If dsLocal.Tables(1).Rows.Count > 0 Then ' Does Table have rows
                ugInvoices.DisplayLayout.Bands(1).Columns("Owner_ID").Hidden = True
                ugInvoices.DisplayLayout.Bands(1).Columns("Due_Date").Hidden = True
                ugInvoices.DisplayLayout.Bands(1).Columns("INV_ID").Hidden = True
                ugInvoices.DisplayLayout.Bands(1).Columns("Advice_ID").Hidden = True
                ugInvoices.DisplayLayout.Bands(1).Columns("InvoiceID").Hidden = True
                ugInvoices.DisplayLayout.Bands(1).Columns("InvoiceDate").Hidden = True
                ugInvoices.DisplayLayout.Bands(1).Columns("UnitPrice").Hidden = True
                ugInvoices.DisplayLayout.Bands(1).Columns("SFY").Hidden = True
                ugInvoices.DisplayLayout.Bands(1).Columns("ITEM_SEQ_NUMBER").Hidden = True

                'ugInvoices.DisplayLayout.Bands(1).Columns("FacilityName").ColSpan = 3
                ugInvoices.DisplayLayout.Bands(1).Columns("Facility_ID").Width = 75
                ugInvoices.DisplayLayout.Bands(1).Columns("FacilityName").Width = 250
                ugInvoices.DisplayLayout.Bands(1).Columns("Fiscal_Year").Width = 55
                ugInvoices.DisplayLayout.Bands(1).Columns("Quantity").Width = 55

                ugInvoices.DisplayLayout.Bands(1).Columns("Charges").Width = 75
                ugInvoices.DisplayLayout.Bands(1).Columns("Payments").Width = 75
                ugInvoices.DisplayLayout.Bands(1).Columns("Credits").Width = 75
                ugInvoices.DisplayLayout.Bands(1).Columns("Adjustments").Width = 75
                ugInvoices.DisplayLayout.Bands(1).Columns("Legal").Width = 75
                ugInvoices.DisplayLayout.Bands(1).Columns("Balance").Width = 75

                ugInvoices.DisplayLayout.Bands(1).Columns("Linedescription").Width = 100

                ugInvoices.DisplayLayout.Bands(1).Columns("Facility_ID").Header.Caption = "Facility ID"
                ugInvoices.DisplayLayout.Bands(1).Columns("FacilityName").Header.Caption = "Facility Name"
                ugInvoices.DisplayLayout.Bands(1).Columns("Fiscal_Year").Header.Caption = "FY"

                ugInvoices.DisplayLayout.Bands(1).Columns("Quantity").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                ugInvoices.DisplayLayout.Bands(1).Columns("Quantity").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

                ugInvoices.DisplayLayout.Bands(1).Columns("Charges").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                ugInvoices.DisplayLayout.Bands(1).Columns("Charges").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

                ugInvoices.DisplayLayout.Bands(1).Columns("Payments").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                ugInvoices.DisplayLayout.Bands(1).Columns("Payments").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

                ugInvoices.DisplayLayout.Bands(1).Columns("Credits").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                ugInvoices.DisplayLayout.Bands(1).Columns("Credits").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

                ugInvoices.DisplayLayout.Bands(1).Columns("Adjustments").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                ugInvoices.DisplayLayout.Bands(1).Columns("Adjustments").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

                ugInvoices.DisplayLayout.Bands(1).Columns("Legal").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                ugInvoices.DisplayLayout.Bands(1).Columns("Legal").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

                ugInvoices.DisplayLayout.Bands(1).Columns("Balance").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                ugInvoices.DisplayLayout.Bands(1).Columns("Balance").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center
            End If
            ProcessInvoiceGrid()
        Else
            ugInvoices.Visible = False
        End If



    End Sub

    Private Sub LoadugReceipts()
        Dim oFeeReceipt As New MUSTER.BusinessLogic.pFeeReceipt
        Dim dsLocal As DataSet
        Dim tmpBand As Int16

        If IsNothing(pOwn) Then Exit Sub

        dsLocal = oFeeReceipt.GetOwnerSummaryReceiptGrid_ByOwnerID(pOwn.ID)
        ugReceipts.DataSource = dsLocal
        ugReceipts.Rows.CollapseAll(True)
        ugReceipts.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        ugReceipts.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        ugReceipts.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        ugReceipts.DisplayLayout.AutoFitColumns = True
        If dsLocal.Tables(0).Rows.Count > 0 Then ' Does Table have rows
            ugReceipts.Visible = True
            ugReceipts.DisplayLayout.Bands(0).Columns("Owner_ID").Hidden = True
            ugReceipts.DisplayLayout.Bands(0).Columns("Fiscal_Year").Hidden = True

            ugReceipts.DisplayLayout.Bands(0).Columns("Amount").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugReceipts.DisplayLayout.Bands(0).Columns("Amount").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            ugReceipts.DisplayLayout.Bands(0).Columns("Receipt_Date").Header.Caption = "Date Received"
            ugReceipts.DisplayLayout.Bands(0).Columns("Check_Number").Header.Caption = "Check Number"
            ugReceipts.DisplayLayout.Bands(0).Columns("Check_Trans_ID").Header.Caption = "Trans ID"
            'ugReceipts.DisplayLayout.Bands(0).Columns("Trans_Type").Header.Caption = "Trans Type"
            ugReceipts.DisplayLayout.Bands(0).Columns("IssueCompany_Reason").Header.Caption = "Issue Company / Reason"

            If dsLocal.Tables(1).Rows.Count > 0 Then ' Does Table have rows
                ugReceipts.DisplayLayout.Bands(1).Columns("Owner_ID").Hidden = True
                ugReceipts.DisplayLayout.Bands(1).Columns("RECPT_ID").Hidden = True
                ugReceipts.DisplayLayout.Bands(1).Columns("Check_Trans_ID").Hidden = True
                ugReceipts.DisplayLayout.Bands(1).Columns("Misapply_Flag").Hidden = True

                ugReceipts.DisplayLayout.Bands(1).Columns("Facility_ID").Header.Caption = "Facility ID"
                ugReceipts.DisplayLayout.Bands(1).Columns("Fiscal_Year").Header.Caption = "Fiscal Year"
                ugReceipts.DisplayLayout.Bands(1).Columns("INV_Number").Header.Caption = "Invoice"
                ugReceipts.DisplayLayout.Bands(1).Columns("Fee_Type").Header.Caption = "Fee Type"
                ugReceipts.DisplayLayout.Bands(1).Columns("Facility_Name").Header.Caption = "Facility Name\Overpayment Reason"

                ugReceipts.DisplayLayout.Bands(1).Columns("Amount").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                ugReceipts.DisplayLayout.Bands(1).Columns("Amount").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            End If

            ProcessReceiptGrid()
        Else
            ugReceipts.Visible = False
        End If


    End Sub


    Private Sub LoadugRefunds()
        Dim oFeeInvoice As New MUSTER.BusinessLogic.pFeeInvoice
        Dim dsLocal As DataSet
        Dim tmpBand As Int16


        If IsNothing(pOwn) Then Exit Sub

        dsLocal = oFeeInvoice.GetRefunds_ByOwnerID(pOwn.ID)
        ugRefunds.DataSource = dsLocal
        ugRefunds.Rows.CollapseAll(True)
        ugRefunds.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        ugRefunds.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        ugRefunds.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        If dsLocal.Tables(0).Rows.Count > 0 Then ' Does Table have rows
            ugRefunds.Visible = True
            ugRefunds.DisplayLayout.Bands(0).Columns("Owner_ID").Hidden = True

            ugRefunds.DisplayLayout.Bands(0).Columns("Issued_To").Header.Caption = "Issued To"
            ugRefunds.DisplayLayout.Bands(0).Columns("Check_Trans_ID").Header.Caption = "Trans ID"


            ugRefunds.DisplayLayout.Bands(0).Columns("Debit").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugRefunds.DisplayLayout.Bands(0).Columns("Debit").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            ugRefunds.DisplayLayout.Bands(0).Columns("Credits").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugRefunds.DisplayLayout.Bands(0).Columns("Credits").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            ugRefunds.DisplayLayout.Bands(0).Columns("Correction").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugRefunds.DisplayLayout.Bands(0).Columns("Correction").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

        Else
            ugRefunds.Visible = False
        End If


    End Sub

    Private Sub LoadugOverages()

        Dim oFeeAdjustment As New MUSTER.BusinessLogic.pFeeAdjustment
        Dim dsLocal As DataSet
        Dim tmpBand As Int16


        If IsNothing(pOwn) Then Exit Sub

        dsLocal = oFeeAdjustment.RetrieveOwnerOverages(pOwn.ID)
        ugOverage.DataSource = dsLocal
        'ugOverage.Rows.CollapseAll(True)
        'ugOverage.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        'ugOverage.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        'ugOverage.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        If dsLocal.Tables(0).Rows.Count > 0 Then ' Does Table have rows
            ugOverage.Visible = True
            'ugOverage.DisplayLayout.Bands(0).Columns("Owner_ID").Hidden = True
            ugOverage.DisplayLayout.Bands(0).Columns("SFY").Header.Caption = "SFY"
            ugOverage.DisplayLayout.Bands(0).Columns("OWNER_ID").Header.Caption = "Owner ID"
            ugOverage.DisplayLayout.Bands(0).Columns("FACILITY_ID").Header.Caption = "Facility ID"
            ugOverage.DisplayLayout.Bands(0).Columns("INV_NUMBER").Header.Caption = "Invoice #"
            ugOverage.DisplayLayout.Bands(0).Columns("INV_NUMBER").Width = 200


            ugOverage.DisplayLayout.Bands(0).Columns("INV_AMT").Header.Caption = "Invoice Amount"
            ugOverage.DisplayLayout.Bands(0).Columns("INV_AMT").Width = 130


            ugOverage.DisplayLayout.Bands(0).Columns("DATE_APPLIED").Header.Caption = "Date Applied"
            ugOverage.DisplayLayout.Bands(0).Columns("CHECK_NUMBER").Header.Caption = "Check #"
            ugOverage.DisplayLayout.Bands(0).Columns("BP2K_Trans_ID").Header.Caption = "BP2K Trans ID"
            ugOverage.DisplayLayout.Bands(0).Columns("TYPE").Header.Caption = "Type"

            ugOverage.DisplayLayout.Bands(0).Columns("SFY").Width = 75
            ugOverage.DisplayLayout.Bands(0).Columns("CHECK_NUMBER").Width = 100


        Else
            ugOverage.Visible = False
        End If


    End Sub

    Private Sub LoadugByTransaction()
        Dim oFeeInvoice As New MUSTER.BusinessLogic.pFeeInvoice
        Dim dsLocal As DataSet
        Dim tmpBand As Int16
        Dim rp As New Remove_Pencil

        dsLocal = oFeeInvoice.GetFacilityFeeTransaction_ByFacility(lblFacilityIDValue.Text)
        ugByTransaction.DrawFilter = rp
        ugByTransaction.DataSource = dsLocal
        ugByTransaction.Rows.CollapseAll(True)
        ugByTransaction.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        ugByTransaction.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        ugByTransaction.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        ugByTransaction.DisplayLayout.AutoFitColumns = True

        If dsLocal.Tables(0).Rows.Count > 0 Then ' Does Table have rows
            ugByTransaction.Visible = True
            ugByTransaction.DisplayLayout.Bands(0).Columns("Facility_ID").Hidden = True
            ugByTransaction.DisplayLayout.Bands(0).Columns("DATE_CREATED").Hidden = True

            ugByTransaction.DisplayLayout.Bands(0).Columns("PRIORITY").Hidden = True

            ugByTransaction.DisplayLayout.Bands(0).Columns("FiscalYear").Header.Caption = "Fiscal Year"
            ugByTransaction.DisplayLayout.Bands(0).Columns("Inv_Date").Header.Caption = "Invoice Date"
            ugByTransaction.DisplayLayout.Bands(0).Columns("INV_Number").Header.Caption = "Invoice Number"
            ugByTransaction.DisplayLayout.Bands(0).Columns("INV_Number").Width = 150
            ugByTransaction.DisplayLayout.Bands(0).Columns("Inv_Date").Width = 120



            ugByTransaction.DisplayLayout.Bands(0).Columns("Due_Date").Header.Caption = "Due Date"
            ugByTransaction.DisplayLayout.Bands(0).Columns("Due_Date").Width = 120


            ugByTransaction.DisplayLayout.Bands(0).Columns("TransactionType").Header.Caption = "Transaction Type"
            ugByTransaction.DisplayLayout.Bands(0).Columns("Balance").Header.Caption = "Balance"

            ugByTransaction.DisplayLayout.Bands(0).Columns("Debit").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            ugByTransaction.DisplayLayout.Bands(0).Columns("Credit").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugByTransaction.DisplayLayout.Bands(0).Columns("Credit").CellAppearance.ForeColor = System.Drawing.Color.Red

            ugByTransaction.DisplayLayout.Bands(0).Columns("Balance").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
        Else
            ugByTransaction.Visible = False
        End If
        ProcessFeesByTransactionRadio()

    End Sub

    Private Sub LoadugByInvoice()
        Dim oFeeInvoice As New MUSTER.BusinessLogic.pFeeInvoice
        Dim dsLocal As DataSet
        Dim tmpBand As Int16
        Dim rp As New Remove_Pencil

        If lblFacilityIDValue.Text = "" Then
            Exit Sub
        End If

        If ugByInvoice.DataSource Is Nothing Then
            dsLocal = oFeeInvoice.GetFacilityFeeInvoice_ByFacility(lblFacilityIDValue.Text)

            ugByInvoice.DrawFilter = rp
            ugByInvoice.DataSource = dsLocal
            ugByInvoice.Rows.CollapseAll(True)
            ugByInvoice.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
            ugByInvoice.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
            ugByInvoice.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
            ugByInvoice.DisplayLayout.AutoFitColumns = True

            If dsLocal.Tables(0).Rows.Count > 0 Then ' Does Table have rows
                ugByInvoice.Visible = True
                ugByInvoice.DisplayLayout.Bands(0).Columns("Inv_ID").Hidden = True
                ugByInvoice.DisplayLayout.Bands(0).Columns("FiscalYear").Header.Caption = "Fiscal Year"
                ugByInvoice.DisplayLayout.Bands(0).Columns("INV_Number").Header.Caption = "Invoice Number"
                ugByInvoice.DisplayLayout.Bands(0).Columns("INV_Number").Width = 200
                ugByInvoice.DisplayLayout.Bands(0).Columns("Advice_ID").Header.Caption = "Advice ID"
                ugByInvoice.DisplayLayout.Bands(0).Columns("Advice_ID").Width = 130


                ugByInvoice.DisplayLayout.Bands(0).Columns("FeeTypeDesc").Header.Caption = "Fee Type"
                ugByInvoice.DisplayLayout.Bands(0).Columns("INV_Line_Amt").Header.Caption = "Invoice Amount"
                ugByInvoice.DisplayLayout.Bands(0).Columns("Balance").Header.Caption = "Invoice Balance"

                ugByInvoice.DisplayLayout.Bands(0).Columns("INV_Line_Amt").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                ugByInvoice.DisplayLayout.Bands(0).Columns("Balance").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                If dsLocal.Tables(1).Rows.Count > 0 Then ' Does Table have rows
                    ugByInvoice.DisplayLayout.Bands(1).Columns("FiscalYear").Header.Caption = "Fiscal Year"
                    ugByInvoice.DisplayLayout.Bands(1).Columns("INV_Number").Header.Caption = "Invoice Number"
                    ugByInvoice.DisplayLayout.Bands(1).Columns("INV_Number").Width = 150


                    ugByInvoice.DisplayLayout.Bands(1).Columns("INV_Date").Header.Caption = "Invoice Date"
                    ugByInvoice.DisplayLayout.Bands(1).Columns("Due_Date").Header.Caption = "Due Date"
                    ugByInvoice.DisplayLayout.Bands(1).Columns("INV_Date").Width = 170


                    ugByInvoice.DisplayLayout.Bands(1).Columns("TransactionType").Header.Caption = "Transaction Type"
                    ugByInvoice.DisplayLayout.Bands(1).Columns("Balance").Header.Caption = "Balance"
                    ugByInvoice.DisplayLayout.Bands(1).Columns("CreditAppliedTo").Header.Caption = "Credit/Debit Number"

                    ugByInvoice.DisplayLayout.Bands(1).Columns("Facility_ID").Hidden = True
                    ugByInvoice.DisplayLayout.Bands(1).Columns("DATE_CREATED").Hidden = True

                    ugByInvoice.DisplayLayout.Bands(1).Columns("Debit").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                    ugByInvoice.DisplayLayout.Bands(1).Columns("Credit").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                    ugByInvoice.DisplayLayout.Bands(1).Columns("Credit").CellAppearance.ForeColor = System.Drawing.Color.Red

                    ugByInvoice.DisplayLayout.Bands(1).Columns("Balance").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                End If
            Else
                ugByInvoice.Visible = False
            End If
        End If

        ProcessFeesByInvoiceRadio()

    End Sub


    Private Sub InvoiceRadioChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdBtnAllInvoices.CheckedChanged, rdBtnLateInvoices.CheckedChanged, rdBtnOutstandingBalance.CheckedChanged, rdbtnInvoiceCurrentFY.CheckedChanged
        'ProcessInvoiceGrid()
        'ugInvoices.Refresh()
        If rdBtnAllInvoices.Checked Or rdBtnLateInvoices.Checked Or rdBtnOutstandingBalance.Checked Or rdbtnInvoiceCurrentFY.Checked Then
            LoadugInvoices()
        End If
    End Sub


    Private Sub ReceiptRadioChange(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdAllFY.CheckedChanged, rdCurrentFY.CheckedChanged
        'ProcessReceiptGrid()
        If rdAllFY.Checked Or rdCurrentFY.Checked Then
            LoadugReceipts()
        End If

    End Sub
    Private Sub ProcessInvoiceGrid()
        Dim RowCount As Int64
        Dim SumCharges As Single
        Dim SumPayments As Single
        Dim SumCredits As Single
        Dim SumAdjustments As Single
        Dim SumLegal As Single
        Dim SumBalance As Single
        Dim CurrentFY As Int16


        Dim strTemp As String
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim ucrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try

            RowCount = 0
            SumCharges = 0
            SumPayments = 0
            SumCredits = 0
            SumAdjustments = 0
            SumLegal = 0
            SumBalance = 0
            'Get Fiscal Year from Fees Basis
            CurrentFY = oFeesBasis.GetFiscalYearForFee

            'CurrentFY = DatePart(DateInterval.Year, Now())
            'If DatePart(DateInterval.Month, Now()) > 6 Then
            '    CurrentFY += 1
            'End If

            strTemp = ""

            If ugInvoices.Rows.Count > 0 Then


                If rdBtnLateInvoices.Checked Then
                    ugInvoices.DisplayLayout.Bands(0).Columns("Payments").Hidden = True
                    ugInvoices.DisplayLayout.Bands(0).Columns("Credits").Hidden = True
                    ugInvoices.DisplayLayout.Bands(0).Columns("Adjustments").Hidden = True
                    ugInvoices.DisplayLayout.Bands(0).Columns("Legal").Hidden = True
                    ugInvoices.DisplayLayout.Bands(0).Columns("Balance").Hidden = True
                    ugInvoices.DisplayLayout.Bands(0).Columns("Recommend").Hidden = False
                    ugInvoices.DisplayLayout.Bands(0).Columns("Status").Hidden = False
                    ugInvoices.DisplayLayout.Bands(0).Columns("Excuse").Hidden = False
                Else
                    ugInvoices.DisplayLayout.Bands(0).Columns("Payments").Hidden = False
                    ugInvoices.DisplayLayout.Bands(0).Columns("Credits").Hidden = False
                    ugInvoices.DisplayLayout.Bands(0).Columns("Adjustments").Hidden = False
                    ugInvoices.DisplayLayout.Bands(0).Columns("Legal").Hidden = False
                    ugInvoices.DisplayLayout.Bands(0).Columns("Balance").Hidden = False
                    ugInvoices.DisplayLayout.Bands(0).Columns("Recommend").Hidden = True
                    ugInvoices.DisplayLayout.Bands(0).Columns("Status").Hidden = True
                    ugInvoices.DisplayLayout.Bands(0).Columns("Excuse").Hidden = True
                End If

                For Each ugrow In ugInvoices.Rows
                    For Each ucrow In ugrow.ChildBands(0).Rows

                        If ucrow.Cells("Payments").Value > 0 Then
                            ucrow.Cells("Payments").Appearance.ForeColor = System.Drawing.Color.Red
                        End If

                        If ucrow.Cells("Credits").Value > 0 Then
                            ucrow.Cells("Credits").Appearance.ForeColor = System.Drawing.Color.Red
                        End If

                        If ucrow.Cells("Adjustments").Value > 0 Then
                            ucrow.Cells("Adjustments").Appearance.ForeColor = System.Drawing.Color.Red
                        End If

                        If ucrow.Cells("Legal").Value > 0 Then
                            ucrow.Cells("Legal").Appearance.ForeColor = System.Drawing.Color.Red
                        End If

                    Next

                    If ugrow.Cells("Balance").Value < 0 Then
                        ugrow.Cells("Balance").Appearance.ForeColor = System.Drawing.Color.Red
                    End If

                    If rdBtnAllInvoices.Checked Then
                        ugrow.Hidden = False
                        RowCount += 1
                    ElseIf rdBtnOutstandingBalance.Checked And ugrow.Cells("Balance").Value > 0 Then
                        ugrow.Hidden = False
                        RowCount += 1
                    ElseIf rdBtnLateInvoices.Checked And ugrow.Cells("Fee_Type").Value = "Late Fee" Then 'ugrow.Cells("Balance").Value > 0 And IIf(IsDBNull(ugrow.Cells("Due_Date").Value), DateAdd(DateInterval.Day, 1, Now), ugrow.Cells("Due_Date").Value) < Now.Date Then
                        ugrow.Hidden = False
                        RowCount += 1
                    ElseIf rdbtnInvoiceCurrentFY.Checked And ugrow.Cells("SFY").Value = CurrentFY Then
                        ugrow.Hidden = False
                        RowCount += 1
                    Else
                        ugrow.Hidden = True
                    End If

                    If ugrow.Hidden = False Then

                        SumCharges += ugrow.Cells("Charges").Value

                        If ugrow.Cells("Payments").Value > 0 Then
                            ugrow.Cells("Payments").Appearance.ForeColor = System.Drawing.Color.Red
                        End If
                        SumPayments += ugrow.Cells("Payments").Value

                        If ugrow.Cells("Credits").Value > 0 Then
                            ugrow.Cells("Credits").Appearance.ForeColor = System.Drawing.Color.Red
                        End If
                        SumCredits += ugrow.Cells("Credits").Value

                        If ugrow.Cells("Adjustments").Value > 0 Then
                            ugrow.Cells("Adjustments").Appearance.ForeColor = System.Drawing.Color.Red
                        End If
                        SumAdjustments += ugrow.Cells("Adjustments").Value

                        If ugrow.Cells("Legal").Value > 0 Then
                            ugrow.Cells("Legal").Appearance.ForeColor = System.Drawing.Color.Red
                        End If
                        SumLegal += ugrow.Cells("Legal").Value

                        If ugrow.Cells("Balance").Value < 0 Then
                            ugrow.Cells("Balance").Appearance.ForeColor = System.Drawing.Color.Red
                        End If
                        SumBalance += ugrow.Cells("Balance").Value

                    End If
                Next
            End If
            lblTotalInvoices.Text = RowCount


            lblTotalCharges.Text = FormatNumber(SumCharges, 2, TriState.True, TriState.False, TriState.True)
            lblTotalPayments.Text = FormatNumber(SumPayments, 2, TriState.True, TriState.False, TriState.True)
            If SumPayments > 0 Then
                lblTotalPayments.ForeColor = System.Drawing.Color.Red
            Else
                lblTotalPayments.ForeColor = System.Drawing.Color.Black
            End If

            lblTotalCredits.Text = FormatNumber(SumCredits, 2, TriState.True, TriState.False, TriState.True)
            If SumCredits > 0 Then
                lblTotalCredits.ForeColor = System.Drawing.Color.Red
            Else
                lblTotalCredits.ForeColor = System.Drawing.Color.Black
            End If

            lblTotalLegal.Text = FormatNumber(SumLegal, 2, TriState.True, TriState.False, TriState.True)
            If SumLegal > 0 Then
                lblTotalLegal.ForeColor = System.Drawing.Color.Red
            Else
                lblTotalLegal.ForeColor = System.Drawing.Color.Black
            End If


            lblTotalAdjustments.Text = FormatNumber(SumAdjustments, 2, TriState.True, TriState.False, TriState.True)
            lblTotalBalance.Text = FormatNumber(SumBalance, 2, TriState.True, TriState.False, TriState.True)

            For Each ugrow In ugInvoices.Rows
                If ugrow.Hidden = False Then
                    ugInvoices.ActiveRow = ugrow
                    Exit For
                End If
            Next

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Setup Manual Invoice Grid" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub


    Private Sub ProcessReceiptGrid()

        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim ucrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim nCurrentFY As Int16
        Dim strParentType As String
        Dim strTemp As String
        Dim rp As New Remove_Pencil

        ugReceipts.DrawFilter = rp
        'Get Fiscal Year from Fees Basis
        nCurrentFY = oFeesBasis.GetFiscalYearForFee

        'nCurrentFY = DatePart(DateInterval.Year, Now.Date)
        'If DatePart(DateInterval.Month, Now.Date) > 6 Then
        '    nCurrentFY += 1
        'End If

        If ugReceipts.Rows.Count > 0 Then

            For Each ugrow In ugReceipts.Rows

                strTemp = ugrow.Cells("Amount").Value
                If strTemp < "0.00" Then
                    ugrow.Cells("Amount").Appearance.ForeColor = System.Drawing.Color.Red
                Else
                    ugrow.Cells("Amount").Appearance.ForeColor = System.Drawing.Color.Black
                End If

                If Not IsDBNull(ugrow.Cells("Fiscal_Year").Value) Then
                    If Me.rdCurrentFY.Checked And nCurrentFY <> ugrow.Cells("Fiscal_Year").Value Then
                        ugrow.Hidden = True
                    Else
                        ugrow.Hidden = False
                    End If
                End If

                For Each ucrow In ugrow.ChildBands(0).Rows

                    'If Not IsDBNull(ucrow.Cells("Fiscal_Year").Value) Then
                    '    If Me.rdCurrentFY.Checked And nCurrentFY <> ucrow.Cells("Fiscal_Year").Value Then
                    '        ucrow.Hidden = True
                    '    Else
                    '        ucrow.Hidden = False
                    '    End If
                    'End If

                    Select Case ucrow.Cells("Trans_Type").Value
                        Case "Payment"
                            ucrow.Cells("Amount").Appearance.ForeColor = System.Drawing.Color.Red
                        Case "Reapplied Payment"
                            ucrow.Cells("Amount").Appearance.ForeColor = System.Drawing.Color.Red
                        Case "Misapplied Payment"
                            ucrow.Cells("Amount").Appearance.ForeColor = System.Drawing.Color.Black
                        Case "Returned Check"
                            ucrow.Cells("Amount").Appearance.ForeColor = System.Drawing.Color.Black
                    End Select
                Next

            Next
        End If
        ugReceipts.DrawFilter = rp

    End Sub



    Private Sub FacilityFeesByTransaction_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdBtnByTransAllYears.Click, rdBtnCurrentFiscalYear.Click
        If bolLoading Then Exit Sub
        ProcessFeesByTransactionRadio()
    End Sub
    Private Sub FacilityFeesByInvoice_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdBtnAllYears.Click, rdBtnCurrentOutstanding.Click
        If bolLoading Then Exit Sub
        LoadugByInvoice()

    End Sub

    Private Sub ProcessFeesByTransactionRadio()
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim CFY As Int16
        Dim sBalance As Decimal
        Dim setFocusRow As Infragistics.Win.UltraWinGrid.UltraGridRow = Nothing
        If ugByTransaction.Rows.Count = 0 Then Exit Sub

        'Get Fiscal Year from Fees Basis
        CFY = oFeesBasis.GetFiscalYearForFee

        For Each ugrow In ugByTransaction.Rows


            If rdBtnCurrentFiscalYear.Checked AndAlso ugrow.Cells("FiscalYear").Value <> CFY And ugrow.Cells("TransactionType").Value <> "Prior Balance" Then
                ugrow.Hidden = True
            Else

                ugrow.Hidden = False

                If setFocusRow Is Nothing Then
                    setFocusRow = ugrow
                End If
            End If


        Next

        ugByTransaction.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill


        sBalance = 0


        For Each ugrow In ugByTransaction.Rows
            If IsDBNull(ugrow.Cells("Debit").Value) = False And ugrow.Hidden = False Then
                sBalance += ugrow.Cells("Debit").Value
            End If
            If IsDBNull(ugrow.Cells("Credit").Value) = False And ugrow.Hidden = False Then
                sBalance -= ugrow.Cells("Credit").Value
            End If
            ugrow.Cells("Balance").Value = FormatNumber(sBalance, 2, TriState.True, TriState.False, TriState.True)
            If sBalance < 0 Then
                ugrow.Cells("Balance").Appearance.ForeColor = System.Drawing.Color.Red
            Else
                ugrow.Cells("Balance").Appearance.ForeColor = System.Drawing.Color.Black
            End If
        Next

    End Sub

    Private Sub ProcessFeesByInvoiceRadio()
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim Childrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim CFY As Int16
        Dim sBalance As Decimal

        If ugByInvoice.Rows.Count = 0 Then Exit Sub

        For Each ugrow In ugByInvoice.Rows
            ugrow.Hidden = False
            For Each Childrow In ugrow.ChildBands(0).Rows
                Childrow.Hidden = False
            Next
        Next

        'Get Fiscal Year from Fees Basis
        CFY = oFeesBasis.GetFiscalYearForFee

        ugByInvoice.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill

        For Each ugrow In ugByInvoice.Rows
            sBalance = 0
            For Each Childrow In ugrow.ChildBands(0).Rows


                'If ugrow.Cells("FiscalYear").Value <> Childrow.Cells("FiscalYear").Value And ugrow.Cells("FeeTypeDesc").Value = Childrow.Cells("TransactionType").Value Then
                '    Childrow.Hidden = True
                'Else
                If IsDBNull(Childrow.Cells("Debit").Value) = False And Childrow.Hidden = False Then
                    sBalance += Childrow.Cells("Debit").Value
                End If
                If IsDBNull(Childrow.Cells("Credit").Value) = False And Childrow.Hidden = False Then
                    sBalance -= Childrow.Cells("Credit").Value
                End If
                Childrow.Cells("Balance").Value = FormatNumber(sBalance, 2, TriState.True, TriState.False, TriState.True)
                If sBalance < 0 Then
                    Childrow.Cells("Balance").Appearance.ForeColor = System.Drawing.Color.Red
                Else
                    Childrow.Cells("Balance").Appearance.ForeColor = System.Drawing.Color.Black
                End If
                'End If
            Next
            ugrow.Cells("Balance").Value = FormatNumber(sBalance, 2, TriState.True, TriState.False, TriState.True)
            If sBalance < 0 Then
                ugrow.Cells("Balance").Appearance.ForeColor = System.Drawing.Color.Red
            Else
                ugrow.Cells("Balance").Appearance.ForeColor = System.Drawing.Color.Black
            End If
        Next

        If rdBtnCurrentOutstanding.Checked Then
            For Each ugrow In ugByInvoice.Rows
                If ugrow.Cells("FiscalYear").Value <> CFY Then
                    If ugrow.Cells("Balance").Value <> 0.0 Then
                        'Nothing
                    Else
                        ugrow.Hidden = True
                    End If
                End If
            Next
        End If

    End Sub

    Private Sub PrepareManualInvoicePanel()
        Dim FiscalYear As Int16
        'Get Fiscal Year from Fees Basis
        FiscalYear = oFeesBasis.GetFiscalYearForFee
        'FiscalYear = IIf(DatePart(DateInterval.Month, Now) > 6, DatePart(DateInterval.Year, Now) + 1, DatePart(DateInterval.Year, Now))
        lblMIOwnerIDValue.Text = pOwn.ID
        If pOwn.Persona.Company > String.Empty Then
            lblMIOwnerNameValue.Text = pOwn.Persona.Company
        Else
            lblMIOwnerNameValue.Text = pOwn.Persona.FirstName & " " & pOwn.Persona.LastName
        End If

        lblTotalTanksValue.Text = "0"
        lblInvoiceAmountValue.Text = "0.00"

        oManualInvoice = New MUSTER.BusinessLogic.pFeeInvoice
        cmbFeeType.SelectedIndex = 0
        oManualInvoice.InvoiceType = "I"
        oManualInvoice.FeeType = "1"
        oManualInvoice.RecType = "ADVIC"
        oManualInvoice.OwnerID = pOwn.ID
        'oManualInvoice.RecType = "I"
        oManualInvoice.FiscalYear = FiscalYear

        dtManualInvoice.Clear()

        ugManualInvoices.DisplayLayout.Override.AllowAddNew() = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom

    End Sub


    Private Sub SetupManualInvoiceGrid()
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim dsLocal As DataSet
        Dim drRow As DataRow


        Try

            dtManualInvoice.Columns.Add("FacilityID", GetType(Int64))
            dtManualInvoice.Columns.Add("FiscalYear", GetType(Integer))
            dtManualInvoice.Columns.Add("Quantity", GetType(Integer))
            dtManualInvoice.Columns.Add("UnitPrice", GetType(Single))
            dtManualInvoice.Columns.Add("LineAmount", GetType(Single))
            dtManualInvoice.Columns.Add("Reason", GetType(String))
            dtManualInvoice.Columns.Add("FacilityName", GetType(String))

            ugManualInvoices.DataSource = dtManualInvoice

            ugManualInvoices.DisplayLayout.ValueLists.Add("FacilityIDValue")
            ugManualInvoices.DisplayLayout.ValueLists.Add("FiscalYearValue")

            For Each ugrow In ugFacilityList.Rows
                ugManualInvoices.DisplayLayout.ValueLists("FacilityIDValue").ValueListItems.Add(ugrow.Cells("Facility_ID").Value, ugrow.Cells("Facility_ID").Value)
            Next
            ugManualInvoices.DisplayLayout.Bands(0).Columns("FacilityID").ValueList = ugManualInvoices.DisplayLayout.ValueLists("FacilityIDValue")
            ugManualInvoices.DisplayLayout.Bands(0).Columns("FacilityID").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate
            ugManualInvoices.DisplayLayout.ValueLists("FacilityIDValue").DisplayStyle = Infragistics.Win.ValueListDisplayStyle.DisplayText

            dsLocal = oFeesBasis.GetFeesBasisGrid
            For Each drRow In dsLocal.Tables(0).Rows
                If Not IsDBNull(drRow.Item("Invoice_App_Date")) Then
                    ugManualInvoices.DisplayLayout.ValueLists("FiscalYearValue").ValueListItems.Add(drRow.Item("Fees_Basis_ID"), drRow.Item("Fiscal_Year"))
                End If
            Next
            ugManualInvoices.DisplayLayout.Bands(0).Columns("FiscalYear").ValueList = ugManualInvoices.DisplayLayout.ValueLists("FiscalYearValue")
            ugManualInvoices.DisplayLayout.Bands(0).Columns("FiscalYear").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate
            ugManualInvoices.DisplayLayout.ValueLists("FiscalYearValue").DisplayStyle = Infragistics.Win.ValueListDisplayStyle.DisplayText

            ugManualInvoices.DisplayLayout.Bands(0).Columns("UnitPrice").TabStop = False
            ugManualInvoices.DisplayLayout.Bands(0).Columns("LineAmount").TabStop = False
            ugManualInvoices.DisplayLayout.Bands(0).Columns("FacilityName").TabStop = False

            ugManualInvoices.DisplayLayout.Bands(0).Columns("UnitPrice").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            ugManualInvoices.DisplayLayout.Bands(0).Columns("LineAmount").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            ugManualInvoices.DisplayLayout.Bands(0).Columns("Quantity").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

            ugManualInvoices.DisplayLayout.Bands(0).Columns("FacilityID").Header.Caption = "Facility ID"
            ugManualInvoices.DisplayLayout.Bands(0).Columns("FacilityName").Header.Caption = "Facility Name"
            ugManualInvoices.DisplayLayout.Bands(0).Columns("FiscalYear").Header.Caption = "Fiscal Year"
            ugManualInvoices.DisplayLayout.Bands(0).Columns("LineAmount").Header.Caption = "Line Amount"
            ugManualInvoices.DisplayLayout.Bands(0).Columns("UnitPrice").Header.Caption = "Unit Price"

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Setup Manual Invoice Grid" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try


    End Sub
    Private Sub cmbFeeType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbFeeType.SelectedIndexChanged
        Select Case cmbFeeType.SelectedIndex
            Case 0
                oManualInvoice.FeeType = "1"
            Case 1
                oManualInvoice.FeeType = "3"
                'Case 0
                '    oManualInvoice.FeeType = "FD"
                'Case 1
                '    oManualInvoice.FeeType = "2"
                'Case 2
                '    oManualInvoice.FeeType = "1"
                'Case 3
                '    oManualInvoice.FeeType = "3"
        End Select
    End Sub

    Private Sub btnIssueInvoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnIssueInvoice.Click
        'Verify all line items
        ValidManualInvoice = True
        ProcessManualInvoice()
        If ValidManualInvoice Then
            LoadugInvoices()
            LoadugByInvoice()
            RecalcOwnerTotals()

            MsgBox("Invoice Created")
            pnlManualInvoice.Visible = False
            pnlInvoices.Visible = True
        Else
            MsgBox("Invalid Invoice Information")
        End If
    End Sub


    Private Sub ugManualInvoices_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugManualInvoices.AfterCellUpdate
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim oFac As New MUSTER.BusinessLogic.pFacility
        Dim iTankCount As Int16
        Dim sLineAmount As Single
        Dim sTotalAmount As Single
        Try

            If bolRowSelfUpdating Then
                bolRowSelfUpdating = False
                Exit Sub
            End If
            iTankCount = 0
            sLineAmount = 0
            sTotalAmount = 0


            If IsDBNull(ugManualInvoices.ActiveRow.Cells("FacilityID").Value) = False Then
                oFac.Retrieve(pOwn.OwnerInfo, ugManualInvoices.ActiveRow.Cells("FacilityID").Value, "SELF", "FACILITY")
                bolRowSelfUpdating = True
                ugManualInvoices.ActiveRow.Cells("FacilityName").Value = oFac.Name
            End If


            If IsDBNull(ugManualInvoices.ActiveRow.Cells("FiscalYear").Value) = False Then
                oFeesBasis.Retrieve(ugManualInvoices.ActiveRow.Cells("FiscalYear").Value)
                bolRowSelfUpdating = True
                If cmbFeeType.Text.Trim.ToUpper = "MISCELLANEOUS INVOICE" Then
                    If oFeesBasis.LateType = 913 Then ' %
                        ugManualInvoices.ActiveRow.Cells("UnitPrice").Value = FormatNumber(oFeesBasis.BaseFee + ((oFeesBasis.BaseFee * oFeesBasis.LateFee) / 100), 2, TriState.True, TriState.False, TriState.True)
                    Else ' flat rate
                        ugManualInvoices.ActiveRow.Cells("UnitPrice").Value = FormatNumber(oFeesBasis.BaseFee + oFeesBasis.LateFee, 2, TriState.True, TriState.False, TriState.True)
                    End If
                Else
                    ugManualInvoices.ActiveRow.Cells("UnitPrice").Value = FormatNumber(oFeesBasis.BaseFee, 2, TriState.True, TriState.False, TriState.True)
                End If
            End If

            If IsDBNull(ugManualInvoices.ActiveRow.Cells("Quantity").Value) = False Then
                If IsDBNull(ugManualInvoices.ActiveRow.Cells("UnitPrice").Value) = False Then
                    sLineAmount = ugManualInvoices.ActiveRow.Cells("Quantity").Value * ugManualInvoices.ActiveRow.Cells("UnitPrice").Value
                End If

                bolRowSelfUpdating = True
                ugManualInvoices.ActiveRow.Cells("LineAmount").Value = FormatNumber(sLineAmount, 2, TriState.True, TriState.False, TriState.True)
                For Each ugrow In ugManualInvoices.Rows
                    If IsDBNull(ugrow.Cells("Quantity").Value) = False And IsDBNull(ugrow.Cells("UnitPrice").Value) = False Then
                        sTotalAmount += ugrow.Cells("LineAmount").Value
                        iTankCount += ugrow.Cells("Quantity").Value

                        lblTotalTanksValue.Text = iTankCount
                        lblInvoiceAmountValue.Text = FormatNumber(sTotalAmount, 2, TriState.True, TriState.False, TriState.True)

                        oManualInvoice.Quantity = iTankCount
                        oManualInvoice.InvoiceAmount = sTotalAmount
                    End If
                Next
            End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Update Manual Invoice Grid" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ProcessManualInvoice()
        Dim oInvoiceLineItem As New MUSTER.Info.FeeInvoiceInfo
        Dim iSequence As Int16
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim strFacs As String = String.Empty
        Dim bolAddRegActivityFee As Boolean = False

        If ugManualInvoices.Rows.Count > 0 Then
            For Each ugrow In ugManualInvoices.Rows
                If IsDBNull(ugrow.Cells("LineAmount").Value) = False _
                    And IsDBNull(ugrow.Cells("FiscalYear").Text) = False _
                    And IsDBNull(ugrow.Cells("FacilityID").Value) = False _
                    And IsDBNull(ugrow.Cells("Reason").Value) = False _
                    And IsDBNull(ugrow.Cells("Quantity").Value) = False _
                    And IsDBNull(ugrow.Cells("UnitPrice").Value) = False Then
                    ValidManualInvoice = True

                    iSequence += 1
                    oInvoiceLineItem = New MUSTER.Info.FeeInvoiceInfo
                    oInvoiceLineItem.InvoiceType = "I"
                    oInvoiceLineItem.FeeType = oManualInvoice.FeeType
                    oInvoiceLineItem.RecType = "ADLN"
                    oInvoiceLineItem.ID = iSequence * -1
                    oInvoiceLineItem.OwnerID = pOwn.ID
                    oInvoiceLineItem.FiscalYear = ugrow.Cells("FiscalYear").Text
                    oInvoiceLineItem.FacilityID = ugrow.Cells("FacilityID").Value
                    oInvoiceLineItem.SequenceNumber = iSequence
                    oInvoiceLineItem.Description = ugrow.Cells("Reason").Value
                    oInvoiceLineItem.InvoiceLineAmount = ugrow.Cells("LineAmount").Value
                    oInvoiceLineItem.Quantity = ugrow.Cells("Quantity").Value
                    oInvoiceLineItem.UnitPrice = ugrow.Cells("UnitPrice").Value

                    oManualInvoice.InvoiceLineItems.Add(oInvoiceLineItem)

                    strFacs += ugrow.Cells("FacilityName").Text + " - " + ugrow.Cells("FacilityID").Value.ToString + vbCrLf
                Else
                    ValidManualInvoice = False
                    Exit Sub
                End If
            Next
        Else
            ValidManualInvoice = False
            Exit Sub
        End If
        If oManualInvoice.ID <= 0 Then
            oManualInvoice.CreatedBy = MusterContainer.AppUser.ID
        Else
            oManualInvoice.ModifiedBy = MusterContainer.AppUser.ID
        End If

        If strFacs <> String.Empty Then
            If MsgBox("Do you want to add Registration Activity (Fee) for Facility(s)" + vbCrLf + strFacs, MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                bolAddRegActivityFee = True
            End If
        End If
        oManualInvoice.SaveNewInvoice(CType(UIUtilsGen.ModuleID.Fees, Integer), MusterContainer.AppUser.UserKey, returnVal, bolAddRegActivityFee)
        If Not UIUtilsGen.HasRights(returnVal) Then
            Exit Sub
        End If
    End Sub


    Private Sub RecalcOwnerTotals()
        Dim oRecalcInvoice As New MUSTER.BusinessLogic.pFeeInvoice
        txtBalance.Text = FormatNumber(oRecalcInvoice.GetCurrentBalance(lblOwnerIDValue.Text), 2, TriState.True, TriState.False, TriState.True)
        txtOverage.Text = FormatNumber(oRecalcInvoice.GetOverpaymentBucket(lblOwnerIDValue.Text), 2, TriState.True, TriState.False, TriState.True)
        txtFinal.Text = FormatNumber(txtBalance.Text - txtOverage.Text, 2, TriState.True, TriState.False, TriState.True)
        '' e.g.  bal = -10; overage = 10; final should be = -20; (final = bal - overage)
        ''       bal = 0; overage = 10; final should be = -10;
        'If txtBalance.Text <= "0.00" Then
        '    txtFinal.Text = FormatNumber(txtBalance.Text + txtOverage.Text, 2, TriState.True, TriState.False, TriState.True)
        'Else
        '    txtFinal.Text = FormatNumber(txtBalance.Text - txtOverage.Text, 2, TriState.True, TriState.False, TriState.True)
        'End If
    End Sub

    Private Sub btnFacGenerateInvoice_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFacGenerateInvoice.Click
        Try
            tbCntrlFees.SelectedTab = tbPageOwnerDetail
            Me.Text = "Fees - Owner Detail (" & txtOwnerName.Text & ")"
            tbCtrlOwnerTransactions.SelectedTab = tbPageInvoices

            PrepareManualInvoicePanel()
            pnlManualInvoice.Visible = True
            pnlInvoices.Visible = False
            RecalcOwnerTotals()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnFacGenerateCreditMemo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFacGenerateCreditMemo.Click
        Dim oInv As New MUSTER.BusinessLogic.pFeeInvoice

        Try

            If GetVisibleRowCount(ugByInvoice) = 0 Then
                MsgBox("No Invoices Available For Credit Memo", MsgBoxStyle.OKOnly, "Invalid Invoice")
                Exit Sub
            End If

            If ugByInvoice.DataSource Is Nothing Then
                Exit Sub
            Else
                If ugByInvoice.Rows Is Nothing Then
                    Exit Sub
                ElseIf ugByInvoice.ActiveRow Is Nothing Then
                    MsgBox("Please select a row")
                    Exit Sub
                ElseIf ugByInvoice.ActiveRow.Band.Index = 1 Then
                    ugByInvoice.ActiveRow = ugByInvoice.ActiveRow.ParentRow
                End If
            End If

            oInv.Retrieve(ugByInvoice.ActiveRow.Cells("INV_ID").Text)
            If IsDBNull(oInv.WarrantNumber) Then
                MsgBox("Invoice must have an Invoice Number from BP2K before a Credit Memo can be issued.", MsgBoxStyle.OKOnly, "Invalid Invoice Number")
                Exit Sub
            End If
            If oInv.WarrantNumber = "" Then
                MsgBox("Invoice must have an Invoice Number from BP2K before a Credit Memo can be issued.", MsgBoxStyle.OKOnly, "Invalid Invoice Number")
                Exit Sub
            End If

            If IsNothing(frmGenCreditMemo) Then
                frmGenCreditMemo = New GenerateCreditMemo
                AddHandler frmGenCreditMemo.Closing, AddressOf frmClosing
                AddHandler frmGenCreditMemo.Closed, AddressOf frmClosed
            Else
                frmGenCreditMemo = New GenerateCreditMemo
            End If
            frmGenCreditMemo.InvoiceID = ugByInvoice.ActiveRow.Cells("INV_ID").Text
            frmGenCreditMemo.OwnerID = pOwn.ID
            frmGenCreditMemo.FacilityID = lblFacilityIDValue.Text
            frmGenCreditMemo.bolUpdateCM = False
            Me.Tag = "0"
            frmGenCreditMemo.CallingForm = Me
            frmGenCreditMemo.ShowDialog()

            If Me.Tag = "1" Then
                LoadugInvoices()
                LoadugByInvoice()
                RecalcOwnerTotals()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnFacEditCreditMemo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFacEditCreditMemo.Click
        Try
            If ugByInvoice.Rows Is Nothing Then
                'btnFacEditCreditMemo.Enabled = False
                Exit Sub
            ElseIf ugByInvoice.ActiveRow Is Nothing Then
                'btnFacEditCreditMemo.Enabled = False
                MsgBox("Please select a Credit Memo Transaction", MsgBoxStyle.OKOnly, "Invalid Invoice")
                Exit Sub
            ElseIf ugByInvoice.ActiveRow.Band.Index = 0 Then
                'btnFacEditCreditMemo.Enabled = False
                MsgBox("Please select a Credit Memo Transaction", MsgBoxStyle.OKOnly, "Invalid Invoice")
                Exit Sub
            ElseIf ugByInvoice.ActiveRow.Band.Index = 1 Then
                If ugByInvoice.ActiveRow.Cells("TransactionType").Text.ToUpper = "CREDIT MEMO" Then
                    If ugByInvoice.ActiveRow.Cells("DATE_CREATED").Value Is DBNull.Value Then
                        'btnFacEditCreditMemo.Enabled = False
                        MsgBox("Invalid Create Date, Cannot Delete Credit Memo Transaction", MsgBoxStyle.OKOnly, "Invalid Create Date")
                        Exit Sub
                    Else
                        If Date.Compare(ugByInvoice.ActiveRow.Cells("DATE_CREATED").Value, Today.Date) < 0 Then
                            'btnFacEditCreditMemo.Enabled = False
                            MsgBox("Cannot Delete Credit Memo Transaction created prior to today", MsgBoxStyle.OKOnly, "Invalid Invoice")
                            Exit Sub
                        End If
                    End If
                Else
                    MsgBox("Please select a Credit Memo Transaction", MsgBoxStyle.OKOnly, "Invalid Invoice")
                    Exit Sub
                End If
            End If

            If IsNothing(frmGenCreditMemo) Then
                frmGenCreditMemo = New GenerateCreditMemo
                AddHandler frmGenCreditMemo.Closing, AddressOf frmClosing
                AddHandler frmGenCreditMemo.Closed, AddressOf frmClosed
            Else
                frmGenCreditMemo = New GenerateCreditMemo
            End If
            frmGenCreditMemo.InvoiceID = ugByInvoice.ActiveRow.ParentRow.Cells("INV_ID").Text
            frmGenCreditMemo.OwnerID = pOwn.ID
            frmGenCreditMemo.FacilityID = lblFacilityIDValue.Text
            frmGenCreditMemo.bolUpdateCM = True
            Me.Tag = "0"
            frmGenCreditMemo.CallingForm = Me
            frmGenCreditMemo.ShowDialog()

            If Me.Tag = "1" Then
                LoadugInvoices()
                LoadugByInvoice()
                RecalcOwnerTotals()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    'Private Sub ugByInvoice_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugByInvoice.AfterRowActivate
    Private Sub lblOwnerTotalsfor_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblOwnerTotalsfor.Click

    End Sub    '    btnFacEditCreditMemo.Enabled = False
    '    Try
    '        If Not ugByInvoice.ActiveRow Is Nothing Then
    '            If ugByInvoice.ActiveRow.Band.Index = 1 Then
    '                If ugByInvoice.ActiveRow.Cells("TransactionType").Text.ToUpper.IndexOf("CREDIT MEMO") > -1 Then
    '                    If Not ugByInvoice.ActiveRow.Cells("DATE_CREATED").Value Is DBNull.Value Then
    '                        If ugByInvoice.ActiveRow.Cells("DATE_CREATED").Value = Today.Date Then
    '                            btnFacEditCreditMemo.Enabled = True
    '                        End If
    '                    End If
    '                End If
    '            End If
    '        End If
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub

    Private Sub ugManualInvoices_BeforeCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles ugManualInvoices.BeforeCellUpdate
        If bolRowSelfUpdating = True Then Exit Sub

        If ugManualInvoices.ActiveRow.Cells("FacilityName").IsActiveCell Then
            MsgBox("Facility Name is not an enterable field.  Value will default based upon FacilityID")
        End If
        If ugManualInvoices.ActiveRow.Cells("UnitPrice").IsActiveCell Then
            MsgBox("Unit Price is not an enterable field.  Value will default based upon Fiscal Year")
        End If
        If ugManualInvoices.ActiveRow.Cells("LineAmount").IsActiveCell Then
            MsgBox("Line Amount is not an enterable field.  Value will default based upon Quantity and Unit Price")
        End If
    End Sub

#Region "Envelopes and Labels"
    Private Sub btnFeesOwnerEnvelopes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFeesOwnerEnvelopes.Click
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
                Dim dsContactsLocal As DataSet = pConStruct.GetFilteredContacts(pOwn.ID, 612)

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


                UIUtilsGen.CreateEnvelopes(Me.txtOwnerName.Text, arrAddress, "FEES", pOwn.ID, strContactName)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnFeesOwnerLabels_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFeesOwnerLabels.Click
        Dim arrAddress(4) As String
        Try
            arrAddress(0) = pOwn.Addresses.AddressLine1
            arrAddress(1) = pOwn.Addresses.AddressLine2
            arrAddress(2) = pOwn.Addresses.City
            arrAddress(3) = pOwn.Addresses.State
            arrAddress(4) = pOwn.Addresses.Zip
            'strAddress = pOwn.Addresses.AddressLine1 + "," + pOwn.Addresses.AddressLine2 + "," + pOwn.Addresses.City + "," + pOwn.Addresses.State + "," + pOwn.Addresses.Zip
            If pOwn.Addresses.AddressId > 0 And pOwn.ID > 0 Then
                UIUtilsGen.CreateLabels(Me.txtOwnerName.Text, arrAddress, "FEES", pOwn.ID)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnFeesFacEnvelopes_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFeesFacEnvelopes.Click
        Dim arrAddress(4) As String
        Try
            arrAddress(0) = pOwn.Facilities.FacilityAddresses.AddressLine1
            arrAddress(1) = pOwn.Facilities.FacilityAddresses.AddressLine2
            arrAddress(2) = pOwn.Facilities.FacilityAddresses.City
            arrAddress(3) = pOwn.Facilities.FacilityAddresses.State
            arrAddress(4) = pOwn.Facilities.FacilityAddresses.Zip
            'strAddress = pOwn.Facilities.FacilityAddresses.AddressLine1 + "," + pOwn.Facilities.FacilityAddresses.AddressLine2 + "," + pOwn.Facilities.FacilityAddresses.City + "," + pOwn.Facilities.FacilityAddresses.State + "," + pOwn.Facilities.FacilityAddresses.Zip
            If pOwn.Facilities.FacilityAddresses.AddressId > 0 And pOwn.ID > 0 Then
                UIUtilsGen.CreateEnvelopes(Me.txtFacilityName.Text, arrAddress, "FEES", pOwn.Facilities.ID)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnFeesFacLabels_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFeesFacLabels.Click
        Dim arrAddress(4) As String
        Try
            arrAddress(0) = pOwn.Facilities.FacilityAddresses.AddressLine1
            arrAddress(1) = pOwn.Facilities.FacilityAddresses.AddressLine2
            arrAddress(2) = pOwn.Facilities.FacilityAddresses.City
            arrAddress(3) = pOwn.Facilities.FacilityAddresses.State
            arrAddress(4) = pOwn.Facilities.FacilityAddresses.Zip
            'strAddress = pOwn.Facilities.FacilityAddresses.AddressLine1 + "," + pOwn.Facilities.FacilityAddresses.AddressLine2 + "," + pOwn.Facilities.FacilityAddresses.City + "," + pOwn.Facilities.FacilityAddresses.State + "," + pOwn.Facilities.FacilityAddresses.Zip
            If pOwn.Facilities.FacilityAddresses.AddressId > 0 And pOwn.ID > 0 Then
                UIUtilsGen.CreateLabels(Me.txtFacilityName.Text, arrAddress, "FEES", pOwn.Facilities.ID)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region



End Class
