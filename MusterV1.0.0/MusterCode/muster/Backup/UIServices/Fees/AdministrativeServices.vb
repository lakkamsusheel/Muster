Public Class AdministrativeServices
    Inherits System.Windows.Forms.Form
    Const FEES_BASIS = 0
    Const INVOICE_REVIEW = 2
    Const LATE_FEE_CERT = 1
    Const WAIVE_LATE_FEE = 3
#Region "User Defined Variables"
    Private frmLateFeeWaiver As LateFeeWaiverRequest
    Private frmFiscalYearBasis As FiscalYearFeeBasis
    Private oFeesBasis As New MUSTER.BusinessLogic.pFeeBasis
    Private oLateFees As New MUSTER.BusinessLogic.pFeeLateFee
    Private bolLoading As Boolean = True
    Private isRedTag As Boolean = False
    Dim returnVal As String = String.Empty
#End Region
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
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
    Friend WithEvents pnlAdminServicesDetails As System.Windows.Forms.Panel
    Friend WithEvents tbCtrlAdminServices As System.Windows.Forms.TabControl
    Friend WithEvents tbPageFYBasis As System.Windows.Forms.TabPage
    Friend WithEvents tbPageInvoiceReview As System.Windows.Forms.TabPage
    Friend WithEvents tbPageLateFeeCertifications As System.Windows.Forms.TabPage
    Friend WithEvents tbPageWaiveLateFees As System.Windows.Forms.TabPage
    Friend WithEvents pnlFYBasisDisplay As System.Windows.Forms.Panel
    Friend WithEvents lblFeeBasisHistory As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents pnlFYBasisBottom As System.Windows.Forms.Panel
    Friend WithEvents btnDeleteFiscalYearBasis As System.Windows.Forms.Button
    Friend WithEvents btnModifyFiscalYearBasis As System.Windows.Forms.Button
    Friend WithEvents btnAddFiscalYearBasis As System.Windows.Forms.Button
    Friend WithEvents pnlFYBasisDetails As System.Windows.Forms.Panel
    Friend WithEvents ugFYBasis As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnInvoiceCaption As System.Windows.Forms.Label
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents ugInvoiceReview As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnRegenerate As System.Windows.Forms.Button
    Friend WithEvents btnApprove As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents lblTotal As System.Windows.Forms.Label
    Friend WithEvents lblTotalNoOfFacs As System.Windows.Forms.Label
    Friend WithEvents lblTotalNoOfTanks As System.Windows.Forms.Label
    Friend WithEvents lblCharges As System.Windows.Forms.Label
    Friend WithEvents pnlLateFeeTop As System.Windows.Forms.Panel
    Friend WithEvents pnlLateFeeBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlLateFeeDetails As System.Windows.Forms.Panel
    Friend WithEvents lblCertifiedNo As System.Windows.Forms.Label
    Friend WithEvents lblApplyto As System.Windows.Forms.Label
    Friend WithEvents rdBtnAll As System.Windows.Forms.RadioButton
    Friend WithEvents rdBtnSelection As System.Windows.Forms.RadioButton
    Friend WithEvents btnApplyTemplate As System.Windows.Forms.Button
    Friend WithEvents btnLateFeeSave As System.Windows.Forms.Button
    Friend WithEvents btnLateFeeProcess As System.Windows.Forms.Button
    Friend WithEvents btnLateFeeCancel As System.Windows.Forms.Button
    Friend WithEvents ugLateFee As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlWaiveLateFeesTop As System.Windows.Forms.Panel
    Friend WithEvents pnlWaiveLateFeeBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlWaiveLateFeeDetails As System.Windows.Forms.Panel
    Friend WithEvents drBtnSelectAll As System.Windows.Forms.RadioButton
    Friend WithEvents ugWaiveLateFee As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnAcceptDecisions As System.Windows.Forms.Button
    Friend WithEvents btnModifyWaiver As System.Windows.Forms.Button
    Friend WithEvents btnDeleteWaiver As System.Windows.Forms.Button
    Friend WithEvents btnWaiveLateFeeCancel As System.Windows.Forms.Button
    Friend WithEvents txtNo1 As System.Windows.Forms.TextBox
    Friend WithEvents txtNo2 As System.Windows.Forms.TextBox
    Friend WithEvents txtNo3 As System.Windows.Forms.TextBox
    Friend WithEvents txtNo4 As System.Windows.Forms.TextBox
    Friend WithEvents txtNo5 As System.Windows.Forms.TextBox
    Friend WithEvents chkHideProcessed As System.Windows.Forms.CheckBox
    Friend WithEvents chkHideProcessed_Waivers As System.Windows.Forms.CheckBox
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents rdBtnRedTag As System.Windows.Forms.RadioButton
    Friend WithEvents rdBtnLateFee As System.Windows.Forms.RadioButton
    Friend WithEvents lblLetterType As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlAdminServicesDetails = New System.Windows.Forms.Panel
        Me.tbCtrlAdminServices = New System.Windows.Forms.TabControl
        Me.tbPageFYBasis = New System.Windows.Forms.TabPage
        Me.pnlFYBasisDetails = New System.Windows.Forms.Panel
        Me.ugFYBasis = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlFYBasisBottom = New System.Windows.Forms.Panel
        Me.btnDeleteFiscalYearBasis = New System.Windows.Forms.Button
        Me.btnModifyFiscalYearBasis = New System.Windows.Forms.Button
        Me.btnAddFiscalYearBasis = New System.Windows.Forms.Button
        Me.pnlFYBasisDisplay = New System.Windows.Forms.Panel
        Me.lblFeeBasisHistory = New System.Windows.Forms.Label
        Me.tbPageLateFeeCertifications = New System.Windows.Forms.TabPage
        Me.pnlLateFeeDetails = New System.Windows.Forms.Panel
        Me.ugLateFee = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlLateFeeBottom = New System.Windows.Forms.Panel
        Me.Button1 = New System.Windows.Forms.Button
        Me.btnLateFeeCancel = New System.Windows.Forms.Button
        Me.btnLateFeeProcess = New System.Windows.Forms.Button
        Me.btnLateFeeSave = New System.Windows.Forms.Button
        Me.btnApplyTemplate = New System.Windows.Forms.Button
        Me.pnlLateFeeTop = New System.Windows.Forms.Panel
        Me.Panel5 = New System.Windows.Forms.Panel
        Me.rdBtnRedTag = New System.Windows.Forms.RadioButton
        Me.rdBtnLateFee = New System.Windows.Forms.RadioButton
        Me.lblLetterType = New System.Windows.Forms.Label
        Me.chkHideProcessed = New System.Windows.Forms.CheckBox
        Me.txtNo5 = New System.Windows.Forms.TextBox
        Me.txtNo4 = New System.Windows.Forms.TextBox
        Me.txtNo3 = New System.Windows.Forms.TextBox
        Me.txtNo2 = New System.Windows.Forms.TextBox
        Me.txtNo1 = New System.Windows.Forms.TextBox
        Me.rdBtnSelection = New System.Windows.Forms.RadioButton
        Me.rdBtnAll = New System.Windows.Forms.RadioButton
        Me.lblApplyto = New System.Windows.Forms.Label
        Me.lblCertifiedNo = New System.Windows.Forms.Label
        Me.tbPageInvoiceReview = New System.Windows.Forms.TabPage
        Me.Panel4 = New System.Windows.Forms.Panel
        Me.ugInvoiceReview = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.lblCharges = New System.Windows.Forms.Label
        Me.lblTotalNoOfTanks = New System.Windows.Forms.Label
        Me.lblTotalNoOfFacs = New System.Windows.Forms.Label
        Me.lblTotal = New System.Windows.Forms.Label
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnApprove = New System.Windows.Forms.Button
        Me.btnRegenerate = New System.Windows.Forms.Button
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.btnInvoiceCaption = New System.Windows.Forms.Label
        Me.tbPageWaiveLateFees = New System.Windows.Forms.TabPage
        Me.pnlWaiveLateFeeDetails = New System.Windows.Forms.Panel
        Me.ugWaiveLateFee = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlWaiveLateFeeBottom = New System.Windows.Forms.Panel
        Me.btnWaiveLateFeeCancel = New System.Windows.Forms.Button
        Me.btnDeleteWaiver = New System.Windows.Forms.Button
        Me.btnModifyWaiver = New System.Windows.Forms.Button
        Me.btnAcceptDecisions = New System.Windows.Forms.Button
        Me.pnlWaiveLateFeesTop = New System.Windows.Forms.Panel
        Me.chkHideProcessed_Waivers = New System.Windows.Forms.CheckBox
        Me.drBtnSelectAll = New System.Windows.Forms.RadioButton
        Me.pnlAdminServicesDetails.SuspendLayout()
        Me.tbCtrlAdminServices.SuspendLayout()
        Me.tbPageFYBasis.SuspendLayout()
        Me.pnlFYBasisDetails.SuspendLayout()
        CType(Me.ugFYBasis, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlFYBasisBottom.SuspendLayout()
        Me.pnlFYBasisDisplay.SuspendLayout()
        Me.tbPageLateFeeCertifications.SuspendLayout()
        Me.pnlLateFeeDetails.SuspendLayout()
        CType(Me.ugLateFee, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlLateFeeBottom.SuspendLayout()
        Me.pnlLateFeeTop.SuspendLayout()
        Me.Panel5.SuspendLayout()
        Me.tbPageInvoiceReview.SuspendLayout()
        Me.Panel4.SuspendLayout()
        CType(Me.ugInvoiceReview, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel3.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.tbPageWaiveLateFees.SuspendLayout()
        Me.pnlWaiveLateFeeDetails.SuspendLayout()
        CType(Me.ugWaiveLateFee, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlWaiveLateFeeBottom.SuspendLayout()
        Me.pnlWaiveLateFeesTop.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlAdminServicesDetails
        '
        Me.pnlAdminServicesDetails.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlAdminServicesDetails.Controls.Add(Me.tbCtrlAdminServices)
        Me.pnlAdminServicesDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlAdminServicesDetails.Location = New System.Drawing.Point(0, 0)
        Me.pnlAdminServicesDetails.Name = "pnlAdminServicesDetails"
        Me.pnlAdminServicesDetails.Size = New System.Drawing.Size(928, 694)
        Me.pnlAdminServicesDetails.TabIndex = 1
        '
        'tbCtrlAdminServices
        '
        Me.tbCtrlAdminServices.Controls.Add(Me.tbPageFYBasis)
        Me.tbCtrlAdminServices.Controls.Add(Me.tbPageLateFeeCertifications)
        Me.tbCtrlAdminServices.Controls.Add(Me.tbPageInvoiceReview)
        Me.tbCtrlAdminServices.Controls.Add(Me.tbPageWaiveLateFees)
        Me.tbCtrlAdminServices.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbCtrlAdminServices.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbCtrlAdminServices.Location = New System.Drawing.Point(0, 0)
        Me.tbCtrlAdminServices.Multiline = True
        Me.tbCtrlAdminServices.Name = "tbCtrlAdminServices"
        Me.tbCtrlAdminServices.SelectedIndex = 0
        Me.tbCtrlAdminServices.ShowToolTips = True
        Me.tbCtrlAdminServices.Size = New System.Drawing.Size(924, 690)
        Me.tbCtrlAdminServices.TabIndex = 0
        '
        'tbPageFYBasis
        '
        Me.tbPageFYBasis.Controls.Add(Me.pnlFYBasisDetails)
        Me.tbPageFYBasis.Controls.Add(Me.pnlFYBasisBottom)
        Me.tbPageFYBasis.Controls.Add(Me.pnlFYBasisDisplay)
        Me.tbPageFYBasis.Location = New System.Drawing.Point(4, 24)
        Me.tbPageFYBasis.Name = "tbPageFYBasis"
        Me.tbPageFYBasis.Size = New System.Drawing.Size(916, 662)
        Me.tbPageFYBasis.TabIndex = 0
        Me.tbPageFYBasis.Text = "F/Y Basis"
        '
        'pnlFYBasisDetails
        '
        Me.pnlFYBasisDetails.Controls.Add(Me.ugFYBasis)
        Me.pnlFYBasisDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlFYBasisDetails.Location = New System.Drawing.Point(0, 24)
        Me.pnlFYBasisDetails.Name = "pnlFYBasisDetails"
        Me.pnlFYBasisDetails.Size = New System.Drawing.Size(916, 598)
        Me.pnlFYBasisDetails.TabIndex = 4
        '
        'ugFYBasis
        '
        Me.ugFYBasis.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugFYBasis.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugFYBasis.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        Me.ugFYBasis.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        Me.ugFYBasis.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle
        Me.ugFYBasis.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugFYBasis.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugFYBasis.Location = New System.Drawing.Point(0, 0)
        Me.ugFYBasis.Name = "ugFYBasis"
        Me.ugFYBasis.Size = New System.Drawing.Size(916, 598)
        Me.ugFYBasis.TabIndex = 0
        '
        'pnlFYBasisBottom
        '
        Me.pnlFYBasisBottom.Controls.Add(Me.btnDeleteFiscalYearBasis)
        Me.pnlFYBasisBottom.Controls.Add(Me.btnModifyFiscalYearBasis)
        Me.pnlFYBasisBottom.Controls.Add(Me.btnAddFiscalYearBasis)
        Me.pnlFYBasisBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlFYBasisBottom.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlFYBasisBottom.Location = New System.Drawing.Point(0, 622)
        Me.pnlFYBasisBottom.Name = "pnlFYBasisBottom"
        Me.pnlFYBasisBottom.Size = New System.Drawing.Size(916, 40)
        Me.pnlFYBasisBottom.TabIndex = 1
        '
        'btnDeleteFiscalYearBasis
        '
        Me.btnDeleteFiscalYearBasis.Location = New System.Drawing.Point(440, 8)
        Me.btnDeleteFiscalYearBasis.Name = "btnDeleteFiscalYearBasis"
        Me.btnDeleteFiscalYearBasis.Size = New System.Drawing.Size(152, 23)
        Me.btnDeleteFiscalYearBasis.TabIndex = 4
        Me.btnDeleteFiscalYearBasis.Text = "Delete Fiscal Year Basis"
        '
        'btnModifyFiscalYearBasis
        '
        Me.btnModifyFiscalYearBasis.Location = New System.Drawing.Point(280, 8)
        Me.btnModifyFiscalYearBasis.Name = "btnModifyFiscalYearBasis"
        Me.btnModifyFiscalYearBasis.Size = New System.Drawing.Size(152, 23)
        Me.btnModifyFiscalYearBasis.TabIndex = 3
        Me.btnModifyFiscalYearBasis.Text = "Modify Fiscal Year Basis"
        '
        'btnAddFiscalYearBasis
        '
        Me.btnAddFiscalYearBasis.Location = New System.Drawing.Point(128, 8)
        Me.btnAddFiscalYearBasis.Name = "btnAddFiscalYearBasis"
        Me.btnAddFiscalYearBasis.Size = New System.Drawing.Size(136, 23)
        Me.btnAddFiscalYearBasis.TabIndex = 2
        Me.btnAddFiscalYearBasis.Text = "Add Fiscal Year Basis"
        '
        'pnlFYBasisDisplay
        '
        Me.pnlFYBasisDisplay.Controls.Add(Me.lblFeeBasisHistory)
        Me.pnlFYBasisDisplay.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlFYBasisDisplay.Location = New System.Drawing.Point(0, 0)
        Me.pnlFYBasisDisplay.Name = "pnlFYBasisDisplay"
        Me.pnlFYBasisDisplay.Size = New System.Drawing.Size(916, 24)
        Me.pnlFYBasisDisplay.TabIndex = 1
        '
        'lblFeeBasisHistory
        '
        Me.lblFeeBasisHistory.Location = New System.Drawing.Point(8, 4)
        Me.lblFeeBasisHistory.Name = "lblFeeBasisHistory"
        Me.lblFeeBasisHistory.Size = New System.Drawing.Size(112, 17)
        Me.lblFeeBasisHistory.TabIndex = 1
        Me.lblFeeBasisHistory.Text = "Fee Basis History"
        '
        'tbPageLateFeeCertifications
        '
        Me.tbPageLateFeeCertifications.Controls.Add(Me.pnlLateFeeDetails)
        Me.tbPageLateFeeCertifications.Controls.Add(Me.pnlLateFeeBottom)
        Me.tbPageLateFeeCertifications.Controls.Add(Me.pnlLateFeeTop)
        Me.tbPageLateFeeCertifications.Location = New System.Drawing.Point(4, 24)
        Me.tbPageLateFeeCertifications.Name = "tbPageLateFeeCertifications"
        Me.tbPageLateFeeCertifications.Size = New System.Drawing.Size(916, 662)
        Me.tbPageLateFeeCertifications.TabIndex = 2
        Me.tbPageLateFeeCertifications.Text = "Late Fee & Red Tag Certifications"
        '
        'pnlLateFeeDetails
        '
        Me.pnlLateFeeDetails.Controls.Add(Me.ugLateFee)
        Me.pnlLateFeeDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlLateFeeDetails.Location = New System.Drawing.Point(0, 64)
        Me.pnlLateFeeDetails.Name = "pnlLateFeeDetails"
        Me.pnlLateFeeDetails.Size = New System.Drawing.Size(916, 558)
        Me.pnlLateFeeDetails.TabIndex = 2
        '
        'ugLateFee
        '
        Me.ugLateFee.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugLateFee.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle
        Me.ugLateFee.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugLateFee.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugLateFee.Location = New System.Drawing.Point(0, 0)
        Me.ugLateFee.Name = "ugLateFee"
        Me.ugLateFee.Size = New System.Drawing.Size(916, 558)
        Me.ugLateFee.TabIndex = 3
        '
        'pnlLateFeeBottom
        '
        Me.pnlLateFeeBottom.Controls.Add(Me.Button1)
        Me.pnlLateFeeBottom.Controls.Add(Me.btnLateFeeCancel)
        Me.pnlLateFeeBottom.Controls.Add(Me.btnLateFeeProcess)
        Me.pnlLateFeeBottom.Controls.Add(Me.btnLateFeeSave)
        Me.pnlLateFeeBottom.Controls.Add(Me.btnApplyTemplate)
        Me.pnlLateFeeBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlLateFeeBottom.Location = New System.Drawing.Point(0, 622)
        Me.pnlLateFeeBottom.Name = "pnlLateFeeBottom"
        Me.pnlLateFeeBottom.Size = New System.Drawing.Size(916, 40)
        Me.pnlLateFeeBottom.TabIndex = 4
        '
        'Button1
        '
        Me.Button1.Enabled = False
        Me.Button1.Location = New System.Drawing.Point(776, 8)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(128, 23)
        Me.Button1.TabIndex = 9
        Me.Button1.Text = "UnProcess RedTag"
        '
        'btnLateFeeCancel
        '
        Me.btnLateFeeCancel.Location = New System.Drawing.Point(520, 8)
        Me.btnLateFeeCancel.Name = "btnLateFeeCancel"
        Me.btnLateFeeCancel.TabIndex = 8
        Me.btnLateFeeCancel.Text = "Cancel"
        '
        'btnLateFeeProcess
        '
        Me.btnLateFeeProcess.Location = New System.Drawing.Point(440, 8)
        Me.btnLateFeeProcess.Name = "btnLateFeeProcess"
        Me.btnLateFeeProcess.TabIndex = 7
        Me.btnLateFeeProcess.Text = "Process"
        '
        'btnLateFeeSave
        '
        Me.btnLateFeeSave.Location = New System.Drawing.Point(360, 8)
        Me.btnLateFeeSave.Name = "btnLateFeeSave"
        Me.btnLateFeeSave.TabIndex = 6
        Me.btnLateFeeSave.Text = "Save"
        '
        'btnApplyTemplate
        '
        Me.btnApplyTemplate.Location = New System.Drawing.Point(248, 8)
        Me.btnApplyTemplate.Name = "btnApplyTemplate"
        Me.btnApplyTemplate.Size = New System.Drawing.Size(104, 23)
        Me.btnApplyTemplate.TabIndex = 5
        Me.btnApplyTemplate.Text = "Apply Template"
        '
        'pnlLateFeeTop
        '
        Me.pnlLateFeeTop.Controls.Add(Me.Panel5)
        Me.pnlLateFeeTop.Controls.Add(Me.chkHideProcessed)
        Me.pnlLateFeeTop.Controls.Add(Me.txtNo5)
        Me.pnlLateFeeTop.Controls.Add(Me.txtNo4)
        Me.pnlLateFeeTop.Controls.Add(Me.txtNo3)
        Me.pnlLateFeeTop.Controls.Add(Me.txtNo2)
        Me.pnlLateFeeTop.Controls.Add(Me.txtNo1)
        Me.pnlLateFeeTop.Controls.Add(Me.rdBtnSelection)
        Me.pnlLateFeeTop.Controls.Add(Me.rdBtnAll)
        Me.pnlLateFeeTop.Controls.Add(Me.lblApplyto)
        Me.pnlLateFeeTop.Controls.Add(Me.lblCertifiedNo)
        Me.pnlLateFeeTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlLateFeeTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlLateFeeTop.Name = "pnlLateFeeTop"
        Me.pnlLateFeeTop.Size = New System.Drawing.Size(916, 64)
        Me.pnlLateFeeTop.TabIndex = 0
        '
        'Panel5
        '
        Me.Panel5.Controls.Add(Me.rdBtnRedTag)
        Me.Panel5.Controls.Add(Me.rdBtnLateFee)
        Me.Panel5.Controls.Add(Me.lblLetterType)
        Me.Panel5.Location = New System.Drawing.Point(8, 8)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(336, 32)
        Me.Panel5.TabIndex = 18
        '
        'rdBtnRedTag
        '
        Me.rdBtnRedTag.Location = New System.Drawing.Point(168, 0)
        Me.rdBtnRedTag.Name = "rdBtnRedTag"
        Me.rdBtnRedTag.Size = New System.Drawing.Size(88, 24)
        Me.rdBtnRedTag.TabIndex = 19
        Me.rdBtnRedTag.Text = "Red Tag"
        '
        'rdBtnLateFee
        '
        Me.rdBtnLateFee.Checked = True
        Me.rdBtnLateFee.Location = New System.Drawing.Point(88, 0)
        Me.rdBtnLateFee.Name = "rdBtnLateFee"
        Me.rdBtnLateFee.Size = New System.Drawing.Size(72, 24)
        Me.rdBtnLateFee.TabIndex = 18
        Me.rdBtnLateFee.TabStop = True
        Me.rdBtnLateFee.Text = "Late Fee"
        '
        'lblLetterType
        '
        Me.lblLetterType.Location = New System.Drawing.Point(16, 0)
        Me.lblLetterType.Name = "lblLetterType"
        Me.lblLetterType.Size = New System.Drawing.Size(96, 23)
        Me.lblLetterType.TabIndex = 20
        Me.lblLetterType.Text = "Letter Type: "
        Me.lblLetterType.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'chkHideProcessed
        '
        Me.chkHideProcessed.Checked = True
        Me.chkHideProcessed.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkHideProcessed.Location = New System.Drawing.Point(752, 32)
        Me.chkHideProcessed.Name = "chkHideProcessed"
        Me.chkHideProcessed.Size = New System.Drawing.Size(136, 24)
        Me.chkHideProcessed.TabIndex = 14
        Me.chkHideProcessed.Text = "Show Unprocessed"
        '
        'txtNo5
        '
        Me.txtNo5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtNo5.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNo5.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNo5.Location = New System.Drawing.Point(360, 40)
        Me.txtNo5.MaxLength = 3
        Me.txtNo5.Name = "txtNo5"
        Me.txtNo5.Size = New System.Drawing.Size(40, 22)
        Me.txtNo5.TabIndex = 12
        Me.txtNo5.Text = ""
        Me.txtNo5.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtNo4
        '
        Me.txtNo4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtNo4.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNo4.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNo4.Location = New System.Drawing.Point(312, 40)
        Me.txtNo4.MaxLength = 4
        Me.txtNo4.Name = "txtNo4"
        Me.txtNo4.Size = New System.Drawing.Size(40, 22)
        Me.txtNo4.TabIndex = 11
        Me.txtNo4.Text = ""
        Me.txtNo4.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtNo3
        '
        Me.txtNo3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtNo3.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNo3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNo3.Location = New System.Drawing.Point(264, 40)
        Me.txtNo3.MaxLength = 4
        Me.txtNo3.Name = "txtNo3"
        Me.txtNo3.Size = New System.Drawing.Size(40, 22)
        Me.txtNo3.TabIndex = 10
        Me.txtNo3.Text = ""
        Me.txtNo3.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtNo2
        '
        Me.txtNo2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtNo2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNo2.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNo2.Location = New System.Drawing.Point(216, 40)
        Me.txtNo2.MaxLength = 4
        Me.txtNo2.Name = "txtNo2"
        Me.txtNo2.Size = New System.Drawing.Size(40, 22)
        Me.txtNo2.TabIndex = 9
        Me.txtNo2.Text = ""
        Me.txtNo2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtNo1
        '
        Me.txtNo1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtNo1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNo1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNo1.Location = New System.Drawing.Point(168, 40)
        Me.txtNo1.MaxLength = 4
        Me.txtNo1.Name = "txtNo1"
        Me.txtNo1.Size = New System.Drawing.Size(40, 22)
        Me.txtNo1.TabIndex = 8
        Me.txtNo1.Text = ""
        Me.txtNo1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'rdBtnSelection
        '
        Me.rdBtnSelection.Location = New System.Drawing.Point(608, 32)
        Me.rdBtnSelection.Name = "rdBtnSelection"
        Me.rdBtnSelection.Size = New System.Drawing.Size(88, 24)
        Me.rdBtnSelection.TabIndex = 2
        Me.rdBtnSelection.Text = "Selection"
        '
        'rdBtnAll
        '
        Me.rdBtnAll.Checked = True
        Me.rdBtnAll.Location = New System.Drawing.Point(568, 32)
        Me.rdBtnAll.Name = "rdBtnAll"
        Me.rdBtnAll.Size = New System.Drawing.Size(48, 24)
        Me.rdBtnAll.TabIndex = 1
        Me.rdBtnAll.TabStop = True
        Me.rdBtnAll.Text = "All"
        '
        'lblApplyto
        '
        Me.lblApplyto.Location = New System.Drawing.Point(512, 32)
        Me.lblApplyto.Name = "lblApplyto"
        Me.lblApplyto.Size = New System.Drawing.Size(64, 23)
        Me.lblApplyto.TabIndex = 7
        Me.lblApplyto.Text = "Apply to:"
        Me.lblApplyto.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblCertifiedNo
        '
        Me.lblCertifiedNo.Location = New System.Drawing.Point(8, 40)
        Me.lblCertifiedNo.Name = "lblCertifiedNo"
        Me.lblCertifiedNo.Size = New System.Drawing.Size(160, 17)
        Me.lblCertifiedNo.TabIndex = 1
        Me.lblCertifiedNo.Text = "Certified Number Template"
        '
        'tbPageInvoiceReview
        '
        Me.tbPageInvoiceReview.Controls.Add(Me.Panel4)
        Me.tbPageInvoiceReview.Controls.Add(Me.Panel3)
        Me.tbPageInvoiceReview.Controls.Add(Me.Panel2)
        Me.tbPageInvoiceReview.Controls.Add(Me.Panel1)
        Me.tbPageInvoiceReview.Location = New System.Drawing.Point(4, 24)
        Me.tbPageInvoiceReview.Name = "tbPageInvoiceReview"
        Me.tbPageInvoiceReview.Size = New System.Drawing.Size(916, 662)
        Me.tbPageInvoiceReview.TabIndex = 1
        Me.tbPageInvoiceReview.Text = "Invoice Review"
        '
        'Panel4
        '
        Me.Panel4.Controls.Add(Me.ugInvoiceReview)
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel4.Location = New System.Drawing.Point(0, 24)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(916, 568)
        Me.Panel4.TabIndex = 3
        '
        'ugInvoiceReview
        '
        Me.ugInvoiceReview.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugInvoiceReview.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle
        Me.ugInvoiceReview.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugInvoiceReview.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugInvoiceReview.Location = New System.Drawing.Point(0, 0)
        Me.ugInvoiceReview.Name = "ugInvoiceReview"
        Me.ugInvoiceReview.Size = New System.Drawing.Size(916, 568)
        Me.ugInvoiceReview.TabIndex = 0
        '
        'Panel3
        '
        Me.Panel3.Controls.Add(Me.lblCharges)
        Me.Panel3.Controls.Add(Me.lblTotalNoOfTanks)
        Me.Panel3.Controls.Add(Me.lblTotalNoOfFacs)
        Me.Panel3.Controls.Add(Me.lblTotal)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel3.Location = New System.Drawing.Point(0, 592)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(916, 30)
        Me.Panel3.TabIndex = 2
        '
        'lblCharges
        '
        Me.lblCharges.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCharges.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblCharges.Location = New System.Drawing.Point(504, 0)
        Me.lblCharges.Name = "lblCharges"
        Me.lblCharges.Size = New System.Drawing.Size(100, 30)
        Me.lblCharges.TabIndex = 3
        Me.lblCharges.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblTotalNoOfTanks
        '
        Me.lblTotalNoOfTanks.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalNoOfTanks.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblTotalNoOfTanks.Location = New System.Drawing.Point(404, 0)
        Me.lblTotalNoOfTanks.Name = "lblTotalNoOfTanks"
        Me.lblTotalNoOfTanks.Size = New System.Drawing.Size(100, 30)
        Me.lblTotalNoOfTanks.TabIndex = 2
        Me.lblTotalNoOfTanks.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblTotalNoOfFacs
        '
        Me.lblTotalNoOfFacs.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalNoOfFacs.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblTotalNoOfFacs.Location = New System.Drawing.Point(304, 0)
        Me.lblTotalNoOfFacs.Name = "lblTotalNoOfFacs"
        Me.lblTotalNoOfFacs.Size = New System.Drawing.Size(100, 30)
        Me.lblTotalNoOfFacs.TabIndex = 1
        Me.lblTotalNoOfFacs.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblTotal
        '
        Me.lblTotal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotal.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblTotal.Location = New System.Drawing.Point(0, 0)
        Me.lblTotal.Name = "lblTotal"
        Me.lblTotal.Size = New System.Drawing.Size(304, 30)
        Me.lblTotal.TabIndex = 0
        Me.lblTotal.Text = "Totals for 3 Advices to Approve"
        Me.lblTotal.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.btnCancel)
        Me.Panel2.Controls.Add(Me.btnApprove)
        Me.Panel2.Controls.Add(Me.btnRegenerate)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel2.Location = New System.Drawing.Point(0, 622)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(916, 40)
        Me.Panel2.TabIndex = 1
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(472, 8)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 4
        Me.btnCancel.Text = "Cancel"
        '
        'btnApprove
        '
        Me.btnApprove.Location = New System.Drawing.Point(392, 8)
        Me.btnApprove.Name = "btnApprove"
        Me.btnApprove.TabIndex = 3
        Me.btnApprove.Text = "Approve"
        '
        'btnRegenerate
        '
        Me.btnRegenerate.Location = New System.Drawing.Point(304, 8)
        Me.btnRegenerate.Name = "btnRegenerate"
        Me.btnRegenerate.Size = New System.Drawing.Size(80, 23)
        Me.btnRegenerate.TabIndex = 2
        Me.btnRegenerate.Text = "Regenerate"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.btnInvoiceCaption)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(916, 24)
        Me.Panel1.TabIndex = 0
        '
        'btnInvoiceCaption
        '
        Me.btnInvoiceCaption.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnInvoiceCaption.Location = New System.Drawing.Point(8, 3)
        Me.btnInvoiceCaption.Name = "btnInvoiceCaption"
        Me.btnInvoiceCaption.Size = New System.Drawing.Size(160, 17)
        Me.btnInvoiceCaption.TabIndex = 0
        Me.btnInvoiceCaption.Text = "Invoice Advices to Approve"
        '
        'tbPageWaiveLateFees
        '
        Me.tbPageWaiveLateFees.Controls.Add(Me.pnlWaiveLateFeeDetails)
        Me.tbPageWaiveLateFees.Controls.Add(Me.pnlWaiveLateFeeBottom)
        Me.tbPageWaiveLateFees.Controls.Add(Me.pnlWaiveLateFeesTop)
        Me.tbPageWaiveLateFees.Location = New System.Drawing.Point(4, 24)
        Me.tbPageWaiveLateFees.Name = "tbPageWaiveLateFees"
        Me.tbPageWaiveLateFees.Size = New System.Drawing.Size(916, 662)
        Me.tbPageWaiveLateFees.TabIndex = 3
        Me.tbPageWaiveLateFees.Text = "Waive Late Fees"
        '
        'pnlWaiveLateFeeDetails
        '
        Me.pnlWaiveLateFeeDetails.Controls.Add(Me.ugWaiveLateFee)
        Me.pnlWaiveLateFeeDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlWaiveLateFeeDetails.Location = New System.Drawing.Point(0, 24)
        Me.pnlWaiveLateFeeDetails.Name = "pnlWaiveLateFeeDetails"
        Me.pnlWaiveLateFeeDetails.Size = New System.Drawing.Size(916, 598)
        Me.pnlWaiveLateFeeDetails.TabIndex = 2
        '
        'ugWaiveLateFee
        '
        Me.ugWaiveLateFee.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugWaiveLateFee.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle
        Me.ugWaiveLateFee.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugWaiveLateFee.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugWaiveLateFee.Location = New System.Drawing.Point(0, 0)
        Me.ugWaiveLateFee.Name = "ugWaiveLateFee"
        Me.ugWaiveLateFee.Size = New System.Drawing.Size(916, 598)
        Me.ugWaiveLateFee.TabIndex = 2
        '
        'pnlWaiveLateFeeBottom
        '
        Me.pnlWaiveLateFeeBottom.Controls.Add(Me.btnWaiveLateFeeCancel)
        Me.pnlWaiveLateFeeBottom.Controls.Add(Me.btnDeleteWaiver)
        Me.pnlWaiveLateFeeBottom.Controls.Add(Me.btnModifyWaiver)
        Me.pnlWaiveLateFeeBottom.Controls.Add(Me.btnAcceptDecisions)
        Me.pnlWaiveLateFeeBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlWaiveLateFeeBottom.Location = New System.Drawing.Point(0, 622)
        Me.pnlWaiveLateFeeBottom.Name = "pnlWaiveLateFeeBottom"
        Me.pnlWaiveLateFeeBottom.Size = New System.Drawing.Size(916, 40)
        Me.pnlWaiveLateFeeBottom.TabIndex = 3
        '
        'btnWaiveLateFeeCancel
        '
        Me.btnWaiveLateFeeCancel.Location = New System.Drawing.Point(624, 8)
        Me.btnWaiveLateFeeCancel.Name = "btnWaiveLateFeeCancel"
        Me.btnWaiveLateFeeCancel.TabIndex = 7
        Me.btnWaiveLateFeeCancel.Text = "Cancel"
        '
        'btnDeleteWaiver
        '
        Me.btnDeleteWaiver.Location = New System.Drawing.Point(520, 8)
        Me.btnDeleteWaiver.Name = "btnDeleteWaiver"
        Me.btnDeleteWaiver.Size = New System.Drawing.Size(96, 23)
        Me.btnDeleteWaiver.TabIndex = 6
        Me.btnDeleteWaiver.Text = "Delete Waiver"
        '
        'btnModifyWaiver
        '
        Me.btnModifyWaiver.Location = New System.Drawing.Point(416, 8)
        Me.btnModifyWaiver.Name = "btnModifyWaiver"
        Me.btnModifyWaiver.Size = New System.Drawing.Size(96, 23)
        Me.btnModifyWaiver.TabIndex = 5
        Me.btnModifyWaiver.Text = "Modify Waiver"
        '
        'btnAcceptDecisions
        '
        Me.btnAcceptDecisions.Location = New System.Drawing.Point(296, 8)
        Me.btnAcceptDecisions.Name = "btnAcceptDecisions"
        Me.btnAcceptDecisions.Size = New System.Drawing.Size(112, 23)
        Me.btnAcceptDecisions.TabIndex = 4
        Me.btnAcceptDecisions.Text = "Accept Decisions"
        '
        'pnlWaiveLateFeesTop
        '
        Me.pnlWaiveLateFeesTop.Controls.Add(Me.chkHideProcessed_Waivers)
        Me.pnlWaiveLateFeesTop.Controls.Add(Me.drBtnSelectAll)
        Me.pnlWaiveLateFeesTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlWaiveLateFeesTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlWaiveLateFeesTop.Name = "pnlWaiveLateFeesTop"
        Me.pnlWaiveLateFeesTop.Size = New System.Drawing.Size(916, 24)
        Me.pnlWaiveLateFeesTop.TabIndex = 0
        '
        'chkHideProcessed_Waivers
        '
        Me.chkHideProcessed_Waivers.Checked = True
        Me.chkHideProcessed_Waivers.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkHideProcessed_Waivers.Location = New System.Drawing.Point(16, 0)
        Me.chkHideProcessed_Waivers.Name = "chkHideProcessed_Waivers"
        Me.chkHideProcessed_Waivers.Size = New System.Drawing.Size(120, 24)
        Me.chkHideProcessed_Waivers.TabIndex = 15
        Me.chkHideProcessed_Waivers.Text = "Hide Processed"
        '
        'drBtnSelectAll
        '
        Me.drBtnSelectAll.Location = New System.Drawing.Point(288, 0)
        Me.drBtnSelectAll.Name = "drBtnSelectAll"
        Me.drBtnSelectAll.Size = New System.Drawing.Size(104, 20)
        Me.drBtnSelectAll.TabIndex = 1
        Me.drBtnSelectAll.Text = "Select All"
        Me.drBtnSelectAll.Visible = False
        '
        'AdministrativeServices
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(928, 694)
        Me.Controls.Add(Me.pnlAdminServicesDetails)
        Me.Name = "AdministrativeServices"
        Me.Text = "Administrative Services"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.pnlAdminServicesDetails.ResumeLayout(False)
        Me.tbCtrlAdminServices.ResumeLayout(False)
        Me.tbPageFYBasis.ResumeLayout(False)
        Me.pnlFYBasisDetails.ResumeLayout(False)
        CType(Me.ugFYBasis, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlFYBasisBottom.ResumeLayout(False)
        Me.pnlFYBasisDisplay.ResumeLayout(False)
        Me.tbPageLateFeeCertifications.ResumeLayout(False)
        Me.pnlLateFeeDetails.ResumeLayout(False)
        CType(Me.ugLateFee, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlLateFeeBottom.ResumeLayout(False)
        Me.pnlLateFeeTop.ResumeLayout(False)
        Me.Panel5.ResumeLayout(False)
        Me.tbPageInvoiceReview.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
        CType(Me.ugInvoiceReview, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel3.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.tbPageWaiveLateFees.ResumeLayout(False)
        Me.pnlWaiveLateFeeDetails.ResumeLayout(False)
        CType(Me.ugWaiveLateFee, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlWaiveLateFeeBottom.ResumeLayout(False)
        Me.pnlWaiveLateFeesTop.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region "Admin Services"
#Region "UI Support Routines"
    Private Sub frmClosing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        If sender.GetType.Name.IndexOf("LateFeeWaiverRequest") >= 0 Then
        ElseIf sender.GetType.Name.IndexOf("FiscalYearFeeBasis") >= 0 Then
            'ElseIf sender.GetType.Name.IndexOf("") >= 0 Then
            'ElseIf sender.GetType.Name.IndexOf("") >= 0 Then
        End If
    End Sub
    Private Sub frmClosed(ByVal sender As Object, ByVal e As System.EventArgs)
        If sender.GetType.Name.IndexOf("LateFeeWaiverRequest") >= 0 Then
            frmLateFeeWaiver = Nothing
        ElseIf sender.GetType.Name.IndexOf("FiscalYearFeeBasis") >= 0 Then
            frmFiscalYearBasis = Nothing
            'ElseIf sender.GetType.Name.IndexOf("") >= 0 Then
            '    frm = Nothing
            'ElseIf sender.GetType.Name.IndexOf("") >= 0 Then
            '    frm = Nothing
        End If
    End Sub
    Private Sub AdministrativeServices_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ugWaiveLateFee.DisplayLayout.ValueLists.Add("Waive_Approval_StatusValue")
        tbCtrlAdminServices.SelectedIndex = FEES_BASIS
        LoadFeesBasisForm()
        bolLoading = False
    End Sub

    Private Sub tbCtrlAdminServices_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbCtrlAdminServices.SelectedIndexChanged

        Select Case tbCtrlAdminServices.SelectedIndex
            Case FEES_BASIS
                LoadFeesBasisForm()
            Case INVOICE_REVIEW
                LoadInvoiceReviewForm()
            Case LATE_FEE_CERT
                LoadLateFeesCertificationForm()
            Case WAIVE_LATE_FEE
                LoadWaiveLateFeesForm()
        End Select
    End Sub

#End Region
#End Region

#Region "Fiscal Year Basis"
#Region "UI Support Routines"
    Private Sub ugFYBasis_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugFYBasis.AfterRowActivate
        If Not IsDBNull(ugFYBasis.ActiveRow.Cells("Invoice_Gen_Date").Value) Then
            If ugFYBasis.ActiveRow.Cells("Invoice_Gen_Date").Value < Now() Then
                btnModifyFiscalYearBasis.Enabled = False
                btnDeleteFiscalYearBasis.Enabled = False
            Else
                btnModifyFiscalYearBasis.Enabled = True
                btnDeleteFiscalYearBasis.Enabled = True
            End If
        End If
    End Sub
    Private Sub btnApprove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnApprove.Click
        Dim FeeBasisID As Int64
        Try
            btnApprove.Enabled = False
            btnRegenerate.Enabled = False

            If ugInvoiceReview.Rows.Count > 0 Then

                FeeBasisID = oFeesBasis.GetMaxFeeBasisID()

                oFeesBasis.MarkCalendarCompleted_ByDesc("Annual Billing was generated")

                oFeesBasis.Retrieve(FeeBasisID)
                oFeesBasis.ApprovedDate = Now.Date
                oFeesBasis.ApprovedTime = Now

                If oFeesBasis.ID <= 0 Then
                    oFeesBasis.CreatedBy = MusterContainer.AppUser.ID
                Else
                    oFeesBasis.ModifiedBy = MusterContainer.AppUser.ID
                End If
                oFeesBasis.Save(CType(UIUtilsGen.ModuleID.FeeAdmin, Integer), MusterContainer.AppUser.UserKey, returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                oFeesBasis.ApproveBilling()
                LoadInvoiceReviewForm()
                MsgBox("Approval Complete", MsgBoxStyle.OKOnly, "Success")
            Else
                MsgBox("There are no Invoices to Approve", MsgBoxStyle.OKOnly, "Fail")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            btnApprove.Enabled = True
            btnRegenerate.Enabled = True
        End Try

    End Sub

    Private Sub LoadFeesBasisForm()
        'ugFYBasis
        Dim dsLocal As DataSet
        Dim tmpBand As Int16

        dsLocal = oFeesBasis.GetFeesBasisGrid
        ugFYBasis.DataSource = dsLocal
        ugFYBasis.Rows.CollapseAll(True)
        ugFYBasis.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        ugFYBasis.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        If dsLocal.Tables(0).Rows.Count > 0 Then ' Does FeesBasis Table have rows

            ugFYBasis.DisplayLayout.Bands(0).Columns("Fees_Basis_ID").Hidden = True

            ugFYBasis.DisplayLayout.Bands(0).Columns("Fiscal_Year").Header.Caption = "Fiscal Year"
            ugFYBasis.DisplayLayout.Bands(0).Columns("Base_Fee").Header.Caption = "Base Fee"
            ugFYBasis.DisplayLayout.Bands(0).Columns("BaseUnit").Header.Caption = "Base Unit"
            ugFYBasis.DisplayLayout.Bands(0).Columns("LateFee").Header.Caption = "Late Fee"
            ugFYBasis.DisplayLayout.Bands(0).Columns("LatePeriod").Header.Caption = "Late Period"
            ugFYBasis.DisplayLayout.Bands(0).Columns("LateType").Header.Caption = "Late Type"
            ugFYBasis.DisplayLayout.Bands(0).Columns("Invoice_Gen_Date").Header.Caption = "Generation Date/Time"
            ugFYBasis.DisplayLayout.Bands(0).Columns("Invoice_APP_Date").Header.Caption = "Approval Date/Time"
            ugFYBasis.DisplayLayout.Bands(0).Columns("Starting_Date").Header.Caption = "Start Date"
            ugFYBasis.DisplayLayout.Bands(0).Columns("Ending_Date").Header.Caption = "End Date"
            ugFYBasis.DisplayLayout.Bands(0).Columns("Early_Grace").Header.Caption = "Early Grace"
            ugFYBasis.DisplayLayout.Bands(0).Columns("Late_Grace").Header.Caption = "Late Grace"

            ugFYBasis.DisplayLayout.Bands(0).Columns("Invoice_Gen_Date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugFYBasis.DisplayLayout.Bands(0).Columns("Invoice_APP_Date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugFYBasis.DisplayLayout.Bands(0).Columns("Starting_Date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugFYBasis.DisplayLayout.Bands(0).Columns("Ending_Date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugFYBasis.DisplayLayout.Bands(0).Columns("Early_Grace").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugFYBasis.DisplayLayout.Bands(0).Columns("Late_Grace").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            'ugFYBasis.DisplayLayout.Bands(0).Columns("PONumber").Width = 100

            ugFYBasis.DisplayLayout.Bands(0).Columns("Base_Fee").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugFYBasis.DisplayLayout.Bands(0).Columns("Base_Fee").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            ugFYBasis.DisplayLayout.Bands(0).Columns("LateFee").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugFYBasis.DisplayLayout.Bands(0).Columns("LateFee").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            ugFYBasis.DisplayLayout.Bands(0).Columns("LateType").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            If Not IsDBNull(ugFYBasis.Rows(0).Cells("Invoice_Gen_Date").Value) Then
                If ugFYBasis.Rows(0).Cells("Invoice_Gen_Date").Value < Now() Then
                    btnModifyFiscalYearBasis.Enabled = False
                    btnDeleteFiscalYearBasis.Enabled = False
                Else
                    btnModifyFiscalYearBasis.Enabled = True
                    btnDeleteFiscalYearBasis.Enabled = True
                End If
            End If
        End If

    End Sub

    Private Sub CreateModifyDeleteFiscalYearBasis(Optional ByVal strMode As String = Nothing)
        Dim MyFrm As MusterContainer
        Try
            If IsNothing(frmFiscalYearBasis) Then
                frmFiscalYearBasis = New FiscalYearFeeBasis
                AddHandler frmFiscalYearBasis.Closing, AddressOf frmClosing
                AddHandler frmFiscalYearBasis.Closed, AddressOf frmClosed
            End If
            frmFiscalYearBasis.pnlFyBasis.Enabled = True
            frmFiscalYearBasis.pblFYBasisDescription.Enabled = True
            frmFiscalYearBasis.btnDelete.Enabled = True
            frmFiscalYearBasis.btnDelete.Visible = True
            If Not strMode Is Nothing Then
                frmFiscalYearBasis.Mode = strMode
                If strMode.ToUpper.IndexOf("ADD") >= 0 Then
                    frmFiscalYearBasis.btnDelete.Enabled = False
                    frmFiscalYearBasis.btnDelete.Visible = False
                    frmFiscalYearBasis.FeeBasisID = 0
                    frmFiscalYearBasis.TemplateID = 0
                    If ugFYBasis.Rows.Count > 0 Then
                        If Not IsDBNull(ugFYBasis.ActiveRow.Cells("Fees_Basis_ID").Value) Then
                            frmFiscalYearBasis.TemplateID = ugFYBasis.ActiveRow.Cells("Fees_Basis_ID").Value
                        End If
                    End If
                ElseIf strMode.ToUpper.IndexOf("REGENERATE") >= 0 Then
                    frmFiscalYearBasis.pnlFyBasis.Enabled = False
                    frmFiscalYearBasis.pblFYBasisDescription.Enabled = False
                    frmFiscalYearBasis.FeeBasisID = 0
                    frmFiscalYearBasis.TemplateID = 0
                ElseIf strMode.ToUpper.IndexOf("MODIFY") >= 0 Then
                    frmFiscalYearBasis.FeeBasisID = ugFYBasis.ActiveRow.Cells("Fees_Basis_ID").Value
                    frmFiscalYearBasis.TemplateID = 0
                Else
                    frmFiscalYearBasis.FeeBasisID = ugFYBasis.ActiveRow.Cells("Fees_Basis_ID").Value
                    frmFiscalYearBasis.TemplateID = 0
                End If
            End If
            frmFiscalYearBasis.ShowDialog()
        Catch ex As Exception
            Throw ex
        Finally
            MyFrm = MdiParent

            MyFrm.RefreshCalendarInfo()
            MyFrm.LoadDueToMeCalendar()
            MyFrm.LoadToDoCalendar()
            LoadFeesBasisForm()
        End Try
    End Sub
#End Region
#Region "UI Control Events"
    Private Sub btnAddFiscalYearBasis_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddFiscalYearBasis.Click
        Try
            CreateModifyDeleteFiscalYearBasis("ADD")
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnModifyFiscalYearBasis_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnModifyFiscalYearBasis.Click
        Try
            CreateModifyDeleteFiscalYearBasis("MODIFY")
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnRegenerate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRegenerate.Click
        Try
            If ugInvoiceReview.Rows.Count > 0 Then
                If MsgBox("Do you wish to Regenerate Invoice Advice?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                    Exit Sub
                End If

                CreateModifyDeleteFiscalYearBasis("REGENERATE")
                LoadInvoiceReviewForm()
            Else
                MsgBox("There are no Invoices to Regenerate", MsgBoxStyle.OKOnly, "Failure")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnDeleteFiscalYearBasis_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDeleteFiscalYearBasis.Click
        Try
            CreateModifyDeleteFiscalYearBasis("DELETE")
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#End Region

#Region "Invoice Review"
    Private Sub LoadInvoiceReviewForm()

        Dim dsLocal As DataSet
        Dim tmpBand As Int16
        Dim AdviceCount As Int64
        Dim FacilityCount As Int64
        Dim TankCount As Int64
        Dim InvoiceTotal As Single
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow

        dsLocal = oFeesBasis.GetFeesPendingInvoiceHeaders
        ugInvoiceReview.DataSource = dsLocal
        ugInvoiceReview.Rows.CollapseAll(True)
        ugInvoiceReview.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        ugInvoiceReview.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect

        AdviceCount = 0
        FacilityCount = 0
        TankCount = 0
        InvoiceTotal = 0

        If dsLocal.Tables(0).Rows.Count > 0 Then ' Does FeesBasis Table have rows

            ugInvoiceReview.DisplayLayout.Bands(0).Columns("SFY").Header.Caption = "Fiscal Year"
            ugInvoiceReview.DisplayLayout.Bands(0).Columns("ADVICE_ID").Header.Caption = "Advice Number"
            ugInvoiceReview.DisplayLayout.Bands(0).Columns("Fee_Type").Header.Caption = "Billing Type"
            ugInvoiceReview.DisplayLayout.Bands(0).Columns("Owner_ID").Header.Caption = "Owner ID"
            ugInvoiceReview.DisplayLayout.Bands(0).Columns("Owner_Name").Header.Caption = "Owner Name"
            ugInvoiceReview.DisplayLayout.Bands(0).Columns("NumFacilities").Header.Caption = "# Of Facilities"
            ugInvoiceReview.DisplayLayout.Bands(0).Columns("NumTanks").Header.Caption = "# Of Tanks"
            ugInvoiceReview.DisplayLayout.Bands(0).Columns("INV_AMT").Header.Caption = "Charges"
            ugInvoiceReview.DisplayLayout.Bands(0).Columns("Date_Created").Header.Caption = "Created Date"
            ugInvoiceReview.DisplayLayout.Bands(0).Columns("Created_By").Header.Caption = "Created By"
            ugInvoiceReview.DisplayLayout.Bands(0).Columns("Description").Header.Caption = "Description"

            ugInvoiceReview.DisplayLayout.Bands(0).Columns("Date_Created").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            ugInvoiceReview.DisplayLayout.Bands(0).Columns("INV_AMT").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugInvoiceReview.DisplayLayout.Bands(0).Columns("INV_AMT").Header.Appearance.TextHAlign = Infragistics.Win.HAlign.Center

            For Each ugrow In ugInvoiceReview.Rows
                AdviceCount += 1
                FacilityCount += ugrow.Cells("NumFacilities").Value
                TankCount += ugrow.Cells("NumTanks").Value
                InvoiceTotal += ugrow.Cells("INV_AMT").Value
            Next
            'If Not IsDBNull(ugInvoiceReview.Rows(0).Cells("Invoice_Gen_Date").Value) Then
            '    If ugInvoiceReview.Rows(0).Cells("Invoice_Gen_Date").Value < Now() Then
            '        btnModifyFiscalYearBasis.Enabled = False
            '        btnDeleteFiscalYearBasis.Enabled = False
            '    Else
            '        btnModifyFiscalYearBasis.Enabled = True
            '        btnDeleteFiscalYearBasis.Enabled = True
            '    End If
            'End If
           
        End If

        lblTotal.Text = "Totals for " & AdviceCount & " Advices to Approve"
        lblTotalNoOfFacs.Text = FacilityCount
        lblTotalNoOfTanks.Text = TankCount
        lblCharges.Text = FormatNumber(InvoiceTotal, 2, TriState.True, TriState.True, TriState.True)
    End Sub
#End Region

#Region "Late Fee Certification"

    Private Sub LoadLateFeesCertificationForm()
        Dim dsLocal As DataSet
        Dim tmpBand As Int16
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow


        Try
            dsLocal = oLateFees.PopulateLateFeeCertificationGrid(isRedTag)

            dsLocal.Tables(0).Columns("Owner_ID").ReadOnly = True
            dsLocal.Tables(0).Columns("OwnerName").ReadOnly = True
            dsLocal.Tables(0).Columns("Inv_Amt").ReadOnly = True

            If Not isRedTag Then
                dsLocal.Tables(0).Columns("Inv_Number").ReadOnly = True
            Else
                dsLocal.Tables(0).Columns("facility_ID").ReadOnly = True
                dsLocal.Tables(0).Columns("facName").ReadOnly = True

                dsLocal.Tables(1).Columns("Facility_ID").ReadOnly = True
                dsLocal.Tables(1).Columns("Redtagremoved").ReadOnly = True
                dsLocal.Tables(1).Columns("dateprocessed").ReadOnly = True

            End If

            ugLateFee.DataSource = dsLocal
            ugLateFee.Rows.CollapseAll(True)
            ugLateFee.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
            'ugLateFee.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect


            If dsLocal.Tables(0).Rows.Count > 0 Then ' Does FeesBasis Table have rows
                ugLateFee.DisplayLayout.Bands(0).Columns("Owner_ID").TabStop = False
                ugLateFee.DisplayLayout.Bands(0).Columns("OwnerName").TabStop = False

                ugLateFee.DisplayLayout.Bands(0).Columns("Inv_Amt").TabStop = False

                If Not isRedTag Then
                    ugLateFee.DisplayLayout.Bands(0).Columns("Inv_Number").TabStop = False
                    ugLateFee.DisplayLayout.Bands(0).Columns("Inv_Number").Header.Caption = "Invoice"
                Else
                    ugLateFee.DisplayLayout.Bands(0).Columns("FacName").Header.Caption = "Facility Name"
                    ugLateFee.DisplayLayout.Bands(0).Columns("FacName").Width = 225
                    ugLateFee.DisplayLayout.Bands(0).Columns("Owner_ID").Width = 60
                    ugLateFee.DisplayLayout.Bands(0).Columns("Facility_ID").Width = 60
                    ugLateFee.DisplayLayout.Bands(0).Columns("Late_Cert_ID").Width = 60
                    ugLateFee.DisplayLayout.Bands(0).Columns("Status").Width = 60
                    ugLateFee.DisplayLayout.Bands(0).Columns("Facility_ID").TabStop = False
                    ugLateFee.DisplayLayout.Bands(0).Columns("Facility_ID").Header.Caption = "Facility #"
                    ugLateFee.DisplayLayout.Bands(0).Columns("Facility_ID").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Left
                    ugLateFee.DisplayLayout.Bands(0).Columns("FacName").TabStop = False
                    ugLateFee.DisplayLayout.Bands(0).Columns("YEAR").Hidden = True


                End If

                ugLateFee.DisplayLayout.Bands(0).Columns("Owner_ID").Header.Caption = "Owner ID"
                ugLateFee.DisplayLayout.Bands(0).Columns("OwnerName").Header.Caption = "Owner Name"
                ugLateFee.DisplayLayout.Bands(0).Columns("OwnerName").Width = 225
                ugLateFee.DisplayLayout.Bands(0).Columns("Inv_Amt").Header.Caption = "Charges"
                ugLateFee.DisplayLayout.Bands(0).Columns("Inv_Amt").Width = 75
                ugLateFee.DisplayLayout.Bands(0).Columns("Inv_Amt").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                ugLateFee.DisplayLayout.Bands(0).Columns("CertifiedMailNumber").Header.Caption = "Certified Letter Number"
                ugLateFee.DisplayLayout.Bands(0).Columns("CertifiedMailNumber").Width = 175
                ugLateFee.DisplayLayout.Bands(0).Columns("PROCESS_CERTIFICATION").Hidden = True
                ugLateFee.DisplayLayout.Bands(0).Columns("LATE_CERT_ID").Hidden = True

                If dsLocal.Tables(1).Rows.Count > 0 Then

                    If Not isRedTag Then
                        ugLateFee.DisplayLayout.Bands(1).Columns("Inv_Number").TabStop = False
                        ugLateFee.DisplayLayout.Bands(1).Columns("INV_LINE_AMT").TabStop = False
                        ugLateFee.DisplayLayout.Bands(1).Columns("Inv_Number").Hidden = True
                        ugLateFee.DisplayLayout.Bands(1).Columns("Inv_Line_Amt").Header.Caption = "Charges"
                        ugLateFee.DisplayLayout.Bands(1).Columns("Inv_Line_Amt").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                        ugLateFee.DisplayLayout.Bands(1).Columns("FacilityName").Header.Caption = "Facility Name"
                        ugLateFee.DisplayLayout.Bands(1).Columns("FacilityName").ColSpan = 2

                    Else
                        ugLateFee.DisplayLayout.Bands(1).Columns("Facility_id").Hidden = True
                        ugLateFee.DisplayLayout.Bands(1).Columns("owner_id").Hidden = True
                        ugLateFee.DisplayLayout.Bands(1).Columns("RedTagremoved").Header.Caption = "Tag Removed?"
                        ugLateFee.DisplayLayout.Bands(1).Columns("dateprocessed").Header.Caption = "Date Processed"
                        ugLateFee.DisplayLayout.Bands(1).Columns("RedTagremoved").Width = 100
                        ugLateFee.DisplayLayout.Bands(1).Columns("dateprocessed").Width = 80


                    End If

                    ugLateFee.DisplayLayout.Bands(1).Columns("Facility_ID").TabStop = False

                    ugLateFee.DisplayLayout.Bands(1).Columns("Facility_ID").Header.Caption = "Facility ID"



                End If

                ugLateFee.DisplayLayout.Bands(0).SortedColumns.Clear()

                If isRedTag Then
                    ugLateFee.DisplayLayout.Bands(0).SortedColumns.Add(ugLateFee.DisplayLayout.Bands(0).Columns("OWNER_ID"), True)
                Else
                    ugLateFee.DisplayLayout.Bands(0).SortedColumns.Add(ugLateFee.DisplayLayout.Bands(0).Columns("OwnerName"), False)

                End If
                ugLateFee.DisplayLayout.Bands(0).SortedColumns.RefreshSort(False)


                'ugLateFee.DisplayLayout.Bands(0).Columns("Invoice_Gen_Date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                ProcessHide()
            End If

            If Me.chkHideProcessed.Checked = False AndAlso Me.rdBtnRedTag.Checked = True Then
                Button1.Enabled = True
            Else
                Button1.Enabled = False
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Sub txtNo1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNo1.TextChanged
        If txtNo1.Text.Length = 4 Then
            txtNo2.Focus()
        End If
    End Sub

    Private Sub txtNo2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNo2.TextChanged
        If txtNo2.Text.Length = 4 Then
            txtNo3.Focus()
        End If
    End Sub

    Private Sub txtNo3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNo3.TextChanged
        If txtNo3.Text.Length = 4 Then
            txtNo4.Focus()
        End If
    End Sub

    Private Sub txtNo4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtNo4.TextChanged
        If txtNo4.Text.Length = 4 Then
            txtNo5.Focus()
        End If
    End Sub

    Private Sub ugLateFee_AfterEnterEditMode(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugLateFee.AfterEnterEditMode
        Dim strTemp As String
        Dim cCell As Infragistics.Win.UltraWinGrid.UltraGridCell
        Try
            cCell = sender.activecell

            cCell.SelStart = cCell.Text.Length

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Sub btnApplyTemplate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnApplyTemplate.Click
        Dim strTemp As String
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        strTemp = txtNo1.Text
        If Trim(txtNo2.Text) > "" Then
            strTemp = strTemp & " " & txtNo2.Text
        End If
        If Trim(txtNo3.Text) > "" Then
            strTemp = strTemp & " " & txtNo3.Text
        End If
        If Trim(txtNo4.Text) > "" Then
            strTemp = strTemp & " " & txtNo4.Text
        End If
        If Trim(txtNo5.Text) > "" Then
            strTemp = strTemp & " " & txtNo5.Text
        End If

        If strTemp.Length = 4 Or strTemp.Length = 9 Or strTemp.Length = 14 Or strTemp.Length = 19 Then
            strTemp = strTemp & " "
        End If
        If rdBtnSelection.Checked Then
            For Each ugrow In ugLateFee.Selected.Rows
                If ugrow.Cells("certifiedMailNumber").Value Is DBNull.Value OrElse (ugrow.Cells("CertifiedMailNumber").Value < " ") Then
                    ugrow.Cells("CertifiedMailNumber").Value = strTemp
                End If
            Next
        Else
            For Each ugrow In ugLateFee.Rows
                If ugrow.Cells("certifiedMailNumber").Value Is DBNull.Value OrElse (ugrow.Cells("CertifiedMailNumber").Value < " ") Then
                    ugrow.Cells("CertifiedMailNumber").Value = strTemp
                End If
            Next
        End If

    End Sub

    Private Sub chkHideProcessed_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkHideProcessed.CheckedChanged
        If bolLoading = True Then Exit Sub
        ProcessHide()

    End Sub

    Private Sub ProcessHide()
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow


        For Each ugrow In ugLateFee.Rows
            ugrow.Hidden = (ugrow.Cells("PROCESS_CERTIFICATION").Value = True And Me.chkHideProcessed.Checked) OrElse (ugrow.Cells("PROCESS_CERTIFICATION").Value = False And Not Me.chkHideProcessed.Checked)
        Next


        If chkHideProcessed.Checked Then
            Button1.Enabled = False
        ElseIf isRedTag Then
            Button1.Enabled = True
        End If

        For Each ugrow In ugLateFee.Rows
            If ugrow.Hidden = False Then
                ugLateFee.ActiveRow = ugrow
                Exit For
            End If
        Next
    End Sub
    Private Sub btnLateFeeSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLateFeeSave.Click
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim iCount As Int16
        Dim tmpstring As String
        Dim rows As ArrayList = New ArrayList
        Try
            iCount = 0



            If Me.rdBtnAll.Checked AndAlso (ugLateFee.Selected.Rows Is Nothing OrElse ugLateFee.Selected.Rows.Count = 0) Then
                For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In ugLateFee.Rows
                    rows.Add(row)
                Next

            Else
                For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In ugLateFee.Selected.Rows
                    rows.Add(row)
                Next
            End If


            If rows.Count > 0 Then

                For Each ugrow In rows
                    If ugrow.Band.Index = 0 Then

                        If ugrow.Cells("Process_Certification").Value > -1 Then

                            If IsDBNull(ugrow.Cells("CertifiedMailNumber").Value) = False Then
                                If Trim(ugrow.Cells("CertifiedMailNumber").Value) <> "" Then
                                    tmpstring = Replace(ugrow.Cells("CertifiedMailNumber").Value, " ", "", , , CompareMethod.Text)
                                    If tmpstring.Length = 20 Then

                                        If Not isRedTag Then
                                            oLateFees.Retrieve(ugrow.Cells("Late_Cert_ID").Value)
                                        Else
                                            oLateFees.Retrieve(0)

                                        End If

                                        Dim test As Integer = TestCertifiedMailNumber(ugrow.Cells("Late_Cert_ID").Value, ugrow.Cells("CertifiedMailNumber").Value, oLateFees, IIf(isRedTag, ugrow.Cells("OWNER_ID").Value, -1))

                                        If test = 0 Then

                                            If Not isRedTag Then

                                                oLateFees.CertLetterNumber = ugrow.Cells("CertifiedMailNumber").Value
                                                If oLateFees.ID <= 0 Then
                                                    oLateFees.CreatedBy = MusterContainer.AppUser.ID
                                                Else
                                                    oLateFees.ModifiedBy = MusterContainer.AppUser.ID
                                                End If  'If oLateFees.ID <= 0 Then

                                                oLateFees.Save(CType(UIUtilsGen.ModuleID.FeeAdmin, Integer), MusterContainer.AppUser.UserKey, returnVal)
                                            Else
                                                oLateFees.SaveRegTag(CType(UIUtilsGen.ModuleID.FeeAdmin, Integer), MusterContainer.AppUser.UserKey, ugrow.Cells("OWNER_ID").Value, ugrow.Cells("Process_Certification").Value, _
                                                                       Now.Year, ugrow.Cells("CertifiedMailNumber").Value, ugrow.Cells("Facility_ID").Value, IIf(ugrow.Cells("Late_Cert_ID").Value Is DBNull.Value, -1, ugrow.Cells("Late_Cert_ID").Value))

                                            End If


                                            If Not UIUtilsGen.HasRights(returnVal) Then
                                                Exit Sub
                                            End If  'If Not UIUtilsGen.HasRights(returnVal) Then


                                            iCount += 1
                                        ElseIf test = 2 Then
                                            MsgBox("owner: " + ugrow.Cells("Owner_ID").Value.ToString + "at fac#: " + ugrow.Cells("facility_ID").Value.ToString + "  --Certified Mail Number (" & ugrow.Cells("CertifiedMailNumber").Value & ") has to have the same certified mailing number as the all of the open facilties for that owner at this time .", MsgBoxStyle.OKOnly, "Invalid Certified Mail Number By Owner")
                                        Else
                                            MsgBox("Certified Mail Number (" & ugrow.Cells("CertifiedMailNumber").Value & ") has already been used.", MsgBoxStyle.OKOnly, "Invalid Certified Mail Number")
                                        End If      'If oLateFees.CheckForExistingCertNumber(ugrow.Cells("CertifiedMailNumber").Value) = False Then
                                    Else
                                        If tmpstring.Length <= 1 Then
                                            MsgBox("Certified Mail Number (" & ugrow.Cells("CertifiedMailNumber").Value & ") is not valid.", MsgBoxStyle.OKOnly, "Invalid Certified Mail Number")
                                        End If      'If tmpstring.Length > 1 Then

                                    End If          'If tmpstring.Length = 20 Then

                                End If              'If Trim(ugrow.Cells("CertifiedMailNumber").Value) <> "" Then

                            End If                  'If IsDBNull(ugrow.Cells("CertifiedMailNumber").Value) = False Then

                        End If                      'If ugrow.Cells("Process_Certification").Value > -1 Then

                    End If                          'If ugrow.Band.Index = 0 Then
                Next

            End If
            LoadLateFeesCertificationForm()
            MsgBox(iCount & " Records Saved.", MsgBoxStyle.Information, "Records Saved")

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            rows.Clear()

            rows = Nothing
        End Try
    End Sub

    Private Sub RdoButtonRedTag(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdBtnRedTag.CheckedChanged, rdBtnLateFee.CheckedChanged

        If Me.rdBtnLateFee.Checked AndAlso isRedTag Then
            isRedTag = False
            LoadLateFeesCertificationForm()
        ElseIf Me.rdBtnRedTag.Checked AndAlso Not isRedTag Then
            isRedTag = True
            LoadLateFeesCertificationForm()
        End If

    End Sub

    Private Sub UnProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim iCount As Int16
        Dim oLetter As New Reg_Letters
        Dim oOwner As New MUSTER.BusinessLogic.pOwner
        Dim tmpString As String
        Dim strFacilities As String = String.Empty

        Try
            iCount = 0
            If ugLateFee.Rows.Count > 0 Then

                For Each ugrow In ugLateFee.Rows
                    If ugrow.Band.Index = 0 Then

                        If ugrow.Cells("Process_Certification").Value = True AndAlso ugrow.Selected Then


                            If False = False Then


                                oOwner.Retrieve(ugrow.Cells("Owner_ID").Value)

                                'Call Letter Generation Here.
                                'oLateFees.ProcessCertification = True
                                ' oLetter.GenerateFeesLetter("Late Fee Waiver", "LateFeeWaiver", "Late Fee Waiver Letter", "Late_Fee_Letter_Template.doc", oOwner, ugrow.Cells("Late_Cert_ID").Value)

                                strFacilities = String.Empty


                                oLateFees.SaveRegTag(CType(UIUtilsGen.ModuleID.FeeAdmin, Integer), MusterContainer.AppUser.UserKey, ugrow.Cells("OWNER_ID").Value, False, _
                                                     Now.Year, ugrow.Cells("CertifiedMailNumber").Value.ToString.Replace("  ", " ").Replace("  ", " ").Trim, ugrow.Cells("Facility_ID").Value, ugrow.Cells("LATE_CERT_ID").Value)


                            End If

                            If Not UIUtilsGen.HasRights(returnVal) Then
                                Exit Sub
                            End If  'If Not UIUtilsGen.HasRights(returnVal) Then

                            iCount += 1
                        End If      'If oLateFees.CheckForExistingCertNumber(ugrow.Cells("CertifiedMailNumber").Value) = False Then

                    End If
                    'If ugrow.Band.Index = 0 Then
                Next


            End If
            LoadLateFeesCertificationForm()
            MsgBox(iCount & " Records Processed.", MsgBoxStyle.Information, "Records Un - Processed")

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Sub btnLateFeeProcess_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLateFeeProcess.Click
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim iCount As Int16
        Dim oLetter As New Reg_Letters
        Dim oOwner As New MUSTER.BusinessLogic.pOwner
        Dim tmpString As String
        Dim strFacilities As String = String.Empty
        Dim rows As System.Collections.ArrayList


        rows = New ArrayList

        Try
            iCount = 0

            ugLateFee.DisplayLayout.Bands(0).SortedColumns.Clear()
            ugLateFee.DisplayLayout.Bands(0).SortedColumns.Add(ugLateFee.DisplayLayout.Bands(0).Columns("OWNER_ID"), True)
            ugLateFee.DisplayLayout.Bands(0).SortedColumns.RefreshSort(False)



            If Me.rdBtnAll.Checked AndAlso (ugLateFee.Selected.Rows Is Nothing OrElse ugLateFee.Selected.Rows.Count = 0) Then
                For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In ugLateFee.Rows
                    rows.Add(row)
                Next

            Else
                For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In ugLateFee.Selected.Rows
                    rows.Add(row)
                Next
            End If



            If rows.Count > 0 Then

                Dim counter As Integer = 0
                Dim IDs As String = ","


                For Each ugrow In rows

                    If ugrow.Band.Index = 0 Then

                        If ugrow.Cells("Process_Certification").Value > -1 Then
                            If IsDBNull(ugrow.Cells("CertifiedMailNumber").Value) = False Then
                                If Trim(ugrow.Cells("CertifiedMailNumber").Value) <> "" Then
                                    tmpString = Replace(ugrow.Cells("CertifiedMailNumber").Value, " ", "", , , CompareMethod.Text)
                                    If tmpString.Length = 20 Then
                                        If Not isRedTag Then
                                            oLateFees.Retrieve(ugrow.Cells("Late_Cert_ID").Value)
                                        Else
                                            oLateFees.Retrieve(0)
                                        End If



                                        Dim test As Integer = TestCertifiedMailNumber(ugrow.Cells("Late_Cert_ID").Value, ugrow.Cells("CertifiedMailNumber").Value, oLateFees, IIf(isRedTag, ugrow.Cells("OWNER_ID").Value, -1))

                                        If test = 0 Then

                                            oOwner.Retrieve(ugrow.Cells("Owner_ID").Value)

                                            'Call Letter Generation Here.
                                            'oLateFees.ProcessCertification = True
                                            ' oLetter.GenerateFeesLetter("Late Fee Waiver", "LateFeeWaiver", "Late Fee Waiver Letter", "Late_Fee_Letter_Template.doc", oOwner, ugrow.Cells("Late_Cert_ID").Value)


                                            If ugrow.HasChild AndAlso Not isRedTag Then

                                                strFacilities = String.Empty

                                                For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In ugrow.ChildBands(0).Rows
                                                    If strFacilities.Length > 0 Then
                                                        strFacilities = String.Format("{3},{0}Fac #{1}  -  {2}", Chr(7), row.Cells("FACILITY_ID").Value, row.Cells("FACILITyNAME").Value, strFacilities)
                                                    Else
                                                        strFacilities = String.Format("Fac #{0}  -  {1}", row.Cells("FACILITY_ID").Value, row.Cells("FACilityName").Value)
                                                    End If
                                                Next
                                            ElseIf isRedTag Then

                                                strFacilities = String.Format("{0}{1},", strFacilities, ugrow.Cells("FACILITY_ID").Value)

                                            End If




                                            If Not isRedTag Then

                                                oLateFees.CertLetterNumber = ugrow.Cells("CertifiedMailNumber").Value.ToString.Replace("  ", " ").Replace("  ", " ").Trim

                                                If oLateFees.ID <= 0 Then
                                                    oLateFees.CreatedBy = MusterContainer.AppUser.ID
                                                Else
                                                    oLateFees.ModifiedBy = MusterContainer.AppUser.ID
                                                End If  'If oLateFees.ID <= 0 Then
                                                oLateFees.Save(CType(UIUtilsGen.ModuleID.FeeAdmin, Integer), MusterContainer.AppUser.UserKey, returnVal)

                                                oLateFees.ProcessCertification = True

                                             

                                                'oLetter.GenerateFeesLetter("Late Fee Waiver", "LateFeeWaiver", "Late Fee Waiver Letter", "Late_Fee_Letter_Template.doc", oOwner, ugrow.Cells("Late_Cert_ID").Value, strFacilities, ugrow.Cells("CertifiedMailNumber").Value.ToString.Replace("  ", " ").Replace("  ", " ").Trim)
                                                IDs = String.Format("{0}{1},", IDs, oLateFees.ID)

                                                oLateFees.Save(CType(UIUtilsGen.ModuleID.FeeAdmin, Integer), MusterContainer.AppUser.UserKey, returnVal)

                                            Else

                                                Dim id As Integer = -1

                                                id = oLateFees.SaveRegTag(CType(UIUtilsGen.ModuleID.FeeAdmin, Integer), MusterContainer.AppUser.UserKey, ugrow.Cells("OWNER_ID").Value, ugrow.Cells("Process_Certification").Value, _
                                                                     Now.Year, ugrow.Cells("CertifiedMailNumber").Value.ToString.Replace("  ", " ").Replace("  ", " ").Trim, ugrow.Cells("Facility_ID").Value, IIf(ugrow.Cells("Late_Cert_ID").Value Is DBNull.Value, -1, ugrow.Cells("Late_Cert_ID").Value))

                                                'If counter = rows.Count - 1 OrElse ugrow.Cells("OWNER_ID").Value <> rows(counter + 1).Cells("OWNER_ID").Value Then
                                                'End If

                                            oLateFees.SaveRegTag(CType(UIUtilsGen.ModuleID.FeeAdmin, Integer), MusterContainer.AppUser.UserKey, ugrow.Cells("OWNER_ID").Value, True, _
                                                                 Now.Year, ugrow.Cells("CertifiedMailNumber").Value.ToString.Replace("  ", " ").Replace("  ", " ").Trim, ugrow.Cells("Facility_ID").Value, id)

                                        End If



                                            If Not UIUtilsGen.HasRights(returnVal) Then
                                                Exit Sub
                                            End If  'If Not UIUtilsGen.HasRights(returnVal) Then


                                            iCount += 1

                                        ElseIf test = 2 Then
                                            MsgBox("owner: " + ugrow.Cells("Owner_ID").Value + "at fac#: " + ugrow.Cells("facility_ID").Value + "  --Certified Mail Number (" & ugrow.Cells("CertifiedMailNumber").Value & ") has to have the same certified mailing number as the all of the open facilties for that owner at this time .", MsgBoxStyle.OKOnly, "Invalid Certified Mail Number By Owner")
                                        Else
                                            MsgBox("Certified Mail Number (" & ugrow.Cells("CertifiedMailNumber").Value & ") has already been used.", MsgBoxStyle.OKOnly, "Invalid Certified Mail Number")
                                        End If      'If oLateFees.CheckForExistingCertNumber(ugrow.Cells("CertifiedMailNumber").Value) = False Then
                                    Else
                                        If tmpString.Length > 1 Then
                                            MsgBox("Certified Mail Number (" & ugrow.Cells("CertifiedMailNumber").Value & ") is not valid.", MsgBoxStyle.OKOnly, "Invalid Certified Mail Number")
                                        End If      'If tmpstring.Length > 1 Then

                                    End If          'If tmpstring.Length = 20 Then

                                End If              'If Trim(ugrow.Cells("CertifiedMailNumber").Value) <> "" Then

                            End If                  'If IsDBNull(ugrow.Cells("CertifiedMailNumber").Value) = False Then

                        End If                      'If ugrow.Cells("Process_Certification").Value > -1 Then

                    End If                          'If ugrow.Band.Index = 0 Then

                    counter += 1
                Next

                If Not isRedTag AndAlso IDs.Length > 0 Then
                    ''Call Letter Generation Here.
                    Dim frmReport As ReportDisplay

                    frmReport = New ReportDisplay
                    frmReport.MdiParent = _container

                    Dim fy As Integer = oFeesBasis.GetFiscalYearForFee(Date.Now)

                    frmReport.GenerateReport("Late Fee Certifications", New Object() {" ", IDs, "1/1/1900", fy}, "Would like print out a PDF of the generated Late Fee Waiver(s)?", String.Format("LateFeeCertificationWaiver_{0}", Now.ToString("MMM_dd_yyyy_HHmmss")), UIUtilsGen.ModuleID.FeeAdmin)

                ElseIf isRedTag AndAlso strFacilities.Length > 0 Then

                    ''Call Letter Generation Here.

                    Dim frmReport2 As ReportDisplay

                    frmReport2 = New ReportDisplay
                    frmReport2.MdiParent = _container


                    frmReport2.GenerateReport("Red Tag Certification Letter", New Object() {0, " ", strFacilities}, "Would like print out a PDF of the generated Red Tag Letter(s)?", String.Format("RedTagCertLetter_{0}", Now.ToString("MMM_dd_yyyy_HHmmss")), UIUtilsGen.ModuleID.FeeAdmin)

                    'oLetter.GenerateFeesLetter("Red Tag Letter", "RedTagletter", "Red Tag Outstanding Balance Letter", "RedTagLetter_Template.doc", oOwner, ugrow.Cells("Late_Cert_ID").Value, strFacilities, ugrow.Cells("CertifiedMailNumber").Value.ToString.Replace("  ", " ").Replace("  ", " ").Trim)

                    strFacilities = String.Empty


                End If
            End If



            LoadLateFeesCertificationForm()
            MsgBox(iCount & " Records Processed.", MsgBoxStyle.Information, "Records Processed")

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()

        Finally

            rows.Clear()
            rows = Nothing

        End Try
    End Sub


    Private Function TestCertifiedMailNumber(ByVal id As Integer, ByVal CertMailNumber As String, ByVal oLtFee As MUSTER.BusinessLogic.pFeeLateFee, Optional ByVal ownerID As Integer = -1) As Integer

        If Not oLtFee Is Nothing AndAlso Not oLtFee.CertLetterNumber Is Nothing AndAlso oLtFee.CertLetterNumber = CertMailNumber Then
            Return 0
        End If

        Dim ownerIsSame
        Dim pass As Boolean = oLtFee.CheckForExistingCertNumber(CertMailNumber, id, Me.isRedTag, ownerID, ownerIsSame)

        If ownerIsSame Then
            Return 2
        ElseIf pass Then
            Return 1
        Else
            Return 0
        End If

    End Function
#End Region

#Region "Waive Late Fees"
    Private Sub LoadWaiveLateFeesForm()
        Dim dsLocal As DataSet
        Dim dsLocal2 As DataTable
        Dim tmpBand As Int16
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim drRow As DataRow

        dsLocal = oLateFees.PopulateWaiveLateFeeGrid

        dsLocal.Tables(0).Columns("Owner_ID").ReadOnly = True
        dsLocal.Tables(0).Columns("OwnerName").ReadOnly = True
        dsLocal.Tables(0).Columns("Inv_Amt").ReadOnly = True
        dsLocal.Tables(0).Columns("Inv_Number").ReadOnly = True
        dsLocal.Tables(1).Columns("Facility_ID").ReadOnly = True
        dsLocal.Tables(1).Columns("FacilityName").ReadOnly = True
        dsLocal.Tables(1).Columns("INV_LINE_AMT").ReadOnly = True
        dsLocal.Tables(1).Columns("Inv_Number").ReadOnly = True


        ugWaiveLateFee.DataSource = dsLocal
        ugWaiveLateFee.Rows.CollapseAll(True)
        ugWaiveLateFee.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        'ugWaiveLateFee.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect

        'If chkHideProcessed.Checked Then
        '    For Each ugrow In ugLateFee.Rows
        '        If ugrow.Cells("CertifiedMailNumber").Value > "" Then
        '            ugrow.Hidden = True
        '        End If
        '    Next
        'End If

        If ugWaiveLateFee.DisplayLayout.ValueLists("Waive_Approval_StatusValue").ValueListItems.Count < 3 Then
            dsLocal2 = oLateFees.PopulateLateFeeWaiverDecisions
            For Each drRow In dsLocal2.Rows
                ugWaiveLateFee.DisplayLayout.ValueLists("Waive_Approval_StatusValue").ValueListItems.Add(drRow.Item("PROPERTY_ID"), drRow.Item("PROPERTY_NAME"))
            Next
            ugWaiveLateFee.DisplayLayout.Bands(0).Columns("Waive_Approval_Status").ValueList = ugWaiveLateFee.DisplayLayout.ValueLists("Waive_Approval_StatusValue")
            ugWaiveLateFee.DisplayLayout.Bands(0).Columns("Waive_Approval_Status").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate
            ugWaiveLateFee.DisplayLayout.ValueLists("Waive_Approval_StatusValue").DisplayStyle = Infragistics.Win.ValueListDisplayStyle.DisplayText
        End If


        If dsLocal.Tables(0).Rows.Count > 0 Then ' Does FeesBasis Table have rows
            ugWaiveLateFee.DisplayLayout.Bands(0).Columns("Owner_ID").TabStop = False
            ugWaiveLateFee.DisplayLayout.Bands(0).Columns("OwnerName").TabStop = False
            ugWaiveLateFee.DisplayLayout.Bands(0).Columns("Inv_Amt").TabStop = False
            ugWaiveLateFee.DisplayLayout.Bands(0).Columns("Inv_Number").TabStop = False
            ugWaiveLateFee.DisplayLayout.Bands(0).Columns("PROCESS_WAIVER").Hidden = True
            ugWaiveLateFee.DisplayLayout.Bands(0).Columns("LATE_CERT_ID").Hidden = True
            If chkHideProcessed_Waivers.Checked Then
                ugWaiveLateFee.DisplayLayout.Bands(0).Columns("WaiverFinalized").Hidden = True
            Else
                ugWaiveLateFee.DisplayLayout.Bands(0).Columns("WaiverFinalized").Hidden = False
            End If

            ugWaiveLateFee.DisplayLayout.Bands(0).Columns("Waive_approval_status").Header.Caption = "Director Decision"
            ugWaiveLateFee.DisplayLayout.Bands(0).Columns("Owner_ID").Header.Caption = "Owner ID"
            ugWaiveLateFee.DisplayLayout.Bands(0).Columns("OwnerName").Header.Caption = "Owner Name"
            ugWaiveLateFee.DisplayLayout.Bands(0).Columns("OwnerName").Width = 200
            ugWaiveLateFee.DisplayLayout.Bands(0).Columns("Inv_Amt").Header.Caption = "Charges"
            ugWaiveLateFee.DisplayLayout.Bands(0).Columns("Inv_Amt").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
            ugWaiveLateFee.DisplayLayout.Bands(0).Columns("Inv_Number").Header.Caption = "Invoice"
            ugWaiveLateFee.DisplayLayout.Bands(0).Columns("Excuse").Width = 150

            If dsLocal.Tables(1).Rows.Count > 0 Then
                ugWaiveLateFee.DisplayLayout.Bands(1).Columns("Facility_ID").TabStop = False
                ugWaiveLateFee.DisplayLayout.Bands(1).Columns("FacilityName").TabStop = False
                ugWaiveLateFee.DisplayLayout.Bands(1).Columns("INV_LINE_AMT").TabStop = False
                ugWaiveLateFee.DisplayLayout.Bands(1).Columns("Inv_Number").TabStop = False

                ugWaiveLateFee.DisplayLayout.Bands(1).Columns("Facility_ID").Header.Caption = "Facility ID"
                ugWaiveLateFee.DisplayLayout.Bands(1).Columns("FacilityName").Header.Caption = "Facility Name"
                ugWaiveLateFee.DisplayLayout.Bands(1).Columns("FacilityName").ColSpan = 3
                ugWaiveLateFee.DisplayLayout.Bands(1).Columns("INV_LINE_AMT").Header.Caption = "Charges"
                ugWaiveLateFee.DisplayLayout.Bands(1).Columns("INV_LINE_AMT").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                ugWaiveLateFee.DisplayLayout.Bands(1).Columns("Inv_Number").Hidden = True
            End If


            'ugWaiveLateFee.DisplayLayout.Bands(0).Columns("Invoice_Gen_Date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            'If Not IsDBNull(ugLateFee.Rows(0).Cells("Invoice_Gen_Date").Value) Then
            '    If ugWaiveLateFee.Rows(0).Cells("Invoice_Gen_Date").Value < Now() Then
            '        btnModifyFiscalYearBasis.Enabled = False
            '        btnDeleteFiscalYearBasis.Enabled = False
            '    Else
            '        btnModifyFiscalYearBasis.Enabled = True
            '        btnDeleteFiscalYearBasis.Enabled = True
            '    End If
            'End If
        End If
        ProcessHideWaiverRows()

    End Sub

    Private Sub ModifyDeleteLateFeeWaiver()
        Try
            If IsNothing(frmLateFeeWaiver) Then
                frmLateFeeWaiver = New LateFeeWaiverRequest
                AddHandler frmLateFeeWaiver.Closing, AddressOf frmClosing
                AddHandler frmLateFeeWaiver.Closed, AddressOf frmClosed
                frmLateFeeWaiver.Text = "Late Waiver Request Maintenance"

            End If
            If ugWaiveLateFee.ActiveRow.Band.Index = 1 Then
                Exit Sub
            End If
            frmLateFeeWaiver.LateCertID = ugWaiveLateFee.ActiveRow.Cells("Late_Cert_ID").Value
            frmLateFeeWaiver.ShowDialog()
            LoadWaiveLateFeesForm()

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnDeleteWaiver_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDeleteWaiver.Click
        Try
            ModifyDeleteLateFeeWaiver()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnModifyWaiver_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnModifyWaiver.Click
        Try
            ModifyDeleteLateFeeWaiver()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugWaiveLateFee_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugWaiveLateFee.DoubleClick
        Try
            ModifyDeleteLateFeeWaiver()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnAcceptDecisions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAcceptDecisions.Click
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim iCount As Int16
        Try
            iCount = 0
            If ugWaiveLateFee.Rows.Count > 0 Then

                For Each ugrow In ugWaiveLateFee.Rows
                    If ugrow.Band.Index = 0 Then
                        If ugrow.Cells("Waive_Approval_Status").Value > -1 And ugrow.Cells("Process_Waiver").Value = 0 Then
                            oLateFees.Retrieve(ugrow.Cells("Late_Cert_ID").Value)

                            oLateFees.ProcessWaiver = True
                            oLateFees.WaiveApprovalStatus = ugrow.Cells("Waive_Approval_Status").Value
                            oLateFees.WaiverFinalizedOn = Now.Date

                            ' Create Credit Memo 
                            If oLateFees.ID <= 0 Then
                                oLateFees.CreatedBy = MusterContainer.AppUser.ID
                            Else
                                oLateFees.ModifiedBy = MusterContainer.AppUser.ID
                            End If
                            oLateFees.Save(CType(UIUtilsGen.ModuleID.FeeAdmin, Integer), MusterContainer.AppUser.UserKey, returnVal)
                            If Not UIUtilsGen.HasRights(returnVal) Then
                                Exit Sub
                            End If


                            iCount += 1
                        End If
                    End If
                Next

            End If
            LoadWaiveLateFeesForm()
            MsgBox(iCount & " Records Processed.", MsgBoxStyle.Information, "Records Processed")

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Sub chkHideProcessed_Waivers_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkHideProcessed_Waivers.CheckedChanged
        If bolLoading Then Exit Sub

        If chkHideProcessed_Waivers.Checked Then
            ugWaiveLateFee.DisplayLayout.Bands(0).Columns("WaiverFinalized").Hidden = True
        Else
            ugWaiveLateFee.DisplayLayout.Bands(0).Columns("WaiverFinalized").Hidden = False
        End If
        ProcessHideWaiverRows()

    End Sub

    Private Sub ProcessHideWaiverRows()

        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow

        For Each ugrow In ugWaiveLateFee.Rows
            If chkHideProcessed_Waivers.Checked Then
                If ugrow.Cells("Process_Waiver").Value Then
                    ugrow.Hidden = True
                End If
            Else
                ugrow.Hidden = False
            End If
        Next
    End Sub
#End Region


    Private Sub btnLateFeeCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLateFeeCancel.Click
        Me.Close()
    End Sub


    
    Private Sub chkHideProcessed_CheckStateChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkHideProcessed.CheckStateChanged

    End Sub
End Class
