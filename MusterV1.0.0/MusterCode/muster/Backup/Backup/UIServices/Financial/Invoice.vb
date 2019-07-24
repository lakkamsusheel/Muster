Public Class Invoice
    Inherits System.Windows.Forms.Form

#Region " Local Variables "

    Friend FinancialEventID As Int64
    Friend FinancialInvoiceID As Int64
    Friend SelectedRequestId As Int64 = 0
    Private NegativeRequest As Boolean = False

    Private bolLoading As Boolean
    Private bolFormatting As Boolean
    Private oFinancialEvent As New MUSTER.BusinessLogic.pFinancial
    Private oFinancialActivity As New MUSTER.BusinessLogic.pFinancialActivity
    Private oFinancialCommitment As New MUSTER.BusinessLogic.pFinancialCommitment
    Private oFinancialInvoice As New MUSTER.BusinessLogic.pFinancialInvoice
    Private oFinancialReimbursement As New MUSTER.BusinessLogic.pFinancialReimbursement
    Private sStartingPaid As Double
    Private sStartingRequested As Double
    Private sFinalTotal As Double
    Private loading As Boolean = True
    Private nContactID As Integer = 0

    Private returnVal As String = String.Empty
    Private bolFinalOnLoad As Boolean


#End Region

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal contactID)
        MyBase.New()

        nContactID = contactID


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
    Friend WithEvents lblID As System.Windows.Forms.Label
    Friend WithEvents lblIDValue As System.Windows.Forms.Label
    Friend WithEvents lblReimbursementRequest As System.Windows.Forms.Label
    Friend WithEvents pnlInvoiceBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlInvoiceTop As System.Windows.Forms.Panel
    Friend WithEvents pnlInvoiceDetails As System.Windows.Forms.Panel
    Friend WithEvents lblCommitment As System.Windows.Forms.Label
    Friend WithEvents cmbCommitment As System.Windows.Forms.ComboBox
    Friend WithEvents lblPOValue As System.Windows.Forms.Label
    Friend WithEvents lblPO As System.Windows.Forms.Label
    Friend WithEvents lblInvoiceNumber As System.Windows.Forms.Label
    Friend WithEvents txtInvoiceNumber As System.Windows.Forms.TextBox
    Friend WithEvents tbCtrlInvoice As System.Windows.Forms.TabControl
    Friend WithEvents tbPageInvoice As System.Windows.Forms.TabPage
    Friend WithEvents tbPageCommitment As System.Windows.Forms.TabPage
    Friend WithEvents tbPageAdjustments As System.Windows.Forms.TabPage
    Friend WithEvents tbPagePastInvoices As System.Windows.Forms.TabPage
    Friend WithEvents chkOnHold As System.Windows.Forms.CheckBox
    Friend WithEvents chkFinal As System.Windows.Forms.CheckBox
    Friend WithEvents lblInvoiced As System.Windows.Forms.Label
    Friend WithEvents txtInvoiced As System.Windows.Forms.TextBox
    Friend WithEvents lblPaid As System.Windows.Forms.Label
    Friend WithEvents txtPaid As System.Windows.Forms.TextBox
    Friend WithEvents cmbDeduction As System.Windows.Forms.ComboBox
    Friend WithEvents btnInsert As System.Windows.Forms.Button
    Friend WithEvents lblDeductionReason As System.Windows.Forms.Label
    Friend WithEvents txtDeductionReason As System.Windows.Forms.TextBox
    Friend WithEvents lblDeduction As System.Windows.Forms.Label
    Friend WithEvents lblComment As System.Windows.Forms.Label
    Friend WithEvents txtComments As System.Windows.Forms.TextBox
    Friend WithEvents lblCommitmentComment As System.Windows.Forms.Label
    Friend WithEvents txtCommitmentComment As System.Windows.Forms.TextBox
    Friend WithEvents lblTechnicalReports As System.Windows.Forms.Label
    Friend WithEvents ugTechReports As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents ugAdjustment As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents ugPastInvoices As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents cmbReimbursementRequest As System.Windows.Forms.ComboBox
    Friend WithEvents txtReimbursementRequest As System.Windows.Forms.TextBox
    Friend WithEvents txtCommitment As System.Windows.Forms.TextBox
    Friend WithEvents cfCommitment As MUSTER.CostFormat
    Friend WithEvents dtPickPaymentDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblPaymentDate As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlInvoiceBottom = New System.Windows.Forms.Panel
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOK = New System.Windows.Forms.Button
        Me.pnlInvoiceTop = New System.Windows.Forms.Panel
        Me.txtInvoiceNumber = New System.Windows.Forms.TextBox
        Me.lblInvoiceNumber = New System.Windows.Forms.Label
        Me.lblPO = New System.Windows.Forms.Label
        Me.lblPOValue = New System.Windows.Forms.Label
        Me.cmbCommitment = New System.Windows.Forms.ComboBox
        Me.lblCommitment = New System.Windows.Forms.Label
        Me.lblReimbursementRequest = New System.Windows.Forms.Label
        Me.lblIDValue = New System.Windows.Forms.Label
        Me.lblID = New System.Windows.Forms.Label
        Me.cmbReimbursementRequest = New System.Windows.Forms.ComboBox
        Me.txtReimbursementRequest = New System.Windows.Forms.TextBox
        Me.txtCommitment = New System.Windows.Forms.TextBox
        Me.dtPickPaymentDate = New System.Windows.Forms.DateTimePicker
        Me.lblPaymentDate = New System.Windows.Forms.Label
        Me.pnlInvoiceDetails = New System.Windows.Forms.Panel
        Me.tbCtrlInvoice = New System.Windows.Forms.TabControl
        Me.tbPageInvoice = New System.Windows.Forms.TabPage
        Me.ugTechReports = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.lblTechnicalReports = New System.Windows.Forms.Label
        Me.txtCommitmentComment = New System.Windows.Forms.TextBox
        Me.lblCommitmentComment = New System.Windows.Forms.Label
        Me.txtComments = New System.Windows.Forms.TextBox
        Me.lblComment = New System.Windows.Forms.Label
        Me.lblDeduction = New System.Windows.Forms.Label
        Me.txtDeductionReason = New System.Windows.Forms.TextBox
        Me.lblDeductionReason = New System.Windows.Forms.Label
        Me.btnInsert = New System.Windows.Forms.Button
        Me.cmbDeduction = New System.Windows.Forms.ComboBox
        Me.txtPaid = New System.Windows.Forms.TextBox
        Me.lblPaid = New System.Windows.Forms.Label
        Me.txtInvoiced = New System.Windows.Forms.TextBox
        Me.lblInvoiced = New System.Windows.Forms.Label
        Me.chkFinal = New System.Windows.Forms.CheckBox
        Me.chkOnHold = New System.Windows.Forms.CheckBox
        Me.tbPageAdjustments = New System.Windows.Forms.TabPage
        Me.ugAdjustment = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.tbPageCommitment = New System.Windows.Forms.TabPage
        Me.cfCommitment = New MUSTER.CostFormat
        Me.tbPagePastInvoices = New System.Windows.Forms.TabPage
        Me.ugPastInvoices = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlInvoiceBottom.SuspendLayout()
        Me.pnlInvoiceTop.SuspendLayout()
        Me.pnlInvoiceDetails.SuspendLayout()
        Me.tbCtrlInvoice.SuspendLayout()
        Me.tbPageInvoice.SuspendLayout()
        CType(Me.ugTechReports, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbPageAdjustments.SuspendLayout()
        CType(Me.ugAdjustment, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbPageCommitment.SuspendLayout()
        Me.tbPagePastInvoices.SuspendLayout()
        CType(Me.ugPastInvoices, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pnlInvoiceBottom
        '
        Me.pnlInvoiceBottom.Controls.Add(Me.btnCancel)
        Me.pnlInvoiceBottom.Controls.Add(Me.btnOK)
        Me.pnlInvoiceBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlInvoiceBottom.Location = New System.Drawing.Point(0, 526)
        Me.pnlInvoiceBottom.Name = "pnlInvoiceBottom"
        Me.pnlInvoiceBottom.Size = New System.Drawing.Size(656, 40)
        Me.pnlInvoiceBottom.TabIndex = 30
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(328, 8)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 9
        Me.btnCancel.Text = "Cancel"
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(248, 8)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.TabIndex = 8
        Me.btnOK.Text = "OK"
        '
        'pnlInvoiceTop
        '
        Me.pnlInvoiceTop.Controls.Add(Me.txtInvoiceNumber)
        Me.pnlInvoiceTop.Controls.Add(Me.lblInvoiceNumber)
        Me.pnlInvoiceTop.Controls.Add(Me.lblPO)
        Me.pnlInvoiceTop.Controls.Add(Me.lblPOValue)
        Me.pnlInvoiceTop.Controls.Add(Me.cmbCommitment)
        Me.pnlInvoiceTop.Controls.Add(Me.lblCommitment)
        Me.pnlInvoiceTop.Controls.Add(Me.lblReimbursementRequest)
        Me.pnlInvoiceTop.Controls.Add(Me.lblIDValue)
        Me.pnlInvoiceTop.Controls.Add(Me.lblID)
        Me.pnlInvoiceTop.Controls.Add(Me.cmbReimbursementRequest)
        Me.pnlInvoiceTop.Controls.Add(Me.txtReimbursementRequest)
        Me.pnlInvoiceTop.Controls.Add(Me.txtCommitment)
        Me.pnlInvoiceTop.Controls.Add(Me.dtPickPaymentDate)
        Me.pnlInvoiceTop.Controls.Add(Me.lblPaymentDate)
        Me.pnlInvoiceTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlInvoiceTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlInvoiceTop.Name = "pnlInvoiceTop"
        Me.pnlInvoiceTop.Size = New System.Drawing.Size(656, 136)
        Me.pnlInvoiceTop.TabIndex = 99
        '
        'txtInvoiceNumber
        '
        Me.txtInvoiceNumber.Location = New System.Drawing.Point(192, 112)
        Me.txtInvoiceNumber.Name = "txtInvoiceNumber"
        Me.txtInvoiceNumber.TabIndex = 0
        Me.txtInvoiceNumber.Text = ""
        '
        'lblInvoiceNumber
        '
        Me.lblInvoiceNumber.Location = New System.Drawing.Point(96, 112)
        Me.lblInvoiceNumber.Name = "lblInvoiceNumber"
        Me.lblInvoiceNumber.Size = New System.Drawing.Size(87, 17)
        Me.lblInvoiceNumber.TabIndex = 89
        Me.lblInvoiceNumber.Text = "Invoice Number:"
        '
        'lblPO
        '
        Me.lblPO.Location = New System.Drawing.Point(152, 87)
        Me.lblPO.Name = "lblPO"
        Me.lblPO.Size = New System.Drawing.Size(32, 17)
        Me.lblPO.TabIndex = 99
        Me.lblPO.Text = "PO#:"
        '
        'lblPOValue
        '
        Me.lblPOValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblPOValue.Location = New System.Drawing.Point(192, 85)
        Me.lblPOValue.Name = "lblPOValue"
        Me.lblPOValue.Size = New System.Drawing.Size(100, 19)
        Me.lblPOValue.TabIndex = 10
        '
        'cmbCommitment
        '
        Me.cmbCommitment.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbCommitment.Location = New System.Drawing.Point(192, 56)
        Me.cmbCommitment.Name = "cmbCommitment"
        Me.cmbCommitment.Size = New System.Drawing.Size(408, 21)
        Me.cmbCommitment.TabIndex = 7
        Me.cmbCommitment.TabStop = False
        '
        'lblCommitment
        '
        Me.lblCommitment.Location = New System.Drawing.Point(116, 56)
        Me.lblCommitment.Name = "lblCommitment"
        Me.lblCommitment.Size = New System.Drawing.Size(72, 23)
        Me.lblCommitment.TabIndex = 99
        Me.lblCommitment.Text = "Commitment:"
        '
        'lblReimbursementRequest
        '
        Me.lblReimbursementRequest.Location = New System.Drawing.Point(55, 32)
        Me.lblReimbursementRequest.Name = "lblReimbursementRequest"
        Me.lblReimbursementRequest.Size = New System.Drawing.Size(133, 17)
        Me.lblReimbursementRequest.TabIndex = 99
        Me.lblReimbursementRequest.Text = "Reimbursement Request:"
        '
        'lblIDValue
        '
        Me.lblIDValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblIDValue.Location = New System.Drawing.Point(192, 7)
        Me.lblIDValue.Name = "lblIDValue"
        Me.lblIDValue.Size = New System.Drawing.Size(100, 19)
        Me.lblIDValue.TabIndex = 6
        '
        'lblID
        '
        Me.lblID.Location = New System.Drawing.Point(165, 8)
        Me.lblID.Name = "lblID"
        Me.lblID.Size = New System.Drawing.Size(22, 17)
        Me.lblID.TabIndex = 120
        Me.lblID.Text = "ID:"
        '
        'cmbReimbursementRequest
        '
        Me.cmbReimbursementRequest.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbReimbursementRequest.Location = New System.Drawing.Point(192, 32)
        Me.cmbReimbursementRequest.Name = "cmbReimbursementRequest"
        Me.cmbReimbursementRequest.Size = New System.Drawing.Size(168, 21)
        Me.cmbReimbursementRequest.TabIndex = 6
        Me.cmbReimbursementRequest.TabStop = False
        '
        'txtReimbursementRequest
        '
        Me.txtReimbursementRequest.Location = New System.Drawing.Point(192, 32)
        Me.txtReimbursementRequest.Name = "txtReimbursementRequest"
        Me.txtReimbursementRequest.Size = New System.Drawing.Size(168, 20)
        Me.txtReimbursementRequest.TabIndex = 9
        Me.txtReimbursementRequest.TabStop = False
        Me.txtReimbursementRequest.Text = ""
        '
        'txtCommitment
        '
        Me.txtCommitment.Location = New System.Drawing.Point(192, 56)
        Me.txtCommitment.Name = "txtCommitment"
        Me.txtCommitment.Size = New System.Drawing.Size(408, 20)
        Me.txtCommitment.TabIndex = 10
        Me.txtCommitment.Text = ""
        '
        'dtPickPaymentDate
        '
        Me.dtPickPaymentDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickPaymentDate.Location = New System.Drawing.Point(504, 112)
        Me.dtPickPaymentDate.Name = "dtPickPaymentDate"
        Me.dtPickPaymentDate.Size = New System.Drawing.Size(96, 20)
        Me.dtPickPaymentDate.TabIndex = 9
        Me.dtPickPaymentDate.TabStop = False
        '
        'lblPaymentDate
        '
        Me.lblPaymentDate.Location = New System.Drawing.Point(424, 112)
        Me.lblPaymentDate.Name = "lblPaymentDate"
        Me.lblPaymentDate.Size = New System.Drawing.Size(76, 17)
        Me.lblPaymentDate.TabIndex = 12
        Me.lblPaymentDate.Text = "PaymentDate:"
        '
        'pnlInvoiceDetails
        '
        Me.pnlInvoiceDetails.Controls.Add(Me.tbCtrlInvoice)
        Me.pnlInvoiceDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlInvoiceDetails.Location = New System.Drawing.Point(0, 136)
        Me.pnlInvoiceDetails.Name = "pnlInvoiceDetails"
        Me.pnlInvoiceDetails.Size = New System.Drawing.Size(656, 390)
        Me.pnlInvoiceDetails.TabIndex = 3
        '
        'tbCtrlInvoice
        '
        Me.tbCtrlInvoice.Controls.Add(Me.tbPageInvoice)
        Me.tbCtrlInvoice.Controls.Add(Me.tbPageAdjustments)
        Me.tbCtrlInvoice.Controls.Add(Me.tbPageCommitment)
        Me.tbCtrlInvoice.Controls.Add(Me.tbPagePastInvoices)
        Me.tbCtrlInvoice.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbCtrlInvoice.Location = New System.Drawing.Point(0, 0)
        Me.tbCtrlInvoice.Name = "tbCtrlInvoice"
        Me.tbCtrlInvoice.SelectedIndex = 0
        Me.tbCtrlInvoice.Size = New System.Drawing.Size(656, 390)
        Me.tbCtrlInvoice.TabIndex = 1
        '
        'tbPageInvoice
        '
        Me.tbPageInvoice.AutoScroll = True
        Me.tbPageInvoice.Controls.Add(Me.ugTechReports)
        Me.tbPageInvoice.Controls.Add(Me.lblTechnicalReports)
        Me.tbPageInvoice.Controls.Add(Me.txtCommitmentComment)
        Me.tbPageInvoice.Controls.Add(Me.lblCommitmentComment)
        Me.tbPageInvoice.Controls.Add(Me.txtComments)
        Me.tbPageInvoice.Controls.Add(Me.lblComment)
        Me.tbPageInvoice.Controls.Add(Me.lblDeduction)
        Me.tbPageInvoice.Controls.Add(Me.txtDeductionReason)
        Me.tbPageInvoice.Controls.Add(Me.lblDeductionReason)
        Me.tbPageInvoice.Controls.Add(Me.btnInsert)
        Me.tbPageInvoice.Controls.Add(Me.cmbDeduction)
        Me.tbPageInvoice.Controls.Add(Me.txtPaid)
        Me.tbPageInvoice.Controls.Add(Me.lblPaid)
        Me.tbPageInvoice.Controls.Add(Me.txtInvoiced)
        Me.tbPageInvoice.Controls.Add(Me.lblInvoiced)
        Me.tbPageInvoice.Controls.Add(Me.chkFinal)
        Me.tbPageInvoice.Controls.Add(Me.chkOnHold)
        Me.tbPageInvoice.Location = New System.Drawing.Point(4, 22)
        Me.tbPageInvoice.Name = "tbPageInvoice"
        Me.tbPageInvoice.Size = New System.Drawing.Size(648, 364)
        Me.tbPageInvoice.TabIndex = 0
        Me.tbPageInvoice.Text = "Invoice"
        '
        'ugTechReports
        '
        Me.ugTechReports.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugTechReports.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True
        Me.ugTechReports.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle
        Me.ugTechReports.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugTechReports.Location = New System.Drawing.Point(16, 320)
        Me.ugTechReports.Name = "ugTechReports"
        Me.ugTechReports.Size = New System.Drawing.Size(608, 112)
        Me.ugTechReports.TabIndex = 6
        '
        'lblTechnicalReports
        '
        Me.lblTechnicalReports.Location = New System.Drawing.Point(16, 301)
        Me.lblTechnicalReports.Name = "lblTechnicalReports"
        Me.lblTechnicalReports.Size = New System.Drawing.Size(100, 17)
        Me.lblTechnicalReports.TabIndex = 23
        Me.lblTechnicalReports.Text = "Technical Reports"
        '
        'txtCommitmentComment
        '
        Me.txtCommitmentComment.Location = New System.Drawing.Point(16, 240)
        Me.txtCommitmentComment.Multiline = True
        Me.txtCommitmentComment.Name = "txtCommitmentComment"
        Me.txtCommitmentComment.ReadOnly = True
        Me.txtCommitmentComment.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtCommitmentComment.Size = New System.Drawing.Size(608, 56)
        Me.txtCommitmentComment.TabIndex = 22
        Me.txtCommitmentComment.TabStop = False
        Me.txtCommitmentComment.Text = ""
        '
        'lblCommitmentComment
        '
        Me.lblCommitmentComment.Location = New System.Drawing.Point(16, 224)
        Me.lblCommitmentComment.Name = "lblCommitmentComment"
        Me.lblCommitmentComment.Size = New System.Drawing.Size(128, 17)
        Me.lblCommitmentComment.TabIndex = 21
        Me.lblCommitmentComment.Text = "Commitment Comment"
        '
        'txtComments
        '
        Me.txtComments.Location = New System.Drawing.Point(16, 168)
        Me.txtComments.Multiline = True
        Me.txtComments.Name = "txtComments"
        Me.txtComments.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtComments.Size = New System.Drawing.Size(608, 56)
        Me.txtComments.TabIndex = 5
        Me.txtComments.Text = ""
        '
        'lblComment
        '
        Me.lblComment.Location = New System.Drawing.Point(16, 152)
        Me.lblComment.Name = "lblComment"
        Me.lblComment.Size = New System.Drawing.Size(56, 17)
        Me.lblComment.TabIndex = 20
        Me.lblComment.Text = "Comment"
        '
        'lblDeduction
        '
        Me.lblDeduction.Location = New System.Drawing.Point(298, 64)
        Me.lblDeduction.Name = "lblDeduction"
        Me.lblDeduction.Size = New System.Drawing.Size(58, 17)
        Me.lblDeduction.TabIndex = 10
        Me.lblDeduction.Text = "Deduction:"
        '
        'txtDeductionReason
        '
        Me.txtDeductionReason.Location = New System.Drawing.Point(16, 88)
        Me.txtDeductionReason.Multiline = True
        Me.txtDeductionReason.Name = "txtDeductionReason"
        Me.txtDeductionReason.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtDeductionReason.Size = New System.Drawing.Size(608, 56)
        Me.txtDeductionReason.TabIndex = 4
        Me.txtDeductionReason.Text = ""
        '
        'lblDeductionReason
        '
        Me.lblDeductionReason.Location = New System.Drawing.Point(16, 64)
        Me.lblDeductionReason.Name = "lblDeductionReason"
        Me.lblDeductionReason.Size = New System.Drawing.Size(100, 17)
        Me.lblDeductionReason.TabIndex = 18
        Me.lblDeductionReason.Text = "Deduction Reason"
        '
        'btnInsert
        '
        Me.btnInsert.Location = New System.Drawing.Point(582, 64)
        Me.btnInsert.Name = "btnInsert"
        Me.btnInsert.Size = New System.Drawing.Size(43, 23)
        Me.btnInsert.TabIndex = 19
        Me.btnInsert.TabStop = False
        Me.btnInsert.Text = "Insert"
        '
        'cmbDeduction
        '
        Me.cmbDeduction.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbDeduction.Location = New System.Drawing.Point(365, 64)
        Me.cmbDeduction.Name = "cmbDeduction"
        Me.cmbDeduction.Size = New System.Drawing.Size(216, 21)
        Me.cmbDeduction.TabIndex = 7
        '
        'txtPaid
        '
        Me.txtPaid.Location = New System.Drawing.Point(408, 34)
        Me.txtPaid.Name = "txtPaid"
        Me.txtPaid.TabIndex = 3
        Me.txtPaid.Text = ""
        Me.txtPaid.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblPaid
        '
        Me.lblPaid.Location = New System.Drawing.Point(376, 34)
        Me.lblPaid.Name = "lblPaid"
        Me.lblPaid.Size = New System.Drawing.Size(32, 17)
        Me.lblPaid.TabIndex = 17
        Me.lblPaid.Text = "Paid:"
        '
        'txtInvoiced
        '
        Me.txtInvoiced.Location = New System.Drawing.Point(208, 34)
        Me.txtInvoiced.Name = "txtInvoiced"
        Me.txtInvoiced.TabIndex = 2
        Me.txtInvoiced.Text = ""
        Me.txtInvoiced.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblInvoiced
        '
        Me.lblInvoiced.Location = New System.Drawing.Point(152, 34)
        Me.lblInvoiced.Name = "lblInvoiced"
        Me.lblInvoiced.Size = New System.Drawing.Size(50, 17)
        Me.lblInvoiced.TabIndex = 16
        Me.lblInvoiced.Text = "Invoiced:"
        '
        'chkFinal
        '
        Me.chkFinal.Location = New System.Drawing.Point(376, 8)
        Me.chkFinal.Name = "chkFinal"
        Me.chkFinal.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkFinal.Size = New System.Drawing.Size(48, 24)
        Me.chkFinal.TabIndex = 15
        Me.chkFinal.TabStop = False
        Me.chkFinal.Text = "Final"
        '
        'chkOnHold
        '
        Me.chkOnHold.Location = New System.Drawing.Point(152, 8)
        Me.chkOnHold.Name = "chkOnHold"
        Me.chkOnHold.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkOnHold.Size = New System.Drawing.Size(72, 24)
        Me.chkOnHold.TabIndex = 14
        Me.chkOnHold.TabStop = False
        Me.chkOnHold.Text = "On Hold"
        '
        'tbPageAdjustments
        '
        Me.tbPageAdjustments.Controls.Add(Me.ugAdjustment)
        Me.tbPageAdjustments.Location = New System.Drawing.Point(4, 22)
        Me.tbPageAdjustments.Name = "tbPageAdjustments"
        Me.tbPageAdjustments.Size = New System.Drawing.Size(648, 364)
        Me.tbPageAdjustments.TabIndex = 2
        Me.tbPageAdjustments.Text = "Adjustments"
        '
        'ugAdjustment
        '
        Me.ugAdjustment.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugAdjustment.Location = New System.Drawing.Point(16, 8)
        Me.ugAdjustment.Name = "ugAdjustment"
        Me.ugAdjustment.Size = New System.Drawing.Size(616, 264)
        Me.ugAdjustment.TabIndex = 0
        '
        'tbPageCommitment
        '
        Me.tbPageCommitment.Controls.Add(Me.cfCommitment)
        Me.tbPageCommitment.Location = New System.Drawing.Point(4, 22)
        Me.tbPageCommitment.Name = "tbPageCommitment"
        Me.tbPageCommitment.Size = New System.Drawing.Size(648, 364)
        Me.tbPageCommitment.TabIndex = 100
        Me.tbPageCommitment.Text = "Commitment"
        '
        'cfCommitment
        '
        Me.cfCommitment.Location = New System.Drawing.Point(24, 24)
        Me.cfCommitment.Name = "cfCommitment"
        Me.cfCommitment.Size = New System.Drawing.Size(600, 232)
        Me.cfCommitment.TabIndex = 111
        '
        'tbPagePastInvoices
        '
        Me.tbPagePastInvoices.Controls.Add(Me.ugPastInvoices)
        Me.tbPagePastInvoices.Location = New System.Drawing.Point(4, 22)
        Me.tbPagePastInvoices.Name = "tbPagePastInvoices"
        Me.tbPagePastInvoices.Size = New System.Drawing.Size(648, 364)
        Me.tbPagePastInvoices.TabIndex = 3
        Me.tbPagePastInvoices.Text = "Past Invoices"
        '
        'ugPastInvoices
        '
        Me.ugPastInvoices.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugPastInvoices.Location = New System.Drawing.Point(16, 16)
        Me.ugPastInvoices.Name = "ugPastInvoices"
        Me.ugPastInvoices.Size = New System.Drawing.Size(616, 264)
        Me.ugPastInvoices.TabIndex = 0
        '
        'Invoice
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(656, 566)
        Me.Controls.Add(Me.pnlInvoiceDetails)
        Me.Controls.Add(Me.pnlInvoiceTop)
        Me.Controls.Add(Me.pnlInvoiceBottom)
        Me.Name = "Invoice"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Invoice"
        Me.pnlInvoiceBottom.ResumeLayout(False)
        Me.pnlInvoiceTop.ResumeLayout(False)
        Me.pnlInvoiceDetails.ResumeLayout(False)
        Me.tbCtrlInvoice.ResumeLayout(False)
        Me.tbPageInvoice.ResumeLayout(False)
        CType(Me.ugTechReports, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbPageAdjustments.ResumeLayout(False)
        CType(Me.ugAdjustment, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbPageCommitment.ResumeLayout(False)
        Me.tbPagePastInvoices.ResumeLayout(False)
        CType(Me.ugPastInvoices, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Page Events"
    Private Sub Invoice_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try

            '---------- retrieve data for this invoice / reimbursement / commitment / --------- 
            '          if this is an add then reimbursement and commitment will not be there yet. 
            '          if this is a modify then the reimbursement and commitment records will be retrieved. 
            oFinancialEvent.Retrieve(FinancialEventID)
            oFinancialInvoice.Retrieve(FinancialInvoiceID)
            oFinancialReimbursement.Retrieve(oFinancialInvoice.ReimbursementID)
            oFinancialCommitment.Retrieve(oFinancialReimbursement.CommitmentID)

            sStartingPaid = 0
            sStartingRequested = 0
            If FinancialInvoiceID = 0 Then      '---- new invoice 
                lblIDValue.Text = "New"
            Else                                '---- modify invoice 
                lblIDValue.Text = FinancialInvoiceID
                sStartingPaid = oFinancialInvoice.PaidAmount
                sStartingRequested = oFinancialInvoice.InvoicedAmount
            End If
            LoadDropDowns()
            bolLoading = True
            If FinancialInvoiceID > 0 Then
                cmbReimbursementRequest.SelectedValue = oFinancialInvoice.ReimbursementID
                txtDeductionReason.Text = oFinancialInvoice.DeductionReason
                txtComments.Text = oFinancialInvoice.Comment
                txtPaid.Text = oFinancialInvoice.PaidAmount
                txtInvoiced.Text = oFinancialInvoice.InvoicedAmount
                txtInvoiceNumber.Text = oFinancialInvoice.VendorInvoice
                chkFinal.Checked = oFinancialInvoice.Final
                bolFinalOnLoad = oFinancialInvoice.Final
                chkOnHold.Checked = oFinancialInvoice.OnHold
                UIUtilsGen.SetDatePickerValue(dtPickPaymentDate, oFinancialReimbursement.PaymentDate)
            Else
                'cmbReimbursementRequest.SelectedValue = oFinancialInvoice.ReimbursementID
                txtDeductionReason.Text = oFinancialInvoice.DeductionReason
                txtComments.Text = oFinancialInvoice.Comment
                txtPaid.Text = oFinancialInvoice.PaidAmount
                txtInvoiced.Text = oFinancialInvoice.InvoicedAmount
                txtInvoiceNumber.Text = oFinancialInvoice.VendorInvoice
                chkFinal.Checked = oFinancialInvoice.Final
                bolFinalOnLoad = oFinancialInvoice.Final
                chkOnHold.Checked = oFinancialInvoice.OnHold
                oFinancialReimbursement.PaymentDate = Today.Date
                UIUtilsGen.SetDatePickerValue(dtPickPaymentDate, oFinancialReimbursement.PaymentDate)
            End If
            bolLoading = False

            If SelectedRequestId <= 0 Then
                If FinancialInvoiceID <= 0 Then
                    If cmbReimbursementRequest.Items.Count = 1 Then
                        cmbReimbursementRequest.SelectedIndex = 0
                    Else
                        cmbReimbursementRequest.SelectedIndex = -1
                    End If
                End If
            Else
                cmbReimbursementRequest.SelectedValue = SelectedRequestId
            End If

            If Me.cmbCommitment.Items Is Nothing OrElse Me.cmbCommitment.Items.Count = 0 Then
                MsgBox("There is no amount left to make a payment or the event is already closed.")
                Me.Close()
            End If

            ProcessReimbursementRequestChange()
            ProcessCommitmentChange()

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Error Loading Invoice Screen. " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()

        End Try


    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click
        Dim frmAdjustment As Adjustment
        Dim nBalance As Double
        Dim strMsgValidation As String = String.Empty
        Dim nAdjustmentID As Integer = -1
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim nRequestedAmount As Double = 0.0
        Dim nPaidAmount As Double = 0.0
        Dim nReimPaidAmnt As Double = 0.0   '--- total amount invoices/paid for a reimbursement 
        Dim nTempReimbursementID As Integer = oFinancialReimbursement.id

        Try
            '------ data entry validations ---------------------------------------------------
            If txtInvoiced.Text = "" Or txtInvoiced.Text = "0" Then
                strMsgValidation += "Invoice Amount Required. " + vbCrLf
            End If
            If txtInvoiceNumber.Text = "" Then
                strMsgValidation += "Invoice Number Required. " + vbCrLf
            End If
            If cmbReimbursementRequest.Items.Count = 0 Then
                strMsgValidation += "Reimbursement Request Required. " + vbCrLf
            End If
            If Me.cmbCommitment.Items.Count = 0 Then
                strMsgValidation += "Commitment Required. " + vbCrLf
            End If
            If txtPaid.Text = "" Then
                oFinancialInvoice.PaidAmount = 0
                txtPaid.Text = "0"
                ' check if change order exists that requires approval and is not approved 
            ElseIf oFinancialInvoice.CommitmentHasOpenChangeOrder(oFinancialCommitment.CommitmentID) Then
                strMsgValidation += "Commitment has an Unapproved Change Order requiring approval. " + vbCrLf
                strMsgValidation += "Please resolve and re-enter Invoice Payment later. " & vbCrLf
            End If

            If txtPaid.Text <> String.Empty And txtInvoiced.Text <> String.Empty Then
                If NegativeRequest Then
                    If CDbl(txtPaid.Text) > 0 Then
                        strMsgValidation += "Payment Amount has to be negative. " + vbCrLf
                    Else
                        If CDbl(txtPaid.Text) < CDbl(txtInvoiced.Text) Then
                            strMsgValidation += "Payment Amount cannot be less than Invoice Amount. " + vbCrLf
                            'Exit Sub
                        End If
                    End If
                Else
                    ' If CDbl(txtPaid.Text) < 0 Then
                    ' strMsgValidation += "Payment Amount has to be positive. " + vbCrLf
                        ' Else
                        If CDbl(txtPaid.Text) > CDbl(txtInvoiced.Text) Then
                            strMsgValidation += "Payment Amount cannot be more than Invoice Amount. " + vbCrLf
                            'Exit Sub
                        End If
                        ' End If
                    End If
            End If
            If cmbReimbursementRequest.Text <> String.Empty And txtPaid.Text <> String.Empty Then

                '---------------------------------------------------------------------------------------
                '    sum all past invoices for this reimbursement id only for the comparison to see if 
                '    the total of the paid invoices is greater than the reimbursement amount 
                '    the grid from which the invoice paid amounts are accumulated is the total list of 
                '    invoices for the commitment ... just pick out the invoice payments for this reimbursement id
                If ugPastInvoices.Rows.Count > 0 Then
                    For Each ugrow In ugPastInvoices.Rows
                        If nTempReimbursementID = Integer.Parse(ugrow.Cells("REIMBURSEMENT_ID").Value) Then
                            nReimPaidAmnt += CDbl(ugrow.Cells("Paid").Value)
                        End If
                    Next
                End If
                '------ if this is a modify invoice ... then subtract the starting amount of the invoice -------
                If FinancialInvoiceID > 0 Then
                    nReimPaidAmnt -= sStartingPaid
                End If
                '-----------------------------------------------------------------------------------------------

                If NegativeRequest Then
                    If (oFinancialReimbursement.RequestedAmount) > FormatNumber(nReimPaidAmnt + CDbl(txtPaid.Text), 2, TriState.False, TriState.False, TriState.True) Then
                        strMsgValidation += "Payment cannot be less than Requested Amount." + vbCrLf
                    End If
                Else
                    If (oFinancialReimbursement.RequestedAmount) < FormatNumber(nReimPaidAmnt + CDbl(txtPaid.Text), 2, TriState.False, TriState.False, TriState.True) Then
                        strMsgValidation += "Payment cannot be more than Requested Amount." + vbCrLf
                    End If
                End If
            End If

            If strMsgValidation <> String.Empty Then
                MsgBox(strMsgValidation)
                Exit Sub
            End If
            '---------- end of data entry validations --------------------------------------------------------

            '----------- if this is a modify invoice, check to see if it began the modify process as a final invoice -----
            '     if this was a final invoice, display a message prompt 
            If FinancialInvoiceID > 0 Then
                If bolFinalOnLoad Then
                    If chkFinal.Checked = False Then
                        If MsgBox("Are you Sure you want to Modify this Previously Final Invoice?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                        Else
                            Exit Sub
                        End If
                    End If
                    If chkFinal.Checked = True Then
                        If MsgBox("Are you Sure you want to Modify this Final Invoice?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                        Else
                            Exit Sub
                        End If
                    End If
                End If
            End If
            '------------------------------------------------------------------------------------------------------

            If oFinancialInvoice.id <= 0 Then
                oFinancialInvoice.CreatedBy = MusterContainer.AppUser.ID
            Else
                oFinancialInvoice.ModifiedBy = MusterContainer.AppUser.ID
            End If

            nBalance = oFinancialEvent.GetCommitmentBalance(oFinancialCommitment.CommitmentID)
            If txtPaid.Text <> String.Empty Then
                nBalance = FormatNumber((nBalance + sStartingPaid) - CDbl(txtPaid.Text), 2, TriState.False, TriState.False, TriState.True)
                If nBalance = 0.0 Then
                    oFinancialInvoice.Final = True
                End If
            Else
                If nBalance = 0.0 Then
                    oFinancialInvoice.Final = True
                End If
            End If

            If nBalance <> 0 Then
                If (nBalance < 0 And Not NegativeRequest) Then 'Or (nBalance > 0 And NegativeRequest) Then
                    MsgBox("Payment Creates An Overpayment. Loading Adjustment Interface Now.")
                    If IsNothing(frmAdjustment) Then

                        frmAdjustment = New Adjustment(nContactID)
                        frmAdjustment.FinancialEventID = oFinancialEvent.ID
                        frmAdjustment.FinancialCommitmentID = oFinancialCommitment.CommitmentID
                        frmAdjustment.AdjustmentID = 0
                        frmAdjustment.Balance = nBalance
                        frmAdjustment.NegativeRequest = NegativeRequest
                        frmAdjustment.SystemComment = "System Generated Change Order"
                        frmAdjustment.ShowDialog()
                        nAdjustmentID = frmAdjustment.nAdjustmentID
                    End If
                End If
            End If

            If chkFinal.Checked Then
                'Check for Over/Under Payment
                If nBalance <> 0 Then
                    If (nBalance > 0 And Not NegativeRequest) Then 'Or (nBalance < 0 And NegativeRequest) Then
                        MsgBox("Payment Creates An Underpayment. Loading Adjustment Interface Now.")
                        'End If
                        If IsNothing(frmAdjustment) Then
                            frmAdjustment = New Adjustment(nContactID)
                            frmAdjustment.FinancialEventID = oFinancialEvent.ID
                            frmAdjustment.FinancialCommitmentID = oFinancialCommitment.CommitmentID
                            frmAdjustment.AdjustmentID = 0
                            frmAdjustment.Balance = nBalance
                            frmAdjustment.NegativeRequest = NegativeRequest
                            frmAdjustment.ShowDialog()
                            nAdjustmentID = frmAdjustment.nAdjustmentID
                        End If
                    End If
                End If
            End If

            If nAdjustmentID = 0 Then
                If chkFinal.Checked Then
                    If MsgBox("Change order Adjustment was not done. Are you sure to save the Invoice?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                        Exit Sub
                    End If
                Else
                    MsgBox("Cannot save invoice, Change order Adjustment was not done")
                    Exit Sub
                End If
            End If

            'If nAdjustmentID > 0 Or nAdjustmentID < 0 Then
            If oFinancialReimbursement.CommitmentID = 0 Then
                If oFinancialCommitment.CommitmentID > 0 Then
                    oFinancialReimbursement.CommitmentID = oFinancialCommitment.CommitmentID
                Else
                    ProcessCommitmentChange()
                    If oFinancialCommitment.CommitmentID > 0 Then
                        oFinancialReimbursement.CommitmentID = oFinancialCommitment.CommitmentID
                    Else
                        MsgBox("Cannot save invoice, Commitment ID cannot be set")
                        Exit Sub
                    End If
                End If

                If oFinancialReimbursement.id <= 0 Then
                    oFinancialReimbursement.CreatedBy = MusterContainer.AppUser.ID
                Else
                    oFinancialReimbursement.ModifiedBy = MusterContainer.AppUser.ID
                End If
            End If
            oFinancialInvoice.PONumber = lblPOValue.Text.ToString()

            oFinancialInvoice.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If
            oFinancialReimbursement.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            'Else        'If nAdjustmentID = 0 Then
            '    MsgBox("Cannot save invoice, Change order Adjustment was not done")
            '    Exit Sub
            'End If

            If ProcessTechReports() Then

                MsgBox("Invoice Changes Saved")
                Me.Close()
            End If

            Me.Close()
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Save Changes " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        oFinancialInvoice.Reset()
        oFinancialReimbursement.Reset()

        Me.Close()

    End Sub

    Private Sub cmbReimbursementRequest_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbReimbursementRequest.SelectedIndexChanged
        If bolLoading Then Exit Sub

        ProcessReimbursementRequestChange()

    End Sub


    Private Sub cmbCommitment_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCommitment.SelectedIndexChanged
        If bolLoading Then Exit Sub

        ProcessCommitmentChange()
    End Sub

    Private Sub txtInvoiceNumber_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtInvoiceNumber.TextChanged
        If bolLoading = True Then Exit Sub
        oFinancialInvoice.VendorInvoice = txtInvoiceNumber.Text
    End Sub

    Private Sub dtPickPaymentDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtPickPaymentDate.ValueChanged
        If bolLoading = True Then Exit Sub
        UIUtilsGen.ToggleDateFormat(dtPickPaymentDate)
        oFinancialReimbursement.PaymentDate = UIUtilsGen.GetDatePickerValue(dtPickPaymentDate)
        oFinancialReimbursement.PaymentDate = oFinancialReimbursement.PaymentDate.Date
    End Sub

    Private Sub txtInvoiced_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtInvoiced.TextChanged
        If bolLoading = True Or bolFormatting = True Then Exit Sub

        If Not IsNumeric(txtInvoiced.Text) Then
            txtInvoiced.Text = 0
        End If

        oFinancialInvoice.InvoicedAmount = txtInvoiced.Text

        'If IsNumeric(txtInvoiced.Text) Then
        '    oFinancialInvoice.InvoicedAmount = txtInvoiced.Text
        'Else
        '    If txtInvoiced.Text = String.Empty Then
        '        oFinancialInvoice.InvoicedAmount = 0
        '    Else
        '        oFinancialInvoice.InvoicedAmount = 0
        '        txtInvoiced.Text = 0
        '        txtInvoiced.Focus()
        '        MsgBox("Invoiced Amount Must Be Numeric.")
        '    End If
        'End If

    End Sub

    Private Sub txtPaid_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPaid.TextChanged
        If bolLoading = True Or bolFormatting = True Then Exit Sub

        If Not IsNumeric(txtPaid.Text) Then
            txtPaid.Text = 0
        End If

        oFinancialInvoice.PaidAmount = txtPaid.Text

        'If IsNumeric(txtPaid.Text) Then
        '    oFinancialInvoice.PaidAmount = txtPaid.Text
        'Else
        '    If txtPaid.Text = String.Empty Then
        '        oFinancialInvoice.PaidAmount = 0
        '    Else
        '        oFinancialInvoice.PaidAmount = 0
        '        txtPaid.Text = 0
        '        txtPaid.Focus()
        '        MsgBox("Paid Amount Must Be Numeric.")
        '    End If
        'End If
    End Sub

    Private Sub chkOnHold_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkOnHold.CheckedChanged
        If bolLoading Then Exit Sub

        oFinancialInvoice.OnHold = chkOnHold.Checked

    End Sub

    Private Sub chkFinal_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkFinal.CheckedChanged
        If bolLoading Then Exit Sub

        oFinancialInvoice.Final = chkFinal.Checked

    End Sub


    Private Sub txtDeductionReason_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDeductionReason.TextChanged
        If bolLoading Then Exit Sub

        oFinancialInvoice.DeductionReason = txtDeductionReason.Text
    End Sub

    Private Sub txtComments_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtComments.TextChanged
        If bolLoading Then Exit Sub

        oFinancialInvoice.Comment = txtComments.Text
    End Sub

    Private Sub txtPaid_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPaid.LostFocus
        bolFormatting = True
        If txtPaid.Text = String.Empty Then
            txtPaid.Text = "0.00"
        ElseIf IsNumeric(txtPaid.Text) Then
            txtPaid.Text = FormatNumber(txtPaid.Text, 2, TriState.False, TriState.False, TriState.True)
        Else
            txtPaid.Text = "0.00"
            txtPaid.Focus()
            MsgBox("Paid Amount Must Be Numeric.")
        End If
        oFinancialInvoice.PaidAmount = txtPaid.Text
        bolFormatting = False
    End Sub

    Private Sub txtInvoiced_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtInvoiced.LostFocus
        bolFormatting = True
        If txtInvoiced.Text = String.Empty Then
            txtInvoiced.Text = "0.00"
        ElseIf IsNumeric(txtInvoiced.Text) Then
            txtInvoiced.Text = FormatNumber(txtInvoiced.Text, 2, TriState.False, TriState.False, TriState.True)
        Else
            txtInvoiced.Text = "0.00"
            txtInvoiced.Focus()
            MsgBox("Invoiced Amount Must Be Numeric.")
        End If
        oFinancialInvoice.InvoicedAmount = txtInvoiced.Text
        bolFormatting = False
    End Sub
#End Region

#Region "General Processes"
    Private Sub LoadDropDowns()
        Dim dtCommitment As DataTable
        Dim dtReimbursementRequest As DataTable

        Try
            bolLoading = True

            If FinancialInvoiceID > 0 Then
                dtCommitment = oFinancialEvent.PopulateInvoiceCommitmentList2(oFinancialCommitment.CommitmentID)
                dtReimbursementRequest = oFinancialEvent.PopulateInvoiceReimbursementRequestsList2(oFinancialReimbursement.id)
            Else
                dtCommitment = oFinancialEvent.PopulateInvoiceCommitmentList(FinancialEventID)
                dtReimbursementRequest = oFinancialEvent.PopulateInvoiceReimbursementRequestsList(FinancialEventID)
            End If
            Dim dtDeductions As DataTable = oFinancialEvent.PopulateInvoiceDeductionReasons


            cmbCommitment.DataSource = dtCommitment
            cmbCommitment.DisplayMember = "CommitmentDesc"
            cmbCommitment.ValueMember = "CommitmentID"
            Dim dRow As DataRow
            Dim isExists As Boolean = False
            If Not dtReimbursementRequest Is Nothing Then
                For Each dRow In dtReimbursementRequest.Rows
                    If SelectedRequestId = CInt(dRow("Reimbursement_ID")) Then
                        isExists = True
                    End If
                Next
            End If
            If Not isExists Then
                SelectedRequestId = 0
            End If

            cmbReimbursementRequest.DataSource = dtReimbursementRequest
            cmbReimbursementRequest.DisplayMember = "ReimbursementDesc"
            cmbReimbursementRequest.ValueMember = "Reimbursement_ID"

            cmbDeduction.DataSource = dtDeductions
            cmbDeduction.DisplayMember = "Text_Name"
            cmbDeduction.ValueMember = "Text_ID"

            'If FinancialInvoiceID = 0 Then
            '    If cmbCommitment.Items.Count > 0 Then
            '        ProcessCommitmentChange()
            '    End If
            'End If



            bolLoading = False
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Load DropDowns " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()

        End Try


    End Sub

    Private Sub ProcessReimbursementRequestChange()
        Try
            oFinancialReimbursement.Retrieve(cmbReimbursementRequest.SelectedValue)
            NegativeRequest = False
            If oFinancialReimbursement.RequestedAmount < 0 Then
                NegativeRequest = True
            End If
            oFinancialInvoice.ReimbursementID = cmbReimbursementRequest.SelectedValue
            If oFinancialReimbursement.CommitmentID > 0 Then
                cmbCommitment.SelectedValue = oFinancialReimbursement.CommitmentID
                txtCommitment.Text = cmbCommitment.Text
                txtCommitment.Visible = True
                If txtCommitment.Text = "" Then
                    txtCommitment.Text = "Commitment Associated With This Reimbursement Has 0.00 Balance"
                End If
                cmbCommitment.Visible = False
            Else
                If cmbCommitment.Items.Count > 0 Then
                    cmbCommitment.SelectedIndex = 0
                    oFinancialReimbursement.CommitmentID = cmbCommitment.SelectedValue
                End If
                txtCommitment.Text = ""
                txtCommitment.Visible = False
                cmbCommitment.Visible = True
            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot ProcessReimbursementRequestChange " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()

        End Try
    End Sub

    Private Sub ProcessCommitmentChange()

        Try

            If cmbCommitment.SelectedValue Is Nothing AndAlso cmbCommitment.Items.Count >= 1 Then
                cmbCommitment.SelectedItem = cmbCommitment.Items(0)
            End If

            If Not cmbCommitment.SelectedValue Is Nothing Then

                oFinancialCommitment.Retrieve(cmbCommitment.SelectedValue)
                If IsNothing(oFinancialCommitment.CommitmentID) = False Then
                    If oFinancialCommitment.CommitmentID > 0 Then
                        oFinancialReimbursement.CommitmentID = oFinancialCommitment.CommitmentID
                    End If
                End If
                lblPOValue.Text = oFinancialCommitment.PONumber
                txtCommitmentComment.Text = oFinancialCommitment.Comments

                ' Load CostFormatGrid
                cfCommitment.CostFormatType = oFinancialCommitment.Case_Letter
                cfCommitment.AssignCommitmentObject(oFinancialCommitment)
                cfCommitment.SetDisplay(False)
                cfCommitment.LoadCommitment()
                cfCommitment.SetReadonly(True)

                'Load Adjustments Grid
                LoadCommitmentAdjustmentsGrid()

                'Load Past Invoices Grid
                LoadCommitmentInvoicesGrid()

                'Load Tech Reports Grid
                LoadReportGrid()

                ' #927
                If lblPOValue.Text.Trim.Length > 0 Then
                    txtPaid.Enabled = True
                    chkFinal.Enabled = True
                Else
                    txtPaid.Enabled = False
                    chkFinal.Enabled = False
                End If
            End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot ProcessCommitmentChange " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()

        End Try
    End Sub

    Private Sub LoadCommitmentInvoicesGrid()
        Dim dsLocal As DataSet

        Try

            dsLocal = oFinancialEvent.CommitmentGridInvoicesOnlyDataset(oFinancialCommitment.CommitmentID)
            ugPastInvoices.DataSource = dsLocal
            ugPastInvoices.Rows.CollapseAll(True)
            ugPastInvoices.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
            ugPastInvoices.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
            If dsLocal.Tables(0).Rows.Count > 0 Then ' Does Commitment Table have rows

                ugPastInvoices.DisplayLayout.Bands(0).Columns("FIN_EVENT_ID").Hidden = True
                ugPastInvoices.DisplayLayout.Bands(0).Columns("CommitmentID").Hidden = True
                ugPastInvoices.DisplayLayout.Bands(0).Columns("ChildID").Hidden = True
                ugPastInvoices.DisplayLayout.Bands(0).Columns("REQUESTED_AMOUNT").Hidden = True
                ugPastInvoices.DisplayLayout.Bands(0).Columns("REIMBURSEMENT_ID").Hidden = True

                ugPastInvoices.DisplayLayout.Bands(0).Columns("Payment_Date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                ugPastInvoices.DisplayLayout.Bands(0).Columns("Payment_Date").Header.Caption = "Date"
                ugPastInvoices.DisplayLayout.Bands(0).Columns("Paid").Header.Caption = "Paid"
                ugPastInvoices.DisplayLayout.Bands(0).Columns("Final").Header.Caption = "Final"
                ugPastInvoices.DisplayLayout.Bands(0).Columns("PONumber").Header.Caption = "PO #"
                ugPastInvoices.DisplayLayout.Bands(0).Columns("Vendor_Inv_Number").Header.Caption = "Invoice #"

                'ugPastInvoices.DisplayLayout.Bands(tmpBand).Columns("Paid").ColSpan = 2
                ugPastInvoices.DisplayLayout.Bands(0).Columns("Paid").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                'ugPastInvoices.DisplayLayout.Bands(tmpBand).Columns("Comment").ColSpan = 4

            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Load CommitmentInvoicesGrid " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()

        End Try
    End Sub
    Private Sub LoadCommitmentAdjustmentsGrid()
        Dim dsLocal As DataSet

        Try

            dsLocal = oFinancialEvent.CommitmentGridAdjustmentsOnlyDataset(oFinancialCommitment.CommitmentID)
            ugAdjustment.DataSource = dsLocal

            ugAdjustment.Rows.CollapseAll(True)
            ugAdjustment.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
            ugAdjustment.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
            If dsLocal.Tables(0).Rows.Count > 0 Then ' Does Commitment Table have rows

                ugAdjustment.DisplayLayout.Bands(0).Columns("Fin_Event_ID").Hidden = True
                ugAdjustment.DisplayLayout.Bands(0).Columns("COMMITMENTID").Hidden = True
                ugAdjustment.DisplayLayout.Bands(0).Columns("ChildID").Hidden = True

                ugAdjustment.DisplayLayout.Bands(0).Columns("Adjust_Date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

                ugAdjustment.DisplayLayout.Bands(0).Columns("Adjust_Date").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right
                ugAdjustment.DisplayLayout.Bands(0).Columns("Adjust_Amount").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

                ugAdjustment.DisplayLayout.Bands(0).Columns("Adjust_Date").Header.Caption = "Date"
                ugAdjustment.DisplayLayout.Bands(0).Columns("Adjust_Type").Header.Caption = "Type"
                ugAdjustment.DisplayLayout.Bands(0).Columns("Adjust_Amount").Header.Caption = "Amount"
                ugAdjustment.DisplayLayout.Bands(0).Columns("Comments").Header.Caption = "Comments"
                ugAdjustment.DisplayLayout.Bands(0).Columns("Director_App_Req").Header.Caption = "Dir. App. Reqd."
                ugAdjustment.DisplayLayout.Bands(0).Columns("Fin_App_Req").Header.Caption = "Fin. App. Reqd."
                ugAdjustment.DisplayLayout.Bands(0).Columns("Approved").Header.Caption = "Approved"

                'ugAdjustments.DisplayLayout.Bands(tmpBand).Columns("Fin_App_Req").ColSpan = 2
                'ugAdjustments.DisplayLayout.Bands(tmpBand).Columns("Comments").ColSpan = 3

            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Load CommitmentAdjustmentsGrid " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()

        End Try
    End Sub

    Private Sub LoadReportGrid()
        Dim dsLocal As DataSet
        Dim tmpBand As Int16
        Dim dtTotals As DataTable

        dsLocal = oFinancialEvent.PopulateCommitmentTecDocList(oFinancialEvent.TecEventID, cmbCommitment.SelectedValue, Convert.ToInt32(oFinancialCommitment.ActivityType), "RPT")

        ugTechReports.DataSource = dsLocal
        ugTechReports.Rows.CollapseAll(True)
        ugTechReports.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        'ugTechReports.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect

        'If dsLocal.Tables(0).Rows.Count > 0 Then ' Does Payment Table have rows


        ugTechReports.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        ugTechReports.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True


        ugTechReports.DisplayLayout.Bands(0).Columns("Event_Activity_Document_ID").Hidden = True
        ugTechReports.DisplayLayout.Bands(0).Columns("Commitment_ID").Hidden = True
        ugTechReports.DisplayLayout.Bands(0).Columns("Due_Date").Header.Caption = "Due Date"
        ugTechReports.DisplayLayout.Bands(0).Columns("Extension_Date").Header.Caption = "Extension"
        ugTechReports.DisplayLayout.Bands(0).Columns("Received_Date").Header.Caption = "Received"
        ugTechReports.DisplayLayout.Bands(0).Columns("Date_Sent_To_Finance").Header.Caption = "Approved"



        ugTechReports.DisplayLayout.Bands(0).Columns("Date_Closed").CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit
        ugTechReports.DisplayLayout.Bands(0).Columns("Due_Date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

        ugTechReports.DisplayLayout.Bands(0).Columns("Extension_Date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        ugTechReports.DisplayLayout.Bands(0).Columns("Received_Date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        ugTechReports.DisplayLayout.Bands(0).Columns("Date_Sent_To_Finance").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        ugTechReports.DisplayLayout.Bands(0).Columns("Paid").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        ugTechReports.DisplayLayout.Bands(0).Columns("event_id").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        ugTechReports.DisplayLayout.Bands(0).Columns("Status").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        ugTechReports.DisplayLayout.Bands(0).Columns("Report").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit


        'ugTechReports.DisplayLayout.Bands(0).Columns("Paid").Hidden = True

        'ugTechReports.DisplayLayout.Bands(0).Columns("Received_date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        'ugTechReports.DisplayLayout.Bands(0).Columns("Date_Sent_To_Finance").Header.Caption = "Approved"
        'ugTechReports.DisplayLayout.Bands(0).Columns("Received_date").CellAppearance.TextHAlign = Infragistics.Win.HAlign.Right

        'ugTechReports.DisplayLayout.Bands(0).Columns("Received_date").Width = 100
        'ugTechReports.DisplayLayout.Bands(0).Columns("Requested_Amount").Width = 100
        'ugTechReports.DisplayLayout.Bands(0).Columns("Requested_Invoiced").Width = 100
        'ugTechReports.DisplayLayout.Bands(0).Columns("Paid").Width = 100


        'End If


    End Sub
#End Region

    Private Function ProcessTechReports() As Boolean
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim oTechDox As New MUSTER.BusinessLogic.pLustEventDocument
        Dim bolFinalNoClosedDate As Boolean = False
        Dim allowUpdate As Boolean = True
        Dim pass As Boolean = True


        For Each ugrow In ugTechReports.Rows

            'get current report data
            oTechDox.Retrieve(ugrow.Cells("Event_Activity_Document_ID").Value)
            oTechDox.Paid = ugrow.Cells("Paid").Value

            'If change found in closed_date
            If oTechDox.DocClosedDate <> IIf(IsDBNull(ugrow.Cells("Date_Closed").Value), Nothing, ugrow.Cells("Date_Closed").Value) Then


                allowUpdate = True



                'ask user if he/she should close the report
                If Not oFinancialInvoice.Final And allowUpdate Then

                    allowUpdate = (MsgBox(String.Format("This Invoice is not final! Do you still wish to change The Date_Closed on report {0}", _
                                         ugrow.Cells("Report").Value), MsgBoxStyle.YesNo, String.Format("Technical Reports Change on Invoice #{0}", _
                                                                                  oFinancialInvoice.id)) = MsgBoxResult.Yes)

                End If


                'If Final  or user opts to close anyway
                If allowUpdate Then
                    If ugrow.Cells("Date_Closed").Value Is DBNull.Value Then
                        oTechDox.DocClosedDate = CDate("01/01/0001")
                        bolFinalNoClosedDate = True
                    Else
                        oTechDox.DocClosedDate = ugrow.Cells("Date_Closed").Value
                    End If
                End If

                'If oFinancialInvoice.Final Then
                '    oTechDox.DocClosedDate = Today.Date
                'End If


                oTechDox.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)

                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Function
                End If

            End If

        Next

        If bolFinalNoClosedDate And pass Then
            MsgBox("One or more technical reports 'Date Closed' is not entered.")
        End If

        Return pass
    End Function

    Private Sub btnInsert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInsert.Click
        Dim oFinText As New MUSTER.BusinessLogic.pFinancialText

        oFinText.Retrieve(cmbDeduction.SelectedValue)
        If oFinText.ID > 0 Then
            txtDeductionReason.Text &= vbCrLf & oFinText.Text
        End If
    End Sub

    Private Sub tbCtrlInvoice_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbCtrlInvoice.Enter
        If Not loading Then
            txtInvoiced.Focus()
        Else
            Me.txtInvoiceNumber.Focus()
        End If
    End Sub

    Private Sub txtInvoiceNumber_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtInvoiceNumber.GotFocus
        loading = False
    End Sub
End Class
