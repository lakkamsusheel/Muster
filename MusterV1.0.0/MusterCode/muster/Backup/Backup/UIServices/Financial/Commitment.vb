'''  Upgrades and Fixes
'''      Thomas Franey               2/23/2009            Added Text Box to show selected report Due date
'''                                                       Added  Drag Drop functionaility between new textnox and Statement box
'''                                                       See Drag Drop events Region for details              
'''                                                       Fix Close Bug on line 875
'''      Thomas Franey               2/25/2009            Placed CostFormat into a Panel and Set Docking to lock
''                                                        control sizing and prevent cut off of words          


Public Class Commitment
    Inherits System.Windows.Forms.Form
#Region " Local Variables "

    Private IsMouseDown As Boolean = False
    Private changed As Boolean = False

    Friend FinancialEventID As Int64
    Friend FinancialCommitmentID As Int64

    Private bolLoading As Boolean
    Private statementLocX As Integer
    Private statementLocY As Integer

    Private oFinancialEvent As New MUSTER.BusinessLogic.pFinancial
    Private oFinancialActivity As New MUSTER.BusinessLogic.pFinancialActivity
    Private oFinancialCommitment As New MUSTER.BusinessLogic.pFinancialCommitment
    Private sStartingTotal As Double
    Private sFinalTotal As Double
    Private returnVal As String = String.Empty
    Friend ugCommitRow As Infragistics.Win.UltraWinGrid.UltraGridRow
    Friend CallingForm As Form
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
    Friend WithEvents pnlCommitment As System.Windows.Forms.Panel
    Friend WithEvents pnlCommitmentBottom As System.Windows.Forms.Panel
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents lblActivitytitle As System.Windows.Forms.Label
    Friend WithEvents cmbActivity As System.Windows.Forms.ComboBox
    Friend WithEvents txtPayee As System.Windows.Forms.TextBox
    Friend WithEvents lblPayee As System.Windows.Forms.Label
    Friend WithEvents chkThirdPartyPayment As System.Windows.Forms.CheckBox
    Friend WithEvents dtPickSOW As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblSOWDate As System.Windows.Forms.Label
    Friend WithEvents txtPO As System.Windows.Forms.TextBox
    Friend WithEvents lblPO As System.Windows.Forms.Label
    Friend WithEvents lblContractTypetitle As System.Windows.Forms.Label
    Friend WithEvents cmbContractType As System.Windows.Forms.ComboBox
    Friend WithEvents lblApprovedDateTitle As System.Windows.Forms.Label
    Friend WithEvents dtPickApproved As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblFundingTypeTitle As System.Windows.Forms.Label
    Friend WithEvents cmbFundingType As System.Windows.Forms.ComboBox
    Friend WithEvents lblCommitmentIDValue As System.Windows.Forms.Label
    Friend WithEvents lblCommitmentID As System.Windows.Forms.Label
    Friend WithEvents lblActivity As System.Windows.Forms.Label
    Friend WithEvents lblContractType As System.Windows.Forms.Label
    Friend WithEvents lblApprovedDate As System.Windows.Forms.Label
    Friend WithEvents lblFundingType As System.Windows.Forms.Label
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents pnlTechReports As System.Windows.Forms.Panel
    Friend WithEvents txDueDate As System.Windows.Forms.TextBox
    Friend WithEvents lblDue_Date As System.Windows.Forms.Label
    Friend WithEvents txtDueDtStatement As System.Windows.Forms.TextBox
    Friend WithEvents ugTechReports As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents lblDueDateStatement As System.Windows.Forms.Label
    Friend WithEvents lblTechReports As System.Windows.Forms.Label
    Friend WithEvents txtComments As System.Windows.Forms.TextBox
    Friend WithEvents lblComments As System.Windows.Forms.Label
    Friend WithEvents txtCondForReimbursement As System.Windows.Forms.TextBox
    Friend WithEvents btnInsertCondition As System.Windows.Forms.Button
    Friend WithEvents cmbAdditionalConditions As System.Windows.Forms.ComboBox
    Friend WithEvents lblAdditionalConditions As System.Windows.Forms.Label
    Friend WithEvents lblConditionsReimbursement As System.Windows.Forms.Label
    Friend WithEvents cfActivityCosts As MUSTER.CostFormat
    Friend WithEvents chkReimburseERAC As System.Windows.Forms.CheckBox
    Friend WithEvents BtnRegisterSOW As System.Windows.Forms.Button
    Friend WithEvents BtnprintScren As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlCommitment = New System.Windows.Forms.Panel
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.cfActivityCosts = New MUSTER.CostFormat
        Me.pnlTechReports = New System.Windows.Forms.Panel
        Me.BtnRegisterSOW = New System.Windows.Forms.Button
        Me.txDueDate = New System.Windows.Forms.TextBox
        Me.lblDue_Date = New System.Windows.Forms.Label
        Me.txtDueDtStatement = New System.Windows.Forms.TextBox
        Me.ugTechReports = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.lblDueDateStatement = New System.Windows.Forms.Label
        Me.lblTechReports = New System.Windows.Forms.Label
        Me.txtComments = New System.Windows.Forms.TextBox
        Me.lblComments = New System.Windows.Forms.Label
        Me.txtCondForReimbursement = New System.Windows.Forms.TextBox
        Me.btnInsertCondition = New System.Windows.Forms.Button
        Me.cmbAdditionalConditions = New System.Windows.Forms.ComboBox
        Me.lblAdditionalConditions = New System.Windows.Forms.Label
        Me.lblConditionsReimbursement = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.chkReimburseERAC = New System.Windows.Forms.CheckBox
        Me.lblActivitytitle = New System.Windows.Forms.Label
        Me.cmbActivity = New System.Windows.Forms.ComboBox
        Me.txtPayee = New System.Windows.Forms.TextBox
        Me.lblPayee = New System.Windows.Forms.Label
        Me.chkThirdPartyPayment = New System.Windows.Forms.CheckBox
        Me.dtPickSOW = New System.Windows.Forms.DateTimePicker
        Me.lblSOWDate = New System.Windows.Forms.Label
        Me.txtPO = New System.Windows.Forms.TextBox
        Me.lblPO = New System.Windows.Forms.Label
        Me.lblContractTypetitle = New System.Windows.Forms.Label
        Me.cmbContractType = New System.Windows.Forms.ComboBox
        Me.lblApprovedDateTitle = New System.Windows.Forms.Label
        Me.dtPickApproved = New System.Windows.Forms.DateTimePicker
        Me.lblFundingTypeTitle = New System.Windows.Forms.Label
        Me.cmbFundingType = New System.Windows.Forms.ComboBox
        Me.lblCommitmentIDValue = New System.Windows.Forms.Label
        Me.lblCommitmentID = New System.Windows.Forms.Label
        Me.lblActivity = New System.Windows.Forms.Label
        Me.lblContractType = New System.Windows.Forms.Label
        Me.lblApprovedDate = New System.Windows.Forms.Label
        Me.lblFundingType = New System.Windows.Forms.Label
        Me.pnlCommitmentBottom = New System.Windows.Forms.Panel
        Me.BtnprintScren = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.pnlCommitment.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.pnlTechReports.SuspendLayout()
        CType(Me.ugTechReports, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.pnlCommitmentBottom.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlCommitment
        '
        Me.pnlCommitment.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlCommitment.AutoScroll = True
        Me.pnlCommitment.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlCommitment.Controls.Add(Me.Panel2)
        Me.pnlCommitment.Controls.Add(Me.Panel1)
        Me.pnlCommitment.Location = New System.Drawing.Point(0, 0)
        Me.pnlCommitment.Name = "pnlCommitment"
        Me.pnlCommitment.Size = New System.Drawing.Size(944, 600)
        Me.pnlCommitment.TabIndex = 0
        '
        'Panel2
        '
        Me.Panel2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel2.Controls.Add(Me.cfActivityCosts)
        Me.Panel2.Controls.Add(Me.pnlTechReports)
        Me.Panel2.Controls.Add(Me.txtComments)
        Me.Panel2.Controls.Add(Me.lblComments)
        Me.Panel2.Controls.Add(Me.txtCondForReimbursement)
        Me.Panel2.Controls.Add(Me.btnInsertCondition)
        Me.Panel2.Controls.Add(Me.cmbAdditionalConditions)
        Me.Panel2.Controls.Add(Me.lblAdditionalConditions)
        Me.Panel2.Controls.Add(Me.lblConditionsReimbursement)
        Me.Panel2.Location = New System.Drawing.Point(0, 128)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(940, 472)
        Me.Panel2.TabIndex = 37
        '
        'cfActivityCosts
        '
        Me.cfActivityCosts.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cfActivityCosts.AutoScroll = True
        Me.cfActivityCosts.Location = New System.Drawing.Point(8, 8)
        Me.cfActivityCosts.Name = "cfActivityCosts"
        Me.cfActivityCosts.Size = New System.Drawing.Size(928, 168)
        Me.cfActivityCosts.TabIndex = 44
        '
        'pnlTechReports
        '
        Me.pnlTechReports.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.pnlTechReports.Controls.Add(Me.BtnRegisterSOW)
        Me.pnlTechReports.Controls.Add(Me.txDueDate)
        Me.pnlTechReports.Controls.Add(Me.lblDue_Date)
        Me.pnlTechReports.Controls.Add(Me.txtDueDtStatement)
        Me.pnlTechReports.Controls.Add(Me.ugTechReports)
        Me.pnlTechReports.Controls.Add(Me.lblDueDateStatement)
        Me.pnlTechReports.Controls.Add(Me.lblTechReports)
        Me.pnlTechReports.Location = New System.Drawing.Point(24, 320)
        Me.pnlTechReports.Name = "pnlTechReports"
        Me.pnlTechReports.Size = New System.Drawing.Size(904, 136)
        Me.pnlTechReports.TabIndex = 43
        '
        'BtnRegisterSOW
        '
        Me.BtnRegisterSOW.Location = New System.Drawing.Point(264, 8)
        Me.BtnRegisterSOW.Name = "BtnRegisterSOW"
        Me.BtnRegisterSOW.Size = New System.Drawing.Size(200, 23)
        Me.BtnRegisterSOW.TabIndex = 40
        Me.BtnRegisterSOW.Text = "Request Send to Finance"
        Me.BtnRegisterSOW.Visible = False
        '
        'txDueDate
        '
        Me.txDueDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txDueDate.Location = New System.Drawing.Point(72, 112)
        Me.txDueDate.Name = "txDueDate"
        Me.txDueDate.ReadOnly = True
        Me.txDueDate.Size = New System.Drawing.Size(136, 20)
        Me.txDueDate.TabIndex = 38
        Me.txDueDate.Text = ""
        '
        'lblDue_Date
        '
        Me.lblDue_Date.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblDue_Date.Location = New System.Drawing.Point(8, 112)
        Me.lblDue_Date.Name = "lblDue_Date"
        Me.lblDue_Date.Size = New System.Drawing.Size(64, 17)
        Me.lblDue_Date.TabIndex = 39
        Me.lblDue_Date.Text = "Due Date:"
        '
        'txtDueDtStatement
        '
        Me.txtDueDtStatement.AllowDrop = True
        Me.txtDueDtStatement.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtDueDtStatement.Location = New System.Drawing.Point(472, 24)
        Me.txtDueDtStatement.Multiline = True
        Me.txtDueDtStatement.Name = "txtDueDtStatement"
        Me.txtDueDtStatement.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtDueDtStatement.Size = New System.Drawing.Size(416, 72)
        Me.txtDueDtStatement.TabIndex = 18
        Me.txtDueDtStatement.Text = ""
        '
        'ugTechReports
        '
        Me.ugTechReports.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ugTechReports.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugTechReports.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugTechReports.Location = New System.Drawing.Point(8, 32)
        Me.ugTechReports.Name = "ugTechReports"
        Me.ugTechReports.Size = New System.Drawing.Size(456, 64)
        Me.ugTechReports.TabIndex = 17
        Me.ugTechReports.UpdateMode = Infragistics.Win.UltraWinGrid.UpdateMode.OnCellChange
        '
        'lblDueDateStatement
        '
        Me.lblDueDateStatement.Location = New System.Drawing.Point(472, 8)
        Me.lblDueDateStatement.Name = "lblDueDateStatement"
        Me.lblDueDateStatement.Size = New System.Drawing.Size(110, 16)
        Me.lblDueDateStatement.TabIndex = 37
        Me.lblDueDateStatement.Text = "Due Date Statement:"
        '
        'lblTechReports
        '
        Me.lblTechReports.Location = New System.Drawing.Point(8, 8)
        Me.lblTechReports.Name = "lblTechReports"
        Me.lblTechReports.Size = New System.Drawing.Size(104, 17)
        Me.lblTechReports.TabIndex = 36
        Me.lblTechReports.Text = "Technical Reports:"
        Me.lblTechReports.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'txtComments
        '
        Me.txtComments.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtComments.Location = New System.Drawing.Point(488, 272)
        Me.txtComments.Multiline = True
        Me.txtComments.Name = "txtComments"
        Me.txtComments.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtComments.Size = New System.Drawing.Size(424, 40)
        Me.txtComments.TabIndex = 39
        Me.txtComments.Text = ""
        '
        'lblComments
        '
        Me.lblComments.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblComments.Location = New System.Drawing.Point(488, 256)
        Me.lblComments.Name = "lblComments"
        Me.lblComments.Size = New System.Drawing.Size(68, 17)
        Me.lblComments.TabIndex = 42
        Me.lblComments.Text = "Comments:"
        '
        'txtCondForReimbursement
        '
        Me.txtCondForReimbursement.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtCondForReimbursement.Location = New System.Drawing.Point(24, 272)
        Me.txtCondForReimbursement.Multiline = True
        Me.txtCondForReimbursement.Name = "txtCondForReimbursement"
        Me.txtCondForReimbursement.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtCondForReimbursement.Size = New System.Drawing.Size(440, 40)
        Me.txtCondForReimbursement.TabIndex = 38
        Me.txtCondForReimbursement.Text = ""
        '
        'btnInsertCondition
        '
        Me.btnInsertCondition.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnInsertCondition.Location = New System.Drawing.Point(480, 232)
        Me.btnInsertCondition.Name = "btnInsertCondition"
        Me.btnInsertCondition.Size = New System.Drawing.Size(44, 23)
        Me.btnInsertCondition.TabIndex = 37
        Me.btnInsertCondition.Text = "Insert"
        '
        'cmbAdditionalConditions
        '
        Me.cmbAdditionalConditions.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmbAdditionalConditions.Location = New System.Drawing.Point(144, 232)
        Me.cmbAdditionalConditions.Name = "cmbAdditionalConditions"
        Me.cmbAdditionalConditions.Size = New System.Drawing.Size(328, 21)
        Me.cmbAdditionalConditions.TabIndex = 36
        '
        'lblAdditionalConditions
        '
        Me.lblAdditionalConditions.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblAdditionalConditions.Location = New System.Drawing.Point(32, 232)
        Me.lblAdditionalConditions.Name = "lblAdditionalConditions"
        Me.lblAdditionalConditions.Size = New System.Drawing.Size(118, 17)
        Me.lblAdditionalConditions.TabIndex = 41
        Me.lblAdditionalConditions.Text = "Additional Conditions:"
        '
        'lblConditionsReimbursement
        '
        Me.lblConditionsReimbursement.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblConditionsReimbursement.Location = New System.Drawing.Point(24, 256)
        Me.lblConditionsReimbursement.Name = "lblConditionsReimbursement"
        Me.lblConditionsReimbursement.Size = New System.Drawing.Size(172, 17)
        Me.lblConditionsReimbursement.TabIndex = 40
        Me.lblConditionsReimbursement.Text = "Conditions for Reimbursement:"
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel1.Controls.Add(Me.chkReimburseERAC)
        Me.Panel1.Controls.Add(Me.lblActivitytitle)
        Me.Panel1.Controls.Add(Me.cmbActivity)
        Me.Panel1.Controls.Add(Me.txtPayee)
        Me.Panel1.Controls.Add(Me.lblPayee)
        Me.Panel1.Controls.Add(Me.chkThirdPartyPayment)
        Me.Panel1.Controls.Add(Me.dtPickSOW)
        Me.Panel1.Controls.Add(Me.lblSOWDate)
        Me.Panel1.Controls.Add(Me.txtPO)
        Me.Panel1.Controls.Add(Me.lblPO)
        Me.Panel1.Controls.Add(Me.lblContractTypetitle)
        Me.Panel1.Controls.Add(Me.cmbContractType)
        Me.Panel1.Controls.Add(Me.lblApprovedDateTitle)
        Me.Panel1.Controls.Add(Me.dtPickApproved)
        Me.Panel1.Controls.Add(Me.lblFundingTypeTitle)
        Me.Panel1.Controls.Add(Me.cmbFundingType)
        Me.Panel1.Controls.Add(Me.lblCommitmentIDValue)
        Me.Panel1.Controls.Add(Me.lblCommitmentID)
        Me.Panel1.Controls.Add(Me.lblActivity)
        Me.Panel1.Controls.Add(Me.lblContractType)
        Me.Panel1.Controls.Add(Me.lblApprovedDate)
        Me.Panel1.Controls.Add(Me.lblFundingType)
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(944, 128)
        Me.Panel1.TabIndex = 36
        '
        'chkReimburseERAC
        '
        Me.chkReimburseERAC.Location = New System.Drawing.Point(608, 8)
        Me.chkReimburseERAC.Name = "chkReimburseERAC"
        Me.chkReimburseERAC.Size = New System.Drawing.Size(120, 16)
        Me.chkReimburseERAC.TabIndex = 61
        Me.chkReimburseERAC.Text = "Reimburse ERAC"
        Me.chkReimburseERAC.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'lblActivitytitle
        '
        Me.lblActivitytitle.Location = New System.Drawing.Point(48, 104)
        Me.lblActivitytitle.Name = "lblActivitytitle"
        Me.lblActivitytitle.Size = New System.Drawing.Size(47, 23)
        Me.lblActivitytitle.TabIndex = 56
        Me.lblActivitytitle.Text = "Activity:"
        '
        'cmbActivity
        '
        Me.cmbActivity.Location = New System.Drawing.Point(104, 104)
        Me.cmbActivity.Name = "cmbActivity"
        Me.cmbActivity.Size = New System.Drawing.Size(616, 21)
        Me.cmbActivity.TabIndex = 51
        '
        'txtPayee
        '
        Me.txtPayee.Location = New System.Drawing.Point(616, 80)
        Me.txtPayee.Name = "txtPayee"
        Me.txtPayee.TabIndex = 50
        Me.txtPayee.Text = ""
        '
        'lblPayee
        '
        Me.lblPayee.Location = New System.Drawing.Point(576, 80)
        Me.lblPayee.Name = "lblPayee"
        Me.lblPayee.Size = New System.Drawing.Size(39, 17)
        Me.lblPayee.TabIndex = 55
        Me.lblPayee.Text = "Payee:"
        '
        'chkThirdPartyPayment
        '
        Me.chkThirdPartyPayment.Location = New System.Drawing.Point(312, 80)
        Me.chkThirdPartyPayment.Name = "chkThirdPartyPayment"
        Me.chkThirdPartyPayment.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkThirdPartyPayment.Size = New System.Drawing.Size(128, 24)
        Me.chkThirdPartyPayment.TabIndex = 49
        Me.chkThirdPartyPayment.Text = "Third Party Payment"
        '
        'dtPickSOW
        '
        Me.dtPickSOW.Checked = False
        Me.dtPickSOW.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickSOW.Location = New System.Drawing.Point(616, 56)
        Me.dtPickSOW.Name = "dtPickSOW"
        Me.dtPickSOW.Size = New System.Drawing.Size(88, 20)
        Me.dtPickSOW.TabIndex = 45
        '
        'lblSOWDate
        '
        Me.lblSOWDate.Location = New System.Drawing.Point(552, 56)
        Me.lblSOWDate.Name = "lblSOWDate"
        Me.lblSOWDate.Size = New System.Drawing.Size(61, 17)
        Me.lblSOWDate.TabIndex = 54
        Me.lblSOWDate.Text = "SOW Date:"
        '
        'txtPO
        '
        Me.txtPO.Location = New System.Drawing.Point(616, 32)
        Me.txtPO.Name = "txtPO"
        Me.txtPO.TabIndex = 42
        Me.txtPO.Text = ""
        '
        'lblPO
        '
        Me.lblPO.Location = New System.Drawing.Point(584, 32)
        Me.lblPO.Name = "lblPO"
        Me.lblPO.Size = New System.Drawing.Size(32, 17)
        Me.lblPO.TabIndex = 53
        Me.lblPO.Text = "PO#:"
        '
        'lblContractTypetitle
        '
        Me.lblContractTypetitle.Location = New System.Drawing.Point(16, 80)
        Me.lblContractTypetitle.Name = "lblContractTypetitle"
        Me.lblContractTypetitle.Size = New System.Drawing.Size(80, 17)
        Me.lblContractTypetitle.TabIndex = 52
        Me.lblContractTypetitle.Text = "Contract Type:"
        '
        'cmbContractType
        '
        Me.cmbContractType.Location = New System.Drawing.Point(104, 80)
        Me.cmbContractType.Name = "cmbContractType"
        Me.cmbContractType.Size = New System.Drawing.Size(121, 21)
        Me.cmbContractType.TabIndex = 47
        '
        'lblApprovedDateTitle
        '
        Me.lblApprovedDateTitle.Location = New System.Drawing.Point(16, 56)
        Me.lblApprovedDateTitle.Name = "lblApprovedDateTitle"
        Me.lblApprovedDateTitle.Size = New System.Drawing.Size(83, 17)
        Me.lblApprovedDateTitle.TabIndex = 48
        Me.lblApprovedDateTitle.Text = "Date Approved:"
        '
        'dtPickApproved
        '
        Me.dtPickApproved.Checked = False
        Me.dtPickApproved.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickApproved.Location = New System.Drawing.Point(104, 56)
        Me.dtPickApproved.Name = "dtPickApproved"
        Me.dtPickApproved.Size = New System.Drawing.Size(86, 20)
        Me.dtPickApproved.TabIndex = 44
        '
        'lblFundingTypeTitle
        '
        Me.lblFundingTypeTitle.Location = New System.Drawing.Point(24, 32)
        Me.lblFundingTypeTitle.Name = "lblFundingTypeTitle"
        Me.lblFundingTypeTitle.Size = New System.Drawing.Size(77, 17)
        Me.lblFundingTypeTitle.TabIndex = 46
        Me.lblFundingTypeTitle.Text = "Funding Type:"
        '
        'cmbFundingType
        '
        Me.cmbFundingType.Location = New System.Drawing.Point(104, 32)
        Me.cmbFundingType.Name = "cmbFundingType"
        Me.cmbFundingType.Size = New System.Drawing.Size(121, 21)
        Me.cmbFundingType.TabIndex = 40
        '
        'lblCommitmentIDValue
        '
        Me.lblCommitmentIDValue.Location = New System.Drawing.Point(104, 8)
        Me.lblCommitmentIDValue.Name = "lblCommitmentIDValue"
        Me.lblCommitmentIDValue.TabIndex = 43
        '
        'lblCommitmentID
        '
        Me.lblCommitmentID.Location = New System.Drawing.Point(72, 8)
        Me.lblCommitmentID.Name = "lblCommitmentID"
        Me.lblCommitmentID.Size = New System.Drawing.Size(19, 17)
        Me.lblCommitmentID.TabIndex = 41
        Me.lblCommitmentID.Text = "ID:"
        '
        'lblActivity
        '
        Me.lblActivity.BackColor = System.Drawing.SystemColors.Window
        Me.lblActivity.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblActivity.Location = New System.Drawing.Point(104, 104)
        Me.lblActivity.Name = "lblActivity"
        Me.lblActivity.Size = New System.Drawing.Size(430, 21)
        Me.lblActivity.TabIndex = 60
        '
        'lblContractType
        '
        Me.lblContractType.BackColor = System.Drawing.SystemColors.Window
        Me.lblContractType.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblContractType.Location = New System.Drawing.Point(104, 80)
        Me.lblContractType.Name = "lblContractType"
        Me.lblContractType.Size = New System.Drawing.Size(121, 21)
        Me.lblContractType.TabIndex = 59
        '
        'lblApprovedDate
        '
        Me.lblApprovedDate.BackColor = System.Drawing.SystemColors.Window
        Me.lblApprovedDate.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblApprovedDate.Location = New System.Drawing.Point(104, 56)
        Me.lblApprovedDate.Name = "lblApprovedDate"
        Me.lblApprovedDate.Size = New System.Drawing.Size(86, 20)
        Me.lblApprovedDate.TabIndex = 58
        '
        'lblFundingType
        '
        Me.lblFundingType.BackColor = System.Drawing.SystemColors.Window
        Me.lblFundingType.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblFundingType.Location = New System.Drawing.Point(104, 32)
        Me.lblFundingType.Name = "lblFundingType"
        Me.lblFundingType.Size = New System.Drawing.Size(121, 21)
        Me.lblFundingType.TabIndex = 57
        '
        'pnlCommitmentBottom
        '
        Me.pnlCommitmentBottom.Controls.Add(Me.BtnprintScren)
        Me.pnlCommitmentBottom.Controls.Add(Me.btnCancel)
        Me.pnlCommitmentBottom.Controls.Add(Me.btnSave)
        Me.pnlCommitmentBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlCommitmentBottom.Location = New System.Drawing.Point(0, 605)
        Me.pnlCommitmentBottom.Name = "pnlCommitmentBottom"
        Me.pnlCommitmentBottom.Size = New System.Drawing.Size(944, 40)
        Me.pnlCommitmentBottom.TabIndex = 19
        '
        'BtnprintScren
        '
        Me.BtnprintScren.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnprintScren.Location = New System.Drawing.Point(856, 8)
        Me.BtnprintScren.Name = "BtnprintScren"
        Me.BtnprintScren.TabIndex = 22
        Me.BtnprintScren.Text = "Print Screen"
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.btnCancel.Location = New System.Drawing.Point(460, 8)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 21
        Me.btnCancel.Text = "Cancel"
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(24, 8)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.TabIndex = 20
        Me.btnSave.Text = "Save"
        '
        'Commitment
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(944, 645)
        Me.Controls.Add(Me.pnlCommitment)
        Me.Controls.Add(Me.pnlCommitmentBottom)
        Me.Name = "Commitment"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Commitment"
        Me.pnlCommitment.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.pnlTechReports.ResumeLayout(False)
        CType(Me.ugTechReports, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.pnlCommitmentBottom.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


#Region "drag Drop events"

    Sub txtChange(ByVal sender As Object, ByVal e As EventArgs) Handles ugTechReports.MouseEnter
        If changed Then
            LoadReportGrid()
            changed = False
        End If
    End Sub

    Private Sub SetText(ByVal sender As Object, ByVal e As EventArgs)

        If TypeOf Me.ugTechReports.ActiveRow.Cells("Due_Date").Value Is DateTime Then
            txDueDate.Text = DirectCast(Me.ugTechReports.ActiveRow.Cells("Due_Date").Value, DateTime).ToString("MMM dd, yyyy")
        End If

    End Sub
    Private Sub AssignDoc(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs)

        If e.Cell.Column.Header.Caption = "Due_Date" Then
            Dim value As DateTime = e.NewValue
            ProcessTechreports(e.Cell.Row, value)

            If value > CDate("1/1/1900") Then
                Me.txDueDate.Text = value.ToString("MMM dd, yyyy")
            End If

        End If

    End Sub
    Private Sub mouseHoveringStatement(ByVal sender As Object, ByVal e As DragEventArgs)

        statementLocX = e.X
        statementLocY = e.Y


        With Me.txtDueDtStatement



            Dim newpnt As Point = .PointToScreen(New Point(0, 0))



            Dim point = (CInt((e.Y - newpnt.Y) / 15) - 1) * .Width

            Dim loc = CInt(((e.X - newpnt.X) + point) / 5)

            .Text = .Text.Replace("~", String.Empty)
            .Text = .Text.Insert(loc, "~")

        End With

        Me.Text = String.Format("{0},{1}", statementLocX, statementLocY)




    End Sub


    Private Sub Cell_MouseDown(ByVal sender As Object, ByVal e As _
System.Windows.Forms.MouseEventArgs)
        ' Set a flag to show that the mouse is down.
        IsMouseDown = True
    End Sub


    Private Sub Cell_MouseMove(ByVal sender As Object, ByVal e As _
    System.Windows.Forms.MouseEventArgs)

        If IsMouseDown Then
            ' Initiate dragging.
            DirectCast(sender, Control).DoDragDrop(DirectCast(sender, TextBox).Text, DragDropEffects.Copy)
        End If

        IsMouseDown = False
    End Sub


    Private Sub BtnRegisterSowClicked(ByVal sender As Object, ByVal e As _
   EventArgs) Handles BtnRegisterSOW.Click

        With DirectCast(sender, Button)

            If Not .Tag Is Nothing Then
                Me.SendToFinance(.Tag)
            End If
        End With

    End Sub

    Private Sub TextDueStatement_DragEnter(ByVal sender As Object, ByVal e As _
    System.Windows.Forms.DragEventArgs)
        ' Check the format of the data being dropped.
        If (e.Data.GetDataPresent(DataFormats.Text)) Then
            ' Display the copy cursor.
            e.Effect = DragDropEffects.Copy
        Else
            ' Display the no-drop cursor.
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub TextDueStatement_DragDrop(ByVal sender As Object, ByVal e As _
    System.Windows.Forms.DragEventArgs)
        ' Paste the text.


        With Me.txtDueDtStatement
            .Text = .Text.Replace("???", String.Empty).Replace("~", e.Data.GetData(DataFormats.Text))
        End With

    End Sub



#End Region


    Private Sub printOutCommitmentForm()
        With Me
            Dim x, y, xx, yy As Integer

            xx = .Width
            yy = .Height
            x = .Left
            y = .Top
            .btnCancel.Visible = False
            .BtnprintScren.Visible = False
            .BtnRegisterSOW.Visible = False
            .btnSave.Visible = False
            .txtCondForReimbursement.Visible = False
            .btnInsertCondition.Visible = False
            .Width = 1280
            .Height = 900
            .Top = 0
            .Left = 0

            .Refresh()

            Threading.Thread.Sleep(300)

            UIUtilsGen.CaptureScreen(Me, .txtCondForReimbursement.Text)

            .Height = xx
            .Width = yy
            .Left = x
            .Top = y



            .btnCancel.Visible = True
            .BtnprintScren.Visible = True
            .BtnRegisterSOW.Visible = True
            .btnSave.Visible = True
            .btnInsertCondition.Visible = True
            .txtCondForReimbursement.Visible = True

        End With

    End Sub

    Private Sub Commitment_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        oFinancialEvent.Retrieve(FinancialEventID)
        oFinancialCommitment.Retrieve(FinancialCommitmentID)
        cfActivityCosts.AssignCommitmentObject(oFinancialCommitment)
        sStartingTotal = 0
        sFinalTotal = 0

        If oFinancialCommitment.CommitmentID = 0 Then
            lblCommitmentIDValue.Text = "New"
        Else
            lblCommitmentIDValue.Text = oFinancialCommitment.CommitmentID
        End If

        LoadDropDowns()
        SetForm()
        'LoadReportGrid()

    End Sub

    Private Sub SetForm()

        Try
            bolLoading = True

            If oFinancialCommitment.CommitmentID = 0 Then
                cmbContractType.SelectedValue = 1071
                cmbFundingType.SelectedValue = 1080
                dtPickApproved.Value = Now.Date
                dtPickSOW.Value = Now.Date
                cmbActivity.SelectedIndex = 0

                oFinancialCommitment.Fin_Event_ID = FinancialEventID
                oFinancialCommitment.ContractType = 1071
                oFinancialCommitment.FundingType = 1080
                oFinancialCommitment.ApprovedDate = Now.Date
                oFinancialCommitment.SOWDate = Now.Date
                oFinancialCommitment.ActivityType = cmbActivity.SelectedValue
                oFinancialCommitment.ReimbursementCondition = txtCondForReimbursement.Text
                oFinancialCommitment.DueDateStatement = txtDueDtStatement.Text
                oFinancialCommitment.ThirdPartyPayee = txtPayee.Text
                oFinancialCommitment.ReimburseERAC = False


                lblActivity.Visible = False
                lblContractType.Visible = False
                lblFundingType.Visible = False
                lblApprovedDate.Visible = False

                txtPayee.Enabled = False
                resynchCostformat()

            Else
                cmbContractType.SelectedValue = oFinancialCommitment.ContractType
                cmbFundingType.SelectedValue = oFinancialCommitment.FundingType
                cmbActivity.SelectedValue = oFinancialCommitment.ActivityType

                cmbActivity.Visible = True
                lblActivity.Visible = False
                lblActivity.Text = cmbActivity.Text

                chkReimburseERAC.Checked = oFinancialCommitment.ReimburseERAC

                dtPickApproved.Value = oFinancialCommitment.ApprovedDate
                dtPickSOW.Value = oFinancialCommitment.SOWDate

                If oFinancialCommitment.PONumber <> String.Empty Then
                    lblContractType.Text = cmbContractType.Text
                    cmbContractType.Visible = False
                    lblContractType.Visible = True

                    lblFundingType.Text = cmbFundingType.Text
                    lblFundingType.Visible = True
                    cmbFundingType.Visible = False

                    lblApprovedDate.Text = dtPickApproved.Value.Date
                    lblApprovedDate.Visible = True
                    dtPickApproved.Visible = False
                Else
                    lblContractType.Visible = False
                    lblFundingType.Visible = False
                    lblApprovedDate.Visible = False
                End If


                cfActivityCosts.CostFormatType = oFinancialCommitment.Case_Letter
                cfActivityCosts.SetDisplay(False)
                cfActivityCosts.LoadCommitment()
                sStartingTotal = cfActivityCosts.GrandTotal

                chkThirdPartyPayment.Checked = oFinancialCommitment.ThirdPartyPayment
                If chkThirdPartyPayment.Checked Then
                    txtPayee.Enabled = True
                Else
                    txtPayee.Enabled = False
                End If


                txtCondForReimbursement.Text = oFinancialCommitment.ReimbursementCondition
                txtDueDtStatement.Text = oFinancialCommitment.DueDateStatement
                txtPayee.Text = oFinancialCommitment.ThirdPartyPayee
                txtComments.Text = oFinancialCommitment.Comments
                txtPO.Text = oFinancialCommitment.PONumber

                Me.LoadReportGrid()

            End If

            AddHandler txtDueDtStatement.DragEnter, AddressOf TextDueStatement_DragEnter
            AddHandler txtDueDtStatement.DragDrop, AddressOf TextDueStatement_DragDrop
            AddHandler txtDueDtStatement.DragOver, AddressOf mouseHoveringStatement



            bolLoading = False
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Load Commitment " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub LoadDropDowns()
        Try
            bolLoading = True
            Dim dtActivityList As DataTable = oFinancialActivity.PopulateFinancialActivityForCommitment
            Dim dtReimbursementConditions As DataTable = oFinancialActivity.PopulateFinancialAdditionalConditions
            Dim dtContractTypes As DataTable = oFinancialCommitment.PopulateFinancialContractTypes
            Dim dtFundingTypes As DataTable = oFinancialCommitment.PopulateFinancialFundingTypes
            Dim v As New BusinessLogic.pCompany


            cmbAdditionalConditions.DataSource = dtReimbursementConditions
            cmbAdditionalConditions.DisplayMember = "Text_Name"
            cmbAdditionalConditions.ValueMember = "Text_ID"

            cmbContractType.DataSource = dtContractTypes
            cmbContractType.DisplayMember = "PROPERTY_NAME"
            cmbContractType.ValueMember = "PROPERTY_ID"

            cmbFundingType.DataSource = dtFundingTypes
            cmbFundingType.DisplayMember = "PROPERTY_NAME"
            cmbFundingType.ValueMember = "PROPERTY_ID"

            cmbActivity.DataSource = dtActivityList
            cmbActivity.DisplayMember = "ActivityDesc"
            cmbActivity.ValueMember = "Activity_ID"

            bolLoading = False
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Load DropDowns " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()

        End Try


    End Sub
    Private Sub LoadReportGrid()
        Dim dsLocal As DataSet
        Dim dsLocal2 As DataSet

        Dim tmpBand As Int16
        Dim dtTotals As DataTable

        dsLocal = oFinancialEvent.PopulateCommitmentTecDocList(oFinancialEvent.TecEventID, FinancialCommitmentID, Convert.ToInt32(oFinancialCommitment.ActivityType), "RPT")

        If dsLocal Is Nothing OrElse dsLocal.Tables(0).Rows.Count = 0 Then
            dsLocal2 = oFinancialEvent.PopulateCommitmentTecDocList(oFinancialEvent.TecEventID, FinancialCommitmentID, Convert.ToInt32(oFinancialCommitment.ActivityType), "SOW")
        Else
            Me.BtnRegisterSOW.Visible = False
            Me.BtnRegisterSOW.Tag = Nothing

        End If


        If Not dsLocal2 Is Nothing AndAlso dsLocal2.Tables(0).Rows.Count > 0 Then

            Me.BtnRegisterSOW.Text = "Request Sent to Financial"
            Me.BtnRegisterSOW.Visible = True
            Me.BtnRegisterSOW.Tag = dsLocal2.Tables(0).Rows(0).Item("Event_Activity_Document_ID")
            Me.BtnRegisterSOW.Enabled = True

        ElseIf Not dsLocal2 Is Nothing Then
            Me.BtnRegisterSOW.Text = "Scope of work Required"
            Me.BtnRegisterSOW.Visible = True
            Me.BtnRegisterSOW.Enabled = False
            Me.BtnRegisterSOW.Tag = Nothing
        End If


        If Not dsLocal Is Nothing AndAlso dsLocal.Tables(0).Rows.Count > 0 Then


            ugTechReports.DataSource = dsLocal
            ugTechReports.Rows.CollapseAll(True)
            ugTechReports.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
            'ugTechReports.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect

            'If dsLocal.Tables(0).Rows.Count > 0 Then ' Does Payment Table have rows

            With ugTechReports.DisplayLayout.Bands(0).Columns

                .Item("Event_Activity_Document_ID").Hidden = True
                .Item("Extension_Date").Hidden = True
                .Item("Received_Date").Hidden = True
                .Item("Date_Sent_To_Finance").Hidden = True
                .Item("Commitment_ID").Hidden = True
                .Item("Paid").Hidden = True
                .Item("Event_ID").Hidden = True

            End With


            RemoveHandler ugTechReports.BeforeCellUpdate, AddressOf AssignDoc
            RemoveHandler ugTechReports.AfterRowActivate, AddressOf SetText

            RemoveHandler txDueDate.MouseDown, AddressOf Cell_MouseDown
            RemoveHandler txDueDate.MouseMove, AddressOf Cell_MouseMove


            AddHandler ugTechReports.BeforeCellUpdate, AddressOf AssignDoc
            AddHandler ugTechReports.AfterRowActivate, AddressOf SetText

            AddHandler txDueDate.MouseDown, AddressOf Cell_MouseDown
            AddHandler txDueDate.MouseMove, AddressOf Cell_MouseMove


            ugTechReports.DisplayLayout.Bands(0).Columns("Report").Width = 140
            ugTechReports.DisplayLayout.Bands(0).Columns("Report").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugTechReports.DisplayLayout.Bands(0).Columns("Status").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit



            ugTechReports.DisplayLayout.Bands(0).Columns("Due_Date").Width = 90
            ugTechReports.DisplayLayout.Bands(0).Columns("Date_Closed").Width = 90
            ugTechReports.DisplayLayout.Bands(0).Columns("Status").Width = 90

        Else
            ugTechReports.DataSource = Nothing
        End If

        'End If


    End Sub

    'Private Sub cfActivityCosts_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cfActivityCosts.Load

    'End Sub

    Private Sub cmbActivity_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbActivity.SelectedIndexChanged
        If bolLoading = True Then Exit Sub

        oFinancialCommitment.ActivityType = cmbActivity.SelectedValue

        resynchCostformat()

        Me.LoadReportGrid()

    End Sub
    Private Sub resynchCostformat()
        Dim strTempFormat As String

        oFinancialActivity.Retrieve(cmbActivity.SelectedValue)
        Select Case Me.cmbContractType.SelectedValue
            Case 1069
                strTempFormat = oFinancialActivity.FixedPriceDesc
            Case 1070
                strTempFormat = oFinancialActivity.CostPlusDesc
            Case 1071
                strTempFormat = oFinancialActivity.TimeAndMaterialsDesc
        End Select
        bolLoading = True

        txtDueDtStatement.Text = oFinancialActivity.DueDateStatement
        oFinancialCommitment.DueDateStatement = txtDueDtStatement.Text

        txtCondForReimbursement.Text = oFinancialActivity.ReimbursementConditionDesc
        oFinancialCommitment.ReimbursementCondition = txtCondForReimbursement.Text

        oFinancialCommitment.Case_Letter = strTempFormat
        bolLoading = False
        cfActivityCosts.CostFormatType = strTempFormat
        cfActivityCosts.SetDisplay(True)


    End Sub

    Private Sub cmbContractType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbContractType.SelectedIndexChanged
        If bolLoading = True Then Exit Sub

        oFinancialCommitment.ContractType = cmbContractType.SelectedValue
        resynchCostformat()
    End Sub

    Private Sub cmbFundingType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbFundingType.SelectedIndexChanged
        If bolLoading = True Then Exit Sub

        oFinancialCommitment.FundingType = cmbFundingType.SelectedValue
    End Sub


    Private Sub chkThirdPartyPayment_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkThirdPartyPayment.CheckedChanged
        If bolLoading = True Then Exit Sub

        oFinancialCommitment.ThirdPartyPayment = chkThirdPartyPayment.Checked

        If chkThirdPartyPayment.Checked Then
            txtPayee.Enabled = True
        Else
            txtPayee.Enabled = False
            oFinancialCommitment.ThirdPartyPayee = ""
            bolLoading = True
            txtPayee.Text = ""
            bolLoading = False
        End If
    End Sub

    Private Sub txtPayee_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPayee.TextChanged
        If bolLoading = True Then Exit Sub
        oFinancialCommitment.ThirdPartyPayee = txtPayee.Text
    End Sub

    Private Sub dtPickApproved_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtPickApproved.ValueChanged
        If bolLoading = True Then Exit Sub
        oFinancialCommitment.ApprovedDate = dtPickApproved.Value.Date
    End Sub

    Private Sub dtPickSOW_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtPickSOW.ValueChanged
        If bolLoading = True Then Exit Sub
        oFinancialCommitment.SOWDate = dtPickSOW.Value.Date
    End Sub

    Private Sub txtCondForReimbursement_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCondForReimbursement.TextChanged
        If bolLoading = True Then Exit Sub

        oFinancialCommitment.ReimbursementCondition = txtCondForReimbursement.Text
    End Sub

    Private Sub txtComments_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtComments.TextChanged
        If bolLoading = True Then Exit Sub

        oFinancialCommitment.Comments = txtComments.Text
    End Sub

    Private Sub txtDueDtStatement_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDueDtStatement.TextChanged
        If bolLoading = True Then Exit Sub

        oFinancialCommitment.DueDateStatement = txtDueDtStatement.Text
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        oFinancialCommitment.Reset()
        oFinancialCommitment.IsDirty = False

        Me.Close()

    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim sFinalTotalMinusPayments As Double
        Dim bolAllowNegativePayment As Boolean = (lblActivity.Text.IndexOf("Cost Recovery") > -1 Or cmbActivity.Text.IndexOf("Cost Recovery") > -1)
        ' #1225 validate only if state
        Dim bolValidate1500000_1200000 As Boolean = (oFinancialCommitment.FundingType = 1078 Or _
                                                        oFinancialCommitment.FundingType = 1079 Or _
                                                        oFinancialCommitment.FundingType = 1080)
        Try
            'validation for less than zero balance.
            If oFinancialCommitment.CommitmentID > 0 Then
                If Not ugCommitRow Is Nothing Then
                    Dim sBalance As Double
                    sBalance = (cfActivityCosts.GrandTotal + IIf(IsDBNull(ugCommitRow.Cells("Adjustment").Value) = True, 0.0, CDbl(ugCommitRow.Cells("Adjustment").Value))) - IIf(IsDBNull(ugCommitRow.Cells("Payment").Value) = True, 0.0, CDbl(ugCommitRow.Cells("Payment").Value))
                    If sBalance < 0 And Not bolAllowNegativePayment Then
                        MsgBox("Balance Cannot Be Less Than Zero.")
                        Exit Sub
                    End If
                End If
            End If

            sFinalTotal = GetCommitmentTotals(False, True)
            sFinalTotalMinusPayments = GetCommitmentTotals(True, True)
            If sFinalTotal > 1500000 And bolValidate1500000_1200000 Then
                MsgBox("Total Commitments Exceed $1,500,000.00 Save Not Allowed")
                Exit Sub
            Else
                If sFinalTotal >= 1200000 And bolValidate1500000_1200000 Then
                    If MsgBox("Total Commitments Exceed $1200,000.00.  So you wish to save?", MsgBoxStyle.YesNo) = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If
                If sFinalTotalMinusPayments < -0.009 Then
                    If MsgBox("Total Commitments Cannot Be Less Than Zero.") Then
                        Exit Sub
                    End If
                End If
                ''''''''''''''''''''''''''''''''''''''''''
                'Commented by Padmaja
                'IsDirty is false on modification of Technical Reports which is not allowing to 
                'Change the commitment
                'If oFinancialCommitment.IsDirty Then

                If oFinancialCommitment.CommitmentID <= 0 Then
                    oFinancialCommitment.CreatedBy = MusterContainer.AppUser.ID
                Else
                    oFinancialCommitment.ModifiedBy = MusterContainer.AppUser.ID
                End If
                oFinancialCommitment.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                CallingForm.Tag = "C" + oFinancialCommitment.CommitmentID.ToString

                ProcessTechreports()
                MsgBox("Commitment Saved")
                Me.Close()
                'End If
                Dim strDescription As String = String.Empty
                If Me.chkThirdPartyPayment.Checked Then
                    strDescription = "Third Party Financial Commitment Exceeds $800,000.00"
                Else
                    strDescription = "Non Third Party Financial Commitment Exceeds $800,000.00"
                End If

                If sFinalTotal >= 1200000 And bolValidate1500000_1200000 Then

                    'Adding a Flag Entry.
                    'If MusterContainer.pFlag.Retrieve(oFinancialCommitment.Fin_Event_ID, 32, strDescription) Is Nothing Then
                    MusterContainer.pFlag.RetrieveFlags(oFinancialCommitment.Fin_Event_ID, UIUtilsGen.EntityTypes.FinancialEvent, , "Financial", , , "SYSTEM", strDescription)
                    If MusterContainer.pFlag.FlagsCol.Values.Count <= 0 Then
                        MusterContainer.pFlag.Add(New MUSTER.Info.FlagInfo(0, _
                             oFinancialCommitment.Fin_Event_ID, _
                             32, _
                             strDescription, _
                             False, _
                             Now.Date, _
                             "Financial", _
                             0, _
                             MusterContainer.AppUser.ID, _
                             Now, _
                             MusterContainer.AppUser.ID, _
                             CDate("01/01/0001"), _
                             CDate("01/01/0001"), _
                             "SYSTEM", _
                             "RED"))

                        MusterContainer.pFlag.Save()
                    End If
                Else
                    ' if flag exists, delete
                    'Dim flagsCol As MUSTER.Info.FlagsCollection
                    MusterContainer.pFlag.RetrieveFlags(oFinancialCommitment.Fin_Event_ID, UIUtilsGen.EntityTypes.FinancialEvent, , "Financial", , , "SYSTEM", strDescription)
                    For Each flagInfo As MUSTER.Info.FlagInfo In MusterContainer.pFlag.FlagsCol.Values
                        If flagInfo.FlagDescription = Trim(strDescription) Then
                            MusterContainer.pFlag.FlagInfo = flagInfo
                            MusterContainer.pFlag.Deleted = True
                            MusterContainer.pFlag.Save()
                        End If
                    Next
                End If
            End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Save Failed " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub txtPO_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPO.TextChanged
        If bolLoading = True Then Exit Sub
        oFinancialCommitment.PONumber = txtPO.Text
    End Sub

    Private Function GetCommitmentTotals(ByVal IncludePayments As Boolean, ByVal excludeFed As Boolean) As Double
        Dim sReturn As Double
        Dim dtTotals As DataTable

        Try
            If Me.chkThirdPartyPayment.Checked Then
                dtTotals = oFinancialEvent.CommitmentTotalsDatatable(2, False, excludeFed)
            Else
                dtTotals = oFinancialEvent.CommitmentTotalsDatatable(1, False, excludeFed)
            End If
            If IsNothing(dtTotals) Then
                sReturn = cfActivityCosts.GrandTotal
            Else
                sReturn = cfActivityCosts.GrandTotal - sStartingTotal + CDbl(dtTotals.Rows(0)("EventCommitmentTotal")) + CDbl(dtTotals.Rows(0)("EventAdjustmentTotal"))
                If IncludePayments Then
                    sReturn -= CDbl(dtTotals.Rows(0)("EventPaymentTotal"))
                End If
            End If

            Return sReturn
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Save Calculation Failed " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try


    End Function


    Private Sub btnInsertCondition_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInsertCondition.Click
        Dim oFinText As New MUSTER.BusinessLogic.pFinancialText

        oFinText.Retrieve(cmbAdditionalConditions.SelectedValue)
        If oFinText.ID > 0 Then
            txtCondForReimbursement.Text &= vbCrLf & oFinText.Text
        End If

    End Sub


    Private Sub ProcessTechreports(ByVal row As Infragistics.Win.UltraWinGrid.UltraGridRow, ByVal newDate As DateTime)

        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow = row
        Dim oTechDox As New MUSTER.BusinessLogic.pLustEventDocument


        If oFinancialCommitment.CommitmentID > 0 Then

            oTechDox.Retrieve(ugrow.Cells("Event_Activity_Document_ID").Value)
            'TecDocument.Retrieve(oTechDox.DocClass)
            If newDate <= CDate("1/1/1900") Then
                oTechDox.DueDate = "01/01/0001"
                'oTechDox.DocClosedDate = "01/01/0001"
                oTechDox.CommitmentID = 0
            Else
                ' #2904
                If oTechDox.DueDate <> newDate Then
                    oTechDox.DueDate = newDate
                    oTechDox.CommitmentID = oFinancialCommitment.CommitmentID
                    'Need to generate the document.
                    ProcessDocumentsAndCalendarEntries(oTechDox)
                    changed = True
                End If
            End If
            If oTechDox.ID <= 0 Then
                oTechDox.CreatedBy = MusterContainer.AppUser.ID
            Else
                oTechDox.ModifiedBy = MusterContainer.AppUser.ID
            End If
            oTechDox.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            Else
                oTechDox.CloseParentDocument(oTechDox.AssocActivity, oTechDox.ID)
            End If
        End If


    End Sub

    Private Sub ProcessTechReports()
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim oTechDox As New MUSTER.BusinessLogic.pLustEventDocument
        For Each ugrow In ugTechReports.Rows
            oTechDox.Retrieve(ugrow.Cells("Event_Activity_Document_ID").Value)
            'TecDocument.Retrieve(oTechDox.DocClass)
            If IsDBNull(ugrow.Cells("Due_Date").Value) Then
                oTechDox.DueDate = "01/01/0001"
                'oTechDox.DocClosedDate = "01/01/0001"
                oTechDox.CommitmentID = 0
            Else
                ' #2904
                If ugrow.Cells("Date_Closed").Value Is DBNull.Value Then
                    oTechDox.DocClosedDate = CDate("01/01/0001")
                ElseIf ugrow.Cells("Date_Closed").Value = CDate("01/01/0001") Then
                    oTechDox.DocClosedDate = CDate("01/01/0001")
                Else
                    oTechDox.DocClosedDate = ugrow.Cells("Date_Closed").Value
                    oTechDox.DocClosedDate = oTechDox.DocClosedDate.Date
                End If
                If oTechDox.DueDate <> ugrow.Cells("Due_Date").Value Then
                    oTechDox.DueDate = ugrow.Cells("Due_Date").Value
                    oTechDox.DueDate = oTechDox.DueDate.Date
                    oTechDox.CommitmentID = oFinancialCommitment.CommitmentID
                    'Need to generate the document.
                    ProcessDocumentsAndCalendarEntries(oTechDox)
                End If
            End If
            If oTechDox.ID <= 0 Then
                oTechDox.CreatedBy = MusterContainer.AppUser.ID
            Else
                oTechDox.ModifiedBy = MusterContainer.AppUser.ID
            End If
            oTechDox.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            Else
                oTechDox.CloseParentDocument(oTechDox.AssocActivity, oTechDox.ID)
            End If
        Next
    End Sub
    Private Sub ProcessTecDocuments(ByVal nDocID As Integer, ByVal nAssocActivity As Integer, ByVal dtClosedDate As Date)
        Dim oTechDox As New MUSTER.BusinessLogic.pLustEventDocument
        Try
            oTechDox.Retrieve(nAssocActivity, nDocID)
            'oTechDox.Retrieve(oDocument.Auto_Doc_1)
            oTechDox.DocClosedDate = dtClosedDate
            If oTechDox.ID <= 0 Then
                oTechDox.CreatedBy = MusterContainer.AppUser.ID
            Else
                oTechDox.ModifiedBy = MusterContainer.AppUser.ID
            End If
            oTechDox.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub ProcessDocumentsAndCalendarEntries(ByVal oLocalLustDocument As MUSTER.BusinessLogic.pLustEventDocument)
        Dim nDocTrigger As Int64
        Dim nDocType As Integer
        Dim tmpDate As Date
        Dim dtNotificationDate As Date = Now()
        Dim nColorCode
        Dim strTaskDesc As String
        Dim strGroupID As String = ""
        Dim strUserID As String = ""
        Dim strSourceUserID As String = "SYSTEM"

        Dim oLocalLustActivity As New MUSTER.BusinessLogic.pLustEventActivity
        'Dim oLustRemediation As New MUSTER.BusinessLogic.pLustRemediation
        Dim oDocument As New MUSTER.BusinessLogic.pTecDoc
        Dim oCalendar As New MUSTER.BusinessLogic.pCalendar
        Dim oCalendarInfo As MUSTER.Info.CalendarInfo
        Dim oLustEvent As New MUSTER.BusinessLogic.pLustEvent
        Dim oFacility As New MUSTER.BusinessLogic.pFacility
        Dim oOwner As New MUSTER.BusinessLogic.pOwner
        Dim oUser As New MUSTER.BusinessLogic.pUser
        Dim nDocumentID As Integer

        Try
            oDocument.Retrieve(oLocalLustDocument.DocClass)
            oLustEvent.Retrieve(oLocalLustDocument.EventId)
            oFacility.Retrieve(oOwner.OwnerInfo, oLustEvent.FacilityID, , "FACILITY")
            oOwner.Retrieve(oFacility.OwnerID)
            oUser.Retrieve(oLustEvent.PM)

            nDocTrigger = oDocument.Trigger_Field
            nDocType = oDocument.DocType
            nDocumentID = oDocument.GetAutoCreatedDocumentParent(oLocalLustDocument.AssocActivity, oLocalLustDocument.DocClass)
            If Not nDocumentID = 0 Then
                ProcessTecDocuments(nDocumentID, oLocalLustDocument.AssocActivity, oLocalLustDocument.DocClosedDate)
            End If
            If nDocType = 917 Then
                oLocalLustDocument.DocumentID = CreateDocument(oLocalLustDocument, oLustEvent, oFacility, oOwner)
                'Set Tag in parent form to let it know a document was generated.
                'Me.CallingForm.Tag = 1

                '•	Create a Due To Me Calendar entry for the user on the Due date of the Document, unless it is a Task, 
                strTaskDesc = "ID : " & oLustEvent.FacilityID & " - Event: " & oLustEvent.EVENTSEQUENCE & " - " & oDocument.Name & " - New Due Date"
                oLocalLustDocument.MarkDueToMeCompleted_ByDesc(oLocalLustDocument.ID, strTaskDesc)

                strUserID = oUser.ID

                oCalendarInfo = New MUSTER.Info.CalendarInfo(0, _
                        dtNotificationDate, _
                        oLocalLustDocument.DueDate, _
                        nColorCode, _
                        strTaskDesc, _
                        strUserID, _
                        strSourceUserID, _
                        strGroupID, _
                        True, _
                        False, _
                        False, _
                        False, _
                        "sdfsdf", _
                        Now(), _
                        "asdf", _
                        Now())

                oCalendarInfo.OwningEntityID = oLocalLustDocument.ID
                oCalendarInfo.OwningEntityType = UIUtilsGen.EntityTypes.LustDocument
                oCalendarInfo.IsDirty = True
                oCalendar.Add(oCalendarInfo)
                oCalendar.Flush()
            End If

            If oLocalLustDocument.IsDirty Then
                oLocalLustDocument.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()

        End Try
    End Sub

    Private Sub SendToFinance(ByVal nDocID As Integer)

        Dim oTechDox As New MUSTER.BusinessLogic.pLustEventDocument
        Dim pm As New BusinessLogic.pProperty
        Dim lEvent As New BusinessLogic.pLustEvent
        Dim frmAddDocument As MUSTER.Document
        Dim newAct As New BusinessLogic.pLustEventActivity

        Try

            oTechDox.Retrieve(nDocID)

            If String.Format("{0} ", pm.GetPropertyNameByID(oTechDox.DocumentType)).IndexOf(" SOW/CE ") > -1 Or _
               String.Format("{0} ", pm.GetPropertyNameByID(oTechDox.DocumentType)).IndexOf(" SOW ") > -1 Then


                lEvent.Retrieve(oTechDox.EventId)


                frmAddDocument = New MUSTER.Document(lEvent)
                frmAddDocument.CallingForm = Me
                frmAddDocument.Mode = 1 ' Add
                frmAddDocument.EventActivityID = oTechDox.AssocActivity
                newAct.Retrieve(oTechDox.AssocActivity)

                frmAddDocument.EventDocumentID = oTechDox.ID
                frmAddDocument.TFStatus = lEvent.MGPTFStatus
                frmAddDocument.ActivityName = pm.GetPropertyNameByID(newAct.Type)

                frmAddDocument.StartDate = oTechDox.STARTDATE
                frmAddDocument.IsTech = False

                frmAddDocument.ShowDialog()

                frmAddDocument.Dispose()

                Me.LoadReportGrid()

            End If
        Catch ex As Exception
            Throw ex

        Finally
            pm = Nothing
            newAct = Nothing
            lEvent = Nothing
            frmAddDocument = Nothing

        End Try
    End Sub


    Private Function CreateDocument(ByVal oLocalLustDocument As MUSTER.BusinessLogic.pLustEventDocument, ByVal oLustEvent As MUSTER.BusinessLogic.pLustEvent, ByVal oFacility As MUSTER.BusinessLogic.pFacility, ByVal oOwner As MUSTER.BusinessLogic.pOwner) As Long
        Dim oLetter As New Reg_Letters
        Dim strShortName As String
        Dim strLongName As String
        Dim strTemplate As String
        Dim oDocument As New MUSTER.BusinessLogic.pTecDoc

        Try
            oDocument.Retrieve(oLocalLustDocument.DocClass)

            strTemplate = oDocument.FileName

            If strTemplate <> "" And strTemplate <> "Default" Then
                strLongName = oDocument.Name
                strShortName = strLongName
                strShortName = strShortName.Replace(" ", "")
                strShortName = strShortName.Replace("a", "")
                strShortName = strShortName.Replace("e", "")
                strShortName = strShortName.Replace("i", "")
                strShortName = strShortName.Replace("o", "")
                strShortName = strShortName.Replace("u", "")
                strShortName = strShortName.Replace("/", "")
                strShortName = strShortName.Replace("\", "")
                strShortName = strShortName.Replace(".", "")
                strShortName = strShortName.Replace("'", "")

                oLetter.GenerateTechLetter(oLustEvent.FacilityID, strLongName, Mid(strShortName, 1, 8), strLongName, strTemplate, oLocalLustDocument.DueDate, oLocalLustDocument.EventId, oOwner, 0, oLustEvent.EVENTSEQUENCE, UIUtilsGen.EntityTypes.LUST_Event)
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()

        End Try
    End Function



    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkReimburseERAC.CheckedChanged
        oFinancialCommitment.ReimburseERAC = chkReimburseERAC.Checked
    End Sub



    Private Sub BtnprintScren_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnprintScren.Click
        printOutCommitmentForm()
    End Sub

    Private Sub lblComments_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblComments.Click

    End Sub
End Class
