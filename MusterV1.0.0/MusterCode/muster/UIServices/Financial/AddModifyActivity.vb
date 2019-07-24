'Fixes up upgrades
'     2.0       Thomas Franey     02/6/09        Added two selections to select templates for the activity
'                                                Fix Close Button



Public Class AddModifyActivity
    Inherits System.Windows.Forms.Form
    Private WithEvents oFinActivity As New MUSTER.BusinessLogic.pFinancialActivity
    Private bolLoading As Boolean = True
    Private nMode As Int16 ' 0 = Add, 1 = Update
    Private coverMode As Integer = 0
    Private noticeMode As Integer = 0
    Dim returnVal As String = String.Empty

    Protected TmpltPath As String = MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_Templates).ProfileValue & "\Financial\"


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
    Friend WithEvents pnlActivityDetails As System.Windows.Forms.Panel
    Friend WithEvents lblActivity As System.Windows.Forms.Label
    Friend WithEvents chkActive As System.Windows.Forms.CheckBox
    Friend WithEvents txtActivity As System.Windows.Forms.TextBox
    Friend WithEvents cmbConditions As System.Windows.Forms.ComboBox
    Friend WithEvents lblConditions As System.Windows.Forms.Label
    Friend WithEvents txtAbbreviation As System.Windows.Forms.TextBox
    Friend WithEvents lblAbbreviation As System.Windows.Forms.Label
    Friend WithEvents txtDueDtStatement As System.Windows.Forms.TextBox
    Friend WithEvents lblDueDtStatement As System.Windows.Forms.Label
    Friend WithEvents lblCostFormatGroups As System.Windows.Forms.Label
    Friend WithEvents cmbActivity As System.Windows.Forms.ComboBox
    Friend WithEvents btnAddNew As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents lblFixedPrice As System.Windows.Forms.Label
    Friend WithEvents cmbCostPlus As System.Windows.Forms.ComboBox
    Friend WithEvents lblCostPlus As System.Windows.Forms.Label
    Friend WithEvents lblTM As System.Windows.Forms.Label
    Friend WithEvents cmbTM As System.Windows.Forms.ComboBox
    Friend WithEvents cmbFixedPrice As System.Windows.Forms.ComboBox
    Friend WithEvents LblReportTemplate As System.Windows.Forms.Label
    Friend WithEvents cmbActivityDocTemplate As System.Windows.Forms.ComboBox
    Friend WithEvents LblNoticetemplate As System.Windows.Forms.Label
    Friend WithEvents cboNoticeDocTemplate As System.Windows.Forms.ComboBox
    Friend WithEvents gBoxActDocRelationship As System.Windows.Forms.GroupBox
    Friend WithEvents BtnRemoveDoc As System.Windows.Forms.Button
    Friend WithEvents BtnAddNewDoc As System.Windows.Forms.Button
    Friend WithEvents BtnRemoveRpt As System.Windows.Forms.Button
    Friend WithEvents BtnAddNewRpt As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents PnlSelectTechDoc As System.Windows.Forms.Panel
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents LstTecDocuments As System.Windows.Forms.ListBox
    Friend WithEvents BtnCloseTecDocPanel As System.Windows.Forms.Button
    Friend WithEvents LstFinDocs As System.Windows.Forms.ListBox
    Friend WithEvents LstTechRpts As System.Windows.Forms.ListBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlActivityDetails = New System.Windows.Forms.Panel
        Me.PnlSelectTechDoc = New System.Windows.Forms.Panel
        Me.BtnCloseTecDocPanel = New System.Windows.Forms.Button
        Me.LstTecDocuments = New System.Windows.Forms.ListBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.gBoxActDocRelationship = New System.Windows.Forms.GroupBox
        Me.BtnRemoveDoc = New System.Windows.Forms.Button
        Me.BtnAddNewDoc = New System.Windows.Forms.Button
        Me.BtnRemoveRpt = New System.Windows.Forms.Button
        Me.BtnAddNewRpt = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.LstFinDocs = New System.Windows.Forms.ListBox
        Me.LstTechRpts = New System.Windows.Forms.ListBox
        Me.txtActivity = New System.Windows.Forms.TextBox
        Me.LblNoticetemplate = New System.Windows.Forms.Label
        Me.cboNoticeDocTemplate = New System.Windows.Forms.ComboBox
        Me.LblReportTemplate = New System.Windows.Forms.Label
        Me.cmbActivityDocTemplate = New System.Windows.Forms.ComboBox
        Me.lblFixedPrice = New System.Windows.Forms.Label
        Me.cmbCostPlus = New System.Windows.Forms.ComboBox
        Me.lblCostPlus = New System.Windows.Forms.Label
        Me.lblTM = New System.Windows.Forms.Label
        Me.cmbTM = New System.Windows.Forms.ComboBox
        Me.cmbFixedPrice = New System.Windows.Forms.ComboBox
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnAddNew = New System.Windows.Forms.Button
        Me.lblCostFormatGroups = New System.Windows.Forms.Label
        Me.lblDueDtStatement = New System.Windows.Forms.Label
        Me.txtDueDtStatement = New System.Windows.Forms.TextBox
        Me.lblAbbreviation = New System.Windows.Forms.Label
        Me.txtAbbreviation = New System.Windows.Forms.TextBox
        Me.lblConditions = New System.Windows.Forms.Label
        Me.cmbConditions = New System.Windows.Forms.ComboBox
        Me.chkActive = New System.Windows.Forms.CheckBox
        Me.lblActivity = New System.Windows.Forms.Label
        Me.cmbActivity = New System.Windows.Forms.ComboBox
        Me.pnlActivityDetails.SuspendLayout()
        Me.PnlSelectTechDoc.SuspendLayout()
        Me.gBoxActDocRelationship.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlActivityDetails
        '
        Me.pnlActivityDetails.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlActivityDetails.Controls.Add(Me.PnlSelectTechDoc)
        Me.pnlActivityDetails.Controls.Add(Me.gBoxActDocRelationship)
        Me.pnlActivityDetails.Controls.Add(Me.txtActivity)
        Me.pnlActivityDetails.Controls.Add(Me.LblNoticetemplate)
        Me.pnlActivityDetails.Controls.Add(Me.cboNoticeDocTemplate)
        Me.pnlActivityDetails.Controls.Add(Me.LblReportTemplate)
        Me.pnlActivityDetails.Controls.Add(Me.cmbActivityDocTemplate)
        Me.pnlActivityDetails.Controls.Add(Me.lblFixedPrice)
        Me.pnlActivityDetails.Controls.Add(Me.cmbCostPlus)
        Me.pnlActivityDetails.Controls.Add(Me.lblCostPlus)
        Me.pnlActivityDetails.Controls.Add(Me.lblTM)
        Me.pnlActivityDetails.Controls.Add(Me.cmbTM)
        Me.pnlActivityDetails.Controls.Add(Me.cmbFixedPrice)
        Me.pnlActivityDetails.Controls.Add(Me.btnDelete)
        Me.pnlActivityDetails.Controls.Add(Me.btnCancel)
        Me.pnlActivityDetails.Controls.Add(Me.btnSave)
        Me.pnlActivityDetails.Controls.Add(Me.btnAddNew)
        Me.pnlActivityDetails.Controls.Add(Me.lblCostFormatGroups)
        Me.pnlActivityDetails.Controls.Add(Me.lblDueDtStatement)
        Me.pnlActivityDetails.Controls.Add(Me.txtDueDtStatement)
        Me.pnlActivityDetails.Controls.Add(Me.lblAbbreviation)
        Me.pnlActivityDetails.Controls.Add(Me.txtAbbreviation)
        Me.pnlActivityDetails.Controls.Add(Me.lblConditions)
        Me.pnlActivityDetails.Controls.Add(Me.cmbConditions)
        Me.pnlActivityDetails.Controls.Add(Me.chkActive)
        Me.pnlActivityDetails.Controls.Add(Me.lblActivity)
        Me.pnlActivityDetails.Controls.Add(Me.cmbActivity)
        Me.pnlActivityDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlActivityDetails.Location = New System.Drawing.Point(0, 0)
        Me.pnlActivityDetails.Name = "pnlActivityDetails"
        Me.pnlActivityDetails.Size = New System.Drawing.Size(584, 517)
        Me.pnlActivityDetails.TabIndex = 0
        '
        'PnlSelectTechDoc
        '
        Me.PnlSelectTechDoc.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.PnlSelectTechDoc.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PnlSelectTechDoc.Controls.Add(Me.BtnCloseTecDocPanel)
        Me.PnlSelectTechDoc.Controls.Add(Me.LstTecDocuments)
        Me.PnlSelectTechDoc.Controls.Add(Me.Label3)
        Me.PnlSelectTechDoc.Location = New System.Drawing.Point(104, 224)
        Me.PnlSelectTechDoc.Name = "PnlSelectTechDoc"
        Me.PnlSelectTechDoc.Size = New System.Drawing.Size(312, 152)
        Me.PnlSelectTechDoc.TabIndex = 34
        Me.PnlSelectTechDoc.Visible = False
        '
        'BtnCloseTecDocPanel
        '
        Me.BtnCloseTecDocPanel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnCloseTecDocPanel.ForeColor = System.Drawing.Color.Firebrick
        Me.BtnCloseTecDocPanel.Location = New System.Drawing.Point(288, 0)
        Me.BtnCloseTecDocPanel.Name = "BtnCloseTecDocPanel"
        Me.BtnCloseTecDocPanel.Size = New System.Drawing.Size(24, 23)
        Me.BtnCloseTecDocPanel.TabIndex = 21
        Me.BtnCloseTecDocPanel.Text = "X"
        '
        'LstTecDocuments
        '
        Me.LstTecDocuments.Location = New System.Drawing.Point(8, 24)
        Me.LstTecDocuments.Name = "LstTecDocuments"
        Me.LstTecDocuments.Size = New System.Drawing.Size(296, 121)
        Me.LstTecDocuments.TabIndex = 20
        '
        'Label3
        '
        Me.Label3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(0, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(312, 24)
        Me.Label3.TabIndex = 19
        Me.Label3.Text = "Select a Technical Document"
        '
        'gBoxActDocRelationship
        '
        Me.gBoxActDocRelationship.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gBoxActDocRelationship.Controls.Add(Me.BtnRemoveDoc)
        Me.gBoxActDocRelationship.Controls.Add(Me.BtnAddNewDoc)
        Me.gBoxActDocRelationship.Controls.Add(Me.BtnRemoveRpt)
        Me.gBoxActDocRelationship.Controls.Add(Me.BtnAddNewRpt)
        Me.gBoxActDocRelationship.Controls.Add(Me.Label2)
        Me.gBoxActDocRelationship.Controls.Add(Me.Label1)
        Me.gBoxActDocRelationship.Controls.Add(Me.LstFinDocs)
        Me.gBoxActDocRelationship.Controls.Add(Me.LstTechRpts)
        Me.gBoxActDocRelationship.Enabled = False
        Me.gBoxActDocRelationship.Location = New System.Drawing.Point(16, 272)
        Me.gBoxActDocRelationship.Name = "gBoxActDocRelationship"
        Me.gBoxActDocRelationship.Size = New System.Drawing.Size(552, 192)
        Me.gBoxActDocRelationship.TabIndex = 33
        Me.gBoxActDocRelationship.TabStop = False
        Me.gBoxActDocRelationship.Text = "Technical Document Associations"
        '
        'BtnRemoveDoc
        '
        Me.BtnRemoveDoc.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnRemoveDoc.Location = New System.Drawing.Point(384, 160)
        Me.BtnRemoveDoc.Name = "BtnRemoveDoc"
        Me.BtnRemoveDoc.Size = New System.Drawing.Size(120, 24)
        Me.BtnRemoveDoc.TabIndex = 40
        Me.BtnRemoveDoc.Text = "Remove Doc"
        '
        'BtnAddNewDoc
        '
        Me.BtnAddNewDoc.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnAddNewDoc.Location = New System.Drawing.Point(264, 160)
        Me.BtnAddNewDoc.Name = "BtnAddNewDoc"
        Me.BtnAddNewDoc.Size = New System.Drawing.Size(120, 24)
        Me.BtnAddNewDoc.TabIndex = 39
        Me.BtnAddNewDoc.Text = "Add New Doc"
        '
        'BtnRemoveRpt
        '
        Me.BtnRemoveRpt.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.BtnRemoveRpt.Location = New System.Drawing.Point(128, 160)
        Me.BtnRemoveRpt.Name = "BtnRemoveRpt"
        Me.BtnRemoveRpt.Size = New System.Drawing.Size(120, 24)
        Me.BtnRemoveRpt.TabIndex = 38
        Me.BtnRemoveRpt.Text = "Remove Doc"
        '
        'BtnAddNewRpt
        '
        Me.BtnAddNewRpt.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.BtnAddNewRpt.Location = New System.Drawing.Point(8, 160)
        Me.BtnAddNewRpt.Name = "BtnAddNewRpt"
        Me.BtnAddNewRpt.Size = New System.Drawing.Size(120, 24)
        Me.BtnAddNewRpt.TabIndex = 37
        Me.BtnAddNewRpt.Text = "Add New Doc"
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.Location = New System.Drawing.Point(260, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(276, 17)
        Me.Label2.TabIndex = 36
        Me.Label2.Text = "Doc Available to Send to Finance  to Start Reports"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 24)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(232, 17)
        Me.Label1.TabIndex = 35
        Me.Label1.Text = "Technical Reports to Tie to Commitments"
        '
        'LstFinDocs
        '
        Me.LstFinDocs.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LstFinDocs.Location = New System.Drawing.Point(264, 48)
        Me.LstFinDocs.Name = "LstFinDocs"
        Me.LstFinDocs.Size = New System.Drawing.Size(272, 108)
        Me.LstFinDocs.TabIndex = 34
        '
        'LstTechRpts
        '
        Me.LstTechRpts.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LstTechRpts.Location = New System.Drawing.Point(8, 48)
        Me.LstTechRpts.Name = "LstTechRpts"
        Me.LstTechRpts.Size = New System.Drawing.Size(240, 108)
        Me.LstTechRpts.TabIndex = 33
        '
        'txtActivity
        '
        Me.txtActivity.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtActivity.Location = New System.Drawing.Point(80, 24)
        Me.txtActivity.Name = "txtActivity"
        Me.txtActivity.Size = New System.Drawing.Size(440, 20)
        Me.txtActivity.TabIndex = 1
        Me.txtActivity.Text = ""
        '
        'LblNoticetemplate
        '
        Me.LblNoticetemplate.Location = New System.Drawing.Point(8, 176)
        Me.LblNoticetemplate.Name = "LblNoticetemplate"
        Me.LblNoticetemplate.Size = New System.Drawing.Size(128, 28)
        Me.LblNoticetemplate.TabIndex = 24
        Me.LblNoticetemplate.Text = "Notice doc template :"
        Me.LblNoticetemplate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cboNoticeDocTemplate
        '
        Me.cboNoticeDocTemplate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboNoticeDocTemplate.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboNoticeDocTemplate.Location = New System.Drawing.Point(136, 184)
        Me.cboNoticeDocTemplate.Name = "cboNoticeDocTemplate"
        Me.cboNoticeDocTemplate.Size = New System.Drawing.Size(392, 21)
        Me.cboNoticeDocTemplate.TabIndex = 23
        '
        'LblReportTemplate
        '
        Me.LblReportTemplate.Location = New System.Drawing.Point(8, 152)
        Me.LblReportTemplate.Name = "LblReportTemplate"
        Me.LblReportTemplate.Size = New System.Drawing.Size(128, 28)
        Me.LblReportTemplate.TabIndex = 22
        Me.LblReportTemplate.Text = "Cover letter template :"
        Me.LblReportTemplate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbActivityDocTemplate
        '
        Me.cmbActivityDocTemplate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbActivityDocTemplate.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbActivityDocTemplate.Location = New System.Drawing.Point(136, 160)
        Me.cmbActivityDocTemplate.Name = "cmbActivityDocTemplate"
        Me.cmbActivityDocTemplate.Size = New System.Drawing.Size(392, 21)
        Me.cmbActivityDocTemplate.TabIndex = 21
        '
        'lblFixedPrice
        '
        Me.lblFixedPrice.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblFixedPrice.Location = New System.Drawing.Point(392, 240)
        Me.lblFixedPrice.Name = "lblFixedPrice"
        Me.lblFixedPrice.Size = New System.Drawing.Size(64, 17)
        Me.lblFixedPrice.TabIndex = 20
        Me.lblFixedPrice.Text = "Fixed Price:"
        '
        'cmbCostPlus
        '
        Me.cmbCostPlus.Location = New System.Drawing.Point(264, 240)
        Me.cmbCostPlus.Name = "cmbCostPlus"
        Me.cmbCostPlus.Size = New System.Drawing.Size(69, 21)
        Me.cmbCostPlus.TabIndex = 7
        '
        'lblCostPlus
        '
        Me.lblCostPlus.Location = New System.Drawing.Point(208, 240)
        Me.lblCostPlus.Name = "lblCostPlus"
        Me.lblCostPlus.Size = New System.Drawing.Size(56, 17)
        Me.lblCostPlus.TabIndex = 19
        Me.lblCostPlus.Text = "Cost Plus:"
        '
        'lblTM
        '
        Me.lblTM.Location = New System.Drawing.Point(32, 240)
        Me.lblTM.Name = "lblTM"
        Me.lblTM.Size = New System.Drawing.Size(40, 17)
        Me.lblTM.TabIndex = 18
        Me.lblTM.Text = "T && M:"
        '
        'cmbTM
        '
        Me.cmbTM.Location = New System.Drawing.Point(80, 240)
        Me.cmbTM.Name = "cmbTM"
        Me.cmbTM.Size = New System.Drawing.Size(69, 21)
        Me.cmbTM.TabIndex = 6
        '
        'cmbFixedPrice
        '
        Me.cmbFixedPrice.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbFixedPrice.Location = New System.Drawing.Point(456, 240)
        Me.cmbFixedPrice.Name = "cmbFixedPrice"
        Me.cmbFixedPrice.Size = New System.Drawing.Size(69, 21)
        Me.cmbFixedPrice.TabIndex = 8
        '
        'btnDelete
        '
        Me.btnDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnDelete.Location = New System.Drawing.Point(192, 488)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.TabIndex = 11
        Me.btnDelete.Text = "Delete"
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnCancel.Location = New System.Drawing.Point(104, 488)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 10
        Me.btnCancel.Text = "Cancel"
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Location = New System.Drawing.Point(16, 488)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.TabIndex = 9
        Me.btnSave.Text = "Save"
        '
        'btnAddNew
        '
        Me.btnAddNew.Location = New System.Drawing.Point(0, 0)
        Me.btnAddNew.Name = "btnAddNew"
        Me.btnAddNew.Size = New System.Drawing.Size(64, 16)
        Me.btnAddNew.TabIndex = 12
        Me.btnAddNew.Text = "Add New"
        '
        'lblCostFormatGroups
        '
        Me.lblCostFormatGroups.Location = New System.Drawing.Point(32, 216)
        Me.lblCostFormatGroups.Name = "lblCostFormatGroups"
        Me.lblCostFormatGroups.Size = New System.Drawing.Size(128, 17)
        Me.lblCostFormatGroups.TabIndex = 12
        Me.lblCostFormatGroups.Text = "Cost Format Groups:"
        '
        'lblDueDtStatement
        '
        Me.lblDueDtStatement.Location = New System.Drawing.Point(8, 104)
        Me.lblDueDtStatement.Name = "lblDueDtStatement"
        Me.lblDueDtStatement.Size = New System.Drawing.Size(128, 29)
        Me.lblDueDtStatement.TabIndex = 10
        Me.lblDueDtStatement.Text = "Due Date Statement : "
        Me.lblDueDtStatement.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtDueDtStatement
        '
        Me.txtDueDtStatement.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDueDtStatement.Location = New System.Drawing.Point(136, 104)
        Me.txtDueDtStatement.Multiline = True
        Me.txtDueDtStatement.Name = "txtDueDtStatement"
        Me.txtDueDtStatement.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtDueDtStatement.Size = New System.Drawing.Size(392, 52)
        Me.txtDueDtStatement.TabIndex = 5
        Me.txtDueDtStatement.Text = ""
        '
        'lblAbbreviation
        '
        Me.lblAbbreviation.Location = New System.Drawing.Point(48, 80)
        Me.lblAbbreviation.Name = "lblAbbreviation"
        Me.lblAbbreviation.Size = New System.Drawing.Size(88, 17)
        Me.lblAbbreviation.TabIndex = 8
        Me.lblAbbreviation.Text = "Abbreviation : "
        Me.lblAbbreviation.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtAbbreviation
        '
        Me.txtAbbreviation.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtAbbreviation.Location = New System.Drawing.Point(136, 80)
        Me.txtAbbreviation.Name = "txtAbbreviation"
        Me.txtAbbreviation.Size = New System.Drawing.Size(320, 20)
        Me.txtAbbreviation.TabIndex = 4
        Me.txtAbbreviation.Text = ""
        '
        'lblConditions
        '
        Me.lblConditions.Location = New System.Drawing.Point(8, 56)
        Me.lblConditions.Name = "lblConditions"
        Me.lblConditions.Size = New System.Drawing.Size(176, 16)
        Me.lblConditions.TabIndex = 6
        Me.lblConditions.Text = "Conditions for Reimbursement : "
        Me.lblConditions.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbConditions
        '
        Me.cmbConditions.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbConditions.Location = New System.Drawing.Point(184, 56)
        Me.cmbConditions.Name = "cmbConditions"
        Me.cmbConditions.Size = New System.Drawing.Size(272, 21)
        Me.cmbConditions.TabIndex = 2
        '
        'chkActive
        '
        Me.chkActive.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkActive.Location = New System.Drawing.Point(472, 56)
        Me.chkActive.Name = "chkActive"
        Me.chkActive.Size = New System.Drawing.Size(54, 24)
        Me.chkActive.TabIndex = 3
        Me.chkActive.Text = "Active"
        '
        'lblActivity
        '
        Me.lblActivity.Location = New System.Drawing.Point(16, 24)
        Me.lblActivity.Name = "lblActivity"
        Me.lblActivity.Size = New System.Drawing.Size(64, 21)
        Me.lblActivity.TabIndex = 1
        Me.lblActivity.Text = "Activity : "
        Me.lblActivity.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbActivity
        '
        Me.cmbActivity.Location = New System.Drawing.Point(112, 24)
        Me.cmbActivity.Name = "cmbActivity"
        Me.cmbActivity.Size = New System.Drawing.Size(339, 21)
        Me.cmbActivity.TabIndex = 13
        Me.cmbActivity.Text = "ComboBox1"
        '
        'AddModifyActivity
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(584, 517)
        Me.Controls.Add(Me.pnlActivityDetails)
        Me.Name = "AddModifyActivity"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Add/Modify Activity"
        Me.pnlActivityDetails.ResumeLayout(False)
        Me.PnlSelectTechDoc.ResumeLayout(False)
        Me.gBoxActDocRelationship.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Page Events "
    Private Sub AddModifyActivity_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadForm()
    End Sub
    Private Sub txtActivity_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtActivity.TextChanged
        If bolLoading Then Exit Sub
        oFinActivity.ActivityDesc = txtActivity.Text
    End Sub

    Private Sub cmbConditions_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbConditions.SelectedIndexChanged
        If bolLoading Then Exit Sub
        oFinActivity.ReimbursementCondition = cmbConditions.SelectedValue
    End Sub


    Private Sub txtAbbreviation_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAbbreviation.TextChanged
        If bolLoading Then Exit Sub
        oFinActivity.ActivityDescShort = txtAbbreviation.Text
    End Sub

    Private Sub txtDueDtStatement_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDueDtStatement.TextChanged
        If bolLoading Then Exit Sub
        oFinActivity.DueDateStatement = txtDueDtStatement.Text
    End Sub

    Private Sub chkActive_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkActive.CheckedChanged
        If bolLoading Then Exit Sub
        oFinActivity.Active = chkActive.Checked
    End Sub
    Private Sub cmbTM_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbTM.SelectedIndexChanged
        If bolLoading Then Exit Sub
        oFinActivity.TimeAndMaterials = cmbTM.SelectedValue
        oFinActivity.TimeAndMaterialsDesc = cmbTM.Text
    End Sub


    Private Sub BtnCloseTecDocPanel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnCloseTecDocPanel.Click
        PnlSelectTechDoc.Visible = False
        LstTecDocuments.DataSource = Nothing
        gBoxActDocRelationship.Enabled = True
    End Sub

    Private Sub cmbCostPlus_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCostPlus.SelectedIndexChanged
        If bolLoading Then Exit Sub
        oFinActivity.CostPlus = cmbCostPlus.SelectedValue
        oFinActivity.CostPlusDesc = cmbCostPlus.Text
    End Sub

    Private Sub cmbCoverDoc_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbActivityDocTemplate.SelectedIndexChanged

        If bolLoading Then Exit Sub
        oFinActivity.CoverTemplate = Me.cmbActivityDocTemplate.SelectedItem

    End Sub


    Private Sub cmbNoticeDoc_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboNoticeDocTemplate.SelectedIndexChanged

        If bolLoading Then Exit Sub
        oFinActivity.NoticeTemplate = Me.cboNoticeDocTemplate.SelectedItem

    End Sub


    Private Sub cmbFixedPrice_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbFixedPrice.SelectedIndexChanged
        If bolLoading Then Exit Sub
        oFinActivity.FixedPrice = cmbFixedPrice.SelectedValue
        oFinActivity.FixedPriceDesc = cmbFixedPrice.Text
    End Sub
    Private Sub cmbActivity_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbActivity.SelectedIndexChanged
        If bolLoading Then Exit Sub
        SetActivityForm()
    End Sub
    Private Sub btnAddNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddNew.Click
        nMode = 0
        SetActivityForm()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click

        Dim SaveMainModule As Boolean = True
        Dim message As String = String.Empty

        If txtActivity.Text = String.Empty Then
            message = "Please enter Activity name."
            SaveMainModule = False
        End If

        If cmbConditions.SelectedIndex = -1 Then
            message = "Please select a Condition."
            SaveMainModule = False
        End If

        If txtAbbreviation.Text = String.Empty Then
            message = "Please enter Abbreviation."
            SaveMainModule = False
        End If

        If txtDueDtStatement.Text = String.Empty Then
            message = "Please enter Statement Due Date."
            SaveMainModule = False
        End If

        If cmbTM.SelectedIndex <= 0 Then
            message = "Please select Time And Materials."
            SaveMainModule = False
        End If

        If cmbCostPlus.SelectedIndex <= 0 Then
            message = "Please select Cost Plus."
            SaveMainModule = False
        End If

        If cmbFixedPrice.SelectedIndex <= 0 Then
            message = "Please select Fixed Price."
            SaveMainModule = False
        End If

        If cmbActivityDocTemplate.SelectedIndex <= IIf(coverMode = 0, 0, -1) Then
            message = "Please select your template cover letter."
            SaveMainModule = False
        End If

        If cboNoticeDocTemplate.SelectedIndex <= IIf(noticeMode = 0, 0, -1) Then
            message = "Please select your template notice doc."
            SaveMainModule = False
        End If

        If SaveMainModule Then
            message = "Activity Saved."
        End If


        Try

            If SaveMainModule Then


                If oFinActivity.Activity_ID <= 0 Then
                    oFinActivity.CreatedBy = MusterContainer.AppUser.ID
                Else
                    oFinActivity.ModifiedBy = MusterContainer.AppUser.ID
                End If

                oFinActivity.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

            End If


            If oFinActivity.Activity_ID > 0 Then


                'Clean Relationships
                oFinActivity.ClearACtivityDocRelationshipByID(oFinActivity.Activity_ID, CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)

                'Set New Relationships
                For Each row As DataRowView In DirectCast(Me.LstFinDocs.DataSource, DataView)
                    oFinActivity.PutFinActivityTechDocRelationship(oFinActivity.Activity_ID, row.Row.Item("TEC_DOC_TYPE_ID"), 1, _
                                                                   CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)
                Next

                For Each row As DataRowView In DirectCast(Me.LstTechRpts.DataSource, DataView)
                    oFinActivity.PutFinActivityTechDocRelationship(oFinActivity.Activity_ID, row.Row.Item("TEC_DOC_TYPE_ID"), 0, _
                                                                   CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)
                Next



                MsgBox(String.Format("{0}{2}{1}", message, " Financial/Tech Doc relationships will be saved", vbCrLf))
            Else
                MsgBox(String.Format("{0}{1}", message))

            End If

        Catch ex As Exception
            MsgBox(message)
        Finally

            nMode = 1
            If SaveMainModule Then
                bolLoading = True
                LoadDropDowns()
                bolLoading = False
                SetActivityForm()
            End If

        End Try

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        nMode = 1
        SetActivityForm()
        Me.Close()

    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click

        If MsgBox("Are You Sure You Want To Delete This Activity.", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            oFinActivity.Deleted = True
            oFinActivity.ModifiedBy = MusterContainer.AppUser.ID

            oFinActivity.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            bolLoading = True
            nMode = 1
            LoadDropDowns()
            bolLoading = False
            SetActivityForm()
        End If

    End Sub

#End Region

#Region " Populate Routines "
    Private Sub LoadForm()
        bolLoading = True
        LoadDropDowns()
        nMode = 1
        'cmbActivity.SelectedIndex = 0
        bolLoading = False
        SetActivityForm()
    End Sub

    Private Sub LoadTechDocs()

        Dim dtTechDocsToSelect As DataTable = oFinActivity.PopulateAllTechnicalDocsForActivity()

        Try

            With Me.LstTecDocuments

                .DataSource = dtTechDocsToSelect
                .DisplayMember = "Property_Name"
                .ValueMember = "DocID"
            End With

            gBoxActDocRelationship.Enabled = False
            PnlSelectTechDoc.Visible = True



        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Load technical Doucment selection for Activity " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try


    End Sub

    Private Sub LoadLists()
        Dim dtFinTechDocs As DataTable = oFinActivity.PopulateDocumentsForActivity
        Dim dtFinTechDocs2 As DataTable = oFinActivity.PopulateDocumentsForActivity

        dtFinTechDocs.DefaultView.RowFilter = "IsSentToFinancialDoc = 1"
        dtFinTechDocs2.DefaultView.RowFilter = "IsSentToFinancialDoc = 0"


        Try


            Me.lstFinDocs.DataSource = dtFinTechDocs.DefaultView
            Me.lstFinDocs.DisplayMember = "Document"
            Me.lstFinDocs.ValueMember = "TEC_DOC_TYPE_ID"

            Me.lstTechRpts.DataSource = dtFinTechDocs2.DefaultView
            Me.lstTechRpts.DisplayMember = "Document"
            Me.lstTechRpts.ValueMember = "TEC_DOC_TYPE_ID"




        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Load Activity Document Lists " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try


    End Sub
    Private Sub LoadDropDowns()

        Dim dtCostFormat1 As DataTable = oFinActivity.PopulateCostFormats
        Dim dtCostFormat2 As DataTable = oFinActivity.PopulateCostFormats
        Dim dtCostFormat3 As DataTable = oFinActivity.PopulateCostFormats
        Dim dtReimbursement As DataTable = oFinActivity.PopulateReimbursementConditions
        Dim dtActivities As DataTable = oFinActivity.PopulateFinancialActivityList


        Try

            cmbConditions.DataSource = dtReimbursement
            cmbConditions.DisplayMember = "Text_Name"
            cmbConditions.ValueMember = "Text_ID"


            cmbActivity.DataSource = dtActivities
            cmbActivity.DisplayMember = "ActivityDesc"
            cmbActivity.ValueMember = "Activity_ID"

            cmbCostPlus.DataSource = dtCostFormat1
            cmbCostPlus.DisplayMember = "PROPERTY_NAME"
            cmbCostPlus.ValueMember = "PROPERTY_ID"

            cmbFixedPrice.DataSource = dtCostFormat2
            cmbFixedPrice.DisplayMember = "PROPERTY_NAME"
            cmbFixedPrice.ValueMember = "PROPERTY_ID"

            cmbTM.DataSource = dtCostFormat3
            cmbTM.DisplayMember = "PROPERTY_NAME"
            cmbTM.ValueMember = "PROPERTY_ID"

            coverMode = 1
            noticeMode = 1

            If Me.oFinActivity.CoverTemplate = String.Empty Then
                cmbActivityDocTemplate.Items.Add("[Select your Cover Letter Template]")
                coverMode = 0
            End If

            If Me.oFinActivity.NoticeTemplate = String.Empty Then
                cboNoticeDocTemplate.Items.Add("[Select your Notice Letter Template]")
                noticeMode = 0
            End If

            Dim doc As String = Dir(String.Format("{0}*.doc", TmpltPath))

            While Not doc Is Nothing AndAlso doc.Length > 0

                cmbActivityDocTemplate.Items.Add(doc)
                cboNoticeDocTemplate.Items.Add(doc)

                doc = Dir()

            End While



        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Load System DropDowns " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub



#End Region

#Region " General Routines "

    Private Sub SetActivityForm()
        bolLoading = True
        If nMode = 0 Then
            btnDelete.Visible = False
            oFinActivity.Retrieve(0)
            cmbCostPlus.SelectedValue = 0
            cmbFixedPrice.SelectedValue = 0
            cmbTM.SelectedValue = 0
            cmbActivity.Visible = False



            txtActivity.Visible = True
            cmbConditions.SelectedIndex = 0
            txtAbbreviation.Text = ""
            txtActivity.Text = ""
            txtDueDtStatement.Text = ""
            chkActive.Checked = True
            cmbActivityDocTemplate.SelectedIndex = 0
            cboNoticeDocTemplate.SelectedIndex = 0

            oFinActivity.Active = True
            oFinActivity.ReimbursementCondition = cmbConditions.SelectedValue

            Me.Text = "Add New Activity"

            Me.gBoxActDocRelationship.Enabled = False

            LoadLists()

        Else



            oFinActivity.Retrieve(cmbActivity.SelectedValue)

            With oFinActivity

                If .IsUsed Then
                    btnDelete.Visible = False
                Else
                    btnDelete.Visible = True
                End If
                cmbActivity.Visible = True
                txtActivity.Visible = False
                cmbActivity.SelectedValue = .Activity_ID

                cmbCostPlus.SelectedValue = .CostPlus
                cmbFixedPrice.SelectedValue = .FixedPrice
                cmbTM.SelectedValue = .TimeAndMaterials
                cmbConditions.SelectedValue = .ReimbursementCondition
                txtAbbreviation.Text = .ActivityDescShort
                txtActivity.Text = .ActivityDesc
                txtDueDtStatement.Text = .DueDateStatement
                chkActive.Checked = .Active
                If Me.cmbActivityDocTemplate.Items.Contains(.CoverTemplate) Then
                    cmbActivityDocTemplate.SelectedItem = .CoverTemplate()
                Else
                    cmbActivityDocTemplate.SelectedIndex = 0

                End If

                If Me.cboNoticeDocTemplate.Items.Contains(.NoticeTemplate) Then
                    cboNoticeDocTemplate.SelectedItem = .NoticeTemplate
                Else
                    cboNoticeDocTemplate.SelectedIndex = 0

                End If


                Me.Text = String.Format("Modify Activity {0} : {1}", .Activity_ID, .ActivityDesc)

            End With

            Me.gBoxActDocRelationship.Enabled = True

            LoadLists()


        End If
        bolLoading = False
    End Sub

#End Region










    Private Sub LstTecDocuments_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LstTecDocuments.Click

        If Not Me.LstTecDocuments.SelectedItem Is Nothing AndAlso Not Me.LstTecDocuments.SelectedValue Is Nothing Then
            Dim listGrid As DataTable
            Dim listGridCOntrol As ListControl

            If Me.PnlSelectTechDoc.Tag = 1 Then
                listGrid = DirectCast(Me.lstFinDocs.DataSource, DataView).Table
                listGridCOntrol = Me.lstFinDocs
            Else
                listGrid = DirectCast(Me.lstTechRpts.DataSource, DataView).Table
                listGridCOntrol = Me.lstTechRpts
            End If

            Dim newRow As DataRow = listGrid.NewRow

            newRow.Item("Document") = DirectCast(LstTecDocuments.SelectedItem, DataRowView).Row.Item(1)
            newRow.Item("TEC_DOC_TYPE_ID") = Me.LstTecDocuments.SelectedValue
            newRow.Item("IsSentToFinancialDoc") = IIf(Me.PnlSelectTechDoc.Tag = 1, True, False)



            gBoxActDocRelationship.Enabled = True

            listGrid.Rows.Add(newRow)
            listGridCOntrol.DataSource = listGrid.DefaultView

            gBoxActDocRelationship.Enabled = False


            'BtnCloseTecDocPanel.PerformClick()
        End If


    End Sub

    Private Sub BtnAddNewDoc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAddNewDoc.Click
        PnlSelectTechDoc.Tag = 1
        LoadTechDocs()
    End Sub

    Private Sub BtnAddNewRpt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAddNewRpt.Click
        PnlSelectTechDoc.Tag = 0
        LoadTechDocs()
    End Sub

    Private Sub BtnRemoveRpt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnRemoveRpt.Click
        If Not Me.lstTechRpts.SelectedItem Is Nothing Then
            DirectCast(Me.lstTechRpts.DataSource, DataView).Table.Rows.Remove(DirectCast(Me.lstTechRpts.SelectedItem, DataRowView).Row)
        End If
    End Sub

    Private Sub BtnRemoveDoc_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnRemoveDoc.Click
        If Not Me.lstFinDocs.SelectedItem Is Nothing Then
            DirectCast(Me.lstFinDocs.DataSource, DataView).Table.Rows.Remove(DirectCast(Me.lstFinDocs.SelectedItem, DataRowView).Row)
        End If
    End Sub

End Class
