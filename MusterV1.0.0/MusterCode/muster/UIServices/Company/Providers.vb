Public Class Providers
    Inherits System.Windows.Forms.Form
#Region "User Defined Variables"
    'Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long
    Private WithEvents oProvider As New MUSTER.BusinessLogic.pProvider
    Private WithEvents oComAdd As New MUSTER.BusinessLogic.pComAddress
    Private oComAddInfo As MUSTER.Info.ComAddressInfo
    Friend WithEvents objAddMaster As AddressMaster
    Private oProvInfo As MUSTER.Info.ProviderInfo
    Private bolValidationFlg As Boolean = False
    Private nProviderID As Integer = 0
    Private bolLoading As Boolean = False
    Private bolNewProvider As Boolean = False
    Private nAddressID As Integer = 0
    Dim strAddress As String = ""
    Dim returnVal As String = String.Empty
#End Region
#Region " Windows Form Designer generated code "

    Public Sub New(Optional ByRef ParentForm As Windows.Forms.Form = Nothing)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        If Not ParentForm Is Nothing Then
            Me.MdiParent = ParentForm
        End If
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
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents txtDepartment As System.Windows.Forms.TextBox
    Friend WithEvents lblDepartment As System.Windows.Forms.Label
    Friend WithEvents txtAbbreviation As System.Windows.Forms.TextBox
    Friend WithEvents lblAbbreviation As System.Windows.Forms.Label
    Friend WithEvents txtProvider As System.Windows.Forms.TextBox
    Friend WithEvents lblProvider As System.Windows.Forms.Label
    Friend WithEvents txtProviderID As System.Windows.Forms.TextBox
    Friend WithEvents lblProviderID As System.Windows.Forms.Label
    Friend WithEvents txtWebsite As System.Windows.Forms.TextBox
    Friend WithEvents lblWebsite As System.Windows.Forms.Label
    Friend WithEvents chkActive As System.Windows.Forms.CheckBox
    Friend WithEvents lblAddress As System.Windows.Forms.Label
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents cmbProviderName As System.Windows.Forms.ComboBox
    Friend WithEvents txtProviderAddress As System.Windows.Forms.TextBox
    Friend WithEvents pnlExtension As System.Windows.Forms.Panel
    Friend WithEvents lblExt2 As System.Windows.Forms.Label
    Friend WithEvents lblExt1 As System.Windows.Forms.Label
    Friend WithEvents txtExt2 As System.Windows.Forms.TextBox
    Friend WithEvents txtExt1 As System.Windows.Forms.TextBox
    Friend WithEvents pnlPhone As System.Windows.Forms.Panel
    Public WithEvents mskTxtPhone1 As AxMSMask.AxMaskEdBox
    Friend WithEvents lblPhone2 As System.Windows.Forms.Label
    Public WithEvents mskTxtCell As AxMSMask.AxMaskEdBox
    Friend WithEvents lblCell As System.Windows.Forms.Label
    Public WithEvents mskTxtPhone2 As AxMSMask.AxMaskEdBox
    Friend WithEvents lblPhone1 As System.Windows.Forms.Label
    Public WithEvents mskTxtFax As AxMSMask.AxMaskEdBox
    Friend WithEvents lblFax As System.Windows.Forms.Label
    Friend WithEvents pnlPhoneComment As System.Windows.Forms.Panel
    Friend WithEvents lblPhone2Comment As System.Windows.Forms.Label
    Friend WithEvents lblPhone1Comment As System.Windows.Forms.Label
    Friend WithEvents txtPhone2Comment As System.Windows.Forms.TextBox
    Friend WithEvents txtPhone1Comment As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Providers))
        Me.txtDepartment = New System.Windows.Forms.TextBox
        Me.lblDepartment = New System.Windows.Forms.Label
        Me.txtAbbreviation = New System.Windows.Forms.TextBox
        Me.lblAbbreviation = New System.Windows.Forms.Label
        Me.txtProvider = New System.Windows.Forms.TextBox
        Me.lblProvider = New System.Windows.Forms.Label
        Me.txtProviderID = New System.Windows.Forms.TextBox
        Me.lblProviderID = New System.Windows.Forms.Label
        Me.txtWebsite = New System.Windows.Forms.TextBox
        Me.lblWebsite = New System.Windows.Forms.Label
        Me.chkActive = New System.Windows.Forms.CheckBox
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.txtProviderAddress = New System.Windows.Forms.TextBox
        Me.lblAddress = New System.Windows.Forms.Label
        Me.btnClose = New System.Windows.Forms.Button
        Me.cmbProviderName = New System.Windows.Forms.ComboBox
        Me.pnlExtension = New System.Windows.Forms.Panel
        Me.lblExt2 = New System.Windows.Forms.Label
        Me.lblExt1 = New System.Windows.Forms.Label
        Me.txtExt2 = New System.Windows.Forms.TextBox
        Me.txtExt1 = New System.Windows.Forms.TextBox
        Me.pnlPhone = New System.Windows.Forms.Panel
        Me.mskTxtPhone1 = New AxMSMask.AxMaskEdBox
        Me.lblPhone2 = New System.Windows.Forms.Label
        Me.mskTxtCell = New AxMSMask.AxMaskEdBox
        Me.lblCell = New System.Windows.Forms.Label
        Me.mskTxtPhone2 = New AxMSMask.AxMaskEdBox
        Me.lblPhone1 = New System.Windows.Forms.Label
        Me.mskTxtFax = New AxMSMask.AxMaskEdBox
        Me.lblFax = New System.Windows.Forms.Label
        Me.pnlPhoneComment = New System.Windows.Forms.Panel
        Me.lblPhone2Comment = New System.Windows.Forms.Label
        Me.lblPhone1Comment = New System.Windows.Forms.Label
        Me.txtPhone2Comment = New System.Windows.Forms.TextBox
        Me.txtPhone1Comment = New System.Windows.Forms.TextBox
        Me.pnlExtension.SuspendLayout()
        Me.pnlPhone.SuspendLayout()
        CType(Me.mskTxtPhone1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtCell, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtPhone2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtFax, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlPhoneComment.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtDepartment
        '
        Me.txtDepartment.Location = New System.Drawing.Point(128, 112)
        Me.txtDepartment.Name = "txtDepartment"
        Me.txtDepartment.Size = New System.Drawing.Size(208, 20)
        Me.txtDepartment.TabIndex = 4
        Me.txtDepartment.Text = ""
        '
        'lblDepartment
        '
        Me.lblDepartment.Location = New System.Drawing.Point(48, 112)
        Me.lblDepartment.Name = "lblDepartment"
        Me.lblDepartment.Size = New System.Drawing.Size(72, 17)
        Me.lblDepartment.TabIndex = 193
        Me.lblDepartment.Text = "Department:"
        Me.lblDepartment.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtAbbreviation
        '
        Me.txtAbbreviation.Location = New System.Drawing.Point(128, 88)
        Me.txtAbbreviation.Name = "txtAbbreviation"
        Me.txtAbbreviation.Size = New System.Drawing.Size(128, 20)
        Me.txtAbbreviation.TabIndex = 3
        Me.txtAbbreviation.Text = ""
        '
        'lblAbbreviation
        '
        Me.lblAbbreviation.Location = New System.Drawing.Point(40, 88)
        Me.lblAbbreviation.Name = "lblAbbreviation"
        Me.lblAbbreviation.Size = New System.Drawing.Size(80, 17)
        Me.lblAbbreviation.TabIndex = 211
        Me.lblAbbreviation.Text = "Abbreviation:"
        Me.lblAbbreviation.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtProvider
        '
        Me.txtProvider.Location = New System.Drawing.Point(128, 64)
        Me.txtProvider.Name = "txtProvider"
        Me.txtProvider.Size = New System.Drawing.Size(304, 20)
        Me.txtProvider.TabIndex = 1
        Me.txtProvider.Text = ""
        '
        'lblProvider
        '
        Me.lblProvider.Location = New System.Drawing.Point(56, 48)
        Me.lblProvider.Name = "lblProvider"
        Me.lblProvider.Size = New System.Drawing.Size(64, 32)
        Me.lblProvider.TabIndex = 213
        Me.lblProvider.Text = "Provider Name"
        Me.lblProvider.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtProviderID
        '
        Me.txtProviderID.Location = New System.Drawing.Point(128, 16)
        Me.txtProviderID.Name = "txtProviderID"
        Me.txtProviderID.ReadOnly = True
        Me.txtProviderID.Size = New System.Drawing.Size(104, 20)
        Me.txtProviderID.TabIndex = 1001
        Me.txtProviderID.Text = ""
        '
        'lblProviderID
        '
        Me.lblProviderID.Location = New System.Drawing.Point(48, 16)
        Me.lblProviderID.Name = "lblProviderID"
        Me.lblProviderID.Size = New System.Drawing.Size(72, 17)
        Me.lblProviderID.TabIndex = 215
        Me.lblProviderID.Text = "ProviderID:"
        Me.lblProviderID.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtWebsite
        '
        Me.txtWebsite.Location = New System.Drawing.Point(128, 240)
        Me.txtWebsite.Name = "txtWebsite"
        Me.txtWebsite.Size = New System.Drawing.Size(208, 20)
        Me.txtWebsite.TabIndex = 6
        Me.txtWebsite.Text = ""
        '
        'lblWebsite
        '
        Me.lblWebsite.Location = New System.Drawing.Point(64, 240)
        Me.lblWebsite.Name = "lblWebsite"
        Me.lblWebsite.Size = New System.Drawing.Size(56, 17)
        Me.lblWebsite.TabIndex = 225
        Me.lblWebsite.Text = "Website:"
        Me.lblWebsite.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'chkActive
        '
        Me.chkActive.Checked = True
        Me.chkActive.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkActive.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkActive.Location = New System.Drawing.Point(456, 16)
        Me.chkActive.Name = "chkActive"
        Me.chkActive.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkActive.Size = New System.Drawing.Size(72, 16)
        Me.chkActive.TabIndex = 7
        Me.chkActive.Tag = "644"
        Me.chkActive.Text = " :Active"
        Me.chkActive.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(328, 288)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 11
        Me.btnDelete.Text = "Delete"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(240, 288)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(80, 26)
        Me.btnCancel.TabIndex = 10
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.Visible = False
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(152, 288)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(80, 26)
        Me.btnSave.TabIndex = 9
        Me.btnSave.Text = "Save"
        '
        'btnNew
        '
        Me.btnNew.Location = New System.Drawing.Point(64, 288)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 8
        Me.btnNew.Text = "New"
        '
        'txtProviderAddress
        '
        Me.txtProviderAddress.Location = New System.Drawing.Point(128, 136)
        Me.txtProviderAddress.Multiline = True
        Me.txtProviderAddress.Name = "txtProviderAddress"
        Me.txtProviderAddress.ReadOnly = True
        Me.txtProviderAddress.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
        Me.txtProviderAddress.Size = New System.Drawing.Size(208, 96)
        Me.txtProviderAddress.TabIndex = 5
        Me.txtProviderAddress.Text = ""
        Me.txtProviderAddress.WordWrap = False
        '
        'lblAddress
        '
        Me.lblAddress.Location = New System.Drawing.Point(56, 136)
        Me.lblAddress.Name = "lblAddress"
        Me.lblAddress.Size = New System.Drawing.Size(64, 16)
        Me.lblAddress.TabIndex = 266
        Me.lblAddress.Text = "Address:"
        Me.lblAddress.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(416, 288)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(80, 26)
        Me.btnClose.TabIndex = 12
        Me.btnClose.Text = "Close"
        '
        'cmbProviderName
        '
        Me.cmbProviderName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbProviderName.Location = New System.Drawing.Point(128, 40)
        Me.cmbProviderName.Name = "cmbProviderName"
        Me.cmbProviderName.Size = New System.Drawing.Size(304, 21)
        Me.cmbProviderName.TabIndex = 0
        '
        'pnlExtension
        '
        Me.pnlExtension.Controls.Add(Me.lblExt2)
        Me.pnlExtension.Controls.Add(Me.lblExt1)
        Me.pnlExtension.Controls.Add(Me.txtExt2)
        Me.pnlExtension.Controls.Add(Me.txtExt1)
        Me.pnlExtension.Location = New System.Drawing.Point(560, 112)
        Me.pnlExtension.Name = "pnlExtension"
        Me.pnlExtension.Size = New System.Drawing.Size(75, 56)
        Me.pnlExtension.TabIndex = 1003
        '
        'lblExt2
        '
        Me.lblExt2.Location = New System.Drawing.Point(0, 30)
        Me.lblExt2.Name = "lblExt2"
        Me.lblExt2.Size = New System.Drawing.Size(33, 16)
        Me.lblExt2.TabIndex = 236
        Me.lblExt2.Text = "Ext 2:"
        Me.lblExt2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblExt1
        '
        Me.lblExt1.Location = New System.Drawing.Point(0, 6)
        Me.lblExt1.Name = "lblExt1"
        Me.lblExt1.Size = New System.Drawing.Size(33, 16)
        Me.lblExt1.TabIndex = 235
        Me.lblExt1.Text = "Ext 1:"
        Me.lblExt1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtExt2
        '
        Me.txtExt2.Location = New System.Drawing.Point(32, 29)
        Me.txtExt2.Name = "txtExt2"
        Me.txtExt2.Size = New System.Drawing.Size(40, 20)
        Me.txtExt2.TabIndex = 234
        Me.txtExt2.Text = ""
        '
        'txtExt1
        '
        Me.txtExt1.Location = New System.Drawing.Point(32, 6)
        Me.txtExt1.Name = "txtExt1"
        Me.txtExt1.Size = New System.Drawing.Size(40, 20)
        Me.txtExt1.TabIndex = 233
        Me.txtExt1.Text = ""
        '
        'pnlPhone
        '
        Me.pnlPhone.Controls.Add(Me.mskTxtPhone1)
        Me.pnlPhone.Controls.Add(Me.lblPhone2)
        Me.pnlPhone.Controls.Add(Me.mskTxtCell)
        Me.pnlPhone.Controls.Add(Me.lblCell)
        Me.pnlPhone.Controls.Add(Me.mskTxtPhone2)
        Me.pnlPhone.Controls.Add(Me.lblPhone1)
        Me.pnlPhone.Controls.Add(Me.mskTxtFax)
        Me.pnlPhone.Controls.Add(Me.lblFax)
        Me.pnlPhone.Location = New System.Drawing.Point(344, 112)
        Me.pnlPhone.Name = "pnlPhone"
        Me.pnlPhone.Size = New System.Drawing.Size(216, 104)
        Me.pnlPhone.TabIndex = 1002
        '
        'mskTxtPhone1
        '
        Me.mskTxtPhone1.ContainingControl = Me
        Me.mskTxtPhone1.Location = New System.Drawing.Point(72, 6)
        Me.mskTxtPhone1.Name = "mskTxtPhone1"
        Me.mskTxtPhone1.OcxState = CType(resources.GetObject("mskTxtPhone1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtPhone1.Size = New System.Drawing.Size(142, 20)
        Me.mskTxtPhone1.TabIndex = 13
        '
        'lblPhone2
        '
        Me.lblPhone2.Location = New System.Drawing.Point(10, 30)
        Me.lblPhone2.Name = "lblPhone2"
        Me.lblPhone2.Size = New System.Drawing.Size(56, 16)
        Me.lblPhone2.TabIndex = 228
        Me.lblPhone2.Text = "Phone 2:"
        Me.lblPhone2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'mskTxtCell
        '
        Me.mskTxtCell.ContainingControl = Me
        Me.mskTxtCell.Location = New System.Drawing.Point(72, 78)
        Me.mskTxtCell.Name = "mskTxtCell"
        Me.mskTxtCell.OcxState = CType(resources.GetObject("mskTxtCell.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtCell.Size = New System.Drawing.Size(142, 20)
        Me.mskTxtCell.TabIndex = 17
        '
        'lblCell
        '
        Me.lblCell.Location = New System.Drawing.Point(10, 78)
        Me.lblCell.Name = "lblCell"
        Me.lblCell.Size = New System.Drawing.Size(56, 16)
        Me.lblCell.TabIndex = 230
        Me.lblCell.Text = "Cell:"
        Me.lblCell.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'mskTxtPhone2
        '
        Me.mskTxtPhone2.ContainingControl = Me
        Me.mskTxtPhone2.Location = New System.Drawing.Point(72, 30)
        Me.mskTxtPhone2.Name = "mskTxtPhone2"
        Me.mskTxtPhone2.OcxState = CType(resources.GetObject("mskTxtPhone2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtPhone2.Size = New System.Drawing.Size(142, 20)
        Me.mskTxtPhone2.TabIndex = 14
        '
        'lblPhone1
        '
        Me.lblPhone1.Location = New System.Drawing.Point(10, 6)
        Me.lblPhone1.Name = "lblPhone1"
        Me.lblPhone1.Size = New System.Drawing.Size(56, 16)
        Me.lblPhone1.TabIndex = 227
        Me.lblPhone1.Text = "Phone 1:"
        Me.lblPhone1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'mskTxtFax
        '
        Me.mskTxtFax.ContainingControl = Me
        Me.mskTxtFax.Location = New System.Drawing.Point(72, 54)
        Me.mskTxtFax.Name = "mskTxtFax"
        Me.mskTxtFax.OcxState = CType(resources.GetObject("mskTxtFax.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtFax.Size = New System.Drawing.Size(142, 20)
        Me.mskTxtFax.TabIndex = 16
        '
        'lblFax
        '
        Me.lblFax.Location = New System.Drawing.Point(10, 54)
        Me.lblFax.Name = "lblFax"
        Me.lblFax.Size = New System.Drawing.Size(56, 16)
        Me.lblFax.TabIndex = 229
        Me.lblFax.Text = "Fax:"
        Me.lblFax.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlPhoneComment
        '
        Me.pnlPhoneComment.Controls.Add(Me.lblPhone2Comment)
        Me.pnlPhoneComment.Controls.Add(Me.lblPhone1Comment)
        Me.pnlPhoneComment.Controls.Add(Me.txtPhone2Comment)
        Me.pnlPhoneComment.Controls.Add(Me.txtPhone1Comment)
        Me.pnlPhoneComment.Location = New System.Drawing.Point(344, 216)
        Me.pnlPhoneComment.Name = "pnlPhoneComment"
        Me.pnlPhoneComment.Size = New System.Drawing.Size(272, 56)
        Me.pnlPhoneComment.TabIndex = 1004
        '
        'lblPhone2Comment
        '
        Me.lblPhone2Comment.Location = New System.Drawing.Point(8, 30)
        Me.lblPhone2Comment.Name = "lblPhone2Comment"
        Me.lblPhone2Comment.Size = New System.Drawing.Size(104, 16)
        Me.lblPhone2Comment.TabIndex = 236
        Me.lblPhone2Comment.Text = "Phone 2 Comment:"
        Me.lblPhone2Comment.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPhone1Comment
        '
        Me.lblPhone1Comment.Location = New System.Drawing.Point(8, 6)
        Me.lblPhone1Comment.Name = "lblPhone1Comment"
        Me.lblPhone1Comment.Size = New System.Drawing.Size(104, 16)
        Me.lblPhone1Comment.TabIndex = 235
        Me.lblPhone1Comment.Text = "Phone 1 Comment:"
        Me.lblPhone1Comment.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtPhone2Comment
        '
        Me.txtPhone2Comment.Location = New System.Drawing.Point(120, 30)
        Me.txtPhone2Comment.Name = "txtPhone2Comment"
        Me.txtPhone2Comment.Size = New System.Drawing.Size(144, 20)
        Me.txtPhone2Comment.TabIndex = 1
        Me.txtPhone2Comment.Text = ""
        '
        'txtPhone1Comment
        '
        Me.txtPhone1Comment.Location = New System.Drawing.Point(120, 6)
        Me.txtPhone1Comment.Name = "txtPhone1Comment"
        Me.txtPhone1Comment.Size = New System.Drawing.Size(144, 20)
        Me.txtPhone1Comment.TabIndex = 0
        Me.txtPhone1Comment.Text = ""
        '
        'Providers
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(656, 438)
        Me.Controls.Add(Me.pnlPhoneComment)
        Me.Controls.Add(Me.pnlExtension)
        Me.Controls.Add(Me.pnlPhone)
        Me.Controls.Add(Me.cmbProviderName)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.txtProviderAddress)
        Me.Controls.Add(Me.txtWebsite)
        Me.Controls.Add(Me.txtProviderID)
        Me.Controls.Add(Me.txtProvider)
        Me.Controls.Add(Me.txtAbbreviation)
        Me.Controls.Add(Me.txtDepartment)
        Me.Controls.Add(Me.chkActive)
        Me.Controls.Add(Me.lblAddress)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.lblWebsite)
        Me.Controls.Add(Me.lblProviderID)
        Me.Controls.Add(Me.lblProvider)
        Me.Controls.Add(Me.lblAbbreviation)
        Me.Controls.Add(Me.lblDepartment)
        Me.Name = "Providers"
        Me.Text = "Providers"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.pnlExtension.ResumeLayout(False)
        Me.pnlPhone.ResumeLayout(False)
        CType(Me.mskTxtPhone1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtCell, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtPhone2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtFax, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlPhoneComment.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Providers_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            bolLoading = True

            cmbProviderName.DisplayMember = "Provider_Name"
            cmbProviderName.ValueMember = "Provider_ID"
            cmbProviderName.DataSource = oProvider.ListProviderNames(False)
            'cmbProviderName.SelectedIndex = IIf(cmbProviderName.Items.Count > 0, 0, -1)
            cmbProviderName.Text = String.Empty
            ClearProviderData()
            bolLoading = False
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            bolLoading = True
            nProviderID = 0
            nAddressID = 0
            ClearProviderData()

            bolLoading = False
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    ' Validates Data according to DDD Specifications
    Public Function ValidateData() As Boolean
        Try
            Dim errStr As String = ""
            Dim validateSuccess As Boolean = True


            If txtProvider.Text <> String.Empty Then
                If txtAbbreviation.Text <> String.Empty Then
                    'If txtDepartment.Text <> String.Empty Then
                    '    validateSuccess = True
                    'Else
                    '    errStr += "Department cannot be empty" + vbCrLf
                    '    validateSuccess = False
                    'End If
                    validateSuccess = True
                Else
                    errStr += "Abbreviation cannot be empty" + vbCrLf
                    validateSuccess = False
                End If
            Else
                errStr += "Provider cannot be empty" + vbCrLf
                validateSuccess = False
            End If

            If errStr.Length > 0 Or Not validateSuccess Then
                MsgBox(errStr)
            End If
            Return validateSuccess
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            If ValidateData() Then
                oProvInfo = New MUSTER.Info.ProviderInfo(nProviderID, _
                                            chkActive.Checked, _
                                            txtProvider.Text, _
                                            txtAbbreviation.Text, _
                                            txtDepartment.Text, _
                                            txtWebsite.Text, _
                                            False, _
                                            IIf(nProviderID <= 0, MusterContainer.AppUser.ID, ""), _
                                            Now, _
                                            IIf(nProviderID > 0, MusterContainer.AppUser.ID, ""), _
                                            CDate("01/01/0001"))

                oProvider.Add(oProvInfo)
                If oProvider.ID <= 0 Then
                    oProvider.CreatedBy = MusterContainer.AppUser.ID
                Else
                    oProvider.ModifiedBy = MusterContainer.AppUser.ID
                End If
                oProvider.Save(CType(UIUtilsGen.ModuleID.CompanyAdmin, Integer), MusterContainer.AppUser.UserKey, returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                oComAdd.ProviderID = oProvider.ProviderInfo.ID
                If oComAdd.AddressId <= 0 Then
                    oComAdd.CreatedBy = MusterContainer.AppUser.ID
                Else
                    oComAdd.ModifiedBy = MusterContainer.AppUser.ID
                End If
                oComAdd.Save(CType(UIUtilsGen.ModuleID.CompanyAdmin, Integer), MusterContainer.AppUser.UserKey, returnVal, False)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                MsgBox("Provider is saved successfully!")
                bolLoading = True

                cmbProviderName.DisplayMember = "Provider_Name"
                cmbProviderName.ValueMember = "Provider_ID"
                cmbProviderName.DataSource = oProvider.ListProviderNames(False)
                'cmbProviderName.SelectedValue = oProvider.ProviderInfo.ID
                UIUtilsGen.SetComboboxItemByValue(cmbProviderName, oProvider.ProviderInfo.ID)
                bolLoading = False
                bolValidationFlg = False

            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.Close()
            '' temporarily added - to be removed
            'Dim obj As New ManageContactTypes
            'obj.ShowDialog()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try
            Dim msgResult As MsgBoxResult = MsgBox("Are you sure you wish to DELETE the Provider : " & cmbProviderName.Text & " ?", MsgBoxStyle.Question & MsgBoxStyle.YesNo, "DELETE PROVIDER")
            If msgResult = MsgBoxResult.Yes Then
                'Delete Associated Address
                If cmbProviderName.SelectedValue Is Nothing Then
                    'oComAdd.GetAddressByType(0, 0, 0, oProvider.ID)
                    oComAdd.Retrieve(oProvider.ID)
                Else
                    oComAdd.Retrieve(cmbProviderName.SelectedValue)
                    'oComAdd.GetAddressByType(0, 0, 0, cmbProviderName.SelectedValue)
                End If

                oComAdd.Deleted = True
                oComAdd.ModifiedBy = MusterContainer.AppUser.ID
                oComAdd.Save(CType(UIUtilsGen.ModuleID.CompanyAdmin, Integer), MusterContainer.AppUser.UserKey, returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                oComAdd.Remove(oComAdd.AddressId)
                'Delete Provider 
                oProvider.Deleted = True
                If oProvider.ID <= 0 Then
                    oProvider.CreatedBy = MusterContainer.AppUser.ID
                Else
                    oProvider.ModifiedBy = MusterContainer.AppUser.ID
                End If

                oProvider.Save(CType(UIUtilsGen.ModuleID.CompanyAdmin, Integer), MusterContainer.AppUser.UserKey, returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                If cmbProviderName.SelectedValue Is Nothing Then
                    oProvider.Remove(oProvider.ID)
                Else
                    oProvider.Remove(cmbProviderName.SelectedValue)
                End If

                MsgBox("Provider Deleted Successfully.")
                bolLoading = True
                cmbProviderName.DisplayMember = "Provider_Name"
                cmbProviderName.ValueMember = "Provider_ID"
                cmbProviderName.DataSource = oProvider.ListProviderNames(False)
                cmbProviderName.SelectedIndex = -1
                cmbProviderName.Text = String.Empty
                ClearProviderData()

                Me.cmbProviderName.Focus()
                bolLoading = False

            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnSearchProvider_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtProviderAddress_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtProviderAddress.DoubleClick
        Try

            'objAddMaster = New AddressMaster(oComAdd, nAddressID, nProviderID, "Provider", IIf(txtProviderAddress.Text = String.Empty, "ADD", "MODIFY"))
            If txtProviderAddress.Text = "" Then
                objAddMaster = New AddressMaster(oComAdd, 0, nProviderID, "Provider", "ADD")
            Else
                objAddMaster = New AddressMaster(oComAdd, nAddressID, nProviderID, "Provider", "MODIFY")
            End If

            Me.Update()
            'LockWindowUpdate(Me.Handle.ToInt64)
            objAddMaster.ShowDialog()
            'LockWindowUpdate(0)
            nAddressID = oComAdd.AddressId
            'strAddress = oComAdd.AddressLine1 + IIf(oComAdd.AddressLine1.Trim.Length = 0, "", vbCrLf + oComAdd.AddressLine2) + vbCrLf + oComAdd.City + ", " + oComAdd.State + " " + oComAdd.Zip
            strAddress = oComAdd.AddressLine1 + IIf(oComAdd.AddressLine1.Trim.Length = 0, "", vbCrLf + oComAdd.AddressLine2) + IIf(oComAdd.City.Length = 0, "", vbCrLf + oComAdd.City) + IIf(oComAdd.State.Length = 0, "", ", " + oComAdd.State) + IIf(oComAdd.Zip.Length = 0, "", " " + oComAdd.Zip)
            If strAddress <> String.Empty Then
                Me.txtProviderAddress.Text = strAddress
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub cmbProviderName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbProviderName.SelectedIndexChanged
        Dim strOldProviderName As String = oProvider.ID
        If bolLoading Then Exit Sub



        Try
            ' If Not cmbProviderName.SelectedValue <> String.Empty Then Exit Sub
            If cmbProviderName.SelectedIndex = -1 Then Exit Sub

            oProvider.Retrieve(cmbProviderName.SelectedValue)

            Me.txtProvider.Text = cmbProviderName.Text
            Me.txtProviderID.Text = cmbProviderName.SelectedValue
            nProviderID = cmbProviderName.SelectedValue
            Me.txtAbbreviation.Text = oProvider.Abbrev
            Me.txtDepartment.Text = oProvider.Department
            Me.txtWebsite.Text = oProvider.Website
            Me.chkActive.Checked = oProvider.Active

            'oComAddInfo = oComAdd.GetAddressByType(0, 0, 0, cmbProviderName.SelectedValue)
            oComAddInfo = oComAdd.Retrieve(cmbProviderName.SelectedValue)
            strAddress = oComAddInfo.AddressLine1 + IIf(oComAddInfo.AddressLine1.Trim.Length = 0, "", vbCrLf + oComAddInfo.AddressLine2) + vbCrLf + oComAddInfo.City + ", " + oComAddInfo.State + " " + oComAddInfo.Zip
            nAddressID = oComAddInfo.AddressId
            Me.txtProviderAddress.Text = strAddress
            mskTxtPhone1.SelText = oComAddInfo.Phone1
            mskTxtPhone2.SelText = oComAddInfo.Phone2
            mskTxtCell.SelText = oComAddInfo.Cell
            mskTxtFax.SelText = oComAddInfo.Fax
            txtExt1.Text = oComAddInfo.Ext1
            txtExt2.Text = oComAddInfo.Ext2
            Me.txtPhone1Comment.Text = oComAddInfo.Phone1Comment
            Me.txtPhone2Comment.Text = oComAddInfo.Phone2Comment
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ClearProviderData()

        bolLoading = True

        Me.txtProvider.Text = String.Empty
        Me.txtProviderID.Text = String.Empty
        Me.txtAbbreviation.Text = String.Empty
        Me.txtDepartment.Text = String.Empty
        Me.txtProviderAddress.Text = String.Empty
        Me.txtWebsite.Text = String.Empty
        Me.cmbProviderName.SelectedIndex = -1
        Me.cmbProviderName.SelectedIndex = -1
        Me.chkActive.Checked = True
        mskTxtPhone1.SelText = String.Empty
        mskTxtPhone2.SelText = String.Empty
        mskTxtCell.SelText = String.Empty
        mskTxtFax.SelText = String.Empty
        txtExt1.Text = String.Empty
        txtExt2.Text = String.Empty
        Me.txtPhone1Comment.Text = String.Empty
        Me.txtPhone2Comment.Text = String.Empty
        bolLoading = False

    End Sub

    Private Sub oProvider_evtProviderErr(ByVal MsgStr As String) Handles oProvider.evtProviderErr
        If MsgStr <> String.Empty And MsgBox(MsgStr) = MsgBoxResult.OK Then
            MsgStr = String.Empty
            bolValidationFlg = True
            Exit Sub
        End If
    End Sub

    Private Sub txtProvider_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProvider.Leave
        Dim item As Object
        Try
            If cmbProviderName.SelectedIndex = -1 And nProviderID = 0 Then
                For Each item In cmbProviderName.Items
                    If item("PROVIDER_NAME") = txtProvider.Text Then
                        'cmbProviderName.SelectedValue = Integer.Parse(item("PROVIDER_ID"))
                        UIUtilsGen.SetComboboxItemByValue(cmbProviderName, Integer.Parse(item("PROVIDER_ID")))
                        Exit Try
                    End If
                Next
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub mskTxtPhone1_KeyUpEvent(ByVal sender As Object, ByVal e As AxMSMask.MaskEdBoxEvents_KeyUpEvent) Handles mskTxtPhone1.KeyUpEvent
        If bolLoading Then Exit Sub
        UIUtilsGen.FillStringObjectValues(oComAdd.Phone1, mskTxtPhone1.FormattedText.ToString)
    End Sub

    Private Sub mskTxtPhone2_KeyUpEvent(ByVal sender As Object, ByVal e As AxMSMask.MaskEdBoxEvents_KeyUpEvent) Handles mskTxtPhone2.KeyUpEvent
        If bolLoading Then Exit Sub
        UIUtilsGen.FillStringObjectValues(oComAdd.Phone2, mskTxtPhone2.FormattedText.ToString)
    End Sub

    Private Sub mskTxtCell_KeyUpEvent(ByVal sender As Object, ByVal e As AxMSMask.MaskEdBoxEvents_KeyUpEvent) Handles mskTxtCell.KeyUpEvent
        If bolLoading Then Exit Sub
        UIUtilsGen.FillStringObjectValues(oComAdd.Cell, mskTxtCell.FormattedText.ToString)
    End Sub

    Private Sub mskTxtFax_KeyUpEvent(ByVal sender As Object, ByVal e As AxMSMask.MaskEdBoxEvents_KeyUpEvent) Handles mskTxtFax.KeyUpEvent
        If bolLoading Then Exit Sub
        UIUtilsGen.FillStringObjectValues(oComAdd.Fax, mskTxtFax.FormattedText.ToString)
    End Sub

    Private Sub txtExt1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExt1.TextChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.FillStringObjectValues(oComAdd.Ext1, txtExt1.Text.Trim)
    End Sub

    Private Sub txtExt2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExt2.TextChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.FillStringObjectValues(oComAdd.Ext2, txtExt2.Text.Trim)
    End Sub

    Private Sub txtPhone1Comment_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPhone1Comment.TextChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.FillStringObjectValues(oComAdd.Phone1Comment, txtPhone1Comment.Text.Trim)
    End Sub

    Private Sub txtPhone2Comment_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPhone2Comment.TextChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.FillStringObjectValues(oComAdd.Phone2Comment, txtPhone2Comment.Text.Trim)
    End Sub
End Class
