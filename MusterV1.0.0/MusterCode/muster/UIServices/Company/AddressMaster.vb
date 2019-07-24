Public Class AddressMaster
    Inherits System.Windows.Forms.Form
#Region "User Defined Variables"
    'Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long
    Private strEntityType As String
    Private nCompanyID As Integer = 0
    Private nAddressID As Integer = 0
    Private WithEvents pComAdd As MUSTER.BusinessLogic.pComAddress
    Private oAddInfo As MUSTER.Info.ComAddressInfo
    Private bolLoading As Boolean = False
    Dim strMode As String = "ADD"
    Private bolValidationFlg As Boolean = False
    Dim strPrevCity As String = String.Empty
    Dim strPrevState As String = String.Empty
    Private bolShowFIPS As Boolean = True
    Private bolSaved As Boolean = False
    Private returnVal As String = String.Empty
#End Region
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub
    Public Sub New(ByRef pAddress As MUSTER.BusinessLogic.pComAddress, Optional ByVal AddressID As Integer = 0, Optional ByVal EntityID As Integer = 0, Optional ByVal EntityType As String = "Company", Optional ByVal Mode As String = "ADD", Optional ByVal RetValue As String = "")
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        pComAdd = pAddress
        nCompanyID = EntityID
        strEntityType = EntityType
        nAddressID = AddressID
        strMode = Mode
        returnVal = RetValue

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
    Friend WithEvents btnClearZip As System.Windows.Forms.Button
    Friend WithEvents btnClearState As System.Windows.Forms.Button
    Friend WithEvents btnClearCity As System.Windows.Forms.Button
    Friend WithEvents btnClearData As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents lblZip As System.Windows.Forms.Label
    Friend WithEvents lblState As System.Windows.Forms.Label
    Friend WithEvents lblCity As System.Windows.Forms.Label
    Friend WithEvents lblCounty As System.Windows.Forms.Label
    Friend WithEvents lblFIPS As System.Windows.Forms.Label
    Friend WithEvents lblAddress2 As System.Windows.Forms.Label
    Friend WithEvents lblAddress1 As System.Windows.Forms.Label
    Friend WithEvents cboCounty As Infragistics.Win.UltraWinGrid.UltraCombo
    Friend WithEvents cboZipCode As Infragistics.Win.UltraWinGrid.UltraCombo
    Friend WithEvents cboState As Infragistics.Win.UltraWinGrid.UltraCombo
    Friend WithEvents cboCity As Infragistics.Win.UltraWinGrid.UltraCombo
    Friend WithEvents cboFIPS As Infragistics.Win.UltraWinGrid.UltraCombo
    Friend WithEvents txtAddress2 As System.Windows.Forms.TextBox
    Friend WithEvents txtAddress1 As System.Windows.Forms.TextBox
    Public WithEvents mskTxtCell As AxMSMask.AxMaskEdBox
    Friend WithEvents lblCell As System.Windows.Forms.Label
    Public WithEvents mskTxtFax As AxMSMask.AxMaskEdBox
    Friend WithEvents lblFax As System.Windows.Forms.Label
    Public WithEvents mskTxtPhone2 As AxMSMask.AxMaskEdBox
    Public WithEvents mskTxtPhone1 As AxMSMask.AxMaskEdBox
    Friend WithEvents lblPhone2 As System.Windows.Forms.Label
    Friend WithEvents lblPhone1 As System.Windows.Forms.Label
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents lblExt2 As System.Windows.Forms.Label
    Friend WithEvents lblExt1 As System.Windows.Forms.Label
    Friend WithEvents txtExt2 As System.Windows.Forms.TextBox
    Friend WithEvents txtExt1 As System.Windows.Forms.TextBox
    Friend WithEvents pnlExtension As System.Windows.Forms.Panel
    Friend WithEvents lblPhone1Comment As System.Windows.Forms.Label
    Friend WithEvents lblPhone2Comment As System.Windows.Forms.Label
    Friend WithEvents txtPhone2Comment As System.Windows.Forms.TextBox
    Friend WithEvents txtPhone1Comment As System.Windows.Forms.TextBox
    Friend WithEvents pnlPhoneComment As System.Windows.Forms.Panel
    Friend WithEvents pnlPhone As System.Windows.Forms.Panel
    Friend WithEvents btnClearCounty As System.Windows.Forms.Button
    Friend WithEvents btnClearFIPS As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(AddressMaster))
        Me.btnClearZip = New System.Windows.Forms.Button
        Me.btnClearState = New System.Windows.Forms.Button
        Me.btnClearCity = New System.Windows.Forms.Button
        Me.btnClearData = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.lblZip = New System.Windows.Forms.Label
        Me.lblState = New System.Windows.Forms.Label
        Me.lblCity = New System.Windows.Forms.Label
        Me.lblCounty = New System.Windows.Forms.Label
        Me.lblFIPS = New System.Windows.Forms.Label
        Me.lblAddress2 = New System.Windows.Forms.Label
        Me.lblAddress1 = New System.Windows.Forms.Label
        Me.cboCounty = New Infragistics.Win.UltraWinGrid.UltraCombo
        Me.cboZipCode = New Infragistics.Win.UltraWinGrid.UltraCombo
        Me.cboState = New Infragistics.Win.UltraWinGrid.UltraCombo
        Me.cboCity = New Infragistics.Win.UltraWinGrid.UltraCombo
        Me.cboFIPS = New Infragistics.Win.UltraWinGrid.UltraCombo
        Me.txtAddress2 = New System.Windows.Forms.TextBox
        Me.txtAddress1 = New System.Windows.Forms.TextBox
        Me.mskTxtCell = New AxMSMask.AxMaskEdBox
        Me.lblCell = New System.Windows.Forms.Label
        Me.mskTxtFax = New AxMSMask.AxMaskEdBox
        Me.lblFax = New System.Windows.Forms.Label
        Me.mskTxtPhone2 = New AxMSMask.AxMaskEdBox
        Me.mskTxtPhone1 = New AxMSMask.AxMaskEdBox
        Me.lblPhone2 = New System.Windows.Forms.Label
        Me.lblPhone1 = New System.Windows.Forms.Label
        Me.pnlExtension = New System.Windows.Forms.Panel
        Me.lblExt2 = New System.Windows.Forms.Label
        Me.lblExt1 = New System.Windows.Forms.Label
        Me.txtExt2 = New System.Windows.Forms.TextBox
        Me.txtExt1 = New System.Windows.Forms.TextBox
        Me.pnlPhoneComment = New System.Windows.Forms.Panel
        Me.lblPhone2Comment = New System.Windows.Forms.Label
        Me.lblPhone1Comment = New System.Windows.Forms.Label
        Me.txtPhone2Comment = New System.Windows.Forms.TextBox
        Me.txtPhone1Comment = New System.Windows.Forms.TextBox
        Me.pnlPhone = New System.Windows.Forms.Panel
        Me.btnClearCounty = New System.Windows.Forms.Button
        Me.btnClearFIPS = New System.Windows.Forms.Button
        CType(Me.cboCounty, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboZipCode, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboState, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboCity, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cboFIPS, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtCell, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtFax, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtPhone2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtPhone1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlExtension.SuspendLayout()
        Me.pnlPhoneComment.SuspendLayout()
        Me.pnlPhone.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnClearZip
        '
        Me.btnClearZip.Image = CType(resources.GetObject("btnClearZip.Image"), System.Drawing.Image)
        Me.btnClearZip.Location = New System.Drawing.Point(256, 176)
        Me.btnClearZip.Name = "btnClearZip"
        Me.btnClearZip.Size = New System.Drawing.Size(17, 17)
        Me.btnClearZip.TabIndex = 12
        Me.btnClearZip.Visible = False
        '
        'btnClearState
        '
        Me.btnClearState.Image = CType(resources.GetObject("btnClearState.Image"), System.Drawing.Image)
        Me.btnClearState.Location = New System.Drawing.Point(192, 160)
        Me.btnClearState.Name = "btnClearState"
        Me.btnClearState.Size = New System.Drawing.Size(17, 17)
        Me.btnClearState.TabIndex = 9
        Me.btnClearState.Visible = False
        '
        'btnClearCity
        '
        Me.btnClearCity.Image = CType(resources.GetObject("btnClearCity.Image"), System.Drawing.Image)
        Me.btnClearCity.Location = New System.Drawing.Point(320, 128)
        Me.btnClearCity.Name = "btnClearCity"
        Me.btnClearCity.Size = New System.Drawing.Size(17, 17)
        Me.btnClearCity.TabIndex = 7
        Me.btnClearCity.Visible = False
        '
        'btnClearData
        '
        Me.btnClearData.Location = New System.Drawing.Point(224, 312)
        Me.btnClearData.Name = "btnClearData"
        Me.btnClearData.Size = New System.Drawing.Size(80, 26)
        Me.btnClearData.TabIndex = 20
        Me.btnClearData.Text = "C&lear"
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(136, 312)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(80, 26)
        Me.btnCancel.TabIndex = 19
        Me.btnCancel.Text = "&Cancel"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(48, 312)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(80, 26)
        Me.btnSave.TabIndex = 18
        Me.btnSave.Text = "&Save"
        '
        'lblZip
        '
        Me.lblZip.Location = New System.Drawing.Point(32, 176)
        Me.lblZip.Name = "lblZip"
        Me.lblZip.Size = New System.Drawing.Size(72, 16)
        Me.lblZip.TabIndex = 38
        Me.lblZip.Text = "Zip Code:"
        Me.lblZip.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblState
        '
        Me.lblState.Location = New System.Drawing.Point(48, 152)
        Me.lblState.Name = "lblState"
        Me.lblState.Size = New System.Drawing.Size(56, 16)
        Me.lblState.TabIndex = 36
        Me.lblState.Text = "State:"
        Me.lblState.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCity
        '
        Me.lblCity.Location = New System.Drawing.Point(48, 128)
        Me.lblCity.Name = "lblCity"
        Me.lblCity.Size = New System.Drawing.Size(56, 16)
        Me.lblCity.TabIndex = 34
        Me.lblCity.Text = "City:"
        Me.lblCity.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCounty
        '
        Me.lblCounty.Location = New System.Drawing.Point(48, 112)
        Me.lblCounty.Name = "lblCounty"
        Me.lblCounty.Size = New System.Drawing.Size(56, 16)
        Me.lblCounty.TabIndex = 32
        Me.lblCounty.Text = "County:"
        Me.lblCounty.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblFIPS
        '
        Me.lblFIPS.Location = New System.Drawing.Point(32, 80)
        Me.lblFIPS.Name = "lblFIPS"
        Me.lblFIPS.Size = New System.Drawing.Size(72, 16)
        Me.lblFIPS.TabIndex = 30
        Me.lblFIPS.Text = "FIPS Code:"
        Me.lblFIPS.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblAddress2
        '
        Me.lblAddress2.Location = New System.Drawing.Point(16, 56)
        Me.lblAddress2.Name = "lblAddress2"
        Me.lblAddress2.Size = New System.Drawing.Size(88, 16)
        Me.lblAddress2.TabIndex = 27
        Me.lblAddress2.Text = "Address Line 2:"
        Me.lblAddress2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblAddress1
        '
        Me.lblAddress1.Location = New System.Drawing.Point(16, 32)
        Me.lblAddress1.Name = "lblAddress1"
        Me.lblAddress1.Size = New System.Drawing.Size(88, 16)
        Me.lblAddress1.TabIndex = 25
        Me.lblAddress1.Text = "Address Line 1:"
        Me.lblAddress1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cboCounty
        '
        Me.cboCounty.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.cboCounty.DisplayMember = ""
        Me.cboCounty.Location = New System.Drawing.Point(112, 104)
        Me.cboCounty.Name = "cboCounty"
        Me.cboCounty.Size = New System.Drawing.Size(200, 21)
        Me.cboCounty.TabIndex = 4
        Me.cboCounty.ValueMember = ""
        '
        'cboZipCode
        '
        Me.cboZipCode.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.cboZipCode.DisplayMember = ""
        Me.cboZipCode.Location = New System.Drawing.Point(112, 176)
        Me.cboZipCode.Name = "cboZipCode"
        Me.cboZipCode.Size = New System.Drawing.Size(136, 21)
        Me.cboZipCode.TabIndex = 11
        Me.cboZipCode.ValueMember = ""
        '
        'cboState
        '
        Me.cboState.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.cboState.DisplayLayout.BorderStyle = Infragistics.Win.UIElementBorderStyle.Inset
        Me.cboState.DisplayMember = ""
        Me.cboState.Location = New System.Drawing.Point(112, 152)
        Me.cboState.Name = "cboState"
        Me.cboState.Size = New System.Drawing.Size(72, 21)
        Me.cboState.TabIndex = 8
        Me.cboState.ValueMember = ""
        '
        'cboCity
        '
        Me.cboCity.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.cboCity.DisplayMember = ""
        Me.cboCity.Location = New System.Drawing.Point(112, 128)
        Me.cboCity.Name = "cboCity"
        Me.cboCity.Size = New System.Drawing.Size(200, 21)
        Me.cboCity.TabIndex = 6
        Me.cboCity.ValueMember = ""
        '
        'cboFIPS
        '
        Me.cboFIPS.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.cboFIPS.DisplayMember = ""
        Me.cboFIPS.Location = New System.Drawing.Point(112, 80)
        Me.cboFIPS.Name = "cboFIPS"
        Me.cboFIPS.Size = New System.Drawing.Size(96, 21)
        Me.cboFIPS.TabIndex = 2
        Me.cboFIPS.ValueMember = ""
        '
        'txtAddress2
        '
        Me.txtAddress2.Location = New System.Drawing.Point(112, 56)
        Me.txtAddress2.Name = "txtAddress2"
        Me.txtAddress2.Size = New System.Drawing.Size(200, 20)
        Me.txtAddress2.TabIndex = 1
        Me.txtAddress2.Text = ""
        '
        'txtAddress1
        '
        Me.txtAddress1.Location = New System.Drawing.Point(112, 32)
        Me.txtAddress1.Name = "txtAddress1"
        Me.txtAddress1.Size = New System.Drawing.Size(200, 20)
        Me.txtAddress1.TabIndex = 0
        Me.txtAddress1.Text = ""
        '
        'mskTxtCell
        '
        Me.mskTxtCell.ContainingControl = Me
        Me.mskTxtCell.Location = New System.Drawing.Point(74, 78)
        Me.mskTxtCell.Name = "mskTxtCell"
        Me.mskTxtCell.OcxState = CType(resources.GetObject("mskTxtCell.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtCell.Size = New System.Drawing.Size(96, 20)
        Me.mskTxtCell.TabIndex = 17
        '
        'lblCell
        '
        Me.lblCell.Location = New System.Drawing.Point(2, 78)
        Me.lblCell.Name = "lblCell"
        Me.lblCell.Size = New System.Drawing.Size(56, 16)
        Me.lblCell.TabIndex = 230
        Me.lblCell.Text = "Cell:"
        Me.lblCell.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'mskTxtFax
        '
        Me.mskTxtFax.ContainingControl = Me
        Me.mskTxtFax.Location = New System.Drawing.Point(74, 54)
        Me.mskTxtFax.Name = "mskTxtFax"
        Me.mskTxtFax.OcxState = CType(resources.GetObject("mskTxtFax.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtFax.Size = New System.Drawing.Size(96, 20)
        Me.mskTxtFax.TabIndex = 16
        '
        'lblFax
        '
        Me.lblFax.Location = New System.Drawing.Point(2, 54)
        Me.lblFax.Name = "lblFax"
        Me.lblFax.Size = New System.Drawing.Size(56, 16)
        Me.lblFax.TabIndex = 229
        Me.lblFax.Text = "Fax:"
        Me.lblFax.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'mskTxtPhone2
        '
        Me.mskTxtPhone2.ContainingControl = Me
        Me.mskTxtPhone2.Location = New System.Drawing.Point(74, 30)
        Me.mskTxtPhone2.Name = "mskTxtPhone2"
        Me.mskTxtPhone2.OcxState = CType(resources.GetObject("mskTxtPhone2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtPhone2.Size = New System.Drawing.Size(96, 20)
        Me.mskTxtPhone2.TabIndex = 14
        '
        'mskTxtPhone1
        '
        Me.mskTxtPhone1.ContainingControl = Me
        Me.mskTxtPhone1.Location = New System.Drawing.Point(74, 6)
        Me.mskTxtPhone1.Name = "mskTxtPhone1"
        Me.mskTxtPhone1.OcxState = CType(resources.GetObject("mskTxtPhone1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtPhone1.Size = New System.Drawing.Size(96, 20)
        Me.mskTxtPhone1.TabIndex = 13
        '
        'lblPhone2
        '
        Me.lblPhone2.Location = New System.Drawing.Point(2, 30)
        Me.lblPhone2.Name = "lblPhone2"
        Me.lblPhone2.Size = New System.Drawing.Size(56, 16)
        Me.lblPhone2.TabIndex = 228
        Me.lblPhone2.Text = "Phone 2:"
        Me.lblPhone2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPhone1
        '
        Me.lblPhone1.Location = New System.Drawing.Point(2, 6)
        Me.lblPhone1.Name = "lblPhone1"
        Me.lblPhone1.Size = New System.Drawing.Size(56, 16)
        Me.lblPhone1.TabIndex = 227
        Me.lblPhone1.Text = "Phone 1:"
        Me.lblPhone1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlExtension
        '
        Me.pnlExtension.Controls.Add(Me.lblExt2)
        Me.pnlExtension.Controls.Add(Me.lblExt1)
        Me.pnlExtension.Controls.Add(Me.txtExt2)
        Me.pnlExtension.Controls.Add(Me.txtExt1)
        Me.pnlExtension.Location = New System.Drawing.Point(216, 200)
        Me.pnlExtension.Name = "pnlExtension"
        Me.pnlExtension.Size = New System.Drawing.Size(112, 56)
        Me.pnlExtension.TabIndex = 15
        '
        'lblExt2
        '
        Me.lblExt2.Location = New System.Drawing.Point(8, 30)
        Me.lblExt2.Name = "lblExt2"
        Me.lblExt2.Size = New System.Drawing.Size(40, 16)
        Me.lblExt2.TabIndex = 236
        Me.lblExt2.Text = "Ext 2:"
        Me.lblExt2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblExt1
        '
        Me.lblExt1.Location = New System.Drawing.Point(8, 6)
        Me.lblExt1.Name = "lblExt1"
        Me.lblExt1.Size = New System.Drawing.Size(40, 16)
        Me.lblExt1.TabIndex = 235
        Me.lblExt1.Text = "Ext 1:"
        Me.lblExt1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtExt2
        '
        Me.txtExt2.Location = New System.Drawing.Point(48, 29)
        Me.txtExt2.Name = "txtExt2"
        Me.txtExt2.Size = New System.Drawing.Size(40, 20)
        Me.txtExt2.TabIndex = 234
        Me.txtExt2.Text = ""
        '
        'txtExt1
        '
        Me.txtExt1.Location = New System.Drawing.Point(48, 6)
        Me.txtExt1.Name = "txtExt1"
        Me.txtExt1.Size = New System.Drawing.Size(40, 20)
        Me.txtExt1.TabIndex = 233
        Me.txtExt1.Text = ""
        '
        'pnlPhoneComment
        '
        Me.pnlPhoneComment.Controls.Add(Me.lblPhone2Comment)
        Me.pnlPhoneComment.Controls.Add(Me.lblPhone1Comment)
        Me.pnlPhoneComment.Controls.Add(Me.txtPhone2Comment)
        Me.pnlPhoneComment.Controls.Add(Me.txtPhone1Comment)
        Me.pnlPhoneComment.Location = New System.Drawing.Point(336, 248)
        Me.pnlPhoneComment.Name = "pnlPhoneComment"
        Me.pnlPhoneComment.Size = New System.Drawing.Size(272, 56)
        Me.pnlPhoneComment.TabIndex = 15
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
        Me.pnlPhone.Location = New System.Drawing.Point(37, 200)
        Me.pnlPhone.Name = "pnlPhone"
        Me.pnlPhone.Size = New System.Drawing.Size(176, 104)
        Me.pnlPhone.TabIndex = 231
        '
        'btnClearCounty
        '
        Me.btnClearCounty.Image = CType(resources.GetObject("btnClearCounty.Image"), System.Drawing.Image)
        Me.btnClearCounty.Location = New System.Drawing.Point(320, 104)
        Me.btnClearCounty.Name = "btnClearCounty"
        Me.btnClearCounty.Size = New System.Drawing.Size(17, 17)
        Me.btnClearCounty.TabIndex = 5
        Me.btnClearCounty.Visible = False
        '
        'btnClearFIPS
        '
        Me.btnClearFIPS.Image = CType(resources.GetObject("btnClearFIPS.Image"), System.Drawing.Image)
        Me.btnClearFIPS.Location = New System.Drawing.Point(216, 80)
        Me.btnClearFIPS.Name = "btnClearFIPS"
        Me.btnClearFIPS.Size = New System.Drawing.Size(17, 17)
        Me.btnClearFIPS.TabIndex = 3
        Me.btnClearFIPS.Visible = False
        '
        'AddressMaster
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(488, 358)
        Me.Controls.Add(Me.pnlPhone)
        Me.Controls.Add(Me.pnlPhoneComment)
        Me.Controls.Add(Me.pnlExtension)
        Me.Controls.Add(Me.btnClearZip)
        Me.Controls.Add(Me.btnClearState)
        Me.Controls.Add(Me.btnClearCity)
        Me.Controls.Add(Me.btnClearCounty)
        Me.Controls.Add(Me.btnClearFIPS)
        Me.Controls.Add(Me.btnClearData)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.lblZip)
        Me.Controls.Add(Me.lblState)
        Me.Controls.Add(Me.lblCity)
        Me.Controls.Add(Me.lblCounty)
        Me.Controls.Add(Me.lblFIPS)
        Me.Controls.Add(Me.lblAddress2)
        Me.Controls.Add(Me.lblAddress1)
        Me.Controls.Add(Me.cboCounty)
        Me.Controls.Add(Me.cboZipCode)
        Me.Controls.Add(Me.cboState)
        Me.Controls.Add(Me.cboCity)
        Me.Controls.Add(Me.cboFIPS)
        Me.Controls.Add(Me.txtAddress2)
        Me.Controls.Add(Me.txtAddress1)
        Me.Name = "AddressMaster"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Address Info"
        CType(Me.cboCounty, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboZipCode, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboState, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboCity, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cboFIPS, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtCell, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtFax, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtPhone2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtPhone1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlExtension.ResumeLayout(False)
        Me.pnlPhoneComment.ResumeLayout(False)
        Me.pnlPhone.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Public Function ValidateData() As Boolean
        Try
            Dim errStr As String = ""
            Dim validateSuccess As Boolean = True

            If txtAddress1.Text.Trim = String.Empty Then
                errStr += "AddressLine1 cannot be empty" + vbCrLf
                validateSuccess = False
            End If
            If cboCity.Text.Trim = String.Empty Then
                errStr += "City cannot be empty" + vbCrLf
                validateSuccess = False
            End If
            If cboState.Text.Trim = String.Empty Then
                errStr += "State cannot be empty" + vbCrLf
                validateSuccess = False
            End If
            If cboZipCode.Text.Trim = String.Empty Then
                errStr += "Zip cannot be empty" + vbCrLf
                validateSuccess = False
            End If

            'If txtAddress1.Text <> String.Empty Then
            '    If cboCity.Text <> String.Empty Then
            '        If cboState.Text <> String.Empty Then
            '            If cboZipCode.Text <> String.Empty Then
            '                validateSuccess = True
            '            Else
            '                errStr += "Zip cannot be empty" + vbCrLf
            '                validateSuccess = False
            '            End If
            '        Else
            '            errStr += "State cannot be empty" + vbCrLf
            '            validateSuccess = False
            '        End If
            '    Else
            '        errStr += "City cannot be empty" + vbCrLf
            '        validateSuccess = False
            '    End If
            'Else
            '    errStr += "AddressLine1 cannot be empty" + vbCrLf
            '    validateSuccess = False
            'End If

            If errStr.Length > 0 Or Not validateSuccess Then
                MsgBox(errStr)
            End If
            Return validateSuccess
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Sub AddressMaster_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            bolLoading = True
            cboCounty.Visible = False
            cboFIPS.Visible = False
            btnClearCounty.Visible = False
            btnClearFIPS.Visible = False
            lblCounty.Visible = False
            lblFIPS.Visible = False

            If strEntityType = "Provider" Then
                ' pnlPhoneComment.Location = New System.Drawing.Point(208, 197)
                'pnlPhoneComment.Visible = True
                pnlPhoneComment.Visible = False
                pnlExtension.Visible = False
                pnlPhone.Visible = False

            Else
                pnlExtension.Visible = True
                pnlPhoneComment.Visible = False
            End If

            If strEntityType = "Licensee" OrElse strEntityType = "Manager" Then
                pnlExtension.Visible = False
                pnlPhone.Visible = False
            End If
            'pComAdd.PopulateAddressDetails()
            cboState.Text = "MS"
            strPrevState = "MS"
            If strMode = "MODIFY" Then

                If nCompanyID > 0 Or nCompanyID < -100 Then
                    If nAddressID <> 0 Then
                        oAddInfo = pComAdd.Retrieve(nAddressID, False)
                    Else
                        Select Case strEntityType
                            Case "Provider"
                                oAddInfo = pComAdd.GetAddressByType(0, 0, 0, nCompanyID)
                            Case "Company"
                                oAddInfo = pComAdd.GetAddressByType(0, nCompanyID, 0, 0)
                            Case "Licensee"
                                oAddInfo = pComAdd.GetAddressByType(0, 0, nCompanyID, 0)
                            Case "Manager"
                                oAddInfo = pComAdd.GetAddressByType(0, 0, nCompanyID, 0)
                        End Select
                    End If
                Else
                    If nAddressID <> 0 Then
                        oAddInfo = pComAdd.Retrieve(nAddressID, False)
                    End If
                End If

                'txtAdd.Text = oAddInfo.AddressId.ToString
                nAddressID = oAddInfo.AddressId

                'txtAddress1.Text = oAddInfo.AddressLine1.ToString
                'txtAddress2.Text = oAddInfo.AddressLine2.ToString
                'cboFIPS.Text = oAddInfo.FIPSCode
                'cboCity.Text = oAddInfo.City.ToString
                'cboState.Text = oAddInfo.State.ToString
                'cboZipCode.Text = oAddInfo.Zip
                'mskTxtZip.Mask = ""
                'mskTxtZip.CtlText = ""
                'mskTxtZip.Mask = "#####-####"

                'mskTxtZip.SelText = oAddInfo.Zip

                mskTxtPhone1.SelText = oAddInfo.Phone1
                mskTxtPhone2.SelText = oAddInfo.Phone2
                mskTxtCell.SelText = oAddInfo.Cell
                mskTxtFax.SelText = oAddInfo.Fax
                txtExt1.Text = oAddInfo.Ext1
                txtExt2.Text = oAddInfo.Ext2
                txtPhone1Comment.Text = oAddInfo.Phone1Comment
                txtPhone2Comment.Text = oAddInfo.Phone2Comment
                UpdateComboBoxes(oAddInfo.State, oAddInfo.City, oAddInfo.Zip)
            ElseIf strMode = "ADD" Then
                pComAdd.Retrieve(nAddressID, False)
                pComAdd.State = "MS"
                UpdateComboBoxes(pComAdd.State, , )
            End If

            'PopulateForm()
            bolLoading = False

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            'Add Address Details.
            ' pComAdd.Retrieve(nAddressID, False)
            If ValidateData() Then
                pComAdd.AddressLine1 = txtAddress1.Text.Trim
                pComAdd.AddressLine2 = txtAddress2.Text.Trim
                pComAdd.City = cboCity.Text.Trim
                pComAdd.State = cboState.Text.Trim
                pComAdd.Zip = cboZipCode.Text.Trim 'IIf(mskTxtZip.FormattedText.Substring(0, 5) Like "_____", String.Empty, mskTxtZip.FormattedText)
                pComAdd.FIPSCode = cboFIPS.Text.Trim
                pComAdd.Phone1 = mskTxtPhone1.FormattedText
                pComAdd.Phone2 = mskTxtPhone2.FormattedText
                pComAdd.Ext1 = txtExt1.Text.Trim
                pComAdd.Ext2 = txtExt2.Text.Trim
                pComAdd.Phone1Comment = txtPhone1Comment.Text.Trim
                pComAdd.Phone2Comment = txtPhone2Comment.Text.Trim
                pComAdd.Cell = mskTxtCell.FormattedText
                pComAdd.Fax = mskTxtFax.FormattedText
                Select Case strEntityType
                    Case "Provider"
                        pComAdd.ProviderID = nCompanyID
                    Case "Company"
                        pComAdd.CompanyId = nCompanyID
                    Case "Licensee"
                        pComAdd.LicenseeID = nCompanyID
                    Case "Manager"
                        pComAdd.LicenseeID = nCompanyID
                End Select

                If pComAdd.AddressId <= 0 Then
                    pComAdd.CreatedBy = MusterContainer.AppUser.ID
                Else
                    pComAdd.ModifiedBy = MusterContainer.AppUser.ID
                End If
                ' issue 3206
                If strEntityType = "Company" Or strEntityType = "Licensee" Or strEntityType = "Manager" Then
                    pComAdd.Save(CType(UIUtilsGen.ModuleID.Company, Integer), MusterContainer.AppUser.UserKey, returnVal)
                End If
                'If pComAdd.Add() Then
                bolSaved = True
                Me.Close()
                'End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.Close()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    'Private Sub clearcomboBox()
    '    pComAdd.FIPSCode = String.Empty
    '    pComAdd.Zip = String.Empty
    '    pComAdd.State = String.Empty
    '    pComAdd.City = String.Empty
    '    UpdateComboBoxes()
    'End Sub
    Private Sub btnClearData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearData.Click
        Dim strtmpMask As String
        Try
            'pComAdd.AddressLine1 = String.Empty
            'pComAdd.AddressLine2 = String.Empty

            pComAdd.City = String.Empty
            pComAdd.FIPSCode = String.Empty
            pComAdd.Zip = String.Empty
            pComAdd.State = String.Empty

            If pnlPhoneComment.Visible Then
                pComAdd.Phone1Comment = String.Empty
                pComAdd.Phone2Comment = String.Empty

                txtPhone1Comment.Text = String.Empty
                txtPhone2Comment.Text = String.Empty
            End If

            If pnlExtension.Visible Then
                pComAdd.Ext1 = String.Empty
                pComAdd.Ext2 = String.Empty

                txtExt1.Text = String.Empty
                txtExt2.Text = String.Empty
            End If

            If pnlPhone.Visible Then

                If Not strEntityType = "Company" Then
                    pComAdd.Phone1 = String.Empty
                    pComAdd.Phone2 = String.Empty
                    pComAdd.Fax = String.Empty
                    pComAdd.Cell = String.Empty

                    strtmpMask = mskTxtPhone1.Mask
                    mskTxtPhone1.Mask = ""
                    mskTxtPhone1.CtlText = ""
                    mskTxtPhone1.Mask = strtmpMask

                    strtmpMask = mskTxtPhone2.Mask
                    mskTxtPhone2.Mask = ""
                    mskTxtPhone2.CtlText = ""
                    mskTxtPhone2.Mask = strtmpMask

                    strtmpMask = mskTxtFax.Mask
                    mskTxtFax.Mask = ""
                    mskTxtFax.CtlText = ""
                    mskTxtFax.Mask = strtmpMask

                    strtmpMask = mskTxtCell.Mask
                    mskTxtCell.Mask = ""
                    mskTxtCell.CtlText = ""
                    mskTxtCell.Mask = strtmpMask
                End If

            End If

            'txtAddress1.Text = String.Empty
            'txtAddress2.Text = String.Empty
            'cboFIPS.Text = String.Empty
            'cboCity.Text = String.Empty
            'cboState.Text = String.Empty

            ''cboState.Text = "MS"

            'cboZipCode.Text = String.Empty
            ''mskTxtZip.Mask = ""
            ''mskTxtZip.CtlText = ""
            ''mskTxtZip.Mask = "#####-####"

            'strtmpMask = mskTxtPhone1.Mask
            'mskTxtPhone1.Mask = ""
            'mskTxtPhone1.CtlText = ""
            'mskTxtPhone1.Mask = strtmpMask
            ' mskTxtPhone1.CtlText = ""
            'strtmpMask = mskTxtPhone2.Mask
            'mskTxtPhone2.Mask = ""
            'mskTxtPhone2.CtlText = ""
            'mskTxtPhone2.Mask = strtmpMask
            'mskTxtPhone2.CtlText = ""
            'strtmpMask = mskTxtCell.Mask
            'mskTxtCell.Mask = ""
            'mskTxtCell.CtlText = ""
            'mskTxtCell.Mask = strtmpMask
            'mskTxtCell.CtlText = ""
            'strtmpMask = mskTxtFax.Mask
            'mskTxtFax.Mask = ""
            'mskTxtFax.CtlText = ""
            'mskTxtFax.Mask = strtmpMask
            'mskTxtFax.CtlText = ""
            'txtExt1.Text = String.Empty
            'txtExt2.Text = String.Empty
            'txtPhone1Comment.Text = String.Empty
            'txtPhone2Comment.Text = String.Empty
            'cboCounty.Text = String.Empty
            UpdateComboBoxes()
            'PopulateForm()
            txtAddress1.Focus()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    'Private Sub cboCity_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCity.AfterCloseUp
    '    Try
    '        If bolLoading Then Exit Sub
    '        If strPrevCity = String.Empty Or strPrevCity <> cboCity.Text Then
    '            pComAdd.PopulateAddressDetails(cboCity.Text, cboState.Text, "")
    '            strPrevCity = cboCity.Text
    '            cboZipCode.Value = String.Empty
    '        Else
    '            pComAdd.PopulateAddressDetails(cboCity.Text, cboState.Text, cboZipCode.Text)
    '        End If
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub cboState_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboState.AfterCloseUp
    '    Try
    '        If bolLoading Then Exit Sub
    '        If strPrevState = String.Empty Or strPrevState <> cboState.Text Then
    '            pComAdd.PopulateAddressDetails("", cboState.Text, "")
    '            strPrevState = cboState.Text
    '            cboCity.Value = String.Empty
    '            cboZipCode.Value = String.Empty
    '        Else
    '            pComAdd.PopulateAddressDetails(cboCity.Text, cboState.Text, cboZipCode.Text)
    '        End If
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub cboZipCode_AfterCloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboZipCode.AfterCloseUp
    '    Try
    '        If bolLoading Then Exit Sub
    '        If mskTxtZip.CtlText.StartsWith(cboZipCode.Text) Then Exit Sub

    '        mskTxtZip.Mask = ""
    '        mskTxtZip.CtlText = ""
    '        mskTxtZip.SelText = ""
    '        mskTxtZip.Mask = "#####-####"

    '        mskTxtZip.SelText = cboZipCode.Text
    '        mskTxtZip.Focus()
    '        pComAdd.PopulateAddressDetails(cboCity.Text, cboState.Text, cboZipCode.Text)
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub cboCity_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCity.Leave
    '    Try
    '        If bolLoading Then Exit Sub
    '        If strPrevCity = String.Empty Or strPrevCity <> cboCity.Text Then
    '            pComAdd.PopulateAddressDetails(cboCity.Text, cboState.Text, "")
    '            strPrevCity = cboCity.Text
    '            cboZipCode.Value = String.Empty
    '        Else
    '            pComAdd.PopulateAddressDetails(cboCity.Text, cboState.Text, cboZipCode.Text)
    '        End If

    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub pComAdd_CitiesChanged(ByVal dsCities As System.Data.DataSet) Handles pComAdd.CitiesChanged
    '    Dim strVal As String
    '    Try
    '        strVal = cboCity.Text
    '        cboCity.DataSource = dsCities.Tables(0).DefaultView
    '        cboCity.DisplayMember = "CITY"
    '        If cboCity.Rows.Count > 0 Then cboCity.Text = strVal
    '        If cboCity.Rows.Count = 1 Then
    '            pComAdd.City = cboCity.Rows.Item(0).Cells(0).Value
    '        End If
    '        cboCity.DataSource = Nothing
    '        cboCity.Text = String.Empty
    '        pComAdd.City = String.Empty
    '        PopulateForm()
    '    Catch ex As Exception
    '        Dim MyErr As ErrorReport
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub pComAdd_StateChanged(ByVal dsState As System.Data.DataSet) Handles pComAdd.StateChanged
    '    Dim strVal As String
    '    Try
    '        strVal = cboState.Text
    '        cboState.DataSource = dsState.Tables(0).DefaultView
    '        cboState.DisplayMember = "STATE"
    '        If cboState.Rows.Count > 0 Then cboState.Text = strVal
    '        If cboState.Rows.Count = 1 Then
    '            pComAdd.State = cboState.Rows.Item(0).Cells(0).Value
    '        End If
    '        PopulateForm()
    '    Catch ex As Exception
    '        Dim MyErr As ErrorReport
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub pComAdd_ZipChanged(ByVal dsZip As System.Data.DataSet) Handles pComAdd.ZipChanged
    '    Try
    '        cboZipCode.DataSource = dsZip.Tables(0).DefaultView
    '        cboZipCode.DisplayMember = "ZIP"

    '        If cboZipCode.Rows.Count > 0 Then
    '            cboZipCode.Rows(0).Selected = True
    '            If Not mskTxtZip.FormattedText.StartsWith(cboZipCode.Text) Then
    '                mskTxtZip.Mask = ""
    '                mskTxtZip.Text = ""
    '                mskTxtZip.SelText = ""
    '                mskTxtZip.Mask = "#####-####"
    '                mskTxtZip.Text = Trim(cboZipCode.Text)
    '            End If
    '            If cboZipCode.Rows.Count = 1 Then
    '                pComAdd.Zip = cboZipCode.Rows.Item(0).Cells(0).Value
    '            End If

    '            cboZipCode.DataSource = Nothing
    '            cboZipCode.Text = String.Empty
    '            pComAdd.Zip = String.Empty
    '            PopulateForm()
    '        End If
    '    Catch ex As Exception
    '        Dim MyErr As ErrorReport
    '        MyErr.ShowDialog()
    '    End Try

    'End Sub
    'Private Sub pComAdd_FipsChanged(ByVal strFIPS As String) Handles pComAdd.FipsChanged
    '    Dim strVal As String
    '    Try
    '        cboFIPS.Text = strFIPS
    '        If cboFIPS.Rows.Count = 1 Then
    '            pComAdd.FIPSCode = cboFIPS.Rows.Item(0).Cells(0).Value
    '        End If
    '        cboFIPS.DataSource = Nothing
    '        cboFIPS.Text = String.Empty
    '        PopulateForm()
    '    Catch ex As Exception
    '        Dim MyErr As ErrorReport
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    Private Sub btnClearCity_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClearCity.Click
        Try
            cboCity.Text = String.Empty
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnClearState_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClearState.Click
        Try
            cboState.Text = String.Empty
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr.ShowDialog()
        End Try

    End Sub
    Private Sub btnClearZip_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClearZip.Click
        Try

            cboZipCode.Text = String.Empty
            'mskTxtZip.Mask = ""
            'mskTxtZip.Text = ""
            'mskTxtZip.SelText = ""
            'mskTxtZip.CtlText = String.Empty
            'mskTxtZip.Mask = "#####-####"
            'mskTxtZip.Text = cboZipCode.Text
            'mskTxtZip.SelText = cboZipCode.Text
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub pComAdd_evtAddressErr(ByVal MsgStr As String) Handles pComAdd.evtAddressErr
        If MsgStr <> String.Empty And MsgBox(MsgStr) = MsgBoxResult.OK Then
            MsgStr = String.Empty
            bolValidationFlg = True
            Exit Sub
        End If
    End Sub
    Public Sub UpdateComboBoxes(Optional ByVal strState As String = "", Optional ByVal strCity As String = "", Optional ByVal strZip As String = "")
        Try
            Dim strStates, strCities, strCounties, strZips, strFips As String
            Dim strOldZip As String
            Dim strNewState As String
            Dim ds As DataSet
            'strNewState = IIf(strState <> String.Empty, strState, pComAdd.State.Trim)

            'strOldZip = pComAdd.Zip.Substring(0, pComAdd.Zip.Length - 5)
            '" AND FIPS LIKE '%" + pComAdd.FIPSCode.Trim + "%'" + _
            If strState = String.Empty Then
                strStates = "SELECT DISTINCT STATE FROM tblSYS_ZIPCODES WHERE" + _
                                " CITY LIKE '%" + strCity + "%'" + _
                                " AND ZIP LIKE '%" + strZip + "%'" + _
                                " ORDER BY STATE"
                cboState.DataSource = pComAdd.GetDataSet(strStates)
                If cboState.Rows.Count = 1 Then
                    pComAdd.State = cboState.Rows.Item(0).Cells(0).Value
                End If
            End If

            '" AND FIPS LIKE '%" + pComAdd.FIPSCode.Trim + "%'" + _
            If pComAdd.State.Trim <> String.Empty Then
                strCities = "SELECT DISTINCT CITY FROM tblSYS_ZIPCODES WHERE" + _
                                " STATE LIKE '%" + pComAdd.State.Trim + "%'" + _
                                " AND ZIP LIKE '%" + strZip + "%'" + _
                                " ORDER BY CITY"
                cboCity.DataSource = pComAdd.GetDataSet(strCities)
                If cboCity.Rows.Count = 1 Then
                    pComAdd.City = cboCity.Rows.Item(0).Cells(0).Value
                End If

                '" AND FIPS LIKE '%" + pComAdd.FIPSCode.Trim + "%'" + _
                strZips = "SELECT DISTINCT ZIP FROM tblSYS_ZIPCODES WHERE" + _
                                " STATE LIKE '%" + pComAdd.State.Trim + "%'" + _
                                " AND CITY LIKE '%" + strCity + "%'" + _
                                " ORDER BY ZIP"
                cboZipCode.DataSource = pComAdd.GetDataSet(strZips)
                If cboZipCode.Rows.Count = 1 Then
                    pComAdd.Zip = cboZipCode.Rows.Item(0).Cells(0).Value
                End If

                'If bolShowCounty Then
                '    strCounties = "SELECT DISTINCT COUNTY FROM tblSYS_ZIPCODES WHERE" + _
                '                    " STATE LIKE '%" + pAddress.State.Trim + "%'" + _
                '                    " AND CITY LIKE '%" + pAddress.City.Trim + "%'" + _
                '                    " AND ZIP LIKE '%" + pAddress.Zip.Trim + "%'" + _
                '                    " AND FIPS LIKE '%" + pAddress.FIPSCode.Trim + "%'" + _
                '                    " ORDER BY COUNTY"
                '    cboCounty.DataSource = pAddress.GetDataSet(strCounties)
                '    If cboCounty.Rows.Count = 1 Then
                '        pAddress.County = cboCounty.Rows.Item(0).Cells(0).Value
                '    End If
                'End If

                If bolShowFIPS Then
                    strFips = "SELECT DISTINCT FIPS FROM tblSYS_ZIPCODES WHERE" + _
                                    " STATE LIKE '%" + strNewState + "%'" + _
                                    " AND ZIP LIKE '%" + strZip + "%'" + _
                                    " AND CITY LIKE '%" + strCity + "%'" + _
                                    " ORDER BY FIPS"
                    cboFIPS.DataSource = pComAdd.GetDataSet(strFips)
                    If cboFIPS.Rows.Count = 1 Then
                        pComAdd.FIPSCode = cboFIPS.Rows.Item(0).Cells(0).Value
                    End If
                End If
            Else
                cboCity.DataSource = Nothing
                cboZipCode.DataSource = Nothing
                cboCounty.DataSource = Nothing
                cboFIPS.DataSource = Nothing
                cboCity.Text = String.Empty
                cboZipCode.Text = String.Empty
                cboCounty.Text = String.Empty
                cboFIPS.Text = String.Empty
                pComAdd.City = String.Empty
                pComAdd.Zip = String.Empty

            End If
            PopulateForm()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub cboCity_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCity.Leave
        pComAdd.City = cboCity.Text
        UpdateComboBoxes(pComAdd.State, cboCity.Text)
        'btnSave.Focus()

        'Try
        '    If bolLoading Then Exit Sub
        '    If strPrevCity = String.Empty Or strPrevCity <> cboCity.Text Then
        '        pComAdd.PopulateAddressDetails(cboCity.Text, cboState.Text, cboZipCode.Text)
        '        strPrevCity = cboCity.Text
        '        cboZipCode.Value = String.Empty
        '    Else
        '        pComAdd.PopulateAddressDetails(cboCity.Text, cboState.Text, cboZipCode.Text)
        '    End If
        'Catch ex As Exception
        '    Dim MyErr As New ErrorReport(ex)
        '    MyErr.ShowDialog()
        'End Try
    End Sub

    Private Sub cboZipCode_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboZipCode.Leave

        pComAdd.Zip = cboZipCode.Text
        UpdateComboBoxes(, , cboZipCode.Text)
        'mskTxtPhone1.Focus()
        'Try
        '    If bolLoading Then Exit Sub
        '    If mskTxtZip.CtlText.StartsWith(cboZipCode.Text) Then Exit Sub

        '    mskTxtZip.Mask = ""
        '    mskTxtZip.CtlText = ""
        '    mskTxtZip.SelText = ""
        '    mskTxtZip.Mask = "#####-####"

        '    mskTxtZip.SelText = cboZipCode.Text
        '    mskTxtZip.Focus()
        '    pComAdd.PopulateAddressDetails(cboCity.Text, cboState.Text, cboZipCode.Text)
        'Catch ex As Exception
        '    Dim MyErr As New ErrorReport(ex)
        '    MyErr.ShowDialog()
        'End Try
    End Sub

    Private Sub cboState_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboState.Leave
        Dim strState As String
        strState = pComAdd.State
        pComAdd.State = cboState.Text
        If strState <> pComAdd.State Then
            UpdateComboBoxes(cboState.Text, , )
        Else
            UpdateComboBoxes(cboState.Text, cboCity.Text, cboZipCode.Text)
        End If

        'Try
        '    If bolLoading Then Exit Sub
        '    If strPrevState = String.Empty Or strPrevState <> cboState.Text Then
        '        pComAdd.PopulateAddressDetails(cboCity.Text, cboState.Text, cboZipCode.Text)
        '        strPrevState = cboState.Text
        '        cboCity.Value = String.Empty
        '        cboZipCode.Value = String.Empty
        '    Else
        '        pComAdd.PopulateAddressDetails(cboCity.Text, cboState.Text, cboZipCode.Text)
        '    End If
        'Catch ex As Exception
        '    Dim MyErr As New ErrorReport(ex)
        '    MyErr.ShowDialog()
        'End Try
    End Sub

    Private Sub cboFIPS_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboFIPS.Leave
        pComAdd.FIPSCode = cboFIPS.Text
        UpdateComboBoxes(cboState.Text, cboCity.Text, cboZipCode.Text)
    End Sub
    Private Sub PopulateForm()

        Try
            txtAddress1.Text = pComAdd.AddressLine1
            txtAddress2.Text = pComAdd.AddressLine2
            cboState.Text = pComAdd.State
            If ContainsValue(cboCity, pComAdd.City) Then
                cboCity.Text = pComAdd.City
            Else
                cboCity.Text = String.Empty
                pComAdd.City = String.Empty
            End If
            'strZip = pComAdd.Zip.Substring(0, pComAdd.Zip.Length - 5)
            If ContainsValue(cboZipCode, pComAdd.Zip.Trim) Then
                cboZipCode.Text = pComAdd.Zip
            Else
                cboZipCode.Text = String.Empty
                pComAdd.Zip = String.Empty
            End If


            If ContainsValue(cboFIPS, pComAdd.FIPSCode) Then
                cboFIPS.Text = pComAdd.FIPSCode
            Else
                cboFIPS.Text = String.Empty
                pComAdd.FIPSCode = String.Empty
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Function ContainsValue(ByVal cbo As Infragistics.Win.UltraWinGrid.UltraCombo, ByVal value As String) As Boolean
        If cbo.Rows.Count > 0 Then
            For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In cbo.Rows
                If Not row.Cells(0).Value Is System.DBNull.Value Then
                    If UCase(row.Cells(0).Value) = UCase(value) Then
                        Return True
                    End If
                End If
            Next
        End If
        Return False
    End Function

    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)

        If pComAdd.IsDirty And Not bolSaved Then

            If pComAdd.AddressId <= 0 And (pComAdd.AddressLine1 = String.Empty Or pComAdd.City = String.Empty Or pComAdd.Zip = String.Empty) Then
                pComAdd.Reset()
                pComAdd.Remove(pComAdd.AddressId)
            ElseIf pComAdd.AddressId > 0 Then
                pComAdd.Reset()
            End If
            bolSaved = False
        End If
    End Sub

    Private Sub txtAddress1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAddress1.TextChanged
        pComAdd.AddressLine1 = txtAddress1.Text
    End Sub

    Private Sub txtAddress2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAddress2.TextChanged
        pComAdd.AddressLine2 = txtAddress2.Text
    End Sub
    Private Sub mskTxtPhone1_KeyUpEvent(ByVal sender As Object, ByVal e As AxMSMask.MaskEdBoxEvents_KeyUpEvent) Handles mskTxtPhone1.KeyUpEvent
        If bolLoading Then Exit Sub
        UIUtilsGen.FillStringObjectValues(pComAdd.Phone1, mskTxtPhone1.FormattedText.ToString)
    End Sub

    Private Sub mskTxtPhone2_KeyUpEvent(ByVal sender As Object, ByVal e As AxMSMask.MaskEdBoxEvents_KeyUpEvent) Handles mskTxtPhone2.KeyUpEvent
        If bolLoading Then Exit Sub
        UIUtilsGen.FillStringObjectValues(pComAdd.Phone2, mskTxtPhone2.FormattedText.ToString)
    End Sub

    Private Sub txtExt1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExt1.TextChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.FillStringObjectValues(pComAdd.Ext1, txtExt1.Text.Trim)
    End Sub

    Private Sub txtExt2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExt2.TextChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.FillStringObjectValues(pComAdd.Ext2, txtExt2.Text.Trim)
    End Sub

    Private Sub txtPhone1Comment_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPhone1Comment.TextChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.FillStringObjectValues(pComAdd.Phone1Comment, txtPhone1Comment.Text.Trim)
    End Sub

    Private Sub txtPhone2Comment_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPhone2Comment.TextChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.FillStringObjectValues(pComAdd.Phone2Comment, txtPhone2Comment.Text.Trim)
    End Sub

    Private Sub mskTxtCell_KeyUpEvent(ByVal sender As Object, ByVal e As AxMSMask.MaskEdBoxEvents_KeyUpEvent) Handles mskTxtCell.KeyUpEvent
        If bolLoading Then Exit Sub
        UIUtilsGen.FillStringObjectValues(pComAdd.Cell, mskTxtCell.FormattedText.ToString)
    End Sub

    Private Sub mskTxtFax_KeyUpEvent(ByVal sender As Object, ByVal e As AxMSMask.MaskEdBoxEvents_KeyUpEvent) Handles mskTxtFax.KeyUpEvent
        If bolLoading Then Exit Sub
        UIUtilsGen.FillStringObjectValues(pComAdd.Fax, mskTxtFax.FormattedText.ToString)
    End Sub

End Class
