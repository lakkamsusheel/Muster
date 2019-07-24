Public Class Contacts
    Inherits System.Windows.Forms.Form
#Region "User Defined Variables"
    Private WithEvents pConStruct As MUSTER.BusinessLogic.pContactStruct
    Private WithEvents objCompSearch As CompanyContactSearch
    Private ConStructInfo As MUSTER.Info.ContactStructInfo
    Private ConDatumInfo As MUSTER.Info.ContactDatumInfo
    Private Selectedrow As Infragistics.Win.UltraWinGrid.UltraGridRow
    Private dtContactType As DataTable
    Private dtContactPreferredAddresses As DataTable
    Private dtContactPreferredAlias As DataTable

    Private dsChildContacts As DataSet
    Private dsContactAsocAliases As DataSet
    Private dsContactAddresses As DataSet
    Private lockMod As Boolean = False

    '------ local variables ---------------------------------
    Dim nContactID As Integer = 0
    Dim nEntityTypeID As Integer = 0
    Dim nEntityID As Int64 = 0
    Dim strModuleName As String = String.Empty
    Dim bolLoading As Boolean = True
    Friend strMode As String = String.Empty
    Friend bolAddChangedFromCompany As Boolean = False
    Dim strPrevCity As String
    Dim strPrevState As String
    Dim strFIPSCode As String
    Dim nModuleID As Integer = 0
    Dim nContactAssocID As Integer = 0
    Dim nEntityAssocID As Integer = 0
    Dim bolOtherInfoChanged As Boolean = False
    Dim strFlow As String
    Dim bolNewCompany As Boolean = False
    Friend Event ContactAdded() ' Used to tell the parent that a Contact was added.
    Friend Event NewCompanyAdded()  'Used to tell the parent that a new Company record was added. 
    Dim returnVal As String = String.Empty
    Dim strEnvelopeLabelAddress As String = String.Empty
    Dim strEnvelopeLabelName As String = String.Empty
    Dim strZip As String
    Dim arrAddress(4) As String
    '--------------------------------------------------------------------------------

#End Region

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub
    Public Sub New(ByVal EntityID As Int64, ByVal EntityType As Integer, ByVal strModule As String, Optional ByVal ContactID As Integer = 0, Optional ByVal Selrow As Infragistics.Win.UltraWinGrid.UltraGridRow = Nothing, Optional ByRef pContactStruct As MUSTER.BusinessLogic.pContactStruct = Nothing, Optional ByVal Mode As String = "ADD", Optional ByVal Flow As String = "FromModule", Optional ByVal bolCompany As Boolean = False)

        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        nEntityTypeID = EntityType
        nEntityID = EntityID
        nContactID = ContactID
        strModuleName = strModule
        Selectedrow = Selrow
        pConStruct = pContactStruct

        'pConStruct = pContactStruct
        strMode = Mode
        strFlow = Flow
        bolNewCompany = bolCompany
        nModuleID = pConStruct.GetModuleID(strModuleName)
        If UCase(strModuleName) = UCase("Financial") Then
            Me.lblVendor.Visible = True
            Me.txtVendorValue.Visible = True
        Else
            Me.lblVendor.Visible = False
            Me.txtVendorValue.Visible = False
        End If
        If strMode = "ADD" Then
            cmbPersonCompany.SelectedIndex = 0
            grpEntitySpecific.Enabled = True
            pnlAssociatedCompanyContacts.Enabled = True
            Me.btnSearchforCompany.Enabled = False
        Else
            Me.btnSearchforCompany.Enabled = True
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
    Friend WithEvents lblEmail As System.Windows.Forms.Label
    Friend WithEvents lblFax As System.Windows.Forms.Label
    Friend WithEvents txtEmail As System.Windows.Forms.TextBox
    Friend WithEvents lblCell As System.Windows.Forms.Label
    Friend WithEvents lblPhone2 As System.Windows.Forms.Label
    Friend WithEvents lblPhone1 As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents lblExt1 As System.Windows.Forms.Label
    Friend WithEvents txtExt1 As System.Windows.Forms.TextBox
    Friend WithEvents lblExt2 As System.Windows.Forms.Label
    Friend WithEvents txtExt2 As System.Windows.Forms.TextBox
    Friend WithEvents grpEntitySpecific As System.Windows.Forms.GroupBox
    Friend WithEvents chkActive As System.Windows.Forms.CheckBox
    Friend WithEvents lblDisplayAs As System.Windows.Forms.Label
    Friend WithEvents txtDisplayAs As System.Windows.Forms.TextBox
    Friend WithEvents cmbCC As System.Windows.Forms.ComboBox
    Friend WithEvents lblCC As System.Windows.Forms.Label
    Friend WithEvents cmbType As System.Windows.Forms.ComboBox
    Friend WithEvents lblType As System.Windows.Forms.Label
    Friend WithEvents pnlAssociatedCompanyContacts As System.Windows.Forms.Panel
    Friend WithEvents btnOk As System.Windows.Forms.Button
    Friend WithEvents lblAssociatedCompanyContacts As System.Windows.Forms.Label
    Friend WithEvents ugAssociatedCompanyContacts As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnSearchforCompany As System.Windows.Forms.Button
    Friend WithEvents lblContactAddress As System.Windows.Forms.Label
    Friend WithEvents txtCity As System.Windows.Forms.TextBox
    Friend WithEvents lblCity As System.Windows.Forms.Label
    Friend WithEvents lblState As System.Windows.Forms.Label
    Friend WithEvents lblZip As System.Windows.Forms.Label
    Public WithEvents mskTxtPhone1 As AxMSMask.AxMaskEdBox
    Public WithEvents mskTxtPhone2 As AxMSMask.AxMaskEdBox
    Public WithEvents mskTxtCell As AxMSMask.AxMaskEdBox
    Public WithEvents mskTxtFax As AxMSMask.AxMaskEdBox
    Friend WithEvents cmbPersonCompany As System.Windows.Forms.ComboBox
    Friend WithEvents lblPersonCompany As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents lblLastName As System.Windows.Forms.Label
    Friend WithEvents txtLastName As System.Windows.Forms.TextBox
    Friend WithEvents txtMiddleName As System.Windows.Forms.TextBox
    Friend WithEvents txtFirstName As System.Windows.Forms.TextBox
    Friend WithEvents lblMiddleName As System.Windows.Forms.Label
    Friend WithEvents lblFirstName As System.Windows.Forms.Label
    Friend WithEvents pnlPerson As System.Windows.Forms.Panel
    Friend WithEvents pnlCompany As System.Windows.Forms.Panel
    Friend WithEvents lblCompanyName As System.Windows.Forms.Label
    Friend WithEvents txtCompanyName As System.Windows.Forms.TextBox
    Friend WithEvents txtEmailPersonal As System.Windows.Forms.TextBox
    Friend WithEvents cmbTitle As System.Windows.Forms.ComboBox
    Friend WithEvents lblTitle As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmbSuffix As System.Windows.Forms.ComboBox
    Friend WithEvents btnClearCity As System.Windows.Forms.Button
    Friend WithEvents btnClearState As System.Windows.Forms.Button
    Friend WithEvents btnClearZip As System.Windows.Forms.Button
    Public WithEvents txtAddressTwo As System.Windows.Forms.TextBox
    Public WithEvents txtAddressOne As System.Windows.Forms.TextBox
    Friend WithEvents txtVendorValue As System.Windows.Forms.TextBox
    Friend WithEvents lblVendor As System.Windows.Forms.Label
    Friend WithEvents CmbCity As Infragistics.Win.UltraWinGrid.UltraCombo
    Friend WithEvents btnLabels As System.Windows.Forms.Button
    Friend WithEvents btnEnvelopes As System.Windows.Forms.Button
    Friend WithEvents CmbZip As Infragistics.Win.UltraWinGrid.UltraCombo
    Friend WithEvents cmbState As Infragistics.Win.UltraWinGrid.UltraCombo
    Friend WithEvents lblAliases As System.Windows.Forms.Label
    Friend WithEvents ugAliases As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents lbladdress As System.Windows.Forms.Label
    Friend WithEvents ugAddresses As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents BtnAddAlias As System.Windows.Forms.Button
    Friend WithEvents BtnDeleteAlias As System.Windows.Forms.Button
    Friend WithEvents BtnDeleteAddress As System.Windows.Forms.Button
    Friend WithEvents cmbAlias As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents cmbAddress As System.Windows.Forms.ComboBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents BtnSetAtMain As System.Windows.Forms.Button
    Friend WithEvents BtnUnassociateCompany As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Contacts))
        Me.lblEmail = New System.Windows.Forms.Label
        Me.lblFax = New System.Windows.Forms.Label
        Me.txtEmail = New System.Windows.Forms.TextBox
        Me.lblContactAddress = New System.Windows.Forms.Label
        Me.lblLastName = New System.Windows.Forms.Label
        Me.txtLastName = New System.Windows.Forms.TextBox
        Me.lblCell = New System.Windows.Forms.Label
        Me.lblPhone2 = New System.Windows.Forms.Label
        Me.lblPhone1 = New System.Windows.Forms.Label
        Me.lblExt1 = New System.Windows.Forms.Label
        Me.txtExt1 = New System.Windows.Forms.TextBox
        Me.lblExt2 = New System.Windows.Forms.Label
        Me.txtExt2 = New System.Windows.Forms.TextBox
        Me.grpEntitySpecific = New System.Windows.Forms.GroupBox
        Me.cmbAddress = New System.Windows.Forms.ComboBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.cmbAlias = New System.Windows.Forms.ComboBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.chkActive = New System.Windows.Forms.CheckBox
        Me.lblDisplayAs = New System.Windows.Forms.Label
        Me.txtDisplayAs = New System.Windows.Forms.TextBox
        Me.cmbCC = New System.Windows.Forms.ComboBox
        Me.lblCC = New System.Windows.Forms.Label
        Me.cmbType = New System.Windows.Forms.ComboBox
        Me.lblType = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnOk = New System.Windows.Forms.Button
        Me.pnlAssociatedCompanyContacts = New System.Windows.Forms.Panel
        Me.BtnSetAtMain = New System.Windows.Forms.Button
        Me.BtnDeleteAddress = New System.Windows.Forms.Button
        Me.BtnDeleteAlias = New System.Windows.Forms.Button
        Me.BtnAddAlias = New System.Windows.Forms.Button
        Me.lbladdress = New System.Windows.Forms.Label
        Me.ugAddresses = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.lblAliases = New System.Windows.Forms.Label
        Me.ugAliases = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.btnSearchforCompany = New System.Windows.Forms.Button
        Me.lblAssociatedCompanyContacts = New System.Windows.Forms.Label
        Me.ugAssociatedCompanyContacts = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.BtnUnassociateCompany = New System.Windows.Forms.Button
        Me.txtAddressTwo = New System.Windows.Forms.TextBox
        Me.txtCity = New System.Windows.Forms.TextBox
        Me.lblCity = New System.Windows.Forms.Label
        Me.lblState = New System.Windows.Forms.Label
        Me.lblZip = New System.Windows.Forms.Label
        Me.mskTxtPhone1 = New AxMSMask.AxMaskEdBox
        Me.mskTxtPhone2 = New AxMSMask.AxMaskEdBox
        Me.mskTxtCell = New AxMSMask.AxMaskEdBox
        Me.mskTxtFax = New AxMSMask.AxMaskEdBox
        Me.txtMiddleName = New System.Windows.Forms.TextBox
        Me.txtFirstName = New System.Windows.Forms.TextBox
        Me.lblMiddleName = New System.Windows.Forms.Label
        Me.lblFirstName = New System.Windows.Forms.Label
        Me.cmbPersonCompany = New System.Windows.Forms.ComboBox
        Me.lblPersonCompany = New System.Windows.Forms.Label
        Me.pnlPerson = New System.Windows.Forms.Panel
        Me.cmbSuffix = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.cmbTitle = New System.Windows.Forms.ComboBox
        Me.lblTitle = New System.Windows.Forms.Label
        Me.pnlCompany = New System.Windows.Forms.Panel
        Me.txtCompanyName = New System.Windows.Forms.TextBox
        Me.lblCompanyName = New System.Windows.Forms.Label
        Me.txtAddressOne = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.txtEmailPersonal = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.btnClearCity = New System.Windows.Forms.Button
        Me.btnClearState = New System.Windows.Forms.Button
        Me.btnClearZip = New System.Windows.Forms.Button
        Me.txtVendorValue = New System.Windows.Forms.TextBox
        Me.lblVendor = New System.Windows.Forms.Label
        Me.CmbCity = New Infragistics.Win.UltraWinGrid.UltraCombo
        Me.btnLabels = New System.Windows.Forms.Button
        Me.btnEnvelopes = New System.Windows.Forms.Button
        Me.CmbZip = New Infragistics.Win.UltraWinGrid.UltraCombo
        Me.cmbState = New Infragistics.Win.UltraWinGrid.UltraCombo
        Me.grpEntitySpecific.SuspendLayout()
        Me.pnlAssociatedCompanyContacts.SuspendLayout()
        CType(Me.ugAddresses, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ugAliases, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ugAssociatedCompanyContacts, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtPhone1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtPhone2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtCell, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtFax, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlPerson.SuspendLayout()
        Me.pnlCompany.SuspendLayout()
        CType(Me.CmbCity, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CmbZip, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmbState, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblEmail
        '
        Me.lblEmail.Location = New System.Drawing.Point(352, 104)
        Me.lblEmail.Name = "lblEmail"
        Me.lblEmail.Size = New System.Drawing.Size(88, 16)
        Me.lblEmail.TabIndex = 53
        Me.lblEmail.Text = "E-mail:"
        Me.lblEmail.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblFax
        '
        Me.lblFax.Location = New System.Drawing.Point(400, 80)
        Me.lblFax.Name = "lblFax"
        Me.lblFax.Size = New System.Drawing.Size(40, 16)
        Me.lblFax.TabIndex = 52
        Me.lblFax.Text = "Fax:"
        Me.lblFax.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtEmail
        '
        Me.txtEmail.Location = New System.Drawing.Point(448, 104)
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.Size = New System.Drawing.Size(136, 20)
        Me.txtEmail.TabIndex = 18
        Me.txtEmail.Text = ""
        '
        'lblContactAddress
        '
        Me.lblContactAddress.Location = New System.Drawing.Point(664, 32)
        Me.lblContactAddress.Name = "lblContactAddress"
        Me.lblContactAddress.Size = New System.Drawing.Size(80, 16)
        Me.lblContactAddress.TabIndex = 51
        Me.lblContactAddress.Text = "Address Two:"
        Me.lblContactAddress.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblLastName
        '
        Me.lblLastName.Location = New System.Drawing.Point(0, 80)
        Me.lblLastName.Name = "lblLastName"
        Me.lblLastName.Size = New System.Drawing.Size(88, 16)
        Me.lblLastName.TabIndex = 49
        Me.lblLastName.Text = "Last Name:"
        Me.lblLastName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtLastName
        '
        Me.txtLastName.Location = New System.Drawing.Point(88, 80)
        Me.txtLastName.Name = "txtLastName"
        Me.txtLastName.Size = New System.Drawing.Size(144, 20)
        Me.txtLastName.TabIndex = 3
        Me.txtLastName.Text = ""
        '
        'lblCell
        '
        Me.lblCell.Location = New System.Drawing.Point(408, 56)
        Me.lblCell.Name = "lblCell"
        Me.lblCell.Size = New System.Drawing.Size(32, 16)
        Me.lblCell.TabIndex = 48
        Me.lblCell.Text = "Cell:"
        Me.lblCell.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPhone2
        '
        Me.lblPhone2.Location = New System.Drawing.Point(384, 32)
        Me.lblPhone2.Name = "lblPhone2"
        Me.lblPhone2.Size = New System.Drawing.Size(56, 16)
        Me.lblPhone2.TabIndex = 47
        Me.lblPhone2.Text = "Phone 2:"
        Me.lblPhone2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPhone1
        '
        Me.lblPhone1.Location = New System.Drawing.Point(384, 8)
        Me.lblPhone1.Name = "lblPhone1"
        Me.lblPhone1.Size = New System.Drawing.Size(56, 16)
        Me.lblPhone1.TabIndex = 45
        Me.lblPhone1.Text = "Phone 1:"
        Me.lblPhone1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblExt1
        '
        Me.lblExt1.Location = New System.Drawing.Point(552, 8)
        Me.lblExt1.Name = "lblExt1"
        Me.lblExt1.Size = New System.Drawing.Size(24, 16)
        Me.lblExt1.TabIndex = 56
        Me.lblExt1.Text = "Ext:"
        Me.lblExt1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtExt1
        '
        Me.txtExt1.Location = New System.Drawing.Point(576, 8)
        Me.txtExt1.Name = "txtExt1"
        Me.txtExt1.Size = New System.Drawing.Size(40, 20)
        Me.txtExt1.TabIndex = 13
        Me.txtExt1.Text = ""
        '
        'lblExt2
        '
        Me.lblExt2.Location = New System.Drawing.Point(552, 32)
        Me.lblExt2.Name = "lblExt2"
        Me.lblExt2.Size = New System.Drawing.Size(24, 16)
        Me.lblExt2.TabIndex = 58
        Me.lblExt2.Text = "Ext:"
        Me.lblExt2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtExt2
        '
        Me.txtExt2.Location = New System.Drawing.Point(576, 32)
        Me.txtExt2.Name = "txtExt2"
        Me.txtExt2.Size = New System.Drawing.Size(40, 20)
        Me.txtExt2.TabIndex = 15
        Me.txtExt2.Text = ""
        '
        'grpEntitySpecific
        '
        Me.grpEntitySpecific.Controls.Add(Me.cmbAddress)
        Me.grpEntitySpecific.Controls.Add(Me.Label3)
        Me.grpEntitySpecific.Controls.Add(Me.cmbAlias)
        Me.grpEntitySpecific.Controls.Add(Me.Label2)
        Me.grpEntitySpecific.Controls.Add(Me.chkActive)
        Me.grpEntitySpecific.Controls.Add(Me.lblDisplayAs)
        Me.grpEntitySpecific.Controls.Add(Me.txtDisplayAs)
        Me.grpEntitySpecific.Controls.Add(Me.cmbCC)
        Me.grpEntitySpecific.Controls.Add(Me.lblCC)
        Me.grpEntitySpecific.Controls.Add(Me.cmbType)
        Me.grpEntitySpecific.Controls.Add(Me.lblType)
        Me.grpEntitySpecific.Location = New System.Drawing.Point(8, 448)
        Me.grpEntitySpecific.Name = "grpEntitySpecific"
        Me.grpEntitySpecific.Size = New System.Drawing.Size(952, 160)
        Me.grpEntitySpecific.TabIndex = 20
        Me.grpEntitySpecific.TabStop = False
        Me.grpEntitySpecific.Text = "Entity Specific"
        '
        'cmbAddress
        '
        Me.cmbAddress.Location = New System.Drawing.Point(560, 64)
        Me.cmbAddress.Name = "cmbAddress"
        Me.cmbAddress.Size = New System.Drawing.Size(248, 21)
        Me.cmbAddress.TabIndex = 73
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(456, 64)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(104, 16)
        Me.Label3.TabIndex = 74
        Me.Label3.Text = "Preferred Address:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbAlias
        '
        Me.cmbAlias.Location = New System.Drawing.Point(560, 24)
        Me.cmbAlias.Name = "cmbAlias"
        Me.cmbAlias.Size = New System.Drawing.Size(248, 21)
        Me.cmbAlias.TabIndex = 71
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(464, 24)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(96, 16)
        Me.Label2.TabIndex = 72
        Me.Label2.Text = "Preferred Alias:"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'chkActive
        '
        Me.chkActive.Checked = True
        Me.chkActive.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkActive.Location = New System.Drawing.Point(16, 96)
        Me.chkActive.Name = "chkActive"
        Me.chkActive.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkActive.Size = New System.Drawing.Size(80, 24)
        Me.chkActive.TabIndex = 3
        Me.chkActive.Text = "  :Active"
        '
        'lblDisplayAs
        '
        Me.lblDisplayAs.Location = New System.Drawing.Point(8, 72)
        Me.lblDisplayAs.Name = "lblDisplayAs"
        Me.lblDisplayAs.Size = New System.Drawing.Size(64, 16)
        Me.lblDisplayAs.TabIndex = 70
        Me.lblDisplayAs.Text = "Display As:"
        Me.lblDisplayAs.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtDisplayAs
        '
        Me.txtDisplayAs.Enabled = False
        Me.txtDisplayAs.Location = New System.Drawing.Point(80, 72)
        Me.txtDisplayAs.Name = "txtDisplayAs"
        Me.txtDisplayAs.Size = New System.Drawing.Size(144, 20)
        Me.txtDisplayAs.TabIndex = 2
        Me.txtDisplayAs.Text = ""
        '
        'cmbCC
        '
        Me.cmbCC.Items.AddRange(New Object() {"YES", "NO"})
        Me.cmbCC.Location = New System.Drawing.Point(80, 48)
        Me.cmbCC.Name = "cmbCC"
        Me.cmbCC.Size = New System.Drawing.Size(56, 21)
        Me.cmbCC.TabIndex = 1
        '
        'lblCC
        '
        Me.lblCC.Location = New System.Drawing.Point(40, 48)
        Me.lblCC.Name = "lblCC"
        Me.lblCC.Size = New System.Drawing.Size(32, 16)
        Me.lblCC.TabIndex = 67
        Me.lblCC.Text = "CC:"
        Me.lblCC.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbType
        '
        Me.cmbType.Location = New System.Drawing.Point(80, 24)
        Me.cmbType.Name = "cmbType"
        Me.cmbType.Size = New System.Drawing.Size(136, 21)
        Me.cmbType.TabIndex = 0
        '
        'lblType
        '
        Me.lblType.Location = New System.Drawing.Point(40, 24)
        Me.lblType.Name = "lblType"
        Me.lblType.Size = New System.Drawing.Size(32, 16)
        Me.lblType.TabIndex = 65
        Me.lblType.Text = "Type:"
        Me.lblType.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(88, 616)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(70, 26)
        Me.btnCancel.TabIndex = 23
        Me.btnCancel.Text = "Cancel"
        '
        'btnOk
        '
        Me.btnOk.Location = New System.Drawing.Point(8, 616)
        Me.btnOk.Name = "btnOk"
        Me.btnOk.Size = New System.Drawing.Size(70, 26)
        Me.btnOk.TabIndex = 22
        Me.btnOk.Text = "OK"
        '
        'pnlAssociatedCompanyContacts
        '
        Me.pnlAssociatedCompanyContacts.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlAssociatedCompanyContacts.Controls.Add(Me.BtnSetAtMain)
        Me.pnlAssociatedCompanyContacts.Controls.Add(Me.BtnDeleteAddress)
        Me.pnlAssociatedCompanyContacts.Controls.Add(Me.BtnDeleteAlias)
        Me.pnlAssociatedCompanyContacts.Controls.Add(Me.BtnAddAlias)
        Me.pnlAssociatedCompanyContacts.Controls.Add(Me.lbladdress)
        Me.pnlAssociatedCompanyContacts.Controls.Add(Me.ugAddresses)
        Me.pnlAssociatedCompanyContacts.Controls.Add(Me.lblAliases)
        Me.pnlAssociatedCompanyContacts.Controls.Add(Me.ugAliases)
        Me.pnlAssociatedCompanyContacts.Controls.Add(Me.btnSearchforCompany)
        Me.pnlAssociatedCompanyContacts.Controls.Add(Me.lblAssociatedCompanyContacts)
        Me.pnlAssociatedCompanyContacts.Controls.Add(Me.ugAssociatedCompanyContacts)
        Me.pnlAssociatedCompanyContacts.Controls.Add(Me.BtnUnassociateCompany)
        Me.pnlAssociatedCompanyContacts.Location = New System.Drawing.Point(8, 232)
        Me.pnlAssociatedCompanyContacts.Name = "pnlAssociatedCompanyContacts"
        Me.pnlAssociatedCompanyContacts.Size = New System.Drawing.Size(952, 208)
        Me.pnlAssociatedCompanyContacts.TabIndex = 21
        '
        'BtnSetAtMain
        '
        Me.BtnSetAtMain.Location = New System.Drawing.Point(656, 168)
        Me.BtnSetAtMain.Name = "BtnSetAtMain"
        Me.BtnSetAtMain.Size = New System.Drawing.Size(120, 26)
        Me.BtnSetAtMain.TabIndex = 142
        Me.BtnSetAtMain.Text = "Set Address As Main"
        '
        'BtnDeleteAddress
        '
        Me.BtnDeleteAddress.Location = New System.Drawing.Point(776, 168)
        Me.BtnDeleteAddress.Name = "BtnDeleteAddress"
        Me.BtnDeleteAddress.Size = New System.Drawing.Size(96, 26)
        Me.BtnDeleteAddress.TabIndex = 141
        Me.BtnDeleteAddress.Text = "Delete Address"
        '
        'BtnDeleteAlias
        '
        Me.BtnDeleteAlias.Location = New System.Drawing.Point(440, 168)
        Me.BtnDeleteAlias.Name = "BtnDeleteAlias"
        Me.BtnDeleteAlias.Size = New System.Drawing.Size(80, 26)
        Me.BtnDeleteAlias.TabIndex = 140
        Me.BtnDeleteAlias.Text = "Delete Alias"
        '
        'BtnAddAlias
        '
        Me.BtnAddAlias.Location = New System.Drawing.Point(352, 168)
        Me.BtnAddAlias.Name = "BtnAddAlias"
        Me.BtnAddAlias.Size = New System.Drawing.Size(88, 26)
        Me.BtnAddAlias.TabIndex = 139
        Me.BtnAddAlias.Text = "Add Alias"
        '
        'lbladdress
        '
        Me.lbladdress.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbladdress.Location = New System.Drawing.Point(648, 0)
        Me.lbladdress.Name = "lbladdress"
        Me.lbladdress.Size = New System.Drawing.Size(240, 16)
        Me.lbladdress.TabIndex = 138
        Me.lbladdress.Text = "Associated Contact Alias Addresses"
        '
        'ugAddresses
        '
        Me.ugAddresses.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugAddresses.Location = New System.Drawing.Point(656, 24)
        Me.ugAddresses.Name = "ugAddresses"
        Me.ugAddresses.Size = New System.Drawing.Size(280, 144)
        Me.ugAddresses.TabIndex = 137
        '
        'lblAliases
        '
        Me.lblAliases.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAliases.Location = New System.Drawing.Point(344, 0)
        Me.lblAliases.Name = "lblAliases"
        Me.lblAliases.Size = New System.Drawing.Size(176, 16)
        Me.lblAliases.TabIndex = 136
        Me.lblAliases.Text = "Associated Contact Aliases"
        '
        'ugAliases
        '
        Me.ugAliases.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugAliases.Location = New System.Drawing.Point(352, 24)
        Me.ugAliases.Name = "ugAliases"
        Me.ugAliases.Size = New System.Drawing.Size(296, 144)
        Me.ugAliases.TabIndex = 135
        '
        'btnSearchforCompany
        '
        Me.btnSearchforCompany.Location = New System.Drawing.Point(16, 168)
        Me.btnSearchforCompany.Name = "btnSearchforCompany"
        Me.btnSearchforCompany.Size = New System.Drawing.Size(144, 26)
        Me.btnSearchforCompany.TabIndex = 1
        Me.btnSearchforCompany.Text = "Search For/Add Company"
        '
        'lblAssociatedCompanyContacts
        '
        Me.lblAssociatedCompanyContacts.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAssociatedCompanyContacts.Location = New System.Drawing.Point(16, 0)
        Me.lblAssociatedCompanyContacts.Name = "lblAssociatedCompanyContacts"
        Me.lblAssociatedCompanyContacts.Size = New System.Drawing.Size(176, 16)
        Me.lblAssociatedCompanyContacts.TabIndex = 134
        Me.lblAssociatedCompanyContacts.Text = "Associated Company Contacts"
        '
        'ugAssociatedCompanyContacts
        '
        Me.ugAssociatedCompanyContacts.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugAssociatedCompanyContacts.Location = New System.Drawing.Point(16, 24)
        Me.ugAssociatedCompanyContacts.Name = "ugAssociatedCompanyContacts"
        Me.ugAssociatedCompanyContacts.Size = New System.Drawing.Size(320, 144)
        Me.ugAssociatedCompanyContacts.TabIndex = 0
        '
        'BtnUnassociateCompany
        '
        Me.BtnUnassociateCompany.Enabled = False
        Me.BtnUnassociateCompany.Location = New System.Drawing.Point(160, 168)
        Me.BtnUnassociateCompany.Name = "BtnUnassociateCompany"
        Me.BtnUnassociateCompany.Size = New System.Drawing.Size(128, 26)
        Me.BtnUnassociateCompany.TabIndex = 141
        Me.BtnUnassociateCompany.Text = "Unassociate Company"
        '
        'txtAddressTwo
        '
        Me.txtAddressTwo.Location = New System.Drawing.Point(752, 32)
        Me.txtAddressTwo.Name = "txtAddressTwo"
        Me.txtAddressTwo.Size = New System.Drawing.Size(176, 20)
        Me.txtAddressTwo.TabIndex = 4
        Me.txtAddressTwo.Text = ""
        '
        'txtCity
        '
        Me.txtCity.Location = New System.Drawing.Point(904, 72)
        Me.txtCity.Name = "txtCity"
        Me.txtCity.Size = New System.Drawing.Size(16, 20)
        Me.txtCity.TabIndex = 4
        Me.txtCity.Text = ""
        Me.txtCity.Visible = False
        '
        'lblCity
        '
        Me.lblCity.Location = New System.Drawing.Point(704, 56)
        Me.lblCity.Name = "lblCity"
        Me.lblCity.Size = New System.Drawing.Size(40, 16)
        Me.lblCity.TabIndex = 160
        Me.lblCity.Text = "City:"
        Me.lblCity.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblState
        '
        Me.lblState.Location = New System.Drawing.Point(696, 80)
        Me.lblState.Name = "lblState"
        Me.lblState.Size = New System.Drawing.Size(48, 16)
        Me.lblState.TabIndex = 159
        Me.lblState.Text = "State:"
        Me.lblState.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblZip
        '
        Me.lblZip.Location = New System.Drawing.Point(704, 104)
        Me.lblZip.Name = "lblZip"
        Me.lblZip.Size = New System.Drawing.Size(40, 16)
        Me.lblZip.TabIndex = 158
        Me.lblZip.Text = "Zip:"
        Me.lblZip.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'mskTxtPhone1
        '
        Me.mskTxtPhone1.Location = New System.Drawing.Point(448, 8)
        Me.mskTxtPhone1.Name = "mskTxtPhone1"
        Me.mskTxtPhone1.OcxState = CType(resources.GetObject("mskTxtPhone1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtPhone1.Size = New System.Drawing.Size(96, 20)
        Me.mskTxtPhone1.TabIndex = 12
        '
        'mskTxtPhone2
        '
        Me.mskTxtPhone2.Location = New System.Drawing.Point(448, 32)
        Me.mskTxtPhone2.Name = "mskTxtPhone2"
        Me.mskTxtPhone2.OcxState = CType(resources.GetObject("mskTxtPhone2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtPhone2.Size = New System.Drawing.Size(96, 20)
        Me.mskTxtPhone2.TabIndex = 14
        '
        'mskTxtCell
        '
        Me.mskTxtCell.Location = New System.Drawing.Point(448, 56)
        Me.mskTxtCell.Name = "mskTxtCell"
        Me.mskTxtCell.OcxState = CType(resources.GetObject("mskTxtCell.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtCell.Size = New System.Drawing.Size(96, 20)
        Me.mskTxtCell.TabIndex = 16
        '
        'mskTxtFax
        '
        Me.mskTxtFax.Location = New System.Drawing.Point(448, 80)
        Me.mskTxtFax.Name = "mskTxtFax"
        Me.mskTxtFax.OcxState = CType(resources.GetObject("mskTxtFax.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtFax.Size = New System.Drawing.Size(96, 20)
        Me.mskTxtFax.TabIndex = 17
        '
        'txtMiddleName
        '
        Me.txtMiddleName.Location = New System.Drawing.Point(88, 56)
        Me.txtMiddleName.Name = "txtMiddleName"
        Me.txtMiddleName.Size = New System.Drawing.Size(144, 20)
        Me.txtMiddleName.TabIndex = 2
        Me.txtMiddleName.Text = ""
        '
        'txtFirstName
        '
        Me.txtFirstName.Location = New System.Drawing.Point(88, 32)
        Me.txtFirstName.Name = "txtFirstName"
        Me.txtFirstName.Size = New System.Drawing.Size(144, 20)
        Me.txtFirstName.TabIndex = 1
        Me.txtFirstName.Text = ""
        '
        'lblMiddleName
        '
        Me.lblMiddleName.Location = New System.Drawing.Point(-8, 56)
        Me.lblMiddleName.Name = "lblMiddleName"
        Me.lblMiddleName.Size = New System.Drawing.Size(96, 16)
        Me.lblMiddleName.TabIndex = 168
        Me.lblMiddleName.Text = "Middle Name:"
        Me.lblMiddleName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblFirstName
        '
        Me.lblFirstName.Location = New System.Drawing.Point(16, 32)
        Me.lblFirstName.Name = "lblFirstName"
        Me.lblFirstName.Size = New System.Drawing.Size(72, 16)
        Me.lblFirstName.TabIndex = 167
        Me.lblFirstName.Text = "First Name:"
        Me.lblFirstName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbPersonCompany
        '
        Me.cmbPersonCompany.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbPersonCompany.Items.AddRange(New Object() {"Person", "Company"})
        Me.cmbPersonCompany.Location = New System.Drawing.Point(104, 8)
        Me.cmbPersonCompany.Name = "cmbPersonCompany"
        Me.cmbPersonCompany.Size = New System.Drawing.Size(120, 21)
        Me.cmbPersonCompany.TabIndex = 0
        '
        'lblPersonCompany
        '
        Me.lblPersonCompany.Location = New System.Drawing.Point(0, 8)
        Me.lblPersonCompany.Name = "lblPersonCompany"
        Me.lblPersonCompany.Size = New System.Drawing.Size(96, 16)
        Me.lblPersonCompany.TabIndex = 171
        Me.lblPersonCompany.Text = "Person/Company:"
        Me.lblPersonCompany.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlPerson
        '
        Me.pnlPerson.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlPerson.Controls.Add(Me.cmbSuffix)
        Me.pnlPerson.Controls.Add(Me.Label1)
        Me.pnlPerson.Controls.Add(Me.cmbTitle)
        Me.pnlPerson.Controls.Add(Me.txtLastName)
        Me.pnlPerson.Controls.Add(Me.lblTitle)
        Me.pnlPerson.Controls.Add(Me.txtFirstName)
        Me.pnlPerson.Controls.Add(Me.txtMiddleName)
        Me.pnlPerson.Controls.Add(Me.lblMiddleName)
        Me.pnlPerson.Controls.Add(Me.lblFirstName)
        Me.pnlPerson.Controls.Add(Me.lblLastName)
        Me.pnlPerson.Location = New System.Drawing.Point(8, 80)
        Me.pnlPerson.Name = "pnlPerson"
        Me.pnlPerson.Size = New System.Drawing.Size(312, 136)
        Me.pnlPerson.TabIndex = 1
        '
        'cmbSuffix
        '
        Me.cmbSuffix.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbSuffix.ItemHeight = 13
        Me.cmbSuffix.Items.AddRange(New Object() {"Jr", "Sr", "I", "II", "III", "IV", "V", "VI"})
        Me.cmbSuffix.Location = New System.Drawing.Point(88, 104)
        Me.cmbSuffix.Name = "cmbSuffix"
        Me.cmbSuffix.Size = New System.Drawing.Size(48, 21)
        Me.cmbSuffix.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(48, 104)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(40, 16)
        Me.Label1.TabIndex = 170
        Me.Label1.Text = "Suffix:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbTitle
        '
        Me.cmbTitle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbTitle.ItemHeight = 13
        Me.cmbTitle.Items.AddRange(New Object() {"Mr", "Mrs", "Ms", "Dr", "Sir"})
        Me.cmbTitle.Location = New System.Drawing.Point(88, 8)
        Me.cmbTitle.Name = "cmbTitle"
        Me.cmbTitle.Size = New System.Drawing.Size(48, 21)
        Me.cmbTitle.TabIndex = 0
        '
        'lblTitle
        '
        Me.lblTitle.Location = New System.Drawing.Point(48, 8)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(40, 16)
        Me.lblTitle.TabIndex = 165
        Me.lblTitle.Text = "Title:"
        Me.lblTitle.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlCompany
        '
        Me.pnlCompany.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlCompany.Controls.Add(Me.txtCompanyName)
        Me.pnlCompany.Controls.Add(Me.lblCompanyName)
        Me.pnlCompany.Location = New System.Drawing.Point(8, 40)
        Me.pnlCompany.Name = "pnlCompany"
        Me.pnlCompany.Size = New System.Drawing.Size(312, 32)
        Me.pnlCompany.TabIndex = 2
        Me.pnlCompany.Visible = False
        '
        'txtCompanyName
        '
        Me.txtCompanyName.Location = New System.Drawing.Point(104, 8)
        Me.txtCompanyName.Name = "txtCompanyName"
        Me.txtCompanyName.Size = New System.Drawing.Size(192, 20)
        Me.txtCompanyName.TabIndex = 0
        Me.txtCompanyName.Text = ""
        '
        'lblCompanyName
        '
        Me.lblCompanyName.Location = New System.Drawing.Point(8, 8)
        Me.lblCompanyName.Name = "lblCompanyName"
        Me.lblCompanyName.Size = New System.Drawing.Size(96, 16)
        Me.lblCompanyName.TabIndex = 74
        Me.lblCompanyName.Text = "Company Name:"
        Me.lblCompanyName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtAddressOne
        '
        Me.txtAddressOne.Location = New System.Drawing.Point(752, 8)
        Me.txtAddressOne.Name = "txtAddressOne"
        Me.txtAddressOne.Size = New System.Drawing.Size(176, 20)
        Me.txtAddressOne.TabIndex = 3
        Me.txtAddressOne.Text = ""
        '
        'Label9
        '
        Me.Label9.Location = New System.Drawing.Point(664, 8)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(80, 16)
        Me.Label9.TabIndex = 175
        Me.Label9.Text = "Address One:"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtEmailPersonal
        '
        Me.txtEmailPersonal.Location = New System.Drawing.Point(448, 128)
        Me.txtEmailPersonal.Name = "txtEmailPersonal"
        Me.txtEmailPersonal.Size = New System.Drawing.Size(136, 20)
        Me.txtEmailPersonal.TabIndex = 19
        Me.txtEmailPersonal.Text = ""
        '
        'Label10
        '
        Me.Label10.Location = New System.Drawing.Point(352, 128)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(88, 16)
        Me.Label10.TabIndex = 177
        Me.Label10.Text = "E-mail Personal:"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnClearCity
        '
        Me.btnClearCity.Image = CType(resources.GetObject("btnClearCity.Image"), System.Drawing.Image)
        Me.btnClearCity.Location = New System.Drawing.Point(912, 56)
        Me.btnClearCity.Name = "btnClearCity"
        Me.btnClearCity.Size = New System.Drawing.Size(17, 17)
        Me.btnClearCity.TabIndex = 6
        '
        'btnClearState
        '
        Me.btnClearState.Image = CType(resources.GetObject("btnClearState.Image"), System.Drawing.Image)
        Me.btnClearState.Location = New System.Drawing.Point(872, 80)
        Me.btnClearState.Name = "btnClearState"
        Me.btnClearState.Size = New System.Drawing.Size(17, 17)
        Me.btnClearState.TabIndex = 8
        '
        'btnClearZip
        '
        Me.btnClearZip.Image = CType(resources.GetObject("btnClearZip.Image"), System.Drawing.Image)
        Me.btnClearZip.Location = New System.Drawing.Point(872, 104)
        Me.btnClearZip.Name = "btnClearZip"
        Me.btnClearZip.Size = New System.Drawing.Size(17, 17)
        Me.btnClearZip.TabIndex = 11
        '
        'txtVendorValue
        '
        Me.txtVendorValue.Location = New System.Drawing.Point(448, 152)
        Me.txtVendorValue.Name = "txtVendorValue"
        Me.txtVendorValue.Size = New System.Drawing.Size(136, 20)
        Me.txtVendorValue.TabIndex = 178
        Me.txtVendorValue.Text = ""
        Me.txtVendorValue.Visible = False
        '
        'lblVendor
        '
        Me.lblVendor.Location = New System.Drawing.Point(360, 160)
        Me.lblVendor.Name = "lblVendor"
        Me.lblVendor.Size = New System.Drawing.Size(87, 16)
        Me.lblVendor.TabIndex = 179
        Me.lblVendor.Text = "Vendor Number:"
        Me.lblVendor.Visible = False
        '
        'CmbCity
        '
        Me.CmbCity.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.CmbCity.Cursor = System.Windows.Forms.Cursors.Default
        Me.CmbCity.DisplayMember = ""
        Me.CmbCity.Location = New System.Drawing.Point(752, 56)
        Me.CmbCity.Name = "CmbCity"
        Me.CmbCity.Size = New System.Drawing.Size(160, 21)
        Me.CmbCity.TabIndex = 180
        Me.CmbCity.ValueMember = ""
        '
        'btnLabels
        '
        Me.btnLabels.Location = New System.Drawing.Point(680, 168)
        Me.btnLabels.Name = "btnLabels"
        Me.btnLabels.Size = New System.Drawing.Size(64, 23)
        Me.btnLabels.TabIndex = 1062
        Me.btnLabels.Text = "Labels"
        '
        'btnEnvelopes
        '
        Me.btnEnvelopes.Location = New System.Drawing.Point(680, 136)
        Me.btnEnvelopes.Name = "btnEnvelopes"
        Me.btnEnvelopes.Size = New System.Drawing.Size(65, 23)
        Me.btnEnvelopes.TabIndex = 1061
        Me.btnEnvelopes.Text = "Envelopes"
        '
        'CmbZip
        '
        Me.CmbZip.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.CmbZip.Cursor = System.Windows.Forms.Cursors.Default
        Me.CmbZip.DisplayMember = ""
        Me.CmbZip.Location = New System.Drawing.Point(752, 104)
        Me.CmbZip.Name = "CmbZip"
        Me.CmbZip.Size = New System.Drawing.Size(112, 21)
        Me.CmbZip.TabIndex = 1063
        Me.CmbZip.ValueMember = ""
        '
        'cmbState
        '
        Me.cmbState.CharacterCasing = System.Windows.Forms.CharacterCasing.Normal
        Me.cmbState.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmbState.DisplayMember = ""
        Me.cmbState.Location = New System.Drawing.Point(752, 80)
        Me.cmbState.Name = "cmbState"
        Me.cmbState.Size = New System.Drawing.Size(112, 21)
        Me.cmbState.TabIndex = 1064
        Me.cmbState.ValueMember = ""
        '
        'Contacts
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(976, 661)
        Me.Controls.Add(Me.cmbState)
        Me.Controls.Add(Me.CmbZip)
        Me.Controls.Add(Me.btnLabels)
        Me.Controls.Add(Me.btnEnvelopes)
        Me.Controls.Add(Me.CmbCity)
        Me.Controls.Add(Me.lblVendor)
        Me.Controls.Add(Me.txtVendorValue)
        Me.Controls.Add(Me.txtEmailPersonal)
        Me.Controls.Add(Me.txtAddressOne)
        Me.Controls.Add(Me.txtCity)
        Me.Controls.Add(Me.txtAddressTwo)
        Me.Controls.Add(Me.txtExt2)
        Me.Controls.Add(Me.txtExt1)
        Me.Controls.Add(Me.txtEmail)
        Me.Controls.Add(Me.btnClearZip)
        Me.Controls.Add(Me.btnClearState)
        Me.Controls.Add(Me.btnClearCity)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.pnlCompany)
        Me.Controls.Add(Me.pnlPerson)
        Me.Controls.Add(Me.cmbPersonCompany)
        Me.Controls.Add(Me.lblPersonCompany)
        Me.Controls.Add(Me.mskTxtFax)
        Me.Controls.Add(Me.mskTxtCell)
        Me.Controls.Add(Me.mskTxtPhone2)
        Me.Controls.Add(Me.mskTxtPhone1)
        Me.Controls.Add(Me.lblCity)
        Me.Controls.Add(Me.lblState)
        Me.Controls.Add(Me.lblZip)
        Me.Controls.Add(Me.pnlAssociatedCompanyContacts)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnOk)
        Me.Controls.Add(Me.grpEntitySpecific)
        Me.Controls.Add(Me.lblExt2)
        Me.Controls.Add(Me.lblExt1)
        Me.Controls.Add(Me.lblEmail)
        Me.Controls.Add(Me.lblFax)
        Me.Controls.Add(Me.lblContactAddress)
        Me.Controls.Add(Me.lblCell)
        Me.Controls.Add(Me.lblPhone2)
        Me.Controls.Add(Me.lblPhone1)
        Me.Name = "Contacts"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Contacts"
        Me.grpEntitySpecific.ResumeLayout(False)
        Me.pnlAssociatedCompanyContacts.ResumeLayout(False)
        CType(Me.ugAddresses, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ugAliases, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ugAssociatedCompanyContacts, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtPhone1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtPhone2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtCell, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtFax, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlPerson.ResumeLayout(False)
        Me.pnlCompany.ResumeLayout(False)
        CType(Me.CmbCity, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CmbZip, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmbState, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region "Load Events"
    Public Sub LoadForm()
        Try
            bolLoading = True
            If bolNewCompany Then
                pConStruct.CompanyPopulateAddressDetails()
            Else
                pConStruct.PopulateAddressDetails()
            End If

            dsContactAddresses = pConStruct.GetContactAliasAddresses(Me.nContactID)

            If Not dsContactAddresses Is Nothing AndAlso dsContactAddresses.Tables.Count > 0 Then
                ugAddresses.DataSource = dsContactAddresses.Tables(0).DefaultView
            End If

            dsContactAsocAliases = pConStruct.GetContactAliases(nContactID)

            If Not dsContactAsocAliases Is Nothing AndAlso dsContactAsocAliases.Tables.Count > 0 Then
                ugAliases.DataSource = dsContactAsocAliases.Tables(0).DefaultView
            End If

            populateAddresses()
            populateAliases()


            bolLoading = True
            cmbState.Text = "MS"

            If strMode = "MODIFY" Then
                lockMod = True
                cmbPersonCompany.Enabled = False
                Dim oConStructInfo As MUSTER.Info.ContactStructInfo
                'oConStructInfo = pConStruct.Retrieve(CInt(Selectedrow.Cells("EntityASSOCID").Value))
                oConStructInfo = pConStruct.GetContactEntityAssociation(CInt(Selectedrow.Cells("EntityASSOCID").Value))



                Dim parentID As Integer = oConStructInfo.ParentContactID
                Dim AssocChildID As Integer = oConStructInfo.ChildContactID

                If parentID <> 0 Then

                    oConStructInfo.childContact = pConStruct.ContactDatum.Retrieve(parentID)
                    oConStructInfo.ChildContactID = 0

                    dsChildContacts = pConStruct.GetChildContacts(CInt(AssocChildID))
                    dsChildContacts.Tables(0).DefaultView.RowFilter = "ContactAssocID= " + oConStructInfo.ContactAssocID.ToString + ""
                    ugAssociatedCompanyContacts.DataSource = dsChildContacts.Tables(0).DefaultView


                End If

                oconstructinfo.parentContact = Nothing

                Dim c As New BusinessLogic.pContactDatum
                oConStructInfo.parentContact = c.Retrieve(AssocChildID)
                c = Nothing

                populateContactType()
                PopulateContactInfo(oConStructInfo.parentContact)
                cmbType.SelectedValue = pConStruct.ConTypeID
                cmbCC.Text = pConStruct.ccInfo
                cmbAlias.SelectedValue = pConStruct.PreferredAlias
                cmbAddress.SelectedValue = pConStruct.PreferredAddress
                txtDisplayAs.Text = pConStruct.displayAs
                chkActive.Checked = pConStruct.Active
                nContactAssocID = oConStructInfo.ContactAssocID
                nEntityAssocID = oConStructInfo.entityAssocID

                'pConStruct.PopulateAddressDetails(CmbCity.Text, cmbState.Text, CmbZip.Text)
                Me.UpdateComboBoxes(cmbState.Text, CmbCity.Text, CmbZip.Text)
                CheckActiveInactive()
            ElseIf strMode = "ASSOCIATE" Then
                cmbPersonCompany.Enabled = False
                Dim oConDatum As MUSTER.Info.ContactDatumInfo
                oConDatum = pConStruct.ContactDatum.Retrieve(CInt(Selectedrow.Cells("ContactID").Value))

                Dim oConStructInfo As Info.ContactStructInfo = pConStruct.contactStructInfo

                oConStructInfo.parentContact = pConStruct.ContactDatum.Retrieve(oconstructinfo.ChildContactID)



                populateContactType()
                PopulateContactInfo(oConDatum)
                'nContactAssocID = CInt(Selectedrow.Cells("ContactAssocID").Value)
                If Not Selectedrow.Cells("ContactID").Value Is System.DBNull.Value Then
                    nContactID = CInt(Selectedrow.Cells("ContactID").Value)


                    dsContactAsocAliases = pConStruct.GetContactAliases(nContactID)

                    If Not dsContactAsocAliases Is Nothing AndAlso dsContactAsocAliases.Tables.Count > 0 Then
                        ugAliases.DataSource = dsContactAsocAliases.Tables(0).DefaultView
                    End If

                End If
                'If Not strFlow.ToUpper = "FromSearch".ToUpper Then
                'nEntityAssocID = CInt(Selectedrow.Cells("EntityAssocID").Value)
                nEntityAssocID = 0
                'End If
                dsChildContacts = pConStruct.GetChildContacts(nContactID)
                ugAssociatedCompanyContacts.DataSource = dsChildContacts.Tables(0).DefaultView


            ElseIf strMode = "ADD" Then
                populateContactType()
                nEntityAssocID = 0
                UpdateComboBoxes(cmbState.Text, , )

            End If
            bolLoading = False

            '----- add a new company -----------------------------------------
            If bolNewCompany = True Then
                cmbPersonCompany.SelectedIndex = 1
                grpEntitySpecific.Enabled = False
                pnlAssociatedCompanyContacts.Visible = False
                Me.Height = 450
                btnOk.Location = New System.Drawing.Point(216, 370)
                btnCancel.Location = New System.Drawing.Point(296, 370)
                Me.Location = New System.Drawing.Point(200, 120)
                Me.Text = "Contacts - Add New Company"
            End If
            '-----------------------------------------------------------------

            If Me.ugAssociatedCompanyContacts.Rows.Count = 1 Then
                Me.BtnUnassociateCompany.Enabled = True
            Else
                Me.BtnUnassociateCompany.Enabled = False
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub
    Private Sub Contacts_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadForm()
    End Sub

    Private Sub populateContactType()
        If Selectedrow Is Nothing Then
            dtContactType = pConStruct.FilterContactTypes(nEntityID, nEntityTypeID, strModuleName, nModuleID)
        Else
            dtContactType = pConStruct.FilterContactTypes(nEntityID, nEntityTypeID, strModuleName, nModuleID, CInt(Selectedrow.Cells("ContactID").Value))
        End If

        If Not dtContactType Is Nothing Then
            cmbType.DisplayMember = "CONTACTTYPE"
            cmbType.ValueMember = "CONTACTTYPEID"
            dtContactType.DefaultView.RowFilter = "MODULEID = " + nModuleID.ToString
            dtContactType.DefaultView.Sort = "CONTACTTYPE ASC"
            cmbType.DataSource = dtContactType
            cmbType.SelectedIndex = -1
        End If
    End Sub


    Private Sub populateAddresses()
        If Not Selectedrow Is Nothing Then
            Dim ds As DataSet = pConStruct.GetContactAliasAddresses(CInt(Selectedrow.Cells("ContactID").Value), True)

            If Not ds Is Nothing AndAlso ds.Tables.Count > 0 Then
                Me.dtContactPreferredAddresses = ds.Tables(0)
            End If
        End If

        If Not Me.dtContactPreferredAddresses Is Nothing Then
            Me.cmbAddress.DisplayMember = "Address"
            cmbAddress.ValueMember = "address_id"
            cmbAddress.DataSource = dtContactPreferredAddresses
            cmbAddress.SelectedValue = pConStruct.PreferredAddress
        End If
    End Sub


    Private Sub populateAliases()
        If Not Selectedrow Is Nothing Then
            Dim ds As DataSet = pConStruct.GetContactAliases(CInt(Selectedrow.Cells("ContactID").Value), True)

            If Not ds Is Nothing AndAlso ds.Tables.Count > 0 Then
                Me.dtContactPreferredAlias = ds.Tables(0)
            End If
        End If

        If Not Me.dtContactPreferredAlias Is Nothing Then
            Me.cmbAlias.DisplayMember = "Alias"
            cmbAlias.ValueMember = "ContactAliasID"
            cmbAlias.DataSource = dtContactPreferredAlias
            cmbAlias.SelectedValue = pConStruct.PreferredAlias
        End If
    End Sub


    Private Sub PopulateContactInfo(ByVal ConDatumInfo As MUSTER.Info.ContactDatumInfo)
        strEnvelopeLabelName = String.Empty
        Try
            If Selectedrow.Cells("IsPerson").Value Then

                pnlPerson.Visible = True
                pnlCompany.Visible = False
                btnSearchforCompany.Enabled = True
                cmbPersonCompany.SelectedIndex = 0
                cmbTitle.Text = ConDatumInfo.Title
                txtFirstName.Text = ConDatumInfo.FirstName
                txtMiddleName.Text = ConDatumInfo.MiddleName
                txtLastName.Text = ConDatumInfo.LastName
                cmbSuffix.Text = ConDatumInfo.suffix
                strEnvelopeLabelName = IIf(ConDatumInfo.Title <> String.Empty, ConDatumInfo.Title + " ", "") + ConDatumInfo.FirstName + " " + IIf(ConDatumInfo.MiddleName <> String.Empty, ConDatumInfo.MiddleName + " ", "") + ConDatumInfo.LastName + IIf(ConDatumInfo.suffix <> String.Empty, " " + ConDatumInfo.suffix, "")
            Else
                pnlCompany.Location = New System.Drawing.Point(40, 48)
                pnlCompany.Visible = True
                pnlPerson.Visible = False
                btnSearchforCompany.Enabled = False
                cmbPersonCompany.SelectedIndex = 1
                txtCompanyName.Text = ConDatumInfo.companyName
                strEnvelopeLabelName = ConDatumInfo.companyName
            End If


            txtAddressOne.Text = ConDatumInfo.AddressLine1
            txtAddressTwo.Text = ConDatumInfo.AddressLine2
            CmbCity.Text = ConDatumInfo.City
            cmbState.Text = ConDatumInfo.State
            'mskTxtZip.SelText = ConDatumInfo.ZipCode
            CmbZip.Text = ConDatumInfo.ZipCode
            mskTxtPhone1.SelText = ConDatumInfo.Phone1
            mskTxtPhone2.SelText = ConDatumInfo.Phone2
            txtExt1.Text = ConDatumInfo.Ext1
            txtExt2.Text = ConDatumInfo.Ext2
            mskTxtCell.SelText = ConDatumInfo.Cell
            mskTxtFax.SelText = ConDatumInfo.Fax
            txtEmail.Text = ConDatumInfo.publicEmail
            txtEmailPersonal.Text = ConDatumInfo.privateEmail
            txtVendorValue.Text = ConDatumInfo.VendorNumber
            If strMode = "ASSOCIATE" Then
                EnableDisableControls(False)
            End If
            strEnvelopeLabelAddress = ConDatumInfo.AddressLine1 + "," + ConDatumInfo.AddressLine2 + "," + ConDatumInfo.City + "," + ConDatumInfo.State + "," + ConDatumInfo.ZipCode
            arrAddress(0) = ConDatumInfo.AddressLine1
            arrAddress(1) = ConDatumInfo.AddressLine2
            arrAddress(2) = ConDatumInfo.City
            arrAddress(3) = ConDatumInfo.State
            arrAddress(4) = ConDatumInfo.ZipCode
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub EnableDisableControls(ByVal bolState As Boolean)
        Try
            pnlCompany.Enabled = True
            pnlPerson.Enabled = bolState
            txtAddressOne.Enabled = bolState
            txtAddressTwo.Enabled = bolState
            CmbCity.Enabled = bolState
            CmbZip.Enabled = bolState
            'mskTxtZip.Enabled = bolState
            CmbZip.Enabled = bolState
            cmbState.Enabled = bolState
            mskTxtPhone1.Enabled = bolState
            mskTxtPhone2.Enabled = bolState
            txtExt1.Enabled = bolState
            txtExt2.Enabled = bolState
            mskTxtCell.Enabled = bolState
            mskTxtFax.Enabled = bolState
            txtEmail.Enabled = bolState
            txtEmailPersonal.Enabled = bolState
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub CheckActiveInactive(Optional ByVal bolCheckInacive As Boolean = True)
        Try
            If chkActive.Checked Then
                EnableDisableControls(True)
                cmbType.Enabled = True
                cmbCC.Enabled = True
                btnOk.Enabled = True
            ElseIf bolCheckInacive Then
                EnableDisableControls(False)
                cmbType.Enabled = False
                cmbCC.Enabled = False
                btnOk.Enabled = False
                MsgBox("Only Active Contacts can be modified")
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
#Region "Button Events"
    Public Sub SaveContact()
        Dim bolAddressDirty As Boolean = False
        Dim results As Long
        Dim nTempEntityAssocID As Integer = 0
        Try
            If ValidateData() Then


                If strMode = "ASSOCIATE NEW COMPANY" Then
                    AssociateNewCompany()
                    Exit Sub
                End If
                If bolAddChangedFromCompany = True Then
                    AddressChangedToCompanyAdress()
                    Exit Sub
                End If
                If txtAddressOne.Text = String.Empty And txtAddressTwo.Text = String.Empty And CmbCity.Text = String.Empty And CmbZip.Text = String.Empty Then 'FormatZIPCode() = String.Empty Then
                    cmbState.Text = String.Empty
                End If






                If strMode = "MODIFY" Then
                    'Save Contacts.
                    'ConDatumInfo = pConStruct.ContactDatum.ContactCollection.Item(nContactID)
                    ConDatumInfo = pConStruct.ContactDatum.Retrieve(nContactID)
                    pConStruct.contactStructInfo.parentContact.companyName = txtCompanyName.Text
                    pConStruct.contactStructInfo.parentContact.Title = cmbTitle.Text
                    pConStruct.contactStructInfo.parentContact.FirstName = txtFirstName.Text
                    pConStruct.contactStructInfo.parentContact.MiddleName = txtMiddleName.Text
                    pConStruct.contactStructInfo.parentContact.LastName = txtLastName.Text
                    pConStruct.contactStructInfo.parentContact.suffix = cmbSuffix.Text
                    pConStruct.contactStructInfo.parentContact.AddressLine1 = txtAddressOne.Text
                    pConStruct.contactStructInfo.parentContact.AddressLine2 = txtAddressTwo.Text
                    pConStruct.contactStructInfo.parentContact.City = CmbCity.Text
                    pConStruct.contactStructInfo.parentContact.State = cmbState.Text
                    pConStruct.contactStructInfo.parentContact.ZipCode = CmbZip.Text 'FormatZIPCode.Trim.ToString
                    pConStruct.contactStructInfo.parentContact.FipsCode = strFIPSCode
                    pConStruct.contactStructInfo.parentContact.publicEmail = txtEmail.Text
                    pConStruct.contactStructInfo.parentContact.privateEmail = txtEmailPersonal.Text
                    pConStruct.contactStructInfo.parentContact.Phone1 = mskTxtPhone1.FormattedText
                    pConStruct.contactStructInfo.parentContact.Ext1 = txtExt1.Text
                    pConStruct.contactStructInfo.parentContact.Phone2 = mskTxtPhone2.FormattedText
                    pConStruct.contactStructInfo.parentContact.Ext2 = txtExt2.Text
                    pConStruct.contactStructInfo.parentContact.Fax = mskTxtFax.FormattedText.Trim.ToString
                    pConStruct.contactStructInfo.parentContact.Cell = mskTxtCell.FormattedText.Trim.ToString
                    pConStruct.contactStructInfo.parentContact.modifiedBy = MusterContainer.AppUser.ID
                    pConStruct.contactStructInfo.parentContact.modifiedOn = Now()
                    pConStruct.contactStructInfo.PreferredAddress = cmbAddress.SelectedValue
                    pConStruct.contactStructInfo.PreferredAlias = cmbAlias.SelectedValue

                    bolAddressDirty = pConStruct.contactStructInfo.parentContact.IsAddressDirty()
                    If UCase(strModuleName) = UCase("Financial") Then
                        pConStruct.contactStructInfo.parentContact.VendorNumber = IIf(txtVendorValue.Text = "0", String.Empty, txtVendorValue.Text)
                    End If
                    ConDatumInfo = pConStruct.contactStructInfo.parentContact

                    'logic to validate modifications other than address.
                    If pConStruct.contactStructInfo.parentContact.IsOthersDirty Then
                        Dim count As Integer = pConStruct.IsAssociated(nContactID)
                        If count > 1 Then
                            Dim result As MsgBoxResult = MsgBox("Contact is associated with other entity(s) or Module(s). Do you want to save the information?", MsgBoxStyle.YesNo)
                            If result = MsgBoxResult.No Then
                                Exit Sub
                            End If
                        End If
                    End If
                    'logic ends

                    If pConStruct.contactStructInfo.entityAssocID > 0 Then
                        nTempEntityAssocID = pConStruct.contactStructInfo.entityAssocID
                    End If
                Else
                    If UIUtilsGen.GetComboBoxValueInt(cmbType) = 0 And bolNewCompany = False Then
                        MsgBox("Please enter the entity Type")
                        Exit Sub
                    End If
                    ''FormatZIPCode.Trim.ToString, _
                    ConDatumInfo = New MUSTER.Info.ContactDatumInfo(nContactID, _
                                                IIf(cmbPersonCompany.Text = "Person", True, False), _
                                                0, _
                                                txtCompanyName.Text, _
                                                cmbTitle.Text, _
                                                String.Empty, _
                                                txtFirstName.Text, _
                                                txtMiddleName.Text, _
                                                txtLastName.Text, _
                                                cmbSuffix.Text, _
                                                txtAddressOne.Text, _
                                                txtAddressTwo.Text, _
                                                CmbCity.Text, _
                                                cmbState.Text, _
                                                CmbZip.Text, _
                                                strFIPSCode, _
                                                mskTxtPhone1.FormattedText.Trim.ToString, _
                                                mskTxtPhone2.FormattedText.Trim.ToString, _
                                                txtExt1.Text, _
                                                txtExt2.Text, _
                                                mskTxtFax.FormattedText.Trim.ToString, _
                                                mskTxtCell.FormattedText.Trim.ToString, _
                                                txtEmail.Text, _
                                                txtEmailPersonal.Text, _
                                                IIf(txtVendorValue.Text = "0", String.Empty, txtVendorValue.Text), _
                                                IIf(nContactID <= 0, MusterContainer.AppUser.ID, ""), _
                                                Now, _
                                                IIf(nContactID > 0, MusterContainer.AppUser.ID, ""), _
                                                CDate("01/01/0001"), _
                                                False)

                End If
                pConStruct.ContactDatum.Add(ConDatumInfo)

                Dim success As Boolean = False
                success = pConStruct.ContactDatum.Save(nModuleID, MusterContainer.AppUser.UserKey, returnVal, , , bolAddressDirty)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                If Not success Then
                    Exit Sub
                End If

                ' Save Contact- Contact Relationship & Entity- Contact Relationship()
                ConStructInfo = New MUSTER.Info.ContactStructInfo(pConStruct.ContactDatum.ID, _
                                                                    pConStruct.childContactID, _
                                                                    nEntityID, _
                                                                    nEntityTypeID, _
                                                                    UIUtilsGen.GetComboBoxValueInt(cmbType), _
                                                                    nModuleID, _
                                                                    chkActive.Checked, _
                                                                    cmbCC.Text, _
                                                                    txtDisplayAs.Text, _
                                                                  False, cmbAddress.SelectedValue, cmbAlias.SelectedValue)
                'ConStructInfo.parentContact = pConStruct.ContactDatum.ContactCollection.Item(ConStructInfo.ParentContactID)
                ConStructInfo.parentContact = pConStruct.ContactDatum.Retrieve(ConStructInfo.ParentContactID)
                If ConStructInfo.ChildContactID <> 0 Then
                    'ConStructInfo.childContact = pConStruct.ContactDatum.ContactCollection.Item(ConStructInfo.ChildContactID)
                    ConStructInfo.childContact = pConStruct.ContactDatum.Retrieve(ConStructInfo.ChildContactID)
                Else
                    ConStructInfo.childContact = Nothing
                End If

                If pConStruct.ContactAssocID > 0 And strMode = "ASSOCIATE NEW" Then
                    ConStructInfo.ContactAssocID = pConStruct.ContactAssocID()
                End If
                If (nTempEntityAssocID > 0 Or nEntityAssocID > 0) Then
                    ConStructInfo.entityAssocID = IIf(nTempEntityAssocID = 0, nEntityAssocID, nTempEntityAssocID)
                End If
                pConStruct.Add(ConStructInfo)


                pConStruct.PreferredAlias = Me.cmbAlias.SelectedValue
                pConStruct.PreferredAddress = Me.cmbAddress.SelectedValue

                If (pConStruct.childContactID > 0) Then

                    If strMode = "ASSOCIATE NEW" Then
                        pConStruct.Save(nModuleID, MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID, "ASSOCIATE", 0, pConStruct.ContactAssocID, pConStruct.entityID, pConStruct.entityType, pConStruct.ModuleID, , )
                    End If
                    If bolAddChangedFromCompany = True Or strMode = "MODIFY" Then
                        pConStruct.Save(nModuleID, MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
                        If Not UIUtilsGen.HasRights(returnVal) Then
                            Exit Sub
                        End If

                        bolAddChangedFromCompany = False
                    End If
                    If strMode = "ASSOCIATE" And Not bolAddChangedFromCompany Then
                        pConStruct.Save(nModuleID, MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID, strMode, nEntityAssocID, nContactAssocID, nEntityID, nEntityTypeID, nModuleID, , )
                    End If
                Else
                    pConStruct.Save(nModuleID, MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID, strMode, nEntityAssocID, nContactAssocID, nEntityID, nEntityTypeID, nModuleID, , )

                End If

                '----- the contact is saved ----------------------
                If Me.Visible Then
                    MsgBox("Contact is saved successfully!")
                End If

                '-----------------------------------------------------------------------------------------------------

                '----- check to see if there is to be a company associated with this contact -------------------------

                If strMode = "ADD" AndAlso Me.Visible Then
                    If bolNewCompany = False And Me.cmbPersonCompany.SelectedIndex <> 1 Then
                        results = MsgBox("Do you want to associate a Company with this new Contact?", MsgBoxStyle.Question & MsgBoxStyle.YesNoCancel, "Associate Contact")
                        If results = MsgBoxResult.Yes Then
                            strMode = "ASSOCIATE NEW"

                            Me.btnSearchforCompany.Enabled = True
                            nContactID = pConStruct.ContactDatum.ID
                            '---- display the company contact search screen ----------------------------
                            objCompSearch = New CompanyContactSearch(nContactID, nEntityID, nEntityTypeID, strModuleName, pConStruct, Me)
                            objCompSearch.ShowDialog()
                            objCompSearch.BringToFront()

                            If Not Selectedrow Is Nothing Then
                                ugAssociatedCompanyContacts.DataSource = pConStruct.GetChildContacts(CInt(Selectedrow.Cells("ContactID").Value))
                            Else
                                ugAssociatedCompanyContacts.DataSource = pConStruct.GetChildContacts(nContactID)
                            End If

                            If Me.ugAssociatedCompanyContacts.Rows.Count = 1 Then
                                Me.BtnUnassociateCompany.Enabled = True
                            Else
                                Me.BtnUnassociateCompany.Enabled = False
                            End If


                            Exit Sub
                        Else
                            strMode = "MODIFY"
                            pConStruct.Save(nModuleID, MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID, strMode, nEntityAssocID, pConStruct.ContactAssocID, nEntityID, nEntityTypeID, nModuleID, , )
                        End If
                    End If
                    If cmbPersonCompany.SelectedIndex = 1 And grpEntitySpecific.Enabled = True Then
                        strMode = "MODIFY"
                        pConStruct.Save(nModuleID, MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID, strMode, nEntityAssocID, pConStruct.ContactAssocID, nEntityID, nEntityTypeID, nModuleID, , )
                    End If
                End If
                If bolNewCompany = False Then
                    RaiseEvent ContactAdded()
                End If
                Me.Close()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            If ex.Message = "CONTACT IS ALREADY ASSOCIATED WITH THE CURRENT ENTITY" Then
                MsgBox(ex.Message)
                Me.Close()
            Else
                MyErr.ShowDialog()
            End If
        End Try

    End Sub
    Private Sub btnOk_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOk.Click
        SaveContact()
    End Sub

    Private Sub btnSearchforCompany_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSearchforCompany.Click
        Try
            'CompanyContactSearch
            objCompSearch = New CompanyContactSearch(nContactID, nEntityID, nEntityTypeID, strModuleName, pConStruct, Me)
            objCompSearch.ShowDialog()
            objCompSearch.BringToFront()

            If Not Selectedrow Is Nothing Then
                ugAssociatedCompanyContacts.DataSource = pConStruct.GetChildContacts(CInt(Selectedrow.Cells("ContactID").Value))
            Else
                ugAssociatedCompanyContacts.DataSource = pConStruct.GetChildContacts(nContactID)
            End If

            If Me.ugAssociatedCompanyContacts.Rows.Count = 1 Then
                Me.BtnUnassociateCompany.Enabled = True
            Else
                Me.BtnUnassociateCompany.Enabled = False
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.Close()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnClearCity_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearCity.Click
        CmbCity.Text = String.Empty
        'cmbCity.SelectedIndex = -1
    End Sub
    Private Sub btnClearState_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearState.Click
        cmbState.Text = String.Empty
        'cmbState.SelectedIndex = -1
    End Sub
    Private Sub btnClearZip_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearZip.Click
        CmbZip.Text = String.Empty
        'mskTxtZip.Mask = ""
        'mskTxtZip.Text = ""
        'mskTxtZip.SelText = ""
        'mskTxtZip.CtlText = String.Empty
        'mskTxtZip.Mask = "#####-####"
        'mskTxtZip.Text = cmbZip.Text
        'mskTxtZip.SelText = cmbZip.Text
        'cmbZip.SelectedIndex = -1
    End Sub
#End Region

#Region "Combo Events"
    Private Sub cmbPersonCompany_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbPersonCompany.SelectedIndexChanged
        Try
            If bolLoading Then Exit Sub

            If cmbPersonCompany.SelectedItem = "Person" Then
                pnlPerson.Visible = True
                pnlCompany.Visible = False
                txtCompanyName.Text = ""
            Else
                pnlCompany.Visible = True
                pnlCompany.Location = New System.Drawing.Point(40, 48)
                pnlPerson.Visible = False
                cmbTitle.Text = ""
                cmbSuffix.Text = ""
                txtFirstName.Text = ""
                txtMiddleName.Text = ""
                txtLastName.Text = ""
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    'Private Sub cmbCity_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Try
    '        If bolLoading Or cmbCity.SelectedIndex = -1 Then Exit Sub
    '        If strPrevCity = String.Empty Or strPrevCity <> cmbCity.Text Then
    '            pConStruct.PopulateAddressDetails(cmbCity.Text, cmbState.Text, "")
    '            strPrevCity = cmbCity.Text
    '        Else
    '            pConStruct.PopulateAddressDetails(cmbCity.Text, cmbState.Text, cmbZip.Text)
    '        End If


    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    Private Sub CmbCity_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles CmbCity.Leave
        Try
            If bolLoading Then Exit Sub
            If strPrevCity = String.Empty Or strPrevCity <> CmbCity.Text Then
                pConStruct.PopulateAddressDetails(CmbCity.Text, cmbState.Text, "")
                strPrevCity = CmbCity.Text
            Else
                pConStruct.PopulateAddressDetails(CmbCity.Text, cmbState.Text, CmbZip.Text)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub cmbState_leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbState.Leave
        Try
            If bolLoading Then Exit Sub
            If strPrevState = String.Empty Or strPrevState <> cmbState.Text Then
                UpdateComboBoxes(cmbState.Text, "", "")
                strPrevState = cmbState.Text
            Else
                UpdateComboBoxes(cmbState.Text, CmbCity.Text, CmbZip.Text)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub cmbZip_leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CmbZip.Leave
        Try
            If bolLoading Then Exit Sub
            'If cmbZip.Text <> String.Empty Then
            '    mskTxtZip.Mask = ""
            '    mskTxtZip.Text = ""
            '    mskTxtZip.SelText = ""
            '    mskTxtZip.CtlText = String.Empty
            '    mskTxtZip.Mask = "#####-####"
            '    mskTxtZip.SelText = cmbZip.Text
            'End If
            If strZip = String.Empty Or strZip <> CmbZip.Text Then
                UpdateComboBoxes("", "", CmbZip.Text)
                strZip = CmbZip.Text
            Else
                'pConStruct.PopulateAddressDetails(CmbCity.Text, cmbState.Text, CmbZip.Text)
                UpdateComboBoxes(cmbState.Text, CmbCity.Text, CmbZip.Text)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Public Sub UpdateComboBoxes(Optional ByVal strState As String = "", Optional ByVal strCity As String = "", Optional ByVal strZip As String = "")
        Try
            Dim strStates, strCities, strCounties, strZips, strFips As String
            Dim strOldZip As String
            Dim strNewState As String
            Dim ds As DataSet
            'If Not strState = String.Empty And Not strCity = String.Empty And Not strZip = String.Empty Then
            '    Exit Sub
            'End If
            'strNewState = IIf(strState <> String.Empty, strState, pComAdd.State.Trim)

            'strOldZip = pComAdd.Zip.Substring(0, pComAdd.Zip.Length - 5)
            '" AND FIPS LIKE '%" + pComAdd.FIPSCode.Trim + "%'" + _
            If strState = String.Empty Then
                strStates = "SELECT DISTINCT STATE FROM tblSYS_ZIPCODES WHERE" + _
                                " CITY LIKE '%" + strCity + "%'" + _
                                " AND ZIP LIKE '%" + strZip + "%'" + _
                                " ORDER BY STATE"
                Me.cmbState.DataSource = pConStruct.PopulateAddressLookUp(strStates)
                If cmbState.Rows.Count = 1 Then
                    cmbState.Text = cmbState.Rows.Item(0).Cells(0).Value

                End If
            End If

            '" AND FIPS LIKE '%" + pComAdd.FIPSCode.Trim + "%'" + _
            If cmbState.Text <> String.Empty Then
                strCities = "SELECT DISTINCT CITY FROM tblSYS_ZIPCODES WHERE" + _
                                " STATE LIKE '%" + cmbState.Text + "%'" + _
                                " AND ZIP LIKE '%" + strZip + "%'" + _
                                " ORDER BY CITY"
                CmbCity.DataSource = pConStruct.PopulateAddressLookUp(strCities)
                If CmbCity.Rows.Count = 1 Then
                    CmbCity.Text = CmbCity.Rows.Item(0).Cells(0).Value

                End If

                '" AND FIPS LIKE '%" + pComAdd.FIPSCode.Trim + "%'" + _
                strZips = "SELECT DISTINCT ZIP FROM tblSYS_ZIPCODES WHERE" + _
                                " STATE LIKE '%" + cmbState.Text + "%'" + _
                                " AND CITY LIKE '%" + strCity + "%'" + _
                                " ORDER BY ZIP"
                CmbZip.DataSource = pConStruct.PopulateAddressLookUp(strZips)
                If CmbZip.Rows.Count = 1 Then
                    CmbZip.Text = CmbZip.Rows.Item(0).Cells(0).Value
                ElseIf Not strZip = String.Empty Then
                    CmbZip.Text = strZip
                End If

            Else
                CmbCity.DataSource = Nothing
                CmbZip.DataSource = Nothing
                CmbCity.Text = String.Empty
                CmbZip.Text = String.Empty

            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region
#Region "Change Events"
    Private Sub pConStruct_CitiesChanged(ByVal dsCities As System.Data.DataSet) Handles pConStruct.CitiesChanged
        Try
            bolLoading = True
            CmbCity.DataSource = dsCities.Tables(0).DefaultView
            CmbCity.DisplayMember = "CITY"
            'cmbCity.SelectedIndex = -1
            CmbCity.Text = String.Empty
            bolLoading = False
        Catch ex As Exception
            ' Dim MyErr As New ErrorReport(ex)
            ' MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub pConStruct_StateChanged(ByVal dsState As System.Data.DataSet) Handles pConStruct.StateChanged
        Try
            bolLoading = True
            cmbState.DataSource = dsState.Tables(0).DefaultView
            cmbState.DisplayMember = "STATE"
            'cmbState.SelectedIndex = -1
            cmbState.Text = String.Empty
            bolLoading = False
        Catch ex As Exception
            'Dim MyErr As New ErrorReport(ex)
            'MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub pConStruct_ZipChanged(ByVal dsZip As System.Data.DataSet) Handles pConStruct.ZipChanged
        Try
            bolLoading = True
            CmbZip.DataSource = dsZip.Tables(0).DefaultView
            CmbZip.DisplayMember = "ZIP"
            'CmbZip.SelectedIndex = -1
            CmbZip.Text = String.Empty
            bolLoading = False
        Catch ex As Exception
            'Dim MyErr As New ErrorReport(ex)
            'MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub pConStruct_FipsChanged(ByVal strFIPS As String) Handles pConStruct.FipsChanged
        strFIPSCode = strFIPS
    End Sub

    '-------------------------------------------------------------------------------------------------------------
    Private Sub Company_CitiesChanged(ByVal dsCities As System.Data.DataSet) Handles pConStruct.CompanyCitiesChanged
        Try
            If bolNewCompany = False Then
                Exit Sub
            End If
            bolLoading = True
            CmbCity.DataSource = dsCities.Tables(0).DefaultView
            CmbCity.DisplayMember = "CITY"
            'cmbCity.SelectedIndex = -1
            CmbCity.Text = String.Empty
            bolLoading = False
        Catch ex As Exception
            'Dim MyErr As New ErrorReport(ex)
            'MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub Company_StateChanged(ByVal dsState As System.Data.DataSet) Handles pConStruct.CompanyStateChanged
        Try
            If bolNewCompany = False Then
                Exit Sub
            End If
            bolLoading = True
            cmbState.DataSource = dsState.Tables(0).DefaultView
            cmbState.DisplayMember = "STATE"
            'cmbState.SelectedIndex = -1
            cmbState.Text = String.Empty
            bolLoading = False
        Catch ex As Exception
            'Dim MyErr As New ErrorReport(ex)
            'MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub Company_ZipChanged(ByVal dsZip As System.Data.DataSet) Handles pConStruct.CompanyZipChanged
        Try
            If bolNewCompany = False Then
                Exit Sub
            End If
            bolLoading = True
            CmbZip.DataSource = dsZip.Tables(0).DefaultView
            CmbZip.DisplayMember = "ZIP"
            'CmbZip.SelectedIndex = -1
            CmbZip.Text = String.Empty
            bolLoading = False
        Catch ex As Exception
            'Dim MyErr As New ErrorReport(ex)
            'MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub Company_FipsChanged(ByVal strFIPS As String) Handles pConStruct.CompanyFipsChanged

        If bolNewCompany = False Then
            Exit Sub
        End If
        strFIPSCode = strFIPS

    End Sub
    Public Sub ValidationErrors(ByVal MsgStr As String) Handles pConStruct.evtContactStructErr
        If MsgStr <> String.Empty And MsgBox(MsgStr) = MsgBoxResult.OK Then
            MsgStr = String.Empty
        End If
    End Sub
#End Region

#Region "Control Events"
    Private Sub chkActive_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkActive.Click
        Try
            CheckActiveInactive(False)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub cmbCC_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCC.SelectedIndexChanged
        If cmbCC.SelectedItem = "YES" Then
            txtDisplayAs.Enabled = True
        Else
            txtDisplayAs.Enabled = False
            txtDisplayAs.Text = String.Empty
        End If
    End Sub


#End Region
#Region "Other Functions"
    Function ValidateData() As Boolean
        Try
            Dim errStr As String = ""
            Dim validateSuccess As Boolean = True

            If (Me.txtFirstName.Text = String.Empty Or Me.txtLastName.Text = String.Empty) And cmbPersonCompany.Text = "Person" Then
                errStr += "Enter FIRST NAME and LAST NAME (required)" + vbCrLf
                validateSuccess = False
            End If
            If Me.txtCompanyName.Text = String.Empty And cmbPersonCompany.Text <> "Person" Then
                errStr += "Enter COMPANY NAME(required)" + vbCrLf
                validateSuccess = False
            End If
            'If (Me.txtAddressOne.Text <> String.Empty Or mskTxtZip.FormattedText <> "_____-____" Or Me.CmbCity.Text <> String.Empty) Or Not cmbPersonCompany.Text = "Person" Then
            If (Me.txtAddressOne.Text <> String.Empty Or CmbZip.Text <> String.Empty Or Me.CmbCity.Text <> String.Empty) Or Not cmbPersonCompany.Text = "Person" Then
                If txtAddressOne.Text = String.Empty Then
                    errStr += "Enter Address One(required)" + vbCrLf
                    validateSuccess = False
                End If
                'If mskTxtZip.FormattedText = "_____-____" Then
                '    errStr += "Enter Zip code(required)" + vbCrLf
                '    validateSuccess = False
                'End If
                If Me.CmbZip.Text = String.Empty Then
                    errStr += "Enter Zip code(required)" + vbCrLf
                    validateSuccess = False
                End If
                If Me.CmbCity.Text = String.Empty Or Me.cmbState.Text = String.Empty Then
                    errStr += "Enter City and State(required)" + vbCrLf
                    validateSuccess = False
                End If
            End If
            If Me.mskTxtPhone1.FormattedText <> "(___)___-____" And Not UIUtilsGen.IsPhoneValid(mskTxtPhone1.FormattedText) Then
                errStr += "Invalid Phone1" + vbCrLf
                validateSuccess = False
            End If
            If Me.mskTxtPhone2.FormattedText <> "(___)___-____" And Not UIUtilsGen.IsPhoneValid(mskTxtPhone2.FormattedText) Then
                errStr += "Invalid Phone2" + vbCrLf
                validateSuccess = False
            End If

            If Me.mskTxtCell.FormattedText <> "(___)___-____" Then
                If Not UIUtilsGen.IsPhoneValid(mskTxtCell.FormattedText) Then
                    errStr += "Invalid Cell Number" + vbCrLf
                    validateSuccess = False
                End If
            End If
            If Me.mskTxtFax.FormattedText <> "(___)___-____" Then
                If Not UIUtilsGen.IsPhoneValid(mskTxtFax.FormattedText) Then
                    errStr += "Invalid Fax Number" + vbCrLf
                    validateSuccess = False
                End If
            End If
            If errStr.Length > 0 And Not validateSuccess Then
                MsgBox(errStr)
            End If
            Return validateSuccess
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    'Private Function FormatZIPCode() As String
    '    Dim strZip As String
    '    strZip = mskTxtZip.FormattedText.Replace("_", "")
    '    If strZip.EndsWith("-") Then
    '        strZip = strZip.Replace("-", "")
    '    End If
    '    Return strZip
    'End Function
    Private Sub AssociateNewCompany()

        Dim bolAddressDirty As Boolean = False
        Dim results As Long
        Try
            If txtAddressOne.Text = String.Empty And txtAddressTwo.Text = String.Empty And CmbCity.Text = String.Empty And CmbZip.Text = String.Empty Then 'FormatZIPCode() = String.Empty Then
                cmbState.Text = String.Empty
            End If
            'FormatZIPCode.Trim.ToString, _
            ConDatumInfo = New MUSTER.Info.ContactDatumInfo(nContactID, _
                                        IIf(cmbPersonCompany.Text = "Person", True, False), _
                                        0, _
                                        txtCompanyName.Text, _
                                        cmbTitle.Text, _
                                        String.Empty, _
                                        txtFirstName.Text, _
                                        txtMiddleName.Text, _
                                        txtLastName.Text, _
                                        cmbSuffix.Text, _
                                        txtAddressOne.Text, _
                                        txtAddressTwo.Text, _
                                        CmbCity.Text, _
                                        cmbState.Text, _
                                        CmbZip.Text, _
                                        strFIPSCode, _
                                        mskTxtPhone1.FormattedText.Trim.ToString, _
                                        mskTxtPhone2.FormattedText.Trim.ToString, _
                                        txtExt1.Text, _
                                        txtExt2.Text, _
                                        mskTxtFax.FormattedText.Trim.ToString, _
                                        mskTxtCell.FormattedText.Trim.ToString, _
                                        txtEmail.Text, _
                                        txtEmailPersonal.Text, _
                                        IIf(txtVendorValue.Text = "0", String.Empty, txtVendorValue.Text), _
                                        IIf(nContactID <= 0, MusterContainer.AppUser.ID, ""), _
                                        Now, _
                                        IIf(nContactID > 0, MusterContainer.AppUser.ID, ""), _
                                        CDate("01/01/0001"), _
                                        False)

            pConStruct.ContactDatum.Add(ConDatumInfo)
            Dim success As Boolean = False
            success = pConStruct.ContactDatum.Save(nModuleID, MusterContainer.AppUser.UserKey, returnVal, , , bolAddressDirty)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            If Not success Then
                Exit Sub
            End If
            pConStruct.childContactID = pConStruct.ContactDatum.ID
            pConStruct.Save(nModuleID, MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID, "SEARCH", IIf(lockMod, CInt(Selectedrow.Cells("EntityASSOCID").Value), Nothing), IIf(lockMod, pConStruct.ContactAssocID, Nothing))
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If


            '----- the contact is saved -------------------------------------------------
            MsgBox("New Company saved and associated with New Contact successfully!")
            '----------------------------------------------------------------------------
            RaiseEvent NewCompanyAdded()
            Me.Close()

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            If ex.Message = "CONTACT IS ALREADY ASSOCIATED WITH THE CURRENT ENTITY" Then
                MsgBox(ex.Message)
                Me.Close()
            Else
                MyErr.ShowDialog()
            End If
        End Try

    End Sub
    Private Sub AddressChangedToCompanyAdress()
        Dim bolAddressDirty As Boolean = False
        Try

            If UIUtilsGen.GetComboBoxValueInt(cmbType) = 0 And bolNewCompany = False Then
                MsgBox("Please enter the entity Type")
                Exit Sub
            End If

            If txtAddressOne.Text = String.Empty And txtAddressTwo.Text = String.Empty And CmbCity.Text = String.Empty And CmbZip.Text = String.Empty Then 'FormatZIPCode() = String.Empty Then
                cmbState.Text = String.Empty
            End If
            'Save Contacts.
            'ConDatumInfo = pConStruct.ContactDatum.ContactCollection.Item(nContactID)
            ConDatumInfo = pConStruct.ContactDatum.Retrieve(nContactID)
            pConStruct.contactStructInfo.parentContact.ID = nContactID
            pConStruct.contactStructInfo.parentContact.companyName = txtCompanyName.Text
            pConStruct.contactStructInfo.parentContact.Title = cmbTitle.Text
            pConStruct.contactStructInfo.parentContact.FirstName = txtFirstName.Text
            pConStruct.contactStructInfo.parentContact.MiddleName = txtMiddleName.Text
            pConStruct.contactStructInfo.parentContact.LastName = txtLastName.Text
            pConStruct.contactStructInfo.parentContact.suffix = cmbSuffix.Text
            pConStruct.contactStructInfo.parentContact.AddressLine1 = txtAddressOne.Text
            pConStruct.contactStructInfo.parentContact.AddressLine2 = txtAddressTwo.Text
            pConStruct.contactStructInfo.parentContact.City = CmbCity.Text
            pConStruct.contactStructInfo.parentContact.State = cmbState.Text
            'pConStruct.contactStructInfo.parentContact.ZipCode = FormatZIPCode.Trim.ToString
            pConStruct.contactStructInfo.parentContact.ZipCode = CmbZip.Text 'FormatZIPCode.Trim.ToString
            pConStruct.contactStructInfo.parentContact.FipsCode = strFIPSCode
            pConStruct.contactStructInfo.parentContact.publicEmail = txtEmail.Text
            pConStruct.contactStructInfo.parentContact.privateEmail = txtEmailPersonal.Text
            pConStruct.contactStructInfo.parentContact.Phone1 = mskTxtPhone1.FormattedText
            pConStruct.contactStructInfo.parentContact.Ext1 = txtExt1.Text
            pConStruct.contactStructInfo.parentContact.Phone2 = mskTxtPhone2.FormattedText
            pConStruct.contactStructInfo.parentContact.Ext2 = txtExt2.Text
            pConStruct.contactStructInfo.parentContact.modifiedBy = IIf(nContactID > 0, MusterContainer.AppUser.ID, "")
            pConStruct.contactStructInfo.parentContact.modifiedOn = Now()
            bolAddressDirty = pConStruct.contactStructInfo.parentContact.IsAddressDirty()
            If UCase(strModuleName) = UCase("Financial") Then
                pConStruct.contactStructInfo.parentContact.VendorNumber = IIf(txtVendorValue.Text = "0", String.Empty, txtVendorValue.Text)
            End If
            ConDatumInfo = pConStruct.contactStructInfo.parentContact

            pConStruct.ContactDatum.Add(ConDatumInfo)
            Dim success As Boolean = False
            success = pConStruct.ContactDatum.Save(nModuleID, MusterContainer.AppUser.UserKey, returnVal, , , bolAddressDirty)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            If Not success Then
                Exit Sub
            End If

            ' Save Contact- Contact Relationship & Entity- Contact Relationship()
            ConStructInfo = New MUSTER.Info.ContactStructInfo(pConStruct.ContactDatum.ID, _
                                                                pConStruct.childContactID, _
                                                                nEntityID, _
                                                                nEntityTypeID, _
                                                                UIUtilsGen.GetComboBoxValueInt(cmbType), _
                                                                nModuleID, _
                                                                chkActive.Checked, _
                                                                cmbCC.Text, _
                                                                txtDisplayAs.Text, _
                                                              False, Me.cmbAddress.SelectedValue, Me.cmbAlias.SelectedValue)

            'ConStructInfo.parentContact = pConStruct.ContactDatum.ContactCollection.Item(ConStructInfo.ParentContactID)
            ConStructInfo.parentContact = pConStruct.ContactDatum.Retrieve(ConStructInfo.ParentContactID)
            If ConStructInfo.ChildContactID <> 0 Then
                ConStructInfo.childContact = pConStruct.ContactDatum.Retrieve(ConStructInfo.ChildContactID)
            End If

            If pConStruct.ContactAssocID > 0 And strMode = "ASSOCIATE NEW" Then
                ConStructInfo.ContactAssocID = pConStruct.ContactAssocID()
            End If
            If pConStruct.contactStructInfo.entityAssocID > 0 Then
                ConStructInfo.entityAssocID = pConStruct.contactStructInfo.entityAssocID
            End If
            pConStruct.Add(ConStructInfo)
            If (pConStruct.childContactID > 0) Then

                If Not lockMod And strMode = "ASSOCIATE NEW" Then
                    pConStruct.Save(nModuleID, MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID, "ASSOCIATE", 0, pConStruct.ContactAssocID, pConStruct.entityID, pConStruct.entityType, pConStruct.ModuleID, , )
                End If
                If bolAddChangedFromCompany = True Then
                    pConStruct.Save(nModuleID, MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID, , 0, pConStruct.ContactAssocID, pConStruct.entityID, pConStruct.entityType, pConStruct.ModuleID, , )
                    bolAddChangedFromCompany = False
                End If

            End If

            dsContactAddresses = pConStruct.GetContactAliasAddresses(Me.nContactID)

            If Not dsContactAddresses Is Nothing AndAlso dsContactAddresses.Tables.Count > 0 Then
                ugAddresses.DataSource = dsContactAddresses.Tables(0).DefaultView
            End If



            '----- the contact is saved ----------------------
            MsgBox("Contact is saved successfully!")
            If bolNewCompany = False Then
                RaiseEvent ContactAdded()
            End If
            Me.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub SearchFrm_NewCompanyAdded() Handles objCompSearch.NewCompanyAdded


        cmbPersonCompany.Enabled = False

        ' pConStruct.PopulateAddressDetails(ConStructInfo.parentContact.City, ConStructInfo.parentContact.State, ConStructInfo.parentContact.ZipCode)
        'pConStruct.PopulateAddressDetails(pConStruct.contactStructInfo.parentContact.City, pConStruct.contactStructInfo.parentContact.State, pConStruct.contactStructInfo.parentContact.ZipCode)
        UpdateComboBoxes(pConStruct.contactStructInfo.parentContact.City, pConStruct.contactStructInfo.parentContact.State, pConStruct.contactStructInfo.parentContact.ZipCode)
        If lockMod Then
            strMode = "MODIFY"
        Else
            strMode = "ASSOCIATE NEW"

        End If

        cmbPersonCompany.Enabled = False

    End Sub

#End Region
#Region "Envelopes and Labels"
    Private Sub btnEnvelopes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEnvelopes.Click

        Try

            If pConStruct.contactStructInfo.parentContact.AddressLine1 <> String.Empty And pConStruct.contactStructInfo.parentContact.ID > 0 Then
                UIUtilsGen.CreateEnvelopes(strEnvelopeLabelName, arrAddress, "CON", pConStruct.contactStructInfo.parentContact.ID)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnLabels_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLabels.Click
        Try

            If pConStruct.contactStructInfo.parentContact.AddressLine1 <> String.Empty And pConStruct.contactStructInfo.parentContact.ID > 0 Then
                UIUtilsGen.CreateLabels(strEnvelopeLabelName, arrAddress, "CON", pConStruct.contactStructInfo.parentContact.ID)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

    Private Sub Contacts_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.GotFocus

    End Sub

    Private Sub Contacts_BindingContextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.BindingContextChanged

    End Sub

    Private Sub ugAddresses_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugAddresses.InitializeLayout

        ugAddresses.DisplayLayout.Bands(0).Columns("ContactID").Hidden = True
        ugAddresses.DisplayLayout.Bands(0).Columns("Address_ID").Hidden = True
        ugAddresses.DisplayLayout.Bands(0).Columns("Address").Width = ugAddresses.Width
        ugAddresses.DisplayLayout.Bands(0).Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False


    End Sub


    Private Sub ugAliases_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugAliases.InitializeLayout
        ugAliases.DisplayLayout.Bands(0).Columns("ContactID").Hidden = True
        ugAliases.DisplayLayout.Bands(0).Columns("ContactAliasID").Hidden = True
        ugAliases.DisplayLayout.Bands(0).Columns("Created_By").Hidden = True
        ugAliases.DisplayLayout.Bands(0).Columns("Date_Created").Hidden = True
        ugAliases.DisplayLayout.Bands(0).Columns("Alias").Width = ugAliases.Width

    End Sub

    Private Sub BtnDeleteAddress_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDeleteAddress.Click

        Try

            If Not ugAddresses.Selected.Rows Is Nothing AndAlso ugAddresses.Selected.Rows.Count > 0 Then
                With ugAddresses.Selected.Rows.Item(0)

                    dsContactAddresses = pConStruct.RemoveContactAddresses(.Cells("ContactID").Value, .Cells("Address_ID").Value)

                    If Not dsContactAddresses Is Nothing AndAlso dsContactAddresses.Tables.Count > 0 Then
                        ugAddresses.DataSource = dsContactAddresses.Tables(0).DefaultView
                    End If

                    populateAddresses()

                End With
            Else
                Throw New Exception("Please Select an Address Row first")
            End If


        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MsgBox(ex.Message)
        End Try



    End Sub

    Private Sub BtnSetAddress_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSetAtMain.Click

        Try

            If Not ugAddresses.Selected.Rows Is Nothing AndAlso ugAddresses.Selected.Rows.Count > 0 Then
                With ugAddresses.Selected.Rows.Item(0)

                    dsContactAddresses = pConStruct.SetContactAddressesToMain(.Cells("ContactID").Value, .Cells("Address_ID").Value)

                    MsgBox(String.Format("The Address '{0}' has been set as the main address for this contact", .Cells("Address").Value), MsgBoxStyle.OKOnly, "Contact Address Management")

                    If Not dsContactAddresses Is Nothing AndAlso dsContactAddresses.Tables.Count > 0 Then
                        ugAddresses.DataSource = dsContactAddresses.Tables(0).DefaultView
                    End If

                    Contacts_Load(Me, Nothing)

                    RaiseEvent ContactAdded()


                End With
            Else
                Throw New Exception("Please Select a Contact Address from the rows first")
            End If


        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MsgBox(ex.Message)
        End Try



    End Sub


    Private Sub BtnDeleteAlias_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnDeleteAlias.Click

        Try

            If Not ugAliases.Selected.Rows Is Nothing AndAlso ugAliases.Selected.Rows.Count > 0 Then
                With ugAliases.Selected.Rows.Item(0)

                    dsContactAsocAliases = pConStruct.RemoveContactAlias(.Cells("ContactID").Value, .Cells("ContactAliasID").Value)

                    If Not dsContactAsocAliases Is Nothing AndAlso dsContactAsocAliases.Tables.Count > 0 Then
                        ugAliases.DataSource = dsContactAsocAliases.Tables(0).DefaultView
                    End If

                End With

                populateAliases()
            Else
                Throw New Exception("Please Select an Address Row first")
            End If




        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MsgBox(ex.Message)
        End Try



    End Sub





    Private Sub BtnAddAlias_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAddAlias.Click

        Dim msg As String = String.Empty

        Try

            msg = InputBox("Please enter new Alias name", "Contact Alias")

            While msg = String.Empty
                msg = InputBox("Please enter new Alias name", "Contact Alias - Blank values are not allowed")
            End While

            dsContactAsocAliases = pConStruct.AddContactAlias(msg, nContactID, MusterContainer.AppUser.ID)

            If Not dsContactAsocAliases Is Nothing AndAlso dsContactAsocAliases.Tables.Count > 0 Then
                ugAliases.DataSource = dsContactAsocAliases.Tables(0).DefaultView
            End If

            populateAliases()

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MsgBox(ex.Message)

        End Try





    End Sub

    Private Sub ugAssociatedCompanyContacts_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugAssociatedCompanyContacts.InitializeLayout


        With ugAssociatedCompanyContacts.DisplayLayout.Bands(0).Columns
            .Item("Address_One").Hidden = False
            .Item("Address_Two").Hidden = True
            .Item("City").Hidden = False
            .Item("State").Hidden = True
            .Item("ZipCode").Hidden = True
            .Item("ext_one").Hidden = True
            .Item("Phone_Number_two").Hidden = True
            .Item("ext_two").Hidden = True
            .Item("cell_number").Hidden = True
            .Item("fax_Number").Hidden = True
            .Item("Email_Address").Hidden = True
            .Item("Email_Address_Personal").Hidden = True
            .Item("Created_By").Hidden = True
            .Item("Date_Created").Hidden = True
            .Item("Last_Edited_By").Hidden = True
            .Item("Date_Last_Edited").Hidden = True
            .Item("ContactAssocID").Hidden = True
            .Item("Child_Contact").Hidden = True

            .Item("Contact_Name").Width = ugAssociatedCompanyContacts.Width / 2
        End With

    End Sub




    Private Sub BtnUnassociateCompany_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BtnUnassociateCompany.Click
        pConStruct.childContactID = 0
        nContactAssocID = 0
        ugAssociatedCompanyContacts.DataSource = Nothing
    End Sub
End Class
