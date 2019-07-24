Public Class Company
    Inherits System.Windows.Forms.Form

#Region "User Defined Variables"
    Public MyGuid As New System.Guid
    'Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long

    ' Company
    Private WithEvents pCompany As MUSTER.BusinessLogic.pCompany
    Private WithEvents oComAdd As New MUSTER.BusinessLogic.pComAddress
    'Private WithEvents AddressForm As Address
    Friend nCompanyID As Integer = 0
    Dim nCompanyAddressID As Integer = 0
    Dim dateFinRespExp, dateEngLiabExpiration, dateEngAppApproval As Date
    Friend WithEvents objAddMaster As AddressMaster

    ' Licensee
    Private WithEvents pLic As New MUSTER.BusinessLogic.pLicensee
    Private WithEvents pMgr As New MUSTER.BusinessLogic.pLicensee
    Private WithEvents pCompanyLicenseeAssociation As New MUSTER.BusinessLogic.pCompanyLicensee
    '  Private WithEvents pCompanyManagerAssociation As New MUSTER.BusinessLogic.pCompanyLicensee
    Friend WithEvents objLicen As Licensees
    Friend WithEvents objMgr As Managers
    Friend WithEvents oLicenseeList As LicenseesList
    Friend WithEvents oManagerList As ManagersList
    Dim nLicenseeID As Integer = 0
    Dim nManagerID As Integer = 0
    Dim nCompanyLicenseeAssocID As Integer = 0
    Dim companyInfo As MUSTER.Info.CompanyInfo
    Dim CompanyLicenseeInfo As MUSTER.Info.CompanyLicenseeInfo

    Private bolValidateSuccess As Boolean = True
    Private bolDisplayErrmessage As Boolean = True
    Private WithEvents SF As ShowFlags
    Dim bolLoading As Boolean = False
    Dim returnVal As String = String.Empty

    '----- variables for the contact management section --------
    Private pConStruct As New MUSTER.BusinessLogic.pContactStruct
    Private WithEvents objCntSearch As ContactSearch
    Dim dsContacts As DataSet
    Dim result As DialogResult
    '-----------------------------------------------------------

#End Region

#Region " Windows Form Designer generated code "

    Public Sub New(Optional ByVal CompanyID As Integer = 0)
        MyBase.New()
        MyGuid = System.Guid.NewGuid
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        'Add any initialization after the InitializeComponent() call
        'Need to tell the AppUser that we've instantiated another Registration form...
        '
        MusterContainer.AppUser.LogEntry("Company", MyGuid.ToString)
        '
        ' The following line enables all forms to detect the visible form in the MDI container
        '
        MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "Company")
        nCompanyID = CompanyID

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
    Friend WithEvents pnlCompanyBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlCompanyContainer As System.Windows.Forms.Panel
    Friend WithEvents btnFlags As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnDeleteCompany As System.Windows.Forms.Button
    Friend WithEvents btnComments As System.Windows.Forms.Button
    Friend WithEvents btnSaveCompany As System.Windows.Forms.Button
    Friend WithEvents lblFinancialRespExpirationDate As System.Windows.Forms.Label
    Friend WithEvents lblCertORResponsibility As System.Windows.Forms.Label
    Friend WithEvents lblCompanyID As System.Windows.Forms.Label
    Friend WithEvents txtCompanyID As System.Windows.Forms.TextBox
    Friend WithEvents txtCompany As System.Windows.Forms.TextBox
    Friend WithEvents txtEmail As System.Windows.Forms.TextBox
    Friend WithEvents lblCompany As System.Windows.Forms.Label
    Friend WithEvents lblEmail As System.Windows.Forms.Label
    Friend WithEvents cmbCertORResponsibility As System.Windows.Forms.ComboBox
    Friend WithEvents grpERACs As System.Windows.Forms.GroupBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents lblLicensees As System.Windows.Forms.Label
    Friend WithEvents lblManagers As System.Windows.Forms.Label
    Friend WithEvents ugLicensees As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents ugManagers As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents grpTypes As System.Windows.Forms.GroupBox
    Friend WithEvents lblPriorLicensees As System.Windows.Forms.Label
    Friend WithEvents ugPriorLicensees As Infragistics.Win.UltraWinGrid.UltraGrid
    Public WithEvents dtFinancialRespExpirationDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtProfGeologist As System.Windows.Forms.TextBox
    Friend WithEvents txtProfEngineer As System.Windows.Forms.TextBox
    Friend WithEvents lblProfGeologist As System.Windows.Forms.Label
    Friend WithEvents lblProfEngineer As System.Windows.Forms.Label
    Friend WithEvents chkCorrosionExpert As System.Windows.Forms.CheckBox
    Friend WithEvents chkTankLining As System.Windows.Forms.CheckBox
    Friend WithEvents chkEnvironmentalDrillers As System.Windows.Forms.CheckBox
    Friend WithEvents chkEnvironmentalConsultants As System.Windows.Forms.CheckBox
    Friend WithEvents chkIRAC As System.Windows.Forms.CheckBox
    Friend WithEvents chkThermalSoilTreatment As System.Windows.Forms.CheckBox
    Friend WithEvents chkWastewaterLaboratories As System.Windows.Forms.CheckBox
    Friend WithEvents chkUSTSuppliesAndEquipment As System.Windows.Forms.CheckBox
    Friend WithEvents chkLeakDetectionCompanies As System.Windows.Forms.CheckBox
    Friend WithEvents chkPrecisionTankTightnessTesters As System.Windows.Forms.CheckBox
    Friend WithEvents chkCertifiedToCloseUSTs As System.Windows.Forms.CheckBox
    Friend WithEvents chkCertifiedtoInstallAlterandCloseUSTs As System.Windows.Forms.CheckBox
    Friend WithEvents btnDisassociate As System.Windows.Forms.Button
    Friend WithEvents btnLicenseesEdit As System.Windows.Forms.Button
    Friend WithEvents btnLicenseesAddNew As System.Windows.Forms.Button
    Friend WithEvents btnLicenseesAddExisting As System.Windows.Forms.Button
    Friend WithEvents btnManagersEdit As System.Windows.Forms.Button
    Friend WithEvents btnManagersAddNew As System.Windows.Forms.Button
    Friend WithEvents btnManagersAddExisting As System.Windows.Forms.Button
    Friend WithEvents btnManagerDisassociate As System.Windows.Forms.Button
    Friend WithEvents txtGeologistNo As System.Windows.Forms.TextBox
    Friend WithEvents txtEngineerNo As System.Windows.Forms.TextBox
    Public WithEvents dtEngineerLiabilityExpiration As System.Windows.Forms.DateTimePicker
    Public WithEvents dtEngineerAppApprovalDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblEngineerAppApprovalDate As System.Windows.Forms.Label
    Friend WithEvents lblEngineerLiabilityExpiration As System.Windows.Forms.Label
    Friend WithEvents pnlCompanyHeader As System.Windows.Forms.Panel
    Friend WithEvents lblCompanyHeader As System.Windows.Forms.Label
    Friend WithEvents lblCompanyDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlCompanyDetails As System.Windows.Forms.Panel
    Friend WithEvents pnlLicenseeHead As System.Windows.Forms.Panel
    Friend WithEvents pnlManagerHead As System.Windows.Forms.Panel
    Friend WithEvents lblLicenseeHead As System.Windows.Forms.Label
    Friend WithEvents lblManagerHead As System.Windows.Forms.Label
    Friend WithEvents lblLicenseeDisplay As System.Windows.Forms.Label
    Friend WithEvents lblManagerDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlEracHead As System.Windows.Forms.Panel
    Friend WithEvents lnlEracHead As System.Windows.Forms.Label
    Friend WithEvents lblEracDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlEracDetails As System.Windows.Forms.Panel
    Friend WithEvents pnlTypesHead As System.Windows.Forms.Panel
    Friend WithEvents lblTypesHead As System.Windows.Forms.Label
    Friend WithEvents lblTypesDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlTypesDetails As System.Windows.Forms.Panel
    Friend WithEvents pnlPriorCompaniesHead As System.Windows.Forms.Panel
    Friend WithEvents lblPriorLicenseesHead As System.Windows.Forms.Label
    Friend WithEvents lblPriorLicenseesDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlPriorLicenseesDetails As System.Windows.Forms.Panel
    Friend WithEvents pnlLicenseesDetails As System.Windows.Forms.Panel
    Friend WithEvents pnlManagersDetails As System.Windows.Forms.Panel
    Friend WithEvents lblDocumentsHead As System.Windows.Forms.Label
    Friend WithEvents lblDocumentsDisplay As System.Windows.Forms.Label
    Friend WithEvents pnlDocumentDetails As System.Windows.Forms.Panel
    Friend WithEvents UCCompanyDocuments As MUSTER.DocumentViewControl
    Friend WithEvents pnlDocumentsHead As System.Windows.Forms.Panel
    Friend WithEvents lblGeologistEmail As System.Windows.Forms.Label
    Friend WithEvents lblEngineerEmail As System.Windows.Forms.Label
    Friend WithEvents TxtEngineerEmail As System.Windows.Forms.TextBox
    Friend WithEvents TxtGeologistEmail As System.Windows.Forms.TextBox
    Friend WithEvents btnEnvelopes As System.Windows.Forms.Button
    Friend WithEvents btnLabels As System.Windows.Forms.Button
    Friend WithEvents lblAddresses As System.Windows.Forms.Label
    Public WithEvents txtCompanyAddress As System.Windows.Forms.TextBox
    Friend WithEvents pnlERACandContacts As System.Windows.Forms.Panel
    Friend WithEvents lblCompanyContactsHead As System.Windows.Forms.Label
    Friend WithEvents lblCompanyContactsDisplay As System.Windows.Forms.Label
    Friend WithEvents lblCompanyContacts As System.Windows.Forms.Label
    Friend WithEvents pnlCompanyContactHeader As System.Windows.Forms.Panel
    Friend WithEvents pnlCompanyContactContainer As System.Windows.Forms.Panel
    Friend WithEvents ugCompanyContacts As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents chkCompanyShowActive As System.Windows.Forms.CheckBox
    Friend WithEvents chkCompanyShowRelatedContacts As System.Windows.Forms.CheckBox
    Friend WithEvents chkCompanyShowContactsforAllModules As System.Windows.Forms.CheckBox
    Friend WithEvents pnlCompanyContactButtons As System.Windows.Forms.Panel
    Friend WithEvents btnCompanyContactModify As System.Windows.Forms.Button
    Friend WithEvents btnCompanyContactDelete As System.Windows.Forms.Button
    Friend WithEvents btnCompanyContactAssociate As System.Windows.Forms.Button
    Friend WithEvents btnCompanyContactAddorSearch As System.Windows.Forms.Button
    Friend WithEvents chkERAC As System.Windows.Forms.CheckBox
    Friend WithEvents chkManager As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlCompanyBottom = New System.Windows.Forms.Panel
        Me.btnFlags = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnDeleteCompany = New System.Windows.Forms.Button
        Me.btnComments = New System.Windows.Forms.Button
        Me.btnSaveCompany = New System.Windows.Forms.Button
        Me.pnlCompanyContainer = New System.Windows.Forms.Panel
        Me.pnlCompanyContactButtons = New System.Windows.Forms.Panel
        Me.btnCompanyContactModify = New System.Windows.Forms.Button
        Me.btnCompanyContactDelete = New System.Windows.Forms.Button
        Me.btnCompanyContactAssociate = New System.Windows.Forms.Button
        Me.btnCompanyContactAddorSearch = New System.Windows.Forms.Button
        Me.pnlCompanyContactContainer = New System.Windows.Forms.Panel
        Me.ugCompanyContacts = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.Label4 = New System.Windows.Forms.Label
        Me.pnlCompanyContactHeader = New System.Windows.Forms.Panel
        Me.chkCompanyShowActive = New System.Windows.Forms.CheckBox
        Me.chkCompanyShowRelatedContacts = New System.Windows.Forms.CheckBox
        Me.chkCompanyShowContactsforAllModules = New System.Windows.Forms.CheckBox
        Me.lblCompanyContacts = New System.Windows.Forms.Label
        Me.pnlERACandContacts = New System.Windows.Forms.Panel
        Me.lblCompanyContactsHead = New System.Windows.Forms.Label
        Me.lblCompanyContactsDisplay = New System.Windows.Forms.Label
        Me.pnlDocumentDetails = New System.Windows.Forms.Panel
        Me.UCCompanyDocuments = New MUSTER.DocumentViewControl
        Me.pnlDocumentsHead = New System.Windows.Forms.Panel
        Me.lblDocumentsHead = New System.Windows.Forms.Label
        Me.lblDocumentsDisplay = New System.Windows.Forms.Label
        Me.pnlPriorLicenseesDetails = New System.Windows.Forms.Panel
        Me.lblPriorLicensees = New System.Windows.Forms.Label
        Me.ugPriorLicensees = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlPriorCompaniesHead = New System.Windows.Forms.Panel
        Me.lblPriorLicenseesHead = New System.Windows.Forms.Label
        Me.lblPriorLicenseesDisplay = New System.Windows.Forms.Label
        Me.pnlTypesDetails = New System.Windows.Forms.Panel
        Me.grpTypes = New System.Windows.Forms.GroupBox
        Me.chkManager = New System.Windows.Forms.CheckBox
        Me.chkCorrosionExpert = New System.Windows.Forms.CheckBox
        Me.chkTankLining = New System.Windows.Forms.CheckBox
        Me.chkEnvironmentalDrillers = New System.Windows.Forms.CheckBox
        Me.chkEnvironmentalConsultants = New System.Windows.Forms.CheckBox
        Me.chkIRAC = New System.Windows.Forms.CheckBox
        Me.chkERAC = New System.Windows.Forms.CheckBox
        Me.chkThermalSoilTreatment = New System.Windows.Forms.CheckBox
        Me.chkWastewaterLaboratories = New System.Windows.Forms.CheckBox
        Me.chkUSTSuppliesAndEquipment = New System.Windows.Forms.CheckBox
        Me.chkLeakDetectionCompanies = New System.Windows.Forms.CheckBox
        Me.chkPrecisionTankTightnessTesters = New System.Windows.Forms.CheckBox
        Me.chkCertifiedToCloseUSTs = New System.Windows.Forms.CheckBox
        Me.chkCertifiedtoInstallAlterandCloseUSTs = New System.Windows.Forms.CheckBox
        Me.pnlTypesHead = New System.Windows.Forms.Panel
        Me.lblTypesHead = New System.Windows.Forms.Label
        Me.lblTypesDisplay = New System.Windows.Forms.Label
        Me.pnlEracDetails = New System.Windows.Forms.Panel
        Me.grpERACs = New System.Windows.Forms.GroupBox
        Me.TxtGeologistEmail = New System.Windows.Forms.TextBox
        Me.TxtEngineerEmail = New System.Windows.Forms.TextBox
        Me.dtEngineerLiabilityExpiration = New System.Windows.Forms.DateTimePicker
        Me.dtEngineerAppApprovalDate = New System.Windows.Forms.DateTimePicker
        Me.lblEngineerAppApprovalDate = New System.Windows.Forms.Label
        Me.lblEngineerLiabilityExpiration = New System.Windows.Forms.Label
        Me.txtGeologistNo = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.lblGeologistEmail = New System.Windows.Forms.Label
        Me.txtEngineerNo = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.lblEngineerEmail = New System.Windows.Forms.Label
        Me.txtProfGeologist = New System.Windows.Forms.TextBox
        Me.txtProfEngineer = New System.Windows.Forms.TextBox
        Me.lblProfGeologist = New System.Windows.Forms.Label
        Me.lblProfEngineer = New System.Windows.Forms.Label
        Me.pnlEracHead = New System.Windows.Forms.Panel
        Me.lnlEracHead = New System.Windows.Forms.Label
        Me.lblEracDisplay = New System.Windows.Forms.Label
        Me.pnlLicenseesDetails = New System.Windows.Forms.Panel
        Me.lblLicensees = New System.Windows.Forms.Label
        Me.btnLicenseesAddExisting = New System.Windows.Forms.Button
        Me.btnLicenseesAddNew = New System.Windows.Forms.Button
        Me.btnLicenseesEdit = New System.Windows.Forms.Button
        Me.btnDisassociate = New System.Windows.Forms.Button
        Me.ugLicensees = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlManagersDetails = New System.Windows.Forms.Panel
        Me.lblManagers = New System.Windows.Forms.Label
        Me.btnManagersAddExisting = New System.Windows.Forms.Button
        Me.btnManagersAddNew = New System.Windows.Forms.Button
        Me.btnManagersEdit = New System.Windows.Forms.Button
        Me.btnManagerDisassociate = New System.Windows.Forms.Button
        Me.ugManagers = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlLicenseeHead = New System.Windows.Forms.Panel
        Me.lblLicenseeHead = New System.Windows.Forms.Label
        Me.lblLicenseeDisplay = New System.Windows.Forms.Label
        Me.pnlManagerHead = New System.Windows.Forms.Panel
        Me.lblManagerHead = New System.Windows.Forms.Label
        Me.lblManagerDisplay = New System.Windows.Forms.Label
        Me.pnlCompanyDetails = New System.Windows.Forms.Panel
        Me.txtCompanyAddress = New System.Windows.Forms.TextBox
        Me.lblAddresses = New System.Windows.Forms.Label
        Me.btnLabels = New System.Windows.Forms.Button
        Me.btnEnvelopes = New System.Windows.Forms.Button
        Me.cmbCertORResponsibility = New System.Windows.Forms.ComboBox
        Me.lblEmail = New System.Windows.Forms.Label
        Me.lblCompany = New System.Windows.Forms.Label
        Me.txtEmail = New System.Windows.Forms.TextBox
        Me.txtCompany = New System.Windows.Forms.TextBox
        Me.txtCompanyID = New System.Windows.Forms.TextBox
        Me.lblFinancialRespExpirationDate = New System.Windows.Forms.Label
        Me.lblCertORResponsibility = New System.Windows.Forms.Label
        Me.lblCompanyID = New System.Windows.Forms.Label
        Me.dtFinancialRespExpirationDate = New System.Windows.Forms.DateTimePicker
        Me.pnlCompanyHeader = New System.Windows.Forms.Panel
        Me.lblCompanyHeader = New System.Windows.Forms.Label
        Me.lblCompanyDisplay = New System.Windows.Forms.Label
        Me.pnlCompanyBottom.SuspendLayout()
        Me.pnlCompanyContainer.SuspendLayout()
        Me.pnlCompanyContactButtons.SuspendLayout()
        Me.pnlCompanyContactContainer.SuspendLayout()
        CType(Me.ugCompanyContacts, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlCompanyContactHeader.SuspendLayout()
        Me.pnlERACandContacts.SuspendLayout()
        Me.pnlDocumentDetails.SuspendLayout()
        Me.pnlDocumentsHead.SuspendLayout()
        Me.pnlPriorLicenseesDetails.SuspendLayout()
        CType(Me.ugPriorLicensees, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlPriorCompaniesHead.SuspendLayout()
        Me.pnlTypesDetails.SuspendLayout()
        Me.grpTypes.SuspendLayout()
        Me.pnlTypesHead.SuspendLayout()
        Me.pnlEracDetails.SuspendLayout()
        Me.grpERACs.SuspendLayout()
        Me.pnlEracHead.SuspendLayout()
        Me.pnlLicenseesDetails.SuspendLayout()
        CType(Me.ugLicensees, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlManagersDetails.SuspendLayout()
        CType(Me.ugManagers, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlLicenseeHead.SuspendLayout()
        Me.pnlManagerHead.SuspendLayout()
        Me.pnlCompanyDetails.SuspendLayout()
        Me.pnlCompanyHeader.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlCompanyBottom
        '
        Me.pnlCompanyBottom.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlCompanyBottom.Controls.Add(Me.btnFlags)
        Me.pnlCompanyBottom.Controls.Add(Me.btnCancel)
        Me.pnlCompanyBottom.Controls.Add(Me.btnDeleteCompany)
        Me.pnlCompanyBottom.Controls.Add(Me.btnComments)
        Me.pnlCompanyBottom.Controls.Add(Me.btnSaveCompany)
        Me.pnlCompanyBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlCompanyBottom.Location = New System.Drawing.Point(0, 713)
        Me.pnlCompanyBottom.Name = "pnlCompanyBottom"
        Me.pnlCompanyBottom.Size = New System.Drawing.Size(992, 40)
        Me.pnlCompanyBottom.TabIndex = 1
        '
        'btnFlags
        '
        Me.btnFlags.Enabled = False
        Me.btnFlags.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFlags.Location = New System.Drawing.Point(528, 8)
        Me.btnFlags.Name = "btnFlags"
        Me.btnFlags.Size = New System.Drawing.Size(104, 26)
        Me.btnFlags.TabIndex = 13
        Me.btnFlags.Text = "Flags"
        '
        'btnCancel
        '
        Me.btnCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.Location = New System.Drawing.Point(200, 8)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(88, 26)
        Me.btnCancel.TabIndex = 10
        Me.btnCancel.Text = "Cancel"
        '
        'btnDeleteCompany
        '
        Me.btnDeleteCompany.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDeleteCompany.Location = New System.Drawing.Point(296, 8)
        Me.btnDeleteCompany.Name = "btnDeleteCompany"
        Me.btnDeleteCompany.Size = New System.Drawing.Size(112, 26)
        Me.btnDeleteCompany.TabIndex = 11
        Me.btnDeleteCompany.Text = "Delete Company"
        '
        'btnComments
        '
        Me.btnComments.Enabled = False
        Me.btnComments.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnComments.Location = New System.Drawing.Point(416, 8)
        Me.btnComments.Name = "btnComments"
        Me.btnComments.Size = New System.Drawing.Size(104, 26)
        Me.btnComments.TabIndex = 12
        Me.btnComments.Text = "Comments"
        '
        'btnSaveCompany
        '
        Me.btnSaveCompany.Enabled = False
        Me.btnSaveCompany.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSaveCompany.Location = New System.Drawing.Point(72, 8)
        Me.btnSaveCompany.Name = "btnSaveCompany"
        Me.btnSaveCompany.Size = New System.Drawing.Size(120, 26)
        Me.btnSaveCompany.TabIndex = 9
        Me.btnSaveCompany.Text = "Save Company"
        '
        'pnlCompanyContainer
        '
        Me.pnlCompanyContainer.AutoScroll = True
        Me.pnlCompanyContainer.Controls.Add(Me.pnlCompanyContactButtons)
        Me.pnlCompanyContainer.Controls.Add(Me.pnlCompanyContactContainer)
        Me.pnlCompanyContainer.Controls.Add(Me.pnlCompanyContactHeader)
        Me.pnlCompanyContainer.Controls.Add(Me.pnlERACandContacts)
        Me.pnlCompanyContainer.Controls.Add(Me.pnlDocumentDetails)
        Me.pnlCompanyContainer.Controls.Add(Me.pnlDocumentsHead)
        Me.pnlCompanyContainer.Controls.Add(Me.pnlPriorLicenseesDetails)
        Me.pnlCompanyContainer.Controls.Add(Me.pnlPriorCompaniesHead)
        Me.pnlCompanyContainer.Controls.Add(Me.pnlTypesDetails)
        Me.pnlCompanyContainer.Controls.Add(Me.pnlTypesHead)
        Me.pnlCompanyContainer.Controls.Add(Me.pnlEracDetails)
        Me.pnlCompanyContainer.Controls.Add(Me.pnlEracHead)
        Me.pnlCompanyContainer.Controls.Add(Me.pnlLicenseesDetails)
        Me.pnlCompanyContainer.Controls.Add(Me.pnlManagersDetails)
        Me.pnlCompanyContainer.Controls.Add(Me.pnlLicenseeHead)
        Me.pnlCompanyContainer.Controls.Add(Me.pnlManagerHead)
        Me.pnlCompanyContainer.Controls.Add(Me.pnlCompanyDetails)
        Me.pnlCompanyContainer.Controls.Add(Me.pnlCompanyHeader)
        Me.pnlCompanyContainer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlCompanyContainer.Location = New System.Drawing.Point(0, 0)
        Me.pnlCompanyContainer.Name = "pnlCompanyContainer"
        Me.pnlCompanyContainer.Size = New System.Drawing.Size(992, 713)
        Me.pnlCompanyContainer.TabIndex = 0
        '
        'pnlCompanyContactButtons
        '
        Me.pnlCompanyContactButtons.Controls.Add(Me.btnCompanyContactModify)
        Me.pnlCompanyContactButtons.Controls.Add(Me.btnCompanyContactDelete)
        Me.pnlCompanyContactButtons.Controls.Add(Me.btnCompanyContactAssociate)
        Me.pnlCompanyContactButtons.Controls.Add(Me.btnCompanyContactAddorSearch)
        Me.pnlCompanyContactButtons.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlCompanyContactButtons.DockPadding.All = 3
        Me.pnlCompanyContactButtons.Location = New System.Drawing.Point(0, 1490)
        Me.pnlCompanyContactButtons.Name = "pnlCompanyContactButtons"
        Me.pnlCompanyContactButtons.Size = New System.Drawing.Size(976, 40)
        Me.pnlCompanyContactButtons.TabIndex = 225
        '
        'btnCompanyContactModify
        '
        Me.btnCompanyContactModify.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCompanyContactModify.Location = New System.Drawing.Point(240, 8)
        Me.btnCompanyContactModify.Name = "btnCompanyContactModify"
        Me.btnCompanyContactModify.Size = New System.Drawing.Size(235, 26)
        Me.btnCompanyContactModify.TabIndex = 1
        Me.btnCompanyContactModify.Text = "Modify Contact"
        '
        'btnCompanyContactDelete
        '
        Me.btnCompanyContactDelete.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCompanyContactDelete.Location = New System.Drawing.Point(472, 8)
        Me.btnCompanyContactDelete.Name = "btnCompanyContactDelete"
        Me.btnCompanyContactDelete.Size = New System.Drawing.Size(235, 26)
        Me.btnCompanyContactDelete.TabIndex = 2
        Me.btnCompanyContactDelete.Text = "Disassociate Contact"
        '
        'btnCompanyContactAssociate
        '
        Me.btnCompanyContactAssociate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCompanyContactAssociate.Location = New System.Drawing.Point(704, 8)
        Me.btnCompanyContactAssociate.Name = "btnCompanyContactAssociate"
        Me.btnCompanyContactAssociate.Size = New System.Drawing.Size(235, 26)
        Me.btnCompanyContactAssociate.TabIndex = 3
        Me.btnCompanyContactAssociate.Text = "Associate Contact from Different Module"
        '
        'btnCompanyContactAddorSearch
        '
        Me.btnCompanyContactAddorSearch.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCompanyContactAddorSearch.Location = New System.Drawing.Point(8, 8)
        Me.btnCompanyContactAddorSearch.Name = "btnCompanyContactAddorSearch"
        Me.btnCompanyContactAddorSearch.Size = New System.Drawing.Size(235, 26)
        Me.btnCompanyContactAddorSearch.TabIndex = 0
        Me.btnCompanyContactAddorSearch.Text = "Add/Search Contact To Associate"
        '
        'pnlCompanyContactContainer
        '
        Me.pnlCompanyContactContainer.AutoScroll = True
        Me.pnlCompanyContactContainer.Controls.Add(Me.ugCompanyContacts)
        Me.pnlCompanyContactContainer.Controls.Add(Me.Label4)
        Me.pnlCompanyContactContainer.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlCompanyContactContainer.Location = New System.Drawing.Point(0, 1386)
        Me.pnlCompanyContactContainer.Name = "pnlCompanyContactContainer"
        Me.pnlCompanyContactContainer.Size = New System.Drawing.Size(976, 104)
        Me.pnlCompanyContactContainer.TabIndex = 224
        '
        'ugCompanyContacts
        '
        Me.ugCompanyContacts.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugCompanyContacts.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugCompanyContacts.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        Me.ugCompanyContacts.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugCompanyContacts.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugCompanyContacts.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugCompanyContacts.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugCompanyContacts.Location = New System.Drawing.Point(0, 0)
        Me.ugCompanyContacts.Name = "ugCompanyContacts"
        Me.ugCompanyContacts.Size = New System.Drawing.Size(976, 104)
        Me.ugCompanyContacts.TabIndex = 0
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(792, 56)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(7, 23)
        Me.Label4.TabIndex = 2
        '
        'pnlCompanyContactHeader
        '
        Me.pnlCompanyContactHeader.Controls.Add(Me.chkCompanyShowActive)
        Me.pnlCompanyContactHeader.Controls.Add(Me.chkCompanyShowRelatedContacts)
        Me.pnlCompanyContactHeader.Controls.Add(Me.chkCompanyShowContactsforAllModules)
        Me.pnlCompanyContactHeader.Controls.Add(Me.lblCompanyContacts)
        Me.pnlCompanyContactHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlCompanyContactHeader.DockPadding.All = 3
        Me.pnlCompanyContactHeader.Location = New System.Drawing.Point(0, 1354)
        Me.pnlCompanyContactHeader.Name = "pnlCompanyContactHeader"
        Me.pnlCompanyContactHeader.Size = New System.Drawing.Size(976, 32)
        Me.pnlCompanyContactHeader.TabIndex = 223
        '
        'chkCompanyShowActive
        '
        Me.chkCompanyShowActive.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkCompanyShowActive.Location = New System.Drawing.Point(635, 8)
        Me.chkCompanyShowActive.Name = "chkCompanyShowActive"
        Me.chkCompanyShowActive.Size = New System.Drawing.Size(144, 16)
        Me.chkCompanyShowActive.TabIndex = 2
        Me.chkCompanyShowActive.Tag = "646"
        Me.chkCompanyShowActive.Text = "Show Active Only"
        '
        'chkCompanyShowRelatedContacts
        '
        Me.chkCompanyShowRelatedContacts.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkCompanyShowRelatedContacts.Location = New System.Drawing.Point(467, 8)
        Me.chkCompanyShowRelatedContacts.Name = "chkCompanyShowRelatedContacts"
        Me.chkCompanyShowRelatedContacts.Size = New System.Drawing.Size(160, 16)
        Me.chkCompanyShowRelatedContacts.TabIndex = 1
        Me.chkCompanyShowRelatedContacts.Tag = "645"
        Me.chkCompanyShowRelatedContacts.Text = "Show Related Contacts"
        '
        'chkCompanyShowContactsforAllModules
        '
        Me.chkCompanyShowContactsforAllModules.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkCompanyShowContactsforAllModules.Location = New System.Drawing.Point(251, 8)
        Me.chkCompanyShowContactsforAllModules.Name = "chkCompanyShowContactsforAllModules"
        Me.chkCompanyShowContactsforAllModules.Size = New System.Drawing.Size(200, 16)
        Me.chkCompanyShowContactsforAllModules.TabIndex = 0
        Me.chkCompanyShowContactsforAllModules.Tag = "644"
        Me.chkCompanyShowContactsforAllModules.Text = "Show Contacts for All Modules"
        '
        'lblCompanyContacts
        '
        Me.lblCompanyContacts.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompanyContacts.Location = New System.Drawing.Point(16, 8)
        Me.lblCompanyContacts.Name = "lblCompanyContacts"
        Me.lblCompanyContacts.Size = New System.Drawing.Size(104, 16)
        Me.lblCompanyContacts.TabIndex = 139
        Me.lblCompanyContacts.Text = "Company Contacts"
        '
        'pnlERACandContacts
        '
        Me.pnlERACandContacts.Controls.Add(Me.lblCompanyContactsHead)
        Me.pnlERACandContacts.Controls.Add(Me.lblCompanyContactsDisplay)
        Me.pnlERACandContacts.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlERACandContacts.Location = New System.Drawing.Point(0, 1330)
        Me.pnlERACandContacts.Name = "pnlERACandContacts"
        Me.pnlERACandContacts.Size = New System.Drawing.Size(976, 24)
        Me.pnlERACandContacts.TabIndex = 222
        '
        'lblCompanyContactsHead
        '
        Me.lblCompanyContactsHead.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblCompanyContactsHead.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblCompanyContactsHead.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblCompanyContactsHead.Location = New System.Drawing.Point(16, 0)
        Me.lblCompanyContactsHead.Name = "lblCompanyContactsHead"
        Me.lblCompanyContactsHead.Size = New System.Drawing.Size(960, 24)
        Me.lblCompanyContactsHead.TabIndex = 1
        Me.lblCompanyContactsHead.Text = "Contacts"
        Me.lblCompanyContactsHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblCompanyContactsDisplay
        '
        Me.lblCompanyContactsDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblCompanyContactsDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblCompanyContactsDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblCompanyContactsDisplay.Name = "lblCompanyContactsDisplay"
        Me.lblCompanyContactsDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblCompanyContactsDisplay.TabIndex = 0
        Me.lblCompanyContactsDisplay.Text = "-"
        Me.lblCompanyContactsDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlDocumentDetails
        '
        Me.pnlDocumentDetails.Controls.Add(Me.UCCompanyDocuments)
        Me.pnlDocumentDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlDocumentDetails.Location = New System.Drawing.Point(0, 1170)
        Me.pnlDocumentDetails.Name = "pnlDocumentDetails"
        Me.pnlDocumentDetails.Size = New System.Drawing.Size(976, 160)
        Me.pnlDocumentDetails.TabIndex = 221
        '
        'UCCompanyDocuments
        '
        Me.UCCompanyDocuments.AutoScroll = True
        Me.UCCompanyDocuments.Location = New System.Drawing.Point(16, 8)
        Me.UCCompanyDocuments.Name = "UCCompanyDocuments"
        Me.UCCompanyDocuments.Size = New System.Drawing.Size(800, 144)
        Me.UCCompanyDocuments.TabIndex = 0
        '
        'pnlDocumentsHead
        '
        Me.pnlDocumentsHead.Controls.Add(Me.lblDocumentsHead)
        Me.pnlDocumentsHead.Controls.Add(Me.lblDocumentsDisplay)
        Me.pnlDocumentsHead.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlDocumentsHead.Location = New System.Drawing.Point(0, 1146)
        Me.pnlDocumentsHead.Name = "pnlDocumentsHead"
        Me.pnlDocumentsHead.Size = New System.Drawing.Size(976, 24)
        Me.pnlDocumentsHead.TabIndex = 220
        '
        'lblDocumentsHead
        '
        Me.lblDocumentsHead.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblDocumentsHead.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblDocumentsHead.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblDocumentsHead.Location = New System.Drawing.Point(16, 0)
        Me.lblDocumentsHead.Name = "lblDocumentsHead"
        Me.lblDocumentsHead.Size = New System.Drawing.Size(960, 24)
        Me.lblDocumentsHead.TabIndex = 1
        Me.lblDocumentsHead.Text = "Documents"
        Me.lblDocumentsHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblDocumentsDisplay
        '
        Me.lblDocumentsDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblDocumentsDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblDocumentsDisplay.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDocumentsDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblDocumentsDisplay.Name = "lblDocumentsDisplay"
        Me.lblDocumentsDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblDocumentsDisplay.TabIndex = 0
        Me.lblDocumentsDisplay.Text = "-"
        Me.lblDocumentsDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlPriorLicenseesDetails
        '
        Me.pnlPriorLicenseesDetails.Controls.Add(Me.lblPriorLicensees)
        Me.pnlPriorLicenseesDetails.Controls.Add(Me.ugPriorLicensees)
        Me.pnlPriorLicenseesDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlPriorLicenseesDetails.Location = New System.Drawing.Point(0, 972)
        Me.pnlPriorLicenseesDetails.Name = "pnlPriorLicenseesDetails"
        Me.pnlPriorLicenseesDetails.Size = New System.Drawing.Size(976, 174)
        Me.pnlPriorLicenseesDetails.TabIndex = 219
        '
        'lblPriorLicensees
        '
        Me.lblPriorLicensees.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPriorLicensees.Location = New System.Drawing.Point(16, 8)
        Me.lblPriorLicensees.Name = "lblPriorLicensees"
        Me.lblPriorLicensees.Size = New System.Drawing.Size(96, 16)
        Me.lblPriorLicensees.TabIndex = 207
        Me.lblPriorLicensees.Text = "Prior Licensees:"
        Me.lblPriorLicensees.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ugPriorLicensees
        '
        Me.ugPriorLicensees.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugPriorLicensees.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugPriorLicensees.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugPriorLicensees.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugPriorLicensees.Location = New System.Drawing.Point(24, 24)
        Me.ugPriorLicensees.Name = "ugPriorLicensees"
        Me.ugPriorLicensees.Size = New System.Drawing.Size(793, 144)
        Me.ugPriorLicensees.TabIndex = 18
        '
        'pnlPriorCompaniesHead
        '
        Me.pnlPriorCompaniesHead.Controls.Add(Me.lblPriorLicenseesHead)
        Me.pnlPriorCompaniesHead.Controls.Add(Me.lblPriorLicenseesDisplay)
        Me.pnlPriorCompaniesHead.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlPriorCompaniesHead.Location = New System.Drawing.Point(0, 948)
        Me.pnlPriorCompaniesHead.Name = "pnlPriorCompaniesHead"
        Me.pnlPriorCompaniesHead.Size = New System.Drawing.Size(976, 24)
        Me.pnlPriorCompaniesHead.TabIndex = 218
        '
        'lblPriorLicenseesHead
        '
        Me.lblPriorLicenseesHead.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblPriorLicenseesHead.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblPriorLicenseesHead.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblPriorLicenseesHead.Location = New System.Drawing.Point(16, 0)
        Me.lblPriorLicenseesHead.Name = "lblPriorLicenseesHead"
        Me.lblPriorLicenseesHead.Size = New System.Drawing.Size(960, 24)
        Me.lblPriorLicenseesHead.TabIndex = 1
        Me.lblPriorLicenseesHead.Text = "Prior Licensees"
        Me.lblPriorLicenseesHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblPriorLicenseesDisplay
        '
        Me.lblPriorLicenseesDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPriorLicenseesDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblPriorLicenseesDisplay.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPriorLicenseesDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblPriorLicenseesDisplay.Name = "lblPriorLicenseesDisplay"
        Me.lblPriorLicenseesDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblPriorLicenseesDisplay.TabIndex = 0
        Me.lblPriorLicenseesDisplay.Text = "-"
        Me.lblPriorLicenseesDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlTypesDetails
        '
        Me.pnlTypesDetails.Controls.Add(Me.grpTypes)
        Me.pnlTypesDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTypesDetails.Location = New System.Drawing.Point(0, 748)
        Me.pnlTypesDetails.Name = "pnlTypesDetails"
        Me.pnlTypesDetails.Size = New System.Drawing.Size(976, 200)
        Me.pnlTypesDetails.TabIndex = 217
        '
        'grpTypes
        '
        Me.grpTypes.Controls.Add(Me.chkManager)
        Me.grpTypes.Controls.Add(Me.chkCorrosionExpert)
        Me.grpTypes.Controls.Add(Me.chkTankLining)
        Me.grpTypes.Controls.Add(Me.chkEnvironmentalDrillers)
        Me.grpTypes.Controls.Add(Me.chkEnvironmentalConsultants)
        Me.grpTypes.Controls.Add(Me.chkIRAC)
        Me.grpTypes.Controls.Add(Me.chkERAC)
        Me.grpTypes.Controls.Add(Me.chkThermalSoilTreatment)
        Me.grpTypes.Controls.Add(Me.chkWastewaterLaboratories)
        Me.grpTypes.Controls.Add(Me.chkUSTSuppliesAndEquipment)
        Me.grpTypes.Controls.Add(Me.chkLeakDetectionCompanies)
        Me.grpTypes.Controls.Add(Me.chkPrecisionTankTightnessTesters)
        Me.grpTypes.Controls.Add(Me.chkCertifiedToCloseUSTs)
        Me.grpTypes.Controls.Add(Me.chkCertifiedtoInstallAlterandCloseUSTs)
        Me.grpTypes.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpTypes.Location = New System.Drawing.Point(14, 8)
        Me.grpTypes.Name = "grpTypes"
        Me.grpTypes.Size = New System.Drawing.Size(584, 192)
        Me.grpTypes.TabIndex = 17
        Me.grpTypes.TabStop = False
        Me.grpTypes.Text = "Types:  "
        '
        'chkManager
        '
        Me.chkManager.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkManager.Location = New System.Drawing.Point(416, 168)
        Me.chkManager.Name = "chkManager"
        Me.chkManager.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkManager.Size = New System.Drawing.Size(144, 16)
        Me.chkManager.TabIndex = 13
        Me.chkManager.Tag = "649"
        Me.chkManager.Text = ":Compliance Manager"
        Me.chkManager.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'chkCorrosionExpert
        '
        Me.chkCorrosionExpert.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkCorrosionExpert.Location = New System.Drawing.Point(432, 148)
        Me.chkCorrosionExpert.Name = "chkCorrosionExpert"
        Me.chkCorrosionExpert.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkCorrosionExpert.Size = New System.Drawing.Size(128, 16)
        Me.chkCorrosionExpert.TabIndex = 12
        Me.chkCorrosionExpert.Tag = "649"
        Me.chkCorrosionExpert.Text = " :Corrosion Expert"
        Me.chkCorrosionExpert.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'chkTankLining
        '
        Me.chkTankLining.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkTankLining.Location = New System.Drawing.Point(464, 124)
        Me.chkTankLining.Name = "chkTankLining"
        Me.chkTankLining.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkTankLining.Size = New System.Drawing.Size(96, 16)
        Me.chkTankLining.TabIndex = 11
        Me.chkTankLining.Tag = "648"
        Me.chkTankLining.Text = " :Tank Lining"
        Me.chkTankLining.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'chkEnvironmentalDrillers
        '
        Me.chkEnvironmentalDrillers.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkEnvironmentalDrillers.Location = New System.Drawing.Point(400, 100)
        Me.chkEnvironmentalDrillers.Name = "chkEnvironmentalDrillers"
        Me.chkEnvironmentalDrillers.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkEnvironmentalDrillers.Size = New System.Drawing.Size(160, 16)
        Me.chkEnvironmentalDrillers.TabIndex = 10
        Me.chkEnvironmentalDrillers.Tag = "647"
        Me.chkEnvironmentalDrillers.Text = " :Environmental Drillers"
        Me.chkEnvironmentalDrillers.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'chkEnvironmentalConsultants
        '
        Me.chkEnvironmentalConsultants.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkEnvironmentalConsultants.Location = New System.Drawing.Point(376, 76)
        Me.chkEnvironmentalConsultants.Name = "chkEnvironmentalConsultants"
        Me.chkEnvironmentalConsultants.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkEnvironmentalConsultants.Size = New System.Drawing.Size(184, 16)
        Me.chkEnvironmentalConsultants.TabIndex = 9
        Me.chkEnvironmentalConsultants.Tag = "646"
        Me.chkEnvironmentalConsultants.Text = " :Environmental Consultants"
        Me.chkEnvironmentalConsultants.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'chkIRAC
        '
        Me.chkIRAC.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkIRAC.Location = New System.Drawing.Point(496, 52)
        Me.chkIRAC.Name = "chkIRAC"
        Me.chkIRAC.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkIRAC.Size = New System.Drawing.Size(64, 16)
        Me.chkIRAC.TabIndex = 8
        Me.chkIRAC.Tag = "645"
        Me.chkIRAC.Text = " :IRAC"
        Me.chkIRAC.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'chkERAC
        '
        Me.chkERAC.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkERAC.Location = New System.Drawing.Point(488, 28)
        Me.chkERAC.Name = "chkERAC"
        Me.chkERAC.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkERAC.Size = New System.Drawing.Size(72, 16)
        Me.chkERAC.TabIndex = 7
        Me.chkERAC.Tag = "644"
        Me.chkERAC.Text = " :ERAC"
        Me.chkERAC.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'chkThermalSoilTreatment
        '
        Me.chkThermalSoilTreatment.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkThermalSoilTreatment.Location = New System.Drawing.Point(104, 168)
        Me.chkThermalSoilTreatment.Name = "chkThermalSoilTreatment"
        Me.chkThermalSoilTreatment.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkThermalSoilTreatment.Size = New System.Drawing.Size(168, 16)
        Me.chkThermalSoilTreatment.TabIndex = 6
        Me.chkThermalSoilTreatment.Tag = "649"
        Me.chkThermalSoilTreatment.Text = " :Thermal Soil Treatment"
        Me.chkThermalSoilTreatment.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'chkWastewaterLaboratories
        '
        Me.chkWastewaterLaboratories.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkWastewaterLaboratories.Location = New System.Drawing.Point(104, 144)
        Me.chkWastewaterLaboratories.Name = "chkWastewaterLaboratories"
        Me.chkWastewaterLaboratories.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkWastewaterLaboratories.Size = New System.Drawing.Size(168, 16)
        Me.chkWastewaterLaboratories.TabIndex = 5
        Me.chkWastewaterLaboratories.Tag = "649"
        Me.chkWastewaterLaboratories.Text = " :Wastewater Laboratories"
        Me.chkWastewaterLaboratories.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'chkUSTSuppliesAndEquipment
        '
        Me.chkUSTSuppliesAndEquipment.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkUSTSuppliesAndEquipment.Location = New System.Drawing.Point(80, 120)
        Me.chkUSTSuppliesAndEquipment.Name = "chkUSTSuppliesAndEquipment"
        Me.chkUSTSuppliesAndEquipment.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkUSTSuppliesAndEquipment.Size = New System.Drawing.Size(192, 16)
        Me.chkUSTSuppliesAndEquipment.TabIndex = 4
        Me.chkUSTSuppliesAndEquipment.Tag = "648"
        Me.chkUSTSuppliesAndEquipment.Text = " :UST Supplies and Equipment"
        Me.chkUSTSuppliesAndEquipment.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'chkLeakDetectionCompanies
        '
        Me.chkLeakDetectionCompanies.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkLeakDetectionCompanies.Location = New System.Drawing.Point(80, 96)
        Me.chkLeakDetectionCompanies.Name = "chkLeakDetectionCompanies"
        Me.chkLeakDetectionCompanies.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkLeakDetectionCompanies.Size = New System.Drawing.Size(192, 16)
        Me.chkLeakDetectionCompanies.TabIndex = 3
        Me.chkLeakDetectionCompanies.Tag = "647"
        Me.chkLeakDetectionCompanies.Text = " :Leak Detection Companies"
        Me.chkLeakDetectionCompanies.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'chkPrecisionTankTightnessTesters
        '
        Me.chkPrecisionTankTightnessTesters.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkPrecisionTankTightnessTesters.Location = New System.Drawing.Point(56, 72)
        Me.chkPrecisionTankTightnessTesters.Name = "chkPrecisionTankTightnessTesters"
        Me.chkPrecisionTankTightnessTesters.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkPrecisionTankTightnessTesters.Size = New System.Drawing.Size(216, 16)
        Me.chkPrecisionTankTightnessTesters.TabIndex = 2
        Me.chkPrecisionTankTightnessTesters.Tag = "646"
        Me.chkPrecisionTankTightnessTesters.Text = " :Precision Tank Tightness Testers"
        Me.chkPrecisionTankTightnessTesters.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'chkCertifiedToCloseUSTs
        '
        Me.chkCertifiedToCloseUSTs.Enabled = False
        Me.chkCertifiedToCloseUSTs.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkCertifiedToCloseUSTs.Location = New System.Drawing.Point(96, 48)
        Me.chkCertifiedToCloseUSTs.Name = "chkCertifiedToCloseUSTs"
        Me.chkCertifiedToCloseUSTs.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkCertifiedToCloseUSTs.Size = New System.Drawing.Size(176, 16)
        Me.chkCertifiedToCloseUSTs.TabIndex = 1
        Me.chkCertifiedToCloseUSTs.TabStop = False
        Me.chkCertifiedToCloseUSTs.Tag = "645"
        Me.chkCertifiedToCloseUSTs.Text = " :Certified to Close USTs"
        Me.chkCertifiedToCloseUSTs.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'chkCertifiedtoInstallAlterandCloseUSTs
        '
        Me.chkCertifiedtoInstallAlterandCloseUSTs.Enabled = False
        Me.chkCertifiedtoInstallAlterandCloseUSTs.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkCertifiedtoInstallAlterandCloseUSTs.Location = New System.Drawing.Point(24, 24)
        Me.chkCertifiedtoInstallAlterandCloseUSTs.Name = "chkCertifiedtoInstallAlterandCloseUSTs"
        Me.chkCertifiedtoInstallAlterandCloseUSTs.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkCertifiedtoInstallAlterandCloseUSTs.Size = New System.Drawing.Size(248, 16)
        Me.chkCertifiedtoInstallAlterandCloseUSTs.TabIndex = 0
        Me.chkCertifiedtoInstallAlterandCloseUSTs.TabStop = False
        Me.chkCertifiedtoInstallAlterandCloseUSTs.Tag = "644"
        Me.chkCertifiedtoInstallAlterandCloseUSTs.Text = " :Certified to Install, Alter and Close USTs"
        Me.chkCertifiedtoInstallAlterandCloseUSTs.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'pnlTypesHead
        '
        Me.pnlTypesHead.Controls.Add(Me.lblTypesHead)
        Me.pnlTypesHead.Controls.Add(Me.lblTypesDisplay)
        Me.pnlTypesHead.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTypesHead.Location = New System.Drawing.Point(0, 724)
        Me.pnlTypesHead.Name = "pnlTypesHead"
        Me.pnlTypesHead.Size = New System.Drawing.Size(976, 24)
        Me.pnlTypesHead.TabIndex = 216
        '
        'lblTypesHead
        '
        Me.lblTypesHead.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblTypesHead.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblTypesHead.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblTypesHead.Location = New System.Drawing.Point(16, 0)
        Me.lblTypesHead.Name = "lblTypesHead"
        Me.lblTypesHead.Size = New System.Drawing.Size(960, 24)
        Me.lblTypesHead.TabIndex = 1
        Me.lblTypesHead.Text = "Types"
        Me.lblTypesHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblTypesDisplay
        '
        Me.lblTypesDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTypesDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblTypesDisplay.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTypesDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblTypesDisplay.Name = "lblTypesDisplay"
        Me.lblTypesDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblTypesDisplay.TabIndex = 0
        Me.lblTypesDisplay.Text = "-"
        Me.lblTypesDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlEracDetails
        '
        Me.pnlEracDetails.Controls.Add(Me.grpERACs)
        Me.pnlEracDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlEracDetails.Location = New System.Drawing.Point(0, 552)
        Me.pnlEracDetails.Name = "pnlEracDetails"
        Me.pnlEracDetails.Size = New System.Drawing.Size(976, 172)
        Me.pnlEracDetails.TabIndex = 215
        '
        'grpERACs
        '
        Me.grpERACs.Controls.Add(Me.TxtGeologistEmail)
        Me.grpERACs.Controls.Add(Me.TxtEngineerEmail)
        Me.grpERACs.Controls.Add(Me.dtEngineerLiabilityExpiration)
        Me.grpERACs.Controls.Add(Me.dtEngineerAppApprovalDate)
        Me.grpERACs.Controls.Add(Me.lblEngineerAppApprovalDate)
        Me.grpERACs.Controls.Add(Me.lblEngineerLiabilityExpiration)
        Me.grpERACs.Controls.Add(Me.txtGeologistNo)
        Me.grpERACs.Controls.Add(Me.Label6)
        Me.grpERACs.Controls.Add(Me.lblGeologistEmail)
        Me.grpERACs.Controls.Add(Me.txtEngineerNo)
        Me.grpERACs.Controls.Add(Me.Label3)
        Me.grpERACs.Controls.Add(Me.lblEngineerEmail)
        Me.grpERACs.Controls.Add(Me.txtProfGeologist)
        Me.grpERACs.Controls.Add(Me.txtProfEngineer)
        Me.grpERACs.Controls.Add(Me.lblProfGeologist)
        Me.grpERACs.Controls.Add(Me.lblProfEngineer)
        Me.grpERACs.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpERACs.Location = New System.Drawing.Point(13, 8)
        Me.grpERACs.Name = "grpERACs"
        Me.grpERACs.Size = New System.Drawing.Size(835, 160)
        Me.grpERACs.TabIndex = 10
        Me.grpERACs.TabStop = False
        Me.grpERACs.Text = "ERACs"
        '
        'TxtGeologistEmail
        '
        Me.TxtGeologistEmail.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtGeologistEmail.Location = New System.Drawing.Point(544, 128)
        Me.TxtGeologistEmail.Name = "TxtGeologistEmail"
        Me.TxtGeologistEmail.Size = New System.Drawing.Size(240, 20)
        Me.TxtGeologistEmail.TabIndex = 218
        Me.TxtGeologistEmail.Text = ""
        '
        'TxtEngineerEmail
        '
        Me.TxtEngineerEmail.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtEngineerEmail.Location = New System.Drawing.Point(160, 128)
        Me.TxtEngineerEmail.Name = "TxtEngineerEmail"
        Me.TxtEngineerEmail.Size = New System.Drawing.Size(240, 20)
        Me.TxtEngineerEmail.TabIndex = 217
        Me.TxtEngineerEmail.Text = ""
        '
        'dtEngineerLiabilityExpiration
        '
        Me.dtEngineerLiabilityExpiration.Checked = False
        Me.dtEngineerLiabilityExpiration.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtEngineerLiabilityExpiration.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtEngineerLiabilityExpiration.Location = New System.Drawing.Point(160, 48)
        Me.dtEngineerLiabilityExpiration.Name = "dtEngineerLiabilityExpiration"
        Me.dtEngineerLiabilityExpiration.ShowCheckBox = True
        Me.dtEngineerLiabilityExpiration.Size = New System.Drawing.Size(104, 20)
        Me.dtEngineerLiabilityExpiration.TabIndex = 214
        '
        'dtEngineerAppApprovalDate
        '
        Me.dtEngineerAppApprovalDate.Checked = False
        Me.dtEngineerAppApprovalDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtEngineerAppApprovalDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtEngineerAppApprovalDate.Location = New System.Drawing.Point(160, 24)
        Me.dtEngineerAppApprovalDate.Name = "dtEngineerAppApprovalDate"
        Me.dtEngineerAppApprovalDate.ShowCheckBox = True
        Me.dtEngineerAppApprovalDate.Size = New System.Drawing.Size(104, 20)
        Me.dtEngineerAppApprovalDate.TabIndex = 213
        '
        'lblEngineerAppApprovalDate
        '
        Me.lblEngineerAppApprovalDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEngineerAppApprovalDate.Location = New System.Drawing.Point(8, 24)
        Me.lblEngineerAppApprovalDate.Name = "lblEngineerAppApprovalDate"
        Me.lblEngineerAppApprovalDate.Size = New System.Drawing.Size(144, 16)
        Me.lblEngineerAppApprovalDate.TabIndex = 216
        Me.lblEngineerAppApprovalDate.Text = "ERAC App Approval Date:"
        Me.lblEngineerAppApprovalDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblEngineerLiabilityExpiration
        '
        Me.lblEngineerLiabilityExpiration.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEngineerLiabilityExpiration.Location = New System.Drawing.Point(24, 48)
        Me.lblEngineerLiabilityExpiration.Name = "lblEngineerLiabilityExpiration"
        Me.lblEngineerLiabilityExpiration.Size = New System.Drawing.Size(128, 24)
        Me.lblEngineerLiabilityExpiration.TabIndex = 215
        Me.lblEngineerLiabilityExpiration.Text = "Professional Liability Expiration:"
        Me.lblEngineerLiabilityExpiration.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtGeologistNo
        '
        Me.txtGeologistNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtGeologistNo.Location = New System.Drawing.Point(544, 104)
        Me.txtGeologistNo.Name = "txtGeologistNo"
        Me.txtGeologistNo.Size = New System.Drawing.Size(128, 20)
        Me.txtGeologistNo.TabIndex = 9
        Me.txtGeologistNo.Text = ""
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(472, 104)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(64, 17)
        Me.Label6.TabIndex = 212
        Me.Label6.Text = "Number:"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblGeologistEmail
        '
        Me.lblGeologistEmail.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGeologistEmail.Location = New System.Drawing.Point(472, 128)
        Me.lblGeologistEmail.Name = "lblGeologistEmail"
        Me.lblGeologistEmail.Size = New System.Drawing.Size(64, 16)
        Me.lblGeologistEmail.TabIndex = 211
        Me.lblGeologistEmail.Text = "Email:"
        Me.lblGeologistEmail.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtEngineerNo
        '
        Me.txtEngineerNo.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtEngineerNo.Location = New System.Drawing.Point(160, 104)
        Me.txtEngineerNo.Name = "txtEngineerNo"
        Me.txtEngineerNo.Size = New System.Drawing.Size(128, 20)
        Me.txtEngineerNo.TabIndex = 3
        Me.txtEngineerNo.Text = ""
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(88, 104)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(64, 17)
        Me.Label3.TabIndex = 191
        Me.Label3.Text = "Number:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblEngineerEmail
        '
        Me.lblEngineerEmail.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEngineerEmail.Location = New System.Drawing.Point(72, 128)
        Me.lblEngineerEmail.Name = "lblEngineerEmail"
        Me.lblEngineerEmail.Size = New System.Drawing.Size(80, 16)
        Me.lblEngineerEmail.TabIndex = 190
        Me.lblEngineerEmail.Text = "Email: "
        Me.lblEngineerEmail.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtProfGeologist
        '
        Me.txtProfGeologist.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtProfGeologist.Location = New System.Drawing.Point(544, 80)
        Me.txtProfGeologist.Name = "txtProfGeologist"
        Me.txtProfGeologist.Size = New System.Drawing.Size(240, 20)
        Me.txtProfGeologist.TabIndex = 6
        Me.txtProfGeologist.Text = ""
        '
        'txtProfEngineer
        '
        Me.txtProfEngineer.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtProfEngineer.Location = New System.Drawing.Point(160, 80)
        Me.txtProfEngineer.Name = "txtProfEngineer"
        Me.txtProfEngineer.Size = New System.Drawing.Size(240, 20)
        Me.txtProfEngineer.TabIndex = 0
        Me.txtProfEngineer.Text = ""
        '
        'lblProfGeologist
        '
        Me.lblProfGeologist.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProfGeologist.Location = New System.Drawing.Point(416, 80)
        Me.lblProfGeologist.Name = "lblProfGeologist"
        Me.lblProfGeologist.Size = New System.Drawing.Size(120, 17)
        Me.lblProfGeologist.TabIndex = 36
        Me.lblProfGeologist.Text = "Professional Geologist"
        Me.lblProfGeologist.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblProfEngineer
        '
        Me.lblProfEngineer.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProfEngineer.Location = New System.Drawing.Point(24, 80)
        Me.lblProfEngineer.Name = "lblProfEngineer"
        Me.lblProfEngineer.Size = New System.Drawing.Size(128, 17)
        Me.lblProfEngineer.TabIndex = 34
        Me.lblProfEngineer.Text = "Professional Engineer:"
        Me.lblProfEngineer.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlEracHead
        '
        Me.pnlEracHead.Controls.Add(Me.lnlEracHead)
        Me.pnlEracHead.Controls.Add(Me.lblEracDisplay)
        Me.pnlEracHead.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlEracHead.Location = New System.Drawing.Point(0, 528)
        Me.pnlEracHead.Name = "pnlEracHead"
        Me.pnlEracHead.Size = New System.Drawing.Size(976, 24)
        Me.pnlEracHead.TabIndex = 214
        '
        'lnlEracHead
        '
        Me.lnlEracHead.BackColor = System.Drawing.SystemColors.Highlight
        Me.lnlEracHead.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lnlEracHead.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lnlEracHead.Location = New System.Drawing.Point(16, 0)
        Me.lnlEracHead.Name = "lnlEracHead"
        Me.lnlEracHead.Size = New System.Drawing.Size(960, 24)
        Me.lnlEracHead.TabIndex = 1
        Me.lnlEracHead.Text = "ERAC"
        Me.lnlEracHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblEracDisplay
        '
        Me.lblEracDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblEracDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblEracDisplay.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEracDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblEracDisplay.Name = "lblEracDisplay"
        Me.lblEracDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblEracDisplay.TabIndex = 0
        Me.lblEracDisplay.Text = "-"
        Me.lblEracDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlLicenseesDetails
        '
        Me.pnlLicenseesDetails.Controls.Add(Me.lblLicensees)
        Me.pnlLicenseesDetails.Controls.Add(Me.btnLicenseesAddExisting)
        Me.pnlLicenseesDetails.Controls.Add(Me.btnLicenseesAddNew)
        Me.pnlLicenseesDetails.Controls.Add(Me.btnLicenseesEdit)
        Me.pnlLicenseesDetails.Controls.Add(Me.btnDisassociate)
        Me.pnlLicenseesDetails.Controls.Add(Me.ugLicensees)
        Me.pnlLicenseesDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlLicenseesDetails.Location = New System.Drawing.Point(0, 372)
        Me.pnlLicenseesDetails.Name = "pnlLicenseesDetails"
        Me.pnlLicenseesDetails.Size = New System.Drawing.Size(976, 156)
        Me.pnlLicenseesDetails.TabIndex = 213
        '
        'lblLicensees
        '
        Me.lblLicensees.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLicensees.Location = New System.Drawing.Point(10, 5)
        Me.lblLicensees.Name = "lblLicensees"
        Me.lblLicensees.Size = New System.Drawing.Size(64, 16)
        Me.lblLicensees.TabIndex = 203
        Me.lblLicensees.Text = "Licensees:"
        Me.lblLicensees.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnLicenseesAddExisting
        '
        Me.btnLicenseesAddExisting.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnLicenseesAddExisting.Location = New System.Drawing.Point(800, 24)
        Me.btnLicenseesAddExisting.Name = "btnLicenseesAddExisting"
        Me.btnLicenseesAddExisting.Size = New System.Drawing.Size(160, 26)
        Me.btnLicenseesAddExisting.TabIndex = 12
        Me.btnLicenseesAddExisting.Text = "Add Existing Licensee"
        '
        'btnLicenseesAddNew
        '
        Me.btnLicenseesAddNew.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnLicenseesAddNew.Location = New System.Drawing.Point(800, 56)
        Me.btnLicenseesAddNew.Name = "btnLicenseesAddNew"
        Me.btnLicenseesAddNew.Size = New System.Drawing.Size(160, 26)
        Me.btnLicenseesAddNew.TabIndex = 13
        Me.btnLicenseesAddNew.Text = "Add New Licensee"
        '
        'btnLicenseesEdit
        '
        Me.btnLicenseesEdit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnLicenseesEdit.Location = New System.Drawing.Point(800, 88)
        Me.btnLicenseesEdit.Name = "btnLicenseesEdit"
        Me.btnLicenseesEdit.Size = New System.Drawing.Size(160, 26)
        Me.btnLicenseesEdit.TabIndex = 14
        Me.btnLicenseesEdit.Text = "Edit Licensee"
        '
        'btnDisassociate
        '
        Me.btnDisassociate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDisassociate.Location = New System.Drawing.Point(800, 120)
        Me.btnDisassociate.Name = "btnDisassociate"
        Me.btnDisassociate.Size = New System.Drawing.Size(160, 26)
        Me.btnDisassociate.TabIndex = 15
        Me.btnDisassociate.Text = "Disassociate Licensee"
        '
        'ugLicensees
        '
        Me.ugLicensees.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugLicensees.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugLicensees.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugLicensees.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugLicensees.Location = New System.Drawing.Point(13, 24)
        Me.ugLicensees.Name = "ugLicensees"
        Me.ugLicensees.Size = New System.Drawing.Size(779, 128)
        Me.ugLicensees.TabIndex = 11
        '
        'pnlManagersDetails
        '
        Me.pnlManagersDetails.Controls.Add(Me.lblManagers)
        Me.pnlManagersDetails.Controls.Add(Me.btnManagersAddExisting)
        Me.pnlManagersDetails.Controls.Add(Me.btnManagersAddNew)
        Me.pnlManagersDetails.Controls.Add(Me.btnManagersEdit)
        Me.pnlManagersDetails.Controls.Add(Me.btnManagerDisassociate)
        Me.pnlManagersDetails.Controls.Add(Me.ugManagers)
        Me.pnlManagersDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlManagersDetails.Location = New System.Drawing.Point(0, 216)
        Me.pnlManagersDetails.Name = "pnlManagersDetails"
        Me.pnlManagersDetails.Size = New System.Drawing.Size(976, 156)
        Me.pnlManagersDetails.TabIndex = 213
        '
        'lblManagers
        '
        Me.lblManagers.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblManagers.Location = New System.Drawing.Point(10, 5)
        Me.lblManagers.Name = "lblManagers"
        Me.lblManagers.Size = New System.Drawing.Size(134, 16)
        Me.lblManagers.TabIndex = 203
        Me.lblManagers.Text = "Compliance Managers:"
        Me.lblManagers.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnManagersAddExisting
        '
        Me.btnManagersAddExisting.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnManagersAddExisting.Location = New System.Drawing.Point(800, 24)
        Me.btnManagersAddExisting.Name = "btnManagersAddExisting"
        Me.btnManagersAddExisting.Size = New System.Drawing.Size(152, 26)
        Me.btnManagersAddExisting.TabIndex = 12
        Me.btnManagersAddExisting.Text = "Add Existing Manager"
        '
        'btnManagersAddNew
        '
        Me.btnManagersAddNew.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnManagersAddNew.Location = New System.Drawing.Point(800, 56)
        Me.btnManagersAddNew.Name = "btnManagersAddNew"
        Me.btnManagersAddNew.Size = New System.Drawing.Size(152, 26)
        Me.btnManagersAddNew.TabIndex = 13
        Me.btnManagersAddNew.Text = "Add New Manager"
        '
        'btnManagersEdit
        '
        Me.btnManagersEdit.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnManagersEdit.Location = New System.Drawing.Point(800, 88)
        Me.btnManagersEdit.Name = "btnManagersEdit"
        Me.btnManagersEdit.Size = New System.Drawing.Size(152, 26)
        Me.btnManagersEdit.TabIndex = 14
        Me.btnManagersEdit.Text = "Edit Manager"
        '
        'btnManagerDisassociate
        '
        Me.btnManagerDisassociate.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnManagerDisassociate.Location = New System.Drawing.Point(800, 120)
        Me.btnManagerDisassociate.Name = "btnManagerDisassociate"
        Me.btnManagerDisassociate.Size = New System.Drawing.Size(152, 26)
        Me.btnManagerDisassociate.TabIndex = 14
        Me.btnManagerDisassociate.Text = "Disassociate a Manager"
        '
        'ugManagers
        '
        Me.ugManagers.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugManagers.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugManagers.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugManagers.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugManagers.Location = New System.Drawing.Point(13, 24)
        Me.ugManagers.Name = "ugManagers"
        Me.ugManagers.Size = New System.Drawing.Size(779, 128)
        Me.ugManagers.TabIndex = 11
        '
        'pnlLicenseeHead
        '
        Me.pnlLicenseeHead.Controls.Add(Me.lblLicenseeHead)
        Me.pnlLicenseeHead.Controls.Add(Me.lblLicenseeDisplay)
        Me.pnlLicenseeHead.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlLicenseeHead.Location = New System.Drawing.Point(0, 192)
        Me.pnlLicenseeHead.Name = "pnlLicenseeHead"
        Me.pnlLicenseeHead.Size = New System.Drawing.Size(976, 24)
        Me.pnlLicenseeHead.TabIndex = 212
        '
        'lblLicenseeHead
        '
        Me.lblLicenseeHead.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblLicenseeHead.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblLicenseeHead.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblLicenseeHead.Location = New System.Drawing.Point(16, 0)
        Me.lblLicenseeHead.Name = "lblLicenseeHead"
        Me.lblLicenseeHead.Size = New System.Drawing.Size(960, 24)
        Me.lblLicenseeHead.TabIndex = 1
        Me.lblLicenseeHead.Text = "Licensee"
        Me.lblLicenseeHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblLicenseeDisplay
        '
        Me.lblLicenseeDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLicenseeDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblLicenseeDisplay.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLicenseeDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblLicenseeDisplay.Name = "lblLicenseeDisplay"
        Me.lblLicenseeDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblLicenseeDisplay.TabIndex = 0
        Me.lblLicenseeDisplay.Text = "-"
        Me.lblLicenseeDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlManagerHead
        '
        Me.pnlManagerHead.Controls.Add(Me.lblManagerHead)
        Me.pnlManagerHead.Controls.Add(Me.lblManagerDisplay)
        Me.pnlManagerHead.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlManagerHead.Location = New System.Drawing.Point(0, 168)
        Me.pnlManagerHead.Name = "pnlManagerHead"
        Me.pnlManagerHead.Size = New System.Drawing.Size(976, 24)
        Me.pnlManagerHead.TabIndex = 212
        '
        'lblManagerHead
        '
        Me.lblManagerHead.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblManagerHead.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblManagerHead.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblManagerHead.Location = New System.Drawing.Point(16, 0)
        Me.lblManagerHead.Name = "lblManagerHead"
        Me.lblManagerHead.Size = New System.Drawing.Size(960, 24)
        Me.lblManagerHead.TabIndex = 1
        Me.lblManagerHead.Text = "Compliance Manager"
        Me.lblManagerHead.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblManagerDisplay
        '
        Me.lblManagerDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblManagerDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblManagerDisplay.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblManagerDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblManagerDisplay.Name = "lblManagerDisplay"
        Me.lblManagerDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblManagerDisplay.TabIndex = 0
        Me.lblManagerDisplay.Text = "-"
        Me.lblManagerDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlCompanyDetails
        '
        Me.pnlCompanyDetails.Controls.Add(Me.txtCompanyAddress)
        Me.pnlCompanyDetails.Controls.Add(Me.lblAddresses)
        Me.pnlCompanyDetails.Controls.Add(Me.btnLabels)
        Me.pnlCompanyDetails.Controls.Add(Me.btnEnvelopes)
        Me.pnlCompanyDetails.Controls.Add(Me.cmbCertORResponsibility)
        Me.pnlCompanyDetails.Controls.Add(Me.lblEmail)
        Me.pnlCompanyDetails.Controls.Add(Me.lblCompany)
        Me.pnlCompanyDetails.Controls.Add(Me.txtEmail)
        Me.pnlCompanyDetails.Controls.Add(Me.txtCompany)
        Me.pnlCompanyDetails.Controls.Add(Me.txtCompanyID)
        Me.pnlCompanyDetails.Controls.Add(Me.lblFinancialRespExpirationDate)
        Me.pnlCompanyDetails.Controls.Add(Me.lblCertORResponsibility)
        Me.pnlCompanyDetails.Controls.Add(Me.lblCompanyID)
        Me.pnlCompanyDetails.Controls.Add(Me.dtFinancialRespExpirationDate)
        Me.pnlCompanyDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlCompanyDetails.Location = New System.Drawing.Point(0, 24)
        Me.pnlCompanyDetails.Name = "pnlCompanyDetails"
        Me.pnlCompanyDetails.Size = New System.Drawing.Size(976, 144)
        Me.pnlCompanyDetails.TabIndex = 209
        '
        'txtCompanyAddress
        '
        Me.txtCompanyAddress.Location = New System.Drawing.Point(432, 8)
        Me.txtCompanyAddress.Multiline = True
        Me.txtCompanyAddress.Name = "txtCompanyAddress"
        Me.txtCompanyAddress.ReadOnly = True
        Me.txtCompanyAddress.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
        Me.txtCompanyAddress.Size = New System.Drawing.Size(256, 88)
        Me.txtCompanyAddress.TabIndex = 5
        Me.txtCompanyAddress.Text = ""
        Me.txtCompanyAddress.WordWrap = False
        '
        'lblAddresses
        '
        Me.lblAddresses.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAddresses.Location = New System.Drawing.Point(352, 8)
        Me.lblAddresses.Name = "lblAddresses"
        Me.lblAddresses.Size = New System.Drawing.Size(72, 16)
        Me.lblAddresses.TabIndex = 1060
        Me.lblAddresses.Text = "Address:"
        Me.lblAddresses.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnLabels
        '
        Me.btnLabels.Location = New System.Drawing.Point(360, 56)
        Me.btnLabels.Name = "btnLabels"
        Me.btnLabels.Size = New System.Drawing.Size(64, 23)
        Me.btnLabels.TabIndex = 1059
        Me.btnLabels.Text = "Labels"
        '
        'btnEnvelopes
        '
        Me.btnEnvelopes.Location = New System.Drawing.Point(360, 32)
        Me.btnEnvelopes.Name = "btnEnvelopes"
        Me.btnEnvelopes.Size = New System.Drawing.Size(65, 23)
        Me.btnEnvelopes.TabIndex = 1058
        Me.btnEnvelopes.Text = "Envelopes"
        '
        'cmbCertORResponsibility
        '
        Me.cmbCertORResponsibility.Items.AddRange(New Object() {"30-Day", "BCCR", "N/A"})
        Me.cmbCertORResponsibility.Location = New System.Drawing.Point(136, 80)
        Me.cmbCertORResponsibility.Name = "cmbCertORResponsibility"
        Me.cmbCertORResponsibility.Size = New System.Drawing.Size(104, 21)
        Me.cmbCertORResponsibility.TabIndex = 3
        '
        'lblEmail
        '
        Me.lblEmail.Location = New System.Drawing.Point(56, 56)
        Me.lblEmail.Name = "lblEmail"
        Me.lblEmail.Size = New System.Drawing.Size(77, 17)
        Me.lblEmail.TabIndex = 31
        Me.lblEmail.Text = "E-mail:"
        Me.lblEmail.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCompany
        '
        Me.lblCompany.Location = New System.Drawing.Point(56, 32)
        Me.lblCompany.Name = "lblCompany"
        Me.lblCompany.Size = New System.Drawing.Size(77, 17)
        Me.lblCompany.TabIndex = 30
        Me.lblCompany.Text = "Company:"
        Me.lblCompany.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtEmail
        '
        Me.txtEmail.Location = New System.Drawing.Point(136, 56)
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.Size = New System.Drawing.Size(128, 20)
        Me.txtEmail.TabIndex = 2
        Me.txtEmail.Text = ""
        '
        'txtCompany
        '
        Me.txtCompany.Location = New System.Drawing.Point(136, 32)
        Me.txtCompany.Name = "txtCompany"
        Me.txtCompany.Size = New System.Drawing.Size(208, 20)
        Me.txtCompany.TabIndex = 1
        Me.txtCompany.Text = ""
        '
        'txtCompanyID
        '
        Me.txtCompanyID.Location = New System.Drawing.Point(136, 8)
        Me.txtCompanyID.Name = "txtCompanyID"
        Me.txtCompanyID.ReadOnly = True
        Me.txtCompanyID.Size = New System.Drawing.Size(128, 20)
        Me.txtCompanyID.TabIndex = 0
        Me.txtCompanyID.Text = ""
        '
        'lblFinancialRespExpirationDate
        '
        Me.lblFinancialRespExpirationDate.Location = New System.Drawing.Point(5, 104)
        Me.lblFinancialRespExpirationDate.Name = "lblFinancialRespExpirationDate"
        Me.lblFinancialRespExpirationDate.Size = New System.Drawing.Size(128, 32)
        Me.lblFinancialRespExpirationDate.TabIndex = 26
        Me.lblFinancialRespExpirationDate.Text = "Financial Responsibility Expiration Date:"
        Me.lblFinancialRespExpirationDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCertORResponsibility
        '
        Me.lblCertORResponsibility.Location = New System.Drawing.Point(-11, 80)
        Me.lblCertORResponsibility.Name = "lblCertORResponsibility"
        Me.lblCertORResponsibility.Size = New System.Drawing.Size(144, 17)
        Me.lblCertORResponsibility.TabIndex = 24
        Me.lblCertORResponsibility.Text = "Certificate of Responsibility:"
        Me.lblCertORResponsibility.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCompanyID
        '
        Me.lblCompanyID.Location = New System.Drawing.Point(56, 8)
        Me.lblCompanyID.Name = "lblCompanyID"
        Me.lblCompanyID.Size = New System.Drawing.Size(77, 17)
        Me.lblCompanyID.TabIndex = 17
        Me.lblCompanyID.Text = "Company ID:"
        Me.lblCompanyID.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dtFinancialRespExpirationDate
        '
        Me.dtFinancialRespExpirationDate.Checked = False
        Me.dtFinancialRespExpirationDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtFinancialRespExpirationDate.Location = New System.Drawing.Point(136, 104)
        Me.dtFinancialRespExpirationDate.Name = "dtFinancialRespExpirationDate"
        Me.dtFinancialRespExpirationDate.ShowCheckBox = True
        Me.dtFinancialRespExpirationDate.Size = New System.Drawing.Size(104, 20)
        Me.dtFinancialRespExpirationDate.TabIndex = 4
        '
        'pnlCompanyHeader
        '
        Me.pnlCompanyHeader.Controls.Add(Me.lblCompanyHeader)
        Me.pnlCompanyHeader.Controls.Add(Me.lblCompanyDisplay)
        Me.pnlCompanyHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlCompanyHeader.Location = New System.Drawing.Point(0, 0)
        Me.pnlCompanyHeader.Name = "pnlCompanyHeader"
        Me.pnlCompanyHeader.Size = New System.Drawing.Size(976, 24)
        Me.pnlCompanyHeader.TabIndex = 208
        '
        'lblCompanyHeader
        '
        Me.lblCompanyHeader.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblCompanyHeader.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblCompanyHeader.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblCompanyHeader.Location = New System.Drawing.Point(16, 0)
        Me.lblCompanyHeader.Name = "lblCompanyHeader"
        Me.lblCompanyHeader.Size = New System.Drawing.Size(960, 24)
        Me.lblCompanyHeader.TabIndex = 1
        Me.lblCompanyHeader.Text = "Company"
        Me.lblCompanyHeader.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblCompanyDisplay
        '
        Me.lblCompanyDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblCompanyDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblCompanyDisplay.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompanyDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblCompanyDisplay.Name = "lblCompanyDisplay"
        Me.lblCompanyDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblCompanyDisplay.TabIndex = 0
        Me.lblCompanyDisplay.Text = "-"
        Me.lblCompanyDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Company
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(992, 753)
        Me.Controls.Add(Me.pnlCompanyContainer)
        Me.Controls.Add(Me.pnlCompanyBottom)
        Me.Name = "Company"
        Me.Text = "Company"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.pnlCompanyBottom.ResumeLayout(False)
        Me.pnlCompanyContainer.ResumeLayout(False)
        Me.pnlCompanyContactButtons.ResumeLayout(False)
        Me.pnlCompanyContactContainer.ResumeLayout(False)
        CType(Me.ugCompanyContacts, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlCompanyContactHeader.ResumeLayout(False)
        Me.pnlERACandContacts.ResumeLayout(False)
        Me.pnlDocumentDetails.ResumeLayout(False)
        Me.pnlDocumentsHead.ResumeLayout(False)
        Me.pnlPriorLicenseesDetails.ResumeLayout(False)
        CType(Me.ugPriorLicensees, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlPriorCompaniesHead.ResumeLayout(False)
        Me.pnlTypesDetails.ResumeLayout(False)
        Me.grpTypes.ResumeLayout(False)
        Me.pnlTypesHead.ResumeLayout(False)
        Me.pnlEracDetails.ResumeLayout(False)
        Me.grpERACs.ResumeLayout(False)
        Me.pnlEracHead.ResumeLayout(False)
        Me.pnlLicenseesDetails.ResumeLayout(False)
        CType(Me.ugLicensees, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlManagersDetails.ResumeLayout(False)
        CType(Me.ugManagers, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlLicenseeHead.ResumeLayout(False)
        Me.pnlManagerHead.ResumeLayout(False)
        Me.pnlCompanyDetails.ResumeLayout(False)
        Me.pnlCompanyHeader.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Form Events"
    Private Sub Company_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Try
            If pCompany Is Nothing Then
                pCompany = New MUSTER.BusinessLogic.pCompany
            End If

            UIUtilsGen.CreateEmptyFormatDatePicker(dtEngineerAppApprovalDate)
            UIUtilsGen.CreateEmptyFormatDatePicker(dtEngineerLiabilityExpiration)
            UIUtilsGen.CreateEmptyFormatDatePicker(dtFinancialRespExpirationDate)
            If nCompanyID > 0 Or nCompanyID < -100 Then
                companyInfo = pCompany.Retrieve(nCompanyID)
                populateCompanyInfo(companyInfo)
                pLic.GetAll(nCompanyID)
                pMgr.GetManagerAll(nCompanyID)
                pCompanyLicenseeAssociation.GetAll(nCompanyID)
                '   pCompanyManagerAssociation.GetAll(nCompanyID)
                PopulatePriorLicenseeGrid()
                btnComments.Enabled = True
                btnFlags.Enabled = True
                UCCompanyDocuments.LoadDocumentsGrid(nCompanyID, 0, 893)
                Me.EnableDisableControls(True)
            Else
                btnComments.Enabled = False
                btnFlags.Enabled = False
                Me.EnableDisableControls(False)
            End If

            'ugLicensees.DataSource = pLic.EntityTable().DefaultView
            populateLicenseesGrid()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)
        If pCompany.IsDirty() Then 'Or oComAdd.colIsDirty Or pLic.colIsDirty Then
            Dim Results As Long = MsgBox("There are unsaved changes.  Do you wish to save the changes?", MsgBoxStyle.Question & MsgBoxStyle.YesNoCancel, "Data Changed")
            If Results = MsgBoxResult.Yes Then
                '     If ugAddresses.Rows.Count = 0 Then
                '     MsgBox("Atleast one address should be entered for the company")
                '     e.Cancel = True
                '     Exit Sub
            End If

            Dim bolSuccess As Boolean = False
            pCompany.ModifiedBy = MusterContainer.AppUser.ID
            bolSuccess = pCompany.Save(CType(UIUtilsGen.ModuleID.Company, Integer), MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            If Not bolSuccess Then
                e.Cancel = True
                bolValidateSuccess = True
                bolDisplayErrmessage = True
                Exit Sub
            End If
        Else
        End If


        'if any other forms are using the company, leave alone. else remove from collection
        If pCompany.ID > 0 Or pCompany.ID < -100 Then
            UIUtilsGen.RemoveOwner(pCompany, Me)
        End If
    End Sub
    Private Sub Company_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.TextChanged
        MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "WindowName", Me.Text, "Company")
    End Sub
    Private Sub Company_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        ' The following line enables all forms to detect the visible form in the MDI container
        '
        MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "Company")
    End Sub
    Private Sub objLicen_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles objLicen.Closing
        Try
            'If pLic.IsDirty() Then
            If pLic.colIsDirty() Then
                Dim Results As Long = MsgBox("There are unsaved changes.  Do you wish to save the changes?", MsgBoxStyle.Question & MsgBoxStyle.YesNoCancel, "Data Changed")
                If Results = MsgBoxResult.Yes Then
                    'If Not pLic.Save() Then
                    If Not objLicen.SaveLicensee() Then
                        e.Cancel = True
                        bolValidateSuccess = True
                        bolDisplayErrmessage = True
                        Exit Sub
                    End If
                ElseIf Results = MsgBoxResult.Cancel Then
                    e.Cancel = True
                    Exit Sub
                ElseIf Results = MsgBoxResult.No Then
                    pLic.Reset()
                    Exit Sub
                End If
            End If
            If objLicen.bolFormClosing Then
                'ugLicensees.DataSource = pLic.EntityTable().DefaultView
                If pCompany.ID > 0 Or pCompany.ID < -100 Then
                    'If ugLicensees.Rows.Count <> 0 Then
                    pCompanyLicenseeAssociation.ModifiedBy = MusterContainer.AppUser.ID
                    pCompanyLicenseeAssociation.Flush(CType(UIUtilsGen.ModuleID.Company, Integer), MusterContainer.AppUser.UserKey, returnVal, pCompany.ID)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If
                    'End If
                End If
                populateLicenseesGrid()
                CompanyTypes()
            ElseIf nLicenseeID > 0 Or nLicenseeID < -100 Then
                pLic.Remove(nLicenseeID)
            ElseIf pLic.ID < 0 And pLic.ID > -100 Then
                pLic.Remove(pLic.ID)
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            objLicen.Dispose()
            objLicen = Nothing
            System.GC.Collect()
            System.GC.WaitForPendingFinalizers()
        End Try
    End Sub
#End Region

#Region "UI Support Routines"
    Private Sub EnableDisableControls(ByVal bolState As Boolean)
        Me.btnEnvelopes.Enabled = bolState
        Me.btnLabels.Enabled = bolState
        Me.btnLicenseesAddExisting.Enabled = bolState
        Me.btnLicenseesAddNew.Enabled = bolState
        Me.btnLicenseesEdit.Enabled = bolState
        Me.btnDisassociate.Enabled = bolState
    End Sub
    Private Sub CompanyTypes()
        Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim nLicensee As Integer = 0
        Dim nClosure As Integer = 0
        Try
            chkCertifiedtoInstallAlterandCloseUSTs.Checked = False
            chkCertifiedToCloseUSTs.Checked = False

            If ugLicensees.Rows.Count > 0 Then
                For Each ugRow In ugLicensees.Rows
                    If UCase(ugRow.Cells("STATUS").Value) = [Enum].GetName(GetType(EnumStatus), 0) Or UCase(ugRow.Cells("STATUS").Value) = [Enum].GetName(GetType(EnumStatus), 1) Then
                        If UCase(ugRow.Cells("Certification Type").Value) = [Enum].GetName(GetType(EnumCertificationType), 0) Then
                            chkCertifiedtoInstallAlterandCloseUSTs.Checked = True
                            nLicensee += 1
                        End If

                        If UCase(ugRow.Cells("Certification Type").Value) = [Enum].GetName(GetType(EnumCertificationType), 1) Then
                            nClosure += 1
                        End If
                    End If
                Next
            End If
            If nLicensee = 0 And nClosure > 0 Then
                chkCertifiedToCloseUSTs.Checked = True
            End If

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub setCompanyDelete(ByVal bolState As Boolean)
        Me.btnDeleteCompany.Enabled = bolState
    End Sub
    Private Sub setCompanySave(ByVal bolState As Boolean, Optional ByVal fromCompany As Boolean = True)
        btnSaveCompany.Enabled = bolState
        If fromCompany = False And bolState = False Then
            btnSaveCompany.Enabled = pCompany.IsDirty
        End If
    End Sub
    Private Function ValidateAddress() As Boolean
        If txtCompanyAddress.Text.Trim = "" Then
            MsgBox("Invalid Address. Please rectify!")
            Return False
        End If
        Return True
    End Function
    'Private Sub settypes(ByRef companyInfo As MUSTER.Info.CompanyInfo)
    '    Dim xlicenseeinfo As MUSTER.Info.LicenseeInfo
    '    Dim boltemp As Boolean = False
    '    Try
    '        For Each xlicenseeinfo In pLic.colLicensee.Values
    '            '1.	Set to True by the system for a Company if at least one Licensee assigned to the company (with Status = Certified or Renewing) has Certification Type = Install.  Otherwise set to False.
    '            ' need to change the code here
    '            If (xlicenseeinfo.CERT_TYPE = "INSTALL" And (xlicenseeinfo.STATUS = "CERTIFIED" Or xlicenseeinfo.STATUS = "RENEWING")) Then
    '                companyInfo.CTIAC = True
    '            End If
    '            '2.	Set CTC to True by the system for a Company if no Licensee in the Company has Certification Type = Install and at least one Licensee in the Company has  Certification Type = Closure.  Otherwise set to False.
    '            If xlicenseeinfo.CERT_TYPE <> "INSTALL" Then ' no licensee has certyType equal to install(1)
    '                boltemp = True
    '            End If
    '        Next
    '        For Each xlicenseeinfo In pLic.colLicensee.Values
    '            If boltemp And xlicenseeinfo.CERT_TYPE = "CLOSURE" Then ' atleast one licensee has certType as closure(0)
    '                companyInfo.CTC = True
    '            End If
    '        Next
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
#End Region

#Region "Company"
    Private Sub populateCompanyInfo(ByVal ocompInfo As MUSTER.Info.CompanyInfo)
        Dim oCompanyAddress As MUSTER.Info.ComAddressInfo
        Try
            bolLoading = True
            If ocompInfo.ID > 0 Or ocompInfo.ID < -100 Then
                Me.EnableDisableControls(True)
                CommentsMaintenance(, , True)
            Else
                Me.EnableDisableControls(False)
                ' #3112
                'CommentsMaintenance(, , True, True)
            End If
            txtCompanyID.Text = ocompInfo.ID
            txtCompany.Text = ocompInfo.COMPANY_NAME
            cmbCertORResponsibility.Text = ocompInfo.CERT_RESPON
            txtEmail.Text = ocompInfo.EMAIL_ADDRESS

            '------
            oCompanyAddress = oComAdd.GetAddressByType(0, nCompanyID, 0, 0)
            txtCompanyAddress.Text = oCompanyAddress.AddressLine1 + IIf(oCompanyAddress.AddressLine1.Trim.Length = 0, "", vbCrLf + oCompanyAddress.AddressLine2) + vbCrLf + oCompanyAddress.City + ", " + oCompanyAddress.State + " " + oCompanyAddress.Zip + vbCrLf + oCompanyAddress.Phone1
            nCompanyAddressID = oCompanyAddress.AddressId
            '------

            UIUtilsGen.SetDatePickerValue(dtFinancialRespExpirationDate, ocompInfo.FIN_RESP_END_DATE)
            txtProfEngineer.Text = ocompInfo.PRO_ENGIN

            txtEngineerNo.Text = ocompInfo.PRO_ENGIN_NUMBER
            TxtEngineerEmail.Text = ocompInfo.PRO_ENGIN_EMAIL

            UIUtilsGen.SetDatePickerValue(dtEngineerAppApprovalDate, ocompInfo.PRO_ENGIN_APP_APRV_DATE)
            UIUtilsGen.SetDatePickerValue(dtEngineerLiabilityExpiration, ocompInfo.PRO_ENGIN_LIABIL_DATE)
            txtProfGeologist.Text = ocompInfo.PRO_GEOLO
            txtGeologistNo.Text = ocompInfo.PRO_GEOLO_NUMBER
            TxtGeologistEmail.Text = ocompInfo.PRO_GEOLO_EMAIL

            chkCertifiedtoInstallAlterandCloseUSTs.Checked = ocompInfo.CTIAC
            chkCertifiedToCloseUSTs.Checked = ocompInfo.CTC
            chkPrecisionTankTightnessTesters.Checked = ocompInfo.PTTT
            chkLeakDetectionCompanies.Checked = ocompInfo.LDC
            chkUSTSuppliesAndEquipment.Checked = ocompInfo.USTSE
            chkWastewaterLaboratories.Checked = ocompInfo.WL
            chkThermalSoilTreatment.Checked = ocompInfo.TST
            chkERAC.Checked = ocompInfo.ERAC
            chkIRAC.Checked = ocompInfo.IRAC
            chkEnvironmentalConsultants.Checked = ocompInfo.EC
            chkEnvironmentalDrillers.Checked = ocompInfo.ED
            chkTankLining.Checked = ocompInfo.TL
            chkCorrosionExpert.Checked = ocompInfo.CE
            chkManager.Checked = ocompInfo.CM
            Me.Text = "Company" & " - Company Name - " & txtCompany.Text
            bolLoading = False

            'MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "CompanyID", ocompInfo.ID, Me.Name)
            MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "CompanyID", ocompInfo.ID, Me.Name)
            MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "CompanyName", ocompInfo.COMPANY_NAME, Me.Name)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub txtCompany_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCompany.TextChanged
        If Not bolLoading Then
            pCompany.COMPANY_NAME = txtCompany.Text
        End If
    End Sub
    Private Sub txtEmail_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEmail.TextChanged
        If Not bolLoading Then
            pCompany.EMAIL_ADDRESS = txtEmail.Text
        End If
    End Sub
    Private Sub cmbCertORResponsibility_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbCertORResponsibility.SelectedValueChanged
        If Not bolLoading Then
            pCompany.CERT_RESPON_ID = cmbCertORResponsibility.Text
        End If
    End Sub
    Private Sub dtFinancialRespExpirationDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtFinancialRespExpirationDate.ValueChanged
        UIUtilsGen.ToggleDateFormat(dtFinancialRespExpirationDate)
        UIUtilsGen.FillDateobjectValues(dateFinRespExp, dtFinancialRespExpirationDate.Text)
        pCompany.FIN_RESP_END_DATE = dtFinancialRespExpirationDate.Value
    End Sub

#Region "Address"
    Private Sub txtCompanyAddress_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCompanyAddress.DoubleClick
        Try
            If txtCompanyAddress.Text = "" Then
                objAddMaster = New AddressMaster(oComAdd, nCompanyAddressID, nCompanyID, "Company", "ADD")
            Else
                objAddMaster = New AddressMaster(oComAdd, nCompanyAddressID, nCompanyID, "Company", "MODIFY")
            End If
            'Me.Update()
            'LockWindowUpdate(Me.Handle.ToInt64)
            objAddMaster.ShowDialog()
            'LockWindowUpdate(0)
            nCompanyAddressID = oComAdd.AddressId
            Me.txtCompanyAddress.Text = oComAdd.AddressLine1 + IIf(oComAdd.AddressLine1.Trim.Length = 0, "", vbCrLf + oComAdd.AddressLine2) + IIf(oComAdd.City.Length = 0, "", vbCrLf + oComAdd.City) + IIf(oComAdd.State.Length = 0, "", ", " + oComAdd.State) + IIf(oComAdd.Zip.Length = 0, "", " " + oComAdd.Zip) + vbCrLf + oComAdd.Phone1
        Catch ex As Exception
            ' ShowError(ex)
        End Try
    End Sub
    'Private Sub txtCompanyAddress_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCompanyAddress.TextChanged
    '    If bolLoading Then Exit Sub
    '    If txtCompanyAddress.Tag > 0 Then
    '        '     pOwn.AddressId = Integer.Parse(Trim(txtOwnerAddress.Tag))
    '    Else
    '        '    pOwn.AddressId = 0
    '    End If
    'End Sub

    Private Sub btnEnvelopes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEnvelopes.Click

        Dim oCompanyAddress As MUSTER.Info.ComAddressInfo
        Dim arrAddress(4) As String

        Try
            '------
            oCompanyAddress = oComAdd.GetAddressByType(0, nCompanyID, 0, 0)
            '------
            arrAddress(0) = oCompanyAddress.AddressLine1
            arrAddress(1) = oCompanyAddress.AddressLine2
            arrAddress(2) = oCompanyAddress.City
            arrAddress(3) = oCompanyAddress.State
            arrAddress(4) = oCompanyAddress.Zip

            If Not oCompanyAddress.AddressId = 0 And (pCompany.ID > 0 Or pCompany.ID < -100) Then
                UIUtilsGen.CreateEnvelopes(Me.txtCompany.Text, arrAddress, "COM", pCompany.ID)
            Else
                MsgBox("Invalid Address")
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnLabels_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLabels.Click
        Dim oCompanyAddress As MUSTER.Info.ComAddressInfo
        Dim arrAddress(4) As String
        Try

            '------
            oCompanyAddress = oComAdd.GetAddressByType(0, nCompanyID, 0, 0)
            '------
            arrAddress(0) = oCompanyAddress.AddressLine1
            arrAddress(1) = oCompanyAddress.AddressLine2
            arrAddress(2) = oCompanyAddress.City
            arrAddress(3) = oCompanyAddress.State
            arrAddress(4) = oCompanyAddress.Zip

            If Not oCompanyAddress.AddressId = 0 And (pCompany.ID > 0 Or pCompany.ID < -100) Then
                UIUtilsGen.CreateLabels(Me.txtCompany.Text, arrAddress, "COM", pCompany.ID)
            Else
                MsgBox("Invalid Address")
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#End Region

#Region "Licensee"
    Private Sub populateLicenseesGrid()
        Try
            If pCompany.ID > 0 Or pCompany.ID < -100 Then
                pLic.GetAll(pCompany.ID)
                pMgr.GetManagerAll(pCompany.ID)
            End If
            ugLicensees.DataSource = pLic.EntityTable().DefaultView
            ugLicensees.DisplayLayout.Bands(0).Columns("Status_ID").Hidden = True
            ugLicensees.DisplayLayout.Bands(0).Columns("CERT_TYPE_ID").Hidden = True
            ugLicensees.DisplayLayout.Bands(0).Columns("Status").Width = 160
            CompanyTypes()
            'If pLic.EntityTable.Rows.Count > 0 Then
            '    setCompanyDelete(False)
            'Else
            '    setCompanyDelete(True)
            'End If
            If ugLicensees.Rows.Count > 0 Then
                setCompanyDelete(False)
            Else
                setCompanyDelete(True)
            End If
            'Populate manager grid
            ugManagers.DataSource = pMgr.EntityTableManager().DefaultView
            ugManagers.DisplayLayout.Bands(0).Columns("CMStatus_ID").Hidden = True
            ugManagers.DisplayLayout.Bands(0).Columns("CERT_TYPE_ID").Hidden = True
            ugManagers.DisplayLayout.Bands(0).Columns("Manager Number").Hidden = True
            ugManagers.DisplayLayout.Bands(0).Columns("Hire Status").Hidden = True
            ugManagers.DisplayLayout.Bands(0).Columns("CMStatus").Width = 160
            ugManagers.DisplayLayout.Bands(0).Columns("Certification Type").Width = 190
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub PopulateLicensees()
        Dim nLicenseeID As Integer = 0
        Dim oComLicInfo As MUSTER.Info.CompanyLicenseeInfo
        Dim oLicCompAddInfo As MUSTER.Info.ComAddressInfo
        Dim nAssociationID As Integer = 0
        Try
            'setCompanySave(True)
            If Not (ugLicensees.ActiveRow Is Nothing) Then
                nLicenseeID = ugLicensees.ActiveRow.Cells("Licensee ID").Value
            Else
                Exit Sub
            End If
            Licensees(nLicenseeID)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub PopulateManagers()
        Dim nManagerID As Integer = 0
        Dim oComLicInfo As MUSTER.Info.CompanyLicenseeInfo
        Dim oLicCompAddInfo As MUSTER.Info.ComAddressInfo
        Dim nAssociationID As Integer = 0
        Try
            'setCompanySave(True)
            If Not (ugManagers.ActiveRow Is Nothing) Then
                nManagerID = ugManagers.ActiveRow.Cells("Manager ID").Value
            Else
                Exit Sub
            End If
            Managers(nManagerID)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Friend Sub Licensees(ByVal nLicenseeID As Integer)
        Dim oComLicInfo As MUSTER.Info.CompanyLicenseeInfo
        Dim oLicCompAddInfo As MUSTER.Info.ComAddressInfo
        Dim strAddress As String = ""
        Dim nAssociationID As Integer = 0
        Try
            For Each oComLicInfo In pCompanyLicenseeAssociation.ComLicCollection.Values
                If oComLicInfo.LicenseeID = nLicenseeID And oComLicInfo.CompanyID = nCompanyID Then
                    nAssociationID = oComLicInfo.ID
                    If nAssociationID > 0 Or nAssociationID < -100 Then
                        oLicCompAddInfo = oComAdd.Retrieve(oComLicInfo.ComLicAddressID, False)
                    Else
                        oLicCompAddInfo = oComAdd.Retrieve(pCompanyLicenseeAssociation.ComLicAddressID, False)
                    End If
                    strAddress = oLicCompAddInfo.AddressLine1 + IIf(oLicCompAddInfo.AddressLine1.Trim.Length = 0, "", vbCrLf + oLicCompAddInfo.AddressLine2) + vbCrLf + oLicCompAddInfo.City + ", " + oLicCompAddInfo.State + " " + oLicCompAddInfo.Zip
                    Exit For
                End If
            Next
            objLicen = New Licensees(pLic, pCompanyLicenseeAssociation, oComAdd, nCompanyID, strAddress, oLicCompAddInfo.AddressId, nLicenseeID, nAssociationID, pCompany.FIN_RESP_END_DATE, txtCompany.Text, True)
            'objLicen.MdiParent = Me
            objLicen.callingForm = Me
            Me.Tag = "0"
            objLicen.ShowDialog()
            If Me.Tag = "1" Then
                ' update licensees grid
                populateLicenseesGrid()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Friend Sub Managers(ByVal nManagerID As Integer)
        Dim oComLicInfo As MUSTER.Info.CompanyLicenseeInfo
        Dim oLicCompAddInfo As MUSTER.Info.ComAddressInfo
        Dim strAddress As String = ""
        Dim nAssociationID As Integer = 0
        Try
            For Each oComLicInfo In pCompanyLicenseeAssociation.ComLicCollection.Values
                If oComLicInfo.LicenseeID = nManagerID And oComLicInfo.CompanyID = nCompanyID Then
                    nAssociationID = oComLicInfo.ID
                    If nAssociationID > 0 Or nAssociationID < -100 Then
                        oLicCompAddInfo = oComAdd.Retrieve(oComLicInfo.ComLicAddressID, False)
                    Else
                        oLicCompAddInfo = oComAdd.Retrieve(pCompanyLicenseeAssociation.ComLicAddressID, False)
                    End If
                    strAddress = oLicCompAddInfo.AddressLine1 + IIf(oLicCompAddInfo.AddressLine1.Trim.Length = 0, "", vbCrLf + oLicCompAddInfo.AddressLine2) + vbCrLf + oLicCompAddInfo.City + ", " + oLicCompAddInfo.State + " " + oLicCompAddInfo.Zip
                    Exit For
                End If
            Next
            objMgr = New Managers(pMgr, pCompanyLicenseeAssociation, oComAdd, nCompanyID, strAddress, oLicCompAddInfo.AddressId, nManagerID, nAssociationID, pCompany.FIN_RESP_END_DATE, txtCompany.Text, True)
            'objLicen.MdiParent = Me
            objMgr.callingForm = Me
            Me.Tag = "0"
            objMgr.ShowDialog()
            If Me.Tag = "1" Then
                ' update licensees grid
                populateLicenseesGrid()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugLicensees_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugLicensees.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            PopulateLicensees()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugManagers_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugManagers.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            PopulateManagers()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnLicenseesAddExisting_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLicenseesAddExisting.Click
        Try
            'setCompanySave(True)
            oLicenseeList = New LicenseesList(pCompany, pLic)
            'oLicenseeList.MdiParent = Me
            oLicenseeList.ShowDialog()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnLicenseesAddNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLicenseesAddNew.Click
        Dim msg As String = ""
        Try
            'setCompanySave(True)
            If txtCompany.Text = "" Then
                msg += "Please enter the company name before proceeding" + vbCrLf
            End If
            '   If ugAddresses.Rows.Count = 0 Then
            '   msg += "Please enter the company address before proceeding"
            '  End If
            If msg <> "" Then
                MsgBox(msg)
                Exit Sub
            End If
            pLic.pLicenseeCourse = New MUSTER.BusinessLogic.pLicenseeCourses
            pLic.pLicenseeCourseTest = New MUSTER.BusinessLogic.pLicenseeCourseTest

            Dim oLicCompAddInfo As MUSTER.Info.ComAddressInfo
            oLicCompAddInfo = oComAdd.Retrieve(nCompanyAddressID, False)
            msg = oLicCompAddInfo.AddressLine1 + IIf(oLicCompAddInfo.AddressLine1.Trim.Length = 0, "", vbCrLf + oLicCompAddInfo.AddressLine2) + vbCrLf + oLicCompAddInfo.City + ", " + oLicCompAddInfo.State + " " + oLicCompAddInfo.Zip
            objLicen = New Licensees(pLic, pCompanyLicenseeAssociation, oComAdd, nCompanyID, msg, oLicCompAddInfo.AddressId, 0, 0, pCompany.FIN_RESP_END_DATE, txtCompany.Text, True, "ADD")
            'objLicen.MdiParent = Me
            objLicen.ShowDialog()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnLicenseesEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLicenseesEdit.Click
        Try
            PopulateLicensees()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnDisassociate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDisassociate.Click
        Dim oCompanyLicenseeInfoLocal As MUSTER.Info.CompanyLicenseeInfo
        Dim result As DialogResult
        Dim resultTemp As DialogResult
        Try
            If ugLicensees.ActiveRow Is Nothing Then
                Exit Sub
            End If
            'setCompanySave(True)
            result = MessageBox.Show("Do you want to Disassociate?", "Disassociation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)
            If result = DialogResult.Yes Then
                For Each oCompanyLicenseeInfoLocal In pCompanyLicenseeAssociation.ComLicCollection.Values
                    If oCompanyLicenseeInfoLocal.LicenseeID = ugLicensees.ActiveRow.Cells("licensee id").Value Then
                        oCompanyLicenseeInfoLocal.Deleted = True
                        Exit For
                    End If
                Next
                pCompanyLicenseeAssociation.ModifiedBy = MusterContainer.AppUser.ID
                pCompanyLicenseeAssociation.Flush(CType(UIUtilsGen.ModuleID.Company, Integer), MusterContainer.AppUser.UserKey, returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                pLic.Retrieve(ugLicensees.ActiveRow.Cells("licensee ID").Value)
                pLic.EMPLOYEE_LETTER = False
                pLic.CERT_TYPE_ID = 0
                pLic.CERT_TYPE_DESC = String.Empty
                pLic.HIRE_STATUS = String.Empty
                'resultTemp = MessageBox.Show("Do you want to associate the Licensee to another Company?", "Association", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)
                'If resultTemp = DialogResult.No Then
                'pLic.STATUS_ID = "NO LONGER WITH COMPANY"
                'End If
                If pLic.ID <= 0 Then
                    pLic.CreatedBy = MusterContainer.AppUser.ID
                Else
                    pLic.ModifiedBy = MusterContainer.AppUser.ID
                End If
                pLic.Save(CType(UIUtilsGen.ModuleID.Company, Integer), MusterContainer.AppUser.UserKey, returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                pLic.Remove(ugLicensees.ActiveRow.Cells("licensee ID").Value)
                'ugLicensees.DataSource = pLic.EntityTable().DefaultView
                populateLicenseesGrid()
                PopulatePriorLicenseeGrid()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub oLicenseeList_CompanyLicenseeAssociation(ByVal LicenseeID As Integer) Handles oLicenseeList.CompanyLicenseeAssociation
        Try
            nLicenseeID = LicenseeID
            Dim msg As String = String.Empty
            Dim oLicCompAddInfo As MUSTER.Info.ComAddressInfo
            oLicCompAddInfo = oComAdd.Retrieve(nCompanyAddressID, False)
            msg = oLicCompAddInfo.AddressLine1 + IIf(oLicCompAddInfo.AddressLine1.Trim.Length = 0, "", vbCrLf + oLicCompAddInfo.AddressLine2) + vbCrLf + oLicCompAddInfo.City + ", " + oLicCompAddInfo.State + " " + oLicCompAddInfo.Zip
            objLicen = New Licensees(pLic, pCompanyLicenseeAssociation, oComAdd, nCompanyID, msg, nCompanyAddressID, LicenseeID, 0, pCompany.FIN_RESP_END_DATE, txtCompany.Text, True)
            'objLicen.MdiParent = Me
            Me.Update()
            'LockWindowUpdate(Me.Handle.ToInt64)
            'objLicen.MdiParent = Me
            objLicen.ShowDialog()
            'LockWindowUpdate(0)
            nLicenseeID = 0
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub oManagerList_CompanyManagerAssociation(ByVal ManagerID As Integer) Handles oManagerList.CompanyManagerAssociation
        Try
            nManagerID = ManagerID
            Dim msg As String = String.Empty
            Dim oLicCompAddInfo As MUSTER.Info.ComAddressInfo
            oLicCompAddInfo = oComAdd.Retrieve(nCompanyAddressID, False)
            msg = oLicCompAddInfo.AddressLine1 + IIf(oLicCompAddInfo.AddressLine1.Trim.Length = 0, "", vbCrLf + oLicCompAddInfo.AddressLine2) + vbCrLf + oLicCompAddInfo.City + ", " + oLicCompAddInfo.State + " " + oLicCompAddInfo.Zip
            objMgr = New Managers(pLic, pCompanyLicenseeAssociation, oComAdd, nCompanyID, msg, nCompanyAddressID, ManagerID, 0, pCompany.FIN_RESP_END_DATE, txtCompany.Text, True)
            'objLicen.MdiParent = Me
            Me.Update()
            'LockWindowUpdate(Me.Handle.ToInt64)
            'objLicen.MdiParent = Me
            objMgr.btnSaveManager.Enabled = True
            objMgr.ShowDialog()
            'LockWindowUpdate(0)
            nManagerID = 0
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "EARC"
    Private Sub dtEngineerAppApprovalDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtEngineerAppApprovalDate.ValueChanged
        UIUtilsGen.ToggleDateFormat(dtEngineerAppApprovalDate)
        UIUtilsGen.FillDateobjectValues(dateEngAppApproval, dtEngineerAppApprovalDate.Text)
        pCompany.PRO_ENGIN_APP_APRV_DATE = dtEngineerAppApprovalDate.Value
    End Sub
    Private Sub dtEngineerLiabilityExpiration_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtEngineerLiabilityExpiration.ValueChanged
        UIUtilsGen.ToggleDateFormat(dtEngineerLiabilityExpiration)
        UIUtilsGen.FillDateobjectValues(dateEngLiabExpiration, dtEngineerLiabilityExpiration.Text)
        pCompany.PRO_ENGIN_LIABIL_DATE = dtEngineerLiabilityExpiration.Value
    End Sub
    Private Sub txtProfEngineer_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProfEngineer.TextChanged
        If Not bolLoading Then
            pCompany.PRO_ENGIN = txtProfEngineer.Text
        End If
    End Sub
    Private Sub txtEngineerNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEngineerNo.TextChanged
        If Not bolLoading Then
            pCompany.PRO_ENGIN_NUMBER = txtEngineerNo.Text
        End If
    End Sub
    Private Sub TxtEngineerEmail_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtEngineerEmail.TextChanged
        If Not bolLoading Then
            pCompany.PRO_ENGIN_EMAIL = TxtEngineerEmail.Text
        End If
    End Sub
    Private Sub txtProfGeologist_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtProfGeologist.TextChanged
        If Not bolLoading Then
            pCompany.PRO_GEOLO = txtProfGeologist.Text
        End If
    End Sub
    Private Sub txtGeologistNo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtGeologistNo.TextChanged
        If Not bolLoading Then
            pCompany.PRO_GEOLO_NUMBER = txtGeologistNo.Text
        End If
    End Sub
    Private Sub txtGeologistEmail_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TxtGeologistEmail.TextChanged
        If Not bolLoading Then
            pCompany.PRO_GEOLO_EMAIL = TxtGeologistEmail.Text
        End If
    End Sub
#End Region

#Region "Types"
    Private Sub chkCertifiedtoInstallAlterandCloseUSTs_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkCertifiedtoInstallAlterandCloseUSTs.CheckedChanged
        If Not bolLoading Then
            pCompany.CTIAC = chkCertifiedtoInstallAlterandCloseUSTs.Checked
        End If
    End Sub
    Private Sub chkCertifiedToCloseUSTs_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkCertifiedToCloseUSTs.CheckedChanged
        If Not bolLoading Then
            pCompany.CTC = chkCertifiedToCloseUSTs.Checked
        End If
    End Sub
    Private Sub chkPrecisionTankTightnessTesters_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkPrecisionTankTightnessTesters.CheckedChanged
        If Not bolLoading Then
            pCompany.PTTT = chkPrecisionTankTightnessTesters.Checked
        End If
    End Sub
    Private Sub chkLeakDetectionCompanies_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkLeakDetectionCompanies.CheckedChanged
        If Not bolLoading Then
            pCompany.LDC = chkLeakDetectionCompanies.Checked
        End If
    End Sub
    Private Sub chkUSTSuppliesAndEquipment_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkUSTSuppliesAndEquipment.CheckedChanged
        If Not bolLoading Then
            pCompany.USTSE = chkUSTSuppliesAndEquipment.Checked
        End If
    End Sub
    Private Sub chkWastewaterLaboratories_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkWastewaterLaboratories.CheckedChanged
        If Not bolLoading Then
            pCompany.WL = chkWastewaterLaboratories.Checked
        End If
    End Sub
    Private Sub chkThermalSoilTreatment_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkThermalSoilTreatment.CheckedChanged
        If Not bolLoading Then
            pCompany.TST = chkThermalSoilTreatment.Checked
        End If
    End Sub
    Private Sub chkERAC_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkERAC.CheckedChanged
        If Not bolLoading Then
            pCompany.ERAC = chkERAC.Checked
        End If
    End Sub
    Private Sub chkERAC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkERAC.Click
        Dim msg As String = ""
        Dim bolTemp As Boolean = False
        Try
            If chkERAC.Checked Then
                If (txtEngineerNo.Text = "" Or txtProfEngineer.Text = "") And (txtGeologistNo.Text = "" Or txtProfGeologist.Text = "") Then
                    msg += "All information for the professional engineer and/or a professional geologist has to be specified" + vbCrLf
                End If
                If Not (dtEngineerLiabilityExpiration.Value.Date > Now.Date) Then
                    msg += "Liablity expiration date should be greater than today" + vbCrLf
                End If
                If DateDiff(DateInterval.Month, dtEngineerAppApprovalDate.Value, Now.Date) > 12 Then
                    msg += "ERAC Application Approval Date should be < 12 months"
                End If
                If msg <> "" Then
                    MsgBox(msg)
                    chkERAC.Checked = False
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub chkIRAC_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkIRAC.CheckedChanged
        If Not bolLoading Then
            pCompany.IRAC = chkIRAC.Checked
        End If
    End Sub
    Private Sub chkIRAC_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkIRAC.Click
        Dim xlicenseeinfo As MUSTER.Info.LicenseeInfo
        Dim bolIRAC As Boolean = False
        Try
            If chkIRAC.Checked Then
                For Each xlicenseeinfo In pLic.colLicensee.Values
                    ' Updated on Oct 3 by Kumar
                    If ((xlicenseeinfo.CertTypeDesc = "CLOSURE" Or xlicenseeinfo.CertTypeDesc = "INSTALL") And (xlicenseeinfo.STATUS = "CERTIFIED" Or xlicenseeinfo.STATUS = "RENEWING")) Then
                        bolIRAC = True
                    End If
                Next
                If Not bolIRAC Then
                    chkIRAC.Checked = False
                    MsgBox("Company should contain at least one licensee associated with status = Certified or Renewing that has Certification type = Closure")
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub chkEnvironmentalConsultants_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkEnvironmentalConsultants.CheckedChanged
        If Not bolLoading Then
            pCompany.EC = chkEnvironmentalConsultants.Checked
        End If
    End Sub
    Private Sub chkEnvironmentalDrillers_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkEnvironmentalDrillers.CheckedChanged
        If Not bolLoading Then
            pCompany.ED = chkEnvironmentalDrillers.Checked
        End If
    End Sub
    Private Sub chkTankLining_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkTankLining.CheckedChanged
        If Not bolLoading Then
            pCompany.TL = chkTankLining.Checked
        End If
    End Sub
    Private Sub chkCorrosionExpert_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkCorrosionExpert.CheckedChanged
        If Not bolLoading Then
            pCompany.CE = chkCorrosionExpert.Checked
        End If
    End Sub
   Private Sub chkManager_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkManager.CheckedChanged
        If Not bolLoading Then
            pCompany.CM= chkManager.Checked
        End If
    End Sub
#End Region

#Region "Prior Licensees"
    Private Sub PopulatePriorLicenseeGrid()
        Try

            ugPriorLicensees.DataSource = pCompany.GetAssociatedLicensees(nCompanyID, True).DefaultView

            ugPriorLicensees.DisplayLayout.Bands(0).Columns("COM_LICENSEE_ID").Hidden = True
            ugPriorLicensees.DisplayLayout.Bands(0).Columns("Licensee ID").Width = 30
            ugPriorLicensees.DisplayLayout.Bands(0).Columns("Licensee Name").Width = 200
            ugPriorLicensees.DisplayLayout.Bands(0).Columns("Licensee Number").Width = 90
            ugPriorLicensees.DisplayLayout.Bands(0).Columns("Status").Width = 160
            ugPriorLicensees.DisplayLayout.Bands(0).Columns("Certification Type").Width = 80
            ugPriorLicensees.DisplayLayout.Bands(0).Columns("Exp. Date").Width = 80
            ugPriorLicensees.DisplayLayout.Bands(0).Columns("Exp. Grant Date").Width = 80

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Contacts"
#Region "Button and Change Events"
    Dim strFilterString As String = String.Empty
    Private Sub SetCompanyFilter()
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
            strFilterString = String.Empty
            If chkCompanyShowActive.Checked Then
                bolActive = True
            Else
                bolActive = False
            End If

            If chkCompanyShowContactsforAllModules.Checked Then
                ' User has the ability to view the contacts associated for the entity in other modules
                ' Dim strFilterForAllModules As String = pConStruct.GetContactsForAllModules(oOwner.Facility.ID.ToString)
                nEntityID = nCompanyID
                nEntityType = UIUtilsGen.EntityTypes.Company
                '  strEntityAssocIDs = strFilterForAllModules
                nModuleID = 893
            Else
                nEntityID = nCompanyID
                nEntityType = UIUtilsGen.EntityTypes.Company
                nModuleID = 893
            End If

            If chkCompanyShowRelatedContacts.Checked Then
                '  strEntities = 
                '  nRelatedEntityType = 
            End If

            Me.LoadContacts(ugCompanyContacts, nEntityID, nEntityType)


            strFilterString = String.Empty
            If chkCompanyShowActive.Checked Then
                strFilterString = "(ACTIVE = 1"
            Else
                strFilterString = "("
            End If

            'If chkCompanyShowContactsforAllModules.Checked Then
            '    ' User has the ability to view the contacts associated for the entity in other modules
            '    'User has the ability to view the contacts associated for the entity in other modules
            '    Dim strFilterForAllModules As String = pConStruct.GetContactsForAllModules(oOwner.Facility.ID.ToString)
            '    If strFilterString = "(" Then
            '        strFilterString += "ENTITYID = " + oOwner.Facility.ID.ToString + " OR " + IIf(Not strFilterForAllModules = String.Empty, " entityassocid in (" + strFilterForAllModules + ")", "")
            '    Else
            '        strFilterString += "AND ENTITYID = " + oOwner.Facility.ID.ToString + " OR " + IIf(Not strFilterForAllModules = String.Empty, " entityassocid in (" + strFilterForAllModules + ")", "")
            '    End If
            '    'Else
            '    '    If strFilterString = "(" Then
            '    '        strFilterString += "ENTITYID = " + strEntityID
            '    '    Else
            '    '        strFilterString += "AND ENTITYID = " + strEntityID
            '    '    End If
            '    'End If
            '    'If strFilterString = "(" Then
            '    '    strFilterString += "ENTITYID = " + oOwner.Facility.ID.ToString + " OR ENTITYID = " + oLustEvent.ID.ToString
            '    'Else
            '    '    strFilterString += "AND ENTITYID = " + oOwner.Facility.ID.ToString + " OR ENTITYID = " + oLustEvent.ID.ToString
            '    'End If
            'Else
            '    If strFilterString = "(" Then
            '        strFilterString += " MODULEID = 614 And ENTITYID = " + oLustEvent.ID.ToString
            '    Else
            '        strFilterString += " AND MODULEID = 614 And ENTITYID = " + oLustEvent.ID.ToString
            '    End If
            'End If

            'If chkCompanyShowRelatedContacts.Checked Then
            'strFilterString += " OR " + IIf(Not strLustEventIdTags = String.Empty, " ENTITYID in (" + strLustEventIdTags + "))", "")
            'Else
            '    strFilterString += ")"
            'End If

            'dsContacts.Tables(0).DefaultView.RowFilter = strFilterString
            'ugCompanyContacts.DataSource = dsContacts.Tables(0).DefaultView

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCompanyContactAddorSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCompanyContactAddorSearch.Click
        Try
            pConStruct = New MUSTER.BusinessLogic.pContactStruct
            objCntSearch = New ContactSearch(nCompanyID, 26, "Company", pConStruct)

            objCntSearch.ShowDialog()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub
    Private Sub btnCompanyContactModify_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCompanyContactModify.Click
        Try
            ModifyContact(ugCompanyContacts)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCompanyContactDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCompanyContactDelete.Click
        Try
            DeleteContact(ugCompanyContacts, nCompanyID)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCompanyContactAssociate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCompanyContactAssociate.Click
        Try
            '  AssociateContact(ugCompanyContacts, oLustEvent.ID, 7)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()

        End Try
    End Sub
#End Region
#Region "Company Contacts"
    Private Sub chkCompanyShowContactsforAllModules_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCompanyShowContactsforAllModules.CheckedChanged
        Try
            'If Not dsContacts Is Nothing Then
            SetCompanyFilter()
            'End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub chkCompanyShowRelatedContacts_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCompanyShowRelatedContacts.CheckedChanged
        Try
            If Not dsContacts Is Nothing Then
                SetCompanyFilter()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub chkCompanyShowActive_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkCompanyShowActive.CheckedChanged
        Try
            If Not dsContacts Is Nothing Then
                SetCompanyFilter()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

#End Region
#Region "Common Functions"

    Private Sub Contact_ContactAdded(ByVal uggrid As Infragistics.Win.UltraWinGrid.UltraGrid)

        LoadContacts(uggrid, nCompanyID, UIUtilsGen.EntityTypes.Company)

        chkCompanyShowContactsforAllModules.Checked = False
    End Sub

    Private Sub LoadContacts(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal EntityID As Integer, ByVal EntityType As Integer)
        Try

            UIUtilsGen.LoadContacts(ugGrid, EntityID, EntityType, pConStruct, 893)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()

        End Try
    End Sub
    Private Function ModifyContact(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid)
        Try

            If UIUtilsGen.ModifyContact(ugGrid, 893, pConStruct) Then
                Me.Contact_ContactAdded(ugGrid)
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function

    Private Function AssociateContact(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal nEntityID As Integer, ByVal nEntityType As Integer)
        Try
            If UIUtilsGen.AssociateContact(ugGrid, nEntityID, nEntityType, 893, pConStruct) Then
                Me.Contact_ContactAdded(ugGrid)
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function

    Private Function DeleteContact(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal nEntityID As Integer)
        Try

            If UIUtilsGen.DeleteContact(ugGrid, nEntityID, 893, pConStruct) Then
                Me.Contact_ContactAdded(ugGrid)
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function

#End Region
#Region "Close Events"
    Private Sub objCntSearch_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles objCntSearch.Closing
        If Not objCntSearch Is Nothing Then
            objCntSearch = Nothing
        End If
    End Sub
#End Region

#End Region

#Region "UI Events"
    Private Sub btnSaveCompany_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveCompany.Click
        Try
            If Not ValidateAddress() Then Exit Sub

            companyInfo = New MUSTER.Info.CompanyInfo(nCompanyID, _
            cmbCertORResponsibility.Text, _
            txtCompany.Text, _
            dateFinRespExp, _
            txtEmail.Text, _
            txtProfEngineer.Text, _
            pCompany.PRO_ENGIN_ADD_ID, _
            txtEngineerNo.Text, _
            dateEngAppApproval, _
            dateEngLiabExpiration, _
            txtProfGeologist.Text, _
            pCompany.PRO_GEOLO_ADD_ID, _
            txtGeologistNo.Text, _
            chkCertifiedtoInstallAlterandCloseUSTs.Checked, _
            chkCertifiedToCloseUSTs.Checked, _
            chkPrecisionTankTightnessTesters.Checked, _
            chkLeakDetectionCompanies.Checked, _
            chkUSTSuppliesAndEquipment.Checked, _
            chkWastewaterLaboratories.Checked, _
            chkThermalSoilTreatment.Checked, _
            chkERAC.Checked, _
            chkIRAC.Checked, _
            chkEnvironmentalConsultants.Checked, _
            chkEnvironmentalDrillers.Checked, _
            chkTankLining.Checked, _
            chkCorrosionExpert.Checked, _
            chkManager.Checked, _
            True, _
            False, _
            IIf((nCompanyID <= 0 And nCompanyID > -100), MusterContainer.AppUser.ID, ""), _
            Now, _
            IIf(nCompanyID > 0 Or nCompanyID < -100, MusterContainer.AppUser.ID, ""), _
            CDate("01/01/0001"), _
            TxtEngineerEmail.Text, _
            TxtGeologistEmail.Text)
            'Me.settypes(companyInfo)
            If Not pCompany.Add(companyInfo) Then
                Exit Sub
            End If

            Dim bolSuccess As Boolean = False
            bolSuccess = pCompany.Save(CType(UIUtilsGen.ModuleID.Company, Integer), MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If


            If bolSuccess Then
                nCompanyID = pCompany.ID
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                If pCompany.ID <= 0 And pCompany.ID > -99 Then ' there are 30 companies with a negative ID
                    pCompany.CREATED_BY = MusterContainer.AppUser.ID
                Else
                    pCompany.ModifiedBy = MusterContainer.AppUser.ID
                End If
                'Call the save again to update created by and modified by 
                pCompany.Save(CType(UIUtilsGen.ModuleID.Company, Integer), MusterContainer.AppUser.UserKey, returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                ' save the Company address
                For Each oComAddInfo As MUSTER.Info.ComAddressInfo In oComAdd.ColCompanyAddresses.Values
                    If oComAddInfo.AddressId = nCompanyAddressID Then
                        ' begin issue 3199
                        oComAdd.Retrieve(oComAddInfo.AddressId, False)
                        ' end issue 3199
                        oComAdd.CompanyId = nCompanyID
                        oComAdd.Save(UIUtilsGen.ModuleID.Company, MusterContainer.AppUser.UserKey, returnVal, True)
                        If Not UIUtilsGen.HasRights(returnVal) Then
                            Exit Sub
                        End If
                        Exit For ' exiting for as the business logic changed for company to have only one address
                    End If
                Next

                'Save the Company-licensees association into the DB
                If ugLicensees.Rows.Count <> 0 Then
                    pCompany.ModifiedBy = MusterContainer.AppUser.ID
                    pCompanyLicenseeAssociation.Flush(CType(UIUtilsGen.ModuleID.Company, Integer), MusterContainer.AppUser.UserKey, returnVal, pCompany.ID)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If
                End If
                MsgBox("Company is saved successfully")
                txtCompanyID.Text = pCompany.ID.ToString
                btnComments.Enabled = True
                btnFlags.Enabled = True
                PopulatePriorLicenseeGrid()

                If pCompany.ID > 0 Or pCompany.ID < -100 Then
                    Me.EnableDisableControls(True)
                Else
                    Me.EnableDisableControls(False)
                End If
                UCCompanyDocuments.LoadDocumentsGrid(pCompany.ID, 0, UIUtilsGen.ModuleID.Company)

            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub
    Private Sub btnDeleteCompany_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteCompany.Click
        Dim result As DialogResult
        Dim xComAssocInfo As MUSTER.Info.CompanyLicenseeInfo
        Dim xComAddInfo As MUSTER.Info.ComAddressInfo
        Try
            result = MessageBox.Show("Are you sure you want to delete the company", "Delete Company", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)
            If result = DialogResult.Yes Then
                pCompany.Retrieve(nCompanyID)
                pCompany.DELETED = True
                pCompany.ModifiedBy = MusterContainer.AppUser.ID
                pCompany.Save(CType(UIUtilsGen.ModuleID.Company, Integer), MusterContainer.AppUser.UserKey, returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                For Each xComAssocInfo In pCompanyLicenseeAssociation.ComLicCollection.Values
                    xComAssocInfo.Deleted = True
                Next
                pCompanyLicenseeAssociation.ModifiedBy = MusterContainer.AppUser.ID
                pCompanyLicenseeAssociation.Flush(CType(UIUtilsGen.ModuleID.Company, Integer), MusterContainer.AppUser.UserKey, returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                MsgBox("Company is deleted successfully")
                Me.Close()
            Else
                Exit Sub
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "Comments and Flags Events"
    Private Sub btnComments_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnComments.Click
        CommentsMaintenance(sender, e)
    End Sub
    Private Sub CommentsMaintenance(Optional ByVal sender As System.Object = Nothing, Optional ByVal e As System.EventArgs = Nothing, Optional ByVal bolSetCounts As Boolean = False, Optional ByVal resetBtnColor As Boolean = False)
        Dim SC As ShowComments
        Dim nEntityType As Integer = 0
        Dim strEntityName As String = String.Empty
        Dim nCommentsCount As Integer = 0
        Try
            If pCompany.ID <= 0 And pCompany.ID > -99 Then 'There are 30 compaines with negative company ID
                MsgBox("Please save Company before entering comments")
                Exit Sub
            End If
            nEntityType = UIUtilsGen.EntityTypes.Company
            strEntityName = "Company : " + CStr(pCompany.ID) + " " + pCompany.COMPANY_NAME

            If Not resetBtnColor Then
                SC = New ShowComments(pCompany.ID, nEntityType, IIf(bolSetCounts, "", "Company"), strEntityName, pCompany.Comments, Me.Text, , False)
                If bolSetCounts Then
                    nCommentsCount = SC.GetCounts()
                Else
                    SC.ShowDialog()
                    nCommentsCount = IIf(SC.nCommentsCount <= 0, SC.GetCounts(), SC.nCommentsCount)
                End If
            End If
            If nEntityType = UIUtilsGen.EntityTypes.Company Then
                If nCommentsCount > 0 Then
                    btnComments.BackColor = Drawing.Color.FromName(UIUtilsGen.CommentsButton_HasCmts)
                Else
                    btnComments.BackColor = Drawing.Color.FromName(UIUtilsGen.CommentsButton_NoCmts)
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnFlags_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnFlags.Click
        Try
            SF = New ShowFlags(CType(txtCompanyID.Text, Integer), UIUtilsGen.EntityTypes.Company, "Company")
            SF.ShowDialog()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub SF_FlagAdded(ByVal entityID As Integer, ByVal entityType As Integer, ByVal [Module] As String, ByVal ParentFormText As String) Handles SF.FlagAdded
        Dim MyFrm As MusterContainer
        MyFrm = Me.MdiParent
        MyFrm.FlagsChanged(entityID, entityType, [Module], ParentFormText)
    End Sub
    Private Sub SF_RefreshCalendar() Handles SF.RefreshCalendar
        Dim mc As MusterContainer = Me.MdiParent
        mc.RefreshCalendarInfo()
    End Sub
    Private Sub objLicen_FlagAdded(ByVal entityID As Integer, ByVal entityType As Integer, ByVal [Module] As String, ByVal ParentFormText As String) Handles objLicen.FlagAdded
        Dim MyFrm As MusterContainer
        MyFrm = Me.MdiParent
        MyFrm.FlagsChanged(txtCompanyID.Text, UIUtilsGen.EntityTypes.Company, "Company", ParentFormText)
    End Sub
    Private Sub objLicen_RefreshCalendar() Handles objLicen.RefreshCalendar
        Dim mc As MusterContainer = Me.MdiParent
        mc.RefreshCalendarInfo()
    End Sub
#End Region

#Region "External Events"
    Public Sub ValidationErrors(ByVal MsgStr As String) Handles pCompany.CompanyErr
        MsgBox(MsgStr)
    End Sub
    Private Sub oComAdd_evtAddressChanged(ByVal bolValue As Boolean) Handles oComAdd.evtAddressChanged
        setCompanySave(bolValue, False)
    End Sub
    Private Sub pLic_LicenseeChanged(ByVal bolValue As Boolean) Handles pLic.LicenseeChanged
        setCompanySave(bolValue, False)
    End Sub
    Private Sub pCompanyLicenseeAssociation_CompanyLicenseeChanged(ByVal bolValue As Boolean) Handles pCompanyLicenseeAssociation.CompanyLicenseeChanged
        setCompanySave(bolValue, False)
    End Sub
    Private Sub pCompany_evtCompanyChanged(ByVal bolValue As Boolean) Handles pCompany.evtCompanyChanged
        setCompanySave(bolValue, True)
    End Sub
    'Private Sub ugAddresses_AfterCellUpdate(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs)
    '    setCompanySave(True)
    'End Sub
#End Region

#Region "Enums"
    Enum EnumStatus
        CERTIFIED
        RENEWING
        APPLICANT
    End Enum
    Enum EnumCertificationType
        INSTALL
        CLOSURE
    End Enum
#End Region

#Region "Event Handlers for Expanding/Collapsing the different Sections of the Company form"
    Private Sub ExpandCollapse(ByRef pnl As Panel, ByRef lbl As Label)
        pnl.Visible = Not pnl.Visible
        lbl.Text = IIf(pnl.Visible, "-", "+")
    End Sub
    Private Sub lblCompanyDisplay_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblCompanyDisplay.Click
        ExpandCollapse(pnlCompanyDetails, lblCompanyDisplay)
    End Sub
    Private Sub lblCompanyHeader_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblCompanyHeader.Click
        ExpandCollapse(pnlCompanyDetails, lblCompanyDisplay)
    End Sub
    Private Sub lblLicenseeDisplay_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblLicenseeDisplay.Click
        ExpandCollapse(pnlLicenseesDetails, lblLicenseeDisplay)
    End Sub
    Private Sub lblManagerDisplay_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblManagerDisplay.Click
        ExpandCollapse(pnlManagersDetails, lblManagerDisplay)
    End Sub
    Private Sub lblLicenseeHead_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblLicenseeHead.Click
        ExpandCollapse(pnlLicenseesDetails, lblLicenseeDisplay)
    End Sub
    Private Sub lblManagerHead_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblManagerHead.Click
        ExpandCollapse(pnlManagersDetails, lblManagerDisplay)
    End Sub
    Private Sub lblEracDisplay_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblEracDisplay.Click
        ExpandCollapse(pnlEracDetails, lblEracDisplay)
    End Sub
    Private Sub lnlEracHead_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnlEracHead.Click
        ExpandCollapse(pnlEracDetails, lblEracDisplay)
    End Sub
    Private Sub lblTypesDisplay_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblTypesDisplay.Click
        ExpandCollapse(pnlTypesDetails, lblTypesDisplay)
    End Sub
    Private Sub lblTypesHead_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblTypesHead.Click
        ExpandCollapse(pnlTypesDetails, lblTypesDisplay)
    End Sub
    Private Sub lblPriorLicenseesDisplay_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblPriorLicenseesDisplay.Click
        ExpandCollapse(pnlPriorLicenseesDetails, lblPriorLicenseesDisplay)
    End Sub
    Private Sub lblPriorLicenseesHead_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblPriorLicenseesHead.Click
        ExpandCollapse(pnlPriorLicenseesDetails, lblPriorLicenseesDisplay)
    End Sub
    Private Sub lblDocumentsDisplay_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblDocumentsDisplay.Click
        ExpandCollapse(pnlDocumentDetails, lblDocumentsDisplay)
    End Sub
    Private Sub lblDocumentsHead_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblDocumentsHead.Click
        ExpandCollapse(pnlDocumentDetails, lblDocumentsDisplay)
    End Sub
    Private Sub lblCompanyContactsDisplay_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblCompanyContactsDisplay.Click
        ExpandCollapse(pnlCompanyContactHeader, lblCompanyContactsDisplay)
    End Sub
    Private Sub lblCompanyContactsHead_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblCompanyContactsHead.Click
        ExpandCollapse(pnlCompanyContactHeader, lblCompanyContactsDisplay)
    End Sub
#End Region

    Private Sub ugManagers_InitializeLayout(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugManagers.InitializeLayout

    End Sub


    Private Sub btnManagersAddNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnManagersAddNew.Click
        Dim msg As String = ""
        Try
            'setCompanySave(True)
            If txtCompany.Text = "" Then
                msg += "Please enter the company name before proceeding" + vbCrLf
            End If
            '   If ugAddresses.Rows.Count = 0 Then
            '   msg += "Please enter the company address before proceeding"
            '  End If
            If msg <> "" Then
                MsgBox(msg)
                Exit Sub
            End If
            pMgr.pLicenseeCourse = New MUSTER.BusinessLogic.pLicenseeCourses
            pMgr.pLicenseeCourseTest = New MUSTER.BusinessLogic.pLicenseeCourseTest

            Dim oLicCompAddInfo As MUSTER.Info.ComAddressInfo
            oLicCompAddInfo = oComAdd.Retrieve(nCompanyAddressID, False)
            msg = oLicCompAddInfo.AddressLine1 + IIf(oLicCompAddInfo.AddressLine1.Trim.Length = 0, "", vbCrLf + oLicCompAddInfo.AddressLine2) + vbCrLf + oLicCompAddInfo.City + ", " + oLicCompAddInfo.State + " " + oLicCompAddInfo.Zip
            objMgr = New Managers(pMgr, pCompanyLicenseeAssociation, oComAdd, nCompanyID, msg, oLicCompAddInfo.AddressId, 0, 0, pCompany.FIN_RESP_END_DATE, txtCompany.Text, True, "ADD")
            'objLicen.MdiParent = Me
            objMgr.ShowDialog()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnManagersEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnManagersEdit.Click
        Try
            PopulateManagers()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnManagersAddExisting_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnManagersAddExisting.Click
        Try
            'setCompanySave(True)
            oManagerList = New ManagersList(pCompany, pMgr)
            'oLicenseeList.MdiParent = Me
            oManagerList.ShowDialog()
            Me.populateLicenseesGrid()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnManagerDisassociate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnManagerDisassociate.Click
        Dim oCompanyLicenseeInfoLocal As MUSTER.Info.CompanyLicenseeInfo
        Dim result As DialogResult
        Dim resultTemp As DialogResult
        Try
            If ugManagers.ActiveRow Is Nothing Then
                Exit Sub
            End If
            'setCompanySave(True)
            result = MessageBox.Show("Do you want to Disassociate " + ugManagers.ActiveRow.Cells("Manager Name").Value + "?", "Manager Disassociation", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)
            If result = DialogResult.Yes Then
                For Each oCompanyLicenseeInfoLocal In pCompanyLicenseeAssociation.ComLicCollection.Values
                    If oCompanyLicenseeInfoLocal.LicenseeID = ugManagers.ActiveRow.Cells("Manager ID").Value Then
                        oCompanyLicenseeInfoLocal.Deleted = True
                        Exit For
                    End If
                Next
                pCompanyLicenseeAssociation.ModifiedBy = MusterContainer.AppUser.ID
                pCompanyLicenseeAssociation.Flush(CType(UIUtilsGen.ModuleID.Company, Integer), MusterContainer.AppUser.UserKey, returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                pMgr.Retrieve(ugManagers.ActiveRow.Cells("Manager ID").Value)
                pMgr.EMPLOYEE_LETTER = False
                pMgr.CERT_TYPE_ID = 0
                pMgr.CERT_TYPE_DESC = String.Empty
                pMgr.HIRE_STATUS = String.Empty
                'resultTemp = MessageBox.Show("Do you want to associate the Licensee to another Company?", "Association", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)
                'If resultTemp = DialogResult.No Then
                'pMgr.STATUS_ID = "NO LONGER WITH COMPANY"
                'End If
                If pMgr.ID <= 0 Then
                    pMgr.CreatedBy = MusterContainer.AppUser.ID
                Else
                    pMgr.ModifiedBy = MusterContainer.AppUser.ID
                End If
                pMgr.Save(CType(UIUtilsGen.ModuleID.Company, Integer), MusterContainer.AppUser.UserKey, returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                pMgr.Remove(ugManagers.ActiveRow.Cells("Manager ID").Value)
                'ugLicensees.DataSource = pMgr.EntityTable().DefaultView
                Me.populateLicenseesGrid()
                ' PopulatePriorLicenseeGrid()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
End Class
