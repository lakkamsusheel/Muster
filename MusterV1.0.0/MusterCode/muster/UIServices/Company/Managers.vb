Imports System.IO
Imports System.Text

Public Class Managers
    Inherits System.Windows.Forms.Form

#Region " User Defined Variables"
    Public MyGuid As New System.Guid
    'Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long

    Friend WithEvents objAddMaster As AddressMaster
    Friend WithEvents objAddresses As Addresses

    ' Company
    Dim WithEvents oCompany As New MUSTER.BusinessLogic.pCompany



    ' Company Address
    Private WithEvents oCompanyAdd As New MUSTER.BusinessLogic.pComAddress
    Private oComAddInfo As MUSTER.Info.ComAddressInfo

    ' Comany Manager Association
    Public WithEvents pCompanyManagerAssociation As MUSTER.BusinessLogic.pCompanyLicensee
    Dim CompanyLicenseeInfo As MUSTER.Info.CompanyLicenseeInfo

    ' Manager
    Public WithEvents pManager As MUSTER.BusinessLogic.pLicensee
    Dim ManagerInfo As MUSTER.Info.LicenseeInfo

    ' Manager Address
    Private WithEvents oLicAdd As New MUSTER.BusinessLogic.pComAddress

    ' Manager Course
    Private oLicCourseInfo As MUSTER.Info.LicenseeCourseInfo

    ' Manager Test
    Private oLicTestInfo As MUSTER.Info.LicenseeCourseTestInfo

    Private ncompAddressID As Integer = 0
    Private nLicAddressID As Integer = 0
    Dim strCompanyAddress As String = ""
    Dim dtNull As Date = CDate("01/01/0001")
    Dim nManagerID As Integer = 0
    Dim nCompanyID As Integer = 0
    Dim nCompanyManagerAssocID As Integer = 0
    Dim FinRespExpirationDate As Date = Now
    Dim strCompanyName As String = ""
    Dim nAssociationID As Integer = 0
    Private WithEvents SF As ShowFlags
    Dim dir As DirectoryInfo
    Dim finfo As FileInfo()
    Dim strPhotoFilePath As String '= strImagesLocation + "NoPhoto.gif"
    Dim strSignatureFilePath As String '= strImagesLocation + "NoSignature.gif"
    Dim strMode As String = ""
    Dim oLetter As New Reg_Letters
    Dim strFacilityIDs As String = String.Empty
    'Dim bolCongratsLetter As Boolean = False
    'Dim bolRenewalLetter As Boolean = False
    'Dim bolInfoNeededLetter As Boolean = False
    'Dim bolManagerCertificate As Boolean = False
    'Dim bolManagerCard As Boolean = False
    Public bolFormClosing As Boolean = False
    Dim bolModelWindow As Boolean = False
    Dim bolLoading As Boolean
    Dim strManagerHireStatus As String = String.Empty
    ' to prevent duplicate letter generation
    'Private slLetterCount As SortedList
    Dim returnVal As String = String.Empty
    Private bolValidatationFailed As Boolean = False
    Friend callingForm As Form
    Private Enum ManagerLetters
        CongLetter = 0
        ManagerCard = 1
        ManagerCertificateLetter = 2
        NoCertificationLetter = 3
        RenewalLetter = 4
    End Enum
    Friend ReadOnly Property ManagerID() As Integer
        Get
            Return nManagerID
        End Get
    End Property

    Public Event FlagAdded(ByVal entityID As Integer, ByVal entityType As Integer, ByVal [Module] As String, ByVal ParentFormText As String)
    Public Event RefreshCalendar()

#End Region

#Region " Windows Form Designer generated code "

    'Public Sub New()
    '    MyBase.New()
    '    bolLoading = True
    '    MyGuid = System.Guid.NewGuid
    '    'This call is required by the Windows Form Designer.
    '    InitializeComponent()
    '    strMode = "ADD"
    '    'Add any initialization after the InitializeComponent() call
    '    'Need to tell the AppUser that we've instantiated another Manager form...
    '    '
    '    MusterContainer.AppUser.LogEntry("Company", MyGuid.ToString)
    '    '
    '    ' The following line enables all forms to detect the visible form in the MDI container
    '    '
    '    MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "Company")

    '    If dir Is Nothing Then
    '        dir = New DirectoryInfo(MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_ManagersImages).ProfileValue + "\")
    '    End If
    '    finfo = dir.GetFiles("NoPhoto*")
    '    If finfo.Length > 0 Then
    '        strPhotoFilePath = finfo(0).FullName
    '    Else
    '        strPhotoFilePath = ""
    '    End If
    '    finfo = dir.GetFiles("NoSignature*")
    '    If finfo.Length > 0 Then
    '        strSignatureFilePath = finfo(0).FullName
    '    Else
    '        strSignatureFilePath = ""
    '    End If
    'End Sub

    Public Sub New(Optional ByVal _ManagerID As Integer = 0, Optional ByVal _CompanyID As Integer = 0, _
                    Optional ByVal _CompanyAddressID As Integer = 0, Optional ByVal _Mode As String = "Search")
        MyBase.New()
        MyGuid = System.Guid.NewGuid
        bolLoading = True
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

        nManagerID = _ManagerID  'MR - 6/5
        nCompanyID = _CompanyID
        ncompAddressID = _CompanyAddressID
        strMode = _Mode

        InitImages()
    End Sub

    Public Sub New(ByRef pMgr As MUSTER.BusinessLogic.pLicensee, _
                    ByRef pCompLicAssoc As MUSTER.BusinessLogic.pCompanyLicensee, _
                    ByRef pComAddress As MUSTER.BusinessLogic.pComAddress, _
                    ByVal CompanyID As Integer, ByVal strAddress As String, ByVal nLicCompanyAddID As Integer, _
                    ByVal _ManagerID As Integer, ByVal AssociationID As Integer, ByVal FinRespExpDate As Date, _
                    Optional ByVal companyName As String = "", _
                    Optional ByVal Modelwindow As Boolean = False, Optional ByVal Mode As String = "")
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        bolLoading = True
        MyGuid = System.Guid.NewGuid
        'Add any initialization after the InitializeComponent() call

        bolModelWindow = Modelwindow
        If Not bolModelWindow Then
            'Need to tell the AppUser that we've instantiated another Registration form...
            '
            MusterContainer.AppUser.LogEntry("Company", MyGuid.ToString)
            '
            ' The following line enables all forms to detect the visible form in the MDI container
            '
            MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "Company")
        End If

        pManager = pMgr
        pCompanyManagerAssociation = pCompLicAssoc
        oCompanyAdd = pComAddress
        nCompanyID = CompanyID
        strCompanyAddress = strAddress
        ncompAddressID = nLicCompanyAddID
        nManagerID = _ManagerID
        nAssociationID = AssociationID
        strCompanyName = companyName
        FinRespExpirationDate = FinRespExpDate
        If Mode <> String.Empty Then
            strMode = Mode
        Else
            strMode = "Company"
        End If

        InitImages()
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If Not bolModelWindow Then
            ' Remove any values from the shared collection for this screen
            '
            MusterContainer.AppSemaphores.Remove(MyGuid.ToString)
            '
            ' Log the disposal of the form (exit from Registration form)
            '
            MusterContainer.AppUser.LogExit(MyGuid.ToString)
        End If

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
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents lblLicenseNo As System.Windows.Forms.Label
    Friend WithEvents btnFlags As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnDeleteManager As System.Windows.Forms.Button
    Friend WithEvents btnComments As System.Windows.Forms.Button
    Friend WithEvents btnSaveManager As System.Windows.Forms.Button
    Friend WithEvents cmbTitle As System.Windows.Forms.ComboBox
    Friend WithEvents txtLastName As System.Windows.Forms.TextBox
    Friend WithEvents txtFirstName As System.Windows.Forms.TextBox
    Friend WithEvents txtMiddleName As System.Windows.Forms.TextBox
    Friend WithEvents lblMiddleName As System.Windows.Forms.Label
    Friend WithEvents lblFirstName As System.Windows.Forms.Label
    Friend WithEvents lblLastName As System.Windows.Forms.Label
    Friend WithEvents pnlManagersBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlManagers As System.Windows.Forms.Panel
    ' Friend WithEvents ugRelations As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents lblPhoto As System.Windows.Forms.Label
    Friend WithEvents lblSignature As System.Windows.Forms.Label
    Public WithEvents dtExceptionGrantedDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblExceptionGrantedDate As System.Windows.Forms.Label
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents lblCertificationType As System.Windows.Forms.Label
    Friend WithEvents txtEmail As System.Windows.Forms.TextBox
    Friend WithEvents lblEmail As System.Windows.Forms.Label
    Friend WithEvents txtLicenseNo As System.Windows.Forms.TextBox
    Friend WithEvents lblLicenseNumber As System.Windows.Forms.Label
    Friend WithEvents lblPersonTitle As System.Windows.Forms.Label
    Friend WithEvents cmbSuffix As System.Windows.Forms.ComboBox
    Friend WithEvents lblSuffix As System.Windows.Forms.Label
    Friend WithEvents txtManagerAddress As System.Windows.Forms.TextBox
    Friend WithEvents lblManagerAddress As System.Windows.Forms.Label
    Friend WithEvents pbSign As System.Windows.Forms.PictureBox
    Friend WithEvents pbPhoto As System.Windows.Forms.PictureBox
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents pnlPhone As System.Windows.Forms.Panel
    Public WithEvents mskTxtPhone1 As AxMSMask.AxMaskEdBox
    Friend WithEvents lblPhone2 As System.Windows.Forms.Label
    Public WithEvents mskTxtCell As AxMSMask.AxMaskEdBox
    Friend WithEvents lblCell As System.Windows.Forms.Label
    Public WithEvents mskTxtPhone2 As AxMSMask.AxMaskEdBox
    Friend WithEvents lblPhone1 As System.Windows.Forms.Label
    Public WithEvents mskTxtFax As AxMSMask.AxMaskEdBox
    Friend WithEvents lblFax As System.Windows.Forms.Label
    Friend WithEvents pnlExtension As System.Windows.Forms.Panel
    Friend WithEvents lblExt2 As System.Windows.Forms.Label
    Friend WithEvents lblExt1 As System.Windows.Forms.Label
    Friend WithEvents txtExt2 As System.Windows.Forms.TextBox
    Friend WithEvents txtExt1 As System.Windows.Forms.TextBox
    Friend WithEvents btnLabels As System.Windows.Forms.Button
    Friend WithEvents btnEnvelopes As System.Windows.Forms.Button
    Friend WithEvents UCManagerDocuments As MUSTER.DocumentViewControl
    Friend WithEvents ugManagerComplianceEvents As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents ugPriorCompanies As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents ugTests As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnTestModify As System.Windows.Forms.Button
    Friend WithEvents btnTestAdd As System.Windows.Forms.Button
    Friend WithEvents btnTestDelete As System.Windows.Forms.Button
    Friend WithEvents pnlCompany As System.Windows.Forms.Panel
    Friend WithEvents btnCompanyTitleSearch As System.Windows.Forms.Button
    Friend WithEvents txtCompanyTitle As System.Windows.Forms.TextBox
    Friend WithEvents txtCompany As System.Windows.Forms.TextBox
    Friend WithEvents lblCompany As System.Windows.Forms.Label
    Friend WithEvents lblCompanyTitle As System.Windows.Forms.Label
    Friend WithEvents pnlManagerInfo As System.Windows.Forms.Panel
    Friend WithEvents pnlManagerInfoHead As System.Windows.Forms.Panel
    Friend WithEvents pnlCoursesHead As System.Windows.Forms.Panel
    Friend WithEvents pnlCourses As System.Windows.Forms.Panel
    Friend WithEvents pnlTestsHead As System.Windows.Forms.Panel
    Friend WithEvents pnlTests As System.Windows.Forms.Panel
    Friend WithEvents pnlPriorCompaniesHead As System.Windows.Forms.Panel
    Friend WithEvents pnlPriorCompanies As System.Windows.Forms.Panel
    Friend WithEvents pnlLCEHead As System.Windows.Forms.Panel
    Friend WithEvents pnlLCE As System.Windows.Forms.Panel
    Friend WithEvents pnlDocumentsHead As System.Windows.Forms.Panel
    Friend WithEvents pnlDocuments As System.Windows.Forms.Panel
    Friend WithEvents lblManagerInfoDisplay As System.Windows.Forms.Label
    Friend WithEvents lblManagerInfo As System.Windows.Forms.Label
    Friend WithEvents lblCoursesDisplay As System.Windows.Forms.Label
    Friend WithEvents lblTestsDisplay As System.Windows.Forms.Label
    Friend WithEvents lblTests As System.Windows.Forms.Label
    Friend WithEvents lblPriorCompaniesDisplay As System.Windows.Forms.Label
    Friend WithEvents lblLCEDisplay As System.Windows.Forms.Label
    Friend WithEvents lblDocumentsDisplay As System.Windows.Forms.Label
    Friend WithEvents lblDocuments As System.Windows.Forms.Label
    Friend WithEvents lblPriorCompanies As System.Windows.Forms.Label
    Friend WithEvents lblLCE As System.Windows.Forms.Label
    Friend WithEvents pnlCoursesBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlTestsBottom As System.Windows.Forms.Panel
    Friend WithEvents btnGenExpirationLetter As System.Windows.Forms.Button
    Friend WithEvents btnGenReminderLetter As System.Windows.Forms.Button
    Friend WithEvents btnCertify As System.Windows.Forms.Button
    Public WithEvents dtExtensionDeadlineDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnCompanyEnvelopes As System.Windows.Forms.Button
    Friend WithEvents btnCompanyLabels As System.Windows.Forms.Button
    Friend WithEvents lblRestCert As System.Windows.Forms.Label
    Friend WithEvents LblInitCertBy As System.Windows.Forms.Label
    Friend WithEvents LblInitCertDate As System.Windows.Forms.Label
    Public WithEvents dtInitCertDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents LblRetrainDate1 As System.Windows.Forms.Label
    Public WithEvents dtRetrainDate2 As System.Windows.Forms.DateTimePicker
    Friend WithEvents LblRetrainDate2 As System.Windows.Forms.Label
    Friend WithEvents LblRetrainDate3 As System.Windows.Forms.Label
    Public WithEvents dtRetrainDate3 As System.Windows.Forms.DateTimePicker
    Friend WithEvents LblRevokeDate As System.Windows.Forms.Label
    Public WithEvents dtRevokeDate As System.Windows.Forms.DateTimePicker
    Public WithEvents dtRetrainDate1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents cmbInitCertBy As System.Windows.Forms.ComboBox
    Friend WithEvents chkIsLicensee As System.Windows.Forms.CheckBox
    Friend WithEvents cmbCMCertificationType As System.Windows.Forms.ComboBox
    Friend WithEvents cmbCMStatus As System.Windows.Forms.ComboBox
    Friend WithEvents lblRelations As System.Windows.Forms.Label
    Friend WithEvents ugRelations As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnRelationAdd As System.Windows.Forms.Button
    Friend WithEvents btnRelationDelete As System.Windows.Forms.Button
    Friend WithEvents btnRelationModify As System.Windows.Forms.Button
    Public WithEvents dtRetrainReqDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblRetrainReqDate As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Managers))
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.lblLicenseNo = New System.Windows.Forms.Label
        Me.pnlManagersBottom = New System.Windows.Forms.Panel
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnFlags = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnDeleteManager = New System.Windows.Forms.Button
        Me.btnComments = New System.Windows.Forms.Button
        Me.btnSaveManager = New System.Windows.Forms.Button
        Me.btnGenExpirationLetter = New System.Windows.Forms.Button
        Me.btnGenReminderLetter = New System.Windows.Forms.Button
        Me.btnCertify = New System.Windows.Forms.Button
        Me.pnlManagers = New System.Windows.Forms.Panel
        Me.pnlCoursesHead = New System.Windows.Forms.Panel
        Me.lblRelations = New System.Windows.Forms.Label
        Me.lblCoursesDisplay = New System.Windows.Forms.Label
        Me.pnlCourses = New System.Windows.Forms.Panel
        Me.ugRelations = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlCoursesBottom = New System.Windows.Forms.Panel
        Me.btnRelationAdd = New System.Windows.Forms.Button
        Me.btnRelationModify = New System.Windows.Forms.Button
        Me.btnRelationDelete = New System.Windows.Forms.Button
        Me.pnlTestsHead = New System.Windows.Forms.Panel
        Me.lblTests = New System.Windows.Forms.Label
        Me.lblTestsDisplay = New System.Windows.Forms.Label
        Me.pnlTests = New System.Windows.Forms.Panel
        Me.pnlTestsBottom = New System.Windows.Forms.Panel
        Me.btnTestAdd = New System.Windows.Forms.Button
        Me.btnTestModify = New System.Windows.Forms.Button
        Me.btnTestDelete = New System.Windows.Forms.Button
        Me.ugTests = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlPriorCompaniesHead = New System.Windows.Forms.Panel
        Me.lblPriorCompanies = New System.Windows.Forms.Label
        Me.lblPriorCompaniesDisplay = New System.Windows.Forms.Label
        Me.pnlPriorCompanies = New System.Windows.Forms.Panel
        Me.ugPriorCompanies = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlLCEHead = New System.Windows.Forms.Panel
        Me.lblLCE = New System.Windows.Forms.Label
        Me.lblLCEDisplay = New System.Windows.Forms.Label
        Me.ugManagerComplianceEvents = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlLCE = New System.Windows.Forms.Panel
        Me.pnlDocumentsHead = New System.Windows.Forms.Panel
        Me.lblDocuments = New System.Windows.Forms.Label
        Me.lblDocumentsDisplay = New System.Windows.Forms.Label
        Me.pnlDocuments = New System.Windows.Forms.Panel
        Me.UCManagerDocuments = New MUSTER.DocumentViewControl
        Me.pnlManagerInfo = New System.Windows.Forms.Panel
        Me.lblRetrainReqDate = New System.Windows.Forms.Label
        Me.dtRetrainReqDate = New System.Windows.Forms.DateTimePicker
        Me.chkIsLicensee = New System.Windows.Forms.CheckBox
        Me.dtRevokeDate = New System.Windows.Forms.DateTimePicker
        Me.LblRevokeDate = New System.Windows.Forms.Label
        Me.dtRetrainDate3 = New System.Windows.Forms.DateTimePicker
        Me.LblRetrainDate3 = New System.Windows.Forms.Label
        Me.dtRetrainDate2 = New System.Windows.Forms.DateTimePicker
        Me.LblRetrainDate2 = New System.Windows.Forms.Label
        Me.dtRetrainDate1 = New System.Windows.Forms.DateTimePicker
        Me.LblRetrainDate1 = New System.Windows.Forms.Label
        Me.dtInitCertDate = New System.Windows.Forms.DateTimePicker
        Me.LblInitCertDate = New System.Windows.Forms.Label
        Me.cmbInitCertBy = New System.Windows.Forms.ComboBox
        Me.LblInitCertBy = New System.Windows.Forms.Label
        Me.lblRestCert = New System.Windows.Forms.Label
        Me.cmbCMCertificationType = New System.Windows.Forms.ComboBox
        Me.cmbCMStatus = New System.Windows.Forms.ComboBox
        Me.pnlCompany = New System.Windows.Forms.Panel
        Me.btnCompanyLabels = New System.Windows.Forms.Button
        Me.btnCompanyEnvelopes = New System.Windows.Forms.Button
        Me.btnCompanyTitleSearch = New System.Windows.Forms.Button
        Me.txtCompanyTitle = New System.Windows.Forms.TextBox
        Me.txtCompany = New System.Windows.Forms.TextBox
        Me.lblCompany = New System.Windows.Forms.Label
        Me.lblCompanyTitle = New System.Windows.Forms.Label
        Me.pbPhoto = New System.Windows.Forms.PictureBox
        Me.lblPhoto = New System.Windows.Forms.Label
        Me.pbSign = New System.Windows.Forms.PictureBox
        Me.lblSignature = New System.Windows.Forms.Label
        Me.dtExtensionDeadlineDate = New System.Windows.Forms.DateTimePicker
        Me.dtExceptionGrantedDate = New System.Windows.Forms.DateTimePicker
        Me.lblExceptionGrantedDate = New System.Windows.Forms.Label
        Me.lblStatus = New System.Windows.Forms.Label
        Me.lblCertificationType = New System.Windows.Forms.Label
        Me.pnlExtension = New System.Windows.Forms.Panel
        Me.pnlPhone = New System.Windows.Forms.Panel
        Me.mskTxtPhone1 = New AxMSMask.AxMaskEdBox
        Me.lblPhone2 = New System.Windows.Forms.Label
        Me.mskTxtCell = New AxMSMask.AxMaskEdBox
        Me.lblCell = New System.Windows.Forms.Label
        Me.mskTxtPhone2 = New AxMSMask.AxMaskEdBox
        Me.lblPhone1 = New System.Windows.Forms.Label
        Me.mskTxtFax = New AxMSMask.AxMaskEdBox
        Me.lblFax = New System.Windows.Forms.Label
        Me.lblExt2 = New System.Windows.Forms.Label
        Me.lblExt1 = New System.Windows.Forms.Label
        Me.txtExt2 = New System.Windows.Forms.TextBox
        Me.txtExt1 = New System.Windows.Forms.TextBox
        Me.txtManagerAddress = New System.Windows.Forms.TextBox
        Me.lblManagerAddress = New System.Windows.Forms.Label
        Me.btnLabels = New System.Windows.Forms.Button
        Me.btnEnvelopes = New System.Windows.Forms.Button
        Me.lblEmail = New System.Windows.Forms.Label
        Me.txtEmail = New System.Windows.Forms.TextBox
        Me.lblSuffix = New System.Windows.Forms.Label
        Me.lblLastName = New System.Windows.Forms.Label
        Me.cmbSuffix = New System.Windows.Forms.ComboBox
        Me.txtLastName = New System.Windows.Forms.TextBox
        Me.lblMiddleName = New System.Windows.Forms.Label
        Me.lblFirstName = New System.Windows.Forms.Label
        Me.lblPersonTitle = New System.Windows.Forms.Label
        Me.lblLicenseNumber = New System.Windows.Forms.Label
        Me.txtLicenseNo = New System.Windows.Forms.TextBox
        Me.cmbTitle = New System.Windows.Forms.ComboBox
        Me.txtFirstName = New System.Windows.Forms.TextBox
        Me.txtMiddleName = New System.Windows.Forms.TextBox
        Me.pnlManagerInfoHead = New System.Windows.Forms.Panel
        Me.lblManagerInfo = New System.Windows.Forms.Label
        Me.lblManagerInfoDisplay = New System.Windows.Forms.Label
        Me.pnlManagersBottom.SuspendLayout()
        Me.pnlManagers.SuspendLayout()
        Me.pnlCoursesHead.SuspendLayout()
        Me.pnlCourses.SuspendLayout()
        CType(Me.ugRelations, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlCoursesBottom.SuspendLayout()
        Me.pnlTestsHead.SuspendLayout()
        Me.pnlTestsBottom.SuspendLayout()
        CType(Me.ugTests, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlPriorCompaniesHead.SuspendLayout()
        Me.pnlPriorCompanies.SuspendLayout()
        CType(Me.ugPriorCompanies, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlLCEHead.SuspendLayout()
        CType(Me.ugManagerComplianceEvents, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlDocumentsHead.SuspendLayout()
        Me.pnlDocuments.SuspendLayout()
        Me.pnlManagerInfo.SuspendLayout()
        Me.pnlCompany.SuspendLayout()
        Me.pnlPhone.SuspendLayout()
        CType(Me.mskTxtPhone1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtCell, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtPhone2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtFax, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlManagerInfoHead.SuspendLayout()
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(104, -56)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(144, 20)
        Me.TextBox1.TabIndex = 181
        Me.TextBox1.Text = "TextBox1"
        '
        'lblLicenseNo
        '
        Me.lblLicenseNo.Location = New System.Drawing.Point(40, -56)
        Me.lblLicenseNo.Name = "lblLicenseNo"
        Me.lblLicenseNo.Size = New System.Drawing.Size(64, 56)
        Me.lblLicenseNo.TabIndex = 182
        Me.lblLicenseNo.Text = "License #:"
        Me.lblLicenseNo.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlManagersBottom
        '
        Me.pnlManagersBottom.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlManagersBottom.Controls.Add(Me.btnClose)
        Me.pnlManagersBottom.Controls.Add(Me.btnFlags)
        Me.pnlManagersBottom.Controls.Add(Me.btnCancel)
        Me.pnlManagersBottom.Controls.Add(Me.btnDeleteManager)
        Me.pnlManagersBottom.Controls.Add(Me.btnComments)
        Me.pnlManagersBottom.Controls.Add(Me.btnSaveManager)
        Me.pnlManagersBottom.Controls.Add(Me.btnGenExpirationLetter)
        Me.pnlManagersBottom.Controls.Add(Me.btnGenReminderLetter)
        Me.pnlManagersBottom.Controls.Add(Me.btnCertify)
        Me.pnlManagersBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlManagersBottom.Location = New System.Drawing.Point(0, 565)
        Me.pnlManagersBottom.Name = "pnlManagersBottom"
        Me.pnlManagersBottom.Size = New System.Drawing.Size(904, 40)
        Me.pnlManagersBottom.TabIndex = 1
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(240, 9)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(72, 26)
        Me.btnClose.TabIndex = 5
        Me.btnClose.Text = "Close"
        '
        'btnFlags
        '
        Me.btnFlags.Enabled = False
        Me.btnFlags.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnFlags.Location = New System.Drawing.Point(640, 8)
        Me.btnFlags.Name = "btnFlags"
        Me.btnFlags.Size = New System.Drawing.Size(96, 26)
        Me.btnFlags.TabIndex = 4
        Me.btnFlags.Text = "Flags"
        '
        'btnCancel
        '
        Me.btnCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCancel.Location = New System.Drawing.Point(80, 8)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(72, 26)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "Cancel"
        '
        'btnDeleteManager
        '
        Me.btnDeleteManager.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDeleteManager.Location = New System.Drawing.Point(160, 9)
        Me.btnDeleteManager.Name = "btnDeleteManager"
        Me.btnDeleteManager.Size = New System.Drawing.Size(72, 26)
        Me.btnDeleteManager.TabIndex = 2
        Me.btnDeleteManager.Text = "Delete"
        '
        'btnComments
        '
        Me.btnComments.Enabled = False
        Me.btnComments.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnComments.Location = New System.Drawing.Point(536, 8)
        Me.btnComments.Name = "btnComments"
        Me.btnComments.Size = New System.Drawing.Size(96, 26)
        Me.btnComments.TabIndex = 3
        Me.btnComments.Text = "Comments"
        '
        'btnSaveManager
        '
        Me.btnSaveManager.Enabled = False
        Me.btnSaveManager.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSaveManager.Location = New System.Drawing.Point(8, 8)
        Me.btnSaveManager.Name = "btnSaveManager"
        Me.btnSaveManager.Size = New System.Drawing.Size(64, 26)
        Me.btnSaveManager.TabIndex = 0
        Me.btnSaveManager.Text = "Save"
        '
        'btnGenExpirationLetter
        '
        Me.btnGenExpirationLetter.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGenExpirationLetter.Location = New System.Drawing.Point(408, 48)
        Me.btnGenExpirationLetter.Name = "btnGenExpirationLetter"
        Me.btnGenExpirationLetter.Size = New System.Drawing.Size(160, 26)
        Me.btnGenExpirationLetter.TabIndex = 3
        Me.btnGenExpirationLetter.Text = "Generate Expiration Letter"
        '
        'btnGenReminderLetter
        '
        Me.btnGenReminderLetter.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGenReminderLetter.Location = New System.Drawing.Point(576, 48)
        Me.btnGenReminderLetter.Name = "btnGenReminderLetter"
        Me.btnGenReminderLetter.Size = New System.Drawing.Size(160, 26)
        Me.btnGenReminderLetter.TabIndex = 3
        Me.btnGenReminderLetter.Text = "Generate Reminder Letter"
        '
        'btnCertify
        '
        Me.btnCertify.Enabled = False
        Me.btnCertify.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCertify.Location = New System.Drawing.Point(328, 48)
        Me.btnCertify.Name = "btnCertify"
        Me.btnCertify.Size = New System.Drawing.Size(72, 26)
        Me.btnCertify.TabIndex = 3
        Me.btnCertify.Text = "Certify"
        '
        'pnlManagers
        '
        Me.pnlManagers.AutoScroll = True
        Me.pnlManagers.Controls.Add(Me.pnlCoursesHead)
        Me.pnlManagers.Controls.Add(Me.pnlCourses)
        Me.pnlManagers.Controls.Add(Me.pnlCoursesBottom)
        Me.pnlManagers.Controls.Add(Me.pnlTestsHead)
        Me.pnlManagers.Controls.Add(Me.pnlTests)
        Me.pnlManagers.Controls.Add(Me.pnlTestsBottom)
        Me.pnlManagers.Controls.Add(Me.pnlPriorCompaniesHead)
        Me.pnlManagers.Controls.Add(Me.pnlPriorCompanies)
        Me.pnlManagers.Controls.Add(Me.pnlLCEHead)
        Me.pnlManagers.Controls.Add(Me.pnlLCE)
        Me.pnlManagers.Controls.Add(Me.pnlDocumentsHead)
        Me.pnlManagers.Controls.Add(Me.pnlDocuments)
        Me.pnlManagers.Controls.Add(Me.pnlManagerInfo)
        Me.pnlManagers.Controls.Add(Me.pnlManagerInfoHead)
        Me.pnlManagers.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlManagers.Location = New System.Drawing.Point(0, 0)
        Me.pnlManagers.Name = "pnlManagers"
        Me.pnlManagers.Size = New System.Drawing.Size(904, 565)
        Me.pnlManagers.TabIndex = 0
        '
        'pnlCoursesHead
        '
        Me.pnlCoursesHead.Controls.Add(Me.lblRelations)
        Me.pnlCoursesHead.Controls.Add(Me.lblCoursesDisplay)
        Me.pnlCoursesHead.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlCoursesHead.Location = New System.Drawing.Point(0, 392)
        Me.pnlCoursesHead.Name = "pnlCoursesHead"
        Me.pnlCoursesHead.Size = New System.Drawing.Size(888, 24)
        Me.pnlCoursesHead.TabIndex = 1063
        '
        'lblRelations
        '
        Me.lblRelations.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblRelations.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblRelations.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblRelations.Location = New System.Drawing.Point(16, 0)
        Me.lblRelations.Name = "lblRelations"
        Me.lblRelations.Size = New System.Drawing.Size(872, 24)
        Me.lblRelations.TabIndex = 256
        Me.lblRelations.Text = "Relation to UST Facility"
        Me.lblRelations.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblCoursesDisplay
        '
        Me.lblCoursesDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblCoursesDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblCoursesDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblCoursesDisplay.Name = "lblCoursesDisplay"
        Me.lblCoursesDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblCoursesDisplay.TabIndex = 255
        Me.lblCoursesDisplay.Text = "-"
        Me.lblCoursesDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'pnlCourses
        '
        Me.pnlCourses.Controls.Add(Me.ugRelations)
        Me.pnlCourses.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlCourses.Location = New System.Drawing.Point(0, 416)
        Me.pnlCourses.Name = "pnlCourses"
        Me.pnlCourses.Size = New System.Drawing.Size(888, 100)
        Me.pnlCourses.TabIndex = 1062
        '
        'ugRelations
        '
        Me.ugRelations.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugRelations.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom
        Me.ugRelations.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugRelations.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugRelations.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugRelations.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugRelations.Location = New System.Drawing.Point(0, 0)
        Me.ugRelations.Name = "ugRelations"
        Me.ugRelations.Size = New System.Drawing.Size(888, 100)
        Me.ugRelations.TabIndex = 21
        '
        'pnlCoursesBottom
        '
        Me.pnlCoursesBottom.Controls.Add(Me.btnRelationAdd)
        Me.pnlCoursesBottom.Controls.Add(Me.btnRelationModify)
        Me.pnlCoursesBottom.Controls.Add(Me.btnRelationDelete)
        Me.pnlCoursesBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlCoursesBottom.Location = New System.Drawing.Point(0, 516)
        Me.pnlCoursesBottom.Name = "pnlCoursesBottom"
        Me.pnlCoursesBottom.Size = New System.Drawing.Size(888, 40)
        Me.pnlCoursesBottom.TabIndex = 269
        '
        'btnRelationAdd
        '
        Me.btnRelationAdd.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRelationAdd.Location = New System.Drawing.Point(40, 8)
        Me.btnRelationAdd.Name = "btnRelationAdd"
        Me.btnRelationAdd.Size = New System.Drawing.Size(64, 26)
        Me.btnRelationAdd.TabIndex = 267
        Me.btnRelationAdd.Text = "Add"
        '
        'btnRelationModify
        '
        Me.btnRelationModify.Enabled = False
        Me.btnRelationModify.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRelationModify.Location = New System.Drawing.Point(184, 8)
        Me.btnRelationModify.Name = "btnRelationModify"
        Me.btnRelationModify.Size = New System.Drawing.Size(64, 26)
        Me.btnRelationModify.TabIndex = 266
        Me.btnRelationModify.Text = "Modify"
        Me.btnRelationModify.Visible = False
        '
        'btnRelationDelete
        '
        Me.btnRelationDelete.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnRelationDelete.Location = New System.Drawing.Point(112, 8)
        Me.btnRelationDelete.Name = "btnRelationDelete"
        Me.btnRelationDelete.Size = New System.Drawing.Size(64, 26)
        Me.btnRelationDelete.TabIndex = 268
        Me.btnRelationDelete.Text = "Delete"
        '
        'pnlTestsHead
        '
        Me.pnlTestsHead.Controls.Add(Me.lblTests)
        Me.pnlTestsHead.Controls.Add(Me.lblTestsDisplay)
        Me.pnlTestsHead.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlTestsHead.Location = New System.Drawing.Point(0, 556)
        Me.pnlTestsHead.Name = "pnlTestsHead"
        Me.pnlTestsHead.Size = New System.Drawing.Size(888, 24)
        Me.pnlTestsHead.TabIndex = 1061
        '
        'lblTests
        '
        Me.lblTests.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblTests.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblTests.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblTests.Location = New System.Drawing.Point(16, 0)
        Me.lblTests.Name = "lblTests"
        Me.lblTests.Size = New System.Drawing.Size(872, 24)
        Me.lblTests.TabIndex = 259
        Me.lblTests.Text = "Tests"
        Me.lblTests.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblTests.Visible = False
        '
        'lblTestsDisplay
        '
        Me.lblTestsDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTestsDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblTestsDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblTestsDisplay.Name = "lblTestsDisplay"
        Me.lblTestsDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblTestsDisplay.TabIndex = 258
        Me.lblTestsDisplay.Text = "-"
        Me.lblTestsDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblTestsDisplay.Visible = False
        '
        'pnlTests
        '
        Me.pnlTests.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlTests.Location = New System.Drawing.Point(0, 580)
        Me.pnlTests.Name = "pnlTests"
        Me.pnlTests.Size = New System.Drawing.Size(888, 100)
        Me.pnlTests.TabIndex = 1063
        '
        'pnlTestsBottom
        '
        Me.pnlTestsBottom.Controls.Add(Me.btnTestAdd)
        Me.pnlTestsBottom.Controls.Add(Me.btnTestModify)
        Me.pnlTestsBottom.Controls.Add(Me.btnTestDelete)
        Me.pnlTestsBottom.Controls.Add(Me.ugTests)
        Me.pnlTestsBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlTestsBottom.Location = New System.Drawing.Point(0, 680)
        Me.pnlTestsBottom.Name = "pnlTestsBottom"
        Me.pnlTestsBottom.Size = New System.Drawing.Size(888, 40)
        Me.pnlTestsBottom.TabIndex = 273
        Me.pnlTestsBottom.Visible = False
        '
        'btnTestAdd
        '
        Me.btnTestAdd.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnTestAdd.Location = New System.Drawing.Point(40, 8)
        Me.btnTestAdd.Name = "btnTestAdd"
        Me.btnTestAdd.Size = New System.Drawing.Size(64, 26)
        Me.btnTestAdd.TabIndex = 271
        Me.btnTestAdd.Text = "Add"
        Me.btnTestAdd.Visible = False
        '
        'btnTestModify
        '
        Me.btnTestModify.Enabled = False
        Me.btnTestModify.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnTestModify.Location = New System.Drawing.Point(184, 8)
        Me.btnTestModify.Name = "btnTestModify"
        Me.btnTestModify.Size = New System.Drawing.Size(64, 26)
        Me.btnTestModify.TabIndex = 272
        Me.btnTestModify.Text = "Modify"
        Me.btnTestModify.Visible = False
        '
        'btnTestDelete
        '
        Me.btnTestDelete.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnTestDelete.Location = New System.Drawing.Point(112, 8)
        Me.btnTestDelete.Name = "btnTestDelete"
        Me.btnTestDelete.Size = New System.Drawing.Size(64, 26)
        Me.btnTestDelete.TabIndex = 270
        Me.btnTestDelete.Text = "Delete"
        Me.btnTestDelete.Visible = False
        '
        'ugTests
        '
        Me.ugTests.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugTests.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom
        Me.ugTests.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugTests.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugTests.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugTests.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugTests.Location = New System.Drawing.Point(0, 0)
        Me.ugTests.Name = "ugTests"
        Me.ugTests.Size = New System.Drawing.Size(888, 40)
        Me.ugTests.TabIndex = 24
        Me.ugTests.Visible = False
        '
        'pnlPriorCompaniesHead
        '
        Me.pnlPriorCompaniesHead.Controls.Add(Me.lblPriorCompanies)
        Me.pnlPriorCompaniesHead.Controls.Add(Me.lblPriorCompaniesDisplay)
        Me.pnlPriorCompaniesHead.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlPriorCompaniesHead.Location = New System.Drawing.Point(0, 720)
        Me.pnlPriorCompaniesHead.Name = "pnlPriorCompaniesHead"
        Me.pnlPriorCompaniesHead.Size = New System.Drawing.Size(888, 24)
        Me.pnlPriorCompaniesHead.TabIndex = 1063
        '
        'lblPriorCompanies
        '
        Me.lblPriorCompanies.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblPriorCompanies.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblPriorCompanies.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblPriorCompanies.Location = New System.Drawing.Point(16, 0)
        Me.lblPriorCompanies.Name = "lblPriorCompanies"
        Me.lblPriorCompanies.Size = New System.Drawing.Size(872, 24)
        Me.lblPriorCompanies.TabIndex = 262
        Me.lblPriorCompanies.Text = "Prior Companies"
        Me.lblPriorCompanies.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblPriorCompanies.Visible = False
        '
        'lblPriorCompaniesDisplay
        '
        Me.lblPriorCompaniesDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPriorCompaniesDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblPriorCompaniesDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblPriorCompaniesDisplay.Name = "lblPriorCompaniesDisplay"
        Me.lblPriorCompaniesDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblPriorCompaniesDisplay.TabIndex = 261
        Me.lblPriorCompaniesDisplay.Text = "-"
        Me.lblPriorCompaniesDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblPriorCompaniesDisplay.Visible = False
        '
        'pnlPriorCompanies
        '
        Me.pnlPriorCompanies.Controls.Add(Me.ugPriorCompanies)
        Me.pnlPriorCompanies.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlPriorCompanies.Location = New System.Drawing.Point(0, 744)
        Me.pnlPriorCompanies.Name = "pnlPriorCompanies"
        Me.pnlPriorCompanies.Size = New System.Drawing.Size(888, 200)
        Me.pnlPriorCompanies.TabIndex = 1063
        '
        'ugPriorCompanies
        '
        Me.ugPriorCompanies.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugPriorCompanies.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugPriorCompanies.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugPriorCompanies.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugPriorCompanies.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugPriorCompanies.Location = New System.Drawing.Point(0, 0)
        Me.ugPriorCompanies.Name = "ugPriorCompanies"
        Me.ugPriorCompanies.Size = New System.Drawing.Size(888, 200)
        Me.ugPriorCompanies.TabIndex = 26
        Me.ugPriorCompanies.Visible = False
        '
        'pnlLCEHead
        '
        Me.pnlLCEHead.Controls.Add(Me.lblLCE)
        Me.pnlLCEHead.Controls.Add(Me.lblLCEDisplay)
        Me.pnlLCEHead.Controls.Add(Me.ugManagerComplianceEvents)
        Me.pnlLCEHead.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlLCEHead.Location = New System.Drawing.Point(0, 944)
        Me.pnlLCEHead.Name = "pnlLCEHead"
        Me.pnlLCEHead.Size = New System.Drawing.Size(888, 24)
        Me.pnlLCEHead.TabIndex = 1063
        '
        'lblLCE
        '
        Me.lblLCE.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblLCE.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblLCE.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblLCE.Location = New System.Drawing.Point(16, 0)
        Me.lblLCE.Name = "lblLCE"
        Me.lblLCE.Size = New System.Drawing.Size(872, 24)
        Me.lblLCE.TabIndex = 265
        Me.lblLCE.Text = "Manager Compliance Events"
        Me.lblLCE.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblLCE.Visible = False
        '
        'lblLCEDisplay
        '
        Me.lblLCEDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLCEDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblLCEDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblLCEDisplay.Name = "lblLCEDisplay"
        Me.lblLCEDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblLCEDisplay.TabIndex = 264
        Me.lblLCEDisplay.Text = "-"
        Me.lblLCEDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblLCEDisplay.Visible = False
        '
        'ugManagerComplianceEvents
        '
        Me.ugManagerComplianceEvents.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugManagerComplianceEvents.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugManagerComplianceEvents.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugManagerComplianceEvents.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugManagerComplianceEvents.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugManagerComplianceEvents.Location = New System.Drawing.Point(0, 0)
        Me.ugManagerComplianceEvents.Name = "ugManagerComplianceEvents"
        Me.ugManagerComplianceEvents.Size = New System.Drawing.Size(888, 24)
        Me.ugManagerComplianceEvents.TabIndex = 27
        '
        'pnlLCE
        '
        Me.pnlLCE.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlLCE.Location = New System.Drawing.Point(0, 968)
        Me.pnlLCE.Name = "pnlLCE"
        Me.pnlLCE.Size = New System.Drawing.Size(888, 200)
        Me.pnlLCE.TabIndex = 1063
        Me.pnlLCE.Visible = False
        '
        'pnlDocumentsHead
        '
        Me.pnlDocumentsHead.Controls.Add(Me.lblDocuments)
        Me.pnlDocumentsHead.Controls.Add(Me.lblDocumentsDisplay)
        Me.pnlDocumentsHead.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlDocumentsHead.Location = New System.Drawing.Point(0, 1168)
        Me.pnlDocumentsHead.Name = "pnlDocumentsHead"
        Me.pnlDocumentsHead.Size = New System.Drawing.Size(888, 24)
        Me.pnlDocumentsHead.TabIndex = 1063
        '
        'lblDocuments
        '
        Me.lblDocuments.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblDocuments.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblDocuments.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblDocuments.Location = New System.Drawing.Point(16, 0)
        Me.lblDocuments.Name = "lblDocuments"
        Me.lblDocuments.Size = New System.Drawing.Size(872, 24)
        Me.lblDocuments.TabIndex = 279
        Me.lblDocuments.Text = "Documents"
        Me.lblDocuments.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblDocuments.Visible = False
        '
        'lblDocumentsDisplay
        '
        Me.lblDocumentsDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblDocumentsDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblDocumentsDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblDocumentsDisplay.Name = "lblDocumentsDisplay"
        Me.lblDocumentsDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblDocumentsDisplay.TabIndex = 278
        Me.lblDocumentsDisplay.Text = "-"
        Me.lblDocumentsDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblDocumentsDisplay.Visible = False
        '
        'pnlDocuments
        '
        Me.pnlDocuments.Controls.Add(Me.UCManagerDocuments)
        Me.pnlDocuments.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlDocuments.Location = New System.Drawing.Point(0, 1192)
        Me.pnlDocuments.Name = "pnlDocuments"
        Me.pnlDocuments.Size = New System.Drawing.Size(888, 200)
        Me.pnlDocuments.TabIndex = 1063
        Me.pnlDocuments.Visible = False
        '
        'UCManagerDocuments
        '
        Me.UCManagerDocuments.AutoScroll = True
        Me.UCManagerDocuments.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UCManagerDocuments.Location = New System.Drawing.Point(0, 0)
        Me.UCManagerDocuments.Name = "UCManagerDocuments"
        Me.UCManagerDocuments.Size = New System.Drawing.Size(888, 200)
        Me.UCManagerDocuments.TabIndex = 278
        '
        'pnlManagerInfo
        '
        Me.pnlManagerInfo.BackColor = System.Drawing.SystemColors.Control
        Me.pnlManagerInfo.Controls.Add(Me.lblRetrainReqDate)
        Me.pnlManagerInfo.Controls.Add(Me.dtRetrainReqDate)
        Me.pnlManagerInfo.Controls.Add(Me.chkIsLicensee)
        Me.pnlManagerInfo.Controls.Add(Me.dtRevokeDate)
        Me.pnlManagerInfo.Controls.Add(Me.LblRevokeDate)
        Me.pnlManagerInfo.Controls.Add(Me.dtRetrainDate3)
        Me.pnlManagerInfo.Controls.Add(Me.LblRetrainDate3)
        Me.pnlManagerInfo.Controls.Add(Me.dtRetrainDate2)
        Me.pnlManagerInfo.Controls.Add(Me.LblRetrainDate2)
        Me.pnlManagerInfo.Controls.Add(Me.dtRetrainDate1)
        Me.pnlManagerInfo.Controls.Add(Me.LblRetrainDate1)
        Me.pnlManagerInfo.Controls.Add(Me.dtInitCertDate)
        Me.pnlManagerInfo.Controls.Add(Me.LblInitCertDate)
        Me.pnlManagerInfo.Controls.Add(Me.cmbInitCertBy)
        Me.pnlManagerInfo.Controls.Add(Me.LblInitCertBy)
        Me.pnlManagerInfo.Controls.Add(Me.lblRestCert)
        Me.pnlManagerInfo.Controls.Add(Me.cmbCMCertificationType)
        Me.pnlManagerInfo.Controls.Add(Me.cmbCMStatus)
        Me.pnlManagerInfo.Controls.Add(Me.pnlCompany)
        Me.pnlManagerInfo.Controls.Add(Me.pbPhoto)
        Me.pnlManagerInfo.Controls.Add(Me.lblPhoto)
        Me.pnlManagerInfo.Controls.Add(Me.pbSign)
        Me.pnlManagerInfo.Controls.Add(Me.lblSignature)
        Me.pnlManagerInfo.Controls.Add(Me.dtExtensionDeadlineDate)
        Me.pnlManagerInfo.Controls.Add(Me.dtExceptionGrantedDate)
        Me.pnlManagerInfo.Controls.Add(Me.lblExceptionGrantedDate)
        Me.pnlManagerInfo.Controls.Add(Me.lblStatus)
        Me.pnlManagerInfo.Controls.Add(Me.lblCertificationType)
        Me.pnlManagerInfo.Controls.Add(Me.pnlExtension)
        Me.pnlManagerInfo.Controls.Add(Me.pnlPhone)
        Me.pnlManagerInfo.Controls.Add(Me.txtManagerAddress)
        Me.pnlManagerInfo.Controls.Add(Me.lblManagerAddress)
        Me.pnlManagerInfo.Controls.Add(Me.btnLabels)
        Me.pnlManagerInfo.Controls.Add(Me.btnEnvelopes)
        Me.pnlManagerInfo.Controls.Add(Me.lblEmail)
        Me.pnlManagerInfo.Controls.Add(Me.txtEmail)
        Me.pnlManagerInfo.Controls.Add(Me.lblSuffix)
        Me.pnlManagerInfo.Controls.Add(Me.lblLastName)
        Me.pnlManagerInfo.Controls.Add(Me.cmbSuffix)
        Me.pnlManagerInfo.Controls.Add(Me.txtLastName)
        Me.pnlManagerInfo.Controls.Add(Me.lblMiddleName)
        Me.pnlManagerInfo.Controls.Add(Me.lblFirstName)
        Me.pnlManagerInfo.Controls.Add(Me.lblPersonTitle)
        Me.pnlManagerInfo.Controls.Add(Me.lblLicenseNumber)
        Me.pnlManagerInfo.Controls.Add(Me.txtLicenseNo)
        Me.pnlManagerInfo.Controls.Add(Me.cmbTitle)
        Me.pnlManagerInfo.Controls.Add(Me.txtFirstName)
        Me.pnlManagerInfo.Controls.Add(Me.txtMiddleName)
        Me.pnlManagerInfo.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlManagerInfo.Location = New System.Drawing.Point(0, 24)
        Me.pnlManagerInfo.Name = "pnlManagerInfo"
        Me.pnlManagerInfo.Size = New System.Drawing.Size(888, 368)
        Me.pnlManagerInfo.TabIndex = 1065
        '
        'lblRetrainReqDate
        '
        Me.lblRetrainReqDate.Location = New System.Drawing.Point(712, 120)
        Me.lblRetrainReqDate.Name = "lblRetrainReqDate"
        Me.lblRetrainReqDate.Size = New System.Drawing.Size(56, 40)
        Me.lblRetrainReqDate.TabIndex = 1078
        Me.lblRetrainReqDate.Text = "Retrain Required By:"
        Me.lblRetrainReqDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dtRetrainReqDate
        '
        Me.dtRetrainReqDate.Checked = False
        Me.dtRetrainReqDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtRetrainReqDate.Location = New System.Drawing.Point(776, 128)
        Me.dtRetrainReqDate.Name = "dtRetrainReqDate"
        Me.dtRetrainReqDate.ShowCheckBox = True
        Me.dtRetrainReqDate.Size = New System.Drawing.Size(104, 20)
        Me.dtRetrainReqDate.TabIndex = 1077
        '
        'chkIsLicensee
        '
        Me.chkIsLicensee.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkIsLicensee.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.chkIsLicensee.Location = New System.Drawing.Point(160, 20)
        Me.chkIsLicensee.Name = "chkIsLicensee"
        Me.chkIsLicensee.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkIsLicensee.Size = New System.Drawing.Size(88, 21)
        Me.chkIsLicensee.TabIndex = 1076
        Me.chkIsLicensee.Tag = "649"
        Me.chkIsLicensee.Text = "Licensee"
        Me.chkIsLicensee.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.chkIsLicensee.Visible = False
        '
        'dtRevokeDate
        '
        Me.dtRevokeDate.Checked = False
        Me.dtRevokeDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtRevokeDate.Location = New System.Drawing.Point(776, 88)
        Me.dtRevokeDate.Name = "dtRevokeDate"
        Me.dtRevokeDate.ShowCheckBox = True
        Me.dtRevokeDate.Size = New System.Drawing.Size(104, 20)
        Me.dtRevokeDate.TabIndex = 1075
        '
        'LblRevokeDate
        '
        Me.LblRevokeDate.Location = New System.Drawing.Point(712, 88)
        Me.LblRevokeDate.Name = "LblRevokeDate"
        Me.LblRevokeDate.Size = New System.Drawing.Size(56, 24)
        Me.LblRevokeDate.TabIndex = 1074
        Me.LblRevokeDate.Text = "Revoke Date:"
        Me.LblRevokeDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dtRetrainDate3
        '
        Me.dtRetrainDate3.Checked = False
        Me.dtRetrainDate3.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtRetrainDate3.Location = New System.Drawing.Point(608, 128)
        Me.dtRetrainDate3.Name = "dtRetrainDate3"
        Me.dtRetrainDate3.ShowCheckBox = True
        Me.dtRetrainDate3.Size = New System.Drawing.Size(104, 20)
        Me.dtRetrainDate3.TabIndex = 1073
        '
        'LblRetrainDate3
        '
        Me.LblRetrainDate3.Location = New System.Drawing.Point(544, 128)
        Me.LblRetrainDate3.Name = "LblRetrainDate3"
        Me.LblRetrainDate3.Size = New System.Drawing.Size(56, 24)
        Me.LblRetrainDate3.TabIndex = 1072
        Me.LblRetrainDate3.Text = "Retrained Date3:"
        Me.LblRetrainDate3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dtRetrainDate2
        '
        Me.dtRetrainDate2.Checked = False
        Me.dtRetrainDate2.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtRetrainDate2.Location = New System.Drawing.Point(416, 128)
        Me.dtRetrainDate2.Name = "dtRetrainDate2"
        Me.dtRetrainDate2.ShowCheckBox = True
        Me.dtRetrainDate2.Size = New System.Drawing.Size(104, 20)
        Me.dtRetrainDate2.TabIndex = 1071
        '
        'LblRetrainDate2
        '
        Me.LblRetrainDate2.Location = New System.Drawing.Point(344, 128)
        Me.LblRetrainDate2.Name = "LblRetrainDate2"
        Me.LblRetrainDate2.Size = New System.Drawing.Size(56, 24)
        Me.LblRetrainDate2.TabIndex = 1070
        Me.LblRetrainDate2.Text = "Retrained Date2:"
        Me.LblRetrainDate2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dtRetrainDate1
        '
        Me.dtRetrainDate1.Checked = False
        Me.dtRetrainDate1.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtRetrainDate1.Location = New System.Drawing.Point(608, 88)
        Me.dtRetrainDate1.Name = "dtRetrainDate1"
        Me.dtRetrainDate1.ShowCheckBox = True
        Me.dtRetrainDate1.Size = New System.Drawing.Size(104, 20)
        Me.dtRetrainDate1.TabIndex = 1069
        '
        'LblRetrainDate1
        '
        Me.LblRetrainDate1.Location = New System.Drawing.Point(544, 88)
        Me.LblRetrainDate1.Name = "LblRetrainDate1"
        Me.LblRetrainDate1.Size = New System.Drawing.Size(56, 24)
        Me.LblRetrainDate1.TabIndex = 1068
        Me.LblRetrainDate1.Text = "Retrained Date1:"
        Me.LblRetrainDate1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dtInitCertDate
        '
        Me.dtInitCertDate.Checked = False
        Me.dtInitCertDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtInitCertDate.Location = New System.Drawing.Point(416, 88)
        Me.dtInitCertDate.Name = "dtInitCertDate"
        Me.dtInitCertDate.ShowCheckBox = True
        Me.dtInitCertDate.Size = New System.Drawing.Size(104, 20)
        Me.dtInitCertDate.TabIndex = 1067
        '
        'LblInitCertDate
        '
        Me.LblInitCertDate.Location = New System.Drawing.Point(312, 88)
        Me.LblInitCertDate.Name = "LblInitCertDate"
        Me.LblInitCertDate.Size = New System.Drawing.Size(96, 24)
        Me.LblInitCertDate.TabIndex = 1066
        Me.LblInitCertDate.Text = "Initial Certification Date:"
        Me.LblInitCertDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbInitCertBy
        '
        Me.cmbInitCertBy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbInitCertBy.Location = New System.Drawing.Point(680, 20)
        Me.cmbInitCertBy.Name = "cmbInitCertBy"
        Me.cmbInitCertBy.Size = New System.Drawing.Size(192, 21)
        Me.cmbInitCertBy.TabIndex = 1065
        '
        'LblInitCertBy
        '
        Me.LblInitCertBy.Location = New System.Drawing.Point(592, 20)
        Me.LblInitCertBy.Name = "LblInitCertBy"
        Me.LblInitCertBy.Size = New System.Drawing.Size(80, 24)
        Me.LblInitCertBy.TabIndex = 1064
        Me.LblInitCertBy.Text = "Initial Certification By:"
        Me.LblInitCertBy.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblRestCert
        '
        Me.lblRestCert.BackColor = System.Drawing.Color.Firebrick
        Me.lblRestCert.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblRestCert.Font = New System.Drawing.Font("Bookman Old Style", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRestCert.Location = New System.Drawing.Point(656, 272)
        Me.lblRestCert.Name = "lblRestCert"
        Me.lblRestCert.Size = New System.Drawing.Size(184, 80)
        Me.lblRestCert.TabIndex = 1061
        Me.lblRestCert.Text = "Restricted Certification"
        Me.lblRestCert.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblRestCert.Visible = False
        '
        'cmbCMCertificationType
        '
        Me.cmbCMCertificationType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbCMCertificationType.Location = New System.Drawing.Point(400, 56)
        Me.cmbCMCertificationType.Name = "cmbCMCertificationType"
        Me.cmbCMCertificationType.Size = New System.Drawing.Size(192, 21)
        Me.cmbCMCertificationType.TabIndex = 276
        '
        'cmbCMStatus
        '
        Me.cmbCMStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbCMStatus.Location = New System.Drawing.Point(400, 20)
        Me.cmbCMStatus.Name = "cmbCMStatus"
        Me.cmbCMStatus.Size = New System.Drawing.Size(192, 21)
        Me.cmbCMStatus.TabIndex = 275
        '
        'pnlCompany
        '
        Me.pnlCompany.Controls.Add(Me.btnCompanyLabels)
        Me.pnlCompany.Controls.Add(Me.btnCompanyEnvelopes)
        Me.pnlCompany.Controls.Add(Me.btnCompanyTitleSearch)
        Me.pnlCompany.Controls.Add(Me.txtCompanyTitle)
        Me.pnlCompany.Controls.Add(Me.txtCompany)
        Me.pnlCompany.Controls.Add(Me.lblCompany)
        Me.pnlCompany.Controls.Add(Me.lblCompanyTitle)
        Me.pnlCompany.Location = New System.Drawing.Point(320, 160)
        Me.pnlCompany.Name = "pnlCompany"
        Me.pnlCompany.Size = New System.Drawing.Size(400, 128)
        Me.pnlCompany.TabIndex = 273
        Me.pnlCompany.Visible = False
        '
        'btnCompanyLabels
        '
        Me.btnCompanyLabels.Location = New System.Drawing.Point(8, 96)
        Me.btnCompanyLabels.Name = "btnCompanyLabels"
        Me.btnCompanyLabels.Size = New System.Drawing.Size(64, 23)
        Me.btnCompanyLabels.TabIndex = 1061
        Me.btnCompanyLabels.Text = "Labels"
        '
        'btnCompanyEnvelopes
        '
        Me.btnCompanyEnvelopes.Location = New System.Drawing.Point(8, 64)
        Me.btnCompanyEnvelopes.Name = "btnCompanyEnvelopes"
        Me.btnCompanyEnvelopes.Size = New System.Drawing.Size(65, 23)
        Me.btnCompanyEnvelopes.TabIndex = 1060
        Me.btnCompanyEnvelopes.Text = "Envelopes"
        '
        'btnCompanyTitleSearch
        '
        Me.btnCompanyTitleSearch.BackColor = System.Drawing.SystemColors.Control
        Me.btnCompanyTitleSearch.Location = New System.Drawing.Point(376, 30)
        Me.btnCompanyTitleSearch.Name = "btnCompanyTitleSearch"
        Me.btnCompanyTitleSearch.Size = New System.Drawing.Size(24, 24)
        Me.btnCompanyTitleSearch.TabIndex = 237
        Me.btnCompanyTitleSearch.Text = "?"
        Me.btnCompanyTitleSearch.Visible = False
        '
        'txtCompanyTitle
        '
        Me.txtCompanyTitle.Location = New System.Drawing.Point(80, 30)
        Me.txtCompanyTitle.Multiline = True
        Me.txtCompanyTitle.Name = "txtCompanyTitle"
        Me.txtCompanyTitle.ReadOnly = True
        Me.txtCompanyTitle.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
        Me.txtCompanyTitle.Size = New System.Drawing.Size(288, 90)
        Me.txtCompanyTitle.TabIndex = 236
        Me.txtCompanyTitle.Text = ""
        Me.txtCompanyTitle.WordWrap = False
        '
        'txtCompany
        '
        Me.txtCompany.Location = New System.Drawing.Point(80, 6)
        Me.txtCompany.Name = "txtCompany"
        Me.txtCompany.ReadOnly = True
        Me.txtCompany.Size = New System.Drawing.Size(288, 20)
        Me.txtCompany.TabIndex = 235
        Me.txtCompany.Text = ""
        '
        'lblCompany
        '
        Me.lblCompany.Location = New System.Drawing.Point(8, 6)
        Me.lblCompany.Name = "lblCompany"
        Me.lblCompany.Size = New System.Drawing.Size(64, 16)
        Me.lblCompany.TabIndex = 239
        Me.lblCompany.Text = "Company:"
        Me.lblCompany.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCompanyTitle
        '
        Me.lblCompanyTitle.Location = New System.Drawing.Point(16, 24)
        Me.lblCompanyTitle.Name = "lblCompanyTitle"
        Me.lblCompanyTitle.Size = New System.Drawing.Size(56, 32)
        Me.lblCompanyTitle.TabIndex = 238
        Me.lblCompanyTitle.Text = "Company Address"
        Me.lblCompanyTitle.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pbPhoto
        '
        Me.pbPhoto.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pbPhoto.Location = New System.Drawing.Point(648, 232)
        Me.pbPhoto.Name = "pbPhoto"
        Me.pbPhoto.Size = New System.Drawing.Size(200, 152)
        Me.pbPhoto.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.pbPhoto.TabIndex = 271
        Me.pbPhoto.TabStop = False
        Me.pbPhoto.Visible = False
        '
        'lblPhoto
        '
        Me.lblPhoto.Location = New System.Drawing.Point(528, 336)
        Me.lblPhoto.Name = "lblPhoto"
        Me.lblPhoto.Size = New System.Drawing.Size(34, 16)
        Me.lblPhoto.TabIndex = 250
        Me.lblPhoto.Text = "Photo"
        Me.lblPhoto.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblPhoto.Visible = False
        '
        'pbSign
        '
        Me.pbSign.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pbSign.Location = New System.Drawing.Point(696, 296)
        Me.pbSign.Name = "pbSign"
        Me.pbSign.Size = New System.Drawing.Size(184, 64)
        Me.pbSign.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.pbSign.TabIndex = 270
        Me.pbSign.TabStop = False
        Me.pbSign.Visible = False
        '
        'lblSignature
        '
        Me.lblSignature.Location = New System.Drawing.Point(336, 336)
        Me.lblSignature.Name = "lblSignature"
        Me.lblSignature.Size = New System.Drawing.Size(53, 16)
        Me.lblSignature.TabIndex = 249
        Me.lblSignature.Text = "Signature"
        Me.lblSignature.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblSignature.Visible = False
        '
        'dtExtensionDeadlineDate
        '
        Me.dtExtensionDeadlineDate.Checked = False
        Me.dtExtensionDeadlineDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtExtensionDeadlineDate.Location = New System.Drawing.Point(752, 328)
        Me.dtExtensionDeadlineDate.Name = "dtExtensionDeadlineDate"
        Me.dtExtensionDeadlineDate.ShowCheckBox = True
        Me.dtExtensionDeadlineDate.Size = New System.Drawing.Size(104, 20)
        Me.dtExtensionDeadlineDate.TabIndex = 21
        Me.dtExtensionDeadlineDate.Visible = False
        '
        'dtExceptionGrantedDate
        '
        Me.dtExceptionGrantedDate.Checked = False
        Me.dtExceptionGrantedDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtExceptionGrantedDate.Location = New System.Drawing.Point(752, 344)
        Me.dtExceptionGrantedDate.Name = "dtExceptionGrantedDate"
        Me.dtExceptionGrantedDate.ShowCheckBox = True
        Me.dtExceptionGrantedDate.Size = New System.Drawing.Size(104, 20)
        Me.dtExceptionGrantedDate.TabIndex = 23
        Me.dtExceptionGrantedDate.Visible = False
        '
        'lblExceptionGrantedDate
        '
        Me.lblExceptionGrantedDate.Location = New System.Drawing.Point(760, 328)
        Me.lblExceptionGrantedDate.Name = "lblExceptionGrantedDate"
        Me.lblExceptionGrantedDate.Size = New System.Drawing.Size(112, 24)
        Me.lblExceptionGrantedDate.TabIndex = 245
        Me.lblExceptionGrantedDate.Text = "Date Manager Requested Extension:"
        Me.lblExceptionGrantedDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblExceptionGrantedDate.Visible = False
        '
        'lblStatus
        '
        Me.lblStatus.Location = New System.Drawing.Point(296, 20)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(96, 16)
        Me.lblStatus.TabIndex = 231
        Me.lblStatus.Text = "Status:"
        Me.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCertificationType
        '
        Me.lblCertificationType.Location = New System.Drawing.Point(280, 56)
        Me.lblCertificationType.Name = "lblCertificationType"
        Me.lblCertificationType.Size = New System.Drawing.Size(112, 16)
        Me.lblCertificationType.TabIndex = 228
        Me.lblCertificationType.Text = "Certification Type:"
        Me.lblCertificationType.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlExtension
        '
        Me.pnlExtension.Location = New System.Drawing.Point(568, 312)
        Me.pnlExtension.Name = "pnlExtension"
        Me.pnlExtension.Size = New System.Drawing.Size(75, 64)
        Me.pnlExtension.TabIndex = 275
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
        Me.pnlPhone.Controls.Add(Me.lblExt2)
        Me.pnlPhone.Controls.Add(Me.lblExt1)
        Me.pnlPhone.Controls.Add(Me.txtExt2)
        Me.pnlPhone.Controls.Add(Me.txtExt1)
        Me.pnlPhone.Location = New System.Drawing.Point(24, 308)
        Me.pnlPhone.Name = "pnlPhone"
        Me.pnlPhone.Size = New System.Drawing.Size(544, 58)
        Me.pnlPhone.TabIndex = 0
        '
        'mskTxtPhone1
        '
        Me.mskTxtPhone1.ContainingControl = Me
        Me.mskTxtPhone1.Location = New System.Drawing.Point(72, 6)
        Me.mskTxtPhone1.Name = "mskTxtPhone1"
        Me.mskTxtPhone1.OcxState = CType(resources.GetObject("mskTxtPhone1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtPhone1.Size = New System.Drawing.Size(142, 20)
        Me.mskTxtPhone1.TabIndex = 8
        '
        'lblPhone2
        '
        Me.lblPhone2.Location = New System.Drawing.Point(10, 35)
        Me.lblPhone2.Name = "lblPhone2"
        Me.lblPhone2.Size = New System.Drawing.Size(56, 16)
        Me.lblPhone2.TabIndex = 228
        Me.lblPhone2.Text = "Phone 2:"
        Me.lblPhone2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'mskTxtCell
        '
        Me.mskTxtCell.ContainingControl = Me
        Me.mskTxtCell.Location = New System.Drawing.Point(384, 35)
        Me.mskTxtCell.Name = "mskTxtCell"
        Me.mskTxtCell.OcxState = CType(resources.GetObject("mskTxtCell.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtCell.Size = New System.Drawing.Size(142, 20)
        Me.mskTxtCell.TabIndex = 13
        '
        'lblCell
        '
        Me.lblCell.Location = New System.Drawing.Point(320, 35)
        Me.lblCell.Name = "lblCell"
        Me.lblCell.Size = New System.Drawing.Size(56, 16)
        Me.lblCell.TabIndex = 230
        Me.lblCell.Text = "Cell:"
        Me.lblCell.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'mskTxtPhone2
        '
        Me.mskTxtPhone2.ContainingControl = Me
        Me.mskTxtPhone2.Location = New System.Drawing.Point(72, 35)
        Me.mskTxtPhone2.Name = "mskTxtPhone2"
        Me.mskTxtPhone2.OcxState = CType(resources.GetObject("mskTxtPhone2.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtPhone2.Size = New System.Drawing.Size(142, 20)
        Me.mskTxtPhone2.TabIndex = 10
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
        Me.mskTxtFax.Location = New System.Drawing.Point(384, 8)
        Me.mskTxtFax.Name = "mskTxtFax"
        Me.mskTxtFax.OcxState = CType(resources.GetObject("mskTxtFax.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtFax.Size = New System.Drawing.Size(142, 20)
        Me.mskTxtFax.TabIndex = 12
        '
        'lblFax
        '
        Me.lblFax.Location = New System.Drawing.Point(320, 8)
        Me.lblFax.Name = "lblFax"
        Me.lblFax.Size = New System.Drawing.Size(56, 16)
        Me.lblFax.TabIndex = 229
        Me.lblFax.Text = "Fax:"
        Me.lblFax.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblExt2
        '
        Me.lblExt2.Location = New System.Drawing.Point(232, 35)
        Me.lblExt2.Name = "lblExt2"
        Me.lblExt2.Size = New System.Drawing.Size(33, 16)
        Me.lblExt2.TabIndex = 236
        Me.lblExt2.Text = "Ext 2:"
        Me.lblExt2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblExt1
        '
        Me.lblExt1.Location = New System.Drawing.Point(232, 8)
        Me.lblExt1.Name = "lblExt1"
        Me.lblExt1.Size = New System.Drawing.Size(33, 16)
        Me.lblExt1.TabIndex = 235
        Me.lblExt1.Text = "Ext 1:"
        Me.lblExt1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtExt2
        '
        Me.txtExt2.Location = New System.Drawing.Point(272, 35)
        Me.txtExt2.Name = "txtExt2"
        Me.txtExt2.Size = New System.Drawing.Size(40, 20)
        Me.txtExt2.TabIndex = 11
        Me.txtExt2.Text = ""
        '
        'txtExt1
        '
        Me.txtExt1.Location = New System.Drawing.Point(272, 8)
        Me.txtExt1.Name = "txtExt1"
        Me.txtExt1.Size = New System.Drawing.Size(40, 20)
        Me.txtExt1.TabIndex = 9
        Me.txtExt1.Text = ""
        '
        'txtManagerAddress
        '
        Me.txtManagerAddress.Location = New System.Drawing.Point(104, 206)
        Me.txtManagerAddress.Multiline = True
        Me.txtManagerAddress.Name = "txtManagerAddress"
        Me.txtManagerAddress.ReadOnly = True
        Me.txtManagerAddress.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
        Me.txtManagerAddress.Size = New System.Drawing.Size(208, 96)
        Me.txtManagerAddress.TabIndex = 7
        Me.txtManagerAddress.Text = ""
        Me.txtManagerAddress.WordWrap = False
        '
        'lblManagerAddress
        '
        Me.lblManagerAddress.Location = New System.Drawing.Point(24, 200)
        Me.lblManagerAddress.Name = "lblManagerAddress"
        Me.lblManagerAddress.Size = New System.Drawing.Size(72, 32)
        Me.lblManagerAddress.TabIndex = 264
        Me.lblManagerAddress.Text = "Manager Address"
        Me.lblManagerAddress.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnLabels
        '
        Me.btnLabels.Location = New System.Drawing.Point(32, 280)
        Me.btnLabels.Name = "btnLabels"
        Me.btnLabels.Size = New System.Drawing.Size(64, 23)
        Me.btnLabels.TabIndex = 1060
        Me.btnLabels.Text = "Labels"
        '
        'btnEnvelopes
        '
        Me.btnEnvelopes.Location = New System.Drawing.Point(32, 248)
        Me.btnEnvelopes.Name = "btnEnvelopes"
        Me.btnEnvelopes.Size = New System.Drawing.Size(65, 23)
        Me.btnEnvelopes.TabIndex = 1059
        Me.btnEnvelopes.Text = "Envelopes"
        '
        'lblEmail
        '
        Me.lblEmail.Location = New System.Drawing.Point(40, 176)
        Me.lblEmail.Name = "lblEmail"
        Me.lblEmail.Size = New System.Drawing.Size(56, 16)
        Me.lblEmail.TabIndex = 222
        Me.lblEmail.Text = "E-mail:"
        Me.lblEmail.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtEmail
        '
        Me.txtEmail.Location = New System.Drawing.Point(104, 176)
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.Size = New System.Drawing.Size(208, 20)
        Me.txtEmail.TabIndex = 5
        Me.txtEmail.Text = ""
        '
        'lblSuffix
        '
        Me.lblSuffix.Location = New System.Drawing.Point(24, 144)
        Me.lblSuffix.Name = "lblSuffix"
        Me.lblSuffix.Size = New System.Drawing.Size(72, 16)
        Me.lblSuffix.TabIndex = 190
        Me.lblSuffix.Text = "Suffix:"
        Me.lblSuffix.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblLastName
        '
        Me.lblLastName.Location = New System.Drawing.Point(16, 120)
        Me.lblLastName.Name = "lblLastName"
        Me.lblLastName.Size = New System.Drawing.Size(80, 16)
        Me.lblLastName.TabIndex = 186
        Me.lblLastName.Text = "Last Name:"
        Me.lblLastName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbSuffix
        '
        Me.cmbSuffix.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbSuffix.ItemHeight = 13
        Me.cmbSuffix.Items.AddRange(New Object() {"", "Jr", "Sr", "I", "II", "III", "IV", "V", "VI"})
        Me.cmbSuffix.Location = New System.Drawing.Point(104, 144)
        Me.cmbSuffix.Name = "cmbSuffix"
        Me.cmbSuffix.Size = New System.Drawing.Size(48, 21)
        Me.cmbSuffix.TabIndex = 4
        '
        'txtLastName
        '
        Me.txtLastName.Location = New System.Drawing.Point(104, 120)
        Me.txtLastName.Name = "txtLastName"
        Me.txtLastName.Size = New System.Drawing.Size(176, 20)
        Me.txtLastName.TabIndex = 3
        Me.txtLastName.Text = ""
        '
        'lblMiddleName
        '
        Me.lblMiddleName.Location = New System.Drawing.Point(16, 96)
        Me.lblMiddleName.Name = "lblMiddleName"
        Me.lblMiddleName.Size = New System.Drawing.Size(80, 16)
        Me.lblMiddleName.TabIndex = 189
        Me.lblMiddleName.Text = "Middle Name:"
        Me.lblMiddleName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblFirstName
        '
        Me.lblFirstName.Location = New System.Drawing.Point(24, 56)
        Me.lblFirstName.Name = "lblFirstName"
        Me.lblFirstName.Size = New System.Drawing.Size(72, 16)
        Me.lblFirstName.TabIndex = 188
        Me.lblFirstName.Text = "First Name:"
        Me.lblFirstName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPersonTitle
        '
        Me.lblPersonTitle.Location = New System.Drawing.Point(56, 22)
        Me.lblPersonTitle.Name = "lblPersonTitle"
        Me.lblPersonTitle.Size = New System.Drawing.Size(40, 16)
        Me.lblPersonTitle.TabIndex = 187
        Me.lblPersonTitle.Text = "Title:"
        Me.lblPersonTitle.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblLicenseNumber
        '
        Me.lblLicenseNumber.Location = New System.Drawing.Point(8, 8)
        Me.lblLicenseNumber.Name = "lblLicenseNumber"
        Me.lblLicenseNumber.Size = New System.Drawing.Size(136, 16)
        Me.lblLicenseNumber.TabIndex = 193
        Me.lblLicenseNumber.Text = "Compliance Manager #:"
        Me.lblLicenseNumber.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblLicenseNumber.Visible = False
        '
        'txtLicenseNo
        '
        Me.txtLicenseNo.Location = New System.Drawing.Point(152, 8)
        Me.txtLicenseNo.Name = "txtLicenseNo"
        Me.txtLicenseNo.ReadOnly = True
        Me.txtLicenseNo.Size = New System.Drawing.Size(96, 20)
        Me.txtLicenseNo.TabIndex = 100
        Me.txtLicenseNo.Text = ""
        Me.txtLicenseNo.Visible = False
        '
        'cmbTitle
        '
        Me.cmbTitle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbTitle.ItemHeight = 13
        Me.cmbTitle.Items.AddRange(New Object() {"", "Mr", "Mrs", "Ms", "Dr", "Sir"})
        Me.cmbTitle.Location = New System.Drawing.Point(104, 20)
        Me.cmbTitle.Name = "cmbTitle"
        Me.cmbTitle.Size = New System.Drawing.Size(48, 21)
        Me.cmbTitle.TabIndex = 0
        '
        'txtFirstName
        '
        Me.txtFirstName.Location = New System.Drawing.Point(104, 56)
        Me.txtFirstName.Name = "txtFirstName"
        Me.txtFirstName.Size = New System.Drawing.Size(176, 20)
        Me.txtFirstName.TabIndex = 1
        Me.txtFirstName.Text = ""
        '
        'txtMiddleName
        '
        Me.txtMiddleName.Location = New System.Drawing.Point(104, 88)
        Me.txtMiddleName.Name = "txtMiddleName"
        Me.txtMiddleName.Size = New System.Drawing.Size(176, 20)
        Me.txtMiddleName.TabIndex = 2
        Me.txtMiddleName.Text = ""
        '
        'pnlManagerInfoHead
        '
        Me.pnlManagerInfoHead.Controls.Add(Me.lblManagerInfo)
        Me.pnlManagerInfoHead.Controls.Add(Me.lblManagerInfoDisplay)
        Me.pnlManagerInfoHead.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlManagerInfoHead.Location = New System.Drawing.Point(0, 0)
        Me.pnlManagerInfoHead.Name = "pnlManagerInfoHead"
        Me.pnlManagerInfoHead.Size = New System.Drawing.Size(888, 24)
        Me.pnlManagerInfoHead.TabIndex = 1064
        '
        'lblManagerInfo
        '
        Me.lblManagerInfo.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblManagerInfo.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblManagerInfo.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblManagerInfo.Location = New System.Drawing.Point(16, 0)
        Me.lblManagerInfo.Name = "lblManagerInfo"
        Me.lblManagerInfo.Size = New System.Drawing.Size(872, 24)
        Me.lblManagerInfo.TabIndex = 1
        Me.lblManagerInfo.Text = "Compliance Manager Info"
        Me.lblManagerInfo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblManagerInfoDisplay
        '
        Me.lblManagerInfoDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblManagerInfoDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblManagerInfoDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblManagerInfoDisplay.Name = "lblManagerInfoDisplay"
        Me.lblManagerInfoDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblManagerInfoDisplay.TabIndex = 0
        Me.lblManagerInfoDisplay.Text = "-"
        Me.lblManagerInfoDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Managers
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(904, 605)
        Me.Controls.Add(Me.pnlManagers)
        Me.Controls.Add(Me.pnlManagersBottom)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.lblLicenseNo)
        Me.Name = "Managers"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Managers"
        Me.pnlManagersBottom.ResumeLayout(False)
        Me.pnlManagers.ResumeLayout(False)
        Me.pnlCoursesHead.ResumeLayout(False)
        Me.pnlCourses.ResumeLayout(False)
        CType(Me.ugRelations, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlCoursesBottom.ResumeLayout(False)
        Me.pnlTestsHead.ResumeLayout(False)
        Me.pnlTestsBottom.ResumeLayout(False)
        CType(Me.ugTests, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlPriorCompaniesHead.ResumeLayout(False)
        Me.pnlPriorCompanies.ResumeLayout(False)
        CType(Me.ugPriorCompanies, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlLCEHead.ResumeLayout(False)
        CType(Me.ugManagerComplianceEvents, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlDocumentsHead.ResumeLayout(False)
        Me.pnlDocuments.ResumeLayout(False)
        Me.pnlManagerInfo.ResumeLayout(False)
        Me.pnlCompany.ResumeLayout(False)
        Me.pnlPhone.ResumeLayout(False)
        CType(Me.mskTxtPhone1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtCell, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtPhone2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtFax, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlManagerInfoHead.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Form Events"
    Private Sub Managers_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Not bolLoading Then bolLoading = True
        Try
            If pManager Is Nothing Then
                pManager = New MUSTER.BusinessLogic.pLicensee
            End If
            If pCompanyManagerAssociation Is Nothing Then
                pCompanyManagerAssociation = New MUSTER.BusinessLogic.pCompanyLicensee
            End If
            '   UIUtilsGen.CreateEmptyFormatDatePicker(dtAppRcvdDate)
            UIUtilsGen.CreateEmptyFormatDatePicker(dtExceptionGrantedDate)
            UIUtilsGen.CreateEmptyFormatDatePicker(dtExtensionDeadlineDate)
            '   UIUtilsGen.CreateEmptyFormatDatePicker(dtManagerxpirationDate)
            '   UIUtilsGen.CreateEmptyFormatDatePicker(dtIssuedDate)
            '   UIUtilsGen.CreateEmptyFormatDatePicker(dtOriginalIssueDate)
            UIUtilsGen.CreateEmptyFormatDatePicker(dtInitCertDate)
            UIUtilsGen.CreateEmptyFormatDatePicker(dtRetrainDate1)
            UIUtilsGen.CreateEmptyFormatDatePicker(dtRetrainDate2)
            UIUtilsGen.CreateEmptyFormatDatePicker(dtRetrainDate3)
            UIUtilsGen.CreateEmptyFormatDatePicker(dtRevokeDate)
            UIUtilsGen.CreateEmptyFormatDatePicker(dtRetrainReqDate)
            txtCompany.Text = strCompanyName
            txtCompanyTitle.Text = strCompanyAddress
            PopulateLicenseeHireStatus()
            PopulateManagerStatus()
            PopulateManagerCertificationType()
            PopulateManagerInitCertBy()
            If nManagerID > 0 Or nManagerID < -100 Then
                Dim oManagerInfo As MUSTER.Info.LicenseeInfo
                oManagerInfo = pManager.Retrieve(nManagerID)
                ManagerInfo = oManagerInfo
                populateManagerInfo(oManagerInfo)
                pManager.pManagerFacRelation.GetAll(nManagerID)
                pManager.pLicenseeCourseTest.GetAll(nManagerID)
                PopulatePriorCompanies()
                PopulateDocuments()
                btnFlags.Enabled = True
                btnComments.Enabled = True
            Else
                btnFlags.Enabled = False
                btnComments.Enabled = False
                If strMode = "ADD" Then
                    ManagerInfo = New MUSTER.Info.LicenseeInfo
                    pManager.Add(ManagerInfo)
                    populateManagerInfo(ManagerInfo, False)
                End If
            End If
            If (strMode = "ADD" Or strMode = "Search" Or strMode = "Company") And ncompAddressID > 0 Then
                pnlCompany.Visible = True
                If txtCompany.Text = String.Empty Then
                    If oCompany Is Nothing Then oCompany = New MUSTER.BusinessLogic.pCompany
                    oCompany.Retrieve(nCompanyID)
                    txtCompany.Text = oCompany.COMPANY_NAME
                End If
                If txtCompanyTitle.Text = String.Empty Then
                    oCompanyAdd.Retrieve(ncompAddressID)
                    txtCompanyTitle.Text = oCompanyAdd.AddressLine1 + IIf(oCompanyAdd.AddressLine1.Trim.Length = 0, "", vbCrLf + oCompanyAdd.AddressLine2) + vbCrLf + oCompanyAdd.City + ", " + oCompanyAdd.State + " " + oCompanyAdd.Zip
                End If
            Else
                pnlCompany.Visible = False
            End If
            displayImages()
            ugRelations.DisplayLayout.ValueLists.Clear()

            ugRelations.DisplayLayout.ValueLists.Add("RelationDesc")
            FillManagerFacilityRelationGrid()

            'Tests
            ugTests.DisplayLayout.ValueLists.Clear()
            ugTests.DisplayLayout.ValueLists.Add("CourseType")
            ugTests.DisplayLayout.ValueLists.Add("Start")
            FillCourseTestGrid()
            'InitializeLetterCount()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            bolLoading = False
        End Try
    End Sub
    Private Sub Managers_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        ' The following line enables all forms to detect the visible form in the MDI container
        '
        If Not bolModelWindow Then
            MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "Company")
        End If
    End Sub
    Private Sub Managers_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.TextChanged
        If Not bolModelWindow Then
            MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "WindowName", Me.Text, "Company")
        End If
    End Sub
    Private Sub Managers_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Cancel()
        dir = Nothing
        finfo = Nothing
        If Not pbPhoto.Image Is Nothing Then
            pbPhoto.Image.Dispose()
            pbPhoto.Image = Nothing
        End If
        If Not pbSign.Image Is Nothing Then
            pbSign.Image.Dispose()
            pbSign.Image = Nothing
        End If
        strPhotoFilePath = String.Empty
        strSignatureFilePath = String.Empty
        System.GC.Collect()
        System.GC.WaitForPendingFinalizers()
    End Sub
#End Region

#Region "UI Support Routines"
    Private Sub populateManagerInfo(ByVal oLicInfo As MUSTER.Info.LicenseeInfo, Optional ByVal loadImages As Boolean = True)
        Dim oLicAddInfo As MUSTER.Info.ComAddressInfo
        Try
            If loadImages Then
                finfo = dir.GetFiles(oLicInfo.LICENSEE_NUMBER.ToString + "*")

                For i As Integer = 0 To UBound(finfo)
                    If finfo(i).Name.IndexOf("-") >= 0 Then
                        strSignatureFilePath = finfo(i).FullName
                    Else
                        strPhotoFilePath = finfo(i).FullName
                    End If
                Next
            End If

            '  dtIssuedDate.Enabled = True
            'dtExtensionDeadlineDate.Enabled = True
            'If oLicInfo.StatusDesc = "NOT CURRENTLY CERTIFIED" Then
            '    chkOverrideExpiration.Enabled = True
            'End If

            If nManagerID > 0 Or nManagerID < -100 Then
                oLicAddInfo = oLicAdd.GetAddressByType(0, 0, nManagerID, 0)
                txtManagerAddress.Text = oLicAddInfo.AddressLine1 + IIf(oLicAddInfo.AddressLine1.Trim.Length = 0, "", vbCrLf + oLicAddInfo.AddressLine2) + vbCrLf + oLicAddInfo.City + ", " + oLicAddInfo.State + " " + oLicAddInfo.Zip
            Else
                oLicAddInfo = New MUSTER.Info.ComAddressInfo
                txtManagerAddress.Text = String.Empty
            End If
            mskTxtPhone1.SelText = oLicAddInfo.Phone1
            mskTxtPhone2.SelText = oLicAddInfo.Phone2
            mskTxtCell.SelText = oLicAddInfo.Cell
            mskTxtFax.SelText = oLicAddInfo.Fax
            txtExt1.Text = oLicAddInfo.Ext1
            txtExt2.Text = oLicAddInfo.Ext2
            PopulateLicenseNumber()
            txtFirstName.Text = oLicInfo.FIRST_NAME
            cmbTitle.Text = oLicInfo.TITLE
            cmbSuffix.Text = oLicInfo.SUFFIX
            txtMiddleName.Text = oLicInfo.MIDDLE_NAME
            txtLastName.Text = oLicInfo.LAST_NAME
            '  cmbHireStatus.Text = oLicInfo.HIRE_STATUS
            strManagerHireStatus = oLicInfo.HIRE_STATUS
            '  chkEmployeeLetter.Checked = oLicInfo.EMPLOYEE_LETTER
            EnableDisableEmployeLetter()
            Me.Text = "Manager - " + txtFirstName.Text + " " + txtLastName.Text
            'chkOverrideExpiration.Checked = oLicInfo.OVERRIDE_EXPIRE
            If oLicInfo.CMSTATUS > 0 Then
                UIUtilsGen.SetComboboxItemByValue(cmbCMStatus, oLicInfo.CMSTATUS)
            Else
                cmbCMStatus.SelectedIndex = 0
                'If cmbStatus.SelectedIndex <> 0 Then cmbStatus.SelectedIndex = 0
            End If
            'cmbStatus.Text = oLicInfo.StatusDesc
            'txtStatus.Text = oLicInfo.STATUS
            If oLicInfo.CMCertType > 0 Then
                UIUtilsGen.SetComboboxItemByValue(cmbCMCertificationType, oLicInfo.CMCertType)
            Else
                cmbCMCertificationType.SelectedIndex = 0
                'If cmbCertificationType.SelectedIndex <> 0 Then cmbCertificationType.SelectedIndex = 0
            End If
            If oLicInfo.INITCERTBY > 0 Then
                UIUtilsGen.SetComboboxItemByValue(cmbInitCertBy, oLicInfo.INITCERTBY)
            Else
                cmbInitCertBy.SelectedIndex = 0
                'If cmbCertificationType.SelectedIndex <> 0 Then cmbCertificationType.SelectedIndex = 0
            End If
            'cmbCertificationType.Text = oLicInfo.CertTypeDesc
            'txtCertificationType.Text = oLicInfo.CERT_TYPE
            '   UIUtilsGen.SetDatePickerValue(dtAppRcvdDate, oLicInfo.APP_RECVD_DATE)
            '   UIUtilsGen.SetDatePickerValue(dtIssuedDate, oLicInfo.ISSUED_DATE)
            '   UIUtilsGen.SetDatePickerValue(dtOriginalIssueDate, oLicInfo.ORIGIN_ISSUED_DATE)
            UIUtilsGen.SetDatePickerValue(dtExtensionDeadlineDate, oLicInfo.EXTENSION_DEADLINE_DATE)
            If oLicInfo.LICENSE_EXPIRE_DATE = "" Then
                '   UIUtilsGen.SetDatePickerValue(dtManagerxpirationDate, CDate("01/01/0001"))
            Else
                '   UIUtilsGen.SetDatePickerValue(dtManagerxpirationDate, oLicInfo.LICENSE_EXPIRE_DATE)
            End If
            If oLicInfo.INITCERTDATE = "" Then
                UIUtilsGen.SetDatePickerValue(dtInitCertDate, CDate("01/01/0001"))
            Else
                UIUtilsGen.SetDatePickerValue(dtInitCertDate, oLicInfo.INITCERTDATE)
            End If
            If oLicInfo.RETRAINDATE1 = "" Then
                UIUtilsGen.SetDatePickerValue(dtRetrainDate1, CDate("01/01/0001"))
            Else
                UIUtilsGen.SetDatePickerValue(dtRetrainDate1, oLicInfo.RETRAINDATE1)
            End If
            If oLicInfo.RETRAINDATE2 = "" Then
                UIUtilsGen.SetDatePickerValue(dtRetrainDate2, CDate("01/01/0001"))
            Else
                UIUtilsGen.SetDatePickerValue(dtRetrainDate2, oLicInfo.RETRAINDATE2)
            End If
            If oLicInfo.RETRAINDATE3 = "" Then
                UIUtilsGen.SetDatePickerValue(dtRetrainDate3, CDate("01/01/0001"))
            Else
                UIUtilsGen.SetDatePickerValue(dtRetrainDate3, oLicInfo.RETRAINDATE3)
            End If
            If oLicInfo.REVOKEDATE = "" Then
                UIUtilsGen.SetDatePickerValue(dtRevokeDate, CDate("01/01/0001"))
            Else
                UIUtilsGen.SetDatePickerValue(dtRevokeDate, oLicInfo.REVOKEDATE)
            End If
            If oLicInfo.RETRAINREQDATE = "" Then
                UIUtilsGen.SetDatePickerValue(dtRetrainReqDate, CDate("01/01/0001"))
            Else
                UIUtilsGen.SetDatePickerValue(dtRetrainReqDate, oLicInfo.RETRAINREQDATE)
            End If
            UIUtilsGen.SetDatePickerValue(dtExceptionGrantedDate, oLicInfo.EXCEPT_GRANT_DATE)
            If (nCompanyID > 0 Or nCompanyID < -100) And strMode = "Search" Then
                Dim currentControl As Control
                Dim myEnumerator As System.Collections.IEnumerator = pnlManagers.Controls.GetEnumerator()
                While myEnumerator.MoveNext()
                    If myEnumerator.Current.GetType.ToString.ToLower = "System.Windows.Forms.TextBox".ToLower Then
                        currentControl = CType(myEnumerator.Current, System.Windows.Forms.TextBox)
                        currentControl.Enabled = False
                    End If
                    If myEnumerator.Current.GetType.ToString.ToLower = "System.Windows.Forms.Button".ToLower Then
                        currentControl = CType(myEnumerator.Current, System.Windows.Forms.Button)
                        currentControl.Enabled = False
                    End If
                    If myEnumerator.Current.GetType.ToString.ToLower = "System.Windows.Forms.DataGrid".ToLower Then
                        currentControl = CType(myEnumerator.Current, System.Windows.Forms.DataGrid)
                        currentControl.Enabled = False
                    End If
                    If myEnumerator.Current.GetType.ToString.ToLower = "System.Windows.Forms.ComboBox".ToLower Then
                        currentControl = CType(myEnumerator.Current, System.Windows.Forms.ComboBox)
                        currentControl.Enabled = False
                    End If
                    If myEnumerator.Current.GetType.ToString.ToLower = "System.Windows.Forms.CheckBox".ToLower Then
                        currentControl = CType(myEnumerator.Current, System.Windows.Forms.CheckBox)
                        currentControl.Enabled = False
                    End If
                    If myEnumerator.Current.GetType.ToString.ToLower = "System.Windows.Forms.DateTimePicker".ToLower Then
                        currentControl = CType(myEnumerator.Current, System.Windows.Forms.DateTimePicker)
                        currentControl.Enabled = False
                    End If
                End While
                ugRelations.Enabled = False
                ugTests.Enabled = False
                ugPriorCompanies.Enabled = False
                ugManagerComplianceEvents.Enabled = False
                pnlCompany.Enabled = False
                btnSaveManager.Enabled = False
                btnCancel.Name = "Close"
                btnComments.Enabled = False
                btnFlags.Enabled = False
                btnDeleteManager.Enabled = False
            End If
            EnableDisableCertifyButton()
            EnableDisableGenExpirationLetterButton()
            EnableDisableGenReminderLetterButton()
            'lblManagerID.Text = oLicInfo.ID
            If (nCompanyID > 0 Or nCompanyID < -100) And (pManager.ID > 0 Or pManager.ID < -100) Then
                CommentsMaintenance(, , True)
            ElseIf loadImages Then
                CommentsMaintenance(, , True, True)
            End If

            If Not bolModelWindow Then
                MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "ManagerID", oLicInfo.ID, "Company")
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub displayImages()
        If strPhotoFilePath <> String.Empty Then
            If System.IO.File.Exists(strPhotoFilePath) Then
                pbPhoto.Image = pbPhoto.Image.FromFile(strPhotoFilePath)
            Else
                strPhotoFilePath = MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_LicenseesImages).ProfileValue + "\NoPhoto.gif"
                pbPhoto.Image = pbPhoto.Image.FromFile(strPhotoFilePath)
            End If
        Else
            pbPhoto.Image = Nothing
        End If

        System.GC.Collect()
        System.GC.WaitForPendingFinalizers()

        If strSignatureFilePath <> String.Empty Then
            If System.IO.File.Exists(strSignatureFilePath) Then
                pbSign.Image = pbSign.Image.FromFile(strSignatureFilePath)
            Else
                strSignatureFilePath = MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_LicenseesImages).ProfileValue + "\NoSignature.gif"
                pbSign.Image = pbSign.Image.FromFile(strSignatureFilePath)
            End If
        Else
            pbSign.Image = Nothing
        End If

        System.GC.Collect()
        System.GC.WaitForPendingFinalizers()
    End Sub
    Private Sub InitImages()
        If dir Is Nothing Then
            dir = New DirectoryInfo(MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_LicenseesImages).ProfileValue)
        End If
        finfo = dir.GetFiles("NoPhoto*")
        If finfo.Length > 0 Then
            strPhotoFilePath = finfo(0).FullName
        Else
            strPhotoFilePath = ""
        End If
        finfo = dir.GetFiles("NoSignature*")
        If finfo.Length > 0 Then
            strSignatureFilePath = finfo(0).FullName
        Else
            strSignatureFilePath = ""
        End If
    End Sub

    Private Sub EnableDisableEmployeLetter()
        ' If cmbHireStatus.Text.EndsWith("Employee") Then
        ' chkEmployeeLetter.Enabled = True
        '  Else
        '      chkEmployeeLetter.Enabled = False
        '   End If
    End Sub
    Private Sub EnableDisableCertifyButton()
        If pManager.STATUS_DESC.ToUpper = "CERTIFIED" Then
            btnCertify.Enabled = True
        Else
            btnCertify.Enabled = False
        End If
    End Sub
    Private Sub EnableDisableGenExpirationLetterButton()
        btnGenExpirationLetter.Enabled = False

        If Date.Compare(IIf(pManager.LICENSE_EXPIRE_DATE = "", CDate("01/01/0001"), pManager.LICENSE_EXPIRE_DATE), Today.Date) <= 0 Then
            btnGenExpirationLetter.Enabled = True
        End If


    End Sub
    Private Sub EnableDisableGenReminderLetterButton()
        btnGenReminderLetter.Enabled = False

        If pManager.STATUS_DESC.ToUpper = "CERTIFIED" Then
            If Date.Compare(IIf(pManager.LICENSE_EXPIRE_DATE = "", CDate("01/01/0001"), pManager.LICENSE_EXPIRE_DATE), DateAdd(DateInterval.Day, -90, Today.Date)) > 0 Then
                btnGenReminderLetter.Enabled = True
            End If
        End If




    End Sub

    Private Sub PopulateLicenseNumber()

        With txtLicenseNo
            .Text = pManager.LICENSEE_NUMBER_PREFIX + pManager.LICENSEE_NUMBER.ToString

            If .Text.ToUpper.IndexOf("RX") > -1 Then
                Me.lblRestCert.Visible = True
            Else
                Me.lblRestCert.Visible = False
            End If
        End With
    End Sub
    Private Sub PopulateLicenseeHireStatus()
        ' cmbHireStatus.DataSource = pManager.GetLicenseeHireStatus(False, True).Tables(0)
        '  cmbHireStatus.DisplayMember = "PROPERTY_NAME"
        '  cmbHireStatus.ValueMember = "PROPERTY_NAME"
    End Sub
    Private Sub PopulateManagerStatus()
        cmbCMStatus.DataSource = pManager.GetManagerStatus(False, True).Tables(0)
        cmbCMStatus.DisplayMember = "PROPERTY_NAME"
        cmbCMStatus.ValueMember = "PROPERTY_ID"
    End Sub
    Private Sub PopulateManagerCertificationType()
        cmbCMCertificationType.DataSource = pManager.GetLicenseeCertificationType(False, True).Tables(0)
        cmbCMCertificationType.DisplayMember = "PROPERTY_NAME"
        cmbCMCertificationType.ValueMember = "PROPERTY_ID"
    End Sub
    Private Sub PopulateManagerInitCertBy()
        cmbInitCertBy.DataSource = pManager.GetManagerInitCertBy(False, True).Tables(0)
        cmbInitCertBy.DisplayMember = "PROPERTY_NAME"
        cmbInitCertBy.ValueMember = "PROPERTY_ID"
    End Sub

    Private Sub Cancel()
        pManager.Reset()
        For Each oComAddInfo In oLicAdd.ColCompanyAddresses.Values
            If oComAddInfo.CompanyId = pManager.ID Then
                oComAddInfo.Reset()
            End If
        Next
        For Each oLicCourseInfo In pManager.pLicenseeCourse.colLicCourse.Values
            If oLicCourseInfo.LicenseeID = pManager.ID Then
                oLicCourseInfo.Reset()
            End If
        Next
        For Each oLicTestInfo In pManager.pLicenseeCourseTest.colLicCourseTest.Values
            If oLicTestInfo.LicenseeID = pManager.ID Then
                oLicTestInfo.Reset()
            End If
        Next
    End Sub
#End Region

#Region "Manager Info"
    Private Sub cmbTitle_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbTitle.SelectedValueChanged
        If bolLoading Then Exit Sub
        pManager.TITLE = cmbTitle.Text
    End Sub
    Private Sub txtFirstName_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFirstName.TextChanged
        If bolLoading Then Exit Sub
        pManager.FIRST_NAME = txtFirstName.Text
    End Sub
    Private Sub txtManagerAddress_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtManagerAddress.TextChanged
        If bolLoading Then Exit Sub
        If Not Me.txtManagerAddress.Tag Is Nothing AndAlso TypeOf Me.txtManagerAddress.Tag Is Integer Then
            nLicAddressID = Me.txtManagerAddress.Tag
            pManager.IsDirty = True
        End If


    End Sub
    Private Sub txtMiddleName_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMiddleName.TextChanged
        If bolLoading Then Exit Sub
        pManager.MIDDLE_NAME = txtMiddleName.Text
    End Sub
    Private Sub txtLastName_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLastName.TextChanged
        If bolLoading Then Exit Sub
        pManager.LAST_NAME = txtLastName.Text
    End Sub
    Private Sub cmbSuffix_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbSuffix.SelectedValueChanged
        If bolLoading Then Exit Sub
        pManager.SUFFIX = cmbSuffix.Text
    End Sub
    Private Sub txtEmail_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEmail.TextChanged
        If bolLoading Then Exit Sub
        pManager.EMAIL_ADDRESS = txtEmail.Text
    End Sub

    Private Sub mskTxtPhone1_KeyUpEvent(ByVal sender As Object, ByVal e As AxMSMask.MaskEdBoxEvents_KeyUpEvent) Handles mskTxtPhone1.KeyUpEvent
        If bolLoading Then Exit Sub
        UIUtilsGen.FillStringObjectValues(oLicAdd.Phone1, mskTxtPhone1.FormattedText.ToString)
    End Sub
    Private Sub txtExt1_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExt1.TextChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.FillStringObjectValues(oLicAdd.Ext1, txtExt1.Text.Trim)
    End Sub
    Private Sub mskTxtPhone2_KeyUpEvent(ByVal sender As Object, ByVal e As AxMSMask.MaskEdBoxEvents_KeyUpEvent) Handles mskTxtPhone2.KeyUpEvent
        If bolLoading Then Exit Sub
        UIUtilsGen.FillStringObjectValues(oLicAdd.Phone2, mskTxtPhone2.FormattedText.ToString)
    End Sub
    Private Sub txtExt2_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtExt2.TextChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.FillStringObjectValues(oLicAdd.Ext2, txtExt2.Text.Trim)
    End Sub
    Private Sub mskTxtFax_KeyUpEvent(ByVal sender As Object, ByVal e As AxMSMask.MaskEdBoxEvents_KeyUpEvent) Handles mskTxtFax.KeyUpEvent
        If bolLoading Then Exit Sub
        UIUtilsGen.FillStringObjectValues(oLicAdd.Fax, mskTxtFax.FormattedText.ToString)
    End Sub
    Private Sub mskTxtCell_KeyUpEvent(ByVal sender As Object, ByVal e As AxMSMask.MaskEdBoxEvents_KeyUpEvent) Handles mskTxtCell.KeyUpEvent
        If bolLoading Then Exit Sub
        UIUtilsGen.FillStringObjectValues(oLicAdd.Cell, mskTxtCell.FormattedText.ToString)
    End Sub

    Private Sub txtManagerAddress_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtManagerAddress.DoubleClick
        Try
            If txtManagerAddress.Text = "" Then
                objAddMaster = New AddressMaster(oLicAdd, nLicAddressID, nManagerID, "Manager", "ADD")
            Else
                objAddMaster = New AddressMaster(oLicAdd, nLicAddressID, nManagerID, "Manager", "MODIFY")
            End If
            Me.Update()
            'LockWindowUpdate(Me.Handle.ToInt64)
            ' objAddMaster.mskTxtPhone1.Enabled = False
            '  objAddMaster.mskTxtPhone2.Enabled = False
            '  objAddMaster.txtExt1.Enabled = False
            ' objAddMaster.txtExt2.Enabled = False
            ' objAddMaster.mskTxtCell.Enabled = False
            ' objAddMaster.mskTxtFax.Enabled = False
            objAddMaster.ShowDialog()
            'LockWindowUpdate(0)
            txtManagerAddress.Tag = oLicAdd.AddressId


            Me.txtManagerAddress.Text = oLicAdd.AddressLine1 + IIf(oLicAdd.AddressLine1.Trim.Length = 0, "", vbCrLf + oLicAdd.AddressLine2) + IIf(oLicAdd.City.Length = 0, "", vbCrLf + oLicAdd.City) + IIf(oLicAdd.State.Length = 0, "", ", " + oLicAdd.State) + IIf(oLicAdd.Zip.Length = 0, "", " " + oLicAdd.Zip)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    '  Private Sub cmbHireStatus_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    If bolLoading Then Exit Sub
    '     pManager.HIRE_STATUS = cmbHireStatus.Text
    '     EnableDisableEmployeLetter()
    '  End Sub
    ' Private Sub chkEmployeeLetter_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '   If bolLoading Then Exit Sub
    '   pManager.EMPLOYEE_LETTER = chkEmployeeLetter.Checked
    '  End Sub
    Private Sub cmbCMStatus_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCMStatus.SelectedIndexChanged
        If Not bolLoading Then
            pManager.CMSTATUS_ID = UIUtilsGen.GetComboBoxValueInt(cmbCMStatus)
            pManager.CMSTATUS_DESC = UIUtilsGen.GetComboBoxText(cmbCMStatus)
        End If


    End Sub
    Private Sub cmbCMCertificationType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCMCertificationType.SelectedIndexChanged
        If Not bolLoading Then
            pManager.CMCERT_TYPE_ID = UIUtilsGen.GetComboBoxValueInt(cmbCMCertificationType)
            pManager.CMCERT_TYPE_DESC = UIUtilsGen.GetComboBoxText(cmbCMCertificationType)
        End If
    End Sub
    Private Sub cmbInitCertBy_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbInitCertBy.SelectedIndexChanged
        If Not bolLoading Then
            pManager.INITCERTBY = UIUtilsGen.GetComboBoxValueInt(cmbInitCertBy)
            pManager.INITCERTBYDESC = UIUtilsGen.GetComboBoxText(cmbInitCertBy)

        End If
    End Sub


    Private Sub dtExtensionDeadlineDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtExtensionDeadlineDate.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(dtExtensionDeadlineDate)
        UIUtilsGen.FillDateobjectValues(pManager.EXTENSION_DEADLINE_DATE, dtExtensionDeadlineDate.Text)


    End Sub

    Private Sub dtInitCertDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtInitCertDate.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(Me.dtInitCertDate)
        UIUtilsGen.FillDateobjectValues(pManager.INITCERTDATE, dtInitCertDate.Text)
    End Sub
    Private Sub dtRetrainDate1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtRetrainDate1.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(Me.dtRetrainDate1)
        UIUtilsGen.FillDateobjectValues(pManager.RETRAINDATE1, dtRetrainDate1.Text)
    End Sub
    Private Sub dtRetrainDate2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtRetrainDate2.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(Me.dtRetrainDate2)
        UIUtilsGen.FillDateobjectValues(pManager.RETRAINDATE2, dtRetrainDate2.Text)
    End Sub
    Private Sub dtRetrainDate3_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtRetrainDate3.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(Me.dtRetrainDate3)
        UIUtilsGen.FillDateobjectValues(pManager.RETRAINDATE3, dtRetrainDate3.Text)
    End Sub
    Private Sub dtRevokeDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtRevokeDate.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(Me.dtRevokeDate)
        UIUtilsGen.FillDateobjectValues(pManager.REVOKEDATE, dtRevokeDate.Text)
    End Sub

    Private Sub dtRetrainReqDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtRetrainReqDate.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(Me.dtRetrainReqDate)
        UIUtilsGen.FillDateobjectValues(pManager.RETRAINREQDATE, dtRetrainReqDate.Text)
    End Sub
    Private Sub dtExceptionGrantedDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtExceptionGrantedDate.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(Me.dtExceptionGrantedDate)
        UIUtilsGen.FillDateobjectValues(pManager.EXCEPT_GRANT_DATE, dtExceptionGrantedDate.Text)


    End Sub
    'Private Sub chkOverrideExpiration_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkOverrideExpiration.CheckedChanged
    '    If bolLoading Then Exit Sub
    '    pManager.OVERRIDE_EXPIRE = chkOverrideExpiration.Checked
    'End Sub

    'Private Sub btnCompanyTitleSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCompanyTitleSearch.Click
    '    Dim tempCompanyAddInfo As MUSTER.Info.ComAddressInfo
    '    Try
    '        objAddresses = New Addresses(oCompanyAdd, "Manager", nAssociationID, pCompanyManagerAssociation, , nCompanyID)
    '        objAddresses.ShowDialog()
    '        If nAssociationID > 0 Or nAssociationID < -100 Then
    '            tempCompanyAddInfo = oCompanyAdd.Retrieve(pCompanyManagerAssociation.ComLicCollection.Item(nAssociationID).ComLicAddressID, False)
    '        Else
    '            tempCompanyAddInfo = oCompanyAdd.ColCompanyAddresses.Item(ncompAddressID)
    '        End If
    '        If Not tempCompanyAddInfo Is Nothing Then
    '            Me.txtCompanyTitle.Text = tempCompanyAddInfo.AddressLine1 + IIf(tempCompanyAddInfo.AddressLine1.Trim.Length = 0, "", vbCrLf + tempCompanyAddInfo.AddressLine2) + vbCrLf + tempCompanyAddInfo.City + ", " + tempCompanyAddInfo.State + " " + tempCompanyAddInfo.Zip
    '        End If
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub objAddresses_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles objAddresses.Closing
    '    Try
    '        ncompAddressID = pCompanyManagerAssociation.ComLicAddressID
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub objAddresses_evtCompLicAssocChanged(ByVal AddressID As Integer) Handles objAddresses.evtCompLicAssocChanged
    '    Try
    '        If nAssociationID > 0 Or nAssociationID < -100 Then
    '            Dim comlicinfo As MUSTER.Info.CompanyLicenseeInfo
    '            For Each comlicinfo In pCompanyManagerAssociation.ComLicCollection.Values
    '                If comlicinfo.ID = nAssociationID Then
    '                    comlicinfo.ComLicAddressID = AddressID
    '                End If
    '            Next
    '        Else
    '            pCompanyManagerAssociation.ComLicAddressID = AddressID
    '        End If
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
#End Region



#Region "ManagerFacilityRelations"
    Private Sub FillManagerFacilityRelationGrid()
        Dim drFac As DataRow
        Try
            pManager.pManagerFacRelation.RelationTable.DefaultView.Sort = "FacilityID"
            ugRelations.DataSource = Nothing
            ugRelations.DataSource = pManager.pManagerFacRelation.RelationTable.DefaultView

            For Each drFac In pManager.pManagerFacRelation.RelationTable.Rows
                strFacilityIDs += IIf(CType(drFac.Item("FacilityID"), String) Is Nothing, String.Empty, CType(drFac.Item("FacilityID"), String)) + ", "
            Next
            ugRelations.DisplayLayout.Bands(0).Columns("MGRFACRELATION_id").Hidden = True
            ugRelations.DisplayLayout.Bands(0).Columns("FacilityID").Hidden = False
            ugRelations.DisplayLayout.Bands(0).Columns("ManagerID").Hidden = True
            ugRelations.DisplayLayout.Bands(0).Columns("RelationDesc").Hidden = True
            ugRelations.DisplayLayout.Bands(0).Columns("Relation").Hidden = False
            ugRelations.DisplayLayout.Bands(0).Columns("Deleted").Hidden = True
            ugRelations.DisplayLayout.Bands(0).Columns("Deleted").CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled

            PopulateMgrFacRelation()
            ugRelations.DisplayLayout.Appearance.BackColor = System.Drawing.Color.White
            ugRelations.DisplayLayout.Bands(0).Columns("FacilityID").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending

        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub PopulateMgrFacRelation()
        Try
            Dim dtRelation As DataTable = pManager.pManagerFacRelation.ListRelations
            Dim drRow As DataRow

            ugRelations.DisplayLayout.ValueLists("RelationDesc").ValueListItems.Clear()
            ugRelations.DisplayLayout.ValueLists("RelationDesc").DropDownListAlignment = Infragistics.Win.DropDownListAlignment.Left
            If Not dtRelation Is Nothing Then
                For Each drRow In dtRelation.Rows
                    ugRelations.DisplayLayout.ValueLists("RelationDesc").ValueListItems.Add(drRow.Item("Relation_ID"), drRow.Item("Relation_Desc").ToString)
                Next
                ugRelations.DisplayLayout.Bands(0).Columns("Relation").ValueList = ugRelations.DisplayLayout.ValueLists("RelationDesc")
                ugRelations.DisplayLayout.Bands(0).Columns("Relation").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate
                ugRelations.DisplayLayout.ValueLists("RelationDesc").DisplayStyle = Infragistics.Win.ValueListDisplayStyle.DisplayText
                ugRelations.DisplayLayout.Bands(0).Columns("Relation").Width = 150
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    'Private Sub PopulateCourseType(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid)
    '    Try
    '        Dim dtCourseType As DataTable = pManager.pLicenseeCourse.ListCourseTypes(True)
    '        Dim drRow As DataRow

    '        ugGrid.DisplayLayout.ValueLists("CourseType").ValueListItems.Clear()
    '        ugGrid.DisplayLayout.ValueLists("CourseType").DropDownListAlignment = Infragistics.Win.DropDownListAlignment.Left

    '        For Each drRow In dtCourseType.Rows
    '            ugGrid.DisplayLayout.ValueLists("CourseType").ValueListItems.Add(drRow.Item("PROPERTY_ID"), drRow.Item("PROPERTY_NAME").ToString)
    '        Next
    '        ugGrid.DisplayLayout.Bands(0).Columns("Type").ValueList = ugGrid.DisplayLayout.ValueLists("CourseType")
    '        ugGrid.DisplayLayout.Bands(0).Columns("Type").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate
    '        ugGrid.DisplayLayout.ValueLists("CourseType").DisplayStyle = Infragistics.Win.ValueListDisplayStyle.DisplayText
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub
    'Private Sub PopulateProvider()
    '    Try
    '        'Dim dtProvider As DataTable = pManager.pLicenseeCourse.ListProviders
    '        'Dim drRow As DataRow

    '        'ugCourses.DisplayLayout.ValueLists("ProviderName").ValueListItems.Clear()
    '        'ugCourses.DisplayLayout.ValueLists("ProviderName").DropDownListAlignment = Infragistics.Win.DropDownListAlignment.Left
    '        'If Not dtProvider Is Nothing Then
    '        '    For Each drRow In dtProvider.Rows
    '        '        ugCourses.DisplayLayout.ValueLists("ProviderName").ValueListItems.Add(drRow.Item("Provider_ID"), drRow.Item("Abbrev").ToString)
    '        '    Next
    '        '    ugCourses.DisplayLayout.Bands(0).Columns("Provider").ValueList = ugCourses.DisplayLayout.ValueLists("ProviderName")
    '        '    ugCourses.DisplayLayout.Bands(0).Columns("Provider").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate
    '        '    ugCourses.DisplayLayout.ValueLists("ProviderName").DisplayStyle = Infragistics.Win.ValueListDisplayStyle.DisplayText
    '        '    ugCourses.DisplayLayout.Bands(0).Columns("Provider").Width = 150
    '        'End If
    '    Catch ex As Exception
    '        Throw ex
    '    End Try

    'End Sub

    Private Sub btnRelationAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRelationAdd.Click
        Try
            If Not bolValidatationFailed Then
                ugRelations.DisplayLayout.Bands(0).AddNew()
                ugRelations.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True
            Else
                MsgBox("Enter Relation Type")
                Exit Sub
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    'Private Sub btnCourseModify_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCourseModify.Click
    '    Try
    '        ugCourses.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    Private Sub btnRelationDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRelationDelete.Click
        Try
            Dim msgResult As MsgBoxResult
            If ugRelations.ActiveRow Is Nothing Then
                MsgBox("Select row to Delete")
                Exit Sub
            End If

            msgResult = MsgBox("Do you want to Delete this row.?", MsgBoxStyle.YesNo, "Manager Facility Relation")
            If msgResult = MsgBoxResult.No Then
                Exit Sub
            Else
                'Remove the Relation Entry.
                If CInt(ugRelations.ActiveRow.Cells("MGRFACRELATION_ID").Text) <= 0 Then
                    pManager.pManagerFacRelation.Remove(Integer.Parse(ugRelations.ActiveRow.Cells("MGRFACRELATION_ID").Text))
                Else
                    pManager.pManagerFacRelation.Retrieve(Integer.Parse(ugRelations.ActiveRow.Cells("MGRFACRELATION_ID").Text))
                    pManager.pManagerFacRelation.Deleted = True

                    pManager.pManagerFacRelation.Save(CType(UIUtilsGen.ModuleID.Company, Integer), MusterContainer.AppUser.UserKey, returnVal, False)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If

                    pManager.pManagerFacRelation.Remove(Integer.Parse(ugRelations.ActiveRow.Cells("MGRFACRELATION_ID").Text))

                End If
                ugRelations.ActiveRow.Delete(False)
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ugRelations_AfterRowInsert(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles ugRelations.AfterRowInsert
        Try
            pManager.pManagerFacRelation.Retrieve(0)
            e.Row.Cells("MGRFACRELATION_ID").Value = pManager.pManagerFacRelation.ID
            Me.ugRelations.ActiveCell = e.Row.Cells("Relation")
            Me.ugRelations.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ToggleCellSel, False, False)
            Me.ugRelations.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode, False, False)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugRelations_AfterRowUpdate(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles ugRelations.AfterRowUpdate
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            ugrow = ugRelations.ActiveRow

            'If pManager.pManagerFacRelation.ID = 0 Then
            '    pManager.pManagerFacRelation.ID = -1
            'End If
            'If ugrow.Cells("MGRFACRELATION_ID").Value Is DBNull.Value Then
            '    ugrow.Cells("MGRFACRELATION_ID").Value = pManager.pManagerFacRelation.ID
            'End If
            pManager.pManagerFacRelation.Retrieve(ugrow.Cells("MGRFACRELATION_ID").Value)

            pManager.pManagerFacRelation.ManagerID = nManagerID
            pManager.pManagerFacRelation.FacilityID = ugrow.Cells("FacilityID").Value
            If ugrow.Cells("Relation").Text <> String.Empty Then
                pManager.pManagerFacRelation.RelationID = ugrow.Cells("Relation").Value
            Else
                pManager.pManagerFacRelation.RelationID = 0
            End If

            'If ugrow.Cells("Provider").Text <> String.Empty Then
            '    pManager.pLicenseeCourse.ProviderID = ugrow.Cells("Provider").Value
            'Else
            '    pManager.pLicenseeCourse.ProviderID = 0
            'End If

            'If ugrow.Cells("Date").Text <> String.Empty Then
            '    pManager.pLicenseeCourse.CourseDate = ugrow.Cells("Date").Value
            'Else
            '    pManager.pLicenseeCourse.CourseDate = dtNull
            'End If

            pManager.pManagerFacRelation.Deleted = ugrow.Cells("Deleted").Value

            pManager.pManagerFacRelation.IsDirty = True
            btnSaveManager.Enabled = True
            'If pManager.pLicenseeCourse.ID <= 0 Then
            '    pManager.pLicenseeCourse.CreatedBy = MusterContainer.AppUser.ID
            'Else
            '    pManager.pLicenseeCourse.ModifiedBy = MusterContainer.AppUser.ID
            'End If

            '  ugRelations.DisplayLayout.Bands(0).Columns("Date").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugRelations_BeforeRowUpdate(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles ugRelations.BeforeRowUpdate
        Try
            If e.Row.Cells("Relation").Text = String.Empty Then
                bolValidatationFailed = True
            Else
                bolValidatationFailed = False
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "Tests"
    Private Sub FillCourseTestGrid()
        Try
            ugTests.DataSource = Nothing
            ugTests.DataSource = pManager.pLicenseeCourseTest.TestTable
            ugTests.DisplayLayout.Bands(0).Columns("TestID").Hidden = True
            ugTests.DisplayLayout.Bands(0).Columns("LicenseeID").Hidden = True
            ugTests.DisplayLayout.Bands(0).Columns("Deleted").Hidden = True
            ugTests.DisplayLayout.Bands(0).Columns("Created By").Hidden = True
            ugTests.DisplayLayout.Bands(0).Columns("Date Created").Hidden = True
            ugTests.DisplayLayout.Bands(0).Columns("Last Edited By").Hidden = True
            ugTests.DisplayLayout.Bands(0).Columns("Date Last Edited").Hidden = True
            ugTests.DisplayLayout.Bands(0).Columns("StartTime").Hidden = True
            ugTests.DisplayLayout.Bands(0).Columns("Deleted").CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled

            ' PopulateCourseType(ugTests)
            'PopulateStartTime()
            ugTests.DisplayLayout.Appearance.BackColor = System.Drawing.Color.White
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    'Private Sub PopulateStartTime()
    '    Try
    '        Dim dtProvider As DataTable = pManager.pLicenseeCourseTest.ListStartTime
    '        Dim drRow As DataRow

    '        ugTests.DisplayLayout.ValueLists("Start").ValueListItems.Clear()
    '        ugTests.DisplayLayout.ValueLists("Start").DropDownListAlignment = Infragistics.Win.DropDownListAlignment.Left

    '        For Each drRow In dtProvider.Rows
    '            ugTests.DisplayLayout.ValueLists("Start").ValueListItems.Add(drRow.Item("PROPERTY_ID"), drRow.Item("PROPERTY_NAME").ToString)
    '        Next
    '        ugTests.DisplayLayout.Bands(0).Columns("StartTime").ValueList = ugTests.DisplayLayout.ValueLists("Start")
    '        ugTests.DisplayLayout.Bands(0).Columns("StartTime").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate
    '        ugTests.DisplayLayout.ValueLists("Start").DisplayStyle = Infragistics.Win.ValueListDisplayStyle.DisplayText

    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub

    Private Sub btnTestAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTestAdd.Click
        Try
            ugTests.DisplayLayout.Bands(0).AddNew()
            ugTests.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    'Private Sub btnTestModify_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTestModify.Click
    '    Try
    '        ugTests.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    Private Sub btnTestDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTestDelete.Click
        Try
            Dim msgResult As MsgBoxResult
            If ugTests.ActiveRow Is Nothing Then
                MsgBox("Select row to Delete")
                Exit Sub
            End If

            msgResult = MsgBox("Do you want to Delete this row.?", MsgBoxStyle.YesNo, "Manager Course Tests")
            If msgResult = MsgBoxResult.No Then
                Exit Sub
            Else
                'Remove the Course Entry.
                If CInt(ugTests.ActiveRow.Cells("TestID").Text) <= 0 Then
                    pManager.pLicenseeCourseTest.Remove(Integer.Parse(ugTests.ActiveRow.Cells("TestID").Text))
                Else
                    pManager.pLicenseeCourseTest.Retrieve(Integer.Parse(ugTests.ActiveRow.Cells("TestID").Text))
                    pManager.pLicenseeCourseTest.Deleted = True
                    pManager.pLicenseeCourseTest.ModifiedBy = MusterContainer.AppUser.ID
                    pManager.pLicenseeCourseTest.Save(CType(UIUtilsGen.ModuleID.Company, Integer), MusterContainer.AppUser.UserKey, returnVal, False)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If


                    pManager.pLicenseeCourseTest.Remove(Integer.Parse(ugTests.ActiveRow.Cells("TestID").Text))

                End If
                ugTests.ActiveRow.Delete(False)
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ugTests_AfterRowInsert(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles ugTests.AfterRowInsert
        Try
            pManager.pLicenseeCourseTest.Retrieve(0)
            e.Row.Cells("TestID").Value = pManager.pLicenseeCourseTest.ID
            Me.ugTests.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ToggleCellSel, False, False)
            Me.ugTests.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode, False, False)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugTests_AfterRowUpdate(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles ugTests.AfterRowUpdate
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow

        Try
            ugrow = ugTests.ActiveRow
            pManager.pLicenseeCourseTest.Retrieve(ugrow.Cells("TestID").Value)

            pManager.pLicenseeCourseTest.LicenseeID = nManagerID

            If ugrow.Cells("Type").Text <> String.Empty Then
                pManager.pLicenseeCourseTest.CourseTypeID = ugrow.Cells("Type").Value
            Else
                pManager.pLicenseeCourse.CourseTypeID = 0
            End If

            'If ugrow.Cells("StartTime").Text <> String.Empty Then
            '    pManager.pLicenseeCourseTest.StartTime = ugrow.Cells("StartTime").Text
            'Else
            '    pManager.pLicenseeCourseTest.StartTime = String.Empty
            'End If

            If ugrow.Cells("Score").Text <> String.Empty Then
                pManager.pLicenseeCourseTest.TestScore = ugrow.Cells("Score").Value
            Else
                pManager.pLicenseeCourseTest.TestScore = 0
            End If

            pManager.pLicenseeCourseTest.TestDate = ugrow.Cells("Date").Value

            If pManager.pLicenseeCourseTest.ID <= 0 Then
                pManager.pLicenseeCourseTest.CreatedBy = MusterContainer.AppUser.ID
            Else
                pManager.pLicenseeCourseTest.ModifiedBy = MusterContainer.AppUser.ID
            End If


            pManager.pLicenseeCourse.Deleted = ugrow.Cells("Deleted").Value


            ugTests.DisplayLayout.Bands(0).Columns("TestID").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            btnTestAdd.Focus()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugTests_BeforeRowUpdate(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles ugTests.BeforeRowUpdate
        Try
            If Not validateManagerTests(e.Row) Then
                e.Cancel = True
                pManager.pLicenseeCourseTest.Remove(pManager.pLicenseeCourseTest.ID)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "Prior Companies"
    Private Sub PopulatePriorCompanies()
        ugPriorCompanies.DataSource = pManager.GetPriorCompanies(nManagerID)
    End Sub
#End Region

#Region "LCE"
#End Region

#Region "Documents"
    Private Sub PopulateDocuments()
        UCManagerDocuments.LoadDocumentsGrid(nManagerID, 0, UIUtilsGen.ModuleID.Company)
    End Sub
#End Region

#Region "UI Events"
    Private Sub btnSaveManager_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveManager.Click
        Try
            Dim msgResult As MsgBoxResult
            If pManager.INITCERTDATE = "12:00:00AM" OrElse pManager.INITCERTDATE = "12:00:00 AM" Then
                pManager.INITCERTDATE = "#12:00:00AM#"
            End If
            If pManager.RETRAINDATE1 = "12:00:00AM" OrElse pManager.RETRAINDATE1 = "12:00:00 AM" Then
                pManager.RETRAINDATE1 = "#12:00:00AM#"
            End If
            If pManager.RETRAINDATE2 = "12:00:00AM" OrElse pManager.RETRAINDATE2 = "12:00:00 AM" Then
                pManager.RETRAINDATE2 = "#12:00:00AM#"
            End If
            If pManager.RETRAINDATE3 = "12:00:00AM" OrElse pManager.RETRAINDATE3 = "12:00:00 AM" Then
                pManager.RETRAINDATE3 = "#12:00:00AM#"
            End If
            If pManager.REVOKEDATE = "12:00:00AM" OrElse pManager.REVOKEDATE = "12:00:00 AM" Then
                pManager.REVOKEDATE = "#12:00:00AM#"
            End If
            If pManager.RETRAINREQDATE = "12:00:00AM" OrElse pManager.RETRAINREQDATE = "12:00:00 AM" Then
                pManager.RETRAINREQDATE = "#12:00:00AM#"
            End If
            '1692 property_id is revoked
            'If (pManager.REVOKEDATE <> "#12:00:00AM#" And pManager.REVOKEDATE <> "#12:00:00 AM#") And pManager.CMSTATUS_ID <> 1692 Then
            '    msgResult = MsgBox("Do you want to change status to revoked?", MsgBoxStyle.YesNo, "Change Status")
            '    If msgResult = MsgBoxResult.Yes Then
            '        '  pManager.CMSTATUS_ID = 1692 'no longer a compliance manager or revoked
            '        Exit Sub
            '    End If
            'End If
            SaveManager()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Cancel()
        Managers_Load(sender, e)
        'dir = Nothing
        'finfo = Nothing
        'If Not pbPhoto Is Nothing Then
        '    If Not pbPhoto.Image Is Nothing Then
        '        pbPhoto.Image.Dispose()
        '    End If
        '    pbPhoto.Image = Nothing
        'End If
        'pbPhoto = Nothing
        'If Not pbSign Is Nothing Then
        '    If Not pbSign.Image Is Nothing Then
        '        pbSign.Image.Dispose()
        '    End If
        '    pbSign.Image = Nothing
        'End If
        'pbSign = Nothing
        'strPhotoFilePath = String.Empty
        'strSignatureFilePath = String.Empty
        'System.GC.Collect()
        'System.GC.WaitForPendingFinalizers()
        'Me.Close()
    End Sub
    Private Sub btnDeleteManager_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteManager.Click
        Dim result As DialogResult
        Try
            result = MessageBox.Show("Are you sure you want to delete this Manager?", "Delete Manager.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)
            If result = DialogResult.Yes Then
                ' The following condition needs to be checked
                ' Manager is not associated with any tank installations or closure events
                If txtCompanyTitle.Text = "" Then
                    pManager.Retrieve(nManagerID)
                    pManager.DELETED = True

                    If pManager.ID <= 0 And pManager.ID > -100 Then
                        pManager.CreatedBy = MusterContainer.AppUser.ID
                    Else
                        pManager.ModifiedBy = MusterContainer.AppUser.ID
                    End If
                    pManager.Save(CType(UIUtilsGen.ModuleID.Company, Integer), MusterContainer.AppUser.UserKey, returnVal)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If

                    MsgBox("Manager is deleted successfully")
                    Me.Close()
                Else
                    MsgBox("Manager cannot be deleted, associated with a company.")
                End If
            Else
                Exit Sub
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnCertify_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCertify.Click
        If pManager.IsDirty Then
            MsgBox("Cannot Certify before saving Manager")
            Exit Sub
        End If
        Try
            GenerateCongLetter()
            GenerateManagerCertificateLetter(pManager.LicenseeInfo)
            GenerateManagerCard(pManager.LicenseeInfo)
        Catch ex As Address.NoAddressException
            MsgBox(ex.Message)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()

        End Try

    End Sub
    Private Sub btnGenExpirationLetter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGenExpirationLetter.Click
        If pManager.IsDirty Then
            MsgBox("Cannot Generate Expiration Letter before saving Manager")
            Exit Sub
        End If
        Dim colParams As New Specialized.NameValueCollection
        colParams = FillParameters()
        If pManager.LicenseeInfo.CertTypeDesc.ToUpper = "INSTALL" Then
            colParams.Add("<Certification Type> ", "Install, Alter, and Permanently Close")
        ElseIf pManager.LicenseeInfo.CertTypeDesc.ToUpper = "CLOSURE" Then
            colParams.Add("<Certification Type>", "Permanently Close")
        Else
            colParams.Add("<Certification Type>", "None")
        End If
        oLetter.GenerateLicenseeLetter(nManagerID, "Licensee Expired", "Expired", "Licensee Expired Letter", "ExpirationLetter.doc", colParams)
        ' per sandy's email dated jun 18, 2009
        ' change status to "not currently certified"
        pManager.STATUS_ID = 1642
        UIUtilsGen.SetComboboxItemByValue(cmbCMStatus, pManager.STATUS_ID)
        pManager.STATUS_DESC = UIUtilsGen.GetComboBoxText(cmbCMStatus)

        pManager.CERT_TYPE_ID = 0
        UIUtilsGen.SetComboboxItemByValue(cmbCMCertificationType, pManager.CERT_TYPE_ID)
        pManager.CERT_TYPE_DESC = UIUtilsGen.GetComboBoxText(cmbCMCertificationType)

        btnSaveManager_Click(sender, e)
    End Sub
    Private Sub btnGenReminderLetter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGenReminderLetter.Click
        If pManager.IsDirty Then
            MsgBox("Cannot Generate Reminder Letter before saving Manager")
            Exit Sub
        End If
        Dim colParams As New Specialized.NameValueCollection
        colParams = FillParameters(True)
        oLetter.GenerateLicenseeLetter(nManagerID, "Info Needed Letter", "InfoNeeded_Letter", "Manager Info Needed Letter", "InfoNeededLetter.doc", colParams)
    End Sub

    Friend Function SaveManager() As Boolean
        Dim LicNo As String = ""
        Dim oCompanyLicenseeInfoLocal As MUSTER.Info.CompanyLicenseeInfo
        Dim bolExist As Boolean = False
        Dim bolsuccess As Boolean = False

        Try
            ' to prevent duplicate letters
            'InitializeLetterCount()

            If ValidateData() Then
                'If Not txtCompany.Text = "" Then
                '    pManager.ManagerLogic(ManagerInfo, FinRespExpirationDate, True, nManagerID, strManagerHireStatus)
                'Else
                '    pManager.ManagerLogic(ManagerInfo, FinRespExpirationDate, , nManagerID, strManagerHireStatus)
                'End If
                'If Not pManager.Add(ManagerInfo) Then
                '    Exit Sub
                'End If

                Dim success As Boolean = False
                Dim msgResult As MsgBoxResult
                'If pManager.REVOKEDATE <> "12:00:00AM" And pManager.REVOKEDATE <> "12:00:00 AM" And pManager.REVOKEDATE <> "#12:00:00AM#" And pManager.REVOKEDATE <> "#12:00:00 AM#" Then
                '    msgResult = MsgBox("Do you want to generate the revoke letter?", MsgBoxStyle.YesNo, "Change Status")
                '    If msgResult = MsgBoxResult.Yes Then
                '        genRevokeLetter()
                '    End If
                'End If
                'have to allow for the - intergers fro migration and the - intergers from collections
                If pManager.ID >= -100 And pManager.ID <= 0 Then
                    pManager.CreatedBy = MusterContainer.AppUser.ID
                Else
                    pManager.ModifiedBy = MusterContainer.AppUser.ID
                End If

                success = pManager.CMSave(UIUtilsGen.ModuleID.Company, MusterContainer.AppUser.UserKey, ncompAddressID, returnVal, False, nCompanyID)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Function
                End If
                If Not success Then
                    bolsuccess = False
                    Exit Function
                End If
                If Not callingForm Is Nothing Then
                    callingForm.Tag = "1"
                End If
                ' begin issue 3198
                If Not btnComments.Enabled Then btnComments.Enabled = True
                ' end issue 3198
                nManagerID = pManager.ID
                'txtCertificationType.Text = pManager.CERT_TYPE_ID
                'txtStatus.Text = pManager.STATUS_ID
                If Not txtCompany.Text = "" Then
                    If nAssociationID > 0 Or nAssociationID <= -100 Then
                        ncompAddressID = pCompanyManagerAssociation.ComLicCollection.Item(nAssociationID).ComLicAddressID
                    End If
                    For Each oCompanyLicenseeInfoLocal In pCompanyManagerAssociation.ComLicCollection.Values
                        If oCompanyLicenseeInfoLocal.LicenseeID = pManager.ID And oCompanyLicenseeInfoLocal.CompanyID = nCompanyID Then
                            oCompanyLicenseeInfoLocal.ComLicAddressID = ncompAddressID
                            oCompanyLicenseeInfoLocal.Deleted = False
                            bolExist = True
                            Exit For
                        End If
                    Next
                    If Not bolExist Then
                        CompanyLicenseeInfo = New MUSTER.Info.CompanyLicenseeInfo(nAssociationID, _
                                                                  nCompanyID, _
                                                                  pManager.ID, _
                                                                  ncompAddressID, _
                                                                  False, _
                                                                  IIf((nAssociationID <= 0 And nAssociationID > -100), MusterContainer.AppUser.ID, ""), _
                                                                    Now, _
                                                                  IIf(nAssociationID > 0, MusterContainer.AppUser.ID, ""), _
                                                                    dtNull)
                        CompanyLicenseeInfo.IsDirty = True
                        pCompanyManagerAssociation.Add(CompanyLicenseeInfo)
                    End If
                End If
                oLicAdd.LicenseeID = pManager.ID

                If oLicAdd.AddressId <= 0 And oLicAdd.AddressId > -100 Then
                    oLicAdd.CreatedBy = MusterContainer.AppUser.ID
                Else
                    oLicAdd.ModifiedBy = MusterContainer.AppUser.ID
                End If

                oLicAdd.Save(CType(UIUtilsGen.ModuleID.Company, Integer), MusterContainer.AppUser.UserKey, returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Function
                End If

                'Fill Course and Test DataGrids.
                FillManagerFacilityRelationGrid()
                FillCourseTestGrid()
                PopulateLicenseNumber()
                populateManagerInfo(ManagerInfo)
                UCManagerDocuments.LoadDocumentsGrid(pManager.ID, 0, 893)
                displayImages()
                bolFormClosing = True

                'If bolCongratsLetter = True Then
                '    If bolRenewalLetter = True Then
                '    Else
                '        GenerateCongLetter()
                '        bolCongratsLetter = False
                '        UIUtilsGen.Delay(, 1)
                '    End If
                'End If
                'If bolManagerCard Then
                '    GenerateManagerCard(ManagerInfo)
                '    bolManagerCard = False
                '    UIUtilsGen.Delay(, 1)
                'End If
                'If bolManagerCertificate Then
                '    GenerateManagerCertificateLetter(ManagerInfo)
                '    bolManagerCertificate = False
                '    UIUtilsGen.Delay(, 1)
                'End If

                MsgBox("Manager is saved successfully")
                bolsuccess = True
            End If
            Return bolsuccess
        Catch ex As Exception
            Throw ex
        Finally
            ' resetting genetared letters count
            'InitializeLetterCount()
        End Try
    End Function
    Public Function ValidateData() As Boolean
        Dim errStr As String = ""
        Dim validateSuccess As Boolean = True
        Try
            If txtLastName.Text = "" Or txtFirstName.Text = "" Then
                errStr += "Manager's First and Last Name" + vbCrLf
                validateSuccess = False
            End If
            If txtCompany.Text <> "" And txtCompanyTitle.Text = "" Then
                MsgBox("Company address")
                Exit Function
            End If
            ' begin issue 3197
            'If cmbHireStatus.Text = "" Then
            '    errStr += "Hire Status" + vbCrLf
            '    validateSuccess = False
            'End If
            'If cmbStatus.Text = "" Then
            '    errStr += "Status" + vbCrLf
            '    validateSuccess = False
            'End If
            'If cmbCertificationType.Text = "" Then
            '    errStr += "Certification Type" + vbCrLf
            '    validateSuccess = False
            'End If
            ' end issue 3197
            If Date.Compare(pManager.EXCEPT_GRANT_DATE, dtNull) <> 0 Then
                If Date.Compare(pManager.EXCEPT_GRANT_DATE, pManager.EXTENSION_DEADLINE_DATE) > 0 Then
                    errStr += "Extension Deadline Date must be greater than Date Manager Requested Extension" + vbCrLf
                End If
            End If
            If txtManagerAddress.Text = "" Then
                errStr += "Invalid Address" + vbCrLf
                validateSuccess = False
            End If
            If errStr.Length > 0 Or Not validateSuccess Then
                MsgBox("The following are required: " + vbCrLf + errStr)
            End If
            Return validateSuccess
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Function validateManagerTests(ByRef ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow) As Boolean
        Dim msg As String = ""
        Dim msgGood As String = ""
        Dim oTestInfo As MUSTER.Info.LicenseeCourseTestInfo
        Dim bolClosure1 As Boolean = False
        Dim closureTestDate1 As DateTime = dtNull
        Dim closureTestDate2 As DateTime = dtNull
        Dim InstallTestDate1 As DateTime = dtNull
        Dim countClosure As Integer = 0
        Dim countInstall As Integer = 0
        Dim bolConsecutiveClosure3 As Boolean = False
        Dim bolConsecutiveInstall3 As Boolean = False
        Dim bolConsecutiveClosure2 As Boolean = False
        Dim bolConsecutiveInstall2 As Boolean = False
        Try
            If ugTests.Rows.Count < 1 Then
                Return True
            End If
            If pManager.pLicenseeCourseTest.colLicCourseTest.Count = 0 Then
                Return True
            End If
            For Each oTestInfo In pManager.pLicenseeCourseTest.colLicCourseTest.Values
                If ugRow.Cells("TestID").Text <> oTestInfo.ID.ToString Then
                    '	The Test Date for a new Install Test (Type = Install), must be greater or equal to the 
                    '   test Date for a Closure Test that has a Score of at least 75. If a Closure Test meeting these 
                    '   conditions does not exist, the Install Test Date is invalid.	
                    If oTestInfo.CourseTypeID = 921 And (Not oTestInfo.TestScore < 75) Then
                        closureTestDate1 = oTestInfo.TestDate
                        bolClosure1 = True
                    End If
                    If oTestInfo.CourseTypeID = 921 And oTestInfo.TestScore < 75 Then
                        If closureTestDate2 = dtNull Then
                            closureTestDate2 = oTestInfo.TestDate
                        Else
                            If Date.Compare(oTestInfo.TestDate, closureTestDate2) > 0 Then
                                closureTestDate2 = oTestInfo.TestDate
                            End If
                        End If
                        countClosure += 1
                    End If
                    If oTestInfo.CourseTypeID = 920 And oTestInfo.TestScore < 75 Then
                        If InstallTestDate1 = dtNull Then
                            InstallTestDate1 = oTestInfo.TestDate
                        Else
                            If Date.Compare(oTestInfo.TestDate, InstallTestDate1) > 0 Then
                                InstallTestDate1 = oTestInfo.TestDate
                            End If
                        End If
                        countInstall += 1
                    End If
                End If
            Next


            If ugRow.Cells("Type").Value = 920 Then  'And oTestInfo.IsDirty Then
                If bolClosure1 Then
                    If Date.Compare(CDate(ugRow.Cells("Date").Value), closureTestDate1) < 0 Then 'Or Date.Compare(oTestInfo.TestDate, closureTestDate1) <> 0 Then
                        msg += "Test date for a new install test must be >= the test date for a closure test that has a score of atleast 75"
                    End If

                ElseIf Not bolClosure1 And Date.Compare(closureTestDate1, dtNull) = 0 Then
                    msg += "Install TestDate is invalid"

                End If
            End If
            If ugRow.Cells("Type").Value = 921 And (countClosure >= 1 And countClosure < 2) Then  'bolConsecutiveClosure2 Then
                If Not closureTestDate2 = dtNull Then
                    If DateDiff(DateInterval.Day, closureTestDate2, CDate(ugRow.Cells("Date").Value)) <= 30 Then
                        msg += "Test date for a new closure test must be 30 days > the most recent one or two consecutive closure tests with score < 75"
                    End If
                End If
            End If
            If ugRow.Cells("Type").Value = 920 And (countInstall >= 1 And countInstall < 2) Then
                If Not InstallTestDate1 = dtNull Then
                    If DateDiff(DateInterval.Day, InstallTestDate1, CDate(ugRow.Cells("Date").Value)) <= 30 Then
                        msg += "Test date for a new Install test must be 30 days > the most recent one or two consecutive Install tests with score < 75"
                    End If
                End If
            End If
            If ugRow.Cells("Type").Value = 921 And countClosure >= 2 Then  'bolConsecutiveClosure3 Then
                If Not closureTestDate2 = dtNull Then
                    ' If DateDiff(DateInterval.Year, closureTestDate2, CDate(ugRow.Cells("Date").Value)) >= 1 Then
                    If Convert.ToInt32(ugRow.Cells("Score").Value) < 75 Then
                        ' good to go - do nothing
                        msgGood += "This applicant must wait one year to re-test because the 3rd test score is < 75."
                    End If
                End If
            End If
            If ugRow.Cells("Type").Value = 920 And countInstall >= 2 Then 'bolConsecutiveInstall3 Then
                If Not InstallTestDate1 = dtNull Then
                    If DateDiff(DateInterval.Year, InstallTestDate1, CDate(ugRow.Cells("Date").Value)) >= 1 Then
                        ' good to go - do nothing
                    Else
                        msg += "Test date for a new Install test must be 1 year  > the most recent 3 or more consecutive Install tests with score < 75"
                    End If
                End If
            End If
            If msg = "" Then
                If msgGood.Length > 0 Then
                    MsgBox(msgGood)
                End If
                Return True
            Else
                MsgBox(msg)
                Return False
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function
#End Region

#Region "Event Handlers for Expanding/Collapsing the different Sections of the Company form"
    Private Sub ExpandCollapse(ByRef pnl As Panel, ByRef lbl As Label, Optional ByVal pnl1 As Panel = Nothing)
        pnl.Visible = Not pnl.Visible
        lbl.Text = IIf(pnl.Visible, "-", "+")
        If Not pnl1 Is Nothing Then pnl1.Visible = Not pnl1.Visible
    End Sub
    Private Sub lblManagerInfoDisplay_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblManagerInfoDisplay.Click
        lblManagerInfo_Click(sender, e)
    End Sub
    Private Sub lblManagerInfo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblManagerInfo.Click
        ExpandCollapse(pnlManagerInfo, lblManagerInfoDisplay)
    End Sub
    Private Sub lblCoursesDisplay_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblCoursesDisplay.Click
        lblCourses_Click(sender, e)
    End Sub
    Private Sub lblCourses_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblRelations.Click
        ExpandCollapse(pnlCourses, lblCoursesDisplay, pnlCoursesBottom)
    End Sub
    Private Sub lblTestsDisplay_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblTestsDisplay.Click
        lblTests_Click(sender, e)
    End Sub
    Private Sub lblTests_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblTests.Click
        ExpandCollapse(pnlTests, lblTestsDisplay, pnlTestsBottom)
    End Sub
    Private Sub lblPriorCompaniesDisplay_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblPriorCompaniesDisplay.Click
        lblPriorCompanies_Click(sender, e)
    End Sub
    Private Sub lblPriorCompanies_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblPriorCompanies.Click
        ExpandCollapse(pnlPriorCompanies, lblPriorCompaniesDisplay)
    End Sub
    Private Sub lblLCEDisplay_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblLCEDisplay.Click
        lblLCE_Click(sender, e)
    End Sub
    Private Sub lblLCE_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblLCE.Click
        ExpandCollapse(pnlLCE, lblLCEDisplay)
    End Sub
    Private Sub lblDocumentsDisplay_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblDocumentsDisplay.Click
        lblDocuments_Click(sender, e)
    End Sub
    Private Sub lblDocuments_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblDocuments.Click
        ExpandCollapse(pnlDocuments, lblDocumentsDisplay)
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
            If pManager.ID >= -100 And pManager.ID <= 0 Then
                MsgBox("Please Save Manager before entering comments")
                Exit Sub
            End If
            nEntityType = UIUtilsGen.EntityTypes.Licensee
            strEntityName = "Manager : " + CStr(pManager.ID) + " " + pManager.FIRST_NAME
            If Not resetBtnColor Then
                SC = New ShowComments(pManager.ID, nEntityType, IIf(bolSetCounts, "", "Manager"), strEntityName, pManager.Comments, Me.Text, , False)
                If bolSetCounts Then
                    nCommentsCount = SC.GetCounts()
                Else
                    SC.ShowDialog()
                    nCommentsCount = IIf(SC.nCommentsCount <= 0, SC.GetCounts(), SC.nCommentsCount)
                End If
            End If
            If nEntityType = UIUtilsGen.EntityTypes.Licensee Then
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
            SF = New ShowFlags(pManager.ID, UIUtilsGen.EntityTypes.Licensee, "Manager")
            SF.ShowDialog()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub SF_FlagAdded(ByVal entityID As Integer, ByVal entityType As Integer, ByVal [Module] As String, ByVal ParentFormText As String) Handles SF.FlagAdded
        RaiseEvent FlagAdded(entityID, entityType, [Module], ParentFormText)
    End Sub
    Private Sub SF_RefreshCalendar() Handles SF.RefreshCalendar
        RaiseEvent RefreshCalendar()
    End Sub
#End Region

#Region "Letter Generation"
    'Private Sub InitializeLetterCount()
    '    slLetterCount = New SortedList
    '    slLetterCount.Add(ManagerLetters.CongLetter, 0)
    '    slLetterCount.Add(ManagerLetters.ManagerCard, 0)
    '    slLetterCount.Add(ManagerLetters.ManagerCertificateLetter, 0)
    '    slLetterCount.Add(ManagerLetters.NoCertificationLetter, 0)
    '    slLetterCount.Add(ManagerLetters.RenewalLetter, 0)
    'End Sub
    Public Sub GenerateCongLetter()
        Dim colParams As New Specialized.NameValueCollection
        Try
            'bolCongratsLetter = True
            colParams = FillParameters()
            oLetter.GenerateLicenseeLetter(nManagerID, "Manager Congratulation Letter", "Manager_Congratulation_Letter", "Manager Congratulation Letter", "CongratLetter.doc", colParams)
            'slLetterCount.Item(ManagerLetters.CongLetter) = 1
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub genRevokeLetter()
        Dim colParams As New Specialized.NameValueCollection
        If pManager.IsDirty Then
            '    MsgBox("Cannot Generate Revoke Letter before saving Manager")
            '    Exit Sub
        End If
        Try
            colParams = FillParameters()
            colParams.Add("<FacilityID>", strFacilityIDs)
            colParams = FillParameters(True)
            oLetter.GenerateLicenseeLetter(nManagerID, "Revoke Letter", "Revoke_Letter", "Manager Revoke Letter", "CMRevokeLetter.doc", colParams)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub GenerateManagerCard(ByVal oManager As MUSTER.Info.LicenseeInfo)
        Dim colParams As New Specialized.NameValueCollection
        Try
            colParams = FillParameters()
            If oManager.CertTypeDesc.ToUpper = "INSTALL" Then
                colParams.Add("<TYPE>", "Install, Alter, and Permanently Close")
            ElseIf oManager.CertTypeDesc.ToUpper = "CLOSURE" Then
                colParams.Add("<TYPE>", "Permanently Close")
            End If
            colParams.Add("<LicenseeID>", pManager.LICENSEE_NUMBER_PREFIX + pManager.LICENSEE_NUMBER.ToString)
            oLetter.GenerateLicenseeLetter(nManagerID, "Licensee Card", "Licensee_Card", "Licensee Card", "LicenseeCard.doc", colParams, strPhotoFilePath, , , strSignatureFilePath)
            'slLetterCount.Item(ManagerLetters.ManagerCard) = 1
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub GenerateManagerCertificateLetter(ByVal oManager As MUSTER.Info.LicenseeInfo)
        Dim colParams As New Specialized.NameValueCollection
        Try
            colParams = FillParameters()
            colParams.Add("<Certification>", oManager.CertTypeDesc)
            If oManager.CertTypeDesc.ToUpper = "INSTALL" Then
                colParams.Add("<TYPE>", "Install, Alter, and Permanently Close")
            ElseIf oManager.CertTypeDesc.ToUpper = "CLOSURE" Then
                colParams.Add("<TYPE>", "Permanently Close")
            End If

            If pManager.LICENSEE_NUMBER_PREFIX = "NRX" Or pManager.LICENSEE_NUMBER_PREFIX = "CRX" Then
                colParams.Add("<Condition>", "Only USTs owned by " + txtCompany.Text)
            ElseIf pManager.LICENSEE_NUMBER_PREFIX = "NHB" Or pManager.LICENSEE_NUMBER_PREFIX = "CHB" Then
                colParams.Add("<Condition>", "Only as an employee of  " + txtCompany.Text)
            ElseIf pManager.LICENSEE_NUMBER_PREFIX = "NHX" Or pManager.LICENSEE_NUMBER_PREFIX = "CHX" Then
                colParams.Add("<Condition>", "None")
            End If
            'If pManager.HIRE_STATUS = "HX - For Hire - Owner" Then
            '    colParams.Add("<Condition>", "Only USTs owned by " + txtCompany.Text)
            'ElseIf pManager.HIRE_STATUS = "HB - For Hire - Employee" Then
            '    colParams.Add("<Condition>", "Only as an employee of  " + txtCompany.Text)
            'End If
            colParams.Add("<LicenseeID>", pManager.LICENSEE_NUMBER_PREFIX + pManager.LICENSEE_NUMBER.ToString)
            oLetter.GenerateLicenseeLetter(nManagerID, "Licensee Certification Letter", "Licensee_Certification_Letter", "Licensee Certification Letter", "CertificationLetter.doc", colParams)
            'slLetterCount.Item(LicenseeLetters.LicenseeCertificateLetter) = 1
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    'Private Sub pManager_GenerateCongLetter() Handles pManager.GenerateCongLetter
    '    If slLetterCount Is Nothing Then
    '        ' initialize sorted list to maintain the letter count to prevent creating multiple letters
    '        InitializeLetterCount()
    '    End If
    '    If CType(slLetterCount.Item(ManagerLetters.CongLetter), Integer) > 0 Then
    '        Exit Sub
    '    End If
    '    bolCongratsLetter = True
    'End Sub
    'Private Sub pManager_GenerateManagerCard() Handles pManager.GenerateManagerCard
    '    If slLetterCount Is Nothing Then
    '        ' initialize sorted list to maintain the letter count to prevent creating multiple letters
    '        InitializeLetterCount()
    '    End If
    '    If CType(slLetterCount.Item(ManagerLetters.ManagerCard), Integer) > 0 Then
    '        Exit Sub
    '    End If
    '    bolManagerCard = True
    'End Sub
    'Private Sub pManager_GenerateManagerCertLetter() Handles pManager.GenerateManagerCertLetter
    '    If slLetterCount Is Nothing Then
    '        ' initialize sorted list to maintain the letter count to prevent creating multiple letters
    '        InitializeLetterCount()
    '    End If
    '    If CType(slLetterCount.Item(ManagerLetters.ManagerCertificateLetter), Integer) > 0 Then
    '        Exit Sub
    '    End If
    '    bolManagerCertificate = True
    'End Sub
    'Private Sub pManager_GenerateNoCertificationLetter(ByVal oManager As MUSTER.Info.LicenseeInfo) Handles pManager.GenerateNoCertificationLetter
    '    Dim colParams As New Specialized.NameValueCollection
    '    Try
    '        colParams = FillParameters()
    '        If oManager.CertTypeDesc = "INSTALL" Then
    '            colParams.Add("<Certification Type> ", "Install, Alter, and Permanently Close")
    '        ElseIf oManager.CertTypeDesc = "CLOSURE" Then
    '            colParams.Add("<Certification Type>", "Permanently Close")
    '        Else
    '            colParams.Add("<Certification Type>", "None")
    '        End If
    '        oLetter.GenerateManagerLetter(nManagerID, "Manager No Certification Letter", "Manager_NoCertification_Letter", "Manager No Certification Letter", "NoLongerCertifiedLetter.doc", colParams)
    '        slLetterCount.Item(ManagerLetters.NoCertificationLetter) = 1
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub pManager_GenerateManagerRenewalLetter() Handles pManager.GenerateManagerRenewalLetter
    '    Dim colParams As New Specialized.NameValueCollection
    '    Try
    '        If slLetterCount Is Nothing Then
    '            ' initialize sorted list to maintain the letter count to prevent creating multiple letters
    '            InitializeLetterCount()
    '        End If
    '        If CType(slLetterCount.Item(ManagerLetters.RenewalLetter), Integer) > 0 Then
    '            Exit Sub
    '        End If
    '        If bolRenewalLetter = True Then
    '            Exit Sub
    '        End If

    '        colParams = FillParameters()
    '        oLetter.GenerateManagerLetter(nManagerID, "Renewal Certificate Letter", "Renewal_Certificate", "Manager Renewal Certificate for Company", "ManagerRenewalLetter.doc", colParams)
    '        slLetterCount.Item(ManagerLetters.RenewalLetter) = 1
    '        bolRenewalLetter = True

    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub pManager_GenerateNoCertificationLetterOption(ByVal oManager As MUSTER.Info.LicenseeInfo) Handles pManager.GenerateNoCertificationLetterOption
    '    Dim result As DialogResult
    '    Try
    '        If slLetterCount Is Nothing Then
    '            ' initialize sorted list to maintain the letter count to prevent creating multiple letters
    '            InitializeLetterCount()
    '        End If
    '        If CType(slLetterCount.Item(ManagerLetters.NoCertificationLetter), Integer) > 0 Then
    '            Exit Sub
    '        End If
    '        result = MessageBox.Show("Do you want to generate the NOT CERTIFIED LETTER?", "Generate Not Certified Letter.", MessageBoxButtons.YesNo)
    '        If result = DialogResult.Yes Then
    '            pManager_GenerateNoCertificationLetter(oManager)
    '        End If
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub

    Private Function FillParameters(Optional ByVal fillInfoNeededInfo As Boolean = False) As Specialized.NameValueCollection
        Try
            Dim colParams As New Specialized.NameValueCollection
            Dim expStr As String
            Dim oLicAddInfo As MUSTER.Info.ComAddressInfo
            'Build NameValueCollection with Tags and Values.
            colParams.Add("<Date>", Format(Now, "MMMM dd, yyyy"))
            expStr = pManager.LICENSE_EXPIRE_DATE
            fillInfoNeededInfo = False
            If expStr.Length > 10 Then
                expStr = expStr.Substring(5, 2) + "/" + expStr.Substring(8, 2) + "/" + expStr.Substring(0, 4)
            End If
            colParams.Add("<Expiration Date>", expStr)
            colParams.Add("<Manager Name>", pManager.FullName)
            colParams.Add("<Company Name>", txtCompany.Text)

            If nManagerID > 0 Or nManagerID < -100 Then
                oLicAddInfo = oLicAdd.GetAddressByType(0, 0, nManagerID, 0)
                colParams.Add("<Manager Address>", oLicAddInfo.AddressLine1 + IIf(oLicAddInfo.AddressLine2.Trim.Length = 0, "", vbCrLf + oLicAddInfo.AddressLine2) + vbCrLf + oLicAddInfo.City + ", " + oLicAddInfo.State + " " + oLicAddInfo.Zip)

                colParams.Add("<Company Address1>", oCompanyAdd.AddressLine1)

              
                colParams.Add("<Manager Greeting>", "Dear " + pManager.FullName)
                colParams.Add("<User>", MusterContainer.AppUser.Name)
                ' colParams.Add("<User Phone>", MusterContainer.AppUser.PhoneNumber)

                If pManager.CERT_TYPE_DESC.ToUpper = "CLOSURE" Then
                    colParams.Add("<Certification Type>", "Permanently Close ")
                ElseIf pManager.CERT_TYPE_DESC.ToUpper = "INSTALL" Then
                    colParams.Add("<Certification Type> ", "Install, Alter, and Permanently Close ")
                End If

                If fillInfoNeededInfo Then
                    Dim strInfoNeeded As String = String.Empty
                    Dim count As Integer = 0

                    If pManager.LICENSEE_NUMBER_PREFIX = "CHB" Or _
                        pManager.LICENSEE_NUMBER_PREFIX = "CRX" Or _
                        pManager.LICENSEE_NUMBER_PREFIX = "CHX" Or _
                        pManager.LICENSEE_NUMBER_PREFIX = "NHB" Or _
                        pManager.LICENSEE_NUMBER_PREFIX = "NRX" Or _
                        pManager.LICENSEE_NUMBER_PREFIX = "NHX" Then
                        If Date.Compare(pManager.APP_RECVD_DATE, dtNull) <> 0 Then
                            If Date.Compare(pManager.ISSUED_DATE, dtNull) <> 0 Then
                                If Date.Compare(pManager.APP_RECVD_DATE, pManager.ISSUED_DATE) < 0 Then
                                    count += 1
                                    strInfoNeeded += count.ToString + ". A completed certification renewal application" + vbCrLf
                                End If
                            End If
                        End If
                    End If

                    Dim bolClosure As Boolean = True
                    Dim bolInstall As Boolean = True

                    If Date.Compare(pManager.ISSUED_DATE, dtNull) = 0 Then
                        bolClosure = False
                        bolInstall = False
                    Else
                        For Each oLicCourseInfo In pManager.pLicenseeCourse.colLicCourse.Values
                            If oLicCourseInfo.LicenseeID = pManager.ID Then
                                If Date.Compare(oLicCourseInfo.CourseDate, pManager.ISSUED_DATE) > 0 Then
                                    If oLicCourseInfo.CourseTypeID = UIUtilsGen.LicenseeCourseType.Closure Then
                                        bolClosure = False
                                    ElseIf oLicCourseInfo.CourseTypeID = UIUtilsGen.LicenseeCourseType.Install Then
                                        bolInstall = False
                                    End If
                                End If
                            End If
                        Next
                    End If

                    If bolClosure Then
                        If pManager.LICENSEE_NUMBER_PREFIX = "CHB" Or _
                            pManager.LICENSEE_NUMBER_PREFIX = "CRX" Or _
                            pManager.LICENSEE_NUMBER_PREFIX = "CHX" Or _
                            pManager.LICENSEE_NUMBER_PREFIX = "NHB" Or _
                            pManager.LICENSEE_NUMBER_PREFIX = "NRX" Or _
                            pManager.LICENSEE_NUMBER_PREFIX = "NHX" Then
                            count += 1
                            strInfoNeeded += count.ToString + ". A Closure course completion certificate" + vbCrLf
                        End If
                    End If

                    If bolInstall Then
                        If pManager.LICENSEE_NUMBER_PREFIX = "NHB" Or _
                            pManager.LICENSEE_NUMBER_PREFIX = "NRX" Or _
                            pManager.LICENSEE_NUMBER_PREFIX = "NHX" Then
                            count += 1
                            strInfoNeeded += count.ToString + ". An Install course completion certificate" + vbCrLf
                        End If
                    End If

                    If pManager.LICENSEE_NUMBER_PREFIX = "CHB" Or _
                        pManager.LICENSEE_NUMBER_PREFIX = "CRB" Or _
                        pManager.LICENSEE_NUMBER_PREFIX = "NHB" Or _
                        pManager.LICENSEE_NUMBER_PREFIX = "NRB" Then
                        If pManager.HIRE_STATUS <> String.Empty Then
                            If pManager.HIRE_STATUS = "HB - For Hire - Employee" Then
                                If Not pManager.EMPLOYEE_LETTER Then
                                    count += 1
                                    strInfoNeeded += count.ToString + ". An Employee letter stating you are a full time employee of company" + vbCrLf
                                End If
                            End If
                        End If
                    End If

                    If pManager.LICENSEE_NUMBER_PREFIX = "CHB" Or _
                        pManager.LICENSEE_NUMBER_PREFIX = "CHX" Or _
                        pManager.LICENSEE_NUMBER_PREFIX = "NHB" Or _
                        pManager.LICENSEE_NUMBER_PREFIX = "NHX" Then
                        If Date.Compare(FinRespExpirationDate, dtNull) <> 0 Then
                            If Date.Compare(pManager.LICENSE_EXPIRE_DATE, dtNull) <> 0 Then
                                If Date.Compare(FinRespExpirationDate, pManager.LICENSE_EXPIRE_DATE) < 0 Then
                                    count += 1
                                    strInfoNeeded += count.ToString + ". Also if you work for hire (you have an ''H'' in your certification number), you must comply with the financial responsibility requirements in one of the following ways: " + vbCrLf
                                    strInfoNeeded += "      (1) submit to MDEQ a copy of an insurance certificate indicating that your company has atleast $50,000 of contractor's general liability insurance. This insurance certificate must list MDEQ as the certificate"
                                    strInfoNeeded += "holder and have 30 or 60 - days cancellation notice." + vbCrLf + "           OR" + vbCrLf
                                    strInfoNeeded += "       (2) submit to MDEQ a copy of your company's certificate of responsibility from the Mississippi Board of Contractors."

                                End If
                            End If
                        End If
                    End If
                    colParams.Add("<InfoNeeded>", strInfoNeeded)
                End If

                Return colParams
            Else
                Throw New Address.NoAddressException

            End If


        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "Event Handlers for all the controls"
    Public Sub setSaveManager(ByVal bolState As Boolean)
        If Not bolLoading Then
            btnSaveManager.Enabled = bolState Or pManager.IsDirty
        End If

    End Sub

    Public Sub ValidationErrors(ByVal MsgStr As String) Handles pManager.LicenseeErr
        MsgBox(MsgStr)
    End Sub
    Private Sub pManager_ManagerChanged(ByVal bolValue As Boolean) Handles pManager.LicenseeChanged
        setSaveManager(bolValue)
    End Sub
    Private Sub pManager_ManagerCourseChanged(ByVal bolValue As Boolean) Handles pManager.LicenseeCourseChanged
        setSaveManager(bolValue)
    End Sub
    Private Sub pManager_ManagerCourseTestsChanged(ByVal bolValue As Boolean) Handles pManager.LicenseeCourseTestsChanged
        setSaveManager(bolValue)
    End Sub
    Private Sub pManager_ColChanged(ByVal bolValue As Boolean) Handles pManager.ColChanged
        setSaveManager(bolValue)
    End Sub

    Private Sub oLicAdd_evtAddressesChanged(ByVal bolValue As Boolean) Handles oLicAdd.evtAddressesChanged
        setSaveManager(bolValue)
    End Sub
    Private Sub oLicAdd_AddressChanged(ByVal bolValue As Boolean) Handles oLicAdd.evtAddressChanged
        setSaveManager(bolValue)
    End Sub
    Private Sub oLicAdd_evtAddressChanged(ByVal bolValue As Boolean) Handles oLicAdd.evtAddressChanged
        setSaveManager(bolValue)
    End Sub

    Private Sub oCompanyAdd_evtAddressesChanged(ByVal bolValue As Boolean) Handles oCompanyAdd.evtAddressesChanged
        setSaveManager(bolValue)
    End Sub
    Private Sub oCompanyAdd_evtAddressChanged(ByVal bolValue As Boolean) Handles oCompanyAdd.evtAddressChanged
        setSaveManager(bolValue)
    End Sub

    Private Sub pCompanyManagerAssociation_CompanyManagerChanged(ByVal bolValue As Boolean) Handles pCompanyManagerAssociation.CompanyLicenseeChanged
        setSaveManager(bolValue)
    End Sub
    Private Sub pCompanyManagerAssociation_ColChanged(ByVal bolValue As Boolean) Handles pCompanyManagerAssociation.ColChanged
        setSaveManager(bolValue)
    End Sub
#End Region

#Region "Envelopes and Labels"
    Private Sub btnEnvelopes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEnvelopes.Click
        Dim strAddress As String
        Dim strName As String = String.Empty
        Dim arrAddress(4) As String
        Try
            strName = IIf(Me.cmbTitle.Text <> String.Empty, cmbTitle.Text + " ", "") + Me.txtFirstName.Text + " " + IIf(Me.txtMiddleName.Text <> String.Empty, Me.txtMiddleName.Text + " ", "") + Me.txtLastName.Text + IIf(Me.cmbSuffix.Text <> String.Empty, " " + Me.cmbSuffix.Text, "")
            'strAddress = oLicAdd.AddressLine1 + "," + oLicAdd.AddressLine2 + "," + oLicAdd.City + "," + oLicAdd.State + "," + oLicAdd.Zip
            arrAddress(0) = oLicAdd.AddressLine1
            arrAddress(1) = oLicAdd.AddressLine2
            arrAddress(2) = oLicAdd.City
            arrAddress(3) = oLicAdd.State
            arrAddress(4) = oLicAdd.Zip
            If (oLicAdd.AddressId > 0 Or oLicAdd.AddressId < -100) And (pManager.ID > 0 Or pManager.ID < -100) Then
                UIUtilsGen.CreateEnvelopes(strName, arrAddress, "COM", pManager.ID)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnLabels_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLabels.Click
        Dim strAddress As String = String.Empty
        Dim strName As String = String.Empty
        Dim arrAddress(4) As String
        Try
            strName = IIf(Me.cmbTitle.Text <> String.Empty, cmbTitle.Text + " ", "") + Me.txtFirstName.Text + " " + IIf(Me.txtMiddleName.Text <> String.Empty, Me.txtMiddleName.Text + " ", "") + Me.txtLastName.Text + IIf(Me.cmbSuffix.Text <> String.Empty, " " + Me.cmbSuffix.Text, "")
            'strAddress = oLicAdd.AddressLine1 + "," + oLicAdd.AddressLine2 + "," + oLicAdd.City + "," + oLicAdd.State + "," + oLicAdd.Zip
            arrAddress(0) = oLicAdd.AddressLine1
            arrAddress(1) = oLicAdd.AddressLine2
            arrAddress(2) = oLicAdd.City
            arrAddress(3) = oLicAdd.State
            arrAddress(4) = oLicAdd.Zip
            If (oLicAdd.AddressId > 0 Or oLicAdd.AddressId < -100) And (pManager.ID > 0 Or pManager.ID < -100) Then
                UIUtilsGen.CreateLabels(strName, arrAddress, "COM", pManager.ID)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnCompanyEnvelopes_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCompanyEnvelopes.Click
        Dim strAddress As String
        Dim strName As String = String.Empty
        Dim arrAddress(4) As String
        Try
            strName = txtCompany.Text
            If ncompAddressID > 0 Then
                oComAddInfo = oCompanyAdd.Retrieve(ncompAddressID, False)
            Else
                MsgBox("Invalid Address")
                Exit Sub
            End If
            arrAddress(0) = oComAddInfo.AddressLine1
            arrAddress(1) = oComAddInfo.AddressLine2
            arrAddress(2) = oComAddInfo.City
            arrAddress(3) = oComAddInfo.State
            arrAddress(4) = oComAddInfo.Zip
            If (oComAddInfo.AddressId > 0 Or oComAddInfo.AddressId < -100) And (pManager.ID > 0 Or pManager.ID < -100) Then
                UIUtilsGen.CreateEnvelopes(strName, arrAddress, "COM", pManager.ID)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCompanyLabels_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCompanyLabels.Click
        Dim strAddress As String = String.Empty
        Dim strName As String = String.Empty
        Dim arrAddress(4) As String
        Try
            strName = txtCompany.Text
            If ncompAddressID > 0 Then
                oComAddInfo = oCompanyAdd.Retrieve(ncompAddressID, False)
            Else
                MsgBox("Invalid Address")
                Exit Sub
            End If
            arrAddress(0) = oComAddInfo.AddressLine1
            arrAddress(1) = oComAddInfo.AddressLine2
            arrAddress(2) = oComAddInfo.City
            arrAddress(3) = oComAddInfo.State
            arrAddress(4) = oComAddInfo.Zip
            If (oComAddInfo.AddressId > 0 Or oComAddInfo.AddressId < -100) And (pManager.ID > 0 Or pManager.ID < -100) Then
                UIUtilsGen.CreateLabels(strName, arrAddress, "COM", pManager.ID)
            Else
                MsgBox("Invalid Address")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region


 

    Private Sub dtRevokeDate_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtRevokeDate.Enter
        Dim msgResult As MsgBoxResult
        If dtRevokeDate.Checked And (pManager.REVOKEDATE = "#12:00:00AM#" Or pManager.REVOKEDATE = "#12:00:00 AM#" Or pManager.REVOKEDATE = "12:00:00AM" Or pManager.REVOKEDATE = "12:00:00 AM") And pManager.CMSTATUS_ID <> 1692 Then
            ' If dtRevokeDate.Checked And (dtRevokeDate.Value.ToString <> "#12:00:00AM#" And dtRevokeDate.Value.ToString <> "#12:00:00 AM#") And pManager.CMSTATUS_ID <> 1692 Then
            msgResult = MsgBox("Do you want to change status to revoked?", MsgBoxStyle.YesNo, "Change Status")
            ' If msgResult = MsgBoxResult.Yes Then
            '  pManager.CMSTATUS_ID = 1692 'no longer a compliance manager or revoked
            '  If pManager.REVOKEDATE <> "12:00:00AM" And pManager.REVOKEDATE <> "12:00:00 AM" And pManager.REVOKEDATE <> "#12:00:00AM#" And pManager.REVOKEDATE <> "#12:00:00 AM#" Then
        End If
        If dtRevokeDate.Checked And (pManager.REVOKEDATE = "#12:00:00AM#" Or pManager.REVOKEDATE = "#12:00:00 AM#" Or pManager.REVOKEDATE = "12:00:00AM" Or pManager.REVOKEDATE = "12:00:00 AM") Then
            msgResult = MsgBox("Do you want to generate the revoke letter?", MsgBoxStyle.YesNo, "Revoke Letter")
            If msgResult = MsgBoxResult.Yes Then
                genRevokeLetter()
            End If

        End If
    End Sub
End Class
