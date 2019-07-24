Imports System.IO
Imports System.Text

Public Class Licensees
    Inherits System.Windows.Forms.Form

#Region " User Defined Variables"
    Public MyGuid As New System.Guid
    Public Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long

    Friend WithEvents objAddMaster As AddressMaster
    Friend WithEvents objAddresses As Addresses

    ' Company
    Private WithEvents oCompany As MUSTER.BusinessLogic.pCompany

    ' Company Address
    Private WithEvents oCompanyAdd As New MUSTER.BusinessLogic.pComAddress
    Private oComAddInfo As MUSTER.Info.ComAddressInfo

    ' Comany Licensee Association
    Public WithEvents pCompanyLicenseeAssociation As MUSTER.BusinessLogic.pCompanyLicensee
    Dim CompanyLicenseeInfo As MUSTER.Info.CompanyLicenseeInfo

    ' Licensee
    Public WithEvents pLicen As MUSTER.BusinessLogic.pLicensee
    Dim licenseeInfo As MUSTER.Info.LicenseeInfo

    ' Licensee Address
    Private WithEvents oLicAdd As New MUSTER.BusinessLogic.pComAddress

    ' Licensee Course
    Private oLicCourseInfo As MUSTER.Info.LicenseeCourseInfo

    ' Licensee Test
    Private oLicTestInfo As MUSTER.Info.LicenseeCourseTestInfo

    Private ncompAddressID As Integer = 0
    Private nLicAddressID As Integer = 0
    Dim strCompanyAddress As String = ""
    Dim dtNull As Date = CDate("01/01/0001")
    Dim nLicenseeID As Integer = 0
    Dim nCompanyID As Integer = 0
    Dim nCompanyLicenseeAssocID As Integer = 0
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
    'Dim bolCongratsLetter As Boolean = False
    'Dim bolRenewalLetter As Boolean = False
    'Dim bolInfoNeededLetter As Boolean = False
    'Dim bolLicenseeCertificate As Boolean = False
    'Dim bolLicenseeCard As Boolean = False
    Public bolFormClosing As Boolean = False
    Dim bolModelWindow As Boolean = False
    Dim bolLoading As Boolean
    Dim strLicenseeHireStatus As String = String.Empty
    ' to prevent duplicate letter generation
    'Private slLetterCount As SortedList
    Dim returnVal As String = String.Empty
    Private bolValidatationFailed As Boolean = False
    Friend callingForm As Form
    Private Enum LicenseeLetters
        CongLetter = 0
        LicenseeCard = 1
        LicenseeCertificateLetter = 2
        NoCertificationLetter = 3
        RenewalLetter = 4
    End Enum
    Friend ReadOnly Property LicenseeID() As Integer
        Get
            Return nLicenseeID
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
    '    'Need to tell the AppUser that we've instantiated another Licensee form...
    '    '
    '    MusterContainer.AppUser.LogEntry("Company", MyGuid.ToString)
    '    '
    '    ' The following line enables all forms to detect the visible form in the MDI container
    '    '
    '    MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "Company")

    '    If dir Is Nothing Then
    '        dir = New DirectoryInfo(MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_LicenseesImages).ProfileValue + "\")
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

    Public Sub New(Optional ByVal _LicenseeID As Integer = 0, Optional ByVal _CompanyID As Integer = 0, _
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

        nLicenseeID = _LicenseeID  'MR - 6/5
        nCompanyID = _CompanyID
        ncompAddressID = _CompanyAddressID
        strMode = _Mode

        InitImages()
    End Sub

    Public Sub New(ByRef pLic As MUSTER.BusinessLogic.pLicensee, _
                    ByRef pCompLicAssoc As MUSTER.BusinessLogic.pCompanyLicensee, _
                    ByRef pComAddress As MUSTER.BusinessLogic.pComAddress, _
                    ByVal CompanyID As Integer, ByVal strAddress As String, ByVal nLicCompanyAddID As Integer, _
                    ByVal _licenseeID As Integer, ByVal AssociationID As Integer, ByVal FinRespExpDate As Date, _
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

        pLicen = pLic
        pCompanyLicenseeAssociation = pCompLicAssoc
        oCompanyAdd = pComAddress
        nCompanyID = CompanyID
        strCompanyAddress = strAddress
        ncompAddressID = nLicCompanyAddID
        nLicenseeID = _licenseeID
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
    Friend WithEvents btnDeleteLicensee As System.Windows.Forms.Button
    Friend WithEvents btnComments As System.Windows.Forms.Button
    Friend WithEvents btnSaveLicensee As System.Windows.Forms.Button
    Friend WithEvents cmbTitle As System.Windows.Forms.ComboBox
    Friend WithEvents txtLastName As System.Windows.Forms.TextBox
    Friend WithEvents txtFirstName As System.Windows.Forms.TextBox
    Friend WithEvents txtMiddleName As System.Windows.Forms.TextBox
    Friend WithEvents lblMiddleName As System.Windows.Forms.Label
    Friend WithEvents lblFirstName As System.Windows.Forms.Label
    Friend WithEvents lblLastName As System.Windows.Forms.Label
    Friend WithEvents pnlLicenseesBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlLicensees As System.Windows.Forms.Panel
    Friend WithEvents ugCourses As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents lblPhoto As System.Windows.Forms.Label
    Friend WithEvents lblSignature As System.Windows.Forms.Label
    Public WithEvents dtExceptionGrantedDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblExceptionGrantedDate As System.Windows.Forms.Label
    Friend WithEvents lblLicenseExpirationDate As System.Windows.Forms.Label
    Friend WithEvents chkOverrideExpiration As System.Windows.Forms.CheckBox
    Friend WithEvents chkEmployeeLetter As System.Windows.Forms.CheckBox
    Friend WithEvents lblAppRcvdDate As System.Windows.Forms.Label
    Friend WithEvents lblStatus As System.Windows.Forms.Label
    Friend WithEvents lblHireStatus As System.Windows.Forms.Label
    Friend WithEvents lblCertificationType As System.Windows.Forms.Label
    Friend WithEvents txtEmail As System.Windows.Forms.TextBox
    Friend WithEvents lblEmail As System.Windows.Forms.Label
    Friend WithEvents txtLicenseNo As System.Windows.Forms.TextBox
    Friend WithEvents lblLicenseNumber As System.Windows.Forms.Label
    Friend WithEvents lblPersonTitle As System.Windows.Forms.Label
    Friend WithEvents cmbSuffix As System.Windows.Forms.ComboBox
    Friend WithEvents lblSuffix As System.Windows.Forms.Label
    Friend WithEvents txtLicenseeAddress As System.Windows.Forms.TextBox
    Friend WithEvents lblLicenseeAddress As System.Windows.Forms.Label
    Friend WithEvents btnCourseModify As System.Windows.Forms.Button
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
    Friend WithEvents UCLicenseeDocuments As MUSTER.DocumentViewControl
    Friend WithEvents ugLicenseeComplianceEvents As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents ugPriorCompanies As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents ugTests As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnTestModify As System.Windows.Forms.Button
    Friend WithEvents btnTestAdd As System.Windows.Forms.Button
    Friend WithEvents btnTestDelete As System.Windows.Forms.Button
    Friend WithEvents btnCourseAdd As System.Windows.Forms.Button
    Friend WithEvents btnCourseDelete As System.Windows.Forms.Button
    Friend WithEvents pnlCompany As System.Windows.Forms.Panel
    Friend WithEvents btnCompanyTitleSearch As System.Windows.Forms.Button
    Friend WithEvents txtCompanyTitle As System.Windows.Forms.TextBox
    Friend WithEvents txtCompany As System.Windows.Forms.TextBox
    Friend WithEvents lblCompany As System.Windows.Forms.Label
    Friend WithEvents lblCompanyTitle As System.Windows.Forms.Label
    Friend WithEvents cmbCertificationType As System.Windows.Forms.ComboBox
    Friend WithEvents cmbStatus As System.Windows.Forms.ComboBox
    Public WithEvents dtLicenseExpirationDate As System.Windows.Forms.DateTimePicker
    Public WithEvents dtIssuedDate As System.Windows.Forms.DateTimePicker
    Public WithEvents dtAppRcvdDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents cmbHireStatus As System.Windows.Forms.ComboBox
    Friend WithEvents pnlLicenseeInfo As System.Windows.Forms.Panel
    Friend WithEvents pnlLicenseeInfoHead As System.Windows.Forms.Panel
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
    Friend WithEvents lblLicenseeInfoDisplay As System.Windows.Forms.Label
    Friend WithEvents lblLicenseeInfo As System.Windows.Forms.Label
    Friend WithEvents lblCoursesDisplay As System.Windows.Forms.Label
    Friend WithEvents lblCourses As System.Windows.Forms.Label
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
    Friend WithEvents lblExtensionDeadlineDate As System.Windows.Forms.Label
    Friend WithEvents btnCompanyEnvelopes As System.Windows.Forms.Button
    Friend WithEvents btnCompanyLabels As System.Windows.Forms.Button
    Friend WithEvents lblIssuedDate As System.Windows.Forms.Label
    Friend WithEvents lblRestCert As System.Windows.Forms.Label
    Friend WithEvents LblOriginalIssueDate As System.Windows.Forms.Label
    Public WithEvents dtOriginalIssueDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents chkComplianceManager As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Licensees))
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.lblLicenseNo = New System.Windows.Forms.Label
        Me.pnlLicenseesBottom = New System.Windows.Forms.Panel
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnFlags = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnDeleteLicensee = New System.Windows.Forms.Button
        Me.btnComments = New System.Windows.Forms.Button
        Me.btnSaveLicensee = New System.Windows.Forms.Button
        Me.btnGenExpirationLetter = New System.Windows.Forms.Button
        Me.btnGenReminderLetter = New System.Windows.Forms.Button
        Me.btnCertify = New System.Windows.Forms.Button
        Me.pnlLicensees = New System.Windows.Forms.Panel
        Me.pnlCoursesHead = New System.Windows.Forms.Panel
        Me.lblCourses = New System.Windows.Forms.Label
        Me.lblCoursesDisplay = New System.Windows.Forms.Label
        Me.pnlCourses = New System.Windows.Forms.Panel
        Me.ugCourses = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlCoursesBottom = New System.Windows.Forms.Panel
        Me.btnCourseAdd = New System.Windows.Forms.Button
        Me.btnCourseModify = New System.Windows.Forms.Button
        Me.btnCourseDelete = New System.Windows.Forms.Button
        Me.pnlTestsHead = New System.Windows.Forms.Panel
        Me.lblTests = New System.Windows.Forms.Label
        Me.lblTestsDisplay = New System.Windows.Forms.Label
        Me.pnlTests = New System.Windows.Forms.Panel
        Me.ugTests = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlTestsBottom = New System.Windows.Forms.Panel
        Me.btnTestAdd = New System.Windows.Forms.Button
        Me.btnTestModify = New System.Windows.Forms.Button
        Me.btnTestDelete = New System.Windows.Forms.Button
        Me.pnlPriorCompaniesHead = New System.Windows.Forms.Panel
        Me.lblPriorCompanies = New System.Windows.Forms.Label
        Me.lblPriorCompaniesDisplay = New System.Windows.Forms.Label
        Me.pnlPriorCompanies = New System.Windows.Forms.Panel
        Me.ugPriorCompanies = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlLCEHead = New System.Windows.Forms.Panel
        Me.lblLCE = New System.Windows.Forms.Label
        Me.lblLCEDisplay = New System.Windows.Forms.Label
        Me.pnlLCE = New System.Windows.Forms.Panel
        Me.ugLicenseeComplianceEvents = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlDocumentsHead = New System.Windows.Forms.Panel
        Me.lblDocuments = New System.Windows.Forms.Label
        Me.lblDocumentsDisplay = New System.Windows.Forms.Label
        Me.pnlDocuments = New System.Windows.Forms.Panel
        Me.UCLicenseeDocuments = New MUSTER.DocumentViewControl
        Me.pnlLicenseeInfo = New System.Windows.Forms.Panel
        Me.dtOriginalIssueDate = New System.Windows.Forms.DateTimePicker
        Me.LblOriginalIssueDate = New System.Windows.Forms.Label
        Me.lblRestCert = New System.Windows.Forms.Label
        Me.cmbCertificationType = New System.Windows.Forms.ComboBox
        Me.cmbStatus = New System.Windows.Forms.ComboBox
        Me.dtLicenseExpirationDate = New System.Windows.Forms.DateTimePicker
        Me.dtIssuedDate = New System.Windows.Forms.DateTimePicker
        Me.dtAppRcvdDate = New System.Windows.Forms.DateTimePicker
        Me.cmbHireStatus = New System.Windows.Forms.ComboBox
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
        Me.lblExtensionDeadlineDate = New System.Windows.Forms.Label
        Me.chkEmployeeLetter = New System.Windows.Forms.CheckBox
        Me.lblHireStatus = New System.Windows.Forms.Label
        Me.lblStatus = New System.Windows.Forms.Label
        Me.lblCertificationType = New System.Windows.Forms.Label
        Me.chkOverrideExpiration = New System.Windows.Forms.CheckBox
        Me.lblAppRcvdDate = New System.Windows.Forms.Label
        Me.lblIssuedDate = New System.Windows.Forms.Label
        Me.lblLicenseExpirationDate = New System.Windows.Forms.Label
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
        Me.txtLicenseeAddress = New System.Windows.Forms.TextBox
        Me.lblLicenseeAddress = New System.Windows.Forms.Label
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
        Me.pnlLicenseeInfoHead = New System.Windows.Forms.Panel
        Me.lblLicenseeInfo = New System.Windows.Forms.Label
        Me.lblLicenseeInfoDisplay = New System.Windows.Forms.Label
        Me.chkComplianceManager = New System.Windows.Forms.CheckBox
        Me.pnlLicenseesBottom.SuspendLayout()
        Me.pnlLicensees.SuspendLayout()
        Me.pnlCoursesHead.SuspendLayout()
        Me.pnlCourses.SuspendLayout()
        CType(Me.ugCourses, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlCoursesBottom.SuspendLayout()
        Me.pnlTestsHead.SuspendLayout()
        Me.pnlTests.SuspendLayout()
        CType(Me.ugTests, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlTestsBottom.SuspendLayout()
        Me.pnlPriorCompaniesHead.SuspendLayout()
        Me.pnlPriorCompanies.SuspendLayout()
        CType(Me.ugPriorCompanies, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlLCEHead.SuspendLayout()
        Me.pnlLCE.SuspendLayout()
        CType(Me.ugLicenseeComplianceEvents, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlDocumentsHead.SuspendLayout()
        Me.pnlDocuments.SuspendLayout()
        Me.pnlLicenseeInfo.SuspendLayout()
        Me.pnlCompany.SuspendLayout()
        Me.pnlExtension.SuspendLayout()
        Me.pnlPhone.SuspendLayout()
        CType(Me.mskTxtPhone1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtCell, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtPhone2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.mskTxtFax, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlLicenseeInfoHead.SuspendLayout()
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
        'pnlLicenseesBottom
        '
        Me.pnlLicenseesBottom.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlLicenseesBottom.Controls.Add(Me.btnClose)
        Me.pnlLicenseesBottom.Controls.Add(Me.btnFlags)
        Me.pnlLicenseesBottom.Controls.Add(Me.btnCancel)
        Me.pnlLicenseesBottom.Controls.Add(Me.btnDeleteLicensee)
        Me.pnlLicenseesBottom.Controls.Add(Me.btnComments)
        Me.pnlLicenseesBottom.Controls.Add(Me.btnSaveLicensee)
        Me.pnlLicenseesBottom.Controls.Add(Me.btnGenExpirationLetter)
        Me.pnlLicenseesBottom.Controls.Add(Me.btnGenReminderLetter)
        Me.pnlLicenseesBottom.Controls.Add(Me.btnCertify)
        Me.pnlLicenseesBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlLicenseesBottom.Location = New System.Drawing.Point(0, 509)
        Me.pnlLicenseesBottom.Name = "pnlLicenseesBottom"
        Me.pnlLicenseesBottom.Size = New System.Drawing.Size(752, 88)
        Me.pnlLicenseesBottom.TabIndex = 1
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
        'btnDeleteLicensee
        '
        Me.btnDeleteLicensee.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDeleteLicensee.Location = New System.Drawing.Point(160, 9)
        Me.btnDeleteLicensee.Name = "btnDeleteLicensee"
        Me.btnDeleteLicensee.Size = New System.Drawing.Size(72, 26)
        Me.btnDeleteLicensee.TabIndex = 2
        Me.btnDeleteLicensee.Text = "Delete"
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
        'btnSaveLicensee
        '
        Me.btnSaveLicensee.Enabled = False
        Me.btnSaveLicensee.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSaveLicensee.Location = New System.Drawing.Point(8, 8)
        Me.btnSaveLicensee.Name = "btnSaveLicensee"
        Me.btnSaveLicensee.Size = New System.Drawing.Size(64, 26)
        Me.btnSaveLicensee.TabIndex = 0
        Me.btnSaveLicensee.Text = "Save"
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
        'pnlLicensees
        '
        Me.pnlLicensees.AutoScroll = True
        Me.pnlLicensees.Controls.Add(Me.pnlCoursesHead)
        Me.pnlLicensees.Controls.Add(Me.pnlCourses)
        Me.pnlLicensees.Controls.Add(Me.pnlCoursesBottom)
        Me.pnlLicensees.Controls.Add(Me.pnlTestsHead)
        Me.pnlLicensees.Controls.Add(Me.pnlTests)
        Me.pnlLicensees.Controls.Add(Me.pnlTestsBottom)
        Me.pnlLicensees.Controls.Add(Me.pnlPriorCompaniesHead)
        Me.pnlLicensees.Controls.Add(Me.pnlPriorCompanies)
        Me.pnlLicensees.Controls.Add(Me.pnlLCEHead)
        Me.pnlLicensees.Controls.Add(Me.pnlLCE)
        Me.pnlLicensees.Controls.Add(Me.pnlDocumentsHead)
        Me.pnlLicensees.Controls.Add(Me.pnlDocuments)
        Me.pnlLicensees.Controls.Add(Me.pnlLicenseeInfo)
        Me.pnlLicensees.Controls.Add(Me.pnlLicenseeInfoHead)
        Me.pnlLicensees.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlLicensees.Location = New System.Drawing.Point(0, 0)
        Me.pnlLicensees.Name = "pnlLicensees"
        Me.pnlLicensees.Size = New System.Drawing.Size(752, 509)
        Me.pnlLicensees.TabIndex = 0
        '
        'pnlCoursesHead
        '
        Me.pnlCoursesHead.Controls.Add(Me.lblCourses)
        Me.pnlCoursesHead.Controls.Add(Me.lblCoursesDisplay)
        Me.pnlCoursesHead.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlCoursesHead.Location = New System.Drawing.Point(0, 560)
        Me.pnlCoursesHead.Name = "pnlCoursesHead"
        Me.pnlCoursesHead.Size = New System.Drawing.Size(736, 24)
        Me.pnlCoursesHead.TabIndex = 1063
        '
        'lblCourses
        '
        Me.lblCourses.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblCourses.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblCourses.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblCourses.Location = New System.Drawing.Point(16, 0)
        Me.lblCourses.Name = "lblCourses"
        Me.lblCourses.Size = New System.Drawing.Size(720, 24)
        Me.lblCourses.TabIndex = 256
        Me.lblCourses.Text = "Courses"
        Me.lblCourses.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
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
        Me.pnlCourses.Controls.Add(Me.ugCourses)
        Me.pnlCourses.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlCourses.Location = New System.Drawing.Point(0, 584)
        Me.pnlCourses.Name = "pnlCourses"
        Me.pnlCourses.Size = New System.Drawing.Size(736, 100)
        Me.pnlCourses.TabIndex = 1062
        '
        'ugCourses
        '
        Me.ugCourses.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugCourses.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.TemplateOnBottom
        Me.ugCourses.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugCourses.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugCourses.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugCourses.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugCourses.Location = New System.Drawing.Point(0, 0)
        Me.ugCourses.Name = "ugCourses"
        Me.ugCourses.Size = New System.Drawing.Size(736, 100)
        Me.ugCourses.TabIndex = 21
        '
        'pnlCoursesBottom
        '
        Me.pnlCoursesBottom.Controls.Add(Me.btnCourseAdd)
        Me.pnlCoursesBottom.Controls.Add(Me.btnCourseModify)
        Me.pnlCoursesBottom.Controls.Add(Me.btnCourseDelete)
        Me.pnlCoursesBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlCoursesBottom.Location = New System.Drawing.Point(0, 684)
        Me.pnlCoursesBottom.Name = "pnlCoursesBottom"
        Me.pnlCoursesBottom.Size = New System.Drawing.Size(736, 40)
        Me.pnlCoursesBottom.TabIndex = 269
        '
        'btnCourseAdd
        '
        Me.btnCourseAdd.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCourseAdd.Location = New System.Drawing.Point(40, 8)
        Me.btnCourseAdd.Name = "btnCourseAdd"
        Me.btnCourseAdd.Size = New System.Drawing.Size(64, 26)
        Me.btnCourseAdd.TabIndex = 267
        Me.btnCourseAdd.Text = "Add"
        '
        'btnCourseModify
        '
        Me.btnCourseModify.Enabled = False
        Me.btnCourseModify.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCourseModify.Location = New System.Drawing.Point(184, 8)
        Me.btnCourseModify.Name = "btnCourseModify"
        Me.btnCourseModify.Size = New System.Drawing.Size(64, 26)
        Me.btnCourseModify.TabIndex = 266
        Me.btnCourseModify.Text = "Modify"
        Me.btnCourseModify.Visible = False
        '
        'btnCourseDelete
        '
        Me.btnCourseDelete.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCourseDelete.Location = New System.Drawing.Point(112, 8)
        Me.btnCourseDelete.Name = "btnCourseDelete"
        Me.btnCourseDelete.Size = New System.Drawing.Size(64, 26)
        Me.btnCourseDelete.TabIndex = 268
        Me.btnCourseDelete.Text = "Delete"
        '
        'pnlTestsHead
        '
        Me.pnlTestsHead.Controls.Add(Me.lblTests)
        Me.pnlTestsHead.Controls.Add(Me.lblTestsDisplay)
        Me.pnlTestsHead.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlTestsHead.Location = New System.Drawing.Point(0, 724)
        Me.pnlTestsHead.Name = "pnlTestsHead"
        Me.pnlTestsHead.Size = New System.Drawing.Size(736, 24)
        Me.pnlTestsHead.TabIndex = 1061
        '
        'lblTests
        '
        Me.lblTests.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblTests.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblTests.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblTests.Location = New System.Drawing.Point(16, 0)
        Me.lblTests.Name = "lblTests"
        Me.lblTests.Size = New System.Drawing.Size(720, 24)
        Me.lblTests.TabIndex = 259
        Me.lblTests.Text = "Tests"
        Me.lblTests.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
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
        '
        'pnlTests
        '
        Me.pnlTests.Controls.Add(Me.ugTests)
        Me.pnlTests.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlTests.Location = New System.Drawing.Point(0, 748)
        Me.pnlTests.Name = "pnlTests"
        Me.pnlTests.Size = New System.Drawing.Size(736, 100)
        Me.pnlTests.TabIndex = 1063
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
        Me.ugTests.Size = New System.Drawing.Size(736, 100)
        Me.ugTests.TabIndex = 24
        '
        'pnlTestsBottom
        '
        Me.pnlTestsBottom.Controls.Add(Me.btnTestAdd)
        Me.pnlTestsBottom.Controls.Add(Me.btnTestModify)
        Me.pnlTestsBottom.Controls.Add(Me.btnTestDelete)
        Me.pnlTestsBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlTestsBottom.Location = New System.Drawing.Point(0, 848)
        Me.pnlTestsBottom.Name = "pnlTestsBottom"
        Me.pnlTestsBottom.Size = New System.Drawing.Size(736, 40)
        Me.pnlTestsBottom.TabIndex = 273
        '
        'btnTestAdd
        '
        Me.btnTestAdd.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnTestAdd.Location = New System.Drawing.Point(40, 8)
        Me.btnTestAdd.Name = "btnTestAdd"
        Me.btnTestAdd.Size = New System.Drawing.Size(64, 26)
        Me.btnTestAdd.TabIndex = 271
        Me.btnTestAdd.Text = "Add"
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
        '
        'pnlPriorCompaniesHead
        '
        Me.pnlPriorCompaniesHead.Controls.Add(Me.lblPriorCompanies)
        Me.pnlPriorCompaniesHead.Controls.Add(Me.lblPriorCompaniesDisplay)
        Me.pnlPriorCompaniesHead.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlPriorCompaniesHead.Location = New System.Drawing.Point(0, 888)
        Me.pnlPriorCompaniesHead.Name = "pnlPriorCompaniesHead"
        Me.pnlPriorCompaniesHead.Size = New System.Drawing.Size(736, 24)
        Me.pnlPriorCompaniesHead.TabIndex = 1063
        '
        'lblPriorCompanies
        '
        Me.lblPriorCompanies.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblPriorCompanies.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblPriorCompanies.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblPriorCompanies.Location = New System.Drawing.Point(16, 0)
        Me.lblPriorCompanies.Name = "lblPriorCompanies"
        Me.lblPriorCompanies.Size = New System.Drawing.Size(720, 24)
        Me.lblPriorCompanies.TabIndex = 262
        Me.lblPriorCompanies.Text = "Prior Companies"
        Me.lblPriorCompanies.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
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
        '
        'pnlPriorCompanies
        '
        Me.pnlPriorCompanies.Controls.Add(Me.ugPriorCompanies)
        Me.pnlPriorCompanies.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlPriorCompanies.Location = New System.Drawing.Point(0, 912)
        Me.pnlPriorCompanies.Name = "pnlPriorCompanies"
        Me.pnlPriorCompanies.Size = New System.Drawing.Size(736, 200)
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
        Me.ugPriorCompanies.Size = New System.Drawing.Size(736, 200)
        Me.ugPriorCompanies.TabIndex = 26
        '
        'pnlLCEHead
        '
        Me.pnlLCEHead.Controls.Add(Me.lblLCE)
        Me.pnlLCEHead.Controls.Add(Me.lblLCEDisplay)
        Me.pnlLCEHead.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlLCEHead.Location = New System.Drawing.Point(0, 1112)
        Me.pnlLCEHead.Name = "pnlLCEHead"
        Me.pnlLCEHead.Size = New System.Drawing.Size(736, 24)
        Me.pnlLCEHead.TabIndex = 1063
        '
        'lblLCE
        '
        Me.lblLCE.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblLCE.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblLCE.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblLCE.Location = New System.Drawing.Point(16, 0)
        Me.lblLCE.Name = "lblLCE"
        Me.lblLCE.Size = New System.Drawing.Size(720, 24)
        Me.lblLCE.TabIndex = 265
        Me.lblLCE.Text = "Licensee Compliance Events"
        Me.lblLCE.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
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
        '
        'pnlLCE
        '
        Me.pnlLCE.Controls.Add(Me.ugLicenseeComplianceEvents)
        Me.pnlLCE.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlLCE.Location = New System.Drawing.Point(0, 1136)
        Me.pnlLCE.Name = "pnlLCE"
        Me.pnlLCE.Size = New System.Drawing.Size(736, 200)
        Me.pnlLCE.TabIndex = 1063
        '
        'ugLicenseeComplianceEvents
        '
        Me.ugLicenseeComplianceEvents.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugLicenseeComplianceEvents.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugLicenseeComplianceEvents.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugLicenseeComplianceEvents.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugLicenseeComplianceEvents.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugLicenseeComplianceEvents.Location = New System.Drawing.Point(0, 0)
        Me.ugLicenseeComplianceEvents.Name = "ugLicenseeComplianceEvents"
        Me.ugLicenseeComplianceEvents.Size = New System.Drawing.Size(736, 200)
        Me.ugLicenseeComplianceEvents.TabIndex = 27
        '
        'pnlDocumentsHead
        '
        Me.pnlDocumentsHead.Controls.Add(Me.lblDocuments)
        Me.pnlDocumentsHead.Controls.Add(Me.lblDocumentsDisplay)
        Me.pnlDocumentsHead.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlDocumentsHead.Location = New System.Drawing.Point(0, 1336)
        Me.pnlDocumentsHead.Name = "pnlDocumentsHead"
        Me.pnlDocumentsHead.Size = New System.Drawing.Size(736, 24)
        Me.pnlDocumentsHead.TabIndex = 1063
        '
        'lblDocuments
        '
        Me.lblDocuments.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblDocuments.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblDocuments.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblDocuments.Location = New System.Drawing.Point(16, 0)
        Me.lblDocuments.Name = "lblDocuments"
        Me.lblDocuments.Size = New System.Drawing.Size(720, 24)
        Me.lblDocuments.TabIndex = 279
        Me.lblDocuments.Text = "Documents"
        Me.lblDocuments.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
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
        '
        'pnlDocuments
        '
        Me.pnlDocuments.Controls.Add(Me.UCLicenseeDocuments)
        Me.pnlDocuments.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlDocuments.Location = New System.Drawing.Point(0, 1360)
        Me.pnlDocuments.Name = "pnlDocuments"
        Me.pnlDocuments.Size = New System.Drawing.Size(736, 200)
        Me.pnlDocuments.TabIndex = 1063
        '
        'UCLicenseeDocuments
        '
        Me.UCLicenseeDocuments.AutoScroll = True
        Me.UCLicenseeDocuments.Dock = System.Windows.Forms.DockStyle.Fill
        Me.UCLicenseeDocuments.Location = New System.Drawing.Point(0, 0)
        Me.UCLicenseeDocuments.Name = "UCLicenseeDocuments"
        Me.UCLicenseeDocuments.Size = New System.Drawing.Size(736, 200)
        Me.UCLicenseeDocuments.TabIndex = 278
        '
        'pnlLicenseeInfo
        '
        Me.pnlLicenseeInfo.BackColor = System.Drawing.SystemColors.Control
        Me.pnlLicenseeInfo.Controls.Add(Me.chkComplianceManager)
        Me.pnlLicenseeInfo.Controls.Add(Me.dtOriginalIssueDate)
        Me.pnlLicenseeInfo.Controls.Add(Me.LblOriginalIssueDate)
        Me.pnlLicenseeInfo.Controls.Add(Me.lblRestCert)
        Me.pnlLicenseeInfo.Controls.Add(Me.cmbCertificationType)
        Me.pnlLicenseeInfo.Controls.Add(Me.cmbStatus)
        Me.pnlLicenseeInfo.Controls.Add(Me.dtLicenseExpirationDate)
        Me.pnlLicenseeInfo.Controls.Add(Me.dtIssuedDate)
        Me.pnlLicenseeInfo.Controls.Add(Me.dtAppRcvdDate)
        Me.pnlLicenseeInfo.Controls.Add(Me.cmbHireStatus)
        Me.pnlLicenseeInfo.Controls.Add(Me.pnlCompany)
        Me.pnlLicenseeInfo.Controls.Add(Me.pbPhoto)
        Me.pnlLicenseeInfo.Controls.Add(Me.lblPhoto)
        Me.pnlLicenseeInfo.Controls.Add(Me.pbSign)
        Me.pnlLicenseeInfo.Controls.Add(Me.lblSignature)
        Me.pnlLicenseeInfo.Controls.Add(Me.dtExtensionDeadlineDate)
        Me.pnlLicenseeInfo.Controls.Add(Me.dtExceptionGrantedDate)
        Me.pnlLicenseeInfo.Controls.Add(Me.lblExceptionGrantedDate)
        Me.pnlLicenseeInfo.Controls.Add(Me.lblExtensionDeadlineDate)
        Me.pnlLicenseeInfo.Controls.Add(Me.chkEmployeeLetter)
        Me.pnlLicenseeInfo.Controls.Add(Me.lblHireStatus)
        Me.pnlLicenseeInfo.Controls.Add(Me.lblStatus)
        Me.pnlLicenseeInfo.Controls.Add(Me.lblCertificationType)
        Me.pnlLicenseeInfo.Controls.Add(Me.chkOverrideExpiration)
        Me.pnlLicenseeInfo.Controls.Add(Me.lblAppRcvdDate)
        Me.pnlLicenseeInfo.Controls.Add(Me.lblIssuedDate)
        Me.pnlLicenseeInfo.Controls.Add(Me.lblLicenseExpirationDate)
        Me.pnlLicenseeInfo.Controls.Add(Me.pnlExtension)
        Me.pnlLicenseeInfo.Controls.Add(Me.pnlPhone)
        Me.pnlLicenseeInfo.Controls.Add(Me.txtLicenseeAddress)
        Me.pnlLicenseeInfo.Controls.Add(Me.lblLicenseeAddress)
        Me.pnlLicenseeInfo.Controls.Add(Me.btnLabels)
        Me.pnlLicenseeInfo.Controls.Add(Me.btnEnvelopes)
        Me.pnlLicenseeInfo.Controls.Add(Me.lblEmail)
        Me.pnlLicenseeInfo.Controls.Add(Me.txtEmail)
        Me.pnlLicenseeInfo.Controls.Add(Me.lblSuffix)
        Me.pnlLicenseeInfo.Controls.Add(Me.lblLastName)
        Me.pnlLicenseeInfo.Controls.Add(Me.cmbSuffix)
        Me.pnlLicenseeInfo.Controls.Add(Me.txtLastName)
        Me.pnlLicenseeInfo.Controls.Add(Me.lblMiddleName)
        Me.pnlLicenseeInfo.Controls.Add(Me.lblFirstName)
        Me.pnlLicenseeInfo.Controls.Add(Me.lblPersonTitle)
        Me.pnlLicenseeInfo.Controls.Add(Me.lblLicenseNumber)
        Me.pnlLicenseeInfo.Controls.Add(Me.txtLicenseNo)
        Me.pnlLicenseeInfo.Controls.Add(Me.cmbTitle)
        Me.pnlLicenseeInfo.Controls.Add(Me.txtFirstName)
        Me.pnlLicenseeInfo.Controls.Add(Me.txtMiddleName)
        Me.pnlLicenseeInfo.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlLicenseeInfo.Location = New System.Drawing.Point(0, 24)
        Me.pnlLicenseeInfo.Name = "pnlLicenseeInfo"
        Me.pnlLicenseeInfo.Size = New System.Drawing.Size(736, 536)
        Me.pnlLicenseeInfo.TabIndex = 1065
        '
        'dtOriginalIssueDate
        '
        Me.dtOriginalIssueDate.Checked = False
        Me.dtOriginalIssueDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtOriginalIssueDate.Location = New System.Drawing.Point(624, 184)
        Me.dtOriginalIssueDate.Name = "dtOriginalIssueDate"
        Me.dtOriginalIssueDate.ShowCheckBox = True
        Me.dtOriginalIssueDate.Size = New System.Drawing.Size(104, 20)
        Me.dtOriginalIssueDate.TabIndex = 1063
        '
        'LblOriginalIssueDate
        '
        Me.LblOriginalIssueDate.Location = New System.Drawing.Point(520, 188)
        Me.LblOriginalIssueDate.Name = "LblOriginalIssueDate"
        Me.LblOriginalIssueDate.Size = New System.Drawing.Size(104, 20)
        Me.LblOriginalIssueDate.TabIndex = 1062
        Me.LblOriginalIssueDate.Text = "Original Issue Date:"
        '
        'lblRestCert
        '
        Me.lblRestCert.BackColor = System.Drawing.Color.Firebrick
        Me.lblRestCert.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblRestCert.Font = New System.Drawing.Font("Bookman Old Style", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblRestCert.Location = New System.Drawing.Point(336, 432)
        Me.lblRestCert.Name = "lblRestCert"
        Me.lblRestCert.Size = New System.Drawing.Size(184, 96)
        Me.lblRestCert.TabIndex = 1061
        Me.lblRestCert.Text = "Restricted Certification"
        Me.lblRestCert.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblRestCert.Visible = False
        '
        'cmbCertificationType
        '
        Me.cmbCertificationType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbCertificationType.Location = New System.Drawing.Point(400, 88)
        Me.cmbCertificationType.Name = "cmbCertificationType"
        Me.cmbCertificationType.Size = New System.Drawing.Size(256, 21)
        Me.cmbCertificationType.TabIndex = 276
        '
        'cmbStatus
        '
        Me.cmbStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbStatus.Location = New System.Drawing.Point(400, 56)
        Me.cmbStatus.Name = "cmbStatus"
        Me.cmbStatus.Size = New System.Drawing.Size(256, 21)
        Me.cmbStatus.TabIndex = 275
        '
        'dtLicenseExpirationDate
        '
        Me.dtLicenseExpirationDate.Checked = False
        Me.dtLicenseExpirationDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtLicenseExpirationDate.Location = New System.Drawing.Point(400, 184)
        Me.dtLicenseExpirationDate.Name = "dtLicenseExpirationDate"
        Me.dtLicenseExpirationDate.ShowCheckBox = True
        Me.dtLicenseExpirationDate.Size = New System.Drawing.Size(104, 20)
        Me.dtLicenseExpirationDate.TabIndex = 279
        '
        'dtIssuedDate
        '
        Me.dtIssuedDate.Checked = False
        Me.dtIssuedDate.Enabled = False
        Me.dtIssuedDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtIssuedDate.Location = New System.Drawing.Point(400, 152)
        Me.dtIssuedDate.Name = "dtIssuedDate"
        Me.dtIssuedDate.ShowCheckBox = True
        Me.dtIssuedDate.Size = New System.Drawing.Size(104, 20)
        Me.dtIssuedDate.TabIndex = 278
        '
        'dtAppRcvdDate
        '
        Me.dtAppRcvdDate.Checked = False
        Me.dtAppRcvdDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtAppRcvdDate.Location = New System.Drawing.Point(400, 120)
        Me.dtAppRcvdDate.Name = "dtAppRcvdDate"
        Me.dtAppRcvdDate.ShowCheckBox = True
        Me.dtAppRcvdDate.Size = New System.Drawing.Size(104, 20)
        Me.dtAppRcvdDate.TabIndex = 277
        '
        'cmbHireStatus
        '
        Me.cmbHireStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbHireStatus.ItemHeight = 13
        Me.cmbHireStatus.Location = New System.Drawing.Point(400, 8)
        Me.cmbHireStatus.Name = "cmbHireStatus"
        Me.cmbHireStatus.Size = New System.Drawing.Size(256, 21)
        Me.cmbHireStatus.TabIndex = 274
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
        Me.pnlCompany.Location = New System.Drawing.Point(320, 208)
        Me.pnlCompany.Name = "pnlCompany"
        Me.pnlCompany.Size = New System.Drawing.Size(408, 128)
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
        Me.btnCompanyTitleSearch.Location = New System.Drawing.Point(384, 30)
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
        Me.txtCompanyTitle.Size = New System.Drawing.Size(296, 90)
        Me.txtCompanyTitle.TabIndex = 236
        Me.txtCompanyTitle.Text = ""
        Me.txtCompanyTitle.WordWrap = False
        '
        'txtCompany
        '
        Me.txtCompany.Location = New System.Drawing.Point(80, 6)
        Me.txtCompany.Name = "txtCompany"
        Me.txtCompany.ReadOnly = True
        Me.txtCompany.Size = New System.Drawing.Size(296, 20)
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
        Me.pbPhoto.Location = New System.Drawing.Point(528, 360)
        Me.pbPhoto.Name = "pbPhoto"
        Me.pbPhoto.Size = New System.Drawing.Size(200, 168)
        Me.pbPhoto.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.pbPhoto.TabIndex = 271
        Me.pbPhoto.TabStop = False
        '
        'lblPhoto
        '
        Me.lblPhoto.Location = New System.Drawing.Point(528, 336)
        Me.lblPhoto.Name = "lblPhoto"
        Me.lblPhoto.Size = New System.Drawing.Size(34, 16)
        Me.lblPhoto.TabIndex = 250
        Me.lblPhoto.Text = "Photo"
        Me.lblPhoto.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pbSign
        '
        Me.pbSign.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pbSign.Location = New System.Drawing.Point(336, 360)
        Me.pbSign.Name = "pbSign"
        Me.pbSign.Size = New System.Drawing.Size(184, 64)
        Me.pbSign.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.pbSign.TabIndex = 270
        Me.pbSign.TabStop = False
        '
        'lblSignature
        '
        Me.lblSignature.Location = New System.Drawing.Point(336, 336)
        Me.lblSignature.Name = "lblSignature"
        Me.lblSignature.Size = New System.Drawing.Size(53, 16)
        Me.lblSignature.TabIndex = 249
        Me.lblSignature.Text = "Signature"
        Me.lblSignature.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dtExtensionDeadlineDate
        '
        Me.dtExtensionDeadlineDate.Checked = False
        Me.dtExtensionDeadlineDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtExtensionDeadlineDate.Location = New System.Drawing.Point(624, 152)
        Me.dtExtensionDeadlineDate.Name = "dtExtensionDeadlineDate"
        Me.dtExtensionDeadlineDate.ShowCheckBox = True
        Me.dtExtensionDeadlineDate.Size = New System.Drawing.Size(104, 20)
        Me.dtExtensionDeadlineDate.TabIndex = 21
        '
        'dtExceptionGrantedDate
        '
        Me.dtExceptionGrantedDate.Checked = False
        Me.dtExceptionGrantedDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtExceptionGrantedDate.Location = New System.Drawing.Point(624, 120)
        Me.dtExceptionGrantedDate.Name = "dtExceptionGrantedDate"
        Me.dtExceptionGrantedDate.ShowCheckBox = True
        Me.dtExceptionGrantedDate.Size = New System.Drawing.Size(104, 20)
        Me.dtExceptionGrantedDate.TabIndex = 23
        '
        'lblExceptionGrantedDate
        '
        Me.lblExceptionGrantedDate.Location = New System.Drawing.Point(504, 120)
        Me.lblExceptionGrantedDate.Name = "lblExceptionGrantedDate"
        Me.lblExceptionGrantedDate.Size = New System.Drawing.Size(112, 24)
        Me.lblExceptionGrantedDate.TabIndex = 245
        Me.lblExceptionGrantedDate.Text = "Date Licensee Requested Extension:"
        Me.lblExceptionGrantedDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblExtensionDeadlineDate
        '
        Me.lblExtensionDeadlineDate.Location = New System.Drawing.Point(504, 160)
        Me.lblExtensionDeadlineDate.Name = "lblExtensionDeadlineDate"
        Me.lblExtensionDeadlineDate.Size = New System.Drawing.Size(112, 24)
        Me.lblExtensionDeadlineDate.TabIndex = 241
        Me.lblExtensionDeadlineDate.Text = "Extension Deadline Date:"
        Me.lblExtensionDeadlineDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'chkEmployeeLetter
        '
        Me.chkEmployeeLetter.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkEmployeeLetter.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.chkEmployeeLetter.Location = New System.Drawing.Point(296, 32)
        Me.chkEmployeeLetter.Name = "chkEmployeeLetter"
        Me.chkEmployeeLetter.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkEmployeeLetter.Size = New System.Drawing.Size(120, 21)
        Me.chkEmployeeLetter.TabIndex = 15
        Me.chkEmployeeLetter.Tag = "649"
        Me.chkEmployeeLetter.Text = " :Employee Letter"
        Me.chkEmployeeLetter.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblHireStatus
        '
        Me.lblHireStatus.Location = New System.Drawing.Point(320, 8)
        Me.lblHireStatus.Name = "lblHireStatus"
        Me.lblHireStatus.Size = New System.Drawing.Size(72, 16)
        Me.lblHireStatus.TabIndex = 230
        Me.lblHireStatus.Text = "Hire Status:"
        Me.lblHireStatus.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblStatus
        '
        Me.lblStatus.Location = New System.Drawing.Point(296, 56)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(96, 16)
        Me.lblStatus.TabIndex = 231
        Me.lblStatus.Text = "Status:"
        Me.lblStatus.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblCertificationType
        '
        Me.lblCertificationType.Location = New System.Drawing.Point(280, 88)
        Me.lblCertificationType.Name = "lblCertificationType"
        Me.lblCertificationType.Size = New System.Drawing.Size(112, 16)
        Me.lblCertificationType.TabIndex = 228
        Me.lblCertificationType.Text = "Certification Type:"
        Me.lblCertificationType.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'chkOverrideExpiration
        '
        Me.chkOverrideExpiration.Enabled = False
        Me.chkOverrideExpiration.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkOverrideExpiration.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.chkOverrideExpiration.Location = New System.Drawing.Point(504, 32)
        Me.chkOverrideExpiration.Name = "chkOverrideExpiration"
        Me.chkOverrideExpiration.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkOverrideExpiration.Size = New System.Drawing.Size(152, 16)
        Me.chkOverrideExpiration.TabIndex = 16
        Me.chkOverrideExpiration.Tag = "649"
        Me.chkOverrideExpiration.Text = " Override Expiration"
        Me.chkOverrideExpiration.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        Me.chkOverrideExpiration.Visible = False
        '
        'lblAppRcvdDate
        '
        Me.lblAppRcvdDate.Location = New System.Drawing.Point(296, 120)
        Me.lblAppRcvdDate.Name = "lblAppRcvdDate"
        Me.lblAppRcvdDate.Size = New System.Drawing.Size(96, 16)
        Me.lblAppRcvdDate.TabIndex = 232
        Me.lblAppRcvdDate.Text = "App Rcvd Date:"
        Me.lblAppRcvdDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblIssuedDate
        '
        Me.lblIssuedDate.Location = New System.Drawing.Point(280, 152)
        Me.lblIssuedDate.Name = "lblIssuedDate"
        Me.lblIssuedDate.Size = New System.Drawing.Size(112, 16)
        Me.lblIssuedDate.TabIndex = 239
        Me.lblIssuedDate.Text = "Issued Date:"
        Me.lblIssuedDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblLicenseExpirationDate
        '
        Me.lblLicenseExpirationDate.Location = New System.Drawing.Point(264, 184)
        Me.lblLicenseExpirationDate.Name = "lblLicenseExpirationDate"
        Me.lblLicenseExpirationDate.Size = New System.Drawing.Size(128, 16)
        Me.lblLicenseExpirationDate.TabIndex = 243
        Me.lblLicenseExpirationDate.Text = "License Expiration Date:"
        Me.lblLicenseExpirationDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlExtension
        '
        Me.pnlExtension.Controls.Add(Me.lblExt2)
        Me.pnlExtension.Controls.Add(Me.lblExt1)
        Me.pnlExtension.Controls.Add(Me.txtExt2)
        Me.pnlExtension.Controls.Add(Me.txtExt1)
        Me.pnlExtension.Location = New System.Drawing.Point(248, 352)
        Me.pnlExtension.Name = "pnlExtension"
        Me.pnlExtension.Size = New System.Drawing.Size(75, 64)
        Me.pnlExtension.TabIndex = 275
        '
        'lblExt2
        '
        Me.lblExt2.Location = New System.Drawing.Point(0, 40)
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
        Me.txtExt2.Location = New System.Drawing.Point(32, 40)
        Me.txtExt2.Name = "txtExt2"
        Me.txtExt2.Size = New System.Drawing.Size(40, 20)
        Me.txtExt2.TabIndex = 11
        Me.txtExt2.Text = ""
        '
        'txtExt1
        '
        Me.txtExt1.Location = New System.Drawing.Point(32, 6)
        Me.txtExt1.Name = "txtExt1"
        Me.txtExt1.Size = New System.Drawing.Size(40, 20)
        Me.txtExt1.TabIndex = 9
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
        Me.pnlPhone.Location = New System.Drawing.Point(24, 352)
        Me.pnlPhone.Name = "pnlPhone"
        Me.pnlPhone.Size = New System.Drawing.Size(216, 128)
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
        Me.lblPhone2.Location = New System.Drawing.Point(10, 40)
        Me.lblPhone2.Name = "lblPhone2"
        Me.lblPhone2.Size = New System.Drawing.Size(56, 16)
        Me.lblPhone2.TabIndex = 228
        Me.lblPhone2.Text = "Phone 2:"
        Me.lblPhone2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'mskTxtCell
        '
        Me.mskTxtCell.ContainingControl = Me
        Me.mskTxtCell.Location = New System.Drawing.Point(72, 104)
        Me.mskTxtCell.Name = "mskTxtCell"
        Me.mskTxtCell.OcxState = CType(resources.GetObject("mskTxtCell.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtCell.Size = New System.Drawing.Size(142, 20)
        Me.mskTxtCell.TabIndex = 13
        '
        'lblCell
        '
        Me.lblCell.Location = New System.Drawing.Point(10, 104)
        Me.lblCell.Name = "lblCell"
        Me.lblCell.Size = New System.Drawing.Size(56, 16)
        Me.lblCell.TabIndex = 230
        Me.lblCell.Text = "Cell:"
        Me.lblCell.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'mskTxtPhone2
        '
        Me.mskTxtPhone2.ContainingControl = Me
        Me.mskTxtPhone2.Location = New System.Drawing.Point(72, 40)
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
        Me.mskTxtFax.Location = New System.Drawing.Point(72, 72)
        Me.mskTxtFax.Name = "mskTxtFax"
        Me.mskTxtFax.OcxState = CType(resources.GetObject("mskTxtFax.OcxState"), System.Windows.Forms.AxHost.State)
        Me.mskTxtFax.Size = New System.Drawing.Size(142, 20)
        Me.mskTxtFax.TabIndex = 12
        '
        'lblFax
        '
        Me.lblFax.Location = New System.Drawing.Point(10, 72)
        Me.lblFax.Name = "lblFax"
        Me.lblFax.Size = New System.Drawing.Size(56, 16)
        Me.lblFax.TabIndex = 229
        Me.lblFax.Text = "Fax:"
        Me.lblFax.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtLicenseeAddress
        '
        Me.txtLicenseeAddress.Location = New System.Drawing.Point(104, 232)
        Me.txtLicenseeAddress.Multiline = True
        Me.txtLicenseeAddress.Name = "txtLicenseeAddress"
        Me.txtLicenseeAddress.ReadOnly = True
        Me.txtLicenseeAddress.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
        Me.txtLicenseeAddress.Size = New System.Drawing.Size(208, 96)
        Me.txtLicenseeAddress.TabIndex = 7
        Me.txtLicenseeAddress.Text = ""
        Me.txtLicenseeAddress.WordWrap = False
        '
        'lblLicenseeAddress
        '
        Me.lblLicenseeAddress.Location = New System.Drawing.Point(24, 224)
        Me.lblLicenseeAddress.Name = "lblLicenseeAddress"
        Me.lblLicenseeAddress.Size = New System.Drawing.Size(72, 32)
        Me.lblLicenseeAddress.TabIndex = 264
        Me.lblLicenseeAddress.Text = "Licensee Address"
        Me.lblLicenseeAddress.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnLabels
        '
        Me.btnLabels.Location = New System.Drawing.Point(32, 304)
        Me.btnLabels.Name = "btnLabels"
        Me.btnLabels.Size = New System.Drawing.Size(64, 23)
        Me.btnLabels.TabIndex = 1060
        Me.btnLabels.Text = "Labels"
        '
        'btnEnvelopes
        '
        Me.btnEnvelopes.Location = New System.Drawing.Point(32, 272)
        Me.btnEnvelopes.Name = "btnEnvelopes"
        Me.btnEnvelopes.Size = New System.Drawing.Size(65, 23)
        Me.btnEnvelopes.TabIndex = 1059
        Me.btnEnvelopes.Text = "Envelopes"
        '
        'lblEmail
        '
        Me.lblEmail.Location = New System.Drawing.Point(40, 200)
        Me.lblEmail.Name = "lblEmail"
        Me.lblEmail.Size = New System.Drawing.Size(56, 16)
        Me.lblEmail.TabIndex = 222
        Me.lblEmail.Text = "E-mail:"
        Me.lblEmail.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtEmail
        '
        Me.txtEmail.Location = New System.Drawing.Point(104, 200)
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.Size = New System.Drawing.Size(208, 20)
        Me.txtEmail.TabIndex = 5
        Me.txtEmail.Text = ""
        '
        'lblSuffix
        '
        Me.lblSuffix.Location = New System.Drawing.Point(24, 168)
        Me.lblSuffix.Name = "lblSuffix"
        Me.lblSuffix.Size = New System.Drawing.Size(72, 16)
        Me.lblSuffix.TabIndex = 190
        Me.lblSuffix.Text = "Suffix:"
        Me.lblSuffix.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblLastName
        '
        Me.lblLastName.Location = New System.Drawing.Point(16, 136)
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
        Me.cmbSuffix.Location = New System.Drawing.Point(104, 168)
        Me.cmbSuffix.Name = "cmbSuffix"
        Me.cmbSuffix.Size = New System.Drawing.Size(48, 21)
        Me.cmbSuffix.TabIndex = 4
        '
        'txtLastName
        '
        Me.txtLastName.Location = New System.Drawing.Point(104, 136)
        Me.txtLastName.Name = "txtLastName"
        Me.txtLastName.Size = New System.Drawing.Size(176, 20)
        Me.txtLastName.TabIndex = 3
        Me.txtLastName.Text = ""
        '
        'lblMiddleName
        '
        Me.lblMiddleName.Location = New System.Drawing.Point(16, 112)
        Me.lblMiddleName.Name = "lblMiddleName"
        Me.lblMiddleName.Size = New System.Drawing.Size(80, 16)
        Me.lblMiddleName.TabIndex = 189
        Me.lblMiddleName.Text = "Middle Name:"
        Me.lblMiddleName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblFirstName
        '
        Me.lblFirstName.Location = New System.Drawing.Point(24, 72)
        Me.lblFirstName.Name = "lblFirstName"
        Me.lblFirstName.Size = New System.Drawing.Size(72, 16)
        Me.lblFirstName.TabIndex = 188
        Me.lblFirstName.Text = "First Name:"
        Me.lblFirstName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblPersonTitle
        '
        Me.lblPersonTitle.Location = New System.Drawing.Point(24, 40)
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
        Me.lblLicenseNumber.Size = New System.Drawing.Size(88, 16)
        Me.lblLicenseNumber.TabIndex = 193
        Me.lblLicenseNumber.Text = "License #:"
        Me.lblLicenseNumber.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtLicenseNo
        '
        Me.txtLicenseNo.Location = New System.Drawing.Point(104, 8)
        Me.txtLicenseNo.Name = "txtLicenseNo"
        Me.txtLicenseNo.ReadOnly = True
        Me.txtLicenseNo.Size = New System.Drawing.Size(96, 20)
        Me.txtLicenseNo.TabIndex = 100
        Me.txtLicenseNo.Text = ""
        '
        'cmbTitle
        '
        Me.cmbTitle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbTitle.ItemHeight = 13
        Me.cmbTitle.Items.AddRange(New Object() {"", "Mr", "Mrs", "Ms", "Dr", "Sir"})
        Me.cmbTitle.Location = New System.Drawing.Point(72, 40)
        Me.cmbTitle.Name = "cmbTitle"
        Me.cmbTitle.Size = New System.Drawing.Size(48, 21)
        Me.cmbTitle.TabIndex = 0
        '
        'txtFirstName
        '
        Me.txtFirstName.Location = New System.Drawing.Point(104, 72)
        Me.txtFirstName.Name = "txtFirstName"
        Me.txtFirstName.Size = New System.Drawing.Size(176, 20)
        Me.txtFirstName.TabIndex = 1
        Me.txtFirstName.Text = ""
        '
        'txtMiddleName
        '
        Me.txtMiddleName.Location = New System.Drawing.Point(104, 104)
        Me.txtMiddleName.Name = "txtMiddleName"
        Me.txtMiddleName.Size = New System.Drawing.Size(176, 20)
        Me.txtMiddleName.TabIndex = 2
        Me.txtMiddleName.Text = ""
        '
        'pnlLicenseeInfoHead
        '
        Me.pnlLicenseeInfoHead.Controls.Add(Me.lblLicenseeInfo)
        Me.pnlLicenseeInfoHead.Controls.Add(Me.lblLicenseeInfoDisplay)
        Me.pnlLicenseeInfoHead.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlLicenseeInfoHead.Location = New System.Drawing.Point(0, 0)
        Me.pnlLicenseeInfoHead.Name = "pnlLicenseeInfoHead"
        Me.pnlLicenseeInfoHead.Size = New System.Drawing.Size(736, 24)
        Me.pnlLicenseeInfoHead.TabIndex = 1064
        '
        'lblLicenseeInfo
        '
        Me.lblLicenseeInfo.BackColor = System.Drawing.SystemColors.Highlight
        Me.lblLicenseeInfo.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblLicenseeInfo.ForeColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.lblLicenseeInfo.Location = New System.Drawing.Point(16, 0)
        Me.lblLicenseeInfo.Name = "lblLicenseeInfo"
        Me.lblLicenseeInfo.Size = New System.Drawing.Size(720, 24)
        Me.lblLicenseeInfo.TabIndex = 1
        Me.lblLicenseeInfo.Text = "Licensee Info"
        Me.lblLicenseeInfo.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblLicenseeInfoDisplay
        '
        Me.lblLicenseeInfoDisplay.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLicenseeInfoDisplay.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblLicenseeInfoDisplay.Location = New System.Drawing.Point(0, 0)
        Me.lblLicenseeInfoDisplay.Name = "lblLicenseeInfoDisplay"
        Me.lblLicenseeInfoDisplay.Size = New System.Drawing.Size(16, 24)
        Me.lblLicenseeInfoDisplay.TabIndex = 0
        Me.lblLicenseeInfoDisplay.Text = "-"
        Me.lblLicenseeInfoDisplay.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'chkComplianceManager
        '
        Me.chkComplianceManager.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkComplianceManager.ImageAlign = System.Drawing.ContentAlignment.TopLeft
        Me.chkComplianceManager.Location = New System.Drawing.Point(128, 40)
        Me.chkComplianceManager.Name = "chkComplianceManager"
        Me.chkComplianceManager.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkComplianceManager.Size = New System.Drawing.Size(152, 21)
        Me.chkComplianceManager.TabIndex = 1064
        Me.chkComplianceManager.Tag = "649"
        Me.chkComplianceManager.Text = "ComplianceManager"
        Me.chkComplianceManager.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Licensees
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(752, 597)
        Me.Controls.Add(Me.pnlLicensees)
        Me.Controls.Add(Me.pnlLicenseesBottom)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.lblLicenseNo)
        Me.Name = "Licensees"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Licensees"
        Me.pnlLicenseesBottom.ResumeLayout(False)
        Me.pnlLicensees.ResumeLayout(False)
        Me.pnlCoursesHead.ResumeLayout(False)
        Me.pnlCourses.ResumeLayout(False)
        CType(Me.ugCourses, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlCoursesBottom.ResumeLayout(False)
        Me.pnlTestsHead.ResumeLayout(False)
        Me.pnlTests.ResumeLayout(False)
        CType(Me.ugTests, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlTestsBottom.ResumeLayout(False)
        Me.pnlPriorCompaniesHead.ResumeLayout(False)
        Me.pnlPriorCompanies.ResumeLayout(False)
        CType(Me.ugPriorCompanies, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlLCEHead.ResumeLayout(False)
        Me.pnlLCE.ResumeLayout(False)
        CType(Me.ugLicenseeComplianceEvents, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlDocumentsHead.ResumeLayout(False)
        Me.pnlDocuments.ResumeLayout(False)
        Me.pnlLicenseeInfo.ResumeLayout(False)
        Me.pnlCompany.ResumeLayout(False)
        Me.pnlExtension.ResumeLayout(False)
        Me.pnlPhone.ResumeLayout(False)
        CType(Me.mskTxtPhone1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtCell, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtPhone2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.mskTxtFax, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlLicenseeInfoHead.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Form Events"
    Private Sub Licensees_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Not bolLoading Then bolLoading = True
        Try
            If pLicen Is Nothing Then
                pLicen = New MUSTER.BusinessLogic.pLicensee
            End If
            If pCompanyLicenseeAssociation Is Nothing Then
                pCompanyLicenseeAssociation = New MUSTER.BusinessLogic.pCompanyLicensee
            End If
            UIUtilsGen.CreateEmptyFormatDatePicker(dtAppRcvdDate)
            UIUtilsGen.CreateEmptyFormatDatePicker(dtExceptionGrantedDate)
            UIUtilsGen.CreateEmptyFormatDatePicker(dtExtensionDeadlineDate)
            UIUtilsGen.CreateEmptyFormatDatePicker(dtLicenseExpirationDate)
            UIUtilsGen.CreateEmptyFormatDatePicker(dtIssuedDate)
            UIUtilsGen.CreateEmptyFormatDatePicker(dtOriginalIssueDate)
            txtCompany.Text = strCompanyName
            txtCompanyTitle.Text = strCompanyAddress
            PopulateLicenseeHireStatus()
            PopulateLicenseeStatus()
            PopulateLicenseeCertificationType()
            If nLicenseeID > 0 Or nLicenseeID < -100 Then
                Dim oLicenseeInfo As MUSTER.Info.LicenseeInfo
                oLicenseeInfo = pLicen.Retrieve(nLicenseeID)
                licenseeInfo = oLicenseeInfo
                populateLicenseeInfo(oLicenseeInfo)
                pLicen.pLicenseeCourse.GetAll(nLicenseeID)
                pLicen.pLicenseeCourseTest.GetAll(nLicenseeID)
                PopulatePriorCompanies()
                PopulateDocuments()
                btnFlags.Enabled = True
                btnComments.Enabled = True
            Else
                btnFlags.Enabled = False
                btnComments.Enabled = False
                If strMode = "ADD" Then
                    licenseeInfo = New MUSTER.Info.LicenseeInfo
                    pLicen.Add(licenseeInfo)
                    populateLicenseeInfo(licenseeInfo, False)
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
            ugCourses.DisplayLayout.ValueLists.Clear()
            ugCourses.DisplayLayout.ValueLists.Add("CourseType")
            ugCourses.DisplayLayout.ValueLists.Add("ProviderName")
            FillLicenseeCourseGrid()

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
    Private Sub Licensees_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        ' The following line enables all forms to detect the visible form in the MDI container
        '
        If Not bolModelWindow Then
            MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "Company")
        End If
    End Sub
    Private Sub Licensees_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.TextChanged
        If Not bolModelWindow Then
            MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "WindowName", Me.Text, "Company")
        End If
    End Sub
    Private Sub Licensees_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
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
    Private Sub populateLicenseeInfo(ByVal oLicInfo As MUSTER.Info.LicenseeInfo, Optional ByVal loadImages As Boolean = True)
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

            dtIssuedDate.Enabled = True
            'dtExtensionDeadlineDate.Enabled = True
            'If oLicInfo.StatusDesc = "NOT CURRENTLY CERTIFIED" Then
            '    chkOverrideExpiration.Enabled = True
            'End If

            If nLicenseeID > 0 Or nLicenseeID < -100 Then
                oLicAddInfo = oLicAdd.GetAddressByType(0, 0, nLicenseeID, 0)
                txtLicenseeAddress.Text = oLicAddInfo.AddressLine1 + IIf(oLicAddInfo.AddressLine1.Trim.Length = 0, "", vbCrLf + oLicAddInfo.AddressLine2) + vbCrLf + oLicAddInfo.City + ", " + oLicAddInfo.State + " " + oLicAddInfo.Zip
            Else
                oLicAddInfo = New MUSTER.Info.ComAddressInfo
                txtLicenseeAddress.Text = String.Empty
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
            cmbHireStatus.Text = oLicInfo.HIRE_STATUS
            strLicenseeHireStatus = oLicInfo.HIRE_STATUS
            chkEmployeeLetter.Checked = oLicInfo.EMPLOYEE_LETTER
            pLicen.ISLICENSEE = 1
            chkComplianceManager.Checked = oLicInfo.COMPLIANCEMANAGER
            EnableDisableEmployeLetter()
            Me.Text = "Licensee - " + txtFirstName.Text + " " + txtLastName.Text
            'chkOverrideExpiration.Checked = oLicInfo.OVERRIDE_EXPIRE
            If oLicInfo.STATUS > 0 Then
                UIUtilsGen.SetComboboxItemByValue(cmbStatus, oLicInfo.STATUS)
            Else
                cmbStatus.SelectedIndex = 0
                'If cmbStatus.SelectedIndex <> 0 Then cmbStatus.SelectedIndex = 0
            End If
            'cmbStatus.Text = oLicInfo.StatusDesc
            'txtStatus.Text = oLicInfo.STATUS
            If oLicInfo.CertType > 0 Then
                UIUtilsGen.SetComboboxItemByValue(cmbCertificationType, oLicInfo.CertType)
            Else
                cmbCertificationType.SelectedIndex = 0
                'If cmbCertificationType.SelectedIndex <> 0 Then cmbCertificationType.SelectedIndex = 0
            End If
            'cmbCertificationType.Text = oLicInfo.CertTypeDesc
            'txtCertificationType.Text = oLicInfo.CERT_TYPE
            UIUtilsGen.SetDatePickerValue(dtAppRcvdDate, oLicInfo.APP_RECVD_DATE)
            UIUtilsGen.SetDatePickerValue(dtIssuedDate, oLicInfo.ISSUED_DATE)
            UIUtilsGen.SetDatePickerValue(dtOriginalIssueDate, oLicInfo.ORIGIN_ISSUED_DATE)
            UIUtilsGen.SetDatePickerValue(dtExtensionDeadlineDate, oLicInfo.EXTENSION_DEADLINE_DATE)
            If oLicInfo.LICENSE_EXPIRE_DATE = "" Then
                UIUtilsGen.SetDatePickerValue(dtLicenseExpirationDate, CDate("01/01/0001"))
            Else
                UIUtilsGen.SetDatePickerValue(dtLicenseExpirationDate, oLicInfo.LICENSE_EXPIRE_DATE)
            End If

            UIUtilsGen.SetDatePickerValue(dtExceptionGrantedDate, oLicInfo.EXCEPT_GRANT_DATE)
            If (nCompanyID > 0 Or nCompanyID < -100) And strMode = "Search" Then
                Dim currentControl As Control
                Dim myEnumerator As System.Collections.IEnumerator = pnlLicensees.Controls.GetEnumerator()
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
                ugCourses.Enabled = False
                ugTests.Enabled = False
                ugPriorCompanies.Enabled = False
                ugLicenseeComplianceEvents.Enabled = False
                pnlCompany.Enabled = False
                btnSaveLicensee.Enabled = False
                btnCancel.Name = "Close"
                btnComments.Enabled = False
                btnFlags.Enabled = False
                btnDeleteLicensee.Enabled = False
            End If
            EnableDisableCertifyButton()
            EnableDisableGenExpirationLetterButton()
            EnableDisableGenReminderLetterButton()
            'lblLicenseeID.Text = oLicInfo.ID
            If (nCompanyID > 0 Or nCompanyID < -100) And (pLicen.ID > 0 Or pLicen.ID < -100) Then
                CommentsMaintenance(, , True)
            ElseIf loadImages Then
                CommentsMaintenance(, , True, True)
            End If

            If Not bolModelWindow Then
                MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "LicenseeID", oLicInfo.ID, "Company")
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
        If cmbHireStatus.Text.EndsWith("Employee") Then
            chkEmployeeLetter.Enabled = True
        Else
            chkEmployeeLetter.Enabled = False
        End If
    End Sub
    Private Sub EnableDisableCertifyButton()
        If pLicen.STATUS_DESC.ToUpper = "CERTIFIED" Then
            btnCertify.Enabled = True
        Else
            btnCertify.Enabled = False
        End If
    End Sub
    Private Sub EnableDisableGenExpirationLetterButton()
        btnGenExpirationLetter.Enabled = False

        If Date.Compare(IIf(pLicen.LICENSE_EXPIRE_DATE = "", CDate("01/01/0001"), pLicen.LICENSE_EXPIRE_DATE), Today.Date) <= 0 Then
            btnGenExpirationLetter.Enabled = True
        End If


    End Sub
    Private Sub EnableDisableGenReminderLetterButton()
        btnGenReminderLetter.Enabled = False

        If pLicen.STATUS_DESC.ToUpper = "CERTIFIED" Then
            If Date.Compare(IIf(pLicen.LICENSE_EXPIRE_DATE = "", CDate("01/01/0001"), pLicen.LICENSE_EXPIRE_DATE), DateAdd(DateInterval.Day, -90, Today.Date)) > 0 Then
                btnGenReminderLetter.Enabled = True
            End If
        End If

       


    End Sub

    Private Sub PopulateLicenseNumber()

        With txtLicenseNo
            .Text = pLicen.LICENSEE_NUMBER_PREFIX + pLicen.LICENSEE_NUMBER.ToString

            If .Text.ToUpper.IndexOf("RX") > -1 Then
                Me.lblRestCert.Visible = True
            Else
                Me.lblRestCert.Visible = False
            End If
        End With
    End Sub
    Private Sub PopulateLicenseeHireStatus()
        cmbHireStatus.DataSource = pLicen.GetLicenseeHireStatus(False, True).Tables(0)
        cmbHireStatus.DisplayMember = "PROPERTY_NAME"
        cmbHireStatus.ValueMember = "PROPERTY_NAME"
    End Sub
    Private Sub PopulateLicenseeStatus()
        cmbStatus.DataSource = pLicen.GetLicenseeStatus(False, True).Tables(0)
        cmbStatus.DisplayMember = "PROPERTY_NAME"
        cmbStatus.ValueMember = "PROPERTY_ID"
    End Sub
    Private Sub PopulateLicenseeCertificationType()
        cmbCertificationType.DataSource = pLicen.GetLicenseeCertificationType(False, True).Tables(0)
        cmbCertificationType.DisplayMember = "PROPERTY_NAME"
        cmbCertificationType.ValueMember = "PROPERTY_ID"
    End Sub

    Private Sub Cancel()
        pLicen.Reset()
        For Each oComAddInfo In oLicAdd.ColCompanyAddresses.Values
            If oComAddInfo.CompanyId = pLicen.ID Then
                oComAddInfo.Reset()
            End If
        Next
        For Each oLicCourseInfo In pLicen.pLicenseeCourse.colLicCourse.Values
            If oLicCourseInfo.LicenseeID = pLicen.ID Then
                oLicCourseInfo.Reset()
            End If
        Next
        For Each oLicTestInfo In pLicen.pLicenseeCourseTest.colLicCourseTest.Values
            If oLicTestInfo.LicenseeID = pLicen.ID Then
                oLicTestInfo.Reset()
            End If
        Next
    End Sub
#End Region

#Region "Licensee Info"
    Private Sub cmbTitle_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbTitle.SelectedValueChanged
        If bolLoading Then Exit Sub
        pLicen.TITLE = cmbTitle.Text
    End Sub
    Private Sub txtFirstName_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFirstName.TextChanged
        If bolLoading Then Exit Sub
        pLicen.FIRST_NAME = txtFirstName.Text
    End Sub
    Private Sub txtLicenseeAddress_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLicenseeAddress.TextChanged
        If bolLoading Then Exit Sub
        If Not Me.txtLicenseeAddress.Tag Is Nothing AndAlso TypeOf Me.txtLicenseeAddress.Tag Is Integer Then
            nLicAddressID = Me.txtLicenseeAddress.Tag
            pLicen.IsDirty = True
        End If


    End Sub
    Private Sub txtMiddleName_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMiddleName.TextChanged
        If bolLoading Then Exit Sub
        pLicen.MIDDLE_NAME = txtMiddleName.Text
    End Sub
    Private Sub txtLastName_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLastName.TextChanged
        If bolLoading Then Exit Sub
        pLicen.LAST_NAME = txtLastName.Text
    End Sub
    Private Sub cmbSuffix_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbSuffix.SelectedValueChanged
        If bolLoading Then Exit Sub
        pLicen.SUFFIX = cmbSuffix.Text
    End Sub
    Private Sub txtEmail_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEmail.TextChanged
        If bolLoading Then Exit Sub
        pLicen.EMAIL_ADDRESS = txtEmail.Text
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

    Private Sub txtLicenseeAddress_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLicenseeAddress.DoubleClick
        Try
            If txtLicenseeAddress.Text = "" Then
                objAddMaster = New AddressMaster(oLicAdd, nLicAddressID, nLicenseeID, "Licensee", "ADD")
            Else
                objAddMaster = New AddressMaster(oLicAdd, nLicAddressID, nLicenseeID, "Licensee", "MODIFY")
            End If
            Me.Update()
            LockWindowUpdate(Me.Handle.ToInt64)
            objAddMaster.ShowDialog()
            LockWindowUpdate(0)
            txtLicenseeAddress.Tag = oLicAdd.AddressId


            Me.txtLicenseeAddress.Text = oLicAdd.AddressLine1 + IIf(oLicAdd.AddressLine1.Trim.Length = 0, "", vbCrLf + oLicAdd.AddressLine2) + IIf(oLicAdd.City.Length = 0, "", vbCrLf + oLicAdd.City) + IIf(oLicAdd.State.Length = 0, "", ", " + oLicAdd.State) + IIf(oLicAdd.Zip.Length = 0, "", " " + oLicAdd.Zip)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub cmbHireStatus_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbHireStatus.SelectedIndexChanged
        If bolLoading Then Exit Sub
        pLicen.HIRE_STATUS = cmbHireStatus.Text
        EnableDisableEmployeLetter()
    End Sub
    Private Sub chkEmployeeLetter_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkEmployeeLetter.CheckedChanged
        If bolLoading Then Exit Sub
        pLicen.EMPLOYEE_LETTER = chkEmployeeLetter.Checked
    End Sub
    Private Sub chkComplianceManager_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkComplianceManager.CheckedChanged
        If bolLoading Then Exit Sub
        pLicen.COMPLIANCEMANAGER = chkComplianceManager.Checked
    End Sub
    Private Sub cmbStatus_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbStatus.SelectedIndexChanged
        If Not bolLoading Then
            pLicen.STATUS_ID = UIUtilsGen.GetComboBoxValueInt(cmbStatus)
            pLicen.STATUS_DESC = UIUtilsGen.GetComboBoxText(cmbStatus)
        End If


    End Sub
    Private Sub cmbCertificationType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbCertificationType.SelectedIndexChanged
        If Not bolLoading Then
            pLicen.CERT_TYPE_ID = UIUtilsGen.GetComboBoxValueInt(cmbCertificationType)
            pLicen.CERT_TYPE_DESC = UIUtilsGen.GetComboBoxText(cmbCertificationType)
        End If
    End Sub
    Private Sub dtAppRcvdDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtAppRcvdDate.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(Me.dtAppRcvdDate)
        UIUtilsGen.FillDateobjectValues(pLicen.APP_RECVD_DATE, dtAppRcvdDate.Text)

        Me.EnableDisableGenExpirationLetterButton()
        Me.EnableDisableGenReminderLetterButton()
    End Sub
    Private Sub dtIssuedDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtIssuedDate.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(Me.dtIssuedDate)
        UIUtilsGen.FillDateobjectValues(pLicen.ISSUED_DATE, dtIssuedDate.Text)

    End Sub
    Private Sub dtExtensionDeadlineDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtExtensionDeadlineDate.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(dtExtensionDeadlineDate)
        UIUtilsGen.FillDateobjectValues(pLicen.EXTENSION_DEADLINE_DATE, dtExtensionDeadlineDate.Text)


    End Sub
    Private Sub dtLicenseExpirationDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtLicenseExpirationDate.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(Me.dtLicenseExpirationDate)
        UIUtilsGen.FillDateobjectValues(pLicen.LICENSE_EXPIRE_DATE, dtLicenseExpirationDate.Text)


    End Sub
    Private Sub dtOriginalIssueDate_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtOriginalIssueDate.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(Me.dtOriginalIssueDate)
        UIUtilsGen.FillDateobjectValues(pLicen.ORGIN_ISSUED_DATE, dtOriginalIssueDate.Text)


    End Sub
    Private Sub dtExceptionGrantedDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtExceptionGrantedDate.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(Me.dtExceptionGrantedDate)
        UIUtilsGen.FillDateobjectValues(pLicen.EXCEPT_GRANT_DATE, dtExceptionGrantedDate.Text)


    End Sub
    'Private Sub chkOverrideExpiration_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkOverrideExpiration.CheckedChanged
    '    If bolLoading Then Exit Sub
    '    pLicen.OVERRIDE_EXPIRE = chkOverrideExpiration.Checked
    'End Sub

    'Private Sub btnCompanyTitleSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCompanyTitleSearch.Click
    '    Dim tempCompanyAddInfo As MUSTER.Info.ComAddressInfo
    '    Try
    '        objAddresses = New Addresses(oCompanyAdd, "Licensee", nAssociationID, pCompanyLicenseeAssociation, , nCompanyID)
    '        objAddresses.ShowDialog()
    '        If nAssociationID > 0 Or nAssociationID < -100 Then
    '            tempCompanyAddInfo = oCompanyAdd.Retrieve(pCompanyLicenseeAssociation.ComLicCollection.Item(nAssociationID).ComLicAddressID, False)
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
    '        ncompAddressID = pCompanyLicenseeAssociation.ComLicAddressID
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub objAddresses_evtCompLicAssocChanged(ByVal AddressID As Integer) Handles objAddresses.evtCompLicAssocChanged
    '    Try
    '        If nAssociationID > 0 Or nAssociationID < -100 Then
    '            Dim comlicinfo As MUSTER.Info.CompanyLicenseeInfo
    '            For Each comlicinfo In pCompanyLicenseeAssociation.ComLicCollection.Values
    '                If comlicinfo.ID = nAssociationID Then
    '                    comlicinfo.ComLicAddressID = AddressID
    '                End If
    '            Next
    '        Else
    '            pCompanyLicenseeAssociation.ComLicAddressID = AddressID
    '        End If
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
#End Region

#Region "Courses"
    Private Sub FillLicenseeCourseGrid()
        Try
            pLicen.pLicenseeCourse.CourseTable.DefaultView.Sort = "Date"
            ugCourses.DataSource = Nothing
            ugCourses.DataSource = pLicen.pLicenseeCourse.CourseTable.DefaultView

            ugCourses.DisplayLayout.Bands(0).Columns("CourseID").Hidden = True
            ugCourses.DisplayLayout.Bands(0).Columns("LicenseeID").Hidden = True
            ugCourses.DisplayLayout.Bands(0).Columns("Deleted").Hidden = True
            ugCourses.DisplayLayout.Bands(0).Columns("Created By").Hidden = True
            ugCourses.DisplayLayout.Bands(0).Columns("Date Created").Hidden = True
            ugCourses.DisplayLayout.Bands(0).Columns("Last Edited By").Hidden = True
            ugCourses.DisplayLayout.Bands(0).Columns("Date Last Edited").Hidden = True
            ugCourses.DisplayLayout.Bands(0).Columns("Deleted").CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled

            PopulateCourseType(ugCourses)
            PopulateProvider()
            ugCourses.DisplayLayout.Appearance.BackColor = System.Drawing.Color.White
            ugCourses.DisplayLayout.Bands(0).Columns("Date").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub PopulateCourseType(ByRef ugGrid As Infragistics.Win.UltraWinGrid.UltraGrid)
        Try
            Dim dtCourseType As DataTable = pLicen.pLicenseeCourse.ListCourseTypes(True)
            Dim drRow As DataRow

            ugGrid.DisplayLayout.ValueLists("CourseType").ValueListItems.Clear()
            ugGrid.DisplayLayout.ValueLists("CourseType").DropDownListAlignment = Infragistics.Win.DropDownListAlignment.Left

            For Each drRow In dtCourseType.Rows
                ugGrid.DisplayLayout.ValueLists("CourseType").ValueListItems.Add(drRow.Item("PROPERTY_ID"), drRow.Item("PROPERTY_NAME").ToString)
            Next
            ugGrid.DisplayLayout.Bands(0).Columns("Type").ValueList = ugGrid.DisplayLayout.ValueLists("CourseType")
            ugGrid.DisplayLayout.Bands(0).Columns("Type").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate
            ugGrid.DisplayLayout.ValueLists("CourseType").DisplayStyle = Infragistics.Win.ValueListDisplayStyle.DisplayText
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub PopulateProvider()
        Try
            Dim dtProvider As DataTable = pLicen.pLicenseeCourse.ListProviders
            Dim drRow As DataRow

            ugCourses.DisplayLayout.ValueLists("ProviderName").ValueListItems.Clear()
            ugCourses.DisplayLayout.ValueLists("ProviderName").DropDownListAlignment = Infragistics.Win.DropDownListAlignment.Left
            If Not dtProvider Is Nothing Then
                For Each drRow In dtProvider.Rows
                    ugCourses.DisplayLayout.ValueLists("ProviderName").ValueListItems.Add(drRow.Item("Provider_ID"), drRow.Item("Abbrev").ToString)
                Next
                ugCourses.DisplayLayout.Bands(0).Columns("Provider").ValueList = ugCourses.DisplayLayout.ValueLists("ProviderName")
                ugCourses.DisplayLayout.Bands(0).Columns("Provider").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.DropDownValidate
                ugCourses.DisplayLayout.ValueLists("ProviderName").DisplayStyle = Infragistics.Win.ValueListDisplayStyle.DisplayText
                ugCourses.DisplayLayout.Bands(0).Columns("Provider").Width = 150
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Sub

    Private Sub btnCourseAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCourseAdd.Click
        Try
            If Not bolValidatationFailed Then
                ugCourses.DisplayLayout.Bands(0).AddNew()
                ugCourses.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True
            Else
                MsgBox("Enter Course Provider")
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
    Private Sub btnCourseDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCourseDelete.Click
        Try
            Dim msgResult As MsgBoxResult
            If ugCourses.ActiveRow Is Nothing Then
                MsgBox("Select row to Delete")
                Exit Sub
            End If

            msgResult = MsgBox("Do you want to Delete this row.?", MsgBoxStyle.YesNo, "Licensee Course")
            If msgResult = MsgBoxResult.No Then
                Exit Sub
            Else
                'Remove the Course Entry.
                If CInt(ugCourses.ActiveRow.Cells("CourseID").Text) <= 0 Then
                    pLicen.pLicenseeCourse.Remove(Integer.Parse(ugCourses.ActiveRow.Cells("CourseID").Text))
                Else
                    pLicen.pLicenseeCourse.Retrieve(Integer.Parse(ugCourses.ActiveRow.Cells("CourseID").Text))
                    pLicen.pLicenseeCourse.Deleted = True
                    If pLicen.pLicenseeCourse.ID <= 0 And pLicen.pLicenseeCourse.ID > -100 Then
                        pLicen.pLicenseeCourse.CreatedBy = MusterContainer.AppUser.ID
                    Else
                        pLicen.pLicenseeCourse.ModifiedBy = MusterContainer.AppUser.ID
                    End If
                    pLicen.pLicenseeCourse.Save(CType(UIUtilsGen.ModuleID.Company, Integer), MusterContainer.AppUser.UserKey, returnVal, False)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If

                    pLicen.pLicenseeCourse.Remove(Integer.Parse(ugCourses.ActiveRow.Cells("CourseID").Text))

                End If
                ugCourses.ActiveRow.Delete(False)
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ugCourses_AfterRowInsert(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles ugCourses.AfterRowInsert
        Try
            pLicen.pLicenseeCourse.Retrieve(0)
            e.Row.Cells("CourseID").Value = pLicen.pLicenseeCourse.ID
            Me.ugCourses.ActiveCell = e.Row.Cells("Type")
            Me.ugCourses.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.ToggleCellSel, False, False)
            Me.ugCourses.PerformAction(Infragistics.Win.UltraWinGrid.UltraGridAction.EnterEditMode, False, False)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugCourses_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles ugCourses.AfterRowUpdate
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            ugrow = ugCourses.ActiveRow

            pLicen.pLicenseeCourse.Retrieve(ugrow.Cells("CourseID").Value)

            pLicen.pLicenseeCourse.LicenseeID = nLicenseeID

            If ugrow.Cells("Type").Text <> String.Empty Then
                pLicen.pLicenseeCourse.CourseTypeID = ugrow.Cells("Type").Value
            Else
                pLicen.pLicenseeCourse.CourseTypeID = 0
            End If

            If ugrow.Cells("Provider").Text <> String.Empty Then
                pLicen.pLicenseeCourse.ProviderID = ugrow.Cells("Provider").Value
            Else
                pLicen.pLicenseeCourse.ProviderID = 0
            End If

            If ugrow.Cells("Date").Text <> String.Empty Then
                pLicen.pLicenseeCourse.CourseDate = ugrow.Cells("Date").Value
            Else
                pLicen.pLicenseeCourse.CourseDate = dtNull
            End If

            pLicen.pLicenseeCourse.Deleted = ugrow.Cells("Deleted").Value

            If pLicen.pLicenseeCourse.ID <= 0 Then
                pLicen.pLicenseeCourse.CreatedBy = MusterContainer.AppUser.ID
            Else
                pLicen.pLicenseeCourse.ModifiedBy = MusterContainer.AppUser.ID
            End If

            ugCourses.DisplayLayout.Bands(0).Columns("Date").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugCourses_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles ugCourses.BeforeRowUpdate
        Try
            If e.Row.Cells("Provider").Text = String.Empty Then
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
            ugTests.DataSource = pLicen.pLicenseeCourseTest.TestTable
            ugTests.DisplayLayout.Bands(0).Columns("TestID").Hidden = True
            ugTests.DisplayLayout.Bands(0).Columns("LicenseeID").Hidden = True
            ugTests.DisplayLayout.Bands(0).Columns("Deleted").Hidden = True
            ugTests.DisplayLayout.Bands(0).Columns("Created By").Hidden = True
            ugTests.DisplayLayout.Bands(0).Columns("Date Created").Hidden = True
            ugTests.DisplayLayout.Bands(0).Columns("Last Edited By").Hidden = True
            ugTests.DisplayLayout.Bands(0).Columns("Date Last Edited").Hidden = True
            ugTests.DisplayLayout.Bands(0).Columns("StartTime").Hidden = True
            ugTests.DisplayLayout.Bands(0).Columns("Deleted").CellActivation = Infragistics.Win.UltraWinGrid.Activation.Disabled

            PopulateCourseType(ugTests)
            'PopulateStartTime()
            ugTests.DisplayLayout.Appearance.BackColor = System.Drawing.Color.White
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    'Private Sub PopulateStartTime()
    '    Try
    '        Dim dtProvider As DataTable = pLicen.pLicenseeCourseTest.ListStartTime
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

            msgResult = MsgBox("Do you want to Delete this row.?", MsgBoxStyle.YesNo, "Licensee Course Tests")
            If msgResult = MsgBoxResult.No Then
                Exit Sub
            Else
                'Remove the Course Entry.
                If CInt(ugTests.ActiveRow.Cells("TestID").Text) <= 0 Then
                    pLicen.pLicenseeCourseTest.Remove(Integer.Parse(ugTests.ActiveRow.Cells("TestID").Text))
                Else
                    pLicen.pLicenseeCourseTest.Retrieve(Integer.Parse(ugTests.ActiveRow.Cells("TestID").Text))
                    pLicen.pLicenseeCourseTest.Deleted = True
                    pLicen.pLicenseeCourseTest.ModifiedBy = MusterContainer.AppUser.ID
                    pLicen.pLicenseeCourseTest.Save(CType(UIUtilsGen.ModuleID.Company, Integer), MusterContainer.AppUser.UserKey, returnVal, False)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If


                    pLicen.pLicenseeCourseTest.Remove(Integer.Parse(ugTests.ActiveRow.Cells("TestID").Text))

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
            pLicen.pLicenseeCourseTest.Retrieve(0)
            e.Row.Cells("TestID").Value = pLicen.pLicenseeCourseTest.ID
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
            pLicen.pLicenseeCourseTest.Retrieve(ugrow.Cells("TestID").Value)

            pLicen.pLicenseeCourseTest.LicenseeID = nLicenseeID

            If ugrow.Cells("Type").Text <> String.Empty Then
                pLicen.pLicenseeCourseTest.CourseTypeID = ugrow.Cells("Type").Value
            Else
                pLicen.pLicenseeCourse.CourseTypeID = 0
            End If

            'If ugrow.Cells("StartTime").Text <> String.Empty Then
            '    pLicen.pLicenseeCourseTest.StartTime = ugrow.Cells("StartTime").Text
            'Else
            '    pLicen.pLicenseeCourseTest.StartTime = String.Empty
            'End If

            If ugrow.Cells("Score").Text <> String.Empty Then
                pLicen.pLicenseeCourseTest.TestScore = ugrow.Cells("Score").Value
            Else
                pLicen.pLicenseeCourseTest.TestScore = 0
            End If

            pLicen.pLicenseeCourseTest.TestDate = ugrow.Cells("Date").Value

            If pLicen.pLicenseeCourseTest.ID <= 0 Then
                pLicen.pLicenseeCourseTest.CreatedBy = MusterContainer.AppUser.ID
            Else
                pLicen.pLicenseeCourseTest.ModifiedBy = MusterContainer.AppUser.ID
            End If


            pLicen.pLicenseeCourse.Deleted = ugrow.Cells("Deleted").Value


            ugTests.DisplayLayout.Bands(0).Columns("TestID").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            btnTestAdd.Focus()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugTests_BeforeRowUpdate(ByVal sender As System.Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles ugTests.BeforeRowUpdate
        Try
            If Not validateLicenseeTests(e.Row) Then
                e.Cancel = True
                pLicen.pLicenseeCourseTest.Remove(pLicen.pLicenseeCourseTest.ID)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "Prior Companies"
    Private Sub PopulatePriorCompanies()
        ugPriorCompanies.DataSource = pLicen.GetPriorCompanies(nLicenseeID)
    End Sub
#End Region

#Region "LCE"
#End Region

#Region "Documents"
    Private Sub PopulateDocuments()
        UCLicenseeDocuments.LoadDocumentsGrid(nLicenseeID, 0, UIUtilsGen.ModuleID.Company)
    End Sub
#End Region

#Region "UI Events"
    Private Sub btnSaveLicensee_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveLicensee.Click
        Try
            SaveLicensee()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Cancel()
        Licensees_Load(sender, e)
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
    Private Sub btnDeleteLicensee_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDeleteLicensee.Click
        Dim result As DialogResult
        Try
            result = MessageBox.Show("Are you sure you want to delete this Licensee?", "Delete Licensee.", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)
            If result = DialogResult.Yes Then
                ' The following condition needs to be checked
                ' Licensee is not associated with any tank installations or closure events
                If txtCompanyTitle.Text = "" Then
                    pLicen.Retrieve(nLicenseeID)
                    pLicen.DELETED = True

                    If pLicen.ID <= 0 And pLicen.ID > -100 Then
                        pLicen.CreatedBy = MusterContainer.AppUser.ID
                    Else
                        pLicen.ModifiedBy = MusterContainer.AppUser.ID
                    End If
                    pLicen.Save(CType(UIUtilsGen.ModuleID.Company, Integer), MusterContainer.AppUser.UserKey, returnVal)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If

                    MsgBox("Licensee is deleted successfully")
                    Me.Close()
                Else
                    MsgBox("Licensee cannot be deleted, associated with a company.")
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
        If pLicen.IsDirty Then
            MsgBox("Cannot Certify before saving Licensee")
            Exit Sub
        End If
        Try
            GenerateCongLetter()
            GenerateLicenseeCertificateLetter(pLicen.LicenseeInfo)
            GenerateLicenseeCard(pLicen.LicenseeInfo)
        Catch ex As Address.NoAddressException
            MsgBox(ex.Message)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()

        End Try

    End Sub
    Private Sub btnGenExpirationLetter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGenExpirationLetter.Click
        If pLicen.IsDirty Then
            MsgBox("Cannot Generate Expiration Letter before saving Licensee")
            Exit Sub
        End If
        Dim colParams As New Specialized.NameValueCollection
        colParams = FillParameters()
        If pLicen.LicenseeInfo.CertTypeDesc.ToUpper = "INSTALL" Then
            colParams.Add("<Certification Type> ", "Install, Alter, and Permanently Close")
        ElseIf pLicen.LicenseeInfo.CertTypeDesc.ToUpper = "CLOSURE" Then
            colParams.Add("<Certification Type>", "Permanently Close")
        Else
            colParams.Add("<Certification Type>", "None")
        End If
        oLetter.GenerateLicenseeLetter(nLicenseeID, "Licensee Expired", "Expired", "Licensee Expired Letter", "ExpirationLetter.doc", colParams)
        ' per sandy's email dated jun 18, 2009
        ' change status to "not currently certified"
        pLicen.STATUS_ID = 1642
        UIUtilsGen.SetComboboxItemByValue(cmbStatus, pLicen.STATUS_ID)
        pLicen.STATUS_DESC = UIUtilsGen.GetComboBoxText(cmbStatus)

        pLicen.CERT_TYPE_ID = 0
        UIUtilsGen.SetComboboxItemByValue(cmbCertificationType, pLicen.CERT_TYPE_ID)
        pLicen.CERT_TYPE_DESC = UIUtilsGen.GetComboBoxText(cmbCertificationType)

        btnSaveLicensee_Click(sender, e)
    End Sub
    Private Sub btnGenReminderLetter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGenReminderLetter.Click
        If pLicen.IsDirty Then
            MsgBox("Cannot Generate Reminder Letter before saving Licensee")
            Exit Sub
        End If
        Dim colParams As New Specialized.NameValueCollection
        colParams = FillParameters(True)
        oLetter.GenerateLicenseeLetter(nLicenseeID, "Info Needed Letter", "InfoNeeded_Letter", "Licensee Info Needed Letter", "InfoNeededLetter.doc", colParams)
    End Sub

    Friend Function SaveLicensee() As Boolean
        Dim LicNo As String = ""
        Dim oCompanyLicenseeInfoLocal As MUSTER.Info.CompanyLicenseeInfo
        Dim bolExist As Boolean = False
        Dim bolsuccess As Boolean = False

        Try
            ' to prevent duplicate letters
            'InitializeLetterCount()

            If ValidateData() Then
                'If Not txtCompany.Text = "" Then
                '    pLicen.LicenseeLogic(licenseeInfo, FinRespExpirationDate, True, nLicenseeID, strLicenseeHireStatus)
                'Else
                '    pLicen.LicenseeLogic(licenseeInfo, FinRespExpirationDate, , nLicenseeID, strLicenseeHireStatus)
                'End If
                'If Not pLicen.Add(licenseeInfo) Then
                '    Exit Sub
                'End If

                Dim success As Boolean = False
                'have to allow for the - intergers fro migration and the - intergers from collections
                If pLicen.ID >= -100 And pLicen.ID <= 0 Then
                    pLicen.CreatedBy = MusterContainer.AppUser.ID
                Else
                    pLicen.ModifiedBy = MusterContainer.AppUser.ID
                End If
                success = pLicen.Save(UIUtilsGen.ModuleID.Company, MusterContainer.AppUser.UserKey, returnVal)
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
                nLicenseeID = pLicen.ID
                'txtCertificationType.Text = pLicen.CERT_TYPE_ID
                'txtStatus.Text = pLicen.STATUS_ID
                If Not txtCompany.Text = "" Then
                    If nAssociationID > 0 Or nAssociationID <= -100 Then
                        ncompAddressID = pCompanyLicenseeAssociation.ComLicCollection.Item(nAssociationID).ComLicAddressID
                    End If
                    For Each oCompanyLicenseeInfoLocal In pCompanyLicenseeAssociation.ComLicCollection.Values
                        If oCompanyLicenseeInfoLocal.LicenseeID = pLicen.ID And oCompanyLicenseeInfoLocal.CompanyID = nCompanyID Then
                            oCompanyLicenseeInfoLocal.ComLicAddressID = ncompAddressID
                            oCompanyLicenseeInfoLocal.Deleted = False
                            bolExist = True
                            Exit For
                        End If
                    Next
                    If Not bolExist Then
                        CompanyLicenseeInfo = New MUSTER.Info.CompanyLicenseeInfo(nAssociationID, _
                                                                  nCompanyID, _
                                                                  pLicen.ID, _
                                                                  ncompAddressID, _
                                                                  False, _
                                                                  IIf((nAssociationID <= 0 And nAssociationID > -100), MusterContainer.AppUser.ID, ""), _
                                                                    Now, _
                                                                  IIf(nAssociationID > 0, MusterContainer.AppUser.ID, ""), _
                                                                    dtNull)
                        CompanyLicenseeInfo.IsDirty = True
                        pCompanyLicenseeAssociation.Add(CompanyLicenseeInfo)
                    End If
                End If
                oLicAdd.LicenseeID = pLicen.ID

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
                FillLicenseeCourseGrid()
                FillCourseTestGrid()
                PopulateLicenseNumber()
                populateLicenseeInfo(licenseeInfo)
                UCLicenseeDocuments.LoadDocumentsGrid(pLicen.ID, 0, 893)
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
                'If bolLicenseeCard Then
                '    GenerateLicenseeCard(licenseeInfo)
                '    bolLicenseeCard = False
                '    UIUtilsGen.Delay(, 1)
                'End If
                'If bolLicenseeCertificate Then
                '    GenerateLicenseeCertificateLetter(licenseeInfo)
                '    bolLicenseeCertificate = False
                '    UIUtilsGen.Delay(, 1)
                'End If

                MsgBox("Licensee is saved successfully")
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
                errStr += "Licensee's First and Last Name" + vbCrLf
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
            If Date.Compare(pLicen.EXCEPT_GRANT_DATE, dtNull) <> 0 Then
                If Date.Compare(pLicen.EXCEPT_GRANT_DATE, pLicen.EXTENSION_DEADLINE_DATE) > 0 Then
                    errStr += "Extension Deadline Date must be greater than Date Licensee Requested Extension" + vbCrLf
                End If
            End If
            If txtLicenseeAddress.Text = "" Then
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
    Private Function validateLicenseeTests(ByRef ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow) As Boolean
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
            If pLicen.pLicenseeCourseTest.colLicCourseTest.Count = 0 Then
                Return True
            End If
            For Each oTestInfo In pLicen.pLicenseeCourseTest.colLicCourseTest.Values
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
    Private Sub lblLicenseeInfoDisplay_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblLicenseeInfoDisplay.Click
        lblLicenseeInfo_Click(sender, e)
    End Sub
    Private Sub lblLicenseeInfo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblLicenseeInfo.Click
        ExpandCollapse(pnlLicenseeInfo, lblLicenseeInfoDisplay)
    End Sub
    Private Sub lblCoursesDisplay_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblCoursesDisplay.Click
        lblCourses_Click(sender, e)
    End Sub
    Private Sub lblCourses_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lblCourses.Click
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
            If pLicen.ID >= -100 And pLicen.ID <= 0 Then
                MsgBox("Please Save Licensee before entering comments")
                Exit Sub
            End If
            nEntityType = UIUtilsGen.EntityTypes.Licensee
            strEntityName = "Licensee : " + CStr(pLicen.ID) + " " + pLicen.FIRST_NAME
            If Not resetBtnColor Then
                SC = New ShowComments(pLicen.ID, nEntityType, IIf(bolSetCounts, "", "Licensee"), strEntityName, pLicen.Comments, Me.Text, , False)
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
            SF = New ShowFlags(pLicen.ID, UIUtilsGen.EntityTypes.Licensee, "LICENSEE")
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
    '    slLetterCount.Add(LicenseeLetters.CongLetter, 0)
    '    slLetterCount.Add(LicenseeLetters.LicenseeCard, 0)
    '    slLetterCount.Add(LicenseeLetters.LicenseeCertificateLetter, 0)
    '    slLetterCount.Add(LicenseeLetters.NoCertificationLetter, 0)
    '    slLetterCount.Add(LicenseeLetters.RenewalLetter, 0)
    'End Sub
    Public Sub GenerateCongLetter()
        Dim colParams As New Specialized.NameValueCollection
        Try
            'bolCongratsLetter = True
            colParams = FillParameters()
            oLetter.GenerateLicenseeLetter(nLicenseeID, "Licensee Congratulation Letter", "Licensee_Congratulation_Letter", "Licensee Congratulation Letter", "CongratLetter.doc", colParams)
            'slLetterCount.Item(LicenseeLetters.CongLetter) = 1
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub GenerateLicenseeCard(ByVal oLicensee As MUSTER.Info.LicenseeInfo)
        Dim colParams As New Specialized.NameValueCollection
        Try
            colParams = FillParameters()
            If oLicensee.CertTypeDesc.ToUpper = "INSTALL" Then
                colParams.Add("<TYPE>", "Install, Alter, and Permanently Close")
            ElseIf oLicensee.CertTypeDesc.ToUpper = "CLOSURE" Then
                colParams.Add("<TYPE>", "Permanently Close")
            End If
            colParams.Add("<LicenseeID>", pLicen.LICENSEE_NUMBER_PREFIX + pLicen.LICENSEE_NUMBER.ToString)
            oLetter.GenerateLicenseeLetter(nLicenseeID, "Licensee Card", "Licensee_Card", "Licensee Card", "LicenseeCard.doc", colParams, strPhotoFilePath, , , strSignatureFilePath)
            'slLetterCount.Item(LicenseeLetters.LicenseeCard) = 1
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub GenerateLicenseeCertificateLetter(ByVal oLicensee As MUSTER.Info.LicenseeInfo)
        Dim colParams As New Specialized.NameValueCollection
        Try
            colParams = FillParameters()
            colParams.Add("<Certification>", oLicensee.CertTypeDesc)
            If oLicensee.CertTypeDesc.ToUpper = "INSTALL" Then
                colParams.Add("<TYPE>", "Install, Alter, and Permanently Close")
            ElseIf oLicensee.CertTypeDesc.ToUpper = "CLOSURE" Then
                colParams.Add("<TYPE>", "Permanently Close")
            End If

            If pLicen.LICENSEE_NUMBER_PREFIX = "NRX" Or pLicen.LICENSEE_NUMBER_PREFIX = "CRX" Then
                colParams.Add("<Condition>", "Only USTs owned by " + txtCompany.Text)
            ElseIf pLicen.LICENSEE_NUMBER_PREFIX = "NHB" Or pLicen.LICENSEE_NUMBER_PREFIX = "CHB" Then
                colParams.Add("<Condition>", "Only as an employee of  " + txtCompany.Text)
            ElseIf pLicen.LICENSEE_NUMBER_PREFIX = "NHX" Or pLicen.LICENSEE_NUMBER_PREFIX = "CHX" Then
                colParams.Add("<Condition>", "None")
            End If
            'If pLicen.HIRE_STATUS = "HX - For Hire - Owner" Then
            '    colParams.Add("<Condition>", "Only USTs owned by " + txtCompany.Text)
            'ElseIf pLicen.HIRE_STATUS = "HB - For Hire - Employee" Then
            '    colParams.Add("<Condition>", "Only as an employee of  " + txtCompany.Text)
            'End If
            colParams.Add("<LicenseeID>", pLicen.LICENSEE_NUMBER_PREFIX + pLicen.LICENSEE_NUMBER.ToString)
            oLetter.GenerateLicenseeLetter(nLicenseeID, "Licensee Certification Letter", "Licensee_Certification_Letter", "Licensee Certification Letter", "CertificationLetter.doc", colParams)
            'slLetterCount.Item(LicenseeLetters.LicenseeCertificateLetter) = 1
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    'Private Sub pLicen_GenerateCongLetter() Handles pLicen.GenerateCongLetter
    '    If slLetterCount Is Nothing Then
    '        ' initialize sorted list to maintain the letter count to prevent creating multiple letters
    '        InitializeLetterCount()
    '    End If
    '    If CType(slLetterCount.Item(LicenseeLetters.CongLetter), Integer) > 0 Then
    '        Exit Sub
    '    End If
    '    bolCongratsLetter = True
    'End Sub
    'Private Sub pLicen_GenerateLicenseeCard() Handles pLicen.GenerateLicenseeCard
    '    If slLetterCount Is Nothing Then
    '        ' initialize sorted list to maintain the letter count to prevent creating multiple letters
    '        InitializeLetterCount()
    '    End If
    '    If CType(slLetterCount.Item(LicenseeLetters.LicenseeCard), Integer) > 0 Then
    '        Exit Sub
    '    End If
    '    bolLicenseeCard = True
    'End Sub
    'Private Sub pLicen_GenerateLicenseeCertLetter() Handles pLicen.GenerateLicenseeCertLetter
    '    If slLetterCount Is Nothing Then
    '        ' initialize sorted list to maintain the letter count to prevent creating multiple letters
    '        InitializeLetterCount()
    '    End If
    '    If CType(slLetterCount.Item(LicenseeLetters.LicenseeCertificateLetter), Integer) > 0 Then
    '        Exit Sub
    '    End If
    '    bolLicenseeCertificate = True
    'End Sub
    'Private Sub pLicen_GenerateNoCertificationLetter(ByVal oLicensee As MUSTER.Info.LicenseeInfo) Handles pLicen.GenerateNoCertificationLetter
    '    Dim colParams As New Specialized.NameValueCollection
    '    Try
    '        colParams = FillParameters()
    '        If oLicensee.CertTypeDesc = "INSTALL" Then
    '            colParams.Add("<Certification Type> ", "Install, Alter, and Permanently Close")
    '        ElseIf oLicensee.CertTypeDesc = "CLOSURE" Then
    '            colParams.Add("<Certification Type>", "Permanently Close")
    '        Else
    '            colParams.Add("<Certification Type>", "None")
    '        End If
    '        oLetter.GenerateLicenseeLetter(nLicenseeID, "Licensee No Certification Letter", "Licensee_NoCertification_Letter", "Licensee No Certification Letter", "NoLongerCertifiedLetter.doc", colParams)
    '        slLetterCount.Item(LicenseeLetters.NoCertificationLetter) = 1
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub pLicen_GenerateLicenseeRenewalLetter() Handles pLicen.GenerateLicenseeRenewalLetter
    '    Dim colParams As New Specialized.NameValueCollection
    '    Try
    '        If slLetterCount Is Nothing Then
    '            ' initialize sorted list to maintain the letter count to prevent creating multiple letters
    '            InitializeLetterCount()
    '        End If
    '        If CType(slLetterCount.Item(LicenseeLetters.RenewalLetter), Integer) > 0 Then
    '            Exit Sub
    '        End If
    '        If bolRenewalLetter = True Then
    '            Exit Sub
    '        End If

    '        colParams = FillParameters()
    '        oLetter.GenerateLicenseeLetter(nLicenseeID, "Renewal Certificate Letter", "Renewal_Certificate", "Licensee Renewal Certificate for Company", "LicenseeRenewalLetter.doc", colParams)
    '        slLetterCount.Item(LicenseeLetters.RenewalLetter) = 1
    '        bolRenewalLetter = True

    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub pLicen_GenerateNoCertificationLetterOption(ByVal oLicensee As MUSTER.Info.LicenseeInfo) Handles pLicen.GenerateNoCertificationLetterOption
    '    Dim result As DialogResult
    '    Try
    '        If slLetterCount Is Nothing Then
    '            ' initialize sorted list to maintain the letter count to prevent creating multiple letters
    '            InitializeLetterCount()
    '        End If
    '        If CType(slLetterCount.Item(LicenseeLetters.NoCertificationLetter), Integer) > 0 Then
    '            Exit Sub
    '        End If
    '        result = MessageBox.Show("Do you want to generate the NOT CERTIFIED LETTER?", "Generate Not Certified Letter.", MessageBoxButtons.YesNo)
    '        If result = DialogResult.Yes Then
    '            pLicen_GenerateNoCertificationLetter(oLicensee)
    '        End If
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub

    Private Function FillParameters(Optional ByVal fillInfoNeededInfo As Boolean = False) As Specialized.NameValueCollection
        Try
            Dim colParams As New Specialized.NameValueCollection
            Dim expStr As String
            'Build NameValueCollection with Tags and Values.
            colParams.Add("<Date>", Format(Now, "MMMM dd, yyyy"))
            expStr = pLicen.LICENSE_EXPIRE_DATE
            If expStr.Length > 10 Then
                expStr = expStr.Substring(5, 2) + "/" + expStr.Substring(8, 2) + "/" + expStr.Substring(0, 4)
            End If
            colParams.Add("<Expiration Date>", expStr)
            colParams.Add("<Licensee Name>", pLicen.FullName)
            colParams.Add("<Company Name>", txtCompany.Text)

            If ncompAddressID > 0 Then
                oCompanyAdd.Retrieve(ncompAddressID, False)


                colParams.Add("<Company Address1>", oCompanyAdd.AddressLine1)

                If oCompanyAdd.AddressLine2 = String.Empty Then
                    colParams.Add("<Company Address2>", oCompanyAdd.City & ", " & oCompanyAdd.State.TrimEnd & " " & oCompanyAdd.Zip.Substring(0, 5))
                    colParams.Add("<City/State/Zip>", "")
                Else
                    colParams.Add("<Company Address2>", oCompanyAdd.AddressLine2)
                    colParams.Add("<City/State/Zip>", oCompanyAdd.City & ", " & oCompanyAdd.State.TrimEnd & " " & oCompanyAdd.Zip.Substring(0, 5))
                End If
                colParams.Add("<Licensee Greeting>", "Dear " + pLicen.FullName)
                colParams.Add("<User>", MusterContainer.AppUser.Name)
                ' colParams.Add("<User Phone>", MusterContainer.AppUser.PhoneNumber)

                If pLicen.CERT_TYPE_DESC.ToUpper = "CLOSURE" Then
                    colParams.Add("<Certification Type>", "Permanently Close ")
                ElseIf pLicen.CERT_TYPE_DESC.ToUpper = "INSTALL" Then
                    colParams.Add("<Certification Type> ", "Install, Alter, and Permanently Close ")
                End If

                If fillInfoNeededInfo Then
                    Dim strInfoNeeded As String = String.Empty
                    Dim count As Integer = 0

                    If pLicen.LICENSEE_NUMBER_PREFIX = "CHB" Or _
                        pLicen.LICENSEE_NUMBER_PREFIX = "CRX" Or _
                        pLicen.LICENSEE_NUMBER_PREFIX = "CHX" Or _
                        pLicen.LICENSEE_NUMBER_PREFIX = "NHB" Or _
                        pLicen.LICENSEE_NUMBER_PREFIX = "NRX" Or _
                        pLicen.LICENSEE_NUMBER_PREFIX = "NHX" Then
                        If Date.Compare(pLicen.APP_RECVD_DATE, dtNull) <> 0 Then
                            If Date.Compare(pLicen.ISSUED_DATE, dtNull) <> 0 Then
                                If Date.Compare(pLicen.APP_RECVD_DATE, pLicen.ISSUED_DATE) < 0 Then
                                    count += 1
                                    strInfoNeeded += count.ToString + ". A completed certification renewal application" + vbCrLf
                                End If
                            End If
                        End If
                    End If

                    Dim bolClosure As Boolean = True
                    Dim bolInstall As Boolean = True

                    If Date.Compare(pLicen.ISSUED_DATE, dtNull) = 0 Then
                        bolClosure = False
                        bolInstall = False
                    Else
                        For Each oLicCourseInfo In pLicen.pLicenseeCourse.colLicCourse.Values
                            If oLicCourseInfo.LicenseeID = pLicen.ID Then
                                If Date.Compare(oLicCourseInfo.CourseDate, pLicen.ISSUED_DATE) > 0 Then
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
                        If pLicen.LICENSEE_NUMBER_PREFIX = "CHB" Or _
                            pLicen.LICENSEE_NUMBER_PREFIX = "CRX" Or _
                            pLicen.LICENSEE_NUMBER_PREFIX = "CHX" Or _
                            pLicen.LICENSEE_NUMBER_PREFIX = "NHB" Or _
                            pLicen.LICENSEE_NUMBER_PREFIX = "NRX" Or _
                            pLicen.LICENSEE_NUMBER_PREFIX = "NHX" Then
                            count += 1
                            strInfoNeeded += count.ToString + ". A Closure course completion certificate" + vbCrLf
                        End If
                    End If

                    If bolInstall Then
                        If pLicen.LICENSEE_NUMBER_PREFIX = "NHB" Or _
                            pLicen.LICENSEE_NUMBER_PREFIX = "NRX" Or _
                            pLicen.LICENSEE_NUMBER_PREFIX = "NHX" Then
                            count += 1
                            strInfoNeeded += count.ToString + ". An Install course completion certificate" + vbCrLf
                        End If
                    End If

                    If pLicen.LICENSEE_NUMBER_PREFIX = "CHB" Or _
                        pLicen.LICENSEE_NUMBER_PREFIX = "CRB" Or _
                        pLicen.LICENSEE_NUMBER_PREFIX = "NHB" Or _
                        pLicen.LICENSEE_NUMBER_PREFIX = "NRB" Then
                        If pLicen.HIRE_STATUS <> String.Empty Then
                            If pLicen.HIRE_STATUS = "HB - For Hire - Employee" Then
                                If Not pLicen.EMPLOYEE_LETTER Then
                                    count += 1
                                    strInfoNeeded += count.ToString + ". An Employee letter stating you are a full time employee of company" + vbCrLf
                                End If
                            End If
                        End If
                    End If

                    If pLicen.LICENSEE_NUMBER_PREFIX = "CHB" Or _
                        pLicen.LICENSEE_NUMBER_PREFIX = "CHX" Or _
                        pLicen.LICENSEE_NUMBER_PREFIX = "NHB" Or _
                        pLicen.LICENSEE_NUMBER_PREFIX = "NHX" Then
                        If Date.Compare(FinRespExpirationDate, dtNull) <> 0 Then
                            If Date.Compare(pLicen.LICENSE_EXPIRE_DATE, dtNull) <> 0 Then
                                If Date.Compare(FinRespExpirationDate, pLicen.LICENSE_EXPIRE_DATE) < 0 Then
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
    Public Sub setSaveLicensee(ByVal bolState As Boolean)
        If Not bolLoading Then
            btnSaveLicensee.Enabled = bolState Or pLicen.IsDirty
        End If

    End Sub

    Public Sub ValidationErrors(ByVal MsgStr As String) Handles pLicen.LicenseeErr
        MsgBox(MsgStr)
    End Sub
    Private Sub pLicen_LicenseeChanged(ByVal bolValue As Boolean) Handles pLicen.LicenseeChanged
        setSaveLicensee(bolValue)
    End Sub
    Private Sub pLicen_LicenseeCourseChanged(ByVal bolValue As Boolean) Handles pLicen.LicenseeCourseChanged
        setSaveLicensee(bolValue)
    End Sub
    Private Sub pLicen_LicenseeCourseTestsChanged(ByVal bolValue As Boolean) Handles pLicen.LicenseeCourseTestsChanged
        setSaveLicensee(bolValue)
    End Sub
    Private Sub pLicen_ColChanged(ByVal bolValue As Boolean) Handles pLicen.ColChanged
        setSaveLicensee(bolValue)
    End Sub

    Private Sub oLicAdd_evtAddressesChanged(ByVal bolValue As Boolean) Handles oLicAdd.evtAddressesChanged
        setSaveLicensee(bolValue)
    End Sub
    Private Sub oLicAdd_AddressChanged(ByVal bolValue As Boolean) Handles oLicAdd.evtAddressChanged
        setSaveLicensee(bolValue)
    End Sub
    Private Sub oLicAdd_evtAddressChanged(ByVal bolValue As Boolean) Handles oLicAdd.evtAddressChanged
        setSaveLicensee(bolValue)
    End Sub

    Private Sub oCompanyAdd_evtAddressesChanged(ByVal bolValue As Boolean) Handles oCompanyAdd.evtAddressesChanged
        setSaveLicensee(bolValue)
    End Sub
    Private Sub oCompanyAdd_evtAddressChanged(ByVal bolValue As Boolean) Handles oCompanyAdd.evtAddressChanged
        setSaveLicensee(bolValue)
    End Sub

    Private Sub pCompanyLicenseeAssociation_CompanyLicenseeChanged(ByVal bolValue As Boolean) Handles pCompanyLicenseeAssociation.CompanyLicenseeChanged
        setSaveLicensee(bolValue)
    End Sub
    Private Sub pCompanyLicenseeAssociation_ColChanged(ByVal bolValue As Boolean) Handles pCompanyLicenseeAssociation.ColChanged
        setSaveLicensee(bolValue)
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
            If (oLicAdd.AddressId > 0 Or oLicAdd.AddressId < -100) And (pLicen.ID > 0 Or pLicen.ID < -100) Then
                UIUtilsGen.CreateEnvelopes(strName, arrAddress, "COM", pLicen.ID)
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
            If (oLicAdd.AddressId > 0 Or oLicAdd.AddressId < -100) And (pLicen.ID > 0 Or pLicen.ID < -100) Then
                UIUtilsGen.CreateLabels(strName, arrAddress, "COM", pLicen.ID)
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
            If (oComAddInfo.AddressId > 0 Or oComAddInfo.AddressId < -100) And (pLicen.ID > 0 Or pLicen.ID < -100) Then
                UIUtilsGen.CreateEnvelopes(strName, arrAddress, "COM", pLicen.ID)
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
            If (oComAddInfo.AddressId > 0 Or oComAddInfo.AddressId < -100) And (pLicen.ID > 0 Or pLicen.ID < -100) Then
                UIUtilsGen.CreateLabels(strName, arrAddress, "COM", pLicen.ID)
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
