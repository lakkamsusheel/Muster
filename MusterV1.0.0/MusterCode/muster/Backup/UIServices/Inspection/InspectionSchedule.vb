Public Class Inspection
    Inherits System.Windows.Forms.Form
#Region "User Defined Variables"
    Private frmChecklist As CheckList
    Private frmInspecHistory As InspectionHistory
    Private frmAssignedInspec As AssignedInspection
    Private WithEvents frmReschedule As RescheduleInspection
    Private frmCLProgress As CheckListProgress
    Private WithEvents oInspection As MUSTER.BusinessLogic.pInspection
    Private WithEvents ltrGen As MUSTER.BusinessLogic.pLetterGen
    Public MyGuid As New System.Guid
    Private bolLoading As Boolean = False
    Private dsTargetFacs As New DataSet
    Private dsTargetFacsForOwner As New DataSet
    Private dsAssignedInspections As New DataSet
    Private MouseDowned As Infragistics.Win.UltraWinGrid.UltraGridRow
    Private arrugMyFacs(1) As Infragistics.Win.UltraWinGrid.UltraGrid
    Private arrugAllFacsforOwner(1) As Infragistics.Win.UltraWinGrid.UltraGrid
    Private arrugAssignedInspections(1) As Infragistics.Win.UltraWinGrid.UltraGrid
    Private nInspectorID As Integer = MusterContainer.AppUser.UserKey
    Private nOwnerID, nFacilityID As Integer
    Private nSelectedFacilityID, nSelectedOwnerID As Integer
    Dim returnVal As String = String.Empty
    Dim commentDataSet As DataSet
    Private ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow

#End Region
#Region "Windows Form Designer generated code "

    Public Sub New(ByRef pInspec As MUSTER.BusinessLogic.pInspection, Optional ByVal nOwnID As Integer = 0, Optional ByVal nFacID As Integer = 0)
        MyBase.New()

        MyGuid = System.Guid.NewGuid
        bolLoading = True

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Cursor.Current = Cursors.AppStarting
        oInspection = pInspec
        arrugMyFacs(0) = ugMyFacs
        arrugMyFacs(1) = ugMyFacilities
        arrugAllFacsforOwner(0) = ugAllFacsForOwner
        arrugAllFacsforOwner(1) = ugAllFacForOwner
        arrugAssignedInspections(0) = ugAssignedInspections
        arrugAssignedInspections(1) = ugAssignedInspec
        nOwnerID = nOwnID
        nFacilityID = nFacID
        nSelectedFacilityID = nFacilityID
        nSelectedOwnerID = nOwnerID
        MusterContainer.AppUser.LogEntry("Inspection", MyGuid.ToString)
        'MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "WindowName", Me.Text, "Inspection")
        MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "Inspection")
        bolLoading = False
        Cursor.Current = Cursors.Default

        Dim pcom As New BusinessLogic.pComments
        commentDataSet = pcom.GetComments(, 6, 0)
        pcom = Nothing
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        MusterContainer.AppSemaphores.Remove(MyGuid.ToString)
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
    Friend WithEvents pnlInspectionDetails As System.Windows.Forms.Panel
    Friend WithEvents tbCtrlInspection As System.Windows.Forms.TabControl
    Friend WithEvents tbPageAllInfo As System.Windows.Forms.TabPage
    Friend WithEvents pnlAllInfoDetails As System.Windows.Forms.Panel
    Friend WithEvents btnPrintInspectionInfo As System.Windows.Forms.Button
    Friend WithEvents btnEnterInspectionInfo As System.Windows.Forms.Button
    Friend WithEvents ugAssignedInspections As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents lblAssignedInspections As System.Windows.Forms.Label
    Friend WithEvents btnViewHistory As System.Windows.Forms.Button
    Friend WithEvents btnEnterChecklistInfo As System.Windows.Forms.Button
    Friend WithEvents btnPrintChecklists As System.Windows.Forms.Button
    Friend WithEvents btnGenerateLetters As System.Windows.Forms.Button
    Friend WithEvents ugAllFacsForOwner As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents lblAllFacsForOwner As System.Windows.Forms.Label
    Friend WithEvents btnRestoreSortOrder As System.Windows.Forms.Button
    Friend WithEvents ugMyFacs As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents lblMyFacs As System.Windows.Forms.Label
    Friend WithEvents tbPageAssignedInspec As System.Windows.Forms.TabPage
    Friend WithEvents ugAssignedInspec As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents lblAssignedInspec As System.Windows.Forms.Label
    Friend WithEvents pnlAssignedInspecBottom As System.Windows.Forms.Panel
    Friend WithEvents btnAssignedInspecPrint As System.Windows.Forms.Button
    Friend WithEvents btnAssignedInspecEnterInspec As System.Windows.Forms.Button
    Friend WithEvents tbPageMyFacilities As System.Windows.Forms.TabPage
    Friend WithEvents ugMyFacilities As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents lblMyFacilities As System.Windows.Forms.Label
    Friend WithEvents pnlMyFacilitiesBottom As System.Windows.Forms.Panel
    Friend WithEvents btnMyFacViewHistory As System.Windows.Forms.Button
    Friend WithEvents btnMyFacEnterCheckInfo As System.Windows.Forms.Button
    Friend WithEvents btnMyFacPrintChecklists As System.Windows.Forms.Button
    Friend WithEvents btnMyFacGenLetters As System.Windows.Forms.Button
    Friend WithEvents tbPageAllFacForOwner As System.Windows.Forms.TabPage
    Friend WithEvents ugAllFacForOwner As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents lblAllFacforOwner As System.Windows.Forms.Label
    Friend WithEvents pnlAllFacForOwnerBottom As System.Windows.Forms.Panel
    Friend WithEvents btnAllFacViewHistory As System.Windows.Forms.Button
    Friend WithEvents btnAllFacEntercheckInfo As System.Windows.Forms.Button
    Friend WithEvents btnAllFacPrintChecklist As System.Windows.Forms.Button
    Friend WithEvents btnAllFacGenLetters As System.Windows.Forms.Button
    Friend WithEvents pnlInspectionHeader As System.Windows.Forms.Panel
    Friend WithEvents chkShowAllTargetFacilities As System.Windows.Forms.CheckBox
    Friend WithEvents grpBxSortBy As System.Windows.Forms.GroupBox
    Friend WithEvents rdBtnScheduled As System.Windows.Forms.RadioButton
    Friend WithEvents rdBtnDue As System.Windows.Forms.RadioButton
    Friend WithEvents lblFacilitiesFor As System.Windows.Forms.Label
    Friend WithEvents btnReschedule As System.Windows.Forms.Button
    Friend WithEvents btnRestoreSortOrder2 As System.Windows.Forms.Button
    Friend WithEvents pnlAllFacsForOwnerTop As System.Windows.Forms.Panel
    Friend WithEvents pnlAllFacForOwnerDetails As System.Windows.Forms.Panel
    Friend WithEvents pnlAllInfoAllFacsTop As System.Windows.Forms.Panel
    Friend WithEvents pnlAllInfoMyFacsDetails As System.Windows.Forms.Panel
    Friend WithEvents pnlAllInfoMyFacsTop As System.Windows.Forms.Panel
    Friend WithEvents pnlMyFacilitiesTop As System.Windows.Forms.Panel
    Friend WithEvents pnlAssignedInspectop As System.Windows.Forms.Panel
    Friend WithEvents pnlAssignedInspecDetails As System.Windows.Forms.Panel
    Friend WithEvents pnlMyFacilitiesDetails As System.Windows.Forms.Panel
    Friend WithEvents pnlAllInfoAllFacsDetails As System.Windows.Forms.Panel
    Friend WithEvents pnlAllInfoAllFacsBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlAllInfoAssignedInspecTop As System.Windows.Forms.Panel
    Friend WithEvents pnlAllInfoAssignedInspecDetails As System.Windows.Forms.Panel
    Friend WithEvents pnlAllInfoAssignedInspecbottom As System.Windows.Forms.Panel
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents btnMyFacSchedule As System.Windows.Forms.Button
    Friend WithEvents btnAllFacSchedule As System.Windows.Forms.Button
    Friend WithEvents cmbInspectors As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlInspectionDetails = New System.Windows.Forms.Panel
        Me.tbCtrlInspection = New System.Windows.Forms.TabControl
        Me.tbPageAllInfo = New System.Windows.Forms.TabPage
        Me.pnlAllInfoDetails = New System.Windows.Forms.Panel
        Me.pnlAllInfoAssignedInspecbottom = New System.Windows.Forms.Panel
        Me.btnPrintInspectionInfo = New System.Windows.Forms.Button
        Me.btnEnterInspectionInfo = New System.Windows.Forms.Button
        Me.pnlAllInfoAssignedInspecDetails = New System.Windows.Forms.Panel
        Me.ugAssignedInspections = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlAllInfoAssignedInspecTop = New System.Windows.Forms.Panel
        Me.lblAssignedInspections = New System.Windows.Forms.Label
        Me.pnlAllInfoAllFacsBottom = New System.Windows.Forms.Panel
        Me.btnReschedule = New System.Windows.Forms.Button
        Me.btnViewHistory = New System.Windows.Forms.Button
        Me.btnEnterChecklistInfo = New System.Windows.Forms.Button
        Me.btnPrintChecklists = New System.Windows.Forms.Button
        Me.btnGenerateLetters = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.pnlAllInfoAllFacsDetails = New System.Windows.Forms.Panel
        Me.ugAllFacsForOwner = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlAllInfoAllFacsTop = New System.Windows.Forms.Panel
        Me.lblAllFacsForOwner = New System.Windows.Forms.Label
        Me.btnRestoreSortOrder = New System.Windows.Forms.Button
        Me.pnlAllInfoMyFacsDetails = New System.Windows.Forms.Panel
        Me.ugMyFacs = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlAllInfoMyFacsTop = New System.Windows.Forms.Panel
        Me.lblMyFacs = New System.Windows.Forms.Label
        Me.tbPageMyFacilities = New System.Windows.Forms.TabPage
        Me.pnlMyFacilitiesDetails = New System.Windows.Forms.Panel
        Me.ugMyFacilities = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlMyFacilitiesTop = New System.Windows.Forms.Panel
        Me.lblMyFacilities = New System.Windows.Forms.Label
        Me.pnlMyFacilitiesBottom = New System.Windows.Forms.Panel
        Me.btnMyFacSchedule = New System.Windows.Forms.Button
        Me.btnMyFacViewHistory = New System.Windows.Forms.Button
        Me.btnMyFacEnterCheckInfo = New System.Windows.Forms.Button
        Me.btnMyFacPrintChecklists = New System.Windows.Forms.Button
        Me.btnMyFacGenLetters = New System.Windows.Forms.Button
        Me.tbPageAllFacForOwner = New System.Windows.Forms.TabPage
        Me.pnlAllFacForOwnerDetails = New System.Windows.Forms.Panel
        Me.ugAllFacForOwner = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlAllFacsForOwnerTop = New System.Windows.Forms.Panel
        Me.lblAllFacforOwner = New System.Windows.Forms.Label
        Me.btnRestoreSortOrder2 = New System.Windows.Forms.Button
        Me.pnlAllFacForOwnerBottom = New System.Windows.Forms.Panel
        Me.btnAllFacSchedule = New System.Windows.Forms.Button
        Me.btnAllFacViewHistory = New System.Windows.Forms.Button
        Me.btnAllFacEntercheckInfo = New System.Windows.Forms.Button
        Me.btnAllFacPrintChecklist = New System.Windows.Forms.Button
        Me.btnAllFacGenLetters = New System.Windows.Forms.Button
        Me.tbPageAssignedInspec = New System.Windows.Forms.TabPage
        Me.pnlAssignedInspecDetails = New System.Windows.Forms.Panel
        Me.ugAssignedInspec = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlAssignedInspectop = New System.Windows.Forms.Panel
        Me.lblAssignedInspec = New System.Windows.Forms.Label
        Me.pnlAssignedInspecBottom = New System.Windows.Forms.Panel
        Me.btnAssignedInspecPrint = New System.Windows.Forms.Button
        Me.btnAssignedInspecEnterInspec = New System.Windows.Forms.Button
        Me.pnlInspectionHeader = New System.Windows.Forms.Panel
        Me.cmbInspectors = New System.Windows.Forms.ComboBox
        Me.chkShowAllTargetFacilities = New System.Windows.Forms.CheckBox
        Me.grpBxSortBy = New System.Windows.Forms.GroupBox
        Me.rdBtnScheduled = New System.Windows.Forms.RadioButton
        Me.rdBtnDue = New System.Windows.Forms.RadioButton
        Me.lblFacilitiesFor = New System.Windows.Forms.Label
        Me.pnlInspectionDetails.SuspendLayout()
        Me.tbCtrlInspection.SuspendLayout()
        Me.tbPageAllInfo.SuspendLayout()
        Me.pnlAllInfoDetails.SuspendLayout()
        Me.pnlAllInfoAssignedInspecbottom.SuspendLayout()
        Me.pnlAllInfoAssignedInspecDetails.SuspendLayout()
        CType(Me.ugAssignedInspections, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlAllInfoAssignedInspecTop.SuspendLayout()
        Me.pnlAllInfoAllFacsBottom.SuspendLayout()
        Me.pnlAllInfoAllFacsDetails.SuspendLayout()
        CType(Me.ugAllFacsForOwner, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlAllInfoAllFacsTop.SuspendLayout()
        Me.pnlAllInfoMyFacsDetails.SuspendLayout()
        CType(Me.ugMyFacs, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlAllInfoMyFacsTop.SuspendLayout()
        Me.tbPageMyFacilities.SuspendLayout()
        Me.pnlMyFacilitiesDetails.SuspendLayout()
        CType(Me.ugMyFacilities, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlMyFacilitiesTop.SuspendLayout()
        Me.pnlMyFacilitiesBottom.SuspendLayout()
        Me.tbPageAllFacForOwner.SuspendLayout()
        Me.pnlAllFacForOwnerDetails.SuspendLayout()
        CType(Me.ugAllFacForOwner, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlAllFacsForOwnerTop.SuspendLayout()
        Me.pnlAllFacForOwnerBottom.SuspendLayout()
        Me.tbPageAssignedInspec.SuspendLayout()
        Me.pnlAssignedInspecDetails.SuspendLayout()
        CType(Me.ugAssignedInspec, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlAssignedInspectop.SuspendLayout()
        Me.pnlAssignedInspecBottom.SuspendLayout()
        Me.pnlInspectionHeader.SuspendLayout()
        Me.grpBxSortBy.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlInspectionDetails
        '
        Me.pnlInspectionDetails.Controls.Add(Me.tbCtrlInspection)
        Me.pnlInspectionDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlInspectionDetails.Location = New System.Drawing.Point(0, 53)
        Me.pnlInspectionDetails.Name = "pnlInspectionDetails"
        Me.pnlInspectionDetails.Size = New System.Drawing.Size(880, 641)
        Me.pnlInspectionDetails.TabIndex = 5
        '
        'tbCtrlInspection
        '
        Me.tbCtrlInspection.Controls.Add(Me.tbPageAllInfo)
        Me.tbCtrlInspection.Controls.Add(Me.tbPageMyFacilities)
        Me.tbCtrlInspection.Controls.Add(Me.tbPageAllFacForOwner)
        Me.tbCtrlInspection.Controls.Add(Me.tbPageAssignedInspec)
        Me.tbCtrlInspection.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbCtrlInspection.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbCtrlInspection.ItemSize = New System.Drawing.Size(64, 18)
        Me.tbCtrlInspection.Location = New System.Drawing.Point(0, 0)
        Me.tbCtrlInspection.Multiline = True
        Me.tbCtrlInspection.Name = "tbCtrlInspection"
        Me.tbCtrlInspection.SelectedIndex = 0
        Me.tbCtrlInspection.ShowToolTips = True
        Me.tbCtrlInspection.Size = New System.Drawing.Size(880, 641)
        Me.tbCtrlInspection.TabIndex = 0
        '
        'tbPageAllInfo
        '
        Me.tbPageAllInfo.Controls.Add(Me.pnlAllInfoDetails)
        Me.tbPageAllInfo.Location = New System.Drawing.Point(4, 22)
        Me.tbPageAllInfo.Name = "tbPageAllInfo"
        Me.tbPageAllInfo.Size = New System.Drawing.Size(872, 615)
        Me.tbPageAllInfo.TabIndex = 0
        Me.tbPageAllInfo.Text = "All Info"
        '
        'pnlAllInfoDetails
        '
        Me.pnlAllInfoDetails.Controls.Add(Me.pnlAllInfoAssignedInspecbottom)
        Me.pnlAllInfoDetails.Controls.Add(Me.pnlAllInfoAssignedInspecDetails)
        Me.pnlAllInfoDetails.Controls.Add(Me.pnlAllInfoAssignedInspecTop)
        Me.pnlAllInfoDetails.Controls.Add(Me.pnlAllInfoAllFacsBottom)
        Me.pnlAllInfoDetails.Controls.Add(Me.pnlAllInfoAllFacsDetails)
        Me.pnlAllInfoDetails.Controls.Add(Me.pnlAllInfoAllFacsTop)
        Me.pnlAllInfoDetails.Controls.Add(Me.pnlAllInfoMyFacsDetails)
        Me.pnlAllInfoDetails.Controls.Add(Me.pnlAllInfoMyFacsTop)
        Me.pnlAllInfoDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlAllInfoDetails.Location = New System.Drawing.Point(0, 0)
        Me.pnlAllInfoDetails.Name = "pnlAllInfoDetails"
        Me.pnlAllInfoDetails.Size = New System.Drawing.Size(872, 615)
        Me.pnlAllInfoDetails.TabIndex = 1
        '
        'pnlAllInfoAssignedInspecbottom
        '
        Me.pnlAllInfoAssignedInspecbottom.Controls.Add(Me.btnPrintInspectionInfo)
        Me.pnlAllInfoAssignedInspecbottom.Controls.Add(Me.btnEnterInspectionInfo)
        Me.pnlAllInfoAssignedInspecbottom.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlAllInfoAssignedInspecbottom.Location = New System.Drawing.Point(0, 440)
        Me.pnlAllInfoAssignedInspecbottom.Name = "pnlAllInfoAssignedInspecbottom"
        Me.pnlAllInfoAssignedInspecbottom.Size = New System.Drawing.Size(872, 23)
        Me.pnlAllInfoAssignedInspecbottom.TabIndex = 15
        '
        'btnPrintInspectionInfo
        '
        Me.btnPrintInspectionInfo.Location = New System.Drawing.Point(376, 0)
        Me.btnPrintInspectionInfo.Name = "btnPrintInspectionInfo"
        Me.btnPrintInspectionInfo.Size = New System.Drawing.Size(128, 23)
        Me.btnPrintInspectionInfo.TabIndex = 14
        Me.btnPrintInspectionInfo.Text = "Print Inspection Info"
        '
        'btnEnterInspectionInfo
        '
        Me.btnEnterInspectionInfo.Location = New System.Drawing.Point(240, 0)
        Me.btnEnterInspectionInfo.Name = "btnEnterInspectionInfo"
        Me.btnEnterInspectionInfo.Size = New System.Drawing.Size(128, 23)
        Me.btnEnterInspectionInfo.TabIndex = 13
        Me.btnEnterInspectionInfo.Text = "Enter Inspection Info"
        '
        'pnlAllInfoAssignedInspecDetails
        '
        Me.pnlAllInfoAssignedInspecDetails.Controls.Add(Me.ugAssignedInspections)
        Me.pnlAllInfoAssignedInspecDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlAllInfoAssignedInspecDetails.Location = New System.Drawing.Point(0, 338)
        Me.pnlAllInfoAssignedInspecDetails.Name = "pnlAllInfoAssignedInspecDetails"
        Me.pnlAllInfoAssignedInspecDetails.Size = New System.Drawing.Size(872, 102)
        Me.pnlAllInfoAssignedInspecDetails.TabIndex = 14
        '
        'ugAssignedInspections
        '
        Me.ugAssignedInspections.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugAssignedInspections.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugAssignedInspections.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugAssignedInspections.Location = New System.Drawing.Point(0, 0)
        Me.ugAssignedInspections.Name = "ugAssignedInspections"
        Me.ugAssignedInspections.Size = New System.Drawing.Size(872, 102)
        Me.ugAssignedInspections.TabIndex = 11
        '
        'pnlAllInfoAssignedInspecTop
        '
        Me.pnlAllInfoAssignedInspecTop.Controls.Add(Me.lblAssignedInspections)
        Me.pnlAllInfoAssignedInspecTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlAllInfoAssignedInspecTop.Location = New System.Drawing.Point(0, 319)
        Me.pnlAllInfoAssignedInspecTop.Name = "pnlAllInfoAssignedInspecTop"
        Me.pnlAllInfoAssignedInspecTop.Size = New System.Drawing.Size(872, 19)
        Me.pnlAllInfoAssignedInspecTop.TabIndex = 13
        '
        'lblAssignedInspections
        '
        Me.lblAssignedInspections.Location = New System.Drawing.Point(2, 2)
        Me.lblAssignedInspections.Name = "lblAssignedInspections"
        Me.lblAssignedInspections.Size = New System.Drawing.Size(136, 18)
        Me.lblAssignedInspections.TabIndex = 0
        Me.lblAssignedInspections.Text = "Assigned Inspections"
        '
        'pnlAllInfoAllFacsBottom
        '
        Me.pnlAllInfoAllFacsBottom.Controls.Add(Me.btnReschedule)
        Me.pnlAllInfoAllFacsBottom.Controls.Add(Me.btnViewHistory)
        Me.pnlAllInfoAllFacsBottom.Controls.Add(Me.btnEnterChecklistInfo)
        Me.pnlAllInfoAllFacsBottom.Controls.Add(Me.btnPrintChecklists)
        Me.pnlAllInfoAllFacsBottom.Controls.Add(Me.btnGenerateLetters)
        Me.pnlAllInfoAllFacsBottom.Controls.Add(Me.Button1)
        Me.pnlAllInfoAllFacsBottom.Controls.Add(Me.Button2)
        Me.pnlAllInfoAllFacsBottom.Controls.Add(Me.Button3)
        Me.pnlAllInfoAllFacsBottom.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlAllInfoAllFacsBottom.Location = New System.Drawing.Point(0, 296)
        Me.pnlAllInfoAllFacsBottom.Name = "pnlAllInfoAllFacsBottom"
        Me.pnlAllInfoAllFacsBottom.Size = New System.Drawing.Size(872, 23)
        Me.pnlAllInfoAllFacsBottom.TabIndex = 12
        '
        'btnReschedule
        '
        Me.btnReschedule.Location = New System.Drawing.Point(584, 0)
        Me.btnReschedule.Name = "btnReschedule"
        Me.btnReschedule.Size = New System.Drawing.Size(80, 23)
        Me.btnReschedule.TabIndex = 10
        Me.btnReschedule.Text = "Schedule"
        '
        'btnViewHistory
        '
        Me.btnViewHistory.Location = New System.Drawing.Point(488, 0)
        Me.btnViewHistory.Name = "btnViewHistory"
        Me.btnViewHistory.Size = New System.Drawing.Size(88, 23)
        Me.btnViewHistory.TabIndex = 9
        Me.btnViewHistory.Text = "View History"
        '
        'btnEnterChecklistInfo
        '
        Me.btnEnterChecklistInfo.Location = New System.Drawing.Point(352, 1)
        Me.btnEnterChecklistInfo.Name = "btnEnterChecklistInfo"
        Me.btnEnterChecklistInfo.Size = New System.Drawing.Size(128, 23)
        Me.btnEnterChecklistInfo.TabIndex = 8
        Me.btnEnterChecklistInfo.Text = "Enter Checklist Info"
        '
        'btnPrintChecklists
        '
        Me.btnPrintChecklists.Location = New System.Drawing.Point(232, 1)
        Me.btnPrintChecklists.Name = "btnPrintChecklists"
        Me.btnPrintChecklists.Size = New System.Drawing.Size(112, 23)
        Me.btnPrintChecklists.TabIndex = 7
        Me.btnPrintChecklists.Text = "Print Checklist(s)"
        '
        'btnGenerateLetters
        '
        Me.btnGenerateLetters.Location = New System.Drawing.Point(112, 1)
        Me.btnGenerateLetters.Name = "btnGenerateLetters"
        Me.btnGenerateLetters.Size = New System.Drawing.Size(113, 23)
        Me.btnGenerateLetters.TabIndex = 6
        Me.btnGenerateLetters.Text = "Generate Letters"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(352, -1)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(128, 23)
        Me.Button1.TabIndex = 8
        Me.Button1.Text = "Enter Checklist Info"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(232, -1)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(112, 23)
        Me.Button2.TabIndex = 7
        Me.Button2.Text = "Print Checklist(s)"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(112, -1)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(113, 23)
        Me.Button3.TabIndex = 6
        Me.Button3.Text = "Generate Letters"
        '
        'pnlAllInfoAllFacsDetails
        '
        Me.pnlAllInfoAllFacsDetails.Controls.Add(Me.ugAllFacsForOwner)
        Me.pnlAllInfoAllFacsDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlAllInfoAllFacsDetails.Location = New System.Drawing.Point(0, 183)
        Me.pnlAllInfoAllFacsDetails.Name = "pnlAllInfoAllFacsDetails"
        Me.pnlAllInfoAllFacsDetails.Size = New System.Drawing.Size(872, 113)
        Me.pnlAllInfoAllFacsDetails.TabIndex = 11
        '
        'ugAllFacsForOwner
        '
        Me.ugAllFacsForOwner.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugAllFacsForOwner.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugAllFacsForOwner.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugAllFacsForOwner.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugAllFacsForOwner.Location = New System.Drawing.Point(0, 0)
        Me.ugAllFacsForOwner.Name = "ugAllFacsForOwner"
        Me.ugAllFacsForOwner.Size = New System.Drawing.Size(872, 113)
        Me.ugAllFacsForOwner.TabIndex = 4
        '
        'pnlAllInfoAllFacsTop
        '
        Me.pnlAllInfoAllFacsTop.Controls.Add(Me.lblAllFacsForOwner)
        Me.pnlAllInfoAllFacsTop.Controls.Add(Me.btnRestoreSortOrder)
        Me.pnlAllInfoAllFacsTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlAllInfoAllFacsTop.Location = New System.Drawing.Point(0, 160)
        Me.pnlAllInfoAllFacsTop.Name = "pnlAllInfoAllFacsTop"
        Me.pnlAllInfoAllFacsTop.Size = New System.Drawing.Size(872, 23)
        Me.pnlAllInfoAllFacsTop.TabIndex = 10
        '
        'lblAllFacsForOwner
        '
        Me.lblAllFacsForOwner.Location = New System.Drawing.Point(2, 3)
        Me.lblAllFacsForOwner.Name = "lblAllFacsForOwner"
        Me.lblAllFacsForOwner.Size = New System.Drawing.Size(184, 17)
        Me.lblAllFacsForOwner.TabIndex = 0
        Me.lblAllFacsForOwner.Text = "All Facilities for Selected Owner"
        '
        'btnRestoreSortOrder
        '
        Me.btnRestoreSortOrder.Enabled = False
        Me.btnRestoreSortOrder.Location = New System.Drawing.Point(346, 1)
        Me.btnRestoreSortOrder.Name = "btnRestoreSortOrder"
        Me.btnRestoreSortOrder.Size = New System.Drawing.Size(120, 23)
        Me.btnRestoreSortOrder.TabIndex = 2
        Me.btnRestoreSortOrder.Text = "Restore Sort Order"
        '
        'pnlAllInfoMyFacsDetails
        '
        Me.pnlAllInfoMyFacsDetails.Controls.Add(Me.ugMyFacs)
        Me.pnlAllInfoMyFacsDetails.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlAllInfoMyFacsDetails.Location = New System.Drawing.Point(0, 19)
        Me.pnlAllInfoMyFacsDetails.Name = "pnlAllInfoMyFacsDetails"
        Me.pnlAllInfoMyFacsDetails.Size = New System.Drawing.Size(872, 141)
        Me.pnlAllInfoMyFacsDetails.TabIndex = 3
        '
        'ugMyFacs
        '
        Me.ugMyFacs.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugMyFacs.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugMyFacs.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugMyFacs.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugMyFacs.Location = New System.Drawing.Point(0, 0)
        Me.ugMyFacs.Name = "ugMyFacs"
        Me.ugMyFacs.Size = New System.Drawing.Size(872, 141)
        Me.ugMyFacs.TabIndex = 1
        '
        'pnlAllInfoMyFacsTop
        '
        Me.pnlAllInfoMyFacsTop.Controls.Add(Me.lblMyFacs)
        Me.pnlAllInfoMyFacsTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlAllInfoMyFacsTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlAllInfoMyFacsTop.Name = "pnlAllInfoMyFacsTop"
        Me.pnlAllInfoMyFacsTop.Size = New System.Drawing.Size(872, 19)
        Me.pnlAllInfoMyFacsTop.TabIndex = 0
        '
        'lblMyFacs
        '
        Me.lblMyFacs.Location = New System.Drawing.Point(8, 0)
        Me.lblMyFacs.Name = "lblMyFacs"
        Me.lblMyFacs.Size = New System.Drawing.Size(80, 19)
        Me.lblMyFacs.TabIndex = 0
        Me.lblMyFacs.Text = "My Facilities"
        '
        'tbPageMyFacilities
        '
        Me.tbPageMyFacilities.Controls.Add(Me.pnlMyFacilitiesDetails)
        Me.tbPageMyFacilities.Controls.Add(Me.pnlMyFacilitiesTop)
        Me.tbPageMyFacilities.Controls.Add(Me.pnlMyFacilitiesBottom)
        Me.tbPageMyFacilities.Location = New System.Drawing.Point(4, 22)
        Me.tbPageMyFacilities.Name = "tbPageMyFacilities"
        Me.tbPageMyFacilities.Size = New System.Drawing.Size(872, 615)
        Me.tbPageMyFacilities.TabIndex = 1
        Me.tbPageMyFacilities.Text = "My Facilities"
        Me.tbPageMyFacilities.Visible = False
        '
        'pnlMyFacilitiesDetails
        '
        Me.pnlMyFacilitiesDetails.Controls.Add(Me.ugMyFacilities)
        Me.pnlMyFacilitiesDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlMyFacilitiesDetails.Location = New System.Drawing.Point(0, 21)
        Me.pnlMyFacilitiesDetails.Name = "pnlMyFacilitiesDetails"
        Me.pnlMyFacilitiesDetails.Size = New System.Drawing.Size(872, 554)
        Me.pnlMyFacilitiesDetails.TabIndex = 3
        '
        'ugMyFacilities
        '
        Me.ugMyFacilities.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugMyFacilities.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugMyFacilities.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugMyFacilities.Location = New System.Drawing.Point(0, 0)
        Me.ugMyFacilities.Name = "ugMyFacilities"
        Me.ugMyFacilities.Size = New System.Drawing.Size(872, 554)
        Me.ugMyFacilities.TabIndex = 1
        '
        'pnlMyFacilitiesTop
        '
        Me.pnlMyFacilitiesTop.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlMyFacilitiesTop.Controls.Add(Me.lblMyFacilities)
        Me.pnlMyFacilitiesTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlMyFacilitiesTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlMyFacilitiesTop.Name = "pnlMyFacilitiesTop"
        Me.pnlMyFacilitiesTop.Size = New System.Drawing.Size(872, 21)
        Me.pnlMyFacilitiesTop.TabIndex = 0
        '
        'lblMyFacilities
        '
        Me.lblMyFacilities.Location = New System.Drawing.Point(4, 2)
        Me.lblMyFacilities.Name = "lblMyFacilities"
        Me.lblMyFacilities.TabIndex = 0
        Me.lblMyFacilities.Text = "My Facilities"
        '
        'pnlMyFacilitiesBottom
        '
        Me.pnlMyFacilitiesBottom.Controls.Add(Me.btnMyFacSchedule)
        Me.pnlMyFacilitiesBottom.Controls.Add(Me.btnMyFacViewHistory)
        Me.pnlMyFacilitiesBottom.Controls.Add(Me.btnMyFacEnterCheckInfo)
        Me.pnlMyFacilitiesBottom.Controls.Add(Me.btnMyFacPrintChecklists)
        Me.pnlMyFacilitiesBottom.Controls.Add(Me.btnMyFacGenLetters)
        Me.pnlMyFacilitiesBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlMyFacilitiesBottom.Location = New System.Drawing.Point(0, 575)
        Me.pnlMyFacilitiesBottom.Name = "pnlMyFacilitiesBottom"
        Me.pnlMyFacilitiesBottom.Size = New System.Drawing.Size(872, 40)
        Me.pnlMyFacilitiesBottom.TabIndex = 2
        '
        'btnMyFacSchedule
        '
        Me.btnMyFacSchedule.Location = New System.Drawing.Point(592, 8)
        Me.btnMyFacSchedule.Name = "btnMyFacSchedule"
        Me.btnMyFacSchedule.Size = New System.Drawing.Size(80, 23)
        Me.btnMyFacSchedule.TabIndex = 11
        Me.btnMyFacSchedule.Text = "Schedule"
        '
        'btnMyFacViewHistory
        '
        Me.btnMyFacViewHistory.Location = New System.Drawing.Point(496, 8)
        Me.btnMyFacViewHistory.Name = "btnMyFacViewHistory"
        Me.btnMyFacViewHistory.Size = New System.Drawing.Size(88, 23)
        Me.btnMyFacViewHistory.TabIndex = 6
        Me.btnMyFacViewHistory.Text = "View History"
        '
        'btnMyFacEnterCheckInfo
        '
        Me.btnMyFacEnterCheckInfo.Location = New System.Drawing.Point(368, 8)
        Me.btnMyFacEnterCheckInfo.Name = "btnMyFacEnterCheckInfo"
        Me.btnMyFacEnterCheckInfo.Size = New System.Drawing.Size(120, 23)
        Me.btnMyFacEnterCheckInfo.TabIndex = 5
        Me.btnMyFacEnterCheckInfo.Text = "Enter Checklist Info"
        '
        'btnMyFacPrintChecklists
        '
        Me.btnMyFacPrintChecklists.Location = New System.Drawing.Point(248, 8)
        Me.btnMyFacPrintChecklists.Name = "btnMyFacPrintChecklists"
        Me.btnMyFacPrintChecklists.Size = New System.Drawing.Size(112, 23)
        Me.btnMyFacPrintChecklists.TabIndex = 4
        Me.btnMyFacPrintChecklists.Text = "Print Checklist(s)"
        '
        'btnMyFacGenLetters
        '
        Me.btnMyFacGenLetters.Location = New System.Drawing.Point(128, 8)
        Me.btnMyFacGenLetters.Name = "btnMyFacGenLetters"
        Me.btnMyFacGenLetters.Size = New System.Drawing.Size(112, 23)
        Me.btnMyFacGenLetters.TabIndex = 3
        Me.btnMyFacGenLetters.Text = "Generate Letters"
        '
        'tbPageAllFacForOwner
        '
        Me.tbPageAllFacForOwner.Controls.Add(Me.pnlAllFacForOwnerDetails)
        Me.tbPageAllFacForOwner.Controls.Add(Me.pnlAllFacsForOwnerTop)
        Me.tbPageAllFacForOwner.Controls.Add(Me.pnlAllFacForOwnerBottom)
        Me.tbPageAllFacForOwner.Location = New System.Drawing.Point(4, 22)
        Me.tbPageAllFacForOwner.Name = "tbPageAllFacForOwner"
        Me.tbPageAllFacForOwner.Size = New System.Drawing.Size(872, 615)
        Me.tbPageAllFacForOwner.TabIndex = 2
        Me.tbPageAllFacForOwner.Text = "All Facilities for Selected Owner"
        Me.tbPageAllFacForOwner.Visible = False
        '
        'pnlAllFacForOwnerDetails
        '
        Me.pnlAllFacForOwnerDetails.Controls.Add(Me.ugAllFacForOwner)
        Me.pnlAllFacForOwnerDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlAllFacForOwnerDetails.Location = New System.Drawing.Point(0, 27)
        Me.pnlAllFacForOwnerDetails.Name = "pnlAllFacForOwnerDetails"
        Me.pnlAllFacForOwnerDetails.Size = New System.Drawing.Size(872, 548)
        Me.pnlAllFacForOwnerDetails.TabIndex = 3
        '
        'ugAllFacForOwner
        '
        Me.ugAllFacForOwner.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugAllFacForOwner.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugAllFacForOwner.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugAllFacForOwner.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugAllFacForOwner.Location = New System.Drawing.Point(0, 0)
        Me.ugAllFacForOwner.Name = "ugAllFacForOwner"
        Me.ugAllFacForOwner.Size = New System.Drawing.Size(872, 548)
        Me.ugAllFacForOwner.TabIndex = 1
        '
        'pnlAllFacsForOwnerTop
        '
        Me.pnlAllFacsForOwnerTop.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlAllFacsForOwnerTop.Controls.Add(Me.lblAllFacforOwner)
        Me.pnlAllFacsForOwnerTop.Controls.Add(Me.btnRestoreSortOrder2)
        Me.pnlAllFacsForOwnerTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlAllFacsForOwnerTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlAllFacsForOwnerTop.Name = "pnlAllFacsForOwnerTop"
        Me.pnlAllFacsForOwnerTop.Size = New System.Drawing.Size(872, 27)
        Me.pnlAllFacsForOwnerTop.TabIndex = 0
        '
        'lblAllFacforOwner
        '
        Me.lblAllFacforOwner.Location = New System.Drawing.Point(3, 5)
        Me.lblAllFacforOwner.Name = "lblAllFacforOwner"
        Me.lblAllFacforOwner.Size = New System.Drawing.Size(136, 23)
        Me.lblAllFacforOwner.TabIndex = 0
        Me.lblAllFacforOwner.Text = "All Facilities for Owner"
        '
        'btnRestoreSortOrder2
        '
        Me.btnRestoreSortOrder2.Enabled = False
        Me.btnRestoreSortOrder2.Location = New System.Drawing.Point(320, 0)
        Me.btnRestoreSortOrder2.Name = "btnRestoreSortOrder2"
        Me.btnRestoreSortOrder2.Size = New System.Drawing.Size(120, 23)
        Me.btnRestoreSortOrder2.TabIndex = 3
        Me.btnRestoreSortOrder2.Text = "Restore Sort Order"
        '
        'pnlAllFacForOwnerBottom
        '
        Me.pnlAllFacForOwnerBottom.Controls.Add(Me.btnAllFacSchedule)
        Me.pnlAllFacForOwnerBottom.Controls.Add(Me.btnAllFacViewHistory)
        Me.pnlAllFacForOwnerBottom.Controls.Add(Me.btnAllFacEntercheckInfo)
        Me.pnlAllFacForOwnerBottom.Controls.Add(Me.btnAllFacPrintChecklist)
        Me.pnlAllFacForOwnerBottom.Controls.Add(Me.btnAllFacGenLetters)
        Me.pnlAllFacForOwnerBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlAllFacForOwnerBottom.Location = New System.Drawing.Point(0, 575)
        Me.pnlAllFacForOwnerBottom.Name = "pnlAllFacForOwnerBottom"
        Me.pnlAllFacForOwnerBottom.Size = New System.Drawing.Size(872, 40)
        Me.pnlAllFacForOwnerBottom.TabIndex = 2
        '
        'btnAllFacSchedule
        '
        Me.btnAllFacSchedule.Location = New System.Drawing.Point(592, 9)
        Me.btnAllFacSchedule.Name = "btnAllFacSchedule"
        Me.btnAllFacSchedule.Size = New System.Drawing.Size(80, 23)
        Me.btnAllFacSchedule.TabIndex = 12
        Me.btnAllFacSchedule.Text = "Schedule"
        '
        'btnAllFacViewHistory
        '
        Me.btnAllFacViewHistory.Location = New System.Drawing.Point(496, 9)
        Me.btnAllFacViewHistory.Name = "btnAllFacViewHistory"
        Me.btnAllFacViewHistory.Size = New System.Drawing.Size(88, 23)
        Me.btnAllFacViewHistory.TabIndex = 6
        Me.btnAllFacViewHistory.Text = "View History"
        '
        'btnAllFacEntercheckInfo
        '
        Me.btnAllFacEntercheckInfo.Location = New System.Drawing.Point(368, 9)
        Me.btnAllFacEntercheckInfo.Name = "btnAllFacEntercheckInfo"
        Me.btnAllFacEntercheckInfo.Size = New System.Drawing.Size(120, 23)
        Me.btnAllFacEntercheckInfo.TabIndex = 5
        Me.btnAllFacEntercheckInfo.Text = "Enter Checklist Info"
        '
        'btnAllFacPrintChecklist
        '
        Me.btnAllFacPrintChecklist.Location = New System.Drawing.Point(248, 9)
        Me.btnAllFacPrintChecklist.Name = "btnAllFacPrintChecklist"
        Me.btnAllFacPrintChecklist.Size = New System.Drawing.Size(112, 23)
        Me.btnAllFacPrintChecklist.TabIndex = 4
        Me.btnAllFacPrintChecklist.Text = "Print Checklist(s)"
        '
        'btnAllFacGenLetters
        '
        Me.btnAllFacGenLetters.Location = New System.Drawing.Point(128, 9)
        Me.btnAllFacGenLetters.Name = "btnAllFacGenLetters"
        Me.btnAllFacGenLetters.Size = New System.Drawing.Size(112, 23)
        Me.btnAllFacGenLetters.TabIndex = 3
        Me.btnAllFacGenLetters.Text = "Generate Letters"
        '
        'tbPageAssignedInspec
        '
        Me.tbPageAssignedInspec.Controls.Add(Me.pnlAssignedInspecDetails)
        Me.tbPageAssignedInspec.Controls.Add(Me.pnlAssignedInspectop)
        Me.tbPageAssignedInspec.Controls.Add(Me.pnlAssignedInspecBottom)
        Me.tbPageAssignedInspec.Location = New System.Drawing.Point(4, 22)
        Me.tbPageAssignedInspec.Name = "tbPageAssignedInspec"
        Me.tbPageAssignedInspec.Size = New System.Drawing.Size(872, 615)
        Me.tbPageAssignedInspec.TabIndex = 3
        Me.tbPageAssignedInspec.Text = "Assigned Inspections"
        Me.tbPageAssignedInspec.Visible = False
        '
        'pnlAssignedInspecDetails
        '
        Me.pnlAssignedInspecDetails.Controls.Add(Me.ugAssignedInspec)
        Me.pnlAssignedInspecDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlAssignedInspecDetails.Location = New System.Drawing.Point(0, 24)
        Me.pnlAssignedInspecDetails.Name = "pnlAssignedInspecDetails"
        Me.pnlAssignedInspecDetails.Size = New System.Drawing.Size(872, 551)
        Me.pnlAssignedInspecDetails.TabIndex = 3
        '
        'ugAssignedInspec
        '
        Me.ugAssignedInspec.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugAssignedInspec.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugAssignedInspec.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugAssignedInspec.Location = New System.Drawing.Point(0, 0)
        Me.ugAssignedInspec.Name = "ugAssignedInspec"
        Me.ugAssignedInspec.Size = New System.Drawing.Size(872, 551)
        Me.ugAssignedInspec.TabIndex = 1
        '
        'pnlAssignedInspectop
        '
        Me.pnlAssignedInspectop.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlAssignedInspectop.Controls.Add(Me.lblAssignedInspec)
        Me.pnlAssignedInspectop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlAssignedInspectop.Location = New System.Drawing.Point(0, 0)
        Me.pnlAssignedInspectop.Name = "pnlAssignedInspectop"
        Me.pnlAssignedInspectop.Size = New System.Drawing.Size(872, 24)
        Me.pnlAssignedInspectop.TabIndex = 0
        '
        'lblAssignedInspec
        '
        Me.lblAssignedInspec.Location = New System.Drawing.Point(3, 2)
        Me.lblAssignedInspec.Name = "lblAssignedInspec"
        Me.lblAssignedInspec.Size = New System.Drawing.Size(128, 23)
        Me.lblAssignedInspec.TabIndex = 0
        Me.lblAssignedInspec.Text = "Assigned Inspections"
        '
        'pnlAssignedInspecBottom
        '
        Me.pnlAssignedInspecBottom.Controls.Add(Me.btnAssignedInspecPrint)
        Me.pnlAssignedInspecBottom.Controls.Add(Me.btnAssignedInspecEnterInspec)
        Me.pnlAssignedInspecBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlAssignedInspecBottom.Location = New System.Drawing.Point(0, 575)
        Me.pnlAssignedInspecBottom.Name = "pnlAssignedInspecBottom"
        Me.pnlAssignedInspecBottom.Size = New System.Drawing.Size(872, 40)
        Me.pnlAssignedInspecBottom.TabIndex = 2
        '
        'btnAssignedInspecPrint
        '
        Me.btnAssignedInspecPrint.Location = New System.Drawing.Point(344, 9)
        Me.btnAssignedInspecPrint.Name = "btnAssignedInspecPrint"
        Me.btnAssignedInspecPrint.Size = New System.Drawing.Size(136, 23)
        Me.btnAssignedInspecPrint.TabIndex = 4
        Me.btnAssignedInspecPrint.Text = "Print Inspection Info"
        '
        'btnAssignedInspecEnterInspec
        '
        Me.btnAssignedInspecEnterInspec.Location = New System.Drawing.Point(208, 9)
        Me.btnAssignedInspecEnterInspec.Name = "btnAssignedInspecEnterInspec"
        Me.btnAssignedInspecEnterInspec.Size = New System.Drawing.Size(128, 23)
        Me.btnAssignedInspecEnterInspec.TabIndex = 3
        Me.btnAssignedInspecEnterInspec.Text = "Enter Inspection Info"
        '
        'pnlInspectionHeader
        '
        Me.pnlInspectionHeader.Controls.Add(Me.cmbInspectors)
        Me.pnlInspectionHeader.Controls.Add(Me.chkShowAllTargetFacilities)
        Me.pnlInspectionHeader.Controls.Add(Me.grpBxSortBy)
        Me.pnlInspectionHeader.Controls.Add(Me.lblFacilitiesFor)
        Me.pnlInspectionHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlInspectionHeader.Location = New System.Drawing.Point(0, 0)
        Me.pnlInspectionHeader.Name = "pnlInspectionHeader"
        Me.pnlInspectionHeader.Size = New System.Drawing.Size(880, 53)
        Me.pnlInspectionHeader.TabIndex = 0
        '
        'cmbInspectors
        '
        Me.cmbInspectors.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbInspectors.Location = New System.Drawing.Point(24, 25)
        Me.cmbInspectors.Name = "cmbInspectors"
        Me.cmbInspectors.Size = New System.Drawing.Size(184, 21)
        Me.cmbInspectors.TabIndex = 3
        '
        'chkShowAllTargetFacilities
        '
        Me.chkShowAllTargetFacilities.Location = New System.Drawing.Point(272, 8)
        Me.chkShowAllTargetFacilities.Name = "chkShowAllTargetFacilities"
        Me.chkShowAllTargetFacilities.Size = New System.Drawing.Size(152, 24)
        Me.chkShowAllTargetFacilities.TabIndex = 1
        Me.chkShowAllTargetFacilities.Text = "Show All Target Facilities"
        '
        'grpBxSortBy
        '
        Me.grpBxSortBy.Controls.Add(Me.rdBtnScheduled)
        Me.grpBxSortBy.Controls.Add(Me.rdBtnDue)
        Me.grpBxSortBy.Location = New System.Drawing.Point(440, 8)
        Me.grpBxSortBy.Name = "grpBxSortBy"
        Me.grpBxSortBy.Size = New System.Drawing.Size(144, 42)
        Me.grpBxSortBy.TabIndex = 2
        Me.grpBxSortBy.TabStop = False
        Me.grpBxSortBy.Text = "Sort By"
        '
        'rdBtnScheduled
        '
        Me.rdBtnScheduled.Location = New System.Drawing.Point(65, 15)
        Me.rdBtnScheduled.Name = "rdBtnScheduled"
        Me.rdBtnScheduled.Size = New System.Drawing.Size(76, 24)
        Me.rdBtnScheduled.TabIndex = 3
        Me.rdBtnScheduled.Text = "Scheduled"
        '
        'rdBtnDue
        '
        Me.rdBtnDue.Checked = True
        Me.rdBtnDue.Location = New System.Drawing.Point(8, 15)
        Me.rdBtnDue.Name = "rdBtnDue"
        Me.rdBtnDue.Size = New System.Drawing.Size(48, 24)
        Me.rdBtnDue.TabIndex = 2
        Me.rdBtnDue.TabStop = True
        Me.rdBtnDue.Text = "Due"
        '
        'lblFacilitiesFor
        '
        Me.lblFacilitiesFor.Location = New System.Drawing.Point(16, 8)
        Me.lblFacilitiesFor.Name = "lblFacilitiesFor"
        Me.lblFacilitiesFor.Size = New System.Drawing.Size(128, 16)
        Me.lblFacilitiesFor.TabIndex = 0
        Me.lblFacilitiesFor.Text = "Facilities for Inspector:"
        Me.lblFacilitiesFor.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Inspection
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(880, 694)
        Me.Controls.Add(Me.pnlInspectionDetails)
        Me.Controls.Add(Me.pnlInspectionHeader)
        Me.Name = "Inspection"
        Me.Text = "Inspection Schedule"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.pnlInspectionDetails.ResumeLayout(False)
        Me.tbCtrlInspection.ResumeLayout(False)
        Me.tbPageAllInfo.ResumeLayout(False)
        Me.pnlAllInfoDetails.ResumeLayout(False)
        Me.pnlAllInfoAssignedInspecbottom.ResumeLayout(False)
        Me.pnlAllInfoAssignedInspecDetails.ResumeLayout(False)
        CType(Me.ugAssignedInspections, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlAllInfoAssignedInspecTop.ResumeLayout(False)
        Me.pnlAllInfoAllFacsBottom.ResumeLayout(False)
        Me.pnlAllInfoAllFacsDetails.ResumeLayout(False)
        CType(Me.ugAllFacsForOwner, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlAllInfoAllFacsTop.ResumeLayout(False)
        Me.pnlAllInfoMyFacsDetails.ResumeLayout(False)
        CType(Me.ugMyFacs, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlAllInfoMyFacsTop.ResumeLayout(False)
        Me.tbPageMyFacilities.ResumeLayout(False)
        Me.pnlMyFacilitiesDetails.ResumeLayout(False)
        CType(Me.ugMyFacilities, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlMyFacilitiesTop.ResumeLayout(False)
        Me.pnlMyFacilitiesBottom.ResumeLayout(False)
        Me.tbPageAllFacForOwner.ResumeLayout(False)
        Me.pnlAllFacForOwnerDetails.ResumeLayout(False)
        CType(Me.ugAllFacForOwner, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlAllFacsForOwnerTop.ResumeLayout(False)
        Me.pnlAllFacForOwnerBottom.ResumeLayout(False)
        Me.tbPageAssignedInspec.ResumeLayout(False)
        Me.pnlAssignedInspecDetails.ResumeLayout(False)
        CType(Me.ugAssignedInspec, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlAssignedInspectop.ResumeLayout(False)
        Me.pnlAssignedInspecBottom.ResumeLayout(False)
        Me.pnlInspectionHeader.ResumeLayout(False)
        Me.grpBxSortBy.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region "Tab Operations"
    Private Sub tbCtrlInspection_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbCtrlInspection.Click
        Cursor.Current = Cursors.AppStarting
        Select Case tbCtrlInspection.SelectedTab.Name
            Case tbPageAllInfo.Name
                If ugAssignedInspections.Rows.Count < 1 Then
                    btnEnterInspectionInfo.Enabled = False
                    btnPrintInspectionInfo.Enabled = False
                Else
                    btnEnterInspectionInfo.Enabled = True
                    btnPrintInspectionInfo.Enabled = True
                End If
                If ugMyFacs.Rows.Count < 1 Then
                    btnRestoreSortOrder.Enabled = False
                    btnGenerateLetters.Enabled = False
                    btnPrintChecklists.Enabled = False
                    btnEnterChecklistInfo.Enabled = False
                    btnViewHistory.Enabled = False
                    btnReschedule.Enabled = False
                Else
                    btnRestoreSortOrder.Enabled = True
                    btnGenerateLetters.Enabled = True
                    btnPrintChecklists.Enabled = True
                    btnEnterChecklistInfo.Enabled = True
                    btnViewHistory.Enabled = True
                    btnReschedule.Enabled = True
                End If
            Case tbPageMyFacilities.Name
                If ugMyFacilities.Rows.Count < 1 Then
                    btnMyFacGenLetters.Enabled = False
                    btnMyFacPrintChecklists.Enabled = False
                    btnMyFacEnterCheckInfo.Enabled = False
                    btnMyFacViewHistory.Enabled = False
                    btnMyFacSchedule.Enabled = False
                Else
                    btnMyFacGenLetters.Enabled = True
                    btnMyFacPrintChecklists.Enabled = True
                    btnMyFacEnterCheckInfo.Enabled = True
                    btnMyFacViewHistory.Enabled = True
                    btnMyFacSchedule.Enabled = True
                End If
            Case tbPageAllFacForOwner.Name
                If ugAllFacForOwner.Rows.Count < 1 Then
                    btnRestoreSortOrder2.Enabled = False
                    btnAllFacGenLetters.Enabled = False
                    btnAllFacPrintChecklist.Enabled = False
                    btnAllFacEntercheckInfo.Enabled = False
                    btnAllFacViewHistory.Enabled = False
                    btnAllFacSchedule.Enabled = False
                Else
                    btnRestoreSortOrder2.Enabled = True
                    btnAllFacGenLetters.Enabled = True
                    btnAllFacPrintChecklist.Enabled = True
                    btnAllFacEntercheckInfo.Enabled = True
                    btnAllFacViewHistory.Enabled = True
                    btnAllFacSchedule.Enabled = True
                End If
            Case tbPageAssignedInspec.Name
                If ugAssignedInspec.Rows.Count < 1 Then
                    btnAssignedInspecEnterInspec.Enabled = False
                    btnAssignedInspecPrint.Enabled = False
                Else
                    btnAssignedInspecEnterInspec.Enabled = True
                    btnAssignedInspecPrint.Enabled = True
                End If
        End Select
        Cursor.Current = Cursors.Default
    End Sub
#End Region
#Region "UI Support Routines"
    Public Sub LoadInspectors()
        Try
            bolLoading = True
            Dim dt As DataTable = oInspection.GetInspectors.Tables(0)
            ' if logged in user is not an inspector, add blank row to the table
            If dt.Select("STAFF_ID = " + MusterContainer.AppUser.UserKey.ToString).Length <= 0 Then

                Dim dr As DataRow
                dr = dt.NewRow
                dr("STAFF_ID") = 0
                dr("USER_NAME") = ""
                dt.Rows.Add(dr)
            End If
            dt.DefaultView.Sort = "USER_NAME"
            cmbInspectors.DataSource = dt.DefaultView
            cmbInspectors.ValueMember = "STAFF_ID"
            cmbInspectors.DisplayMember = "USER_NAME"
            bolLoading = False
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Public Sub ClearFacsForOwner()
        ugAllFacsForOwner.DataSource = Nothing
        ugAllFacForOwner.DataSource = Nothing
    End Sub
    Public Sub Populate(ByVal inspectorID As Integer, Optional ByVal nOwnerID As Integer = 0, Optional ByVal nFacID As Integer = 0, Optional ByVal promptUser As Boolean = False, Optional ByVal reLoadInspectionID As Integer = 0, Optional ByVal bolChangeTab As Boolean = True)
        Try
            PopulateMyFacilities(inspectorID, nOwnerID, nFacID, False, promptUser, reLoadInspectionID, bolChangeTab)
            PopulateAssignedInspections(inspectorID)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Public Sub PopulateMyFacilities(ByVal inspectorID As Integer, Optional ByVal nOwnerID As Integer = 0, Optional ByVal nFacID As Integer = 0, Optional ByVal allFacilities As Boolean = False, Optional ByVal promptUser As Boolean = False, Optional ByVal reLoadInspectionID As Integer = 0, Optional ByVal bolChangeTab As Boolean = True)
        Try
            nSelectedFacilityID = nFacID
            ClearFacsForOwner()
            ' reLoadInspectionID = 0, load all facs from db
            ' reLoadInspectionID = -1, do not load any thing from db
            ' reLoadInspectionID <> 0 and reLoadInspectionID <> -1, load only reLoadInspectionID's facility from db
            If reLoadInspectionID = 0 Or dsTargetFacs Is Nothing Then
                ' get facs from db
                dsTargetFacs = oInspection.GetTargetFacilities(inspectorID, , allFacilities)
            ElseIf reLoadInspectionID <> -1 Then
                ' get only the updated fac from db and modify the global dataset
                Dim ds As DataSet
                If reLoadInspectionID = -99 Then
                    ds = oInspection.GetTargetFacilities(inspectorID, , , , nSelectedFacilityID)
                Else
                    ds = oInspection.GetTargetFacilities(inspectorID, , , reLoadInspectionID)
                End If
                If Not ds Is Nothing Then
                    If ds.Tables.Count > 0 Then
                        If ds.Tables(0).Rows.Count > 0 Then
                            ' remove row from dsTargetFacs
                            Dim dr, drNew As DataRow
                            Dim drows() As DataRow = dsTargetFacs.Tables(0).Select("[FACILITY ID] = " + ds.Tables(0).Rows(0)("FACILITY ID").ToString)
                            If drows.Length > 0 Then
                                For Each dr In drows
                                    dsTargetFacs.Tables(0).Rows.Remove(dr)
                                Next
                                dsTargetFacs.AcceptChanges()
                            End If
                            For Each dr In ds.Tables(0).Rows
                                drNew = dsTargetFacs.Tables(0).NewRow
                                For Each dtCol As DataColumn In ds.Tables(0).Columns
                                    drNew(dtCol.ColumnName) = dr(dtCol.ColumnName)
                                Next
                                dsTargetFacs.Tables(0).Rows.Add(drNew)
                            Next
                            dsTargetFacs.AcceptChanges()
                        End If
                    End If
                End If
            End If

            SetupMyFacilities()



            If nOwnerID <> 0 And nSelectedFacilityID = 0 Then
                PopulateFacsforOwner(inspectorID, nOwnerID, nSelectedFacilityID, promptUser, reLoadInspectionID, bolChangeTab)
                ugMyFacs.Selected.Rows.Clear()
                ugMyFacilities.Selected.Rows.Clear()
                For Each ugRow In ugMyFacs.Rows
                    If ugRow.Cells("OWNER_ID").Value = nOwnerID Then
                        ugMyFacs.ActiveRow = ugRow
                        ugMyFacs.ActiveRow.Selected = True
                        ugMyFacs.ActiveRow.Activate()
                        Exit For
                    End If
                Next
                For Each ugRow In ugMyFacilities.Rows
                    If ugRow.Cells("OWNER_ID").Value = nOwnerID Then
                        ugMyFacilities.ActiveRow = ugRow
                        ugMyFacilities.ActiveRow.Selected = True
                        ugMyFacilities.ActiveRow.Activate()
                        Exit For
                    End If
                Next
            ElseIf nOwnerID <> 0 And nSelectedFacilityID <> 0 Then
                PopulateFacsforOwner(inspectorID, nOwnerID, nSelectedFacilityID, promptUser, reLoadInspectionID, bolChangeTab)
                'ugMyFacs.Selected.Rows.Clear()
                'ugMyFacilities.Selected.Rows.Clear()
                For Each ugRow In ugMyFacs.Rows
                    If ugRow.Cells("FACILITY ID").Value = nSelectedFacilityID Then
                        ugMyFacs.ActiveRow = ugRow
                        ugMyFacs.ActiveRow.Selected = True
                        ugMyFacs.ActiveRow.Activate()
                        Exit For
                    End If
                Next
                For Each ugRow In ugMyFacilities.Rows
                    If ugRow.Cells("FACILITY ID").Value = nSelectedFacilityID Then
                        ugMyFacilities.ActiveRow = ugRow
                        ugMyFacilities.ActiveRow.Selected = True
                        ugMyFacilities.ActiveRow.Activate()
                        Exit For
                    End If
                Next
            ElseIf nOwnerID = 0 And nSelectedFacilityID <> 0 Then
                ugMyFacs.Selected.Rows.Clear()
                ugMyFacilities.Selected.Rows.Clear()
                For Each ugRow In ugMyFacs.Rows
                    If ugRow.Cells("FACILITY ID").Value = nSelectedFacilityID Then
                        ugMyFacs.ActiveRow = ugRow
                        ugMyFacs.ActiveRow.Selected = True
                        Exit For
                    End If
                Next
                For Each ugRow In ugMyFacilities.Rows
                    If ugRow.Cells("FACILITY ID").Value = nSelectedFacilityID Then
                        ugMyFacilities.ActiveRow = ugRow
                        ugMyFacilities.ActiveRow.Selected = True
                        ugMyFacilities.ActiveRow.Activate()
                        Exit For
                    End If
                Next
            Else
                If ugMyFacs.Rows.Count > 0 Then
                    ugMyFacs.ActiveRow = ugMyFacs.Rows(0)
                    ugMyFacs.ActiveRow.Selected = True
                    ugMyFacs.ActiveRow.Activate()
                End If
                If ugMyFacilities.Rows.Count > 0 Then
                    ugMyFacilities.ActiveRow = ugMyFacilities.Rows(0)
                    ugMyFacilities.ActiveRow.Selected = True
                    ugMyFacilities.ActiveRow.Activate()
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Public Sub PopulateFacsforOwner(ByVal inspectorID As Integer, ByVal nOwnerID As Integer, ByVal nFacID As Integer, Optional ByVal promptUser As Boolean = False, Optional ByVal reLoadInspectionID As Integer = 0, Optional ByVal bolChangeTab As Boolean = True)
        Try
            nSelectedFacilityID = nFacID

            ' reLoadInspectionID = 0, load all facs from db
            ' reLoadInspectionID = -1, do not load any thing from db
            ' reLoadInspectionID <> 0 and reLoadInspectionID <> -1, load only reLoadInspectionID's facility from db
            If reLoadInspectionID = 0 Or dsTargetFacsForOwner Is Nothing Then
                dsTargetFacsForOwner = oInspection.GetTargetFacilities(inspectorID, nOwnerID)
            ElseIf reLoadInspectionID <> -1 Then
                ' get only the updated fac from db and modify the global dataset
                Dim ds As DataSet
                If reLoadInspectionID = -99 Then
                    ds = oInspection.GetTargetFacilities(inspectorID, , , , nSelectedFacilityID)
                Else
                    ds = oInspection.GetTargetFacilities(inspectorID, , , reLoadInspectionID)
                End If
                If Not ds Is Nothing Then
                    If ds.Tables.Count > 0 Then
                        If ds.Tables(0).Rows.Count > 0 Then
                            ' remove row from dsTargetFacsForOwner
                            Dim dr, drNew As DataRow
                            Dim drows() As DataRow = dsTargetFacsForOwner.Tables(0).Select("[FACILITY ID] = " + ds.Tables(0).Rows(0)("FACILITY ID").ToString)
                            If drows.Length > 0 Then
                                For Each dr In drows
                                    dsTargetFacsForOwner.Tables(0).Rows.Remove(dr)
                                Next
                                dsTargetFacsForOwner.AcceptChanges()
                            End If
                            For Each dr In ds.Tables(0).Rows
                                drNew = dsTargetFacsForOwner.Tables(0).NewRow
                                For Each dtCol As DataColumn In ds.Tables(0).Columns
                                    drNew(dtCol.ColumnName) = dr(dtCol.ColumnName)
                                Next
                                dsTargetFacsForOwner.Tables(0).Rows.Add(drNew)
                            Next
                            dsTargetFacsForOwner.AcceptChanges()
                        End If
                    End If
                End If
            End If

            ' set the search box with facility id
            If nFacID <> 0 Then
                Dim mc As MusterContainer = Me.MdiParent
                If Not mc Is Nothing Then
                    mc.txtOwnerQSKeyword.Text = nFacID
                End If
            End If

            SetupFacsforOwner(nOwnerID, nFacID)

            'if there were no facs - prompt user
            'else if user in my facilities tab switch to all facs for selected owner tab
            If dsTargetFacsForOwner.Tables.Count > 0 Then
                If dsTargetFacsForOwner.Tables(0).Rows.Count > 0 Then
                    If nOwnerID <> 0 And nFacID = 0 Then
                        If dsTargetFacsForOwner.Tables(0).Select("OWNER_ID = " + nOwnerID.ToString).Length > 0 Then
                            promptUser = False
                            ' if in my facs tab - switch to all facs for sel owner
                            If bolChangeTab Then
                                If tbCtrlInspection.SelectedTab Is Nothing Then
                                    tbCtrlInspection.SelectedTab = tbPageAllFacForOwner
                                Else
                                    If tbCtrlInspection.SelectedTab.Name = tbPageMyFacilities.Name Then
                                        tbCtrlInspection.SelectedTab = tbPageAllFacForOwner
                                    End If
                                End If
                            End If
                        End If
                    ElseIf nOwnerID <> 0 And nFacID <> 0 Then
                        If dsTargetFacsForOwner.Tables(0).Select("[FACILITY ID] = " + nFacID.ToString).Length > 0 Then
                            promptUser = False
                            ' if in my facs tab - switch to all facs for sel owner
                            If bolChangeTab Then
                                If tbCtrlInspection.SelectedTab Is Nothing Then
                                    tbCtrlInspection.SelectedTab = tbPageAllFacForOwner
                                Else
                                    If tbCtrlInspection.SelectedTab.Name = tbPageMyFacilities.Name Then
                                        tbCtrlInspection.SelectedTab = tbPageAllFacForOwner
                                    End If
                                End If
                            End If
                        End If
                    Else
                        promptUser = False
                    End If
                End If
            End If

            If promptUser Then
                MsgBox("No Records found", , "Search Results")
            End If

            If Not (ugMyFacs.ActiveRow Is Nothing) Then
                'ugMyFacs.ActiveRow.Selected = False
                ugMyFacs.ActiveRow = Nothing
            End If
            If Not (ugMyFacilities.ActiveRow Is Nothing) Then
                'ugMyFacilities.ActiveRow.Selected = False
                ugMyFacilities.ActiveRow = Nothing
            End If
            If ugMyFacs.Selected.Rows.Count > 0 Then
                ugMyFacs.Selected.Rows.Clear()
                'For Each ugRow In ugMyFacs.Selected.Rows
                '    ugRow.Selected = False
                'Next
            End If
            If ugMyFacilities.Selected.Rows.Count > 0 Then
                ugMyFacilities.Selected.Rows.Clear()
                'For Each ugRow In ugMyFacilities.Selected.Rows
                '    ugRow.Selected = False
                'Next
            End If
            If nSelectedFacilityID > 0 Then
                If ugAllFacsForOwner.Rows.Count > 0 Then
                    For Each ugRow In ugAllFacsForOwner.Rows
                        If ugRow.Cells("FACILITY ID").Value = nSelectedFacilityID Then
                            ugRow.Hidden = False
                            ugAllFacsForOwner.Rows.ExpandAll(False)
                            ugRow.Selected = True
                            ugRow.Activate()
                            Exit For
                        End If
                    Next
                End If
                If ugAllFacForOwner.Rows.Count > 0 Then
                    For Each ugRow In ugAllFacForOwner.Rows
                        If ugRow.Cells("FACILITY ID").Value = nSelectedFacilityID Then
                            ugRow.Hidden = False
                            ugAllFacForOwner.Rows.ExpandAll(False)
                            ugRow.Selected = True
                            ugRow.Activate()
                            Exit For
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Public Sub PopulateAssignedInspections(ByVal inspectorID As Integer)
        Try
            dsAssignedInspections = oInspection.GetAssignedFacilities(inspectorID)
            SetupAssignedInspections()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub



    Private Sub CreateToolTipMessage(ByVal row As Infragistics.Win.UltraWinGrid.UltraGridRow, ByVal sender As Object)

        Dim comments As New Form

        Dim txBox As TextBox

        comments.Text = String.Format("All Comments on Facility: {0}", row.Cells("FACILITY").Value)
        comments.Width = 600
        comments.Height = 400
        comments.WindowState = FormWindowState.Normal
        comments.MaximumSize = comments.Size
        comments.MinimumSize = comments.Size


        txBox = New TextBox
        With txBox
            .Multiline = True
            .Dock = DockStyle.None

            .BorderStyle = BorderStyle.Fixed3D
            .Top = 10
            .Left = 10
            .Width = 560
            .Height = 320
            .Text = row.Description.Replace(vbCrLf, vbCrLf + vbCrLf)
        End With

        comments.Controls.Add(txBox)
        comments.ShowDialog()

    End Sub


    Private Sub setSortAndStatus(ByRef ds As DataSet)

        'calculate status and sort status
        If Not ds.Tables Is Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each ugRow As DataRow In ds.Tables(0).Rows
                With ugRow

                    If .Item("TOTAL_TANKS") - .Item("POU_TANKS") = .Item("PENDING_CLOSURE_TANKS") Then
                        If .Item("TOTAL_TANKS") > 0 And .Item("PENDING_CLOSURE_TANKS") > 0 Then
                            .Item("STATUS") = "PENDING CLOSURE (ALL TANKS)"
                            .Item("STATUS FOR SORT") = 9
                        Else
                            .Item("STATUS") = "UPCOMMING INSTALL"
                            .Item("STATUS FOR SORT") = 10
                        End If

                    ElseIf .Item("TOTAL_TANKS") = .Item("POU_TANKS") And Not .Item("UPCOMING_INSTALLATION_DATE") Is DBNull.Value Then
                        .Item("STATUS FOR SORT") = 2
                        .Item("STATUS") = "UPCOMING INSTALL"
                    Else
                        .Item("STATUS FOR SORT") = 8
                        If .Item("LAST OWNER INSPECTION DATE") Is DBNull.Value Then
                            .Item("STATUS") = "NEW OWNER"
                            .Item("STATUS FOR SORT") = 0
                        ElseIf .Item("LAST INSPECTED ON") Is DBNull.Value Then
                            .Item("STATUS") = "NEW FACILITY"
                            .Item("STATUS FOR SORT") = 1
                        End If

                        If .Item("INSPECTION_ID") Is DBNull.Value Then
                            If Not .Item("LAST INSPECTED ON") Is DBNull.Value Then
                                If Date.Compare(.Item("LAST INSPECTED ON"), DateAdd(DateInterval.Year, -3, Today.Date)) < 0 Then
                                    .Item("STATUS") = "PAST DUE"
                                    .Item("STATUS FOR SORT") = 3
                                ElseIf Date.Compare(.Item("LAST INSPECTED ON"), DateAdd(DateInterval.Month, 3, DateAdd(DateInterval.Year, -3, Today.Date))) < 0 Then
                                    .Item("STATUS") = "DUE"
                                    .Item("STATUS FOR SORT") = 4

                                Else
                                    .Item("STATUS") = "INSPECTED"
                                    .Item("STATUS FOR SORT") = 7
                                End If
                            End If
                        Else
                            If Not .Item("SUBMITTED DATE") Is DBNull.Value AndAlso Date.Compare(IIf(.Item("LAST INSPECTED ON") Is DBNull.Value, .Item("SUBMITTED DATE"), .Item("LAST INSPECTED ON")), DateAdd(DateInterval.Year, -3, Today.Date)) >= 0 Then
                                .Item("STATUS") = "INSPECTED"
                                .Item("STATUS FOR SORT") = 7
                            ElseIf Not .Item("LETTER GENERATED") Is DBNull.Value AndAlso .Item("LETTER GENERATED") Then
                                .Item("STATUS") = "PENDING"
                                .Item("STATUS FOR SORT") = 5

                            ElseIf Not .Item("SCHEDULED DATE") Is DBNull.Value Then

                                .Item("STATUS") = "SCHEDULED"
                                .Item("STATUS FOR SORT") = 6

                            End If
                        End If
                        If .Item("LAST OWNER INSPECTION DATE") Is DBNull.Value Then
                            .Item("STATUS") = "NEW OWNER"
                            .Item("STATUS FOR SORT") = 0
                        End If




                    End If



                    ' scheduled date for sort
                    If .Item("SCHEDULED DATE") Is DBNull.Value Then
                        .Item("SCHEDULED DATE FOR SORT") = 1
                    Else
                        .Item("SCHEDULED DATE FOR SORT") = 0
                    End If

                    ' scheduled time for sort
                    If .Item("SCHEDULED TIME") Is DBNull.Value Then
                        .Item("SCHEDULED TIME FOR SORT") = 1
                    Else
                        .Item("SCHEDULED TIME FOR SORT") = 0
                    End If
                End With


            Next
        End If


    End Sub


    Private Sub SetupMyFacilities()
        Dim ug As Infragistics.Win.UltraWinGrid.UltraGrid
        Dim bolLoadingLocal As Boolean = bolLoading
        Try
            bolLoading = True
            'Select Case tbCtrlInspection.SelectedTab.Name
            '    Case tbPageAllInfo.Name
            '        ug = ugMyFacs
            '    Case tbPageMyFacilities.Name
            '        ug = ugMyFacilities
            '    Case Else
            '        Exit Sub
            'End Select











            For Each ug In arrugMyFacs

                RemoveHandler ug.MouseEnterElement, AddressOf ug_MouseDown
                AddHandler ug.MouseEnterElement, AddressOf ug_MouseDown

                RemoveHandler ug.MouseDown, AddressOf ug_MouseDownClick
                AddHandler ug.MouseDown, AddressOf ug_MouseDownClick

                setSortAndStatus(dsTargetFacs)

                ug.DataSource = dsTargetFacs
                ug.DisplayLayout.Override.TipStyleCell = Infragistics.Win.UltraWinGrid.TipStyle.Hide
                ug.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.False
                ug.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti

                ug.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
                ug.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
                ug.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False

                ug.DisplayLayout.Bands(0).Columns("LETTER GENERATED").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox

                ug.DisplayLayout.Bands(0).Columns("STATUS FOR SORT").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("OWNER_ID").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("STAFF_ID").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("INSPECTION_ID").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("RESCHEDULED DATE").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("RESCHEDULED TIME").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("DELETED").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("SCHEDULED DATE FOR SORT").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("SCHEDULED TIME FOR SORT").Hidden = True
                'ug.DisplayLayout.Bands(0).Columns("NEW FOR SORT").Hidden = True
                'ug.DisplayLayout.Bands(0).Columns("LAST OWNER INSPECTION DATE FOR SORT").Hidden = True
                'ug.DisplayLayout.Bands(0).Columns("LAST INSPECTED ON FOR SORT").Hidden = True
                'ug.DisplayLayout.Bands(0).Columns("LAST INSPECTED ON AT OWNER").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("SCHEDULED BY FOR SORT").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("UPCOMING_INSTALLATION_DATE").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("TOTAL_TANKS").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("POU_TANKS").Hidden = True

                ug.DisplayLayout.UseFixedHeaders = True
                ug.DisplayLayout.Override.FixedHeaderIndicator = Infragistics.Win.UltraWinGrid.FixedHeaderIndicator.None
                ug.DisplayLayout.Bands(0).Columns("STATUS").Header.Fixed = True
                ug.DisplayLayout.Bands(0).Columns("OWNER NAME").Header.Fixed = True
                ug.DisplayLayout.Bands(0).Columns("FACILITY").Header.Fixed = True
                ug.DisplayLayout.Bands(0).Columns("FACILITY ID").Header.Fixed = True

                Dim cnt As Integer = 0

                For Each ugRow In ug.Rows

                    If Not commentDataSet Is Nothing AndAlso commentDataSet.Tables.Count > 0 Then


                        If Me.chkShowAllTargetFacilities.Checked Then

                            Dim rows() As DataRow
                            Dim comments As String
                            Dim msg As String


                            rows = commentDataSet.Tables(0).Select(String.Format("[MasterID] = {0}", ugRow.Cells("FACILITY ID").Value))

                            If Not rows Is Nothing AndAlso rows.GetUpperBound(0) > -1 Then

                                For Each row As DataRow In rows
                                    msg = String.Format("{0:D}: ({3})  {1}   -{2}", row.Item("CREATED ON"), row.Item("COMMENT"), row.Item("CREATEDBY"), _
                                    row.Item("ENTITY DESC")).Replace(vbCrLf, " ")

                                    If Not comments Is Nothing AndAlso comments.IndexOf(msg) <= -1 Then
                                        comments = String.Format("{0}{1}{2}", comments, IIf(comments.Length > 0, vbCrLf, String.Empty), msg)
                                    ElseIf comments Is Nothing Then
                                        comments = msg
                                    End If

                                    msg = Nothing

                                Next row

                                rows = Nothing
                                ugRow.Description = comments


                                comments = String.Empty
                            End If

                            ugRow.Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
                            ugRow.Appearance.BackColor = Color.FromArgb(245, 245, 250)


                        Else
                            'ugRow.Appearance.BackColor = Color.White
                        End If


                    End If

                    If ugRow.Cells("SCHEDULED BY").Text.Trim.ToUpper <> MusterContainer.AppUser.Name.ToUpper.Trim And ugRow.Cells("SCHEDULED BY").Text.Trim <> String.Empty Then
                        ugRow.Cells("SCHEDULED BY").Appearance.BackColor = Color.Gray
                    End If


                    If chkShowAllTargetFacilities.Checked Then
                        If ugRow.Cells("STAFF_ID").Value Is System.DBNull.Value Then
                            ugRow.Cells("OWNER NAME").Appearance.BackColor = Color.Gray
                            ugRow.Cells("FACILITY ID").Appearance.BackColor = Color.Gray
                            ugRow.Cells("FACILITY").Appearance.BackColor = Color.Gray
                        End If
                    End If
                    cnt += 1
                Next

                ug.DisplayLayout.Bands(0).SortedColumns.Clear()
                ug.DisplayLayout.Bands(0).Columns("STATUS FOR SORT").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending

                If cnt = 0 Then
                    ChangeTargetFacilitiesLayout(ug)
                End If


                If rdBtnScheduled.Checked Then
                    'ChangeTargetFacilitiesLayout(4, 5, 6, 7, 8, ug)
                    ug.DisplayLayout.Bands(0).Columns("SCHEDULED DATE FOR SORT").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                    ug.DisplayLayout.Bands(0).Columns("SCHEDULED DATE").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                    ug.DisplayLayout.Bands(0).Columns("SCHEDULED TIME FOR SORT").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                    ug.DisplayLayout.Bands(0).Columns("SCHEDULED TIME").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                    ug.DisplayLayout.Bands(0).Columns("OWNER NAME").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                    ug.DisplayLayout.Bands(0).Columns("FACILITY").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                Else
                    'ChangeTargetFacilitiesLayout(7, 8, 4, 5, 6, ug)
                    ug.DisplayLayout.Bands(0).Columns("STATUS FOR SORT").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                    'ug.DisplayLayout.Bands(0).Columns("NEW FOR SORT").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                    ug.DisplayLayout.Bands(0).Columns("LAST INSPECTED ON").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                    ug.DisplayLayout.Bands(0).Columns("LAST OWNER INSPECTION DATE").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                    ug.DisplayLayout.Bands(0).Columns("COUNTY").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                    ug.DisplayLayout.Bands(0).Columns("CITY").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                End If
            Next



        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            bolLoading = bolLoadingLocal
        End Try
    End Sub
    Private Sub SetupFacsforOwner(ByVal nOwnerID As Integer, ByVal nFacID As Integer)
        Dim ug As Infragistics.Win.UltraWinGrid.UltraGrid
        Dim bolLoadingLocal As Boolean = bolLoading
        Try
            bolLoading = True
            'Select Case tbCtrlInspection.SelectedTab.Name
            '    Case tbPageAllInfo.Name
            '        ug = ugAllFacsForOwner
            '    Case tbPageAllFacForOwner.Name
            '        ug = ugAllFacForOwner
            '    Case Else
            '        Exit Sub
            'End Select

            For Each ug In arrugAllFacsforOwner
                ug.DataSource = dsTargetFacsForOwner

                ug.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.False
                ug.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti

                ug.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
                ug.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
                ug.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False

                ug.DisplayLayout.Bands(0).Columns("LETTER GENERATED").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox

                ug.DisplayLayout.Bands(0).Columns("STATUS FOR SORT").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("OWNER_ID").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("STAFF_ID").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("INSPECTION_ID").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("RESCHEDULED DATE").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("RESCHEDULED TIME").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("DELETED").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("SCHEDULED DATE FOR SORT").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("SCHEDULED TIME FOR SORT").Hidden = True
                'ug.DisplayLayout.Bands(0).Columns("NEW FOR SORT").Hidden = True
                'ug.DisplayLayout.Bands(0).Columns("LAST OWNER INSPECTION DATE FOR SORT").Hidden = True
                'ug.DisplayLayout.Bands(0).Columns("LAST INSPECTED ON FOR SORT").Hidden = True
                'ug.DisplayLayout.Bands(0).Columns("LAST INSPECTED ON AT OWNER").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("SCHEDULED BY FOR SORT").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("UPCOMING_INSTALLATION_DATE").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("TOTAL_TANKS").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("POU_TANKS").Hidden = True

                ug.DisplayLayout.UseFixedHeaders = True
                ug.DisplayLayout.Override.FixedHeaderIndicator = Infragistics.Win.UltraWinGrid.FixedHeaderIndicator.None
                ug.DisplayLayout.Bands(0).Columns("STATUS").Header.Fixed = True
                ug.DisplayLayout.Bands(0).Columns("OWNER NAME").Header.Fixed = True
                ug.DisplayLayout.Bands(0).Columns("FACILITY").Header.Fixed = True
                ug.DisplayLayout.Bands(0).Columns("FACILITY ID").Header.Fixed = True

                For Each ugRow In ug.Rows


                    If ugRow.Cells("SCHEDULED BY").Text.Trim.ToUpper <> MusterContainer.AppUser.Name.ToUpper.Trim And ugRow.Cells("SCHEDULED BY").Text.Trim <> String.Empty Then
                        ugRow.Cells("SCHEDULED BY").Appearance.BackColor = Color.Gray
                    End If
                    If ugRow.Cells("STAFF_ID").Value Is System.DBNull.Value Then
                        ugRow.Cells("OWNER NAME").Appearance.BackColor = Color.Gray
                        ugRow.Cells("FACILITY ID").Appearance.BackColor = Color.Gray
                        ugRow.Cells("FACILITY").Appearance.BackColor = Color.Gray
                        If nFacID <> 0 Then
                            If ugRow.Cells("OWNER_ID").Value = nOwnerID And ugRow.Cells("FACILITY ID").Value = nFacID Then
                                ug.ActiveRow = ugRow
                            End If
                        End If
                    End If

                    ' calculate status and sort status
                    If ugRow.Cells("TOTAL_TANKS").Value - ugRow.Cells("POU_TANKS").Value = ugRow.Cells("PENDING_CLOSURE_TANKS").Value Then
                        If ugRow.Cells("TOTAL_TANKS").Value > 0 And ugRow.Cells("PENDING_CLOSURE_TANKS").Value > 0 Then
                            ugRow.Cells("STATUS").Value = "PENDING CLOSURE (ALL TANKS)"
                            ugRow.Cells("STATUS FOR SORT").Value = 9
                        Else
                            ugRow.Cells("STATUS").Value = "UPCOMMING INSTALL"
                            ugRow.Cells("STATUS FOR SORT").Value = 10
                        End If

                    ElseIf ugRow.Cells("TOTAL_TANKS").Value = ugRow.Cells("POU_TANKS").Value And Not ugRow.Cells("UPCOMING_INSTALLATION_DATE").Value Is DBNull.Value Then
                        ugRow.Cells("STATUS FOR SORT").Value = 2
                        ugRow.Cells("STATUS").Value = "UPCOMING INSTALL"
                    Else
                        ugRow.Cells("STATUS FOR SORT").Value = 8
                        If ugRow.Cells("LAST OWNER INSPECTION DATE").Value Is DBNull.Value Then
                            ugRow.Cells("STATUS").Value = "NEW OWNER"
                            ugRow.Cells("STATUS FOR SORT").Value = 0
                        ElseIf ugRow.Cells("LAST INSPECTED ON").Value Is DBNull.Value Then
                            ugRow.Cells("STATUS").Value = "NEW FACILITY"
                            ugRow.Cells("STATUS FOR SORT").Value = 1
                        End If




                        If ugRow.Cells("INSPECTION_ID").Value Is DBNull.Value Then
                            If Not ugRow.Cells("LAST INSPECTED ON").Value Is DBNull.Value Then
                                If Date.Compare(ugRow.Cells("LAST INSPECTED ON").Value, DateAdd(DateInterval.Year, -3, Today.Date)) < 0 Then
                                    ugRow.Cells("STATUS").Value = "PAST DUE"
                                    ugRow.Cells("STATUS FOR SORT").Value = 3
                                ElseIf Date.Compare(ugRow.Cells("LAST INSPECTED ON").Value, DateAdd(DateInterval.Month, -3, DateAdd(DateInterval.Year, -3, Today.Date))) < 0 Then
                                    ugRow.Cells("STATUS").Value = "DUE"
                                    ugRow.Cells("STATUS FOR SORT").Value = 4

                                Else
                                    ugRow.Cells("STATUS").Value = "INSPECTED"
                                    ugRow.Cells("STATUS FOR SORT").Value = 7
                                End If
                            End If
                        Else
                            If Not ugRow.Cells("SUBMITTED DATE").Value Is DBNull.Value AndAlso Date.Compare(IIf(ugRow.Cells("LAST INSPECTED ON").Value Is DBNull.Value, ugRow.Cells("SUBMITTED DATE").Value, ugRow.Cells("LAST INSPECTED ON").Value), DateAdd(DateInterval.Year, -3, Today.Date)) >= 0 Then
                                ugRow.Cells("STATUS").Value = "INSPECTED"
                                ugRow.Cells("STATUS FOR SORT").Value = 7
                            ElseIf Not ugRow.Cells("LETTER GENERATED").Value Is DBNull.Value AndAlso ugRow.Cells("LETTER GENERATED").Value Then
                                ugRow.Cells("STATUS").Value = "PENDING"
                                ugRow.Cells("STATUS FOR SORT").Value = 5

                            ElseIf Not ugRow.Cells("SCHEDULED DATE").Value Is DBNull.Value Then

                                ugRow.Cells("STATUS").Value = "SCHEDULED"
                                ugRow.Cells("STATUS FOR SORT").Value = 6

                            End If
                        End If
                        If ugRow.Cells("LAST OWNER INSPECTION DATE").Value Is DBNull.Value Then
                            ugRow.Cells("STATUS").Value = "NEW OWNER"
                            ugRow.Cells("STATUS FOR SORT").Value = 0
                        End If
                    End If

                    'If ugRow.Cells("STATUS").Text.Trim.ToUpper = "NEW OWNER" Then
                    '    ugRow.Cells("STATUS FOR SORT").Value = 0
                    'ElseIf ugRow.Cells("STATUS").Text.Trim.ToUpper = "NEW FACILITY" Then
                    '    ugRow.Cells("STATUS FOR SORT").Value = 1
                    'ElseIf ugRow.Cells("STATUS").Text.Trim.ToUpper = "DUE" Then
                    '    ugRow.Cells("STATUS FOR SORT").Value = 2
                    'ElseIf ugRow.Cells("STATUS").Text.Trim.ToUpper = "SCHEDULED" Then
                    '    ugRow.Cells("STATUS FOR SORT").Value = 3
                    'ElseIf ugRow.Cells("STATUS").Text.Trim.ToUpper = "PENDING" Then
                    '    ugRow.Cells("STATUS FOR SORT").Value = 4
                    'ElseIf ugRow.Cells("STATUS").Text.Trim.ToUpper = "INSPECTED" Then
                    '    ugRow.Cells("STATUS FOR SORT").Value = 5
                    'Else
                    '    ugRow.Cells("STATUS FOR SORT").Value = 6
                    'End If

                    ' scheduled date for sort
                    If ugRow.Cells("SCHEDULED DATE").Value Is DBNull.Value Then
                        ugRow.Cells("SCHEDULED DATE FOR SORT").Value = 1
                    Else
                        ugRow.Cells("SCHEDULED DATE FOR SORT").Value = 0
                    End If

                    ' scheduled time for sort
                    If ugRow.Cells("SCHEDULED TIME").Value Is DBNull.Value Then
                        ugRow.Cells("SCHEDULED TIME FOR SORT").Value = 1
                    Else
                        ugRow.Cells("SCHEDULED TIME FOR SORT").Value = 0
                    End If
                Next

                ug.DisplayLayout.Bands(0).SortedColumns.Clear()

                ChangeTargetFacilitiesLayout(ug)
                'If rdBtnScheduled.Checked Then
                '    ChangeTargetFacilitiesLayout(4, 5, 6, 7, 8, ug)
                'Else
                '    ChangeTargetFacilitiesLayout(7, 8, 4, 5, 6, ug)
                'End If

                ug.DisplayLayout.Bands(0).Columns("SCHEDULED BY FOR SORT").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                ug.DisplayLayout.Bands(0).Columns("STATUS FOR SORT").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                ug.DisplayLayout.Bands(0).Columns("LAST INSPECTED ON").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                ug.DisplayLayout.Bands(0).Columns("LAST OWNER INSPECTION DATE").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                ug.DisplayLayout.Bands(0).Columns("COUNTY").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                ug.DisplayLayout.Bands(0).Columns("CITY").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending

                'ug.DisplayLayout.Bands(0).Columns("STATUS FOR SORT").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                'ug.DisplayLayout.Bands(0).Columns("STAFF_ID").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Descending
                'ug.DisplayLayout.Bands(0).Columns("SCHEDULED DATE FOR SORT").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Descending
                'ug.DisplayLayout.Bands(0).Columns("SCHEDULED DATE").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                'ug.DisplayLayout.Bands(0).Columns("SCHEDULED TIME FOR SORT").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Descending
                'ug.DisplayLayout.Bands(0).Columns("SCHEDULED TIME").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending

                If dsTargetFacsForOwner.Tables(0).Rows.Count > 0 Then
                    btnRestoreSortOrder.Enabled = True
                    btnRestoreSortOrder2.Enabled = True
                Else
                    btnRestoreSortOrder.Enabled = False
                    btnRestoreSortOrder2.Enabled = False
                End If
            Next
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            bolLoading = bolLoadingLocal
        End Try
    End Sub
    Private Sub SetupAssignedInspections()
        Dim ug As Infragistics.Win.UltraWinGrid.UltraGrid
        Try
            'Select Case tbCtrlInspection.SelectedTab.Name
            '    Case tbPageAllInfo.Name
            '        ug = ugAssignedInspections
            '    Case tbPageAssignedInspec.Name
            '        ug = ugAssignedInspec
            '    Case Else
            '        Exit Sub
            'End Select

            For Each ug In arrugAssignedInspections
                ug.DataSource = dsAssignedInspections

                ug.DisplayLayout.Bands(0).SortedColumns.Clear()

                ug.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.False
                ug.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti

                ug.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
                ug.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
                ug.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False

                ug.DisplayLayout.Bands(0).Columns("INSPECTION_ID").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("INSPECTION_TYPE_ID").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("STAFF_ID").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("OWNER PHONE").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("COUNTY").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("ADDRESS_LINE_ONE").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("CITY").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("ADMIN COMMENTS").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("SUBMITTED").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("INSPECTOR COMMENTS").Hidden = True
                ug.DisplayLayout.Bands(0).Columns("INSPECTOR").Hidden = True

                ug.DisplayLayout.Bands(0).Columns("ASSIGNED DATE").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending

                If ug.Rows.Count <= 0 Then
                    btnEnterInspectionInfo.Enabled = False
                    btnAssignedInspecEnterInspec.Enabled = False
                    btnPrintInspectionInfo.Enabled = False
                    btnAssignedInspecPrint.Enabled = False
                    btnEnterInspectionInfo.Text = "Enter Inspection Info"
                    btnAssignedInspecEnterInspec.Text = "Enter Inspection Info"
                Else
                    btnEnterInspectionInfo.Enabled = True
                    btnAssignedInspecEnterInspec.Enabled = True
                    If ug.Rows(0).Cells("INSPECTOR").Text.Trim.ToUpper = MusterContainer.AppUser.Name.ToUpper.Trim Then
                        btnEnterInspectionInfo.Text = "Enter Inspection Info"
                        btnAssignedInspecEnterInspec.Text = "Enter Inspection Info"
                        btnPrintInspectionInfo.Enabled = True
                        btnAssignedInspecPrint.Enabled = True
                    Else
                        btnEnterInspectionInfo.Text = "View Inspection Info"
                        btnAssignedInspecEnterInspec.Text = "View Inspection Info"
                        btnPrintInspectionInfo.Enabled = False
                        btnAssignedInspecPrint.Enabled = False
                    End If
                End If
            Next
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ChangeTargetFacilitiesLayout(ByRef ThisGrid As Infragistics.Win.UltraWinGrid.UltraGrid)
        Try
            With ThisGrid.DisplayLayout
                .Bands(0).Columns("STATUS").Header.VisiblePosition = 0
                .Bands(0).Columns("OWNER NAME").Header.VisiblePosition = 1
                .Bands(0).Columns("FACILITY ID").Header.VisiblePosition = 2
                .Bands(0).Columns("FACILITY").Header.VisiblePosition = 3

                If rdBtnScheduled.Checked Then
                    .Bands(0).Columns("LAST INSPECTED ON").Header.VisiblePosition = 4
                    .Bands(0).Columns("SCHEDULED DATE").Header.VisiblePosition = 5
                    .Bands(0).Columns("SCHEDULED TIME").Header.VisiblePosition = 6
                    .Bands(0).Columns("CITY").Header.VisiblePosition = 7
                    .Bands(0).Columns("COUNTY").Header.VisiblePosition = 8
                    .Bands(0).Columns("SCHEDULED BY").Header.VisiblePosition = 9
                    .Bands(0).Columns("SUBMITTED DATE").Header.VisiblePosition = 10
                    .Bands(0).Columns("LETTER GENERATED").Header.VisiblePosition = 11
                    .Bands(0).Columns("CHECKLIST GENERATED").Header.VisiblePosition = 12
                    .Bands(0).Columns("LAST OWNER INSPECTION DATE").Header.VisiblePosition = 13
                Else
                    .Bands(0).Columns("CITY").Header.VisiblePosition = 4
                    .Bands(0).Columns("COUNTY").Header.VisiblePosition = 5
                    .Bands(0).Columns("LAST INSPECTED ON").Header.VisiblePosition = 6
                    .Bands(0).Columns("SCHEDULED DATE").Header.VisiblePosition = 7
                    .Bands(0).Columns("SCHEDULED TIME").Header.VisiblePosition = 8
                    .Bands(0).Columns("SCHEDULED BY").Header.VisiblePosition = 9
                    .Bands(0).Columns("SUBMITTED DATE").Header.VisiblePosition = 10
                    .Bands(0).Columns("LETTER GENERATED").Header.VisiblePosition = 11
                    .Bands(0).Columns("CHECKLIST GENERATED").Header.VisiblePosition = 12
                    .Bands(0).Columns("LAST OWNER INSPECTION DATE").Header.VisiblePosition = 13
                End If
                '.Bands(0).Columns("OWNER_ID").Header.VisiblePosition = 14
                '.Bands(0).Columns("STAFF_ID").Header.VisiblePosition = 15
                '.Bands(0).Columns("INSPECTION_ID").Header.VisiblePosition = 16
                '.Bands(0).Columns("STATUS FOR SORT").Header.VisiblePosition = 17
                '.Bands(0).Columns("SCHEDULED DATE FOR SORT").Header.VisiblePosition = 18
                '.Bands(0).Columns("SCHEDULED TIME FOR SORT").Header.VisiblePosition = 19
                '.Bands(0).Columns("SCHEDULED BY FOR SORT").Header.VisiblePosition = 20
            End With
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    'Private Sub ChangeTargetFacilitiesLayout(ByVal nSchDTPos As Int64, ByVal nSchTimePos As Int64, ByVal nCityPos As Int64, ByVal nCountyPos As Int64, ByVal nLstInspecOnPos As Int64, Optional ByRef ThisGrid As Infragistics.Win.UltraWinGrid.UltraGrid = Nothing)
    '    Try
    '        With ThisGrid.DisplayLayout
    '            .Bands(0).Columns("STATUS").Header.VisiblePosition = 0
    '            .Bands(0).Columns("OWNER NAME").Header.VisiblePosition = 1
    '            .Bands(0).Columns("FACILITY ID").Header.VisiblePosition = 2
    '            .Bands(0).Columns("FACILITY").Header.VisiblePosition = 3
    '            .Bands(0).Columns("SCHEDULED DATE").Header.VisiblePosition = nSchDTPos
    '            .Bands(0).Columns("SCHEDULED TIME").Header.VisiblePosition = nSchTimePos
    '            .Bands(0).Columns("CITY").Header.VisiblePosition = nCityPos
    '            .Bands(0).Columns("COUNTY").Header.VisiblePosition = nCountyPos
    '            .Bands(0).Columns("LAST INSPECTED ON").Header.VisiblePosition = nLstInspecOnPos
    '            .Bands(0).Columns("SCHEDULED BY").Header.VisiblePosition = 9
    '            .Bands(0).Columns("SUBMITTED DATE").Header.VisiblePosition = 10
    '            .Bands(0).Columns("LETTER GENERATED").Header.VisiblePosition = 11
    '            .Bands(0).Columns("CHECKLIST GENERATED").Header.VisiblePosition = 12
    '            .Bands(0).Columns("LAST OWNER INSPECTION DATE").Header.VisiblePosition = 13
    '            .Bands(0).Columns("OWNER_ID").Header.VisiblePosition = 14
    '            .Bands(0).Columns("STAFF_ID").Header.VisiblePosition = 15
    '            .Bands(0).Columns("INSPECTION_ID").Header.VisiblePosition = 16
    '            .Bands(0).Columns("STATUS FOR SORT").Header.VisiblePosition = 17
    '            .Bands(0).Columns("SCHEDULED DATE FOR SORT").Header.VisiblePosition = 18
    '            .Bands(0).Columns("SCHEDULED TIME FOR SORT").Header.VisiblePosition = 19
    '            .Bands(0).Columns("SCHEDULED BY FOR SORT").Header.VisiblePosition = 20
    '        End With
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub

    Public Function getSelectedRow() As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            ugRow = Nothing
            Select Case tbCtrlInspection.SelectedTab.Name
                Case tbPageAllInfo.Name
                    If Me.ugAllFacsForOwner.Rows.Count > 0 Then
                        If ugAllFacsForOwner.ActiveRow Is Nothing Then
                            ' if there is no active row for all facs for owner, 
                            ' check if there is active row for my facs
                            If Not ugMyFacs.ActiveRow Is Nothing Then
                                If ugMyFacs.ActiveRow.Selected Then
                                    ugRow = ugMyFacs.ActiveRow
                                    ugRow.Tag = ugMyFacs.Name
                                End If
                            End If
                        Else
                            If ugAllFacsForOwner.ActiveRow.Selected Then
                                ugRow = ugAllFacsForOwner.ActiveRow
                                ugRow.Tag = ugAllFacsForOwner.Name
                            End If
                        End If
                    ElseIf ugMyFacs.Rows.Count > 0 Then
                        If Not ugMyFacs.ActiveRow Is Nothing Then
                            If ugMyFacs.ActiveRow.Selected Then
                                ugRow = ugMyFacs.ActiveRow
                                ugRow.Tag = ugMyFacs.Name
                            End If
                        End If
                    End If
                Case tbPageMyFacilities.Name
                    If ugMyFacilities.Rows.Count > 0 Then
                        If Not ugMyFacilities.ActiveRow Is Nothing Then
                            If ugMyFacilities.ActiveRow.Selected Then
                                ugRow = ugMyFacilities.ActiveRow
                                ugRow.Tag = ugMyFacilities.Name
                            End If
                        End If
                    End If
                Case tbPageAllFacForOwner.Name
                    If ugAllFacForOwner.Rows.Count > 0 Then
                        If Not ugAllFacForOwner.ActiveRow Is Nothing Then
                            If ugAllFacForOwner.ActiveRow.Selected Then
                                ugRow = ugAllFacForOwner.ActiveRow
                                ugRow.Tag = ugAllFacForOwner.Name
                            End If
                        End If
                    End If
                Case tbPageAssignedInspec.Name
                    If ugAssignedInspec.Rows.Count > 0 Then
                        If Not ugAssignedInspec.ActiveRow Is Nothing Then
                            If ugAssignedInspec.ActiveRow.Selected Then
                                ugRow = ugAssignedInspec.ActiveRow
                                ugRow.Tag = ugAssignedInspec.Name
                            End If
                        End If
                    End If
            End Select
            Return ugRow
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function
    Public Function getSelectedRows() As Infragistics.Win.UltraWinGrid.SelectedRowsCollection
        Dim ugRows As Infragistics.Win.UltraWinGrid.SelectedRowsCollection
        Try
            Select Case tbCtrlInspection.SelectedTab.Name
                Case tbPageAllInfo.Name
                    If Me.ugAllFacsForOwner.Rows.Count > 0 Then
                        If ugAllFacsForOwner.Selected.Rows.Count > 0 Then
                            ugRows = ugAllFacsForOwner.Selected.Rows
                        End If
                    ElseIf ugMyFacs.Rows.Count > 0 Then
                        If ugMyFacs.Selected.Rows.Count > 0 Then
                            ugRows = ugMyFacs.Selected.Rows
                        End If
                    End If
                Case tbPageMyFacilities.Name
                    If ugMyFacilities.Rows.Count > 0 Then
                        If ugMyFacilities.Selected.Rows.Count > 0 Then
                            ugRows = ugMyFacilities.Selected.Rows
                        End If
                    End If
                Case tbPageAllFacForOwner.Name
                    If ugAllFacForOwner.Rows.Count > 0 Then
                        If ugAllFacForOwner.Selected.Rows.Count > 0 Then
                            ugRows = ugAllFacForOwner.Selected.Rows
                        End If
                    End If
            End Select
            Return ugRows
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function

    'Private Sub frmClosed(ByVal sender As Object, ByVal e As System.EventArgs)
    '    If sender.GetType.Name.IndexOf("Checklist") >= 0 Then
    '        frmChecklist = Nothing
    '    ElseIf sender.GetType.Name.IndexOf("InspectionHistory") >= 0 Then
    '        frmInspecHistory = Nothing
    '    ElseIf sender.GetType.Name.IndexOf("AssignedInspection") >= 0 Then
    '        frmAssignedInspec = Nothing
    '    ElseIf sender.GetType.Name.IndexOf("RescheduleInspection") >= 0 Then
    '        frmReschedule = Nothing
    '    End If
    'End Sub
#End Region
#Region "UI Events"
    ' ugMyFacs
    Private Sub ugMyFacs_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugMyFacs.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            Cursor.Current = Cursors.AppStarting
            bolLoading = True
            'ugMyFacs.ActiveRow.Selected = False
            nSelectedFacilityID = ugMyFacs.ActiveRow.Cells("FACILITY ID").Value
            PopulateFacsforOwner(nInspectorID, ugMyFacs.ActiveRow.Cells("OWNER_ID").Value, nSelectedFacilityID)
            ugMyFacs.ActiveRow = Nothing
            ugMyFacilities.ActiveRow = Nothing
            'If Not ugMyFacs.ActiveRow Is Nothing Then
            '    ugMyFacs.ActiveRow.Selected = False
            'End If
            'If ugMyFacs.Selected.Rows.Count > 0 Then
            '    For Each ugRow In ugMyFacs.Selected.Rows
            '        ugRow.Selected = False
            '    Next
            '    For Each ugRow In ugMyFacilities.Selected.Rows
            '        ugRow.Selected = False
            '    Next
            'End If
            'If Not (ugMyFacs.ActiveRow Is Nothing) Then
            '    ugMyFacs.ActiveRow.Selected = False
            '    ugMyFacs.ActiveRow = Nothing
            'End If
            'If Not (ugMyFacilities.ActiveRow Is Nothing) Then
            '    ugMyFacilities.ActiveRow.Selected = False
            '    ugMyFacilities.ActiveRow = Nothing
            'End If
            'If ugAllFacsForOwner.Rows.Count > 0 Then
            '    For Each ugRow In ugAllFacsForOwner.Rows
            '        If ugRow.Cells("FACILITY").Value = nSelectedFacilityID Then
            '            ugAllFacsForOwner.ActiveRow = ugRow
            '            ugAllFacsForOwner.Select()
            '            Exit For
            '        End If
            '    Next
            'End If
            'If ugAllFacForOwner.Rows.Count > 0 Then
            '    For Each ugRow In ugAllFacForOwner.Rows
            '        If ugRow.Cells("FACILITY").Value = nSelectedFacilityID Then
            '            ugAllFacForOwner.ActiveRow = ugRow
            '            ugAllFacForOwner.Select()
            '            Exit For
            '        End If
            '    Next
            'End If
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            bolLoading = False
        End Try
    End Sub
    Private Sub ugMyFacs_AfterSortChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BandEventArgs) Handles ugMyFacs.AfterSortChange
        If bolLoading Then Exit Sub
        Dim bolLoadingLocal As Boolean = bolLoading
        Try
            bolLoading = True
            rdBtnScheduled.Checked = False
            rdBtnDue.Checked = False
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            bolLoading = bolLoadingLocal
        End Try
        Try
            If ugMyFacs.Rows.Count > 0 Then
                If ugMyFacs.Selected.Rows.Count > 0 Then
                    'ugMyFacs.ActiveRow = ugMyFacs.Selected.Rows.Item(0)
                    ugMyFacs.Selected.Rows.Clear()
                    For Each ugRow In ugMyFacs.Selected.Rows
                        If ugRow.Cells("FACILITY ID").Value = nSelectedFacilityID Then
                            ugRow.Selected = True
                            ugMyFacs.ActiveRow = ugRow
                        End If
                    Next
                End If
            End If
            If ugMyFacilities.Rows.Count > 0 Then
                If ugMyFacilities.Selected.Rows.Count > 0 Then
                    'ugMyFacilities.ActiveRow = ugMyFacilities.Selected.Rows.Item(0)
                    ugMyFacilities.Selected.Rows.Clear()
                    For Each ugRow In ugMyFacilities.Selected.Rows
                        If ugRow.Cells("FACILITY ID").Value = nSelectedFacilityID Then
                            ugRow.Selected = True
                            ugMyFacilities.ActiveRow = ugRow
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugMyFacs_BeforeRowActivate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles ugMyFacs.BeforeRowActivate
        If bolLoading Then Exit Sub
        Try
            Cursor.Current = Cursors.AppStarting
            nSelectedFacilityID = e.Row.Cells("FACILITY ID").Value
            If nSelectedOwnerID <> e.Row.Cells("OWNER_ID").Value Then
                nSelectedOwnerID = e.Row.Cells("OWNER_ID").Value
                For Each ug As Infragistics.Win.UltraWinGrid.UltraGrid In arrugAllFacsforOwner
                    ug.DataSource = Nothing
                Next
            End If
            If ugAllFacsForOwner.Rows.Count > 0 Then
                If ugAllFacsForOwner.Selected.Rows.Count > 0 Then
                    For Each ugRow In ugAllFacsForOwner.Selected.Rows
                        ugRow.Selected = False
                    Next
                    For Each ugRow In ugAllFacForOwner.Selected.Rows
                        ugRow.Selected = False
                    Next
                End If
                If Not (ugAllFacsForOwner.ActiveRow Is Nothing Or ugAllFacForOwner.ActiveRow Is Nothing) Then
                    ugAllFacsForOwner.ActiveRow = Nothing
                    ugAllFacForOwner.ActiveRow = Nothing
                End If
            End If
            If (e.Row.Cells("SCHEDULED BY").Text.Trim.ToUpper <> MusterContainer.AppUser.Name.ToUpper.Trim And _
                e.Row.Cells("SCHEDULED BY").Text.Trim <> String.Empty) Or _
                e.Row.Cells("STATUS").Text.IndexOf("UPCOMING INSTALL") > -1 Then
                btnReschedule.Enabled = False
                btnReschedule.Text = "Schedule"
                btnEnterChecklistInfo.Enabled = True
                btnEnterChecklistInfo.Text = "View Checklist Info"
                btnGenerateLetters.Enabled = False
                btnPrintChecklists.Enabled = False
            ElseIf e.Row.Cells("SCHEDULED BY").Text.Trim.ToUpper <> MusterContainer.AppUser.Name.ToUpper.Trim And e.Row.Cells("SCHEDULED BY").Text.Trim = String.Empty Then
                btnReschedule.Enabled = True
                btnReschedule.Text = "Schedule"
                btnEnterChecklistInfo.Enabled = False
                btnEnterChecklistInfo.Text = "Enter Checklist Info"
                btnGenerateLetters.Enabled = False
                btnPrintChecklists.Enabled = False
            ElseIf e.Row.Cells("RESCHEDULED DATE").Value Is DBNull.Value And e.Row.Cells("SCHEDULED DATE").Value Is DBNull.Value Then
                btnReschedule.Enabled = True
                btnReschedule.Text = "Schedule"
                btnEnterChecklistInfo.Enabled = False
                btnEnterChecklistInfo.Text = "Enter Checklist Info"
                btnGenerateLetters.Enabled = False
                btnPrintChecklists.Enabled = False
                ' In My Facilities Tab
                'btnMyFacSchedule.Text = "Schedule"
                'btnMyFacEnterCheckInfo.Enabled = False
                'btnMyFacEnterCheckInfo.Text = "Enter Checklist Info"
            Else
                btnReschedule.Enabled = True
                btnReschedule.Text = "Reschedule"
                btnGenerateLetters.Enabled = True
                btnPrintChecklists.Enabled = True
                btnEnterChecklistInfo.Enabled = True
                ' In My Facilities Tab
                'btnMyFacSchedule.Text = "Reschedule"
                'btnMyFacEnterCheckInfo.Enabled = True
                If e.Row.Cells("SUBMITTED DATE").Value Is DBNull.Value Then
                    btnEnterChecklistInfo.Text = "Enter Checklist Info"
                    ' In My Facilities Tab
                    'btnMyFacEnterCheckInfo.Text = "Enter Checklist Info"
                Else
                    btnEnterChecklistInfo.Text = "View Checklist Info"
                    btnReschedule.Enabled = False
                    ' In My Facilities Tab
                    'btnMyFacEnterCheckInfo.Text = "View Checklist Info"
                End If
            End If
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    ' ugMyFacilities
    Private Sub ugMyFacilities_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugMyFacilities.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            Cursor.Current = Cursors.AppStarting
            bolLoading = True
            nSelectedFacilityID = ugMyFacilities.ActiveRow.Cells("FACILITY ID").Value
            PopulateFacsforOwner(nInspectorID, ugMyFacilities.ActiveRow.Cells("OWNER_ID").Value, nSelectedFacilityID)
            'If ugMyFacs.Selected.Rows.Count > 0 Then
            '    For Each ugRow In ugMyFacs.Selected.Rows
            '        ugRow.Selected = False
            '    Next
            '    For Each ugRow In ugMyFacilities.Selected.Rows
            '        ugRow.Selected = False
            '    Next
            'End If
            'If Not (ugMyFacs.ActiveRow Is Nothing) Then
            '    ugMyFacs.ActiveRow.Selected = False
            '    ugMyFacs.ActiveRow = Nothing
            'End If
            'If Not (ugMyFacilities.ActiveRow Is Nothing) Then
            '    ugMyFacilities.ActiveRow.Selected = False
            '    ugMyFacilities.ActiveRow = Nothing
            'End If
            'If ugAllFacsForOwner.Rows.Count > 0 Then
            '    For Each ugRow In ugAllFacsForOwner.Rows
            '        If ugRow.Cells("FACILITY").Value = nSelectedFacilityID Then
            '            ugAllFacsForOwner.ActiveRow = ugRow
            '            ugAllFacsForOwner.Select()
            '            Exit For
            '        End If
            '    Next
            'End If
            'If ugAllFacForOwner.Rows.Count > 0 Then
            '    For Each ugRow In ugAllFacForOwner.Rows
            '        If ugRow.Cells("FACILITY").Value = nSelectedFacilityID Then
            '            ugAllFacForOwner.ActiveRow = ugRow
            '            ugAllFacForOwner.Select()
            '            Exit For
            '        End If
            '    Next
            'End If
            tbCtrlInspection.SelectedTab = tbPageAllFacForOwner
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            Cursor.Current = Cursors.Default
            bolLoading = False
        End Try
    End Sub

    Private Sub ug_MouseDown(ByVal sender As Object, ByVal e As Infragistics.Win.UIElementEventArgs)


        ' find the cell that the cursor is over, if any
        If e.Element.GetType().Equals(GetType(Infragistics.Win.UltraWinGrid.CellUIElement)) Then


            Dim cell As Infragistics.Win.UltraWinGrid.UltraGridCell = e.Element.GetContext(GetType(Infragistics.Win.UltraWinGrid.UltraGridCell))
            If Not cell Is Nothing Then
                MouseDowned = cell.Row
            End If
        End If



    End Sub

    Private Sub ug_MouseDownClick(ByVal sender As Object, ByVal e As MouseEventArgs)
        Try

            If e.Button = MouseButtons.Right AndAlso Not MouseDowned Is Nothing AndAlso MouseDowned.Description <> String.Empty Then
                CreateToolTipMessage(MouseDowned, sender)
            End If

            MouseDowned = Nothing

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            Cursor.Current = Cursors.Default
            bolLoading = False
        End Try
    End Sub


    Private Sub ugMyFacilities_AfterSortChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BandEventArgs) Handles ugMyFacilities.AfterSortChange
        If bolLoading Then Exit Sub
        Dim bolLoadingLocal As Boolean = bolLoading
        Try
            bolLoading = True
            rdBtnScheduled.Checked = False
            rdBtnDue.Checked = False
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            bolLoading = bolLoadingLocal
        End Try
        Try
            If ugMyFacs.Rows.Count > 0 Then
                If ugMyFacs.Selected.Rows.Count > 0 Then
                    'ugMyFacs.ActiveRow = ugMyFacs.Selected.Rows.Item(0)
                    ugMyFacs.Selected.Rows.Clear()
                    For Each ugRow In ugMyFacs.Selected.Rows
                        If ugRow.Cells("FACILITY ID").Value = nSelectedFacilityID Then
                            ugRow.Selected = True
                            ugMyFacs.ActiveRow = ugRow
                        End If
                    Next
                End If
            End If
            If ugMyFacilities.Rows.Count > 0 Then
                If ugMyFacilities.Selected.Rows.Count > 0 Then
                    'ugMyFacilities.ActiveRow = ugMyFacilities.Selected.Rows.Item(0)
                    ugMyFacilities.Selected.Rows.Clear()
                    For Each ugRow In ugMyFacilities.Selected.Rows
                        If ugRow.Cells("FACILITY ID").Value = nSelectedFacilityID Then
                            ugRow.Selected = True
                            ugMyFacilities.ActiveRow = ugRow
                        End If
                    Next
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugMyFacilities_BeforeRowActivate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles ugMyFacilities.BeforeRowActivate
        If bolLoading Then Exit Sub
        Try
            Cursor.Current = Cursors.AppStarting
            nSelectedFacilityID = e.Row.Cells("FACILITY ID").Value
            If nSelectedOwnerID <> e.Row.Cells("OWNER_ID").Value Then
                nSelectedOwnerID = e.Row.Cells("OWNER_ID").Value
                For Each ug As Infragistics.Win.UltraWinGrid.UltraGrid In arrugAllFacsforOwner
                    ug.DataSource = Nothing
                Next
            End If
            If ugAllFacsForOwner.Rows.Count > 0 Then
                If ugAllFacsForOwner.Selected.Rows.Count > 0 Then
                    For Each ugRow In ugAllFacsForOwner.Selected.Rows
                        ugRow.Selected = False
                    Next
                    For Each ugRow In ugAllFacForOwner.Selected.Rows
                        ugRow.Selected = False
                    Next
                End If
                If Not (ugAllFacsForOwner.ActiveRow Is Nothing) Then
                    ugAllFacsForOwner.ActiveRow = Nothing
                    ugAllFacForOwner.ActiveRow = Nothing
                End If
            End If
            If (e.Row.Cells("SCHEDULED BY").Text.Trim.ToUpper <> MusterContainer.AppUser.Name.ToUpper.Trim And _
                e.Row.Cells("SCHEDULED BY").Text.Trim <> String.Empty) Or _
                e.Row.Cells("STATUS").Text.IndexOf("UPCOMING INSTALL") > -1 Then
                btnMyFacSchedule.Enabled = False
                btnMyFacSchedule.Text = "Schedule"
                btnMyFacEnterCheckInfo.Enabled = True
                btnMyFacEnterCheckInfo.Text = "View Checklist Info"
                btnMyFacGenLetters.Enabled = False
                btnMyFacPrintChecklists.Enabled = False
            ElseIf e.Row.Cells("SCHEDULED BY").Text.Trim.ToUpper <> MusterContainer.AppUser.Name.ToUpper.Trim And e.Row.Cells("SCHEDULED BY").Text.Trim = String.Empty Then
                btnMyFacSchedule.Enabled = True
                btnMyFacSchedule.Text = "Schedule"
                btnMyFacEnterCheckInfo.Enabled = False
                btnMyFacEnterCheckInfo.Text = "Enter Checklist Info"
                btnMyFacGenLetters.Enabled = False
                btnMyFacPrintChecklists.Enabled = False
            ElseIf e.Row.Cells("RESCHEDULED DATE").Value Is DBNull.Value And e.Row.Cells("SCHEDULED DATE").Value Is DBNull.Value Then
                btnMyFacSchedule.Enabled = True
                btnMyFacSchedule.Text = "Schedule"
                btnMyFacEnterCheckInfo.Enabled = False
                btnMyFacEnterCheckInfo.Text = "Enter Checklist Info"
                btnMyFacGenLetters.Enabled = False
                btnMyFacPrintChecklists.Enabled = False
                ' In all Info Tab
                'btnReschedule.Text = "Schedule"
                'btnEnterChecklistInfo.Enabled = False
                'btnEnterChecklistInfo.Text = "Enter Checklist Info"
            Else
                btnMyFacSchedule.Enabled = True
                btnMyFacSchedule.Text = "Reschedule"
                btnMyFacGenLetters.Enabled = True
                btnMyFacPrintChecklists.Enabled = True
                btnMyFacEnterCheckInfo.Enabled = True
                ' In all Info Tab
                'btnReschedule.Text = "Reschedule"
                'btnEnterChecklistInfo.Enabled = True
                If e.Row.Cells("SUBMITTED DATE").Value Is DBNull.Value Then
                    btnMyFacEnterCheckInfo.Text = "Enter Checklist Info"
                    ' In all Info Tab
                    'btnEnterChecklistInfo.Text = "Enter Checklist Info"
                Else
                    btnMyFacEnterCheckInfo.Text = "View Checklist Info"
                    btnMyFacSchedule.Enabled = False
                    ' In all Info Tab
                    'btnEnterChecklistInfo.Text = "View Checklist Info"
                End If
            End If
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    ' ugAllFacsForOwner
    Private Sub ugAllFacsForOwner_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugAllFacsForOwner.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            Cursor.Current = Cursors.AppStarting
            EnterCheckList()
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    'Private Sub ugAllFacsForOwner_AfterSortChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BandEventArgs) Handles ugAllFacsForOwner.AfterSortChange
    '    Try
    '        If ugAllFacsForOwner.Rows.Count > 0 Then
    '            If ugAllFacsForOwner.Selected.Rows.Count > 0 Then
    '                ugAllFacsForOwner.ActiveRow = ugAllFacsForOwner.Selected.Rows.Item(0)
    '                For Each ugRow In ugAllFacsForOwner.Selected.Rows
    '                    ugRow.Selected = False
    '                Next
    '            End If
    '        End If
    '        If ugAllFacForOwner.Rows.Count > 0 Then
    '            If ugAllFacForOwner.Selected.Rows.Count > 0 Then
    '                ugAllFacForOwner.ActiveRow = ugAllFacForOwner.Selected.Rows.Item(0)
    '                For Each ugRow In ugAllFacForOwner.Selected.Rows
    '                    ugRow.Selected = False
    '                Next
    '            End If
    '        End If
    '        If ugAllFacsForOwner.Rows.Count > 0 Then
    '            ugAllFacsForOwner.ActiveRow = ugAllFacsForOwner.Rows(0)
    '        End If
    '        If ugAllFacForOwner.Rows.Count > 0 Then
    '            ugAllFacForOwner.ActiveRow = ugAllFacForOwner.Rows(0)
    '        End If
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    Private Sub ugAllFacsForOwner_BeforeRowActivate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles ugAllFacsForOwner.BeforeRowActivate
        If bolLoading Then Exit Sub
        Try
            Cursor.Current = Cursors.AppStarting

            If TypeOf e.Row.Cells("FACILITY ID").Value Is String AndAlso IsNumeric(e.Row.Cells("FACILITY ID").Value) Then
                nSelectedFacilityID = Convert.ToInt32(e.Row.Cells("FACILITY ID").Value)

            ElseIf TypeOf e.Row.Cells("FACILITY ID").Value Is Integer Then

                nSelectedFacilityID = e.Row.Cells("FACILITY ID").Value

            End If

            If ugMyFacs.Rows.Count > 0 Then
                If ugMyFacs.Selected.Rows.Count > 0 Then
                    For Each ugRow In ugMyFacs.Selected.Rows
                        ugRow.Selected = False
                    Next
                    For Each ugRow In ugMyFacilities.Selected.Rows
                        ugRow.Selected = False
                    Next
                End If
                If Not (ugMyFacs.ActiveRow Is Nothing) Then
                    ugMyFacs.ActiveRow = Nothing
                    ugMyFacilities.ActiveRow = Nothing
                End If
            End If
            If (e.Row.Cells("SCHEDULED BY").Text.Trim.ToUpper <> MusterContainer.AppUser.Name.ToUpper.Trim And _
                e.Row.Cells("SCHEDULED BY").Text.Trim <> String.Empty) Or _
                e.Row.Cells("STATUS").Text.IndexOf("UPCOMING INSTALL") > -1 Then
                btnReschedule.Enabled = False
                btnReschedule.Text = "Schedule"
                btnEnterChecklistInfo.Enabled = True
                btnEnterChecklistInfo.Text = "View Checklist Info"
                btnGenerateLetters.Enabled = False
                btnPrintChecklists.Enabled = False
            ElseIf e.Row.Cells("SCHEDULED BY").Text.Trim.ToUpper <> MusterContainer.AppUser.Name.ToUpper.Trim And e.Row.Cells("SCHEDULED BY").Text.Trim = String.Empty Then
                btnReschedule.Enabled = True
                btnReschedule.Text = "Schedule"
                btnEnterChecklistInfo.Enabled = False
                btnEnterChecklistInfo.Text = "Enter Checklist Info"
                btnGenerateLetters.Enabled = False
                btnPrintChecklists.Enabled = False
            ElseIf e.Row.Cells("RESCHEDULED DATE").Value Is DBNull.Value And e.Row.Cells("SCHEDULED DATE").Value Is DBNull.Value Then
                btnReschedule.Enabled = True
                btnReschedule.Text = "Schedule"
                btnEnterChecklistInfo.Enabled = False
                btnEnterChecklistInfo.Text = "Enter Checklist Info"
                btnGenerateLetters.Enabled = False
                btnPrintChecklists.Enabled = False
                ' In All Facs for Selected Owner Tab
                'btnAllFacSchedule.Text = "Schedule"
                'btnAllFacEntercheckInfo.Enabled = False
                'btnAllFacEntercheckInfo.Text = "Enter Checklist Info"
            Else
                btnReschedule.Enabled = True
                btnReschedule.Text = "Reschedule"
                btnGenerateLetters.Enabled = True
                btnPrintChecklists.Enabled = True
                btnEnterChecklistInfo.Enabled = True
                ' In All Facs for Selected Owner Tab
                'btnAllFacSchedule.Text = "Reschedule"
                'btnAllFacEntercheckInfo.Enabled = True
                If e.Row.Cells("SUBMITTED DATE").Value Is DBNull.Value Then
                    btnEnterChecklistInfo.Text = "Enter Checklist Info"
                    ' In All Facs for Selected Owner Tab
                    'btnAllFacEntercheckInfo.Text = "Enter Checklist Info"
                Else
                    btnEnterChecklistInfo.Text = "View Checklist Info"
                    btnReschedule.Enabled = False
                    ' In All Facs for Selected Owner Tab
                    'btnAllFacEntercheckInfo.Text = "View Checklist Info"
                End If
            End If
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    ' ugAllFacForOwner
    Private Sub ugAllFacForOwner_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugAllFacForOwner.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            Cursor.Current = Cursors.AppStarting
            EnterCheckList()
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    'Private Sub ugAllFacForOwner_AfterSortChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BandEventArgs) Handles ugAllFacForOwner.AfterSortChange
    '    Try
    '        If ugAllFacForOwner.Rows.Count > 0 Then
    '            ugAllFacForOwner.ActiveRow = ugAllFacForOwner.Rows(0)
    '        End If
    '        If ugAllFacsForOwner.Rows.Count > 0 Then
    '            ugAllFacsForOwner.ActiveRow = ugAllFacsForOwner.Rows(0)
    '        End If
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    Private Sub ugAllFacForOwner_BeforeRowActivate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles ugAllFacForOwner.BeforeRowActivate
        If bolLoading Then Exit Sub
        Try
            Cursor.Current = Cursors.AppStarting
            nSelectedFacilityID = e.Row.Cells("FACILITY ID").Value
            If ugMyFacs.Rows.Count > 0 Then
                If ugMyFacs.Selected.Rows.Count > 0 Then
                    For Each ugRow In ugMyFacs.Selected.Rows
                        ugRow.Selected = False
                    Next
                    For Each ugRow In ugMyFacilities.Selected.Rows
                        ugRow.Selected = False
                    Next
                End If
                If Not (ugMyFacs.ActiveRow Is Nothing) Then
                    ugMyFacs.ActiveRow = Nothing
                    ugMyFacilities.ActiveRow = Nothing
                End If
            End If
            If (e.Row.Cells("SCHEDULED BY").Text.Trim.ToUpper <> MusterContainer.AppUser.Name.ToUpper.Trim And _
                e.Row.Cells("SCHEDULED BY").Text.Trim <> String.Empty) Or _
                e.Row.Cells("STATUS").Text.IndexOf("UPCOMING INSTALL") > -1 Then
                btnAllFacSchedule.Enabled = False
                btnAllFacSchedule.Text = "Schedule"
                btnAllFacEntercheckInfo.Enabled = True
                btnAllFacEntercheckInfo.Text = "View Checklist Info"
                btnAllFacGenLetters.Enabled = False
                btnAllFacPrintChecklist.Enabled = False
            ElseIf e.Row.Cells("SCHEDULED BY").Text.Trim.ToUpper <> MusterContainer.AppUser.Name.ToUpper.Trim And e.Row.Cells("SCHEDULED BY").Text.Trim = String.Empty Then
                btnAllFacSchedule.Enabled = True
                btnAllFacSchedule.Text = "Schedule"
                btnAllFacEntercheckInfo.Enabled = False
                btnAllFacEntercheckInfo.Text = "Enter Checklist Info"
                btnAllFacGenLetters.Enabled = False
                btnAllFacPrintChecklist.Enabled = False
            ElseIf e.Row.Cells("RESCHEDULED DATE").Value Is DBNull.Value And e.Row.Cells("SCHEDULED DATE").Value Is DBNull.Value Then
                btnAllFacSchedule.Enabled = True
                btnAllFacSchedule.Text = "Schedule"
                btnAllFacEntercheckInfo.Enabled = False
                btnAllFacEntercheckInfo.Text = "Enter Checklist Info"
                btnAllFacGenLetters.Enabled = False
                btnAllFacPrintChecklist.Enabled = False
                ' In All Info Tab
                'btnReschedule.Text = "Schedule"
                'btnEnterChecklistInfo.Enabled = False
                'btnEnterChecklistInfo.Text = "Enter Checklist Info"
            Else
                btnAllFacSchedule.Enabled = True
                btnAllFacSchedule.Text = "Reschedule"
                btnAllFacGenLetters.Enabled = True
                btnAllFacPrintChecklist.Enabled = True
                btnAllFacEntercheckInfo.Enabled = True
                ' In All Info Tab
                'btnReschedule.Text = "Reschedule"
                'btnEnterChecklistInfo.Enabled = True
                If e.Row.Cells("SUBMITTED DATE").Value Is DBNull.Value Then
                    btnAllFacEntercheckInfo.Text = "Enter Checklist Info"
                    ' In All Info Tab
                    'btnEnterChecklistInfo.Text = "Enter Checklist Info"
                Else
                    btnAllFacEntercheckInfo.Text = "View Checklist Info"
                    btnAllFacSchedule.Enabled = False
                    ' In All Info Tab
                    'btnEnterChecklistInfo.Text = "View Checklist Info"
                End If
            End If
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub tbCtrlInspection_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbCtrlInspection.SelectedIndexChanged
        Try
            If nSelectedFacilityID > 0 Then
                Select Case tbCtrlInspection.SelectedTab.Name
                    Case tbPageAllInfo.Name
                        If ugMyFacs.Rows.Count > 0 Then
                            ugMyFacs.Selected.Rows.Clear()
                            For Each ugRow In ugMyFacs.Rows
                                If ugRow.Cells("FACILITY ID").Value = nSelectedFacilityID Then
                                    ugMyFacs.ActiveRow = ugRow
                                    ugMyFacs.ActiveRow.Selected = True
                                    ugMyFacs.ActiveRow.Activate()
                                    Exit For
                                End If
                            Next
                        End If
                        If ugAllFacsForOwner.Rows.Count > 0 Then
                            ugAllFacsForOwner.Selected.Rows.Clear()
                            For Each ugRow In ugAllFacsForOwner.Rows
                                If ugRow.Cells("FACILITY ID").Value = nSelectedFacilityID Then
                                    ugAllFacsForOwner.ActiveRow = ugRow
                                    ugAllFacsForOwner.ActiveRow.Selected = True
                                    'ugAllFacsForOwner.Select()
                                    ugAllFacsForOwner.ActiveRow.Activate()
                                    Exit For
                                End If
                            Next
                        End If
                    Case tbPageMyFacilities.Name
                        If ugMyFacilities.Rows.Count > 0 Then
                            ugMyFacilities.Selected.Rows.Clear()
                            For Each ugRow In ugMyFacilities.Rows
                                If ugRow.Cells("FACILITY ID").Value = nSelectedFacilityID Then
                                    ugMyFacilities.ActiveRow = ugRow
                                    ugMyFacilities.ActiveRow.Selected = True
                                    ugMyFacilities.ActiveRow.Activate()
                                    Exit For
                                End If
                            Next
                        End If
                    Case tbPageAllFacForOwner.Name
                        If ugAllFacForOwner.Rows.Count > 0 Then
                            ugAllFacForOwner.Selected.Rows.Clear()
                            For Each ugRow In ugAllFacForOwner.Rows
                                If ugRow.Cells("FACILITY ID").Value = nSelectedFacilityID Then
                                    ugAllFacForOwner.ActiveRow = ugRow
                                    ugAllFacForOwner.ActiveRow.Selected = True
                                    'ugAllFacForOwner.Select()
                                    ugAllFacForOwner.ActiveRow.Activate()
                                    Exit For
                                End If
                            Next
                        End If
                End Select
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub rdBtnScheduled_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdBtnScheduled.CheckedChanged
        If bolLoading Then Exit Sub
        Dim bolLoadingLocal As Boolean = bolLoading
        Dim ug As Infragistics.Win.UltraWinGrid.UltraGrid
        Try
            bolLoading = True
            Cursor.Current = Cursors.AppStarting
            If rdBtnScheduled.Checked Then
                For Each ug In arrugMyFacs
                    If ug.Rows.Count > 0 Then
                        ChangeTargetFacilitiesLayout(ug)
                        'ChangeTargetFacilitiesLayout(4, 5, 6, 7, 8, ug)

                        ug.DisplayLayout.Bands(0).SortedColumns.Clear()
                        ug.DisplayLayout.Bands(0).Columns("SCHEDULED DATE FOR SORT").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                        ug.DisplayLayout.Bands(0).Columns("SCHEDULED DATE").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                        ug.DisplayLayout.Bands(0).Columns("SCHEDULED TIME FOR SORT").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                        ug.DisplayLayout.Bands(0).Columns("SCHEDULED TIME").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                        ug.DisplayLayout.Bands(0).Columns("OWNER NAME").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                        ug.DisplayLayout.Bands(0).Columns("FACILITY").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                    End If
                Next
                'For Each ug In arrugAllFacsforOwner
                '    If ug.Rows.Count > 0 Then
                '        ChangeTargetFacilitiesLayout(4, 5, 6, 7, 8, ug)

                '        ug.DisplayLayout.Bands(0).SortedColumns.Clear()
                '        ug.DisplayLayout.Bands(0).Columns("SCHEDULED DATE").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                '        ug.DisplayLayout.Bands(0).Columns("SCHEDULED TIME").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                '        ug.DisplayLayout.Bands(0).Columns("OWNER NAME").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                '        ug.DisplayLayout.Bands(0).Columns("FACILITY").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                '    End If
                'Next

                'Select Case tbCtrlInspection.SelectedTab.Name
                '    Case tbPageAllInfo.Name
                '        ug1 = ugMyFacs
                '        ug2 = ugMyFacilities
                '        If ugAllFacsForOwner.Rows.Count > 0 Then
                '            ChangeTargetFacilitiesLayout(4, 5, 6, 7, 8, ugAllFacsForOwner)
                '            ChangeTargetFacilitiesLayout(4, 5, 6, 7, 8, ugAllFacForOwner)
                '        End If
                '    Case tbPageMyFacilities.Name
                '        ug1 = ugMyFacilities
                '        ug2 = ugMyFacs
                'End Select

                'If Not ug1 Is Nothing Then
                '    If ug1.Rows.Count > 0 Then
                '        ChangeTargetFacilitiesLayout(4, 5, 6, 7, 8, ug1)

                '        ug1.DisplayLayout.Bands(0).SortedColumns.Clear()
                '        ug1.DisplayLayout.Bands(0).Columns("SCHEDULED DATE").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Descending
                '        ug1.DisplayLayout.Bands(0).Columns("SCHEDULED TIME").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Descending
                '        ug1.DisplayLayout.Bands(0).Columns("OWNER NAME").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Descending
                '        ug1.DisplayLayout.Bands(0).Columns("FACILITY").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Descending
                '    End If
                'End If
                'If Not ug2 Is Nothing Then
                '    If ug2.Rows.Count > 0 Then
                '        ChangeTargetFacilitiesLayout(4, 5, 6, 7, 8, ug2)

                '        ug2.DisplayLayout.Bands(0).SortedColumns.Clear()
                '        ug2.DisplayLayout.Bands(0).Columns("SCHEDULED DATE").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Descending
                '        ug2.DisplayLayout.Bands(0).Columns("SCHEDULED TIME").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Descending
                '        ug2.DisplayLayout.Bands(0).Columns("OWNER NAME").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Descending
                '        ug2.DisplayLayout.Bands(0).Columns("FACILITY").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Descending
                '    End If
                'End If
            End If
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            bolLoading = bolLoadingLocal
        End Try
    End Sub
    Private Sub rdBtnDue_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdBtnDue.CheckedChanged
        If bolLoading Then Exit Sub
        Dim bolLoadingLocal As Boolean = bolLoading
        Dim ug As Infragistics.Win.UltraWinGrid.UltraGrid
        Try
            bolLoading = True
            Cursor.Current = Cursors.AppStarting
            If rdBtnDue.Checked Then
                For Each ug In arrugMyFacs
                    If ug.Rows.Count > 0 Then
                        ChangeTargetFacilitiesLayout(ug)
                        'ChangeTargetFacilitiesLayout(7, 8, 4, 5, 6, ug)

                        ug.DisplayLayout.Bands(0).SortedColumns.Clear()
                        ug.DisplayLayout.Bands(0).Columns("STATUS FOR SORT").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                        'ug.DisplayLayout.Bands(0).Columns("NEW FOR SORT").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                        'ug.DisplayLayout.Bands(0).Columns("LAST INSPECTED ON FOR SORT").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                        'ug.DisplayLayout.Bands(0).Columns("LAST OWNER INSPECTION DATE FOR SORT").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                        ug.DisplayLayout.Bands(0).Columns("LAST INSPECTED ON").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                        ug.DisplayLayout.Bands(0).Columns("LAST OWNER INSPECTION DATE").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                        ug.DisplayLayout.Bands(0).Columns("COUNTY").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                        ug.DisplayLayout.Bands(0).Columns("CITY").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                    End If
                Next
                'For Each ug In arrugAllFacsforOwner
                '    If ug.Rows.Count > 0 Then
                '        ChangeTargetFacilitiesLayout(7, 8, 4, 5, 6, ug)

                '        ug.DisplayLayout.Bands(0).SortedColumns.Clear()
                '        ug.DisplayLayout.Bands(0).Columns("LAST OWNER INSPECTION DATE").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                '        ug.DisplayLayout.Bands(0).Columns("LAST INSPECTED ON").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                '        ug.DisplayLayout.Bands(0).Columns("COUNTY").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                '        ug.DisplayLayout.Bands(0).Columns("CITY").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                '    End If
                'Next
                'Select Case tbCtrlInspection.SelectedTab.Name
                '    Case tbPageAllInfo.Name
                '        ug1 = ugMyFacs
                '        ug2 = ugMyFacilities
                '        If ugAllFacsForOwner.Rows.Count > 0 Then
                '            ChangeTargetFacilitiesLayout(7, 8, 5, 5, 6, ugAllFacsForOwner)
                '            ChangeTargetFacilitiesLayout(4, 5, 6, 7, 8, ugAllFacForOwner)
                '        End If
                '    Case tbPageMyFacilities.Name
                '        ug1 = ugMyFacilities
                '        ug2 = ugMyFacs
                'End Select
                '
                'If Not ug1 Is Nothing Then
                '    If ug1.Rows.Count > 0 Then
                '        ChangeTargetFacilitiesLayout(7, 8, 4, 5, 6, ug1)

                '        ug1.DisplayLayout.Bands(0).SortedColumns.Clear()
                '        ug1.DisplayLayout.Bands(0).Columns("LAST OWNER INSPECTION DATE").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                '        ug1.DisplayLayout.Bands(0).Columns("LAST INSPECTED ON").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                '        ug1.DisplayLayout.Bands(0).Columns("COUNTY").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                '        ug1.DisplayLayout.Bands(0).Columns("CITY").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                '    End If
                'End If
                '
                'If Not ug2 Is Nothing Then
                '    If ug2.Rows.Count > 0 Then
                '        Me.ChangeTargetFacilitiesLayout(7, 8, 4, 5, 6, ug2)

                '        ug2.DisplayLayout.Bands(0).SortedColumns.Clear()
                '        ug2.DisplayLayout.Bands(0).Columns("LAST OWNER INSPECTION DATE").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                '        ug2.DisplayLayout.Bands(0).Columns("LAST INSPECTED ON").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                '        ug2.DisplayLayout.Bands(0).Columns("COUNTY").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                '        ug2.DisplayLayout.Bands(0).Columns("CITY").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                '    End If
                'End If
            End If
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            bolLoading = bolLoadingLocal
        End Try
    End Sub
    Private Sub btnRestoreSortOrder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRestoreSortOrder.Click
        Dim ug As Infragistics.Win.UltraWinGrid.UltraGrid
        Try
            Cursor.Current = Cursors.AppStarting

            For Each ug In arrugAllFacsforOwner
                If ug.Rows.Count > 0 Then
                    ug.DisplayLayout.Bands(0).SortedColumns.Clear()
                    ChangeTargetFacilitiesLayout(ug)
                    'If rdBtnDue.Checked Then
                    '    ChangeTargetFacilitiesLayout(7, 8, 4, 5, 6, ug)
                    'ElseIf rdBtnScheduled.Checked Then
                    '    ChangeTargetFacilitiesLayout(4, 5, 6, 7, 8, ug)
                    'End If
                    ug.DisplayLayout.Bands(0).Columns("SCHEDULED BY FOR SORT").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                    ug.DisplayLayout.Bands(0).Columns("STATUS FOR SORT").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                    ug.DisplayLayout.Bands(0).Columns("LAST INSPECTED ON").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                    ug.DisplayLayout.Bands(0).Columns("LAST OWNER INSPECTION DATE").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                    ug.DisplayLayout.Bands(0).Columns("COUNTY").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                    ug.DisplayLayout.Bands(0).Columns("CITY").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending

                    'ug.DisplayLayout.Bands(0).Columns("SCHEDULED BY FOR SORT").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                    'ug.DisplayLayout.Bands(0).Columns("STAFF_ID").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Descending
                    'ug.DisplayLayout.Bands(0).Columns("SCHEDULED DATE FOR SORT").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                    'ug.DisplayLayout.Bands(0).Columns("SCHEDULED DATE").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                    'ug.DisplayLayout.Bands(0).Columns("SCHEDULED TIME FOR SORT").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                    'ug.DisplayLayout.Bands(0).Columns("SCHEDULED TIME").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                End If
            Next

            'Select Case tbCtrlInspection.SelectedTab.Name
            '    Case tbPageAllInfo.Name
            '        ug1 = ugAllFacsForOwner
            '        ug2 = ugAllFacForOwner
            '    Case tbPageAllFacForOwner.Name
            '        ug1 = ugAllFacForOwner
            '        ug2 = ugAllFacsForOwner
            '    Case Else
            '        Exit Sub
            'End Select
            '
            'If ug1.Rows.Count > 0 Then
            '    ug1.DisplayLayout.Bands(0).SortedColumns.Clear()
            '    'ug1.DisplayLayout.Bands(0).Columns("SCHEDULED BY").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Descending
            '    ug1.DisplayLayout.Bands(0).Columns("STAFF_ID").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Descending
            '    ug1.DisplayLayout.Bands(0).Columns("SCHEDULED DATE").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            '    ug1.DisplayLayout.Bands(0).Columns("SCHEDULED TIME").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            '
            '    ug2.DisplayLayout.Bands(0).SortedColumns.Clear()
            '    'ug2.DisplayLayout.Bands(0).Columns("SCHEDULED BY").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Descending
            '    ug2.DisplayLayout.Bands(0).Columns("STAFF_ID").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Descending
            '    ug2.DisplayLayout.Bands(0).Columns("SCHEDULED DATE").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            '    ug2.DisplayLayout.Bands(0).Columns("SCHEDULED TIME").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            'End If

            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnRestoreSortOrder2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRestoreSortOrder2.Click
        Try
            btnRestoreSortOrder_Click(sender, e)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub chkShowAllTargetFacilities_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkShowAllTargetFacilities.CheckedChanged
        If bolLoading Then Exit Sub
        Try
            Cursor.Current = Cursors.AppStarting
            PopulateMyFacilities(nInspectorID, 0, 0, chkShowAllTargetFacilities.Checked)
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub cmbInspectors_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbInspectors.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Dim bolLoadingLocal As Boolean = bolLoading
        Try
            bolLoading = True
            chkShowAllTargetFacilities.Checked = False
            nInspectorID = UIUtilsGen.GetComboBoxValue(cmbInspectors)
            Populate(nInspectorID)
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            bolLoading = bolLoadingLocal
        End Try
    End Sub
#End Region
#Region "Checklist"
    Public Sub EnterCheckList()
        Dim bolReadOnly As Boolean = False
        Try
            ugRow = getSelectedRow()
            If ugRow Is Nothing Then
                MsgBox("Please select a facilty")
                Exit Sub
            End If
            nSelectedFacilityID = ugRow.Cells("FACILITY ID").Value
            'ugrow.Cells("SCHEDULED DATE").Value Is System.DBNull.Value

            If ugRow.Cells("INSPECTION_ID").Value Is System.DBNull.Value Then
                MsgBox("Please schedule inspection before entering checklist")
                Exit Sub
            Else
                'oInspection.Retrieve(CType(ugRow.Cells("INSPECTION_ID").Value, Int64), , CType(ugRow.Cells("FACILITY ID").Value, Int64), CType(ugRow.Cells("OWNER_ID").Value, Int64))
                oInspection.Retrieve(CType(ugRow.Cells("INSPECTION_ID").Value, Int64), , CType(ugRow.Cells("FACILITY ID").Value, Int64))
            End If

            Dim id As Integer = ugRow.Cells("INSPECTION_ID").Value

            bolReadOnly = IIf(ugRow.Cells("SUBMITTED DATE").Value Is DBNull.Value, False, True)

            Dim newowner As Integer = oInspection.OwnerID

            If ugRow.Cells("SCHEDULED BY").Text.Trim.ToUpper <> MusterContainer.AppUser.Name.ToUpper.Trim And ugRow.Cells("SCHEDULED BY").Text.Trim <> String.Empty Then
                bolReadOnly = True
            Else
                ' #2897 if the current owner is different for the inspection, user cannot enter checklist to modify

                If oInspection.OwnerID <> ugRow.Cells("OWNER_ID").Value Then

                    newowner = ugRow.Cells("OWNER_ID").Value

                    If MsgBox("This facility has been recently transferred to a new owner, would you like to transfer the Inspection as well.", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                        oInspection.OwnerID = newowner
                        oInspection.ModifiedBy = MusterContainer.AppUser.ID
                        oInspection.Save(CType(UIUtilsGen.ModuleID.Inspection, Integer), MusterContainer.AppUser.UserKey, returnVal, True)
                    Else
                        newowner = oInspection.OwnerID
                    End If

                End If

            End If

            oInspection.Clear()
            oInspection.Retrieve(id, , , newowner)
            frmChecklist = New CheckList(oInspection, bolReadOnly, )
            frmChecklist.CallingForm = Me
            Me.Tag = "0"
            frmChecklist.WindowState = FormWindowState.Maximized
            'frmChecklist.TopMost = True
            'AddHandler frmChecklist.Closing, AddressOf frmClosing
            'AddHandler frmChecklist.Closed, AddressOf frmClosed
            frmChecklist.ShowDialog()
            If Me.Tag = "1" Then
                oInspection.CheckListMaster.Owner.Remove(oInspection.CheckListMaster.Owner.ID)
                oInspection.CheckListMaster = New MUSTER.BusinessLogic.pInspectionChecklistMaster
                oInspection.Remove(oInspection.ID)
                oInspection = New MUSTER.BusinessLogic.pInspection
                If ugAllFacsForOwner.Rows.Count > 0 Then
                    PopulateMyFacilities(nInspectorID, ugRow.Cells("OWNER_ID").Value, nSelectedFacilityID, chkShowAllTargetFacilities.Checked, , ugRow.Cells("INSPECTION_ID").Value, False)
                Else
                    PopulateMyFacilities(nInspectorID, , nSelectedFacilityID, , , ugRow.Cells("INSPECTION_ID").Value, False)
                End If
            End If
            ' #2960
            'oInspection.Remove(oInspection.ID)
            oInspection = New MUSTER.BusinessLogic.pInspection
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            If Not frmChecklist Is Nothing Then
                frmChecklist = Nothing
            End If
        End Try
    End Sub
#Region "Events"
    Private Sub btnEnterChecklistInfo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEnterChecklistInfo.Click
        Try
            Me.EnterCheckList()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnMyFacEnterCheckInfo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMyFacEnterCheckInfo.Click
        Try
            Me.EnterCheckList()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnAllFacEntercheckInfo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAllFacEntercheckInfo.Click
        Try
            Me.EnterCheckList()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#End Region
#Region "Inspection History"
    Private Sub btnViewHistory_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnViewHistory.Click
        Dim dsViewHistory As DataSet
        Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim drRow As DataRow
        Try
            ugRow = Me.getSelectedRow()
            If ugRow Is Nothing Then
                MsgBox("Please select a facility")
                Exit Sub
            End If
            dsViewHistory = oInspection.RetrieveInspectionHistory(Integer.Parse(ugRow.Cells("FACILITY ID").Value))
            If dsViewHistory.Tables(0).Rows.Count = 0 Then
                MsgBox("Facility has no Inspection history")
                Exit Sub
            Else
                frmInspecHistory = New InspectionHistory(oInspection, dsViewHistory, ugRow.Cells("FACILITY").Value)
                frmInspecHistory.ShowDialog()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnMyFacViewHistory_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMyFacViewHistory.Click
        btnViewHistory_Click(sender, e)
    End Sub
    Private Sub btnAllFacViewHistory_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAllFacViewHistory.Click
        btnViewHistory_Click(sender, e)
    End Sub
#End Region
#Region "Assigned Inspections"
    Private Function getSelectedAssignedFacility() As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            ugRow = Nothing
            Select Case tbCtrlInspection.SelectedTab.Name
                Case tbPageAllInfo.Name
                    If Me.ugAssignedInspections.Rows.Count > 0 Then
                        If ugAssignedInspections.ActiveRow.Selected Then
                            ugRow = ugAssignedInspections.ActiveRow
                        End If
                    ElseIf ugMyFacs.Rows.Count > 0 Then
                        If ugMyFacs.ActiveRow.Selected Then
                            ugRow = ugMyFacs.ActiveRow
                        End If
                    End If
                Case tbPageAssignedInspec.Name
                    If ugAssignedInspec.Rows.Count > 0 Then
                        If ugAssignedInspec.ActiveRow.Selected Then
                            ugRow = ugAssignedInspec.ActiveRow
                        End If
                    End If
            End Select
            Return ugRow
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function
    Private Sub EnterAssignedInspection(Optional ByVal frmShow As Boolean = True)
        Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            ugRow = getSelectedAssignedFacility()
            If ugRow Is Nothing Then
                MsgBox("Please select an Assigned Facility")
                Exit Sub
            End If
            If ugRow.Cells("INSPECTION_ID").Value Is System.DBNull.Value Then
                MsgBox("Inspection ID cannot be NULL")
                Exit Sub
            Else
                oInspection.Retrieve(CType(ugRow.Cells("INSPECTION_ID").Value, Int64))
            End If
            Dim bolReadOnly As Boolean = False
            If ugRow.Cells("INSPECTOR").Text.Trim.ToUpper <> MusterContainer.AppUser.Name.ToUpper.Trim Then
                bolReadOnly = True
            End If
            frmAssignedInspec = New AssignedInspection(ugRow, oInspection, bolReadOnly)
            If frmShow Then
                frmAssignedInspec.CallingForm = Me
                frmAssignedInspec.ShowDialog()
                frmAssignedInspec = Nothing
                If Me.Tag = "1" Then
                    PopulateAssignedInspections(nInspectorID)
                End If
            End If
        Catch ex As Exception
            If frmShow Then
                Dim MyErr As New ErrorReport(ex)
                MyErr.ShowDialog()
            Else
                Throw ex
            End If
        End Try
    End Sub
#Region "Events"
    ' ugAssignedInspections
    Private Sub ugAssignedInspections_AfterSortChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BandEventArgs) Handles ugAssignedInspections.AfterSortChange
        Try
            If ugAssignedInspections.Rows.Count > 0 Then
                ugAssignedInspections.ActiveRow = ugAssignedInspections.Rows(0)
            End If
            If ugAssignedInspec.Rows.Count > 0 Then
                ugAssignedInspec.ActiveRow = ugAssignedInspec.Rows(0)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugAssignedInspections_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugAssignedInspections.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            Me.EnterAssignedInspection()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ugAssignedInspec_AfterSortChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BandEventArgs) Handles ugAssignedInspec.AfterSortChange
        Try
            If ugAssignedInspec.Rows.Count > 0 Then
                ugAssignedInspec.ActiveRow = ugAssignedInspec.Rows(0)
            End If
            If ugAssignedInspections.Rows.Count > 0 Then
                ugAssignedInspections.ActiveRow = ugAssignedInspections.Rows(0)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugAssignedInspec_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugAssignedInspec.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            Me.EnterAssignedInspection()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnEnterInspectionInfo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEnterInspectionInfo.Click
        Try
            Me.EnterAssignedInspection()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnAssignedInspecEnterInspec_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAssignedInspecEnterInspec.Click
        Try
            Me.EnterAssignedInspection()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnPrintInspectionInfo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrintInspectionInfo.Click
        Try
            Me.EnterAssignedInspection(False)
            If Not frmAssignedInspec Is Nothing Then
                frmAssignedInspec.PrintAssignedInspection()
                frmAssignedInspec = Nothing
                oInspection.Remove(oInspection.ID)
                oInspection = New MUSTER.BusinessLogic.pInspection
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnAssignedInspecPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAssignedInspecPrint.Click
        btnPrintInspectionInfo_Click(sender, e)
    End Sub
#End Region
#End Region
#Region "Schedule/Reschedule Inspection"
    Private Sub ScheduleInspection()
        Try
            ugRow = Nothing
            ugRow = getSelectedRow()
            If ugRow Is Nothing Then
                MsgBox("Please select a facilty")
                Exit Sub
            End If
            If ugRow.Cells("INSPECTION_ID").Value Is System.DBNull.Value Then
                'oInspection = New MUSTER.BusinessLogic.pInspection
                oInspection.Add(0)
                oInspection.OwnerID = ugRow.Cells("OWNER_ID").Value
                oInspection.FacilityID = ugRow.Cells("FACILITY ID").Value
                oInspection.StaffID = MusterContainer.AppUser.UserKey
                nSelectedFacilityID = ugRow.Cells("FACILITY ID").Value
            Else
                'oInspection.Retrieve(ugRow.Cells("INSPECTION_ID").Value, , ugRow.Cells("FACILITY ID").Value, ugRow.Cells("OWNER_ID").Value)
                oInspection.Retrieve(ugRow.Cells("INSPECTION_ID").Value, , ugRow.Cells("FACILITY ID").Value)
                nSelectedFacilityID = ugRow.Cells("FACILITY ID").Value
            End If
            frmReschedule = New RescheduleInspection(oInspection, dsTargetFacs.Tables(0))
            frmReschedule.CallingForm = Me
            Me.Tag = "0"
            frmReschedule.ShowDialog()
            If Me.Tag = "1" Then
                If ugAllFacsForOwner.Rows.Count > 0 Then
                    PopulateMyFacilities(nInspectorID, ugRow.Cells("OWNER_ID").Value, nSelectedFacilityID, chkShowAllTargetFacilities.Checked, , oInspection.ID)
                Else
                    PopulateMyFacilities(nInspectorID, , nSelectedFacilityID, , , oInspection.ID)
                End If
            ElseIf Me.Tag = "2" Then
                ' inspection deleted
                If ugAllFacsForOwner.Rows.Count > 0 Then
                    PopulateMyFacilities(nInspectorID, ugRow.Cells("OWNER_ID").Value, nSelectedFacilityID, chkShowAllTargetFacilities.Checked, , -99)
                Else
                    PopulateMyFacilities(nInspectorID, , nSelectedFacilityID, , , -99)
                End If
            End If
            oInspection.Remove(oInspection.ID)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            If Not frmReschedule Is Nothing Then
                frmReschedule = Nothing
            End If
        End Try
    End Sub
#Region "Events"
    Private Sub btnReschedule_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReschedule.Click
        Try
            Me.ScheduleInspection()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnAllFacSchedule_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAllFacSchedule.Click
        Try
            Me.ScheduleInspection()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnMyFacSchedule_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMyFacSchedule.Click
        Try
            Me.ScheduleInspection()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#End Region
#Region "Generate Letters"
    ' generate letters
    Private Sub btnGenerateLetters_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGenerateLetters.Click
        Dim ugFacRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim dtFacs As New DataTable
        'Dim dtFacsforOwner As New DataTable
        Dim strErr As String
        Dim strSelectedOwner, strSelectedFacs As String
        'Dim dr, drNew As DataRow
        Dim drNew As DataRow
        'Dim alOwner As New ArrayList
        Dim nOwner As Integer
        Dim facCount As Integer = 0
        Dim progressBarValueIncrement As Single
        'Dim facLastInspDate As Date
        Dim oLetter As New Reg_Letters
        Dim bolIncludeAllFacsForOwner As Boolean = True
        Dim bolRegenerateForOwner As Boolean = False
        Try
            ugRow = Nothing
            ugRow = getSelectedRow()
            If ugRow Is Nothing Then
                MsgBox("Please select an owner to Generate Inspection Announcement Letters")
                Exit Sub
            End If

            nOwner = ugRow.Cells("OWNER_ID").Value
            nSelectedFacilityID = ugRow.Cells("FACILITY ID").Value
            strSelectedOwner = ugRow.Cells("OWNER NAME").Value.ToString + " (" + ugRow.Cells("OWNER_ID").Value.ToString + ")"

            If ugRow.Cells("RESCHEDULED DATE").Value Is DBNull.Value And _
                    ugRow.Cells("SCHEDULED DATE").Value Is DBNull.Value Then
                MsgBox("Cannot generate letter(s) for unscheduled" + vbCrLf + _
                "Facility: " + ugRow.Cells("FACILITY").Text + " (" + nSelectedFacilityID.ToString + ")" + vbCrLf + _
                vbCrLf + _
                "Aborting Generating Letter(s)")
                Exit Sub
            End If

            ' list of facs for the given owner to be included in the letter
            dtFacs.Columns.Add("INSPECTION_ID")
            dtFacs.Columns.Add("OWNER_ID")
            dtFacs.Columns.Add("OWNER_NAME")
            dtFacs.Columns.Add("FACILITY_ID")
            dtFacs.Columns.Add("FACILITY_NAME")
            dtFacs.Columns.Add("SCHEDULE_DATE")
            dtFacs.Columns.Add("SCHEDULE_TIME")
            dtFacs.Columns.Add("LETTER_GENERATED")
            dtFacs.Columns.Add("LAST_INSPECTED_ON")

            'dtFacsforOwner.Columns.Add("INSPECTION_ID")
            'dtFacsforOwner.Columns.Add("OWNER_ID")
            'dtFacsforOwner.Columns.Add("OWNER_NAME")
            'dtFacsforOwner.Columns.Add("FACILITY_ID")
            'dtFacsforOwner.Columns.Add("FACILITY_NAME")
            'dtFacsforOwner.Columns.Add("SCHEDULE_DATE")
            'dtFacsforOwner.Columns.Add("SCHEDULE_TIME")
            'dtFacsforOwner.Columns.Add("LETTER_GENERATED")

            If ugRow.Tag Is Nothing Then
                ' not selected from the 2 associated grids
            ElseIf ugRow.Tag.ToString.Equals(ugMyFacs.Name) Or ugRow.Tag.ToString.Equals(ugMyFacilities.Name) Then
                ' double click the row
                ' to get all the facs for the selected owner
                'ugMyFacs_DoubleClick(sender, e)
                Cursor.Current = Cursors.AppStarting
                PopulateFacsforOwner(nInspectorID, ugRow.Cells("OWNER_ID").Value, ugRow.Cells("FACILITY ID").Value, , , False)
                'ugAllFacsForOwner.ActiveRow = Nothing
                Cursor.Current = Cursors.Default
            ElseIf ugRow.Tag.ToString.Equals(ugAllFacsForOwner.Name) Or ugRow.Tag.ToString.Equals(ugAllFacForOwner.Name) Then
                ' all facilities to generate the letter are in the grid - continue
            Else
                ' invalid grid
                MsgBox("Invalid Grid row selected")
                Exit Sub
            End If

            ' If Letter generated, prompt user if he/she wants to re-generate letter for the selected facility
            ' or for all the facilities (whose letter was generated) for the selected owner
            If Not ugRow.Cells("LETTER GENERATED").Value Is DBNull.Value Then
                If ugRow.Cells("LETTER GENERATED").Value Then
                    If ugAllFacsForOwner.Rows.Count = 1 Then
                        drNew = dtFacs.NewRow
                        drNew("INSPECTION_ID") = ugRow.Cells("INSPECTION_ID").Value
                        drNew("OWNER_ID") = ugRow.Cells("OWNER_ID").Value
                        drNew("OWNER_NAME") = ugRow.Cells("OWNER NAME").Value
                        drNew("FACILITY_ID") = ugRow.Cells("FACILITY ID").Value
                        drNew("FACILITY_NAME") = ugRow.Cells("FACILITY").Value
                        drNew("SCHEDULE_DATE") = ugRow.Cells("SCHEDULED DATE").Value
                        drNew("SCHEDULE_TIME") = ugRow.Cells("SCHEDULED TIME").Value
                        drNew("LETTER_GENERATED") = ugRow.Cells("LETTER GENERATED").Value
                        drNew("LAST_INSPECTED_ON") = ugRow.Cells("LAST INSPECTED ON").Value
                        dtFacs.Rows.Add(drNew)
                        facCount += 1
                        strSelectedFacs += "Facility: " + ugRow.Cells("FACILITY").Value.ToString + " (" + _
                                            ugRow.Cells("FACILITY ID").Value.ToString + ")" + vbCrLf
                    ElseIf ugAllFacsForOwner.Rows.Count > 1 Then
                        Dim strMsg As String = "Do you want to re-generate Inspection Announcement letter for the selected Facility - " + vbCrLf + _
                                                ugRow.Cells("FACILITY").Value.ToString + " (" + nSelectedFacilityID.ToString + ") only, OR" + vbCrLf + _
                                                "for all the facilities (for selected owner) with announcement letter(s) generated previously?" + vbCrLf + _
                                                "Yes - re-generate for " + nSelectedFacilityID.ToString + " only" + vbCrLf + _
                                                "No - all facilities with letter(s) generated previously"
                        Dim result As MsgBoxResult = MsgBox(strMsg, MsgBoxStyle.YesNoCancel, "Regenerate Inspection Announcement Letter")
                        If result = MsgBoxResult.Cancel Then
                            Exit Sub
                        ElseIf result = MsgBoxResult.Yes Then
                            drNew = dtFacs.NewRow
                            drNew("INSPECTION_ID") = ugRow.Cells("INSPECTION_ID").Value
                            drNew("OWNER_ID") = ugRow.Cells("OWNER_ID").Value
                            drNew("OWNER_NAME") = ugRow.Cells("OWNER NAME").Value
                            drNew("FACILITY_ID") = ugRow.Cells("FACILITY ID").Value
                            drNew("FACILITY_NAME") = ugRow.Cells("FACILITY").Value
                            drNew("SCHEDULE_DATE") = ugRow.Cells("SCHEDULED DATE").Value
                            drNew("SCHEDULE_TIME") = ugRow.Cells("SCHEDULED TIME").Value
                            drNew("LETTER_GENERATED") = ugRow.Cells("LETTER GENERATED").Value
                            drNew("LAST_INSPECTED_ON") = ugRow.Cells("LAST INSPECTED ON").Value
                            dtFacs.Rows.Add(drNew)
                            facCount += 1
                            strSelectedFacs += "Facility: " + ugRow.Cells("FACILITY").Value.ToString + " (" + _
                                                ugRow.Cells("FACILITY ID").Value.ToString + ")" + vbCrLf
                        Else
                            bolRegenerateForOwner = True
                        End If
                    End If
                End If
            End If

            Cursor.Current = Cursors.AppStarting
            frmCLProgress = New CheckListProgress
            'facCount = 0
            strErr = ""

            If dtFacs.Rows.Count = 0 Then
                If ugRow.Cells("STAFF_ID").Value Is DBNull.Value Then
                    ' if fac is not assigned to the user, check all the facs for the owner
                    bolIncludeAllFacsForOwner = True
                ElseIf ugRow.Cells("STAFF_ID").Value = nInspectorID Then
                    ' if fac is assigned to the user, check only the facs assigned for the user
                    bolIncludeAllFacsForOwner = False
                Else
                    ' if fac is not assigned to the user, check all the facs for the owner
                    bolIncludeAllFacsForOwner = True
                End If

                For Each ugFacRow In ugAllFacsForOwner.Rows
                    If ugFacRow.Cells("OWNER_ID").Value = ugRow.Cells("OWNER_ID").Value Then
                        ' if there are any facs not scheduled by the user and assigned to the user, issue warning
                        If Not ugFacRow.Cells("SCHEDULED BY").Text = String.Empty Then
                            If (ugFacRow.Cells("SCHEDULED BY").Text.Trim.ToUpper <> MusterContainer.AppUser.Name.ToUpper.Trim Or _
                                ugFacRow.Cells("INSPECTION_ID").Value Is DBNull.Value) And _
                                Not bolRegenerateForOwner Then
                                If bolIncludeAllFacsForOwner Then
                                    strErr += "Facility: " + ugFacRow.Cells("FACILITY").Value.ToString + " (" + ugFacRow.Cells("FACILITY ID").Value.ToString + ")" + vbCrLf
                                Else
                                    If Not ugFacRow.Cells("STAFF_ID").Value Is DBNull.Value Then
                                        If ugFacRow.Cells("STAFF_ID").Value = MusterContainer.AppUser.UserKey Then
                                            strErr += "Facility: " + ugFacRow.Cells("FACILITY").Value.ToString + " (" + ugFacRow.Cells("FACILITY ID").Value.ToString + ")" + vbCrLf
                                        End If
                                    End If
                                End If
                                'strErr += "Facility: " + ugFacRow.Cells("FACILITY ID").Value.ToString + ", "
                            Else
                                ' if scheduled date and scheduled time is not null and letter has not been
                                ' generated add info into datatable  to be processed
                                If Not ugFacRow.Cells("SCHEDULED DATE").Value Is DBNull.Value Then
                                    If Not ugFacRow.Cells("SCHEDULED TIME").Value Is DBNull.Value Then
                                        ' Only need to include the facs scheduled by the inspector
                                        If ugFacRow.Cells("SCHEDULED BY").Text.Trim.ToUpper = MusterContainer.AppUser.Name.ToUpper.Trim Then
                                            If bolRegenerateForOwner Then
                                                If ugFacRow.Cells("LETTER GENERATED").Value Then
                                                    drNew = dtFacs.NewRow
                                                    drNew("INSPECTION_ID") = ugFacRow.Cells("INSPECTION_ID").Value
                                                    drNew("OWNER_ID") = ugFacRow.Cells("OWNER_ID").Value
                                                    drNew("OWNER_NAME") = ugFacRow.Cells("OWNER NAME").Value
                                                    drNew("FACILITY_ID") = ugFacRow.Cells("FACILITY ID").Value
                                                    drNew("FACILITY_NAME") = ugFacRow.Cells("FACILITY").Value
                                                    drNew("SCHEDULE_DATE") = ugFacRow.Cells("SCHEDULED DATE").Value
                                                    drNew("SCHEDULE_TIME") = ugFacRow.Cells("SCHEDULED TIME").Value
                                                    drNew("LETTER_GENERATED") = ugFacRow.Cells("LETTER GENERATED").Value
                                                    drNew("LAST_INSPECTED_ON") = ugFacRow.Cells("LAST INSPECTED ON").Value
                                                    dtFacs.Rows.Add(drNew)
                                                    facCount += 1
                                                    strSelectedFacs += "Facility: " + ugFacRow.Cells("FACILITY").Value.ToString + " (" + _
                                                                        ugFacRow.Cells("FACILITY ID").Value.ToString + ")" + vbCrLf
                                                End If
                                            Else
                                                If Not ugFacRow.Cells("LETTER GENERATED").Value Then
                                                    drNew = dtFacs.NewRow
                                                    drNew("INSPECTION_ID") = ugFacRow.Cells("INSPECTION_ID").Value
                                                    drNew("OWNER_ID") = ugFacRow.Cells("OWNER_ID").Value
                                                    drNew("OWNER_NAME") = ugFacRow.Cells("OWNER NAME").Value
                                                    drNew("FACILITY_ID") = ugFacRow.Cells("FACILITY ID").Value
                                                    drNew("FACILITY_NAME") = ugFacRow.Cells("FACILITY").Value
                                                    drNew("SCHEDULE_DATE") = ugFacRow.Cells("SCHEDULED DATE").Value
                                                    drNew("SCHEDULE_TIME") = ugFacRow.Cells("SCHEDULED TIME").Value
                                                    drNew("LETTER_GENERATED") = ugFacRow.Cells("LETTER GENERATED").Value
                                                    drNew("LAST_INSPECTED_ON") = ugFacRow.Cells("LAST INSPECTED ON").Value
                                                    dtFacs.Rows.Add(drNew)
                                                    facCount += 1
                                                    strSelectedFacs += "Facility: " + ugFacRow.Cells("FACILITY").Value.ToString + " (" + _
                                                                        ugFacRow.Cells("FACILITY ID").Value.ToString + ")" + vbCrLf
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        ElseIf Not bolRegenerateForOwner Then
                            'strErr += "Facility: " + ugFacRow.Cells("FACILITY").Value.ToString + " (" + ugFacRow.Cells("FACILITY ID").Value.ToString + ")" + vbCrLf
                            If bolIncludeAllFacsForOwner Then
                                strErr += "Facility: " + ugFacRow.Cells("FACILITY").Value.ToString + " (" + ugFacRow.Cells("FACILITY ID").Value.ToString + ")" + vbCrLf
                            Else
                                If Not ugFacRow.Cells("STAFF_ID").Value Is DBNull.Value Then
                                    If ugFacRow.Cells("STAFF_ID").Value = MusterContainer.AppUser.UserKey Then
                                        strErr += "Facility: " + ugFacRow.Cells("FACILITY").Value.ToString + " (" + ugFacRow.Cells("FACILITY ID").Value.ToString + ")" + vbCrLf
                                    End If
                                End If
                            End If
                        End If ' If Not ugFacRow.Cells("SCHEDULED BY").Text = String.Empty Then
                    End If ' If ugFacRow.Cells("OWNER_ID").Value = ugRow.Cells("OWNER_ID").Value Then
                Next
            End If ' If dtFacs.Rows.Count = 0 Then

            Cursor.Current = Cursors.Default
            If strErr.Length > 0 Then
                Dim mbc As MessageBoxCustom
                If mbc.Show("One or more Facilities are incompletely scheduled for Owner : " + strSelectedOwner + _
                    vbCrLf + strErr + "Do you want to continue?", , MessageBoxButtons.YesNo) = DialogResult.No Then
                    Exit Sub
                End If
                'Dim msgResult As MsgBoxResult
                'msgResult = MsgBox("One or more Facilities are incompletely scheduled for Owner : " + strSelectedOwner + _
                'vbCrLf + strErr + "Do you want to continue?", MsgBoxStyle.YesNo)
                'If msgResult = MsgBoxResult.No Then
                '    Exit Sub
                'End If
            End If
            If dtFacs.Rows.Count = 0 Then
                MsgBox("There are no scheduled Inspections" + vbCrLf + "OR" + vbCrLf + "Letter(s) generated for the scheduled Inspections")
            Else
                Cursor.Current = Cursors.AppStarting
                ' progress bar
                progressBarValueIncrement = 100 / facCount
                frmCLProgress.HeaderText = "Generating Inspection Announcement Letter(s)"
                frmCLProgress.Show()
                UIUtilsGen.Delay(, 0.5)

                Dim oOwn As New MUSTER.BusinessLogic.pOwner
                ltrGen = New MUSTER.BusinessLogic.pLetterGen
                strErr = String.Empty
                oOwn.RetrieveAll(nOwner, "INSPECTION")
                oLetter.GenerateInspectionAnnouncementLetters(nOwner, "Inspection Announcement Letter", "AnnouncementLetter", "Inspection Announcement Letter", "InspectionAnnouncementLetter.doc", dtFacs, oOwn, ltrGen, progressBarValueIncrement)
                oOwn.Remove(nOwner)
                ' Delay after generating a word document to resolve RPC Server is Unavailable Issue
                UIUtilsGen.Delay(, 1)
                'strErr = strSelectedOwner + vbCrLf + strSelectedFacs
                ' update letter generated in dataset
                Dim dr As DataRow
                For Each dr In dsTargetFacs.Tables(0).Select("OWNER_ID = " + ugRow.Cells("OWNER_ID").Text)
                    If dtFacs.Select("FACILITY_ID = " + dr("FACILITY ID").ToString).Length > 0 Then
                        dr("LETTER GENERATED") = True
                    End If
                Next
                For Each dr In dsTargetFacsForOwner.Tables(0).Select("OWNER_ID = " + ugRow.Cells("OWNER_ID").Text)
                    If dtFacs.Select("FACILITY_ID = " + dr("FACILITY ID").ToString).Length > 0 Then
                        dr("LETTER GENERATED") = True
                    End If
                Next
                Populate(nInspectorID, nOwner, nSelectedFacilityID, , -1, False)
                Cursor.Current = Cursors.Default
                Dim mbc As MessageBoxCustom
                mbc.Show("Letters Generated Successfully for " + vbCrLf + strSelectedOwner + vbCrLf + strSelectedFacs)
                'MsgBox("Letters Generated Successfully for " + vbCrLf + strErr)
            End If
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            Cursor.Current = Cursors.Default
            If Not frmCLProgress Is Nothing Then frmCLProgress.Close()
        End Try
    End Sub
    Private Sub btnMyFacGenLetters_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMyFacGenLetters.Click
        Try
            btnGenerateLetters_Click(sender, e)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnAllFacGenLetters_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAllFacGenLetters.Click
        Try
            btnGenerateLetters_Click(sender, e)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#Region "Print Checklist"
    Private Sub PrintChecklist(ByVal alugRows As ArrayList)
        Dim dsTanks As New DataSet
        Dim dtTank As New DataTable
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim dtFacsPrinted As New DataTable
        Dim dr, dr1 As DataRow
        Dim dopics As Integer = 0
        Dim cnt As Integer = 0
        Try
            dtFacsPrinted.Columns.Add("FACILITY ID", GetType(Integer))
            dtFacsPrinted.Columns.Add("CHECKLIST GENERATED", GetType(Date))

            For Each ugrow In alugRows
                Dim doMsg As Boolean = False

                If cnt >= alugRows.Count - 1 Then
                    doMsg = True
                End If
                oInspection.Retrieve(CType(ugrow.Cells("INSPECTION_ID").Value, Int64), , CType(ugrow.Cells("FACILITY ID").Value, Int64), CType(ugrow.Cells("OWNER_ID").Value, Int64))
                frmChecklist = New CheckList(oInspection, , True)
                If oInspection.IsDirty Then
                    If oInspection.ID <= 0 Then
                        oInspection.CreatedBy = MusterContainer.AppUser.ID
                    Else
                        oInspection.ModifiedBy = MusterContainer.AppUser.ID
                    End If

                    oInspection.Save(CType(UIUtilsGen.ModuleID.Inspection, Integer), MusterContainer.AppUser.UserKey, returnVal)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If
                End If

                frmChecklist.PrintCheckList(dopics, "Printing CL for FacID : " + ugrow.Cells("FACILITY ID").Value.ToString, doMsg)
                frmChecklist = Nothing

                dr = dtFacsPrinted.NewRow
                dr("FACILITY ID") = ugrow.Cells("FACILITY ID").Value
                dr("CHECKLIST GENERATED") = oInspection.CheckListGenDate
                dtFacsPrinted.Rows.Add(dr)

                oInspection.Remove(oInspection.InspectionInfo)
                oInspection = New MUSTER.BusinessLogic.pInspection
                cnt += 1
            Next
            dopics = 0
            ' update both the datasets ("my facs" and "facs for sel owner" if present)
            If dtFacsPrinted.Rows.Count > 0 Then
                Dim bolUpdatedsTargetFacsForOwner As Boolean = False
                Dim selectedOwnerID As Integer = 0

                If Not dsTargetFacsForOwner Is Nothing Then
                    If dsTargetFacsForOwner.Tables.Count > 0 Then
                        If dsTargetFacsForOwner.Tables(0).Rows.Count > 0 Then
                            bolUpdatedsTargetFacsForOwner = True
                            selectedOwnerID = dsTargetFacsForOwner.Tables(0).Rows(0)("OWNER_ID")
                        End If
                    End If
                End If

                For Each dr In dtFacsPrinted.Rows
                    For Each dr1 In dsTargetFacs.Tables(0).Select("[FACILITY ID] = " + dr("FACILITY ID").ToString)
                        If dr1("FACILITY ID") = dr("FACILITY ID") Then
                            dr1("CHECKLIST GENERATED") = dr("CHECKLIST GENERATED")
                        End If
                    Next
                    If bolUpdatedsTargetFacsForOwner Then
                        For Each dr1 In dsTargetFacsForOwner.Tables(0).Select("[FACILITY ID] = " + dr("FACILITY ID").ToString)
                            If dr1("FACILITY ID") = dr("FACILITY ID") Then
                                dr1("CHECKLIST GENERATED") = dr("CHECKLIST GENERATED")
                            End If
                        Next
                    End If
                Next
                Populate(nInspectorID, selectedOwnerID, , , -1, False)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnPrintChecklists_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrintChecklists.Click
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim ugrows As Infragistics.Win.UltraWinGrid.SelectedRowsCollection
        Dim strFacs As String = String.Empty
        Dim facCount As Int64
        Dim alugRows As New ArrayList
        Try
            facCount = 0
            ugrows = getSelectedRows()
            If ugrows Is Nothing Then
                MsgBox("Please select a facilty")
                Exit Sub
            End If
            Cursor.Current = Cursors.AppStarting
            For Each ugrow In ugrows
                If ugrow.Cells("RESCHEDULED DATE").Value Is DBNull.Value And _
                    ugrow.Cells("SCHEDULED DATE").Value Is DBNull.Value Then
                    strFacs += ugrow.Cells("FACILITY ID").Value.ToString + ", "
                    facCount += 1
                Else
                    alugRows.Add(ugrow)
                End If
            Next
            Cursor.Current = Cursors.Default
            If facCount > 0 Then
                strFacs = strFacs.Trim.Trim(",")
                If facCount = ugrows.Count Then
                    MsgBox("Cannot print checklist for unscheduled" + vbCrLf + _
                    IIf(facCount = 1, "Facility: ", "Facilities: ") + strFacs + _
                    vbCrLf + vbCrLf + _
                    "Aborting Print Checklist")
                    Exit Sub
                Else
                    Dim msgResult As MsgBoxResult
                    msgResult = MsgBox("Cannot print checklist for unscheduled" + vbCrLf + _
                    IIf(facCount = 1, "Facility: ", "Facilities: ") + strFacs + _
                    vbCrLf + vbCrLf + _
                    "Do you want to continue printing " + _
                    IIf(ugrows.Count - facCount > 1, "checklists for scheduled facilities?", "checklist for scheduled facility?"), MsgBoxStyle.YesNo)
                    If msgResult = MsgBoxResult.No Then
                        Exit Sub
                    End If
                End If
            End If
            Cursor.Current = Cursors.AppStarting
            PrintChecklist(alugRows)
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnMyFacPrintChecklists_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMyFacPrintChecklists.Click
        Try
            btnPrintChecklists_Click(sender, e)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnAllFacPrintChecklist_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAllFacPrintChecklist.Click
        Try
            btnPrintChecklists_Click(sender, e)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#Region "External Events"
    Private Sub ltrGen_CheckListProgress(ByVal percent As Single) Handles ltrGen.CheckListProgress
        Try
            If frmCLProgress.ProgressBarValue >= frmCLProgress.ProgressBarMax Then
                frmCLProgress.Close()
            Else
                frmCLProgress.ProgressBarValue += percent
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ltrGen_CloseCheckListProgress() Handles ltrGen.CloseCheckListProgress
        Try
            If Not frmCLProgress Is Nothing Then
                frmCLProgress.Close()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#Region "Form Events"
    Private Sub Inspection_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            LoadInspectors()
            bolLoading = True
            UIUtilsGen.SetComboboxItemByValue(cmbInspectors, MusterContainer.AppUser.UserKey)
            If cmbInspectors.SelectedIndex = -1 Then
                UIUtilsGen.SetComboboxItemByText(cmbInspectors, "")
            End If
            bolLoading = False
            Populate(nInspectorID, nOwnerID, nFacilityID)
            If Not rdBtnDue.Checked Then rdBtnDue.Checked = True


            If Not ugMyFacilities Is Nothing AndAlso Not ugMyFacilities.Rows Is Nothing AndAlso ugMyFacilities.Rows.Count = 1 Then
                ugMyFacilities.Rows(0).Selected = True
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub Inspection_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Try
            MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "Inspection")
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ugMyFacs_selected(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs) Handles ugMyFacs.AfterSelectChange
        With Me.ugMyFacs
            If Not _container Is Nothing AndAlso .Selected.Rows.Count > 0 Then
                _container.txtOwnerQSKeyword.Text = .Selected.Rows.Item(0).Cells("FACILITY ID").Value
                _container.cmbQuickSearchFilter.SelectedValue = "Facility ID"
                _container.cmbSearchModule.SelectedValue = 615

            End If
        End With
    End Sub


    Private Sub ugMyFacilties_selected(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs) Handles ugMyFacilities.AfterSelectChange
        With Me.ugMyFacilities
            If Not _container Is Nothing AndAlso .Selected.Rows.Count > 0 Then
                _container.txtOwnerQSKeyword.Text = .Selected.Rows.Item(0).Cells("FACILITY ID").Value
                _container.cmbQuickSearchFilter.SelectedValue = "Facility ID"
                _container.cmbSearchModule.SelectedValue = 615

            End If
        End With
    End Sub

    Private Sub ugAllFacForOwners_selected(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs) Handles ugAllFacForOwner.AfterSelectChange
        With Me.ugAllFacForOwner
            If Not _container Is Nothing AndAlso .Selected.Rows.Count > 0 Then
                _container.txtOwnerQSKeyword.Text = .Selected.Rows.Item(0).Cells("FACILITY ID").Value
                _container.cmbQuickSearchFilter.SelectedValue = "Facility ID"
                _container.cmbSearchModule.SelectedValue = 615

            End If
        End With
    End Sub

    Private Sub ugAllFacsForOwners_selected(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs) Handles ugAllFacsForOwner.AfterSelectChange
        With Me.ugAllFacsForOwner
            If Not _container Is Nothing AndAlso .Selected.Rows.Count > 0 Then
                _container.txtOwnerQSKeyword.Text = .Selected.Rows.Item(0).Cells("FACILITY ID").Value
                _container.cmbQuickSearchFilter.SelectedValue = "Facility ID"
                _container.cmbSearchModule.SelectedValue = 615

            End If
        End With
    End Sub

    Private Sub ugAssignedInspec_selected(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs) Handles ugAssignedInspec.AfterSelectChange
        With Me.ugAssignedInspec
            If Not _container Is Nothing AndAlso .Selected.Rows.Count > 0 Then
                _container.txtOwnerQSKeyword.Text = .Selected.Rows.Item(0).Cells("FACILITY ID").Value
                _container.cmbQuickSearchFilter.SelectedValue = "Facility ID"
                _container.cmbSearchModule.SelectedValue = 615

            End If
        End With
    End Sub

    Private Sub ugAssignedInspections_selected(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs) Handles ugAssignedInspections.AfterSelectChange
        With Me.ugAssignedInspections
            If Not _container Is Nothing AndAlso .Selected.Rows.Count > 0 Then
                _container.txtOwnerQSKeyword.Text = .Selected.Rows.Item(0).Cells("FACILITY ID").Value
                _container.cmbQuickSearchFilter.SelectedValue = "Facility ID"
                _container.cmbSearchModule.SelectedValue = 615
            End If
        End With
    End Sub

    Private Sub Inspection_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.TextChanged
        Try
            MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "WindowName", Me.Text, "Inspection")
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

    Private Sub Inspection_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Disposed
        If Not commentDataSet Is Nothing Then
            commentDataSet.Dispose()
        End If
    End Sub



End Class
