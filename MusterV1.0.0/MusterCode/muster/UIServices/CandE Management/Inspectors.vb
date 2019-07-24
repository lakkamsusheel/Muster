Public Class Inspectors
    Inherits System.Windows.Forms.Form
#Region "User Defined Variables"
    Dim WithEvents oInspectorCounty As New MUSTER.BusinessLogic.pInspectorCountyAssociation
    Dim WithEvents oInspectorOwner As New MUSTER.BusinessLogic.pInspectorOwnerAssignment
    Dim dtAVailableCountyFacilities As New DataTable
    Dim dtAvailableOwnerFacilities As New DataTable

    Dim dtAvailableInspectors As New DataTable
    Dim dtAssignedInspectors As New DataTable

    Dim dtAssCountyOwners As New DataTable
    Dim dtAssInspectorOwners As New DataTable
    Dim firstTimeCounty As Boolean = True
    Dim firstTimeOwner As Boolean = True
    Public MyGuid As New System.Guid
    ' to be changed
    Dim nOwnerID As Integer = 1
    Dim nManagerID As Integer = -1
    Dim returnVal As String = String.Empty

    Dim AddOnInspectors As New Collections.ArrayList
    Dim RemoveInspectors As New Collections.ArrayList
    Dim _dtInspectorData As DataTable

#End Region

#Region "public properties"

    Public Property dtInspectorData() As DataTable
        Get
            If _dtInspectorData Is Nothing Then
                _dtInspectorData = New DataTable

                _dtInspectorData.Columns.Add("PropertyID", (1).GetType)
                _dtInspectorData.Columns.Add("PropertyName", String.Empty.GetType)

            End If

            Return _dtInspectorData
        End Get

        Set(ByVal Value As DataTable)
            _dtInspectorData = Value
        End Set
    End Property

#End Region
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()
        MyGuid = System.Guid.NewGuid
        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        MusterContainer.AppUser.LogEntry("CandEManagement", MyGuid.ToString)
        MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "CandEManagement")
    End Sub



    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)

        AddOnInspectors.Clear()
        AddOnInspectors = Nothing
        RemoveInspectors.Clear()
        RemoveInspectors = Nothing

        If Not _dtInspectorData Is Nothing Then
            _dtInspectorData.Dispose()
        End If



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
    Friend WithEvents ugAvailableCounties As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents ugAssignedCounties As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents ugAssignedOwnersFacCount As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents ugAvailableOwnersFacCount As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents lblAvailableCounties As System.Windows.Forms.Label
    Friend WithEvents lblAssignedCounties As System.Windows.Forms.Label
    Friend WithEvents lblAvailableOwnersFacilityCount As System.Windows.Forms.Label
    Friend WithEvents lblAssignedOwnersFacilityCount As System.Windows.Forms.Label
    Friend WithEvents chkInactive As System.Windows.Forms.CheckBox
    Friend WithEvents lblOwnerFacilities As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents cmbInspector As System.Windows.Forms.ComboBox
    Friend WithEvents btnCountyShiftLeftAll As System.Windows.Forms.Button
    Friend WithEvents btnCountyShiftLeft As System.Windows.Forms.Button
    Friend WithEvents btnCountyShiftRightAll As System.Windows.Forms.Button
    Friend WithEvents btnCountyShiftRight As System.Windows.Forms.Button
    Friend WithEvents btninspectorShiftLeftAll As System.Windows.Forms.Button
    Friend WithEvents btninspectorShiftLeft As System.Windows.Forms.Button
    Friend WithEvents btninspectorShiftRightAll As System.Windows.Forms.Button
    Friend WithEvents btninspectorShiftRight As System.Windows.Forms.Button
    Friend WithEvents lblCountiesSumValue As System.Windows.Forms.Label
    Friend WithEvents lblCountiesSum As System.Windows.Forms.Label
    Friend WithEvents txtOwnerFacilities As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lblAvailCountiesSumValue As System.Windows.Forms.Label
    Friend WithEvents cmbInspectorActive As System.Windows.Forms.ComboBox
    Friend WithEvents lblActive As System.Windows.Forms.Label
    Friend WithEvents cboManagers As System.Windows.Forms.ComboBox
    Friend WithEvents lblManager As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents BtnOwnerShiftLeftAll As System.Windows.Forms.Button
    Friend WithEvents btnOwnerShiftLeft As System.Windows.Forms.Button
    Friend WithEvents BtnOwnerShiftRightAll As System.Windows.Forms.Button
    Friend WithEvents BtnOwnerShiftRight As System.Windows.Forms.Button
    Friend WithEvents lblInspector As System.Windows.Forms.Label
    Friend WithEvents ugAssignedInspectorsToManagers As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents ugAvailableInspectors As Infragistics.Win.UltraWinGrid.UltraGrid
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmbInspector = New System.Windows.Forms.ComboBox
        Me.lblInspector = New System.Windows.Forms.Label
        Me.ugAvailableCounties = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.btnCountyShiftLeftAll = New System.Windows.Forms.Button
        Me.btnCountyShiftLeft = New System.Windows.Forms.Button
        Me.btnCountyShiftRightAll = New System.Windows.Forms.Button
        Me.btnCountyShiftRight = New System.Windows.Forms.Button
        Me.ugAssignedCounties = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.ugAssignedOwnersFacCount = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.BtnOwnerShiftLeftAll = New System.Windows.Forms.Button
        Me.btnOwnerShiftLeft = New System.Windows.Forms.Button
        Me.BtnOwnerShiftRightAll = New System.Windows.Forms.Button
        Me.BtnOwnerShiftRight = New System.Windows.Forms.Button
        Me.btninspectorShiftLeftAll = New System.Windows.Forms.Button
        Me.btninspectorShiftLeft = New System.Windows.Forms.Button
        Me.btninspectorShiftRightAll = New System.Windows.Forms.Button
        Me.btninspectorShiftRight = New System.Windows.Forms.Button
        Me.ugAvailableOwnersFacCount = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.lblAvailableCounties = New System.Windows.Forms.Label
        Me.lblAssignedCounties = New System.Windows.Forms.Label
        Me.lblAvailableOwnersFacilityCount = New System.Windows.Forms.Label
        Me.lblAssignedOwnersFacilityCount = New System.Windows.Forms.Label
        Me.chkInactive = New System.Windows.Forms.CheckBox
        Me.lblOwnerFacilities = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.lblCountiesSumValue = New System.Windows.Forms.Label
        Me.lblCountiesSum = New System.Windows.Forms.Label
        Me.txtOwnerFacilities = New System.Windows.Forms.Label
        Me.lblAvailCountiesSumValue = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.cmbInspectorActive = New System.Windows.Forms.ComboBox
        Me.lblActive = New System.Windows.Forms.Label
        Me.cboManagers = New System.Windows.Forms.ComboBox
        Me.lblManager = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.ugAssignedInspectorsToManagers = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.ugAvailableInspectors = New Infragistics.Win.UltraWinGrid.UltraGrid
        CType(Me.ugAvailableCounties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ugAssignedCounties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ugAssignedOwnersFacCount, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ugAvailableOwnersFacCount, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ugAssignedInspectorsToManagers, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ugAvailableInspectors, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmbInspector
        '
        Me.cmbInspector.Location = New System.Drawing.Point(104, 24)
        Me.cmbInspector.Name = "cmbInspector"
        Me.cmbInspector.Size = New System.Drawing.Size(192, 21)
        Me.cmbInspector.TabIndex = 0
        '
        'lblInspector
        '
        Me.lblInspector.Location = New System.Drawing.Point(32, 24)
        Me.lblInspector.Name = "lblInspector"
        Me.lblInspector.Size = New System.Drawing.Size(64, 16)
        Me.lblInspector.TabIndex = 250
        Me.lblInspector.Text = "Inspector:"
        Me.lblInspector.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'ugAvailableCounties
        '
        Me.ugAvailableCounties.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugAvailableCounties.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True
        Me.ugAvailableCounties.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugAvailableCounties.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugAvailableCounties.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugAvailableCounties.Location = New System.Drawing.Point(8, 80)
        Me.ugAvailableCounties.Name = "ugAvailableCounties"
        Me.ugAvailableCounties.Size = New System.Drawing.Size(344, 144)
        Me.ugAvailableCounties.TabIndex = 5
        '
        'btnCountyShiftLeftAll
        '
        Me.btnCountyShiftLeftAll.BackColor = System.Drawing.SystemColors.Control
        Me.btnCountyShiftLeftAll.Location = New System.Drawing.Point(368, 184)
        Me.btnCountyShiftLeftAll.Name = "btnCountyShiftLeftAll"
        Me.btnCountyShiftLeftAll.Size = New System.Drawing.Size(32, 24)
        Me.btnCountyShiftLeftAll.TabIndex = 9
        Me.btnCountyShiftLeftAll.Text = "<<"
        '
        'btnCountyShiftLeft
        '
        Me.btnCountyShiftLeft.BackColor = System.Drawing.SystemColors.Control
        Me.btnCountyShiftLeft.Location = New System.Drawing.Point(368, 160)
        Me.btnCountyShiftLeft.Name = "btnCountyShiftLeft"
        Me.btnCountyShiftLeft.Size = New System.Drawing.Size(32, 24)
        Me.btnCountyShiftLeft.TabIndex = 8
        Me.btnCountyShiftLeft.Text = "<"
        '
        'btnCountyShiftRightAll
        '
        Me.btnCountyShiftRightAll.BackColor = System.Drawing.SystemColors.Control
        Me.btnCountyShiftRightAll.Location = New System.Drawing.Point(368, 112)
        Me.btnCountyShiftRightAll.Name = "btnCountyShiftRightAll"
        Me.btnCountyShiftRightAll.Size = New System.Drawing.Size(32, 24)
        Me.btnCountyShiftRightAll.TabIndex = 7
        Me.btnCountyShiftRightAll.Text = ">>"
        '
        'btnCountyShiftRight
        '
        Me.btnCountyShiftRight.BackColor = System.Drawing.SystemColors.Control
        Me.btnCountyShiftRight.Location = New System.Drawing.Point(368, 88)
        Me.btnCountyShiftRight.Name = "btnCountyShiftRight"
        Me.btnCountyShiftRight.Size = New System.Drawing.Size(32, 24)
        Me.btnCountyShiftRight.TabIndex = 6
        Me.btnCountyShiftRight.Text = ">"
        '
        'ugAssignedCounties
        '
        Me.ugAssignedCounties.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugAssignedCounties.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugAssignedCounties.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugAssignedCounties.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugAssignedCounties.Location = New System.Drawing.Point(416, 80)
        Me.ugAssignedCounties.Name = "ugAssignedCounties"
        Me.ugAssignedCounties.Size = New System.Drawing.Size(344, 144)
        Me.ugAssignedCounties.TabIndex = 10
        '
        'ugAssignedOwnersFacCount
        '
        Me.ugAssignedOwnersFacCount.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugAssignedOwnersFacCount.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugAssignedOwnersFacCount.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugAssignedOwnersFacCount.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugAssignedOwnersFacCount.Location = New System.Drawing.Point(416, 320)
        Me.ugAssignedOwnersFacCount.Name = "ugAssignedOwnersFacCount"
        Me.ugAssignedOwnersFacCount.Size = New System.Drawing.Size(344, 144)
        Me.ugAssignedOwnersFacCount.TabIndex = 16
        '
        'BtnOwnerShiftLeftAll
        '
        Me.BtnOwnerShiftLeftAll.Location = New System.Drawing.Point(368, 432)
        Me.BtnOwnerShiftLeftAll.Name = "BtnOwnerShiftLeftAll"
        Me.BtnOwnerShiftLeftAll.Size = New System.Drawing.Size(32, 23)
        Me.BtnOwnerShiftLeftAll.TabIndex = 289
        Me.BtnOwnerShiftLeftAll.Text = "<<"
        '
        'btnOwnerShiftLeft
        '
        Me.btnOwnerShiftLeft.Location = New System.Drawing.Point(368, 408)
        Me.btnOwnerShiftLeft.Name = "btnOwnerShiftLeft"
        Me.btnOwnerShiftLeft.Size = New System.Drawing.Size(32, 23)
        Me.btnOwnerShiftLeft.TabIndex = 290
        Me.btnOwnerShiftLeft.Text = "<"
        '
        'BtnOwnerShiftRightAll
        '
        Me.BtnOwnerShiftRightAll.Location = New System.Drawing.Point(368, 352)
        Me.BtnOwnerShiftRightAll.Name = "BtnOwnerShiftRightAll"
        Me.BtnOwnerShiftRightAll.Size = New System.Drawing.Size(32, 23)
        Me.BtnOwnerShiftRightAll.TabIndex = 291
        Me.BtnOwnerShiftRightAll.Text = ">>"
        '
        'BtnOwnerShiftRight
        '
        Me.BtnOwnerShiftRight.Location = New System.Drawing.Point(368, 328)
        Me.BtnOwnerShiftRight.Name = "BtnOwnerShiftRight"
        Me.BtnOwnerShiftRight.Size = New System.Drawing.Size(32, 23)
        Me.BtnOwnerShiftRight.TabIndex = 292
        Me.BtnOwnerShiftRight.Text = ">"
        '
        'btninspectorShiftLeftAll
        '
        Me.btninspectorShiftLeftAll.Location = New System.Drawing.Point(368, 616)
        Me.btninspectorShiftLeftAll.Name = "btninspectorShiftLeftAll"
        Me.btninspectorShiftLeftAll.Size = New System.Drawing.Size(32, 23)
        Me.btninspectorShiftLeftAll.TabIndex = 0
        Me.btninspectorShiftLeftAll.Text = "<<"
        '
        'btninspectorShiftLeft
        '
        Me.btninspectorShiftLeft.Location = New System.Drawing.Point(368, 592)
        Me.btninspectorShiftLeft.Name = "btninspectorShiftLeft"
        Me.btninspectorShiftLeft.Size = New System.Drawing.Size(32, 23)
        Me.btninspectorShiftLeft.TabIndex = 0
        Me.btninspectorShiftLeft.Tag = ""
        Me.btninspectorShiftLeft.Text = "<"
        '
        'btninspectorShiftRightAll
        '
        Me.btninspectorShiftRightAll.Location = New System.Drawing.Point(368, 544)
        Me.btninspectorShiftRightAll.Name = "btninspectorShiftRightAll"
        Me.btninspectorShiftRightAll.Size = New System.Drawing.Size(32, 23)
        Me.btninspectorShiftRightAll.TabIndex = 0
        Me.btninspectorShiftRightAll.Text = ">>"
        '
        'btninspectorShiftRight
        '
        Me.btninspectorShiftRight.Location = New System.Drawing.Point(368, 520)
        Me.btninspectorShiftRight.Name = "btninspectorShiftRight"
        Me.btninspectorShiftRight.Size = New System.Drawing.Size(32, 23)
        Me.btninspectorShiftRight.TabIndex = 0
        Me.btninspectorShiftRight.Text = ">"
        '
        'ugAvailableOwnersFacCount
        '
        Me.ugAvailableOwnersFacCount.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugAvailableOwnersFacCount.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True
        Me.ugAvailableOwnersFacCount.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugAvailableOwnersFacCount.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugAvailableOwnersFacCount.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugAvailableOwnersFacCount.Location = New System.Drawing.Point(8, 320)
        Me.ugAvailableOwnersFacCount.Name = "ugAvailableOwnersFacCount"
        Me.ugAvailableOwnersFacCount.Size = New System.Drawing.Size(344, 144)
        Me.ugAvailableOwnersFacCount.TabIndex = 11
        '
        'lblAvailableCounties
        '
        Me.lblAvailableCounties.Location = New System.Drawing.Point(16, 56)
        Me.lblAvailableCounties.Name = "lblAvailableCounties"
        Me.lblAvailableCounties.Size = New System.Drawing.Size(104, 16)
        Me.lblAvailableCounties.TabIndex = 266
        Me.lblAvailableCounties.Text = "Available Counties"
        '
        'lblAssignedCounties
        '
        Me.lblAssignedCounties.Location = New System.Drawing.Point(416, 56)
        Me.lblAssignedCounties.Name = "lblAssignedCounties"
        Me.lblAssignedCounties.Size = New System.Drawing.Size(104, 16)
        Me.lblAssignedCounties.TabIndex = 267
        Me.lblAssignedCounties.Text = "Assigned Counties"
        '
        'lblAvailableOwnersFacilityCount
        '
        Me.lblAvailableOwnersFacilityCount.Location = New System.Drawing.Point(16, 296)
        Me.lblAvailableOwnersFacilityCount.Name = "lblAvailableOwnersFacilityCount"
        Me.lblAvailableOwnersFacilityCount.Size = New System.Drawing.Size(272, 16)
        Me.lblAvailableOwnersFacilityCount.TabIndex = 268
        Me.lblAvailableOwnersFacilityCount.Text = "Available Owners in Conflict - Facility Count"
        '
        'lblAssignedOwnersFacilityCount
        '
        Me.lblAssignedOwnersFacilityCount.Location = New System.Drawing.Point(416, 296)
        Me.lblAssignedOwnersFacilityCount.Name = "lblAssignedOwnersFacilityCount"
        Me.lblAssignedOwnersFacilityCount.Size = New System.Drawing.Size(280, 16)
        Me.lblAssignedOwnersFacilityCount.TabIndex = 269
        Me.lblAssignedOwnersFacilityCount.Text = "Assigned Owners to Managers  - Facility Count"
        '
        'chkInactive
        '
        Me.chkInactive.Location = New System.Drawing.Point(320, 48)
        Me.chkInactive.Name = "chkInactive"
        Me.chkInactive.Size = New System.Drawing.Size(72, 24)
        Me.chkInactive.TabIndex = 2
        Me.chkInactive.Text = "Inactive"
        Me.chkInactive.Visible = False
        '
        'lblOwnerFacilities
        '
        Me.lblOwnerFacilities.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOwnerFacilities.Location = New System.Drawing.Point(0, 464)
        Me.lblOwnerFacilities.Name = "lblOwnerFacilities"
        Me.lblOwnerFacilities.Size = New System.Drawing.Size(88, 24)
        Me.lblOwnerFacilities.TabIndex = 272
        Me.lblOwnerFacilities.Text = "Owner Count ="
        Me.lblOwnerFacilities.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(672, 672)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(80, 23)
        Me.btnCancel.TabIndex = 17
        Me.btnCancel.Text = "Cancel"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(576, 672)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(80, 23)
        Me.btnSave.TabIndex = 16
        Me.btnSave.Text = "Save"
        '
        'lblCountiesSumValue
        '
        Me.lblCountiesSumValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCountiesSumValue.Location = New System.Drawing.Point(456, 232)
        Me.lblCountiesSumValue.Name = "lblCountiesSumValue"
        Me.lblCountiesSumValue.Size = New System.Drawing.Size(64, 16)
        Me.lblCountiesSumValue.TabIndex = 250
        '
        'lblCountiesSum
        '
        Me.lblCountiesSum.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCountiesSum.Location = New System.Drawing.Point(416, 232)
        Me.lblCountiesSum.Name = "lblCountiesSum"
        Me.lblCountiesSum.Size = New System.Drawing.Size(40, 16)
        Me.lblCountiesSum.TabIndex = 250
        Me.lblCountiesSum.Text = "SUM ="
        Me.lblCountiesSum.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtOwnerFacilities
        '
        Me.txtOwnerFacilities.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOwnerFacilities.Location = New System.Drawing.Point(96, 464)
        Me.txtOwnerFacilities.Name = "txtOwnerFacilities"
        Me.txtOwnerFacilities.Size = New System.Drawing.Size(104, 24)
        Me.txtOwnerFacilities.TabIndex = 272
        Me.txtOwnerFacilities.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblAvailCountiesSumValue
        '
        Me.lblAvailCountiesSumValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAvailCountiesSumValue.Location = New System.Drawing.Point(48, 232)
        Me.lblAvailCountiesSumValue.Name = "lblAvailCountiesSumValue"
        Me.lblAvailCountiesSumValue.Size = New System.Drawing.Size(64, 16)
        Me.lblAvailCountiesSumValue.TabIndex = 276
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(8, 232)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(40, 16)
        Me.Label2.TabIndex = 275
        Me.Label2.Text = "SUM ="
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'cmbInspectorActive
        '
        Me.cmbInspectorActive.Location = New System.Drawing.Point(104, 0)
        Me.cmbInspectorActive.Name = "cmbInspectorActive"
        Me.cmbInspectorActive.Size = New System.Drawing.Size(192, 21)
        Me.cmbInspectorActive.TabIndex = 0
        Me.cmbInspectorActive.Visible = False
        '
        'lblActive
        '
        Me.lblActive.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblActive.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblActive.Location = New System.Drawing.Point(304, 26)
        Me.lblActive.Name = "lblActive"
        Me.lblActive.Size = New System.Drawing.Size(88, 16)
        Me.lblActive.TabIndex = 272
        Me.lblActive.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'cboManagers
        '
        Me.cboManagers.Location = New System.Drawing.Point(104, 264)
        Me.cboManagers.Name = "cboManagers"
        Me.cboManagers.Size = New System.Drawing.Size(512, 21)
        Me.cboManagers.TabIndex = 279
        '
        'lblManager
        '
        Me.lblManager.Location = New System.Drawing.Point(32, 264)
        Me.lblManager.Name = "lblManager"
        Me.lblManager.Size = New System.Drawing.Size(64, 16)
        Me.lblManager.TabIndex = 280
        Me.lblManager.Text = "Manager:"
        Me.lblManager.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(416, 496)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(280, 16)
        Me.Label3.TabIndex = 288
        Me.Label3.Text = "Assigned Inspectors to Managers"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(16, 496)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(272, 16)
        Me.Label4.TabIndex = 287
        Me.Label4.Text = "Available Inspectors with Open OCE's"
        '
        'ugAssignedInspectorsToManagers
        '
        Me.ugAssignedInspectorsToManagers.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugAssignedInspectorsToManagers.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugAssignedInspectorsToManagers.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugAssignedInspectorsToManagers.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugAssignedInspectorsToManagers.Location = New System.Drawing.Point(416, 520)
        Me.ugAssignedInspectorsToManagers.Name = "ugAssignedInspectorsToManagers"
        Me.ugAssignedInspectorsToManagers.Size = New System.Drawing.Size(344, 144)
        Me.ugAssignedInspectorsToManagers.TabIndex = 286
        '
        'ugAvailableInspectors
        '
        Me.ugAvailableInspectors.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugAvailableInspectors.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.True
        Me.ugAvailableInspectors.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugAvailableInspectors.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugAvailableInspectors.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugAvailableInspectors.Location = New System.Drawing.Point(8, 520)
        Me.ugAvailableInspectors.Name = "ugAvailableInspectors"
        Me.ugAvailableInspectors.Size = New System.Drawing.Size(344, 144)
        Me.ugAvailableInspectors.TabIndex = 281
        '
        'Inspectors
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(768, 702)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.ugAssignedInspectorsToManagers)
        Me.Controls.Add(Me.ugAvailableInspectors)
        Me.Controls.Add(Me.cboManagers)
        Me.Controls.Add(Me.lblManager)
        Me.Controls.Add(Me.lblAvailCountiesSumValue)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.lblOwnerFacilities)
        Me.Controls.Add(Me.chkInactive)
        Me.Controls.Add(Me.lblAssignedOwnersFacilityCount)
        Me.Controls.Add(Me.lblAvailableOwnersFacilityCount)
        Me.Controls.Add(Me.lblAssignedCounties)
        Me.Controls.Add(Me.lblAvailableCounties)
        Me.Controls.Add(Me.ugAssignedOwnersFacCount)
        Me.Controls.Add(Me.BtnOwnerShiftLeftAll)
        Me.Controls.Add(Me.btnOwnerShiftLeft)
        Me.Controls.Add(Me.BtnOwnerShiftRightAll)
        Me.Controls.Add(Me.BtnOwnerShiftRight)
        Me.Controls.Add(Me.btninspectorShiftLeftAll)
        Me.Controls.Add(Me.btninspectorShiftLeft)
        Me.Controls.Add(Me.btninspectorShiftRightAll)
        Me.Controls.Add(Me.btninspectorShiftRight)
        Me.Controls.Add(Me.ugAvailableOwnersFacCount)
        Me.Controls.Add(Me.ugAssignedCounties)
        Me.Controls.Add(Me.btnCountyShiftLeftAll)
        Me.Controls.Add(Me.btnCountyShiftLeft)
        Me.Controls.Add(Me.btnCountyShiftRightAll)
        Me.Controls.Add(Me.btnCountyShiftRight)
        Me.Controls.Add(Me.ugAvailableCounties)
        Me.Controls.Add(Me.cmbInspector)
        Me.Controls.Add(Me.lblInspector)
        Me.Controls.Add(Me.lblCountiesSumValue)
        Me.Controls.Add(Me.lblCountiesSum)
        Me.Controls.Add(Me.txtOwnerFacilities)
        Me.Controls.Add(Me.cmbInspectorActive)
        Me.Controls.Add(Me.lblActive)
        Me.Name = "Inspectors"
        Me.Text = "C and E Assignment Manager"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.ugAvailableCounties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ugAssignedCounties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ugAssignedOwnersFacCount, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ugAvailableOwnersFacCount, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ugAssignedInspectorsToManagers, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ugAvailableInspectors, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region "Form Load Events"
    Private Sub Inspectors_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim drRow, drRow1 As DataRow
        Try


            btnOwnerShiftLeft.Enabled = False
            BtnOwnerShiftLeftAll.Enabled = False
            BtnOwnerShiftRight.Enabled = False
            BtnOwnerShiftRightAll.Enabled = False

            btninspectorShiftLeft.Enabled = False
            btninspectorShiftLeftAll.Enabled = False
            btninspectorShiftRight.Enabled = False
            btninspectorShiftRightAll.Enabled = False


            cmbInspector.DisplayMember = "USER_NAME"
            cmbInspector.ValueMember = "STAFF_ID"
            cmbInspector.DataSource = oInspectorCounty.getInspectors


            cmbInspectorActive.DisplayMember = "ACTIVE"
            cmbInspectorActive.ValueMember = "STAFF_ID"
            cmbInspectorActive.DataSource = oInspectorCounty.getInspectors

            cmbInspectorActive.SelectedIndex = cmbInspector.SelectedIndex
            If cmbInspectorActive.Text.ToUpper = "TRUE" Then
                lblActive.Text = "Inactive"
                'chkInactive.Checked = True
            Else
                lblActive.Text = "Active"
                'chkInactive.Checked = False
            End If

            cboManagers.DisplayMember = "Description"
            cboManagers.ValueMember = "STAFF_ID"
            cboManagers.DataSource = oInspectorOwner.getCNEManagers

            getTotalFacilities()
            getOwnerFacilities()



            setSaveEnabled()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "Get County and Owner Facilities"
    Private Sub getSummaryValue(ByVal uGrid As Infragistics.Win.UltraWinGrid.UltraGrid, ByVal lbl As Label)
        If uGrid.Rows.SummaryValues.Count > 0 Then
            lbl.Text = CType(uGrid.Rows.SummaryValues.Item(0).Value, Integer).ToString
        Else
            lbl.Text = 0
        End If
    End Sub

    Private Sub getCountyFacilities()
        Try
            'Get the assigned counties
            oInspectorCounty.GetAll(nOwnerID)
            dtAssCountyOwners = oInspectorCounty.EntityTable()
            ugAssignedCounties.DataSource = Nothing
            ugAssignedCounties.DataSource = dtAssCountyOwners

            'Get the available counties
            dtAVailableCountyFacilities = oInspectorCounty.getAvailableCountyFacilities(nOwnerID)
            ugAvailableCounties.DataSource = Nothing
            ugAvailableCounties.DataSource = dtAVailableCountyFacilities
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub getOwnerFacilities()
        Try
            'get the assigned Owners
            oInspectorOwner.GetAll(nManagerID)
            dtAssInspectorOwners = oInspectorOwner.EntityTable()
            ugAssignedOwnersFacCount.DataSource = Nothing
            ugAssignedOwnersFacCount.DataSource = dtAssInspectorOwners

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub getTotalFacilitiesForManager()

        Try
            dtAvailableOwnerFacilities = oInspectorOwner.getOwnersInConflictOfManagerTerritory(nManagerID)
            ugAvailableOwnersFacCount.DataSource = Nothing
            ugAvailableOwnersFacCount.DataSource = Me.dtAvailableOwnerFacilities
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub


    Private Sub getTotalInspectorsForManager()

        Try
            dtAssignedInspectors = oInspectorOwner.getInspectorsUnderManager(nManagerID)
            ugAssignedInspectorsToManagers.DataSource = Nothing
            ugAssignedInspectorsToManagers.DataSource = Me.dtAssignedInspectors

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub



    Private Sub getAvailableInspectorsForManagers()

        Try
            dtAvailableInspectors = oInspectorOwner.getAvailableInspectorsForManagement
            ugAvailableInspectors.DataSource = Nothing
            ugAvailableInspectors.DataSource = Me.dtAvailableInspectors
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub
    Private Sub getTotalFacilities()
        Dim drRow As DataRow
        Dim totalFacilities As Integer = 0
        Try

            If ugAvailableOwnersFacCount.Rows.Count > 0 Then
                txtOwnerFacilities.Text = Me.ugAvailableOwnersFacCount.Rows.Count
            Else
                txtOwnerFacilities.Text = "0"
            End If


        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "Button Events"
    Private Sub btnCountyShiftRight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCountyShiftRight.Click
        Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            If Not Me.ugAvailableCounties.ActiveRow Is Nothing Then
                ugRow = ugAvailableCounties.ActiveRow
                AssignedCountyInspections(ugRow)
                ugAssignedCounties.DataSource = dtAssCountyOwners
                getSummaryValue(ugAssignedCounties, lblCountiesSumValue)

                ugAvailableCounties.ActiveRow.Delete(False)
                ugAvailableCounties.Refresh()
                getSummaryValue(ugAvailableCounties, lblAvailCountiesSumValue)

                getTotalFacilities()
                setSaveEnabled(True)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnCountyShiftRightAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCountyShiftRightAll.Click
        Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            If Not ugAvailableCounties.Rows.Count = 0 Then
                For Each ugRow In ugAvailableCounties.Rows
                    AssignedCountyInspections(ugRow)
                Next
                dtAssCountyOwners.DefaultView.Sort = "COUNTY"
                dtAVailableCountyFacilities.Clear()
                ugAssignedCounties.DataSource = Nothing
                ugAssignedCounties.DataSource = dtAssCountyOwners
                getSummaryValue(ugAssignedCounties, lblCountiesSumValue)

                ugAvailableCounties.DataSource = dtAVailableCountyFacilities
                getSummaryValue(ugAvailableCounties, lblAvailCountiesSumValue)

                firstTimeCounty = False
                getTotalFacilities()
                setSaveEnabled(True)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnCountyShiftLeft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCountyShiftLeft.Click
        Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            If Not Me.ugAssignedCounties.ActiveRow Is Nothing Then
                ugRow = ugAssignedCounties.ActiveRow
                ugAvailableCounties.DataSource = AvailableCountyInspections(ugRow)
                getSummaryValue(ugAvailableCounties, lblAvailCountiesSumValue)

                ugAssignedCounties.ActiveRow.Delete(False)
                ugAssignedCounties.Refresh()
                getSummaryValue(ugAssignedCounties, lblCountiesSumValue)
                firstTimeCounty = False
                getTotalFacilities()
                setSaveEnabled(True)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnCountyShiftLeftAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCountyShiftLeftAll.Click
        Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            If Not ugAssignedCounties.Rows.Count = 0 Then
                For Each ugRow In ugAssignedCounties.Rows
                    AvailableCountyInspections(ugRow)
                Next
                dtAVailableCountyFacilities.DefaultView.Sort = "COUNTY"
                dtAssCountyOwners.Clear()
                ugAvailableCounties.DataSource = Nothing
                ugAvailableCounties.DataSource = dtAVailableCountyFacilities
                getSummaryValue(ugAvailableCounties, lblAvailCountiesSumValue)

                ugAssignedCounties.DataSource = dtAssCountyOwners
                getSummaryValue(ugAssignedCounties, lblCountiesSumValue)
                getTotalFacilities()
                setSaveEnabled(True)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnOwnerShiftRight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOwnerShiftRight.Click
        Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            If Not Me.ugAvailableOwnersFacCount.ActiveRow Is Nothing Then
                ugRow = ugAvailableOwnersFacCount.ActiveRow
                AssignedOwnerInspections(ugRow)
                ugAssignedOwnersFacCount.DataSource = dtAssInspectorOwners

                ugAvailableOwnersFacCount.ActiveRow.Delete(False)
                ugAvailableOwnersFacCount.Refresh()

                getTotalFacilities()
                setSaveEnabled(True)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnInspectorShiftRight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btninspectorShiftRight.Click
        Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            If Not Me.ugAvailableInspectors.ActiveRow Is Nothing Then
                ugRow = ugAvailableInspectors.ActiveRow
                AssignedInspectors(ugRow)
                ugAssignedInspectorsToManagers.DataSource = dtAssignedInspectors

                ugAvailableInspectors.ActiveRow.Delete(False)
                ugAvailableInspectors.Refresh()

                setSaveEnabled(True)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnOwnerShiftRightAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOwnerShiftRightAll.Click
        Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            If Not ugAvailableOwnersFacCount.Rows.Count = 0 Then
                For Each ugRow In ugAvailableOwnersFacCount.Rows
                    AssignedOwnerInspections(ugRow)
                Next
                dtAssInspectorOwners.DefaultView.Sort = "Owner"
                dtAvailableOwnerFacilities.Clear()
                ugAssignedOwnersFacCount.DataSource = Nothing
                ugAssignedOwnersFacCount.DataSource = dtAssInspectorOwners

                ugAvailableOwnersFacCount.DataSource = dtAvailableOwnerFacilities
                firstTimeOwner = False
                getTotalFacilities()
                setSaveEnabled(True)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub


    Private Sub btnOwnerInspectorRightAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btninspectorShiftRightAll.Click
        Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            If Not ugAvailableInspectors.Rows.Count = 0 Then
                For Each ugRow In ugAvailableInspectors.Rows
                    Me.AssignedInspectors(ugRow)
                Next

                dtAssignedInspectors.DefaultView.Sort = "Inspector Name"

                dtAvailableInspectors.Clear()

                ugAssignedInspectorsToManagers.DataSource = Nothing
                ugAssignedInspectorsToManagers.DataSource = dtAssignedInspectors

                ugAvailableInspectors.DataSource = Me.dtAvailableInspectors

                setSaveEnabled(True)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub


    Private Sub btnOwnerShiftLeft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOwnerShiftLeft.Click
        Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            If Not Me.ugAssignedOwnersFacCount.ActiveRow Is Nothing Then
                ugRow = ugAssignedOwnersFacCount.ActiveRow
                ugAvailableOwnersFacCount.DataSource = AvailableOwnerInspections(ugRow)

                ugAssignedOwnersFacCount.ActiveRow.Delete(False)
                ugAssignedOwnersFacCount.Refresh()

                firstTimeOwner = False
                getTotalFacilities()
                setSaveEnabled(True)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnInspectorShiftLeft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btninspectorShiftLeft.Click
        Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            If Not ugAssignedInspectorsToManagers.ActiveRow Is Nothing Then
                ugRow = ugAssignedInspectorsToManagers.ActiveRow
                ugAvailableInspectors.DataSource = AvailableInspectors(ugRow)

                ugAssignedInspectorsToManagers.ActiveRow.Delete(False)
                ugAssignedInspectorsToManagers.Refresh()

                setSaveEnabled(True)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub


    Private Sub btnInspectorsShiftLeftAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btninspectorShiftLeftAll.Click
        Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            If Not Me.ugAssignedInspectorsToManagers.Rows.Count = 0 Then
                For Each ugRow In Me.ugAssignedInspectorsToManagers.Rows
                    AvailableInspectors(ugRow)
                Next
                dtAssignedInspectors.Clear()

                ugAssignedInspectorsToManagers.DataSource = Nothing
                ugAssignedInspectorsToManagers.DataSource = dtAssignedInspectors

                Me.ugAvailableInspectors.DataSource = Nothing
                ugAvailableInspectors.DataSource = Me.dtAvailableInspectors


                setSaveEnabled(True)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnOwnerShiftLeftAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnOwnerShiftLeftAll.Click
        Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            If Not ugAssignedOwnersFacCount.Rows.Count = 0 Then
                For Each ugRow In ugAssignedOwnersFacCount.Rows
                    AvailableOwnerInspections(ugRow)
                Next
                dtAssInspectorOwners.Clear()




                getTotalFacilities()

                setSaveEnabled(True)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try

            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
            oInspectorCounty.Flush(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal)

            If Not MusterContainer.AppUser.HEAD_CANDE AndAlso MusterContainer.AppUser.ID.ToUpper <> "ADMIN" Then
                returnVal = "User must be the Head of Compliance & Enforcement to set this change"
            End If
            If Not UIUtilsGen.HasRights(returnVal) Then
                Me.Cursor = System.Windows.Forms.Cursors.Default
                Exit Sub
            End If
            oInspectorOwner.Flush(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Me.Cursor = System.Windows.Forms.Cursors.Default
                Exit Sub
            End If


            For Each row As DataRow In Me.AddOnInspectors
                oInspectorOwner.AssignInspectorToManager(Me.nManagerID, row.Item("PropertyID"))
            Next

            Me.AddOnInspectors.Clear()

            For Each row As DataRow In Me.RemoveInspectors
                oInspectorOwner.RemoveInspectorFromManager(Me.nManagerID, row.Item("PropertyID"))
            Next

            Me.RemoveInspectors.Clear()

            Dim inx As Integer = cboManagers.SelectedIndex


            cboManagers.DisplayMember = "Description"
            cboManagers.ValueMember = "STAFF_ID"
            cboManagers.DataSource = oInspectorOwner.getCNEManagers

            cboManagers.SelectedIndex = inx

            Cursor = System.Windows.Forms.Cursors.Default
            MsgBox("Changes saved successfully")

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Dim result As DialogResult
        Try
            'result = MsgBox("Do you want to cancel the changes that you have done", MsgBoxStyle.YesNo)
            ' If result = DialogResult.Yes Then
            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
            getCountyFacilities()
            getOwnerFacilities()
            getTotalFacilities()
            Me.Cursor = System.Windows.Forms.Cursors.Default

            Me.Close()


            ' End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

#End Region
#Region "Change Events"

    Private Sub cmbmanager_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboManagers.SelectedIndexChanged
        Try
            nManagerID = cboManagers.SelectedValue

            If nManagerID > 0 Then

                'get the total number of facilities associated with an Owner
                getOwnerFacilities()
                getTotalFacilitiesForManager()
                getTotalFacilities()

                getAvailableInspectorsForManagers()
                getTotalInspectorsForManager()

                btnOwnerShiftLeft.Enabled = True
                BtnOwnerShiftLeftAll.Enabled = True
                BtnOwnerShiftRight.Enabled = True
                BtnOwnerShiftRightAll.Enabled = True

                btninspectorShiftLeft.Enabled = True
                btninspectorShiftLeftAll.Enabled = True
                btninspectorShiftRight.Enabled = True
                btninspectorShiftRightAll.Enabled = True

                ugAvailableOwnersFacCount.Enabled = True
                ugAssignedOwnersFacCount.Enabled = True
                ugAvailableInspectors.Enabled = True
                ugAssignedInspectorsToManagers.Enabled = True


            Else
                getOwnerFacilities()
                getTotalFacilitiesForManager()
                getTotalFacilities()



                ugAvailableOwnersFacCount.Enabled = False
                ugAssignedOwnersFacCount.Enabled = False
                ugAvailableInspectors.Enabled = False
                ugAssignedInspectorsToManagers.Enabled = False


                btnOwnerShiftLeft.Enabled = False
                BtnOwnerShiftLeftAll.Enabled = False
                BtnOwnerShiftRight.Enabled = False
                BtnOwnerShiftRightAll.Enabled = False

                btninspectorShiftLeft.Enabled = False
                btninspectorShiftLeftAll.Enabled = False
                btninspectorShiftRight.Enabled = False
                btninspectorShiftRightAll.Enabled = False

            End If




        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub


    Private Sub cmbInspector_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbInspector.SelectedIndexChanged
        Try
            If cmbInspector.SelectedIndex = -1 Then
                chkInactive.Checked = False
                Exit Sub
            End If
            nOwnerID = cmbInspector.SelectedValue

            cmbInspectorActive.SelectedValue = cmbInspector.SelectedValue
            If cmbInspectorActive.Text.ToUpper = "TRUE" Then
                lblActive.Text = "Inactive"
                'chkInactive.Checked = True
            Else
                lblActive.Text = "Active"
                'chkInactive.Checked = False
            End If

            'Get the counties asociated with the Owner 
            getCountyFacilities()
            'Get the Owners associated with the Owner
            getOwnerFacilities()
            'get the total number of facilities associated with an Owner
            getTotalFacilities()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "Other Functions"
    Private Function AssignedCountyInspections(ByVal ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow) As DataTable
        Dim drRow As DataRow
        Dim oCountyInfo As MUSTER.Info.InspectorCountyAssociationInfo
        Try
            'If dtAssCountyOwners.Rows.Count = 0 And firstTimeCounty Then
            '    dtAssCountyOwners.Columns.Add("COUNTY")
            '    dtAssCountyOwners.Columns.Add("FIPS")
            '    dtAssCountyOwners.Columns.Add("FACILITIES")
            '    dtAssCountyOwners.Columns.Add("ID")
            '    firstTimeCounty = False
            'End If
            drRow = dtAssCountyOwners.NewRow
            drRow("COUNTY") = ugRow.Cells("COUNTY").Value
            drRow("FACILITIES") = Integer.Parse(ugRow.Cells("FACILITIES").Value)
            drRow("FIPS") = Integer.Parse(ugRow.Cells("FIPS").Value)

            'dtAVailableCountyFacilities.Rows.Remove(drRow)
            'dtAVailableCountyFacilities.DefaultView.Sort = "COUNTY"
            dtAssCountyOwners.DefaultView.Sort = "COUNTY"
            'Instantiate the info and add it to the collection
            oCountyInfo = New MUSTER.Info.InspectorCountyAssociationInfo(0, nOwnerID, drRow("FIPS"), _
                                                                         MusterContainer.AppUser.ID, CDate("01/01/0001"), _
                                                                        String.Empty, CDate("01/01/0001"), 0)
            drRow("ID") = oInspectorCounty.Add(oCountyInfo)
            oCountyInfo.County = drRow("COUNTY")
            oCountyInfo.Facilities = drRow("FACILITIES")
            dtAssCountyOwners.Rows.Add(drRow)
            Return dtAssCountyOwners
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function
    Private Function AssignedOwnerInspections(ByVal ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow) As DataTable
        Dim drRow As DataRow
        Dim oOwnerInfo As MUSTER.Info.InspectorOwnerAssignmentInfo
        Try
            'If dtAssInspectorOwners.Rows.Count = 0 And firstTimeOwner Then
            '    dtAssInspectorOwners.Columns.Add("Owner")
            '    dtAssInspectorOwners.Columns.Add("ID")
            '    dtAssInspectorOwners.Columns.Add("FACILITIES")
            '    firstTimeOwner = False
            'End If
            drRow = dtAssInspectorOwners.NewRow
            drRow("Owner") = ugRow.Cells("PropertyName").Value.ToString.Substring(0, ugRow.Cells("PropertyName").Value.ToString.IndexOf(":"))

            drRow("Owner_ID") = Integer.Parse(ugRow.Cells("PropertyID").Value)
            drRow("FACILITIES") = Integer.Parse(ugRow.Cells("PropertyName").Value.ToString.Substring(ugRow.Cells("PropertyName").Value.ToString.LastIndexOf(":") + 1)) * -1
            oOwnerInfo = New MUSTER.Info.InspectorOwnerAssignmentInfo(0, nManagerID, drRow("Owner_ID"), _
                                                                      MusterContainer.AppUser.ID, CDate("01/01/0001"), _
                                                                      String.Empty, CDate("01/01/0001"), 0)
            drRow("ID") = oInspectorOwner.Add(oOwnerInfo)
            oOwnerInfo.Owner = drRow("Owner")
            oOwnerInfo.Facilities = drRow("FACILITIES")
            dtAssInspectorOwners.Rows.Add(drRow)
            dtAssInspectorOwners.DefaultView.Sort = "Owner"
            Return dtAssInspectorOwners
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function


    Private Function AssignedInspectors(ByVal ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow) As DataTable
        Dim drRow As DataRow
        Dim drrRow As DataRow

        Dim oOwnerInfo As MUSTER.Info.InspectorOwnerAssignmentInfo
        Try

            drRow = Me.dtAssignedInspectors.NewRow

            drRow("Inspector Name") = ugRow.Cells("User_Name").Value.ToString
            drRow("Inspector ID") = Integer.Parse(ugRow.Cells("STAFF_ID").Value)



            If dtInspectorData.Rows.Count = 0 OrElse Not dtInspectorData.Select(String.Format("Staff_id = {0}", drRow("Inspector ID"))) Is Nothing Then
                drrRow = Me.dtInspectorData.NewRow

                drrRow("PropertyName") = ugRow.Cells("user_Name").Value.ToString
                drrRow("PropertyID") = Integer.Parse(ugRow.Cells("STAFF_ID").Value)

            Else
                drrRow = Me.dtInspectorData.Select(String.Format("Staff_id = {0}", drRow("Inspector ID")))(0)
            End If


            If RemoveInspectors.Contains(drrRow) Then
                RemoveInspectors.Remove(drrRow)
            Else
                AddOnInspectors.Add(drrRow)
            End If

            dtAssignedInspectors.Rows.Add(drRow)

            dtAssignedInspectors.DefaultView.Sort = "Inspector Name"
            Return dtAssignedInspectors

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function


    Private Function AvailableCountyInspections(ByVal ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow) As DataTable
        Try
            Dim drRow As DataRow
            If Not ugRow Is Nothing Then
                drRow = dtAVailableCountyFacilities.NewRow
                drRow("COUNTY") = ugRow.Cells("COUNTY").Value
                drRow("FACILITIES") = Integer.Parse(ugRow.Cells("FACILITIES").Value)
                drRow("FIPS") = Integer.Parse(ugRow.Cells("FIPS").Value)
                drRow("ID") = Integer.Parse(ugRow.Cells("ID").Value)
                dtAVailableCountyFacilities.Rows.Add(drRow)
                'dtAssCountyOwners.Rows.Remove(drRow)
                'dtAssCountyOwners.DefaultView.Sort() = "COUNTY"
                dtAVailableCountyFacilities.DefaultView.Sort = "COUNTY"
                If drRow("ID") <= 0 Then
                    oInspectorCounty.Remove(drRow("ID"))
                Else
                    oInspectorCounty.colInspectorCounties.Item(drRow("ID")).DELETED = 1
                End If
                Return dtAVailableCountyFacilities
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function
    Private Function AvailableOwnerInspections(ByVal ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow) As DataTable
        Try
            Dim drRow As DataRow
            Dim ID As Integer
            If Not ugRow Is Nothing Then
                ID = ugRow.Cells("ID").Value
                drRow = dtAvailableOwnerFacilities.NewRow

                drRow("PropertyName") = String.Format("{0}:      FACILITIES: +{1}", ugRow.Cells("Owner").Value, ugRow.Cells("FACILITIES").Value * -1)
                drRow("PropertyID") = Integer.Parse(ugRow.Cells("Owner_ID").Value)

                dtAvailableOwnerFacilities.Rows.Add(drRow)
                'dtAssInspectorOwners.Rows.Remove(drRow)
                'dtAssInspectorOwners.DefaultView.Sort = "Owner"
                dtAvailableOwnerFacilities.DefaultView.Sort = "PropertyName"

                If ID <= 0 Then
                    oInspectorOwner.Remove(ID)
                Else
                    oInspectorOwner.colInspectorOwners.Item(ID.ToString).DELETED = 1
                End If

                Return dtAvailableOwnerFacilities
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function


    Private Function AvailableInspectors(ByVal ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow) As DataTable
        Try
            Dim drRow As DataRow
            Dim drrRow As DataRow

            Dim ID As Integer

            If Not ugRow Is Nothing Then

                drRow = Me.dtAvailableInspectors.NewRow

                drRow("USER_NAME") = ugRow.Cells("Inspector Name").Value.ToString
                drRow("STAFF_ID") = Integer.Parse(ugRow.Cells("Inspector ID").Value)

                If dtInspectorData.Rows.Count = 0 OrElse Not dtInspectorData.Select(String.Format("Staff_id = {0}", drRow("STAFF_ID"))) Is Nothing Then
                    drrRow = Me.dtInspectorData.NewRow

                    drrRow("PropertyName") = ugRow.Cells("Inspector Name").Value.ToString
                    drrRow("PropertyID") = Integer.Parse(ugRow.Cells("Inspector ID").Value)

                Else
                    drrRow = Me.dtInspectorData.Select(String.Format("Staff_id = {0}", drRow("STAFF_ID")))(0)
                End If



                If AddOnInspectors.Contains(drrRow) Then
                    AddOnInspectors.Remove(drrRow)
                Else
                    RemoveInspectors.Add(drrRow)
                End If

                dtAvailableInspectors.Rows.Add(drRow)

                dtAvailableInspectors.DefaultView.Sort = "USER_NAME"

                Return dtAvailableInspectors

            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Function

    Private Function setSaveEnabled(Optional ByVal bolEnable As Boolean = False)
        btnSave.Enabled = bolEnable
    End Function
#End Region
#Region "Ultragrid InitialiazeLayout"
    Private Sub ugAssignedCounties_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugAssignedCounties.InitializeLayout
        '   Set up the grid to display row summaries for Facilities columns
        e.Layout.Override.AllowRowSummaries = Infragistics.Win.UltraWinGrid.AllowRowSummaries.False
        If e.Layout.Bands(0).Columns.Exists("Facilities") Then
            e.Layout.Bands(0).Summaries.Add(Infragistics.Win.UltraWinGrid.SummaryType.Sum, e.Layout.Bands(0).Columns("Facilities"), Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn)
        End If
        getSummaryValue(sender, lblCountiesSumValue)
        e.Layout.Bands(0).Columns("ID").Hidden = True
        e.Layout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
    End Sub

    Private Sub ugAssignedOwnersFacCount_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugAssignedOwnersFacCount.InitializeLayout
        '   Set up the grid to display row summaries for Facilities columns
        e.Layout.Override.AllowRowSummaries = Infragistics.Win.UltraWinGrid.AllowRowSummaries.False
        If e.Layout.Bands(0).Columns.Exists("Facilities") Then
            e.Layout.Bands(0).Summaries.Add(Infragistics.Win.UltraWinGrid.SummaryType.Sum, e.Layout.Bands(0).Columns("Facilities"), Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn)
        End If
        e.Layout.Bands(0).Columns("ID").Hidden = True
        e.Layout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
    End Sub

    Private Sub ugAvailableCounties_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugAvailableCounties.InitializeLayout
        '   Set up the grid to display row summaries for Facilities columns
        e.Layout.Override.AllowRowSummaries = Infragistics.Win.UltraWinGrid.AllowRowSummaries.False
        If e.Layout.Bands(0).Columns.Exists("Facilities") Then
            e.Layout.Bands(0).Summaries.Add(Infragistics.Win.UltraWinGrid.SummaryType.Sum, e.Layout.Bands(0).Columns("Facilities"), Infragistics.Win.UltraWinGrid.SummaryPosition.UseSummaryPositionColumn)
        End If
        getSummaryValue(sender, lblAvailCountiesSumValue)
        e.Layout.Bands(0).Columns("ID").Hidden = True
        e.Layout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
    End Sub

    Private Sub ugAvailableOwnersFacCount_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugAvailableOwnersFacCount.InitializeLayout
        '   Set up the grid to display row summaries for Facilities columns
        e.Layout.Override.AllowRowSummaries = Infragistics.Win.UltraWinGrid.AllowRowSummaries.False


        e.Layout.Bands(0).Columns("PropertyID").Hidden = True
        e.Layout.AutoFitColumns = True
        e.Layout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
    End Sub
#End Region

    Private Sub Owners_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Dim result As DialogResult
        Try
            If oInspectorCounty.colIsDirty Or oInspectorOwner.colIsDirty Then
                result = MsgBox("There are unsaved changes. Do you wish to save them", MsgBoxStyle.YesNoCancel)
                If result = DialogResult.Yes Then
                    oInspectorCounty.Flush(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        e.Cancel = True
                        Exit Sub
                    End If

                    oInspectorOwner.Flush(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        e.Cancel = True
                        Exit Sub
                    End If

                ElseIf result = DialogResult.Cancel Then
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub Owners_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        Try
            MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "CandEManagement")
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub Owners_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.TextChanged
        Try
            MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "WindowName", Me.Text, "CandEManagement")
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnOwnerShiftLeft_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOwnerShiftLeft.Click

    End Sub
End Class
