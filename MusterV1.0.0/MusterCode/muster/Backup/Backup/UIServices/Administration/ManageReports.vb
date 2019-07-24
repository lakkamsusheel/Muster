Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class ManageReports
    Inherits System.Windows.Forms.Form
    '-------------------------------------------------------------------------------
    ' MUSTER.MUSTER.UserReports
    '   Provides the UI for managing reports
    '
    ' Copyright (C) 2004 CIBER, Inc.
    ' All rights reserved.
    '
    ' Release   Initials    Date        Description
    '  1.0        ??      12/??/04    Original class definition.
    '  1.1        JC      01/03/05    Added code to handle MDI interface for app
    '  1.2        AN      01/05/05    Added Event Handler for ReportChanged and formated functions and events into regions
    '  1.3        AN      01/10/05    minor tweaks to allow events to function properly
    '  1.4        PN      01/12/05    Removed chkRptShowDeleted_CheckedChanged and added chkRptShowDeleted_Click event(bug 640) .
    '                                  Added code for Isdirty in OnClosing event(bug 637)  
    '                                 Removed loadReportNames() that is called after Reports.save() from btnSave_Click (bug 641) 
    '  1.5        JVC2    01/12/05    Added TextChanged event handling.
    '  1.6        PN      01/24/05     Set Controls tab order(bug 661) 
    '  1.7        AN      02/10/05    Integrated AppFlags new object model
    '-------------------------------------------------------------------------------
    '
    'TODO - Remove comment from VSS version 2/9/05 - JVC 2
#Region "Private Member Variables"
    Public WithEvents Reports As Muster.BusinessLogic.pReport
    Private dt As DataTable
    Private bolLoading As Boolean = False
    Private bolIsNewReport As Boolean = False
    Private bolNameLeave As Boolean = False
    Private strPreviousReportName As String = String.Empty
    Private strReportName As String = String.Empty
    Private bolShowDeleted As Boolean = False
    Private itemExists As Boolean
    Friend MyGUID As New System.Guid
    Dim returnVal As String = String.Empty
    Dim reportGroupRelInfo As MUSTER.Info.ReportGroupRelationInfo
    Dim dtAvailable, dtAccess As DataTable
    Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
    Dim dr As DataRow
    Protected REPORT_PATH As String = MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_Reports).ProfileValue & "\"
#End Region
#Region " Windows Form Designer generated code "

    Public Sub New(Optional ByRef frm As Windows.Forms.Form = Nothing)
        MyBase.New()
        bolLoading = True

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        MyGUID = System.Guid.NewGuid
        MusterContainer.AppUser.LogEntry(Me.Text, MyGUID.ToString)
        MusterContainer.AppSemaphores.Retrieve(MyGUID.ToString, "WindowName", Me.Text)
        MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGUID)
        If Not frm Is Nothing Then
            If frm.IsMdiContainer Then
                Me.MdiParent = frm
            End If
        End If
        InitDataTable()

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
    Friend WithEvents lblDescription As System.Windows.Forms.Label
    Friend WithEvents txtDescription As System.Windows.Forms.TextBox
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents lblName As System.Windows.Forms.Label
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents cboModule As System.Windows.Forms.ComboBox
    Friend WithEvents lblModule As System.Windows.Forms.Label
    Friend WithEvents lblGroupswithAccess As System.Windows.Forms.Label
    Friend WithEvents lblAvailableGroups As System.Windows.Forms.Label
    Friend WithEvents btnShiftLeftAll As System.Windows.Forms.Button
    Friend WithEvents btnShiftLeft As System.Windows.Forms.Button
    Friend WithEvents btnShiftRightAll As System.Windows.Forms.Button
    Friend WithEvents btnShiftRight As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnSearchReport As System.Windows.Forms.Button
    Friend WithEvents lblReportFilePath As System.Windows.Forms.Label
    Friend WithEvents txtReportFilePath As System.Windows.Forms.TextBox
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents chkInactive As System.Windows.Forms.CheckBox
    Friend WithEvents cboReportName As System.Windows.Forms.ComboBox
    Friend WithEvents chkRptShowDeleted As System.Windows.Forms.CheckBox
    Friend WithEvents btnReportParams As System.Windows.Forms.Button
    Friend WithEvents txtReportName As System.Windows.Forms.TextBox
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents lblReportsList As System.Windows.Forms.Label
    Friend WithEvents ugAvailableGroups As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents ugGroupwithAccess As Infragistics.Win.UltraWinGrid.UltraGrid
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnShiftLeftAll = New System.Windows.Forms.Button
        Me.btnShiftLeft = New System.Windows.Forms.Button
        Me.btnShiftRightAll = New System.Windows.Forms.Button
        Me.btnShiftRight = New System.Windows.Forms.Button
        Me.lblGroupswithAccess = New System.Windows.Forms.Label
        Me.lblAvailableGroups = New System.Windows.Forms.Label
        Me.lblDescription = New System.Windows.Forms.Label
        Me.txtDescription = New System.Windows.Forms.TextBox
        Me.btnSearch = New System.Windows.Forms.Button
        Me.lblName = New System.Windows.Forms.Label
        Me.txtName = New System.Windows.Forms.TextBox
        Me.cboModule = New System.Windows.Forms.ComboBox
        Me.lblModule = New System.Windows.Forms.Label
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnSearchReport = New System.Windows.Forms.Button
        Me.lblReportFilePath = New System.Windows.Forms.Label
        Me.txtReportFilePath = New System.Windows.Forms.TextBox
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.chkInactive = New System.Windows.Forms.CheckBox
        Me.cboReportName = New System.Windows.Forms.ComboBox
        Me.chkRptShowDeleted = New System.Windows.Forms.CheckBox
        Me.btnReportParams = New System.Windows.Forms.Button
        Me.txtReportName = New System.Windows.Forms.TextBox
        Me.btnDelete = New System.Windows.Forms.Button
        Me.lblReportsList = New System.Windows.Forms.Label
        Me.ugAvailableGroups = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.ugGroupwithAccess = New Infragistics.Win.UltraWinGrid.UltraGrid
        CType(Me.ugAvailableGroups, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ugGroupwithAccess, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnShiftLeftAll
        '
        Me.btnShiftLeftAll.BackColor = System.Drawing.SystemColors.Control
        Me.btnShiftLeftAll.Enabled = False
        Me.btnShiftLeftAll.Location = New System.Drawing.Point(216, 360)
        Me.btnShiftLeftAll.Name = "btnShiftLeftAll"
        Me.btnShiftLeftAll.Size = New System.Drawing.Size(32, 24)
        Me.btnShiftLeftAll.TabIndex = 14
        Me.btnShiftLeftAll.Text = "<<"
        '
        'btnShiftLeft
        '
        Me.btnShiftLeft.BackColor = System.Drawing.SystemColors.Control
        Me.btnShiftLeft.Enabled = False
        Me.btnShiftLeft.Location = New System.Drawing.Point(216, 336)
        Me.btnShiftLeft.Name = "btnShiftLeft"
        Me.btnShiftLeft.Size = New System.Drawing.Size(32, 24)
        Me.btnShiftLeft.TabIndex = 13
        Me.btnShiftLeft.Text = "<"
        '
        'btnShiftRightAll
        '
        Me.btnShiftRightAll.BackColor = System.Drawing.SystemColors.Control
        Me.btnShiftRightAll.Enabled = False
        Me.btnShiftRightAll.Location = New System.Drawing.Point(216, 304)
        Me.btnShiftRightAll.Name = "btnShiftRightAll"
        Me.btnShiftRightAll.Size = New System.Drawing.Size(32, 24)
        Me.btnShiftRightAll.TabIndex = 12
        Me.btnShiftRightAll.Text = ">>"
        '
        'btnShiftRight
        '
        Me.btnShiftRight.BackColor = System.Drawing.SystemColors.Control
        Me.btnShiftRight.Enabled = False
        Me.btnShiftRight.Location = New System.Drawing.Point(216, 280)
        Me.btnShiftRight.Name = "btnShiftRight"
        Me.btnShiftRight.Size = New System.Drawing.Size(32, 24)
        Me.btnShiftRight.TabIndex = 11
        Me.btnShiftRight.Text = ">"
        '
        'lblGroupswithAccess
        '
        Me.lblGroupswithAccess.Location = New System.Drawing.Point(272, 256)
        Me.lblGroupswithAccess.Name = "lblGroupswithAccess"
        Me.lblGroupswithAccess.Size = New System.Drawing.Size(168, 16)
        Me.lblGroupswithAccess.TabIndex = 171
        Me.lblGroupswithAccess.Text = "Groups with Access To Report"
        '
        'lblAvailableGroups
        '
        Me.lblAvailableGroups.Location = New System.Drawing.Point(32, 256)
        Me.lblAvailableGroups.Name = "lblAvailableGroups"
        Me.lblAvailableGroups.Size = New System.Drawing.Size(104, 16)
        Me.lblAvailableGroups.TabIndex = 169
        Me.lblAvailableGroups.Text = "Available Groups"
        '
        'lblDescription
        '
        Me.lblDescription.Location = New System.Drawing.Point(32, 112)
        Me.lblDescription.Name = "lblDescription"
        Me.lblDescription.Size = New System.Drawing.Size(64, 20)
        Me.lblDescription.TabIndex = 159
        Me.lblDescription.Text = "Description:"
        Me.lblDescription.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtDescription
        '
        Me.txtDescription.Location = New System.Drawing.Point(128, 112)
        Me.txtDescription.MaxLength = 50
        Me.txtDescription.Multiline = True
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(280, 56)
        Me.txtDescription.TabIndex = 7
        Me.txtDescription.Text = ""
        '
        'btnSearch
        '
        Me.btnSearch.BackColor = System.Drawing.SystemColors.Control
        Me.btnSearch.Enabled = False
        Me.btnSearch.Location = New System.Drawing.Point(392, 40)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(24, 24)
        Me.btnSearch.TabIndex = 2
        Me.btnSearch.Text = "?"
        '
        'lblName
        '
        Me.lblName.Location = New System.Drawing.Point(32, 40)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(56, 20)
        Me.lblName.TabIndex = 155
        Me.lblName.Text = "Name:"
        Me.lblName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(376, 128)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(168, 20)
        Me.txtName.TabIndex = 154
        Me.txtName.Text = ""
        Me.txtName.Visible = False
        '
        'cboModule
        '
        Me.cboModule.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboModule.DropDownWidth = 180
        Me.cboModule.ItemHeight = 13
        Me.cboModule.Location = New System.Drawing.Point(128, 184)
        Me.cboModule.Name = "cboModule"
        Me.cboModule.Size = New System.Drawing.Size(144, 21)
        Me.cboModule.TabIndex = 8
        '
        'lblModule
        '
        Me.lblModule.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblModule.Location = New System.Drawing.Point(32, 184)
        Me.lblModule.Name = "lblModule"
        Me.lblModule.Size = New System.Drawing.Size(88, 20)
        Me.lblModule.TabIndex = 181
        Me.lblModule.Text = "Module:"
        Me.lblModule.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(384, 416)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(80, 26)
        Me.btnClose.TabIndex = 20
        Me.btnClose.Text = "Close"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(296, 416)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(80, 26)
        Me.btnCancel.TabIndex = 19
        Me.btnCancel.Text = "Cancel"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(120, 416)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(80, 26)
        Me.btnSave.TabIndex = 17
        Me.btnSave.Text = "Save"
        '
        'btnNew
        '
        Me.btnNew.Location = New System.Drawing.Point(32, 416)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 16
        Me.btnNew.Text = "New"
        '
        'btnSearchReport
        '
        Me.btnSearchReport.BackColor = System.Drawing.SystemColors.Control
        Me.btnSearchReport.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSearchReport.Location = New System.Drawing.Point(392, 80)
        Me.btnSearchReport.Name = "btnSearchReport"
        Me.btnSearchReport.Size = New System.Drawing.Size(32, 24)
        Me.btnSearchReport.TabIndex = 5
        Me.btnSearchReport.Text = " ..."
        Me.btnSearchReport.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'lblReportFilePath
        '
        Me.lblReportFilePath.Location = New System.Drawing.Point(32, 80)
        Me.lblReportFilePath.Name = "lblReportFilePath"
        Me.lblReportFilePath.Size = New System.Drawing.Size(90, 20)
        Me.lblReportFilePath.TabIndex = 190
        Me.lblReportFilePath.Text = "Report Path\File:"
        Me.lblReportFilePath.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtReportFilePath
        '
        Me.txtReportFilePath.Location = New System.Drawing.Point(128, 80)
        Me.txtReportFilePath.Name = "txtReportFilePath"
        Me.txtReportFilePath.Size = New System.Drawing.Size(256, 20)
        Me.txtReportFilePath.TabIndex = 4
        Me.txtReportFilePath.Text = ""
        '
        'chkInactive
        '
        Me.chkInactive.Location = New System.Drawing.Point(448, 40)
        Me.chkInactive.Name = "chkInactive"
        Me.chkInactive.Size = New System.Drawing.Size(72, 24)
        Me.chkInactive.TabIndex = 3
        Me.chkInactive.Text = "Inactive"
        '
        'cboReportName
        '
        Me.cboReportName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboReportName.Location = New System.Drawing.Point(128, 8)
        Me.cboReportName.Name = "cboReportName"
        Me.cboReportName.Size = New System.Drawing.Size(256, 21)
        Me.cboReportName.TabIndex = 0
        '
        'chkRptShowDeleted
        '
        Me.chkRptShowDeleted.Location = New System.Drawing.Point(448, 8)
        Me.chkRptShowDeleted.Name = "chkRptShowDeleted"
        Me.chkRptShowDeleted.TabIndex = 6
        Me.chkRptShowDeleted.Text = "Show Deleted"
        '
        'btnReportParams
        '
        Me.btnReportParams.Location = New System.Drawing.Point(128, 224)
        Me.btnReportParams.Name = "btnReportParams"
        Me.btnReportParams.Size = New System.Drawing.Size(224, 23)
        Me.btnReportParams.TabIndex = 9
        Me.btnReportParams.Text = "Modify Report Parameters"
        '
        'txtReportName
        '
        Me.txtReportName.Location = New System.Drawing.Point(128, 40)
        Me.txtReportName.Name = "txtReportName"
        Me.txtReportName.ReadOnly = True
        Me.txtReportName.Size = New System.Drawing.Size(256, 20)
        Me.txtReportName.TabIndex = 1
        Me.txtReportName.Text = ""
        '
        'btnDelete
        '
        Me.btnDelete.Enabled = False
        Me.btnDelete.Location = New System.Drawing.Point(208, 416)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 18
        Me.btnDelete.Text = "Delete"
        '
        'lblReportsList
        '
        Me.lblReportsList.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblReportsList.Location = New System.Drawing.Point(32, 8)
        Me.lblReportsList.Name = "lblReportsList"
        Me.lblReportsList.Size = New System.Drawing.Size(88, 20)
        Me.lblReportsList.TabIndex = 191
        Me.lblReportsList.Text = "Reports List: "
        Me.lblReportsList.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ugAvailableGroups
        '
        Me.ugAvailableGroups.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugAvailableGroups.Location = New System.Drawing.Point(32, 272)
        Me.ugAvailableGroups.Name = "ugAvailableGroups"
        Me.ugAvailableGroups.Size = New System.Drawing.Size(152, 128)
        Me.ugAvailableGroups.TabIndex = 192
        '
        'ugGroupwithAccess
        '
        Me.ugGroupwithAccess.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugGroupwithAccess.Location = New System.Drawing.Point(272, 272)
        Me.ugGroupwithAccess.Name = "ugGroupwithAccess"
        Me.ugGroupwithAccess.Size = New System.Drawing.Size(152, 128)
        Me.ugGroupwithAccess.TabIndex = 192
        '
        'ManageReports
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(576, 462)
        Me.Controls.Add(Me.ugAvailableGroups)
        Me.Controls.Add(Me.lblReportsList)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.txtReportName)
        Me.Controls.Add(Me.txtReportFilePath)
        Me.Controls.Add(Me.txtDescription)
        Me.Controls.Add(Me.txtName)
        Me.Controls.Add(Me.btnReportParams)
        Me.Controls.Add(Me.chkRptShowDeleted)
        Me.Controls.Add(Me.cboReportName)
        Me.Controls.Add(Me.chkInactive)
        Me.Controls.Add(Me.btnSearchReport)
        Me.Controls.Add(Me.lblReportFilePath)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.cboModule)
        Me.Controls.Add(Me.lblModule)
        Me.Controls.Add(Me.btnShiftLeftAll)
        Me.Controls.Add(Me.btnShiftLeft)
        Me.Controls.Add(Me.btnShiftRightAll)
        Me.Controls.Add(Me.btnShiftRight)
        Me.Controls.Add(Me.lblGroupswithAccess)
        Me.Controls.Add(Me.lblAvailableGroups)
        Me.Controls.Add(Me.lblDescription)
        Me.Controls.Add(Me.btnSearch)
        Me.Controls.Add(Me.lblName)
        Me.Controls.Add(Me.ugGroupwithAccess)
        Me.Name = "ManageReports"
        Me.Text = "Manage Reports"
        CType(Me.ugAvailableGroups, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ugGroupwithAccess, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Form Level Events"
    Private Sub ManageReports_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            resetform()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            bolLoading = False
        End Try
    End Sub
    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)
        'P1-Added code to Check isDirty
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        If Not Reports Is Nothing Then
            If Reports.IsDirty Then
                Dim Results As Long = MsgBox("There are unsaved changes. Do you want to save changes before closing?", MsgBoxStyle.YesNoCancel)
                If Results = MsgBoxResult.Yes Then
                    If Reports.ID <= 0 Then
                        Reports.CreatedBy = MusterContainer.AppUser.ID
                    Else
                        Reports.ModifiedBy = MusterContainer.AppUser.ID
                    End If
                    Reports.Save(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If
                Else
                    If Results = MsgBoxResult.Cancel Then
                        e.Cancel = True
                        Exit Sub
                    End If
                End If
            End If
        End If
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Remove any values from the shared collection for this screen
        '
        MusterContainer.AppSemaphores.Remove(MyGUID.ToString)
        '
        ' Log the disposal of the form (exit from Registration form)
        '
        MusterContainer.AppUser.LogExit(MyGUID.ToString)

    End Sub
    Private Sub ManageReports_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Try
            'If Not rpt Is Nothing And Me.cboReportName.Text <> String.Empty Then
            If Not Reports Is Nothing And Me.cboReportName.Text <> String.Empty Then

                If Reports.IsDirty Then
                    Dim Results As Long = MsgBox("There are unsaved changes. Do you want to save changes before closing?", MsgBoxStyle.YesNo)
                    If Results = MsgBoxResult.No Then
                        Reports.Reset()
                    Else
                        If chkInactive.Checked = False Then
                            Reports.Deleted = False
                        End If
                        If Reports.ID <= 0 Then
                            Reports.CreatedBy = MusterContainer.AppUser.ID
                        Else
                            Reports.ModifiedBy = MusterContainer.AppUser.ID
                        End If
                        Reports.Save(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal)
                        If Not UIUtilsGen.HasRights(returnVal) Then
                            Exit Sub
                        End If
                    End If
                End If
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "UI Support Routines"
    Private Sub SetupGroupGrid(ByRef ug As Infragistics.Win.UltraWinGrid.UltraGrid)
        ug.DisplayLayout.Bands(0).Columns("GROUP_ID").Hidden = True
        ug.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        ug.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        ug.DisplayLayout.Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Free
        'ug.DisplayLayout.Bands(0).Columns("INACTIVE").Width = 60
        ug.DisplayLayout.Bands(0).Columns("INACTIVE").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
        'ug.DisplayLayout.Bands(0).Columns("GROUP").Width = 130
        ug.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.False
        ug.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
    End Sub
    Private Sub InitDataTable()
        dtAccess = New DataTable
        dtAvailable = New DataTable

        dtAccess.Columns.Add("GROUP_ID", GetType(Integer))
        dtAccess.Columns.Add("GROUP", GetType(String))
        dtAccess.Columns.Add("INACTIVE", GetType(Boolean))

        dtAvailable = dtAccess.Clone
    End Sub
    Private Sub SetSaveCancel(ByVal bolState As Boolean)
        btnSave.Enabled = bolState
        btnCancel.Enabled = bolState
    End Sub
    Function resetform()
        'Dim oModProp As New MUSTER.BusinessLogic.pPropertyType("Modules")
        Dim dtModules As DataTable
        Reports = New MUSTER.BusinessLogic.pReport
        Me.btnCancel.Enabled = False
        Me.btnSave.Enabled = False
        'Me.lstAvailableGroups.Items.Clear()
        'Me.lstGroupwithAccess.Items.Clear()
        Me.txtReportFilePath.Text = String.Empty
        Me.cboReportName.Text = String.Empty
        Me.txtDescription.Text = String.Empty
        Me.btnReportParams.Enabled = False
        Try
            dtModules = MusterContainer.AppUser.ListPrimaryModules
            dtModules.DefaultView.Sort = "PROPERTY_NAME"
            dtModules.DefaultView.RowFilter = "PROPERTY_ID NOT IN (1303,1311,1312)"
            Me.cboModule.DataSource = dtModules.DefaultView
            Me.cboModule.DisplayMember = "PROPERTY_NAME"
            Me.cboModule.ValueMember = "PROPERTY_ID"
            Me.cboModule.SelectedIndex = -1
            loadReportNames()
            Me.cboReportName.SelectedIndex = -1
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Sub loadReportNames()
        Try
            'TODO Add code back - Ask Adam what this is about JVC2 2/8/05
            'Me.cboReportName.Tag = Me.cboReportName.Text
            'rpt = New InfoRepository.Report


            Me.cboReportName.DisplayMember = "REPORT_NAME"
            Me.cboReportName.ValueMember = "Report_ID"
            Dim dt As DataTable = Reports.GetReportsForUser()
            Dim dr As DataRow
            dr = dt.NewRow
            If dt.Rows.Count > 0 Then
                dr("REPORT_NAME") = " - Please select a Report - "
            Else
                dr("REPORT_NAME") = " - No Reports - "
            End If
            dr("REPORT_ID") = 0
            dt.Rows.InsertAt(dr, 0)
            Me.cboReportName.DataSource = dt 'rpt.ListReportNames(bolShowDeleted)
            Me.cboReportName.SelectedIndex = -1
            'TODO Add this code back - Ask Adam what this is about JVC2 - 2/8/05
            'If Me.cboReportName.Tag <> String.Empty Then
            '    Me.cboReportName.SelectedValue = Me.cboReportName.Tag.ToString
            'End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub loadGrids()
        InitDataTable()
        For Each reportGroupRelInfo In Reports.ReportGroupRelationCollection.Values
            If reportGroupRelInfo.ReportID = 0 Or reportGroupRelInfo.Deleted Then
                dr = dtAvailable.NewRow
                dr("GROUP_ID") = reportGroupRelInfo.GroupID
                dr("GROUP") = reportGroupRelInfo.GroupName
                dr("INACTIVE") = reportGroupRelInfo.Active
                dtAvailable.Rows.Add(dr)
            Else
                dr = dtAccess.NewRow
                dr("GROUP_ID") = reportGroupRelInfo.GroupID
                dr("GROUP") = reportGroupRelInfo.GroupName
                dr("INACTIVE") = reportGroupRelInfo.Active
                dtAccess.Rows.Add(dr)
            End If
        Next
        ugAvailableGroups.DataSource = dtAvailable
        ugGroupwithAccess.DataSource = dtAccess
        LeftRightEnable()
    End Sub
    Private Sub newReportName()
        itemExists = False
        If cboReportName.Text <> String.Empty Then
            Try
                Dim dtNewReportname As DataTable = Me.cboReportName.DataSource
                Dim drNewReportname As DataRow
                'Check if the item is existing in the list,if not add to the combobox list.
                'If rpt.Name <> String.Empty Then
                If Reports.Name <> String.Empty Then
                    For Each drNewReportname In dtNewReportname.Rows
                        If Trim(UCase(drNewReportname.Item("REPORT_NAME"))) = Trim(UCase(Reports.Name)) Then
                            Me.cboReportName.SelectedIndex = -1
                            Me.cboReportName.SelectedValue = Reports.Name
                            itemExists = True
                        End If
                    Next
                    If Not itemExists Then
                        drNewReportname = dtNewReportname.NewRow
                        drNewReportname("REPORT_NAME") = Reports.Name
                        dtNewReportname.Rows.Add(drNewReportname)
                        Me.cboReportName.DataSource = dtNewReportname
                        Me.cboReportName.DisplayMember = "REPORT_NAME"
                        Me.cboReportName.ValueMember = "REPORT_NAME"
                        Me.cboReportName.SelectedValue = Reports.Name
                    End If
                End If
            Catch ex As Exception
                Throw ex
            End Try
        End If
    End Sub
    Private Function buildReports()
        Dim availableGroupItem As Object
        Dim groupWithAcessItem As Object
        Dim itemExists As Boolean = False
        Dim arr As ArrayList
        Dim i As Integer = 0
        Dim dtRow As DataRow

        Try
            'If Not rpt Is Nothing Then
            If Not Reports Is Nothing Then
                Me.btnDelete.Enabled = True
                strPreviousReportName = Reports.ID
                Me.cboReportName.SelectedValue = Reports.ID
                Me.txtReportName.ReadOnly = True
                Me.txtReportName.Text = Reports.Name
                'If rpt.Name = String.Empty Then
                If Me.cboReportName.SelectedIndex <= 0 Then
                    newReportName()
                End If
                If Reports.Path.StartsWith("\") Then
                    Me.txtReportFilePath.Text = Reports.Path
                Else
                    Me.txtReportFilePath.Text = REPORT_PATH + Reports.Path
                End If
                Me.txtDescription.Text = Reports.Description
                Me.cboModule.SelectedValue = IIf(Reports.Module = String.Empty, 0, Reports.Module)
                If Reports.Active = False Then
                    Me.chkInactive.Checked = False
                Else
                    Me.chkInactive.Checked = True
                End If
                ugGroupwithAccess.DataSource = Nothing
                ugAvailableGroups.DataSource = Nothing

                'dt = rpt.ListGroupsWithAccess()
                'arr = rpt.ListGroups()
                loadGrids()

                'Dim UserGroups As New MUSTER.BusinessLogic.pUserGroup

                'For Each dtRow In dt.Rows
                '    Me.lstGroupwithAccess.Items.Add(dtRow.Item("GROUP_NAME"))
                'Next
                'Dim row As DataRow
                'For Each row In UserGroups.ListUserGroups.Rows
                '    Me.lstAvailableGroups.Items.Add(row.Item("GROUP_NAME"))
                'Next

                'If Me.lstGroupwithAccess.Items.Count > 0 Then
                '    For Each groupWithAcessItem In Me.lstGroupwithAccess.Items
                '        If Me.lstAvailableGroups.Items.Count > 0 Then
                '            For Each availableGroupItem In Me.lstAvailableGroups.Items
                '                If groupWithAcessItem = availableGroupItem Then
                '                    Me.lstAvailableGroups.Items.Remove(availableGroupItem)
                '                    Exit For
                '                End If
                '            Next
                '        End If
                '    Next
                'End If

                'LeftRightEnable()
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Function ReportValidate() As Boolean

        Dim bolValidation As Boolean = True
        Dim strErrMsg As String = String.Empty

        Try

            'If Me.cboReportName.Text = String.Empty Then
            If Me.txtReportName.Text = String.Empty Then
                strErrMsg += vbTab + "Report name is missing" + vbCrLf
            End If
            If Me.cboModule.SelectedIndex = -1 Then
                strErrMsg += vbTab + "Report Module is missing" + vbCrLf
            End If
            If Me.txtReportFilePath.Text = String.Empty Then
                ' strErrMsg += vbTab + "Report Location is missing"
                strErrMsg += vbTab + "Report Path/File is missing"
            End If

            If strErrMsg.Length > 0 Then
                MsgBox("Invalid/Incomplete Report" + vbCrLf + strErrMsg)
                bolValidation = False
            End If

            Return bolValidation

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Sub Clear()
        Try
            Me.ugAvailableGroups.DataSource = Nothing
            Me.ugGroupwithAccess.DataSource = Nothing
            Me.txtName.Text = String.Empty
            Me.txtReportFilePath.Text = String.Empty
            Me.txtDescription.Text = String.Empty
            Me.chkInactive.Checked = False
            Me.btnDelete.Enabled = False
            Me.cboModule.SelectedIndex = -1
            Me.cboModule.SelectedIndex = -1
            Me.txtReportName.Text = String.Empty
            If Me.cboReportName.Text <> String.Empty Then
                Me.cboReportName.SelectedIndex = -1
                Me.cboReportName.SelectedIndex = -1
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Function LeftRightEnable()
        Try

            If Me.ugAvailableGroups.Rows.Count > 0 Then
                Me.btnShiftRight.Enabled = True
                If Me.ugAvailableGroups.Rows.Count > 1 Then
                    Me.btnShiftRightAll.Enabled = True
                Else
                    Me.btnShiftRightAll.Enabled = False
                End If
            Else
                Me.btnShiftRight.Enabled = False
                Me.btnShiftRightAll.Enabled = False
            End If


            If Me.ugGroupwithAccess.Rows.Count > 0 Then
                Me.btnShiftLeft.Enabled = True
                If Me.ugGroupwithAccess.Rows.Count > 1 Then
                    Me.btnShiftLeftAll.Enabled = True
                Else
                    Me.btnShiftLeftAll.Enabled = False
                End If
            Else
                Me.btnShiftLeft.Enabled = False
                Me.btnShiftLeftAll.Enabled = False
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region

#Region "UI Control Events"
    Private Sub btnSearchReport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearchReport.Click
        'P1 start
        Dim strFilePath As String
        Dim fileName As String
        Dim nReportID As Integer
        'Me.OpenFileDialog1.ShowDialog()
        Dim dlg As New OpenFileDialog
        dlg.Title = "Select Crystal Reports file"
        dlg.Filter = "Crystal Reports (*.rpt)|*.rpt|All Files (*.*)|*.*"
        If txtReportFilePath.Text = String.Empty Then
            If REPORT_PATH <> "\" Then
                dlg.InitialDirectory = REPORT_PATH
            End If
        Else
            If txtReportFilePath.Text.StartsWith("\") Then
                dlg.InitialDirectory = txtReportFilePath.Text
            Else
                dlg.InitialDirectory = REPORT_PATH + txtReportFilePath.Text
            End If
        End If
        Try

            'If Not Reports Is Nothing Then
            '    If Reports.IsDirty Then
            '        Dim Results As Long = MsgBox("There are unsaved changes. Do you want to save changes before closing?", MsgBoxStyle.YesNo)
            '        If Results = MsgBoxResult.Yes Then
            '            'rpt.Save()
            '            Reports.Save()
            '            loadReportNames()
            '        End If
            '    End If
            'End If
            If (dlg.ShowDialog() = DialogResult().OK) Then
                If System.IO.Path.GetExtension(dlg.FileName) <> ".rpt" Then
                    MsgBox("Invalid file")
                    Exit Sub
                End If

                'strFilePath = Me.OpenFileDialog1.FileName.ToString
                strFilePath = dlg.FileName.ToString

                fileName = System.IO.Path.GetFileNameWithoutExtension(strFilePath)

                Me.txtReportFilePath.Text = strFilePath

                If strFilePath.StartsWith(REPORT_PATH) Then
                    Reports.Path = strFilePath.Substring(REPORT_PATH.Length)
                Else
                    Reports.Path = strFilePath
                End If
            Else
                Me.btnCancel_Click(sender, e)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            Me.cboReportName.Enabled = True
            Me.txtReportName.Enabled = False
            Me.txtReportName.Text = ""

            If Not Reports Is Nothing Then
                If Not bolNameLeave Then Reports.Reset()
                If bolNameLeave Then bolNameLeave = False
                If Reports.ID <> 0 Then
                    buildReports()
                Else
                    Clear()
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click

        Me.Close()
    End Sub
    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim strValidPath As String
        Dim strErrMsg As String = String.Empty
        Try
            If Not Reports Is Nothing Then
                If chkInactive.Checked = False Then
                    Reports.Active = False
                End If
                If ReportValidate() Then
                    If Not System.IO.File.Exists(Me.txtReportFilePath.Text) Then
                        Dim Results As Long = MsgBox("The path  " & txtReportFilePath.Text & " not found on this system -" & vbCrLf & "Do you want to save them anyway?", MsgBoxStyle.YesNo)
                        If Results = MsgBoxResult.No Then
                            Exit Sub
                        End If
                    End If
                    'Path Validation
                    If Me.txtReportFilePath.Text <> String.Empty Then
                        'Checking whether the path is valid or not
                        strValidPath = UIUtilsGen.IsPathValid(Me.txtReportFilePath.Text)
                        If strValidPath <> String.Empty Then
                            strErrMsg += strValidPath
                            'Checking whether the file has extension or not
                            If System.IO.Path.HasExtension(Me.txtReportFilePath.Text) = False Then
                                strErrMsg += vbCrLf + "The file " + Me.txtReportFilePath.Text + " has no extension."
                            Else
                                'Checking if the file is of extension ".rpt"
                                If Not System.IO.Path.GetExtension(Me.txtReportFilePath.Text).StartsWith(".rpt") Then
                                    strErrMsg += vbCrLf + "Invalid file name.The file should be of extention .rpt."
                                End If
                            End If
                        End If
                        If strErrMsg.Length > 0 Then
                            Dim Results As Long = MsgBox(strErrMsg & " Do you want to save them anyway?", MsgBoxStyle.YesNo)
                            If Results = MsgBoxResult.No Then
                                Clear()
                                loadReportNames()
                                Exit Sub
                            End If
                        End If
                    End If
                    If Reports.ID <= 0 Then
                        Reports.CreatedBy = MusterContainer.AppUser.ID
                    Else
                        Reports.ModifiedBy = MusterContainer.AppUser.ID
                    End If
                    Reports.Save(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If
                    If bolIsNewReport Then
                        bolIsNewReport = False
                        Me.cboReportName.Enabled = True
                        Me.loadReportNames()
                        cboReportName.SelectedIndex = cboReportName.FindStringExact(Reports.Name)
                    End If

                    MsgBox("Report Save Successful", 0, "MUSTER Data Access Services")
                    Me.btnNew.Tag = String.Empty
                    If chkInactive.Checked = True Then
                    End If
                    txtReportName.ReadOnly = True
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Dim arr As ArrayList
        Dim i As Integer = 0
        Me.btnDelete.Enabled = False
        Try
            If Me.cboReportName.Text <> String.Empty Then
                If Not Reports Is Nothing Then
                    If Reports.IsDirty Then
                        Dim Results As Long = MsgBox("There are unsaved changes. Do you want to save changes before closing?", MsgBoxStyle.YesNo)
                        If Results = MsgBoxResult.No Then
                            Reports.Reset()
                        Else
                            If chkInactive.Checked = False Then
                                Reports.Active = False
                            End If
                            Reports.CreatedBy = MusterContainer.AppUser.ID
                            Reports.Save(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal)
                            If Not UIUtilsGen.HasRights(returnVal) Then
                                Exit Sub
                            End If

                        End If
                    End If
                End If
            End If
            resetform()
            txtReportName.ReadOnly = False
            txtReportName.Enabled = True
            txtReportName.Text = ""
            Reports = New MUSTER.BusinessLogic.pReport
            Me.Clear()
            loadGrids()
            'Dim UserGroups As New MUSTER.BusinessLogic.pUserGroup
            'Dim row As DataRow
            'For Each row In UserGroups.ListUserGroups.Rows
            'Me.lstAvailableGroups.Items.Add(row.Item("GROUP_NAME"))
            'Next

            Me.btnNew.Tag = "New Report"
            Me.LeftRightEnable()
            Me.txtReportName.Focus()
            Me.cboReportName.Enabled = False
            bolIsNewReport = True
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnShiftLeft_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnShiftLeft.Click
        Try
            If Not ugGroupwithAccess.Selected.Rows Is Nothing Then
                If ugGroupwithAccess.Selected.Rows.Count > 0 Then
                    For Each ugRow In ugGroupwithAccess.Selected.Rows
                        reportGroupRelInfo = Reports.ReportGroupRelationCollection.Item(Reports.ID.ToString + "|" + ugRow.Cells("GROUP_ID").Text)
                        If reportGroupRelInfo.isNew Then
                            Reports.ReportGroupRelationCollection.ChangeKey(reportGroupRelInfo.ID, "0|" + reportGroupRelInfo.GroupID.ToString)
                            reportGroupRelInfo.ReportID = 0
                        Else
                            reportGroupRelInfo.Deleted = True
                        End If
                    Next
                    SetSaveCancel(Reports.IsDirty)
                    loadGrids()
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnShiftLeftAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnShiftLeftAll.Click
        Try
            For Each ugRow In ugGroupwithAccess.Rows
                reportGroupRelInfo = Reports.ReportGroupRelationCollection.Item(Reports.ID.ToString + "|" + ugRow.Cells("GROUP_ID").Text)
                If reportGroupRelInfo.isNew Then
                    Reports.ReportGroupRelationCollection.ChangeKey(reportGroupRelInfo.ID, "0|" + reportGroupRelInfo.GroupID.ToString)
                    reportGroupRelInfo.ReportID = 0
                Else
                    reportGroupRelInfo.Deleted = True
                End If
            Next
            SetSaveCancel(Reports.IsDirty)
            loadGrids()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnShiftRight_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnShiftRight.Click
        Try
            If Not ugAvailableGroups.Selected.Rows Is Nothing Then
                If ugAvailableGroups.Selected.Rows.Count > 0 Then
                    For Each ugRow In ugAvailableGroups.Selected.Rows
                        reportGroupRelInfo = Reports.ReportGroupRelationCollection.Item(Reports.ID.ToString + "|" + ugRow.Cells("GROUP_ID").Text)
                        If reportGroupRelInfo Is Nothing Then
                            reportGroupRelInfo = Reports.ReportGroupRelationCollection.Item("0|" + ugRow.Cells("GROUP_ID").Text)
                            reportGroupRelInfo.ReportID = Reports.ID
                            Reports.ReportGroupRelationCollection.ChangeKey("0|" + ugRow.Cells("GROUP_ID").Text, Reports.ID.ToString + "|" + ugRow.Cells("GROUP_ID").Text)
                        End If
                        reportGroupRelInfo.Deleted = False
                    Next
                    SetSaveCancel(Reports.IsDirty)
                    loadGrids()
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnShiftRightAll_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnShiftRightAll.Click
        Try
            For Each ugRow In ugAvailableGroups.Rows
                reportGroupRelInfo = Reports.ReportGroupRelationCollection.Item(Reports.ID.ToString + "|" + ugRow.Cells("GROUP_ID").Text)
                If reportGroupRelInfo Is Nothing Then
                    reportGroupRelInfo = Reports.ReportGroupRelationCollection.Item("0|" + ugRow.Cells("GROUP_ID").Text)
                    reportGroupRelInfo.ReportID = Reports.ID
                    Reports.ReportGroupRelationCollection.ChangeKey("0|" + ugRow.Cells("GROUP_ID").Text, Reports.ID.ToString + "|" + ugRow.Cells("GROUP_ID").Text)
                End If
                reportGroupRelInfo.Deleted = False
            Next
            SetSaveCancel(Reports.IsDirty)
            loadGrids()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtName_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtName.Leave
        Try
            If Me.txtName.Text <> String.Empty Then
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtDescription_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDescription.Leave
        Try
            If Me.txtDescription.Text <> String.Empty Then
                Reports.Description = Me.txtDescription.Text
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub cboModule_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboModule.SelectedIndexChanged
        If bolLoading Then Exit Sub

        Try
            If cboModule.SelectedIndex > -1 Then
                If Not Reports Is Nothing Then
                    Reports.Module = UIUtilsGen.GetComboBoxValueInt(cboModule)
                End If
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub cboReportName_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboReportName.SelectedValueChanged
        If bolLoading Then Exit Sub
        Try
            If Me.cboReportName.SelectedIndex > 0 Then
                Try
                    If Not Reports Is Nothing Then
                        If Reports.IsDirty Then
                            Dim Results As Long = MsgBox("There are unsaved changes. Do you want to save changes before closing?", MsgBoxStyle.YesNo)
                            If Results = MsgBoxResult.Yes Then
                                If chkInactive.Checked = False Then
                                    Reports.Active = False
                                End If
                                If Reports.ID <= 0 Then
                                    Reports.CreatedBy = MusterContainer.AppUser.ID
                                Else
                                    Reports.ModifiedBy = MusterContainer.AppUser.ID
                                End If
                                Reports.Save(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal)
                                If Not UIUtilsGen.HasRights(returnVal) Then
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If

                Catch ex As Exception
                    Throw ex
                Finally
                    Reports = New MUSTER.BusinessLogic.pReport
                    Reports.Retrieve(Me.cboReportName.SelectedValue)
                    ReportParams = New MUSTER.BusinessLogic.pReportParams
                    ReportParams.ReportID = Reports.ID
                    Me.btnReportParams.Enabled = True
                    loadReportNames()
                    Me.cboReportName.SelectedValue = Reports.ID
                    Me.btnNew.Tag = String.Empty
                    buildReports()
                End Try
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtReportFilePath_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtReportFilePath.Leave
        Try
            If Me.txtReportFilePath.Text <> String.Empty Then
                If Reports.ReportFileExists(Me.txtReportFilePath.Text, Reports.ID) Then
                    MsgBox("The specified file is already being used by another report.")
                    Me.txtReportFilePath.Text = ""
                    Me.txtReportFilePath.Focus()
                Else
                    Reports.Path = Me.txtReportFilePath.Text
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub chkInactive_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkInactive.Click
        Try
            Reports.Active = chkInactive.Checked
            'If chkInactive.Checked = True Then
            '    chkRptShowDeleted.Checked = True
            'End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnReportParams_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReportParams.Click
        Try
            Dim cr As New ReportDocument
            Dim rname As String = Reports.Path
            If Not rname.ToUpper.StartsWith(REPORT_PATH.ToUpper) Then
                If Not rname.StartsWith("\") Then
                    rname = REPORT_PATH + rname
                End If
            End If
            If System.IO.File.Exists(rname) Then
                cr.Load(rname)
            Else
                If Reports Is Nothing Then
                    MsgBox("The report is not found.")
                Else
                    MsgBox("the report " & Reports.Name & " is not found.")
                End If
                Exit Sub
            End If
            If cr.DataDefinition.ParameterFields.Count > 0 Then
                Dim form As New CrystalReportsParamMaint
                form.Cr = cr
                form.ReportParams = Me.ReportParams
                form.ShowDialog()
            Else
                MsgBox("There are no parameters for this report.")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub chkRptShowDeleted_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkRptShowDeleted.Click
        bolShowDeleted = chkRptShowDeleted.Checked
        Try
            Me.loadReportNames()
            Me.cboReportName.SelectedIndex = -1

            If (Me.chkInactive.Checked = True And Me.chkRptShowDeleted.Checked = False) Then 'Or Me.cboReportName.Tag.startswith("New Report") Then '("New Report")
                Me.Clear()
                Me.cboReportName.Tag = String.Empty
                'Reports = Nothing
            Else
                If Not Reports Is Nothing And Not Reports.Deleted Then
                    Me.cboReportName.SelectedValue = Reports.ID 'Me.cboReportName.Tag
                End If
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtDescription_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDescription.TextChanged
        Reports.Description = Me.txtDescription.Text
    End Sub
    Private Sub txtReportFilePath_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtReportFilePath.TextChanged
        '
    End Sub
    Private Sub txtReportName_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtReportName.Leave
        Try
            bolNameLeave = False
            If Me.txtReportName.ReadOnly = False Then
                If Me.txtReportName.Text <> String.Empty Then
                    If Not (Me.cboReportName.FindStringExact(Me.txtReportName.Text) > -1) Then
                        Dim strUserName As String = Me.txtReportName.Text
                        Dim oLocalReport As New MUSTER.Info.ReportInfo(0, Me.txtReportName.Text, "", "", "", False, MusterContainer.AppUser.ID, Now, String.Empty, CDate("01/01/0001"), False)
                        Reports.Add(oLocalReport)
                    Else
                        '
                        ' Need this to perform UI reset if not a new user
                        '
                        If Me.txtReportName.Text <> String.Empty Then
                            MsgBox("A report with that name already exists.")
                            cboReportName.Text = Me.txtReportName.Text
                            Reports.Retrieve(txtReportName.Text)
                        End If
                    End If
                Else
                    Dim msgResult As MsgBoxResult = MsgBox("Report name cannot be blank.Would you like to continue creating a new report?", MsgBoxStyle.Question & MsgBoxStyle.YesNo, "NEW REPORT")
                    If msgResult = MsgBoxResult.No Then
                        If Me.ActiveControl.Name = "btnCancel" Then
                            bolNameLeave = True
                            Me.btnCancel_Click(sender, e)
                        Else
                            If strPreviousReportName <> String.Empty Then
                                Me.cboReportName.SelectedValue = strPreviousReportName
                            Else
                                Me.cboReportName.SelectedIndex = 1
                            End If
                            cboReportName.Enabled = True
                            Me.cboReportName_SelectedValueChanged(sender, e)
                        End If
                    Else
                        Me.txtReportName.Focus()
                        Exit Sub
                    End If
                End If
                bolLoading = False
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub chkInactive_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkInactive.CheckedChanged

    End Sub
    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim msgResult As MsgBoxResult = MsgBox("Are you sure you wish to DELETE the report: " & Me.cboReportName.Text & " ?", MsgBoxStyle.Question & MsgBoxStyle.YesNo, "DELETE REPORT")
        If msgResult = MsgBoxResult.Yes Then
            Reports.Deleted = True
            If Reports.ID <= 0 Then
                Reports.CreatedBy = MusterContainer.AppUser.ID
            Else
                Reports.ModifiedBy = MusterContainer.AppUser.ID
            End If
            For Each reportGroupRelInfo In Reports.ReportGroupRelationCollection.Values
                reportGroupRelInfo.Deleted = True
            Next
            Reports.Save(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If
            MsgBox("Report Delete Successful", 0, "MUSTER Data Access Services")
            Clear()
            loadReportNames()
        Else
            'do nothing
        End If
    End Sub
    Private Sub ugAvailableGroups_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugAvailableGroups.InitializeLayout
        SetupGroupGrid(sender)
    End Sub
    Private Sub ugGroupwithAccess_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugGroupwithAccess.InitializeLayout
        SetupGroupGrid(sender)
    End Sub
    Private Sub ugAvailableGroups_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugAvailableGroups.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            btnShiftRight_Click(sender, e)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugGroupwithAccess_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugGroupwithAccess.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            btnShiftLeft_Click(sender, e)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "External Event Handlers"
    Private Sub thisReportChanged(ByVal bolvalue As Boolean) Handles Reports.ReportChanged
        SetSaveCancel(bolvalue)
    End Sub
#End Region

    Public Property ReportParams() As MUSTER.BusinessLogic.pReportParams
        Get
            Return Me.Reports.ReportParams
        End Get
        Set(ByVal Value As MUSTER.BusinessLogic.pReportParams)
            Me.Reports.ReportParams = Value
        End Set
    End Property

End Class
