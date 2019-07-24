Public Class GroupAdmin
    '-------------------------------------------------------------------------------
    ' MUSTER.MUSTER.FilePaths
    '   Provides the UI for managing system user groups
    '
    ' Copyright (C) 2004 CIBER, Inc.
    ' All rights reserved.
    '
    ' Release   Initials    Date        Description
    '  1.0        AN      12/??/04    Original class definition.
    '  1.1        JVC2    12/28/04    Altered to use new UserGroup object
    '                                 Removed reference to txtName (No longer in form)
    '  1.2        JC      01/03/05    Added code to handle MDI interface for app
    '  1.3        JC      01/12/05    Reworked code to handle text changes and 
    '                                  to handle Save on exit properly.  Also
    '                                  added code to handle textchanged events.
    '  1.4        PN      01/19/05     Modified method ClearGroupData(bug 651).  
    '  1.5        PN      01/24/05     Set Controls tab order(bug 660) 
    '  1.6        JVC2    01/31/05     Added Close GotFocus to handle New/Close 
    '                                   combination (bug 670).
    '  1.7        PN      01/31/05    added function isEmptyGroupName and two events
    '                                 txtDescription_Enter,chkInactive_Enter(bug 670,616)  
    '  1.8        AN      02/10/05    Integrated AppFlags new object model
    '
    '  1.9        JVC2    08/08/05    Added ability to delete calendar entries if the 
    '                                   group delete is denied due to existing, active
    '                                   calendar entries.
    '-------------------------------------------------------------------------------
    '
    'TODO - Remove comment from VSS version 2/9/05 - JVC 2
    '
#Region "Private Member Variables"
    Inherits System.Windows.Forms.Form
    Dim WithEvents oGroups As MUSTER.BusinessLogic.pUserGroup
    Dim strPreviousGroupName As String = String.Empty
    Dim bolIsNewGroup As Boolean = False
    Dim bolIsLoading As Boolean = True
    Dim bolErrorOccurred As Boolean = False

    Dim groupModuleInfo As MUSTER.Info.GroupModuleRelationInfo
    Dim dtRO, dtRW, dtAvailableRO, dtAvailableRW As DataTable
    Dim dr As DataRow
    Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow

    'Dim bolIsClosing As Boolean = False
    Friend MyGUID As New System.Guid
    Dim returnVal As String = String.Empty
#End Region
#Region " Windows Form Designer generated code "

    Public Sub New(Optional ByRef frm As Windows.Forms.Form = Nothing)
        MyBase.New()

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
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents lblName As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents chkInactive As System.Windows.Forms.CheckBox
    Friend WithEvents lblDescription As System.Windows.Forms.Label
    Friend WithEvents txtDescription As System.Windows.Forms.TextBox
    Friend WithEvents lblReadOnlyForms As System.Windows.Forms.Label
    Friend WithEvents lblReadWriteForms As System.Windows.Forms.Label
    Friend WithEvents lblAvailableFormsforRW As System.Windows.Forms.Label
    Friend WithEvents lblAvailableFormsforRO As System.Windows.Forms.Label
    Friend WithEvents btnRWShiftLeftAll As System.Windows.Forms.Button
    Friend WithEvents btnRWShiftLeft As System.Windows.Forms.Button
    Friend WithEvents btnRWShiftRightAll As System.Windows.Forms.Button
    Friend WithEvents btnRWShiftRight As System.Windows.Forms.Button
    Friend WithEvents btnROShiftLeftAll As System.Windows.Forms.Button
    Friend WithEvents btnROShiftLeft As System.Windows.Forms.Button
    Friend WithEvents btnROShiftRightAll As System.Windows.Forms.Button
    Friend WithEvents btnROShiftRight As System.Windows.Forms.Button
    Friend WithEvents ComboGroupName As System.Windows.Forms.ComboBox
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents txtGroupName As System.Windows.Forms.TextBox
    Friend WithEvents lblGroupList As System.Windows.Forms.Label
    Friend WithEvents ugAvailableRO As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents ugAvailableRW As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents ugRW As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents ugRO As Infragistics.Win.UltraWinGrid.UltraGrid
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnSearch = New System.Windows.Forms.Button
        Me.chkInactive = New System.Windows.Forms.CheckBox
        Me.lblName = New System.Windows.Forms.Label
        Me.lblDescription = New System.Windows.Forms.Label
        Me.txtDescription = New System.Windows.Forms.TextBox
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnRWShiftLeftAll = New System.Windows.Forms.Button
        Me.btnRWShiftLeft = New System.Windows.Forms.Button
        Me.btnRWShiftRightAll = New System.Windows.Forms.Button
        Me.btnRWShiftRight = New System.Windows.Forms.Button
        Me.lblReadWriteForms = New System.Windows.Forms.Label
        Me.lblAvailableFormsforRW = New System.Windows.Forms.Label
        Me.btnROShiftLeftAll = New System.Windows.Forms.Button
        Me.btnROShiftLeft = New System.Windows.Forms.Button
        Me.btnROShiftRightAll = New System.Windows.Forms.Button
        Me.btnROShiftRight = New System.Windows.Forms.Button
        Me.lblReadOnlyForms = New System.Windows.Forms.Label
        Me.lblAvailableFormsforRO = New System.Windows.Forms.Label
        Me.ComboGroupName = New System.Windows.Forms.ComboBox
        Me.btnDelete = New System.Windows.Forms.Button
        Me.txtGroupName = New System.Windows.Forms.TextBox
        Me.lblGroupList = New System.Windows.Forms.Label
        Me.ugAvailableRO = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.ugAvailableRW = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.ugRW = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.ugRO = New Infragistics.Win.UltraWinGrid.UltraGrid
        CType(Me.ugAvailableRO, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ugAvailableRW, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ugRW, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ugRO, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnSearch
        '
        Me.btnSearch.BackColor = System.Drawing.SystemColors.Control
        Me.btnSearch.Enabled = False
        Me.btnSearch.Location = New System.Drawing.Point(304, 8)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(24, 24)
        Me.btnSearch.TabIndex = 1
        Me.btnSearch.Text = "?"
        '
        'chkInactive
        '
        Me.chkInactive.Location = New System.Drawing.Point(336, 40)
        Me.chkInactive.Name = "chkInactive"
        Me.chkInactive.Size = New System.Drawing.Size(72, 24)
        Me.chkInactive.TabIndex = 3
        Me.chkInactive.Text = "Inactive"
        '
        'lblName
        '
        Me.lblName.Location = New System.Drawing.Point(32, 40)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(106, 16)
        Me.lblName.TabIndex = 112
        Me.lblName.Text = "New Group Name:"
        Me.lblName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblDescription
        '
        Me.lblDescription.Location = New System.Drawing.Point(32, 72)
        Me.lblDescription.Name = "lblDescription"
        Me.lblDescription.Size = New System.Drawing.Size(64, 16)
        Me.lblDescription.TabIndex = 116
        Me.lblDescription.Text = "Description:"
        Me.lblDescription.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtDescription
        '
        Me.txtDescription.Location = New System.Drawing.Point(128, 72)
        Me.txtDescription.Multiline = True
        Me.txtDescription.Name = "txtDescription"
        Me.txtDescription.Size = New System.Drawing.Size(216, 64)
        Me.txtDescription.TabIndex = 4
        Me.txtDescription.Text = ""
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(384, 464)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(80, 26)
        Me.btnClose.TabIndex = 21
        Me.btnClose.Text = "Close"
        '
        'btnCancel
        '
        Me.btnCancel.Enabled = False
        Me.btnCancel.Location = New System.Drawing.Point(296, 464)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(80, 26)
        Me.btnCancel.TabIndex = 20
        Me.btnCancel.Text = "Cancel"
        '
        'btnSave
        '
        Me.btnSave.Enabled = False
        Me.btnSave.Location = New System.Drawing.Point(120, 464)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(80, 26)
        Me.btnSave.TabIndex = 18
        Me.btnSave.Text = "Save"
        '
        'btnNew
        '
        Me.btnNew.Location = New System.Drawing.Point(32, 464)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 17
        Me.btnNew.Text = "New"
        '
        'btnRWShiftLeftAll
        '
        Me.btnRWShiftLeftAll.BackColor = System.Drawing.SystemColors.Control
        Me.btnRWShiftLeftAll.Enabled = False
        Me.btnRWShiftLeftAll.Location = New System.Drawing.Point(208, 408)
        Me.btnRWShiftLeftAll.Name = "btnRWShiftLeftAll"
        Me.btnRWShiftLeftAll.Size = New System.Drawing.Size(32, 24)
        Me.btnRWShiftLeftAll.TabIndex = 15
        Me.btnRWShiftLeftAll.Text = "<<"
        '
        'btnRWShiftLeft
        '
        Me.btnRWShiftLeft.BackColor = System.Drawing.SystemColors.Control
        Me.btnRWShiftLeft.Enabled = False
        Me.btnRWShiftLeft.Location = New System.Drawing.Point(208, 384)
        Me.btnRWShiftLeft.Name = "btnRWShiftLeft"
        Me.btnRWShiftLeft.Size = New System.Drawing.Size(32, 24)
        Me.btnRWShiftLeft.TabIndex = 14
        Me.btnRWShiftLeft.Text = "<"
        '
        'btnRWShiftRightAll
        '
        Me.btnRWShiftRightAll.BackColor = System.Drawing.SystemColors.Control
        Me.btnRWShiftRightAll.Enabled = False
        Me.btnRWShiftRightAll.Location = New System.Drawing.Point(208, 352)
        Me.btnRWShiftRightAll.Name = "btnRWShiftRightAll"
        Me.btnRWShiftRightAll.Size = New System.Drawing.Size(32, 24)
        Me.btnRWShiftRightAll.TabIndex = 13
        Me.btnRWShiftRightAll.Text = ">>"
        '
        'btnRWShiftRight
        '
        Me.btnRWShiftRight.BackColor = System.Drawing.SystemColors.Control
        Me.btnRWShiftRight.Enabled = False
        Me.btnRWShiftRight.Location = New System.Drawing.Point(208, 328)
        Me.btnRWShiftRight.Name = "btnRWShiftRight"
        Me.btnRWShiftRight.Size = New System.Drawing.Size(32, 24)
        Me.btnRWShiftRight.TabIndex = 12
        Me.btnRWShiftRight.Text = ">"
        '
        'lblReadWriteForms
        '
        Me.lblReadWriteForms.Location = New System.Drawing.Point(256, 304)
        Me.lblReadWriteForms.Name = "lblReadWriteForms"
        Me.lblReadWriteForms.Size = New System.Drawing.Size(107, 16)
        Me.lblReadWriteForms.TabIndex = 145
        Me.lblReadWriteForms.Text = "Read/Write Modules"
        '
        'lblAvailableFormsforRW
        '
        Me.lblAvailableFormsforRW.Location = New System.Drawing.Point(32, 304)
        Me.lblAvailableFormsforRW.Name = "lblAvailableFormsforRW"
        Me.lblAvailableFormsforRW.Size = New System.Drawing.Size(160, 16)
        Me.lblAvailableFormsforRW.TabIndex = 143
        Me.lblAvailableFormsforRW.Text = "Read/Write Available Modules"
        '
        'btnROShiftLeftAll
        '
        Me.btnROShiftLeftAll.BackColor = System.Drawing.SystemColors.Control
        Me.btnROShiftLeftAll.Enabled = False
        Me.btnROShiftLeftAll.Location = New System.Drawing.Point(208, 256)
        Me.btnROShiftLeftAll.Name = "btnROShiftLeftAll"
        Me.btnROShiftLeftAll.Size = New System.Drawing.Size(32, 24)
        Me.btnROShiftLeftAll.TabIndex = 9
        Me.btnROShiftLeftAll.Text = "<<"
        '
        'btnROShiftLeft
        '
        Me.btnROShiftLeft.BackColor = System.Drawing.SystemColors.Control
        Me.btnROShiftLeft.Enabled = False
        Me.btnROShiftLeft.Location = New System.Drawing.Point(208, 232)
        Me.btnROShiftLeft.Name = "btnROShiftLeft"
        Me.btnROShiftLeft.Size = New System.Drawing.Size(32, 24)
        Me.btnROShiftLeft.TabIndex = 8
        Me.btnROShiftLeft.Text = "<"
        '
        'btnROShiftRightAll
        '
        Me.btnROShiftRightAll.BackColor = System.Drawing.SystemColors.Control
        Me.btnROShiftRightAll.Enabled = False
        Me.btnROShiftRightAll.Location = New System.Drawing.Point(208, 200)
        Me.btnROShiftRightAll.Name = "btnROShiftRightAll"
        Me.btnROShiftRightAll.Size = New System.Drawing.Size(32, 24)
        Me.btnROShiftRightAll.TabIndex = 7
        Me.btnROShiftRightAll.Text = ">>"
        '
        'btnROShiftRight
        '
        Me.btnROShiftRight.BackColor = System.Drawing.SystemColors.Control
        Me.btnROShiftRight.Enabled = False
        Me.btnROShiftRight.Location = New System.Drawing.Point(208, 176)
        Me.btnROShiftRight.Name = "btnROShiftRight"
        Me.btnROShiftRight.Size = New System.Drawing.Size(32, 24)
        Me.btnROShiftRight.TabIndex = 6
        Me.btnROShiftRight.Text = ">"
        '
        'lblReadOnlyForms
        '
        Me.lblReadOnlyForms.Location = New System.Drawing.Point(256, 152)
        Me.lblReadOnlyForms.Name = "lblReadOnlyForms"
        Me.lblReadOnlyForms.Size = New System.Drawing.Size(104, 16)
        Me.lblReadOnlyForms.TabIndex = 137
        Me.lblReadOnlyForms.Text = "Read/Only Modules"
        '
        'lblAvailableFormsforRO
        '
        Me.lblAvailableFormsforRO.Location = New System.Drawing.Point(32, 152)
        Me.lblAvailableFormsforRO.Name = "lblAvailableFormsforRO"
        Me.lblAvailableFormsforRO.Size = New System.Drawing.Size(153, 16)
        Me.lblAvailableFormsforRO.TabIndex = 135
        Me.lblAvailableFormsforRO.Text = "Read/Only Available Modules"
        '
        'ComboGroupName
        '
        Me.ComboGroupName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboGroupName.Location = New System.Drawing.Point(128, 8)
        Me.ComboGroupName.Name = "ComboGroupName"
        Me.ComboGroupName.Size = New System.Drawing.Size(160, 21)
        Me.ComboGroupName.TabIndex = 0
        '
        'btnDelete
        '
        Me.btnDelete.Enabled = False
        Me.btnDelete.Location = New System.Drawing.Point(208, 464)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 19
        Me.btnDelete.Text = "Delete"
        '
        'txtGroupName
        '
        Me.txtGroupName.Location = New System.Drawing.Point(128, 40)
        Me.txtGroupName.Name = "txtGroupName"
        Me.txtGroupName.ReadOnly = True
        Me.txtGroupName.Size = New System.Drawing.Size(160, 20)
        Me.txtGroupName.TabIndex = 2
        Me.txtGroupName.Text = ""
        '
        'lblGroupList
        '
        Me.lblGroupList.Location = New System.Drawing.Point(32, 10)
        Me.lblGroupList.Name = "lblGroupList"
        Me.lblGroupList.Size = New System.Drawing.Size(66, 20)
        Me.lblGroupList.TabIndex = 146
        Me.lblGroupList.Text = "Group List: "
        Me.lblGroupList.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ugAvailableRO
        '
        Me.ugAvailableRO.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugAvailableRO.Location = New System.Drawing.Point(32, 168)
        Me.ugAvailableRO.Name = "ugAvailableRO"
        Me.ugAvailableRO.Size = New System.Drawing.Size(152, 120)
        Me.ugAvailableRO.TabIndex = 148
        '
        'ugAvailableRW
        '
        Me.ugAvailableRW.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugAvailableRW.Location = New System.Drawing.Point(32, 320)
        Me.ugAvailableRW.Name = "ugAvailableRW"
        Me.ugAvailableRW.Size = New System.Drawing.Size(152, 120)
        Me.ugAvailableRW.TabIndex = 148
        '
        'ugRW
        '
        Me.ugRW.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugRW.Location = New System.Drawing.Point(256, 320)
        Me.ugRW.Name = "ugRW"
        Me.ugRW.Size = New System.Drawing.Size(152, 120)
        Me.ugRW.TabIndex = 148
        '
        'ugRO
        '
        Me.ugRO.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugRO.Location = New System.Drawing.Point(256, 168)
        Me.ugRO.Name = "ugRO"
        Me.ugRO.Size = New System.Drawing.Size(152, 120)
        Me.ugRO.TabIndex = 148
        '
        'GroupAdmin
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(544, 518)
        Me.Controls.Add(Me.ugAvailableRO)
        Me.Controls.Add(Me.lblGroupList)
        Me.Controls.Add(Me.txtGroupName)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.ComboGroupName)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnRWShiftLeftAll)
        Me.Controls.Add(Me.btnRWShiftLeft)
        Me.Controls.Add(Me.btnRWShiftRightAll)
        Me.Controls.Add(Me.btnRWShiftRight)
        Me.Controls.Add(Me.lblReadWriteForms)
        Me.Controls.Add(Me.lblAvailableFormsforRW)
        Me.Controls.Add(Me.btnROShiftLeftAll)
        Me.Controls.Add(Me.btnROShiftLeft)
        Me.Controls.Add(Me.btnROShiftRightAll)
        Me.Controls.Add(Me.btnROShiftRight)
        Me.Controls.Add(Me.lblReadOnlyForms)
        Me.Controls.Add(Me.lblAvailableFormsforRO)
        Me.Controls.Add(Me.lblDescription)
        Me.Controls.Add(Me.txtDescription)
        Me.Controls.Add(Me.btnSearch)
        Me.Controls.Add(Me.chkInactive)
        Me.Controls.Add(Me.lblName)
        Me.Controls.Add(Me.ugAvailableRW)
        Me.Controls.Add(Me.ugRW)
        Me.Controls.Add(Me.ugRO)
        Me.Name = "GroupAdmin"
        Me.Text = "Manage Groups"
        CType(Me.ugAvailableRO, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ugAvailableRW, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ugRW, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ugRO, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region "Form Level Events"
    Protected Overrides Sub OnLoad(ByVal e As System.EventArgs)
        bolIsLoading = True
        Me.btnDelete.Enabled = False
        oGroups = New MUSTER.BusinessLogic.pUserGroup
        oGroups.ShowDeleted = False
        Me.InitUserGroups()
        ComboGroupName.SelectedIndex = IIf(ComboGroupName.Items.Count > 0, 0, -1)
        bolIsLoading = False
        If ComboGroupName.SelectedIndex >= 0 Then
            'Dim sender As New Object
            'ComboGroupName_Leave()
            btnDelete.Enabled = True
            GetGroupData()
            DisplayGroupData()
        End If
    End Sub
    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)
        If oGroups.colIsDirty Then
            Dim msgRet As MsgBoxResult = MsgBox("There are changes for the groups.  Do you wish to save the changes?", MsgBoxStyle.Question & MsgBoxStyle.YesNoCancel, "Data Changed")
            If msgRet = MsgBoxResult.Yes Then
                oGroups.Flush(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If
            Else
                If msgRet = MsgBoxResult.Cancel Then
                    e.Cancel = True
                End If
            End If
        End If
        ' Remove any values from the shared collection for this screen
        '
        MusterContainer.AppSemaphores.Remove(MyGUID.ToString)
        '
        ' Log the disposal of the form (exit from Registration form)
        '
        MusterContainer.AppUser.LogExit(MyGUID.ToString)

    End Sub
#End Region
#Region "UI Support Routines"
    Private Sub InitDataTable()
        dtRO = New DataTable
        dtRW = New DataTable
        dtAvailableRO = New DataTable
        dtAvailableRW = New DataTable

        dtRO.Columns.Add("MODULE_ID", GetType(Integer))
        dtRO.Columns.Add("MODULE", GetType(String))

        dtRW = dtRO.Clone
        dtAvailableRO = dtRO.Clone
        dtAvailableRW = dtRO.Clone
    End Sub
    Private Sub SetupGrid(ByRef ug As Infragistics.Win.UltraWinGrid.UltraGrid)
        ug.DisplayLayout.Bands(0).Columns("MODULE_ID").Hidden = True
        ug.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        ug.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        'ug.DisplayLayout.Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Free
        ug.DisplayLayout.Bands(0).Columns("MODULE").Width = 133
        ug.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.False
    End Sub
    Private Sub InitUserGroups()
        ComboGroupName.DataSource = oGroups.ListUserGroups
        ComboGroupName.DisplayMember = "GROUP_NAME"
        ComboGroupName.ValueMember = "GROUP_ID"
    End Sub
    Private Sub GetGroupData()
        Dim strOldGroupName As String = oGroups.Name
        If bolIsNewGroup Then
            Exit Sub
        End If

        strPreviousGroupName = ComboGroupName.Text
        If ComboGroupName.SelectedIndex > -1 Then
            oGroups.Retrieve(UIUtilsGen.GetComboBoxValueInt(ComboGroupName))
        End If
    End Sub

    Private Sub DisplayGroupData()
        If ComboGroupName.SelectedIndex < 0 Then Exit Sub
        ComboGroupName.Enabled = True
        txtGroupName.ReadOnly = True
        Me.btnDelete.Enabled = True
        txtGroupName.Text = oGroups.Name
        txtDescription.Text = oGroups.Description
        chkInactive.Checked = oGroups.Active
        bolIsLoading = False
        DisplayScreenDisposition()
    End Sub

    Private Sub DisplayScreenDisposition()
        If ComboGroupName.SelectedValue.ToString <> "System.Data.DataRowView" Then
            If ComboGroupName.SelectedValue.ToString <> String.Empty Then
                LoadScreenListviews()
            End If
        End If
    End Sub

    Private Sub ClearScreenData()
        ugRO.DataSource = Nothing
        ugRW.DataSource = Nothing
        ugAvailableRO.DataSource = Nothing
        ugAvailableRW.DataSource = Nothing

        dtRO.Rows.Clear()
        dtRW.Rows.Clear()
        dtAvailableRO.Rows.Clear()
        dtAvailableRW.Rows.Clear()
    End Sub

    Private Sub ClearGroupData()
        bolIsLoading = True
        ComboGroupName.SelectedIndex = -1
        Me.txtGroupName.Text = String.Empty
        Me.chkInactive.Checked = False
        'P1 19/01/04 next line only
        'bolIsLoading = False
        txtDescription.Text = String.Empty
        'P1 19/01/04 next line only
        bolIsLoading = False
        Me.btnDelete.Enabled = False
    End Sub

    Private Sub LoadScreenListviews()
        'Dim oScreen As MUSTER.Info.ProfileInfo
        ClearScreenData()
        For Each groupModuleInfo In oGroups.GroupModuleRelationCollection.Values
            If groupModuleInfo.WriteAccess Then
                dr = dtRW.NewRow
                dr("MODULE_ID") = groupModuleInfo.ModuleID
                dr("MODULE") = groupModuleInfo.ModuleName
                dtRW.Rows.Add(dr)
            ElseIf groupModuleInfo.ReadAccess Then
                dr = dtRO.NewRow
                dr("MODULE_ID") = groupModuleInfo.ModuleID
                dr("MODULE") = groupModuleInfo.ModuleName
                dtRO.Rows.Add(dr)
            Else
                dr = dtAvailableRO.NewRow
                dr("MODULE_ID") = groupModuleInfo.ModuleID
                dr("MODULE") = groupModuleInfo.ModuleName
                dtAvailableRO.Rows.Add(dr)
                dr = dtAvailableRW.NewRow
                dr("MODULE_ID") = groupModuleInfo.ModuleID
                dr("MODULE") = groupModuleInfo.ModuleName
                dtAvailableRW.Rows.Add(dr)
            End If
        Next

        ugRO.DataSource = dtRO
        ugRW.DataSource = dtRW
        ugAvailableRO.DataSource = dtAvailableRO
        ugAvailableRW.DataSource = dtAvailableRW

        'For Each oScreen In oGroups.Screens.Values
        '    If oScreen.User = oGroups.Name Then
        '        Select Case oScreen.ProfileValue
        '            Case "NONE"
        '                lstAvailableFormsforRO.Items.Add(oScreen.ProfileMod1)
        '                lstAvailableFormsforRW.Items.Add(oScreen.ProfileMod1)
        '            Case "RO"
        '                lstReadOnlyForms.Items.Add(oScreen.ProfileMod1)
        '            Case "RW"
        '                lstReadWriteForms.Items.Add(oScreen.ProfileMod1)
        '        End Select
        '    End If
        'Next
        LeftRightEnable()
    End Sub

    Private Sub SetSaveCancel(ByVal bolState As Boolean)
        btnSave.Enabled = bolState
        btnCancel.Enabled = bolState
    End Sub
    Private Function LeftRightEnable()
        Try
            If Me.ugAvailableRO.Rows.Count > 0 Then
                btnROShiftRight.Enabled = True
                If Me.ugAvailableRO.Rows.Count > 1 Then
                    Me.btnROShiftRightAll.Enabled = True
                Else
                    Me.btnROShiftRightAll.Enabled = False
                End If
            Else
                Me.btnROShiftRight.Enabled = False
                Me.btnROShiftRightAll.Enabled = False
            End If

            If Me.ugAvailableRW.Rows.Count > 0 Then
                btnRWShiftRight.Enabled = True
                If Me.ugAvailableRW.Rows.Count > 1 Then
                    Me.btnRWShiftRightAll.Enabled = True
                Else
                    Me.btnRWShiftRightAll.Enabled = False
                End If
            Else
                Me.btnRWShiftRight.Enabled = False
                Me.btnRWShiftRightAll.Enabled = False
            End If

            If Me.ugRO.Rows.Count > 0 Then
                btnROShiftLeft.Enabled = True
                If Me.ugRO.Rows.Count > 1 Then
                    Me.btnROShiftLeftAll.Enabled = True
                Else
                    Me.btnROShiftLeftAll.Enabled = False
                End If
            Else
                Me.btnROShiftLeft.Enabled = False
                Me.btnROShiftLeftAll.Enabled = False
            End If

            If Me.ugRW.Rows.Count > 0 Then
                btnRWShiftLeft.Enabled = True
                If Me.ugRW.Rows.Count > 1 Then
                    Me.btnRWShiftLeftAll.Enabled = True
                Else
                    Me.btnRWShiftLeftAll.Enabled = False
                End If
            Else
                Me.btnRWShiftLeft.Enabled = False
                Me.btnRWShiftLeftAll.Enabled = False
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Sub CheckIfDataDirty()
        If oGroups.colIsDirty Then
            Dim msgRet As MsgBoxResult = MsgBox("The user group data for " & oGroups.Name & " has changed.  Do you wish to save the changes?", MsgBoxStyle.Question & MsgBoxStyle.YesNo, "Data Changed")
            If msgRet = MsgBoxResult.Yes Then
                oGroups.Flush(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If
            Else
                oGroups.Reset()
            End If
        End If
    End Sub
#End Region
#Region "UI Control Events"

    Private Sub txtDescription_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDescription.TextChanged
        If bolIsLoading Then Exit Sub
        oGroups.Description = txtDescription.Text
    End Sub

    Private Sub txtDescription_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDescription.Leave
        Try
            If txtDescription.Text <> String.Empty Then
                oGroups.Description = txtDescription.Text
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        If bolIsNewGroup Then
            Dim msgResult As MsgBoxResult = MsgBox("Would you like to save the new user group you are currently creating?", MsgBoxStyle.Question & MsgBoxStyle.YesNoCancel, "NEW USER GROUP")
            If msgResult = MsgBoxResult.No Then
                'dont save, continue creating new group
                oGroups.Remove(oGroups.Name)
                oGroups.colIsDirty = False
                oGroups.IsDirty = False
            ElseIf msgResult = MsgBoxResult.Yes Then
                If oGroups.ID <= 0 Then
                    oGroups.CreatedBy = MusterContainer.AppUser.ID
                Else
                    oGroups.ModifiedBy = MusterContainer.AppUser.ID
                End If
                oGroups.Save(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If
            ElseIf msgResult = MsgBoxResult.Cancel Then
                Exit Sub
            End If
        End If

        bolIsLoading = True
        bolIsNewGroup = True
        ClearGroupData()
        ClearScreenData()
        ComboGroupName.SelectedIndex = -1
        'oGroups.Add("New Group")
        'LoadScreenListviews()
        txtGroupName.ReadOnly = False
        txtGroupName.Text = ""
        txtDescription.Text = ""
        Me.txtGroupName.Focus()
        ComboGroupName.Enabled = False
        'LoadScreenListviews()
        bolIsLoading = False
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        
        oGroups.Reset()
        'ComboGroupName_Leave()
        If bolIsNewGroup Then bolIsNewGroup = False
        If bolIsNewGroup Then oGroups.Remove(oGroups.ID)
        If strPreviousGroupName <> String.Empty Then
            ComboGroupName.SelectedIndex = ComboGroupName.FindStringExact(strPreviousGroupName)
        Else
            ComboGroupName.SelectedIndex = 0
        End If
        GetGroupData()
        DisplayGroupData()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try

            If oGroups.ID <= 0 Then
                oGroups.CreatedBy = MusterContainer.AppUser.ID
            Else
                oGroups.ModifiedBy = MusterContainer.AppUser.ID
            End If
            oGroups.Save(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If
            If bolIsNewGroup Then
                bolIsNewGroup = False
                ComboGroupName.Enabled = True
                bolIsLoading = True
                Me.InitUserGroups()
                bolIsLoading = False
                ComboGroupName.SelectedIndex = ComboGroupName.FindStringExact(oGroups.Name)
            End If
            MsgBox("User Group Save Successful", 0, "MUSTER Data Access")
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        If oGroups.HasUsers() Then
            MsgBox("This user group has users associated to it. You can not delete this user group at this time.")
        Else
            If oGroups.HasCalendarEntries Then
                Dim MsgBoxResp As MsgBoxResult = MsgBox("There are associated calendar entries for this group!  Do you wish to delete the calendar entries before deleting the group?", MsgBoxStyle.Critical + MsgBoxStyle.YesNo)
                '
                '  If the user responds "Yes" then simply remove calendar entries and try again
                '
                If MsgBoxResp = MsgBoxResult.Yes Then
                    oGroups.DeleteCalendarEntries()
                    Me.btnDelete_Click(sender, e)
                End If
            Else
                Dim msgResult As MsgBoxResult = MsgBox("Are you sure you wish to DELETE the User Group: " & Me.ComboGroupName.Text & " ?", MsgBoxStyle.Question & MsgBoxStyle.YesNo, "DELETE USER GROUP")
                If msgResult = MsgBoxResult.Yes Then
                    oGroups.Deleted = True
                    For Each groupModuleRelInfo As MUSTER.Info.GroupModuleRelationInfo In oGroups.GroupModuleRelationCollection.Values
                        groupModuleRelInfo.Deleted = True
                    Next
                    If oGroups.ID <= 0 Then
                        oGroups.CreatedBy = MusterContainer.AppUser.ID
                    Else
                        oGroups.ModifiedBy = MusterContainer.AppUser.ID
                    End If
                    oGroups.Save(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If
                    MsgBox("User Group Delete Successful", 0, "MUSTER Data Access")
                    Me.InitUserGroups()
                    ClearScreenData()
                    Me.ClearGroupData()
                Else
                    'do nothing
                End If
            End If
        End If
    End Sub

    Private Sub ComboGroupName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboGroupName.SelectedIndexChanged
        '
        ' Dont try to do anything if the group combo is being loaded
        '
        If bolIsLoading Then Exit Sub
        'MsgBox("after" & ComboGroupName.Value)
        'If ComboGroupName.SelectedIndex > 0 Then Me.strPreviousGroupName = ComboGroupName.SelectedValue
        CheckIfDataDirty()
        'ComboGroupName_Leave()
        GetGroupData()
        DisplayGroupData()
    End Sub

    'Private Sub ComboGroupName_Leave()
    'If bolIsClosing Then Exit Sub
    'bolIsNewGroup = Not (ComboGroupName.FindStringExact(ComboGroupName.Text) > -1)
    'If bolIsNewGroup Then
    '    CheckIfDataDirty()
    '    ClearGroupData()
    '    bolIsNewGroup = False
    '    ComboGroupName.Enabled = True
    '    bolIsLoading = True
    '    Me.InitUserGroups()
    '    bolIsLoading = False
    'Else
    'ComboGroupName.SelectedIndex = ComboGroupName.FindStringExact(ComboGroupName.Text)
    'Me.btnDelete.Enabled = True
    'End If
    'GetGroupData()
    'DisplayGroupData()
    'End Sub

    Private Sub chkInactive_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkInactive.CheckedChanged
        '
        ' If loading data, then don't bother to update the object since we're loading from it
        '
        If bolIsLoading Then Exit Sub
        oGroups.Active = chkInactive.Checked
    End Sub

    Private Sub btnROShiftRight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnROShiftRight.Click
        Try
            If Not ugAvailableRO.Selected.Rows Is Nothing Then
                If ugAvailableRO.Selected.Rows.Count > 0 Then
                    For Each ugRow In ugAvailableRO.Selected.Rows
                        oGroups.GroupModuleRelationCollection.Item(oGroups.ID.ToString + "|" + ugRow.Cells("MODULE_ID").Text).ReadAccess = True
                        oGroups.GroupModuleRelationCollection.Item(oGroups.ID.ToString + "|" + ugRow.Cells("MODULE_ID").Text).WriteAccess = False
                    Next
                    SetSaveCancel(oGroups.IsDirty)
                    LoadScreenListviews()
                End If
            End If
            'If Me.lstAvailableFormsforRO.SelectedIndex > -1 Then
            '    oGroups.Screens.Retrieve(oGroups.Name & "|SCREENS|" & lstAvailableFormsforRO.SelectedItem.ToString & "|NONE").ProfileValue = "RO"
            '    Me.lstReadOnlyForms.Items.Add(Me.lstAvailableFormsforRO.SelectedItem)
            '    Me.lstAvailableFormsforRW.SelectedIndex = Me.lstAvailableFormsforRO.SelectedIndex
            '    Me.lstAvailableFormsforRO.Items.RemoveAt(Me.lstAvailableFormsforRO.SelectedIndex)
            '    Me.lstAvailableFormsforRW.Items.RemoveAt(Me.lstAvailableFormsforRW.SelectedIndex)
            '    LeftRightEnable()
            'End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnRWShiftRight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRWShiftRight.Click
        Try
            If Not ugAvailableRW.Selected.Rows Is Nothing Then
                If ugAvailableRW.Selected.Rows.Count > 0 Then
                    For Each ugRow In ugAvailableRW.Selected.Rows
                        oGroups.GroupModuleRelationCollection.Item(oGroups.ID.ToString + "|" + ugRow.Cells("MODULE_ID").Text).ReadAccess = True
                        oGroups.GroupModuleRelationCollection.Item(oGroups.ID.ToString + "|" + ugRow.Cells("MODULE_ID").Text).WriteAccess = True
                    Next
                    SetSaveCancel(oGroups.IsDirty)
                    LoadScreenListviews()
                End If
            End If
            'If Me.lstAvailableFormsforRW.SelectedIndex > -1 Then
            '    oGroups.Screens.Retrieve(oGroups.Name & "|SCREENS|" & lstAvailableFormsforRW.SelectedItem.ToString & "|NONE").ProfileValue = "RW"
            '    Me.lstReadWriteForms.Items.Add(Me.lstAvailableFormsforRW.SelectedItem)
            '    Me.lstAvailableFormsforRO.SelectedIndex = Me.lstAvailableFormsforRW.SelectedIndex
            '    Me.lstAvailableFormsforRO.Items.RemoveAt(Me.lstAvailableFormsforRO.SelectedIndex)
            '    Me.lstAvailableFormsforRW.Items.RemoveAt(Me.lstAvailableFormsforRW.SelectedIndex)
            '    LeftRightEnable()
            'End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnROShiftRightAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnROShiftRightAll.Click
        'Dim item As Object
        Try
            For Each ugRow In ugAvailableRO.Rows
                oGroups.GroupModuleRelationCollection.Item(oGroups.ID.ToString + "|" + ugRow.Cells("MODULE_ID").Text).ReadAccess = True
                oGroups.GroupModuleRelationCollection.Item(oGroups.ID.ToString + "|" + ugRow.Cells("MODULE_ID").Text).WriteAccess = False
            Next
            SetSaveCancel(oGroups.IsDirty)
            LoadScreenListviews()
            'While lstAvailableFormsforRO.Items.Count > 0
            '    lstAvailableFormsforRO.SelectedIndex = 0
            '    btnROShiftRight_Click(sender, e)
            'End While
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnRWShiftRightAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRWShiftRightAll.Click
        'Dim item As Object
        Try
            For Each ugRow In ugAvailableRW.Rows
                oGroups.GroupModuleRelationCollection.Item(oGroups.ID.ToString + "|" + ugRow.Cells("MODULE_ID").Text).ReadAccess = True
                oGroups.GroupModuleRelationCollection.Item(oGroups.ID.ToString + "|" + ugRow.Cells("MODULE_ID").Text).WriteAccess = True
            Next
            SetSaveCancel(oGroups.IsDirty)
            LoadScreenListviews()
            'While lstAvailableFormsforRW.Items.Count > 0
            '    lstAvailableFormsforRW.SelectedIndex = 0
            '    btnRWShiftRight_Click(sender, e)
            'End While
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnROShiftLeft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnROShiftLeft.Click
        Try
            If Not ugRO.Selected.Rows Is Nothing Then
                If ugRO.Selected.Rows.Count > 0 Then
                    For Each ugRow In ugRO.Selected.Rows
                        oGroups.GroupModuleRelationCollection.Item(oGroups.ID.ToString + "|" + ugRow.Cells("MODULE_ID").Text).ReadAccess = False
                        oGroups.GroupModuleRelationCollection.Item(oGroups.ID.ToString + "|" + ugRow.Cells("MODULE_ID").Text).WriteAccess = False
                    Next
                    SetSaveCancel(oGroups.IsDirty)
                    LoadScreenListviews()
                End If
            End If
            'If Me.lstReadOnlyForms.SelectedIndex > -1 Then
            '    oGroups.Screens.Retrieve(oGroups.Name & "|SCREENS|" & lstReadOnlyForms.SelectedItem.ToString & "|NONE").ProfileValue = "NONE"
            '    Me.lstAvailableFormsforRO.Items.Add(Me.lstReadOnlyForms.SelectedItem)
            '    Me.lstAvailableFormsforRW.Items.Add(Me.lstReadOnlyForms.SelectedItem)
            '    Me.lstReadOnlyForms.Items.RemoveAt(Me.lstReadOnlyForms.SelectedIndex)
            '    LeftRightEnable()
            'End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnRWShiftLeft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRWShiftLeft.Click
        Try
            If Not ugRW.Selected.Rows Is Nothing Then
                If ugRW.Selected.Rows.Count > 0 Then
                    For Each ugRow In ugRW.Selected.Rows
                        oGroups.GroupModuleRelationCollection.Item(oGroups.ID.ToString + "|" + ugRow.Cells("MODULE_ID").Text).ReadAccess = False
                        oGroups.GroupModuleRelationCollection.Item(oGroups.ID.ToString + "|" + ugRow.Cells("MODULE_ID").Text).WriteAccess = False
                    Next
                    SetSaveCancel(oGroups.IsDirty)
                    LoadScreenListviews()
                End If
            End If
            'If Me.lstReadWriteForms.SelectedIndex > -1 Then
            '    oGroups.Screens.Retrieve(oGroups.Name & "|SCREENS|" & lstReadWriteForms.SelectedItem.ToString & "|NONE").ProfileValue = "NONE"
            '    Me.lstAvailableFormsforRO.Items.Add(Me.lstReadWriteForms.SelectedItem)
            '    Me.lstAvailableFormsforRW.Items.Add(Me.lstReadWriteForms.SelectedItem)
            '    Me.lstReadWriteForms.Items.RemoveAt(Me.lstReadWriteForms.SelectedIndex)
            '    LeftRightEnable()
            'End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnROShiftLeftAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnROShiftLeftAll.Click
        'Dim item As Object
        Try
            For Each ugRow In ugRO.Rows
                oGroups.GroupModuleRelationCollection.Item(oGroups.ID.ToString + "|" + ugRow.Cells("MODULE_ID").Text).ReadAccess = False
                oGroups.GroupModuleRelationCollection.Item(oGroups.ID.ToString + "|" + ugRow.Cells("MODULE_ID").Text).WriteAccess = False
            Next
            SetSaveCancel(oGroups.IsDirty)
            LoadScreenListviews()
            'While lstReadOnlyForms.Items.Count > 0
            '    lstReadOnlyForms.SelectedIndex = 0
            '    btnROShiftLeft_Click(sender, e)
            'End While
            'LeftRightEnable()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnRWShiftLeftAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRWShiftLeftAll.Click
        'Dim item As Object
        Try
            For Each ugRow In ugRW.Rows
                oGroups.GroupModuleRelationCollection.Item(oGroups.ID.ToString + "|" + ugRow.Cells("MODULE_ID").Text).ReadAccess = False
                oGroups.GroupModuleRelationCollection.Item(oGroups.ID.ToString + "|" + ugRow.Cells("MODULE_ID").Text).WriteAccess = False
            Next
            SetSaveCancel(oGroups.IsDirty)
            LoadScreenListviews()
            'While lstReadWriteForms.Items.Count > 0
            '    lstReadWriteForms.SelectedIndex = 0
            '    btnRWShiftLeft_Click(sender, e)
            'End While
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    'Private Sub lstAvailableFormsforRO_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Try
    '        btnROShiftRight_Click(sender, e)
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub

    'Private Sub lstAvailableFormsforRW_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Try
    '        btnRWShiftRight_Click(sender, e)
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub

    'Private Sub lstReadOnlyForms_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Try
    '        btnROShiftLeft_Click(sender, e)
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub

    'Private Sub lstReadWriteForms_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Try
    '        btnRWShiftLeft_Click(sender, e)
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub

    'Private Sub txtDescription_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDescription.Enter
    '    IsEmptyGroupName()
    'End Sub

    'Private Sub chkInactive_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkInactive.Enter
    '    IsEmptyGroupName()
    'End Sub

    Private Sub txtGroupName_Leave(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtGroupName.Leave
        Dim strGroupName As String = txtGroupName.Text
        If txtGroupName.ReadOnly = False Then
            If txtGroupName.Text <> String.Empty Then
                If Not (Me.ComboGroupName.FindStringExact(Me.txtGroupName.Text) > -1) Then
                    'bolIsLoading = True
                    oGroups.Add(strGroupName)
                    oGroups.IsDirty = True
                    'bolIsNewGroup = True
                    'oGroups.Name = txtGroupName.Text
                    LoadScreenListviews()
                Else
                    MsgBox("A group with that name already exists.")
                    ComboGroupName.Text = Me.txtGroupName.Text
                    '    ComboUserName.SelectedIndex = ComboUserName.FindStringExact(strOldUserName)
                    bolIsNewGroup = False
                    bolIsLoading = False
                    oGroups.Retrieve(txtGroupName.Text)
                    ComboGroupName.SelectedIndex = ComboGroupName.FindStringExact(txtGroupName.Text)
                End If
            Else
                Dim msgResult As MsgBoxResult = MsgBox("Group name cannot be blank.Would you like to continue creating a new user group?", MsgBoxStyle.Question & MsgBoxStyle.YesNo, "NEW USER GROUP")
                If msgResult = MsgBoxResult.No Then
                    bolIsNewGroup = False
                    bolIsLoading = False
                    oGroups.Remove(oGroups.Name)
                    oGroups.colIsDirty = False
                    oGroups.IsDirty = False
                    oGroups.ResetCollection()
                    If strPreviousGroupName <> String.Empty Then
                        ComboGroupName.SelectedIndex = ComboGroupName.FindStringExact(strPreviousGroupName)
                    Else
                        ComboGroupName.SelectedIndex = 0
                    End If
                Else
                    Me.txtGroupName.Focus()
                    Exit Sub
                End If
            End If
        End If
        'LoadScreenListviews()
    End Sub

    Private Sub ugAvailableRO_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugAvailableRO.InitializeLayout
        SetupGrid(sender)
    End Sub

    Private Sub ugAvailableRW_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugAvailableRW.InitializeLayout
        SetupGrid(sender)
    End Sub

    Private Sub ugRO_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugRO.InitializeLayout
        SetupGrid(sender)
    End Sub

    Private Sub ugRW_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugRW.InitializeLayout
        SetupGrid(sender)
    End Sub
#End Region
#Region "External Event Handlers"
    'Private Sub ScreensChanged(ByVal bolState As Boolean) Handles oGroups.ScreensChanged
    '    Dim nInt As Int64
    '    If bolIsLoading Then Exit Sub
    '    SetSaveCancel(bolState)
    'End Sub

    Private Sub GroupChanged(ByVal bolState As Boolean) Handles oGroups.GroupChanged
        SetSaveCancel(bolState)
    End Sub

    Private Sub GroupsChanged(ByVal bolState As Boolean) Handles oGroups.GroupsChanged
        SetSaveCancel(bolState)
    End Sub

    Private Sub UserGroupError(ByVal strErr As String, ByVal strSource As String) Handles oGroups.GroupError
        MsgBox(strErr, MsgBoxStyle.Exclamation & MsgBoxStyle.OKOnly, strSource & " Error")
        bolErrorOccurred = True
    End Sub
#End Region

    'Private Sub btnClose_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.GotFocus
    '    bolIsClosing = True
    'End Sub
End Class
