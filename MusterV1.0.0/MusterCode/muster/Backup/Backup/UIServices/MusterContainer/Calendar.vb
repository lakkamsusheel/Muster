Public Class Calendar
    '-------------------------------------------------------------------------------
    ' MUSTER.MUSTER.Calendar
    '   Provides the mechanism for making manual calendar entries.
    '
    ' Copyright (C) 2004 CIBER, Inc.
    ' All rights reserved.
    '
    ' Release   Initials    Date        Description
    '  1.0        ??      8/??/04    Original class definition.
    '  1.1        JC      1/20/05    Altered to use Calendar object and to use
    '                                   information from MusterContainer and AppUser to 
    '                                   populate user ID and Groups.  Also added
    '                                   code to log user access and make form child
    '                                   of an MDI form if provided in the NEW().
    '
    ' 1.2          MR     1/27/05    Add functions to catch the Events for Data Validation. 
    '                                Modified Save method and Removed CalendarEntryValidations function.
    ' 1.3          AN     02/10/05   Integrated AppFlags new object model
    '-------------------------------------------------------------------------------
    Inherits System.Windows.Forms.Form
    Friend strMode As String
    Dim strUserID As String
    Dim nCalendarInfoId As Integer = 0
    Friend MyGuid As New System.Guid
    Private WithEvents oCalendar As MUSTER.BusinessLogic.pCalendar
    Friend bolValidationFailed As Boolean = False
    Friend bolLoading As Boolean = False
    Dim returnVal As String = String.Empty
    Friend dtUsers As DataTable

#Region " Windows Form Designer generated code "

    Public Sub New(Optional ByRef ParentForm As Windows.Forms.Form = Nothing)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        '
        '
        'Need to tell the AppUser that we've instantiated another Calendar Entry form...
        '
        If Not (ParentForm Is Nothing) Then
            Me.MdiParent = ParentForm
        End If
        MyGuid = System.Guid.NewGuid
        MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "WindowName", Me.Text)
        MusterContainer.AppUser.LogEntry("Calendar", MyGuid.ToString)
        '
        ' The following line enables all forms to detect the visible form in the MDI container
        '
        MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid)

        strUserID = MusterContainer.AppUser.ID

    End Sub

    Public Sub New(ByVal strUser As String, Optional ByRef ParentForm As Windows.Forms.Form = Nothing)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        '
        'Need to tell the AppUser that we've instantiated another Calendar Entry form...
        '
        If Not (ParentForm Is Nothing) Then
            Me.MdiParent = ParentForm
        End If
        MyGuid = System.Guid.NewGuid
        MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "WindowName", Me.Text)
        MusterContainer.AppUser.LogEntry("Calendar", MyGuid.ToString)
        '
        ' The following line enables all forms to detect the visible form in the MDI container
        '
        MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid)

        strUserID = strUser

    End Sub
    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        '
        ' Remove any values from the shared collection for this screen
        '
        MusterContainer.AppSemaphores.Remove(MyGuid.ToString)
        '
        ' Log the disposal of the form (exit from Calendar form)
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

    Friend WithEvents grpBxCalendarCategory As System.Windows.Forms.GroupBox
    Friend WithEvents RdCalCategoryToDo As System.Windows.Forms.RadioButton
    Friend WithEvents rdCalCategoryDueToMe As System.Windows.Forms.RadioButton
    Friend WithEvents grpBxCalendarTarget As System.Windows.Forms.GroupBox
    Friend WithEvents rdCalTargetUser As System.Windows.Forms.RadioButton
    Friend WithEvents rdCalTargetGroup As System.Windows.Forms.RadioButton
    Friend WithEvents cmbCalTargetGroup As System.Windows.Forms.ComboBox
    Friend WithEvents txtCalDescription As System.Windows.Forms.TextBox
    Friend WithEvents lblCalDescription As System.Windows.Forms.Label
    Friend WithEvents btnCalCancel As System.Windows.Forms.Button
    Friend WithEvents btnSaveCalEntry As System.Windows.Forms.Button
    Friend WithEvents pnlCalDueDate As System.Windows.Forms.Panel
    Friend WithEvents lblCalendarDueDate As System.Windows.Forms.Label
    Friend WithEvents pnlLblTaskDescription As System.Windows.Forms.Panel
    Friend WithEvents pnlTaskEntryButtons As System.Windows.Forms.Panel
    Friend WithEvents dtPickCalendarDueDateValue As System.Windows.Forms.DateTimePicker
    Friend WithEvents lstBoxUsers As System.Windows.Forms.ListBox
    Friend WithEvents pnlCalDescription As System.Windows.Forms.Panel
    Friend WithEvents lblUser As System.Windows.Forms.Label

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.grpBxCalendarCategory = New System.Windows.Forms.GroupBox
        Me.rdCalCategoryDueToMe = New System.Windows.Forms.RadioButton
        Me.RdCalCategoryToDo = New System.Windows.Forms.RadioButton
        Me.grpBxCalendarTarget = New System.Windows.Forms.GroupBox
        Me.cmbCalTargetGroup = New System.Windows.Forms.ComboBox
        Me.rdCalTargetGroup = New System.Windows.Forms.RadioButton
        Me.rdCalTargetUser = New System.Windows.Forms.RadioButton
        Me.lstBoxUsers = New System.Windows.Forms.ListBox
        Me.txtCalDescription = New System.Windows.Forms.TextBox
        Me.lblCalDescription = New System.Windows.Forms.Label
        Me.btnCalCancel = New System.Windows.Forms.Button
        Me.btnSaveCalEntry = New System.Windows.Forms.Button
        Me.pnlTaskEntryButtons = New System.Windows.Forms.Panel
        Me.pnlLblTaskDescription = New System.Windows.Forms.Panel
        Me.pnlCalDescription = New System.Windows.Forms.Panel
        Me.pnlCalDueDate = New System.Windows.Forms.Panel
        Me.dtPickCalendarDueDateValue = New System.Windows.Forms.DateTimePicker
        Me.lblCalendarDueDate = New System.Windows.Forms.Label
        Me.lblUser = New System.Windows.Forms.Label
        Me.grpBxCalendarCategory.SuspendLayout()
        Me.grpBxCalendarTarget.SuspendLayout()
        Me.pnlTaskEntryButtons.SuspendLayout()
        Me.pnlLblTaskDescription.SuspendLayout()
        Me.pnlCalDescription.SuspendLayout()
        Me.pnlCalDueDate.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpBxCalendarCategory
        '
        Me.grpBxCalendarCategory.Controls.Add(Me.rdCalCategoryDueToMe)
        Me.grpBxCalendarCategory.Controls.Add(Me.RdCalCategoryToDo)
        Me.grpBxCalendarCategory.Dock = System.Windows.Forms.DockStyle.Top
        Me.grpBxCalendarCategory.Location = New System.Drawing.Point(0, 35)
        Me.grpBxCalendarCategory.Name = "grpBxCalendarCategory"
        Me.grpBxCalendarCategory.Size = New System.Drawing.Size(552, 80)
        Me.grpBxCalendarCategory.TabIndex = 0
        Me.grpBxCalendarCategory.TabStop = False
        Me.grpBxCalendarCategory.Text = "Category"
        '
        'rdCalCategoryDueToMe
        '
        Me.rdCalCategoryDueToMe.Location = New System.Drawing.Point(32, 48)
        Me.rdCalCategoryDueToMe.Name = "rdCalCategoryDueToMe"
        Me.rdCalCategoryDueToMe.Size = New System.Drawing.Size(80, 17)
        Me.rdCalCategoryDueToMe.TabIndex = 1
        Me.rdCalCategoryDueToMe.Text = "Due To Me"
        '
        'RdCalCategoryToDo
        '
        Me.RdCalCategoryToDo.Location = New System.Drawing.Point(32, 24)
        Me.RdCalCategoryToDo.Name = "RdCalCategoryToDo"
        Me.RdCalCategoryToDo.Size = New System.Drawing.Size(64, 15)
        Me.RdCalCategoryToDo.TabIndex = 0
        Me.RdCalCategoryToDo.Text = "To Do"
        '
        'grpBxCalendarTarget
        '
        Me.grpBxCalendarTarget.Controls.Add(Me.lblUser)
        Me.grpBxCalendarTarget.Controls.Add(Me.cmbCalTargetGroup)
        Me.grpBxCalendarTarget.Controls.Add(Me.rdCalTargetGroup)
        Me.grpBxCalendarTarget.Controls.Add(Me.rdCalTargetUser)
        Me.grpBxCalendarTarget.Controls.Add(Me.lstBoxUsers)
        Me.grpBxCalendarTarget.Dock = System.Windows.Forms.DockStyle.Top
        Me.grpBxCalendarTarget.Location = New System.Drawing.Point(0, 115)
        Me.grpBxCalendarTarget.Name = "grpBxCalendarTarget"
        Me.grpBxCalendarTarget.Size = New System.Drawing.Size(552, 176)
        Me.grpBxCalendarTarget.TabIndex = 3
        Me.grpBxCalendarTarget.TabStop = False
        Me.grpBxCalendarTarget.Text = "Target"
        '
        'cmbCalTargetGroup
        '
        Me.cmbCalTargetGroup.Enabled = False
        Me.cmbCalTargetGroup.Location = New System.Drawing.Point(104, 24)
        Me.cmbCalTargetGroup.Name = "cmbCalTargetGroup"
        Me.cmbCalTargetGroup.Size = New System.Drawing.Size(192, 21)
        Me.cmbCalTargetGroup.TabIndex = 0
        '
        'rdCalTargetGroup
        '
        Me.rdCalTargetGroup.Location = New System.Drawing.Point(32, 24)
        Me.rdCalTargetGroup.Name = "rdCalTargetGroup"
        Me.rdCalTargetGroup.Size = New System.Drawing.Size(56, 15)
        Me.rdCalTargetGroup.TabIndex = 1
        Me.rdCalTargetGroup.Text = "Group"
        '
        'rdCalTargetUser
        '
        Me.rdCalTargetUser.Checked = True
        Me.rdCalTargetUser.Location = New System.Drawing.Point(32, 56)
        Me.rdCalTargetUser.Name = "rdCalTargetUser"
        Me.rdCalTargetUser.Size = New System.Drawing.Size(59, 15)
        Me.rdCalTargetUser.TabIndex = 2
        Me.rdCalTargetUser.TabStop = True
        Me.rdCalTargetUser.Text = "User(s)"
        '
        'lstBoxUsers
        '
        Me.lstBoxUsers.Location = New System.Drawing.Point(104, 56)
        Me.lstBoxUsers.Name = "lstBoxUsers"
        Me.lstBoxUsers.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.lstBoxUsers.Size = New System.Drawing.Size(192, 108)
        Me.lstBoxUsers.TabIndex = 3
        '
        'txtCalDescription
        '
        Me.txtCalDescription.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtCalDescription.Location = New System.Drawing.Point(0, 32)
        Me.txtCalDescription.MaxLength = 200
        Me.txtCalDescription.Multiline = True
        Me.txtCalDescription.Name = "txtCalDescription"
        Me.txtCalDescription.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtCalDescription.Size = New System.Drawing.Size(552, 74)
        Me.txtCalDescription.TabIndex = 0
        Me.txtCalDescription.Text = ""
        '
        'lblCalDescription
        '
        Me.lblCalDescription.Location = New System.Drawing.Point(8, 8)
        Me.lblCalDescription.Name = "lblCalDescription"
        Me.lblCalDescription.Size = New System.Drawing.Size(72, 20)
        Me.lblCalDescription.TabIndex = 5
        Me.lblCalDescription.Text = "Description"
        '
        'btnCalCancel
        '
        Me.btnCalCancel.Location = New System.Drawing.Point(96, 8)
        Me.btnCalCancel.Name = "btnCalCancel"
        Me.btnCalCancel.Size = New System.Drawing.Size(75, 20)
        Me.btnCalCancel.TabIndex = 1
        Me.btnCalCancel.Text = "Cancel"
        '
        'btnSaveCalEntry
        '
        Me.btnSaveCalEntry.Location = New System.Drawing.Point(8, 8)
        Me.btnSaveCalEntry.Name = "btnSaveCalEntry"
        Me.btnSaveCalEntry.Size = New System.Drawing.Size(75, 20)
        Me.btnSaveCalEntry.TabIndex = 0
        Me.btnSaveCalEntry.Text = "Save Entry"
        '
        'pnlTaskEntryButtons
        '
        Me.pnlTaskEntryButtons.Controls.Add(Me.btnCalCancel)
        Me.pnlTaskEntryButtons.Controls.Add(Me.btnSaveCalEntry)
        Me.pnlTaskEntryButtons.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlTaskEntryButtons.Location = New System.Drawing.Point(0, 397)
        Me.pnlTaskEntryButtons.Name = "pnlTaskEntryButtons"
        Me.pnlTaskEntryButtons.Size = New System.Drawing.Size(552, 40)
        Me.pnlTaskEntryButtons.TabIndex = 14
        '
        'pnlLblTaskDescription
        '
        Me.pnlLblTaskDescription.Controls.Add(Me.lblCalDescription)
        Me.pnlLblTaskDescription.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlLblTaskDescription.Location = New System.Drawing.Point(0, 0)
        Me.pnlLblTaskDescription.Name = "pnlLblTaskDescription"
        Me.pnlLblTaskDescription.Size = New System.Drawing.Size(552, 32)
        Me.pnlLblTaskDescription.TabIndex = 12
        '
        'pnlCalDescription
        '
        Me.pnlCalDescription.Controls.Add(Me.txtCalDescription)
        Me.pnlCalDescription.Controls.Add(Me.pnlLblTaskDescription)
        Me.pnlCalDescription.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlCalDescription.Location = New System.Drawing.Point(0, 291)
        Me.pnlCalDescription.Name = "pnlCalDescription"
        Me.pnlCalDescription.Size = New System.Drawing.Size(552, 106)
        Me.pnlCalDescription.TabIndex = 10
        '
        'pnlCalDueDate
        '
        Me.pnlCalDueDate.Controls.Add(Me.dtPickCalendarDueDateValue)
        Me.pnlCalDueDate.Controls.Add(Me.lblCalendarDueDate)
        Me.pnlCalDueDate.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlCalDueDate.Location = New System.Drawing.Point(0, 0)
        Me.pnlCalDueDate.Name = "pnlCalDueDate"
        Me.pnlCalDueDate.Size = New System.Drawing.Size(552, 35)
        Me.pnlCalDueDate.TabIndex = 9
        '
        'dtPickCalendarDueDateValue
        '
        Me.dtPickCalendarDueDateValue.Checked = False
        Me.dtPickCalendarDueDateValue.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickCalendarDueDateValue.Location = New System.Drawing.Point(72, 8)
        Me.dtPickCalendarDueDateValue.Name = "dtPickCalendarDueDateValue"
        Me.dtPickCalendarDueDateValue.ShowCheckBox = True
        Me.dtPickCalendarDueDateValue.Size = New System.Drawing.Size(96, 20)
        Me.dtPickCalendarDueDateValue.TabIndex = 0
        '
        'lblCalendarDueDate
        '
        Me.lblCalendarDueDate.Location = New System.Drawing.Point(8, 8)
        Me.lblCalendarDueDate.Name = "lblCalendarDueDate"
        Me.lblCalendarDueDate.Size = New System.Drawing.Size(56, 20)
        Me.lblCalendarDueDate.TabIndex = 0
        Me.lblCalendarDueDate.Text = "Due Date"
        Me.lblCalendarDueDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblUser
        '
        Me.lblUser.Location = New System.Drawing.Point(104, 56)
        Me.lblUser.Name = "lblUser"
        Me.lblUser.Size = New System.Drawing.Size(384, 23)
        Me.lblUser.TabIndex = 4
        '
        'Calendar
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(552, 437)
        Me.Controls.Add(Me.pnlCalDescription)
        Me.Controls.Add(Me.grpBxCalendarTarget)
        Me.Controls.Add(Me.pnlTaskEntryButtons)
        Me.Controls.Add(Me.grpBxCalendarCategory)
        Me.Controls.Add(Me.pnlCalDueDate)
        Me.Name = "Calendar"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "New Calender Entry"
        Me.grpBxCalendarCategory.ResumeLayout(False)
        Me.grpBxCalendarTarget.ResumeLayout(False)
        Me.pnlTaskEntryButtons.ResumeLayout(False)
        Me.pnlLblTaskDescription.ResumeLayout(False)
        Me.pnlCalDescription.ResumeLayout(False)
        Me.pnlCalDueDate.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub Calendar_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            bolLoading = True
            If strMode = "ADD" Then
                UIUtilsGen.CreateEmptyFormatDatePicker(dtPickCalendarDueDateValue)
            Else
                nCalendarInfoId = Me.txtCalDescription.Tag
            End If
            'rdCalTargetUser.Text = "User : " & strUserID
            dtPickCalendarDueDateValue.Focus()
            oCalendar = MusterContainer.pCalendar

            cmbCalTargetGroup.DataSource = MusterContainer.AppUser.ListAllGroups
            cmbCalTargetGroup.DisplayMember = "USER_GROUP"
            cmbCalTargetGroup.ValueMember = "USER_GROUP"
            cmbCalTargetGroup.SelectedIndex = -1
            If cmbCalTargetGroup.SelectedIndex <> -1 Then
                cmbCalTargetGroup.SelectedIndex = -1
            End If

            If strMode = "ADD" Then
                lstBoxUsers.Visible = True
                lblUser.Visible = False
                dtUsers = MusterContainer.AppUser.ListAllUsers(True, False)
                dtUsers.DefaultView.Sort = "USERNAME"
                dtUsers.DefaultView.RowFilter = "LEN(USERNAME) > 0"
                lstBoxUsers.DataSource = dtUsers.DefaultView
                lstBoxUsers.DisplayMember = "USERNAME"
                lstBoxUsers.ValueMember = "USER_ID"
            Else
                lstBoxUsers.Visible = False
                lblUser.Visible = True
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            bolLoading = False
        End Try
    End Sub
    Private Sub btnSaveCalEntry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveCalEntry.Click

        Try
            Dim dtTemp As Date
            Dim calEntry As MUSTER.Info.CalendarInfo
            Dim strUser As String
            If rdCalTargetUser.Checked And strMode = "ADD" Then
                For i As Integer = 0 To lstBoxUsers.SelectedItems.Count - 1
                    strUser = lstBoxUsers.SelectedItems(i)(0)
                    calEntry = New MUSTER.Info.CalendarInfo(0, _
                                IIf(Me.dtPickCalendarDueDateValue.Text = "__/__/____", dtTemp, Me.dtPickCalendarDueDateValue.Text), _
                                IIf(Me.dtPickCalendarDueDateValue.Text = "__/__/____", dtTemp, Me.dtPickCalendarDueDateValue.Text), _
                                0, _
                                Me.txtCalDescription.Text, _
                                strUser, _
                                MusterContainer.AppUser.ID, _
                                String.Empty, _
                                Me.rdCalCategoryDueToMe.Checked, _
                                Me.RdCalCategoryToDo.Checked, _
                                False, _
                                False, _
                                IIf(strMode = "ADD", MusterContainer.AppUser.ID, String.Empty), _
                                IIf(strMode = "ADD", Now, CDate("01/01/0001")), _
                                MusterContainer.AppUser.ID, _
                                Now)
                    SaveCalEntry(calEntry)
                Next
            Else
                calEntry = New MUSTER.Info.CalendarInfo(IIf(strMode = "ADD", 0, nCalendarInfoId), _
                            IIf(Me.dtPickCalendarDueDateValue.Text = "__/__/____", dtTemp, Me.dtPickCalendarDueDateValue.Text), _
                            IIf(Me.dtPickCalendarDueDateValue.Text = "__/__/____", dtTemp, Me.dtPickCalendarDueDateValue.Text), _
                            0, _
                            Me.txtCalDescription.Text, _
                            IIf(rdCalTargetUser.Checked, lblUser.Text, String.Empty), _
                            MusterContainer.AppUser.ID, _
                            IIf(rdCalTargetGroup.Checked, cmbCalTargetGroup.Text, String.Empty), _
                            Me.rdCalCategoryDueToMe.Checked, _
                            Me.RdCalCategoryToDo.Checked, _
                            False, _
                            False, _
                            IIf(strMode = "ADD", MusterContainer.AppUser.ID, String.Empty), _
                            IIf(strMode = "ADD", Now, CDate("01/01/0001")), _
                            MusterContainer.AppUser.ID, _
                            Now)
                SaveCalEntry(calEntry)
            End If

            If Not bolValidationFailed Then
                Me.Close()
            Else
                bolValidationFailed = False
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub SaveCalEntry(ByVal calEntry As MUSTER.Info.CalendarInfo)
        If nCalendarInfoId <= 0 Then
            calEntry.CreatedBy = MusterContainer.AppUser.ID
        Else
            calEntry.ModifiedBy = MusterContainer.AppUser.ID
        End If
        MusterContainer.pCalendar.Add(calEntry)
        MusterContainer.pCalendar.Save(UIUtilsGen.ModuleID.Global, MusterContainer.AppUser.UserKey, returnVal)
        If Not UIUtilsGen.HasRights(returnVal) Then
            Exit Sub
        End If

        If nCalendarInfoId > 0 Then
            nCalendarInfoId = MusterContainer.pCalendar.CalendarId
            MusterContainer.pFlag.FlagsCol = New MUSTER.Info.FlagsCollection
            ' if there are flags associated with the calendar entry, update the due date and description
            MusterContainer.pFlag.RetrieveFlags(, , , , , nCalendarInfoId)
            For Each flagInfo As MUSTER.Info.FlagInfo In MusterContainer.pFlag.FlagsCol.Values
                flagInfo.FlagDescription = Me.txtCalDescription.Text
                flagInfo.DueDate = UIUtilsGen.GetDatePickerValue(dtPickCalendarDueDateValue)
            Next
            MusterContainer.pFlag.Flush()
            MusterContainer.pFlag.FlagsCol = New MUSTER.Info.FlagsCollection
        End If
    End Sub
    Private Sub rdCalTargetGroup_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdCalTargetGroup.CheckedChanged
        Try
            If rdCalTargetGroup.Checked = True Then
                cmbCalTargetGroup.Enabled = True
                lstBoxUsers.Enabled = False
            Else
                cmbCalTargetGroup.Enabled = False
                lstBoxUsers.Enabled = True
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCalCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCalCancel.Click
        Try
            Me.Close()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub dtPickCalendarDueDateValue_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickCalendarDueDateValue.ValueChanged
        Try
            UIUtilsGen.ToggleDateFormat(dtPickCalendarDueDateValue)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub dtPickCalendarDueDateValue_CloseUp(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickCalendarDueDateValue.CloseUp
        Try
            Dim dtPick As DateTimePicker
            dtPick = CType(sender, DateTimePicker)
            Dim dtPickValue As Date = dtPick.Value
            If dtPick.Format <> DateTimePickerFormat.Short Then
                dtPick.Format = DateTimePickerFormat.Short
                dtPick.Value = dtPickValue
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub CalendarChanged(ByVal bolDirty As Boolean) Handles oCalendar.CalendarChanged
        If bolLoading Then Exit Sub
        If bolDirty Then
            btnSaveCalEntry.Enabled = True
        End If
    End Sub
    Private Sub CalendarError(ByVal strMsg As String) Handles oCalendar.evtCalErr
        MsgBox("Validation Failed : " + strMsg)
        bolValidationFailed = True
    End Sub

End Class
