Public Class LicenseeManagement
    Inherits System.Windows.Forms.Form

#Region "User Defined Variables"
    Public WithEvents pLicen As MUSTER.BusinessLogic.pLicensee
    Dim bolLoading As Boolean = True
    Private dsRRE As DataSet
    Private ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
    Dim oLetter As New Reg_Letters
    Dim returnVal As String = String.Empty
    Dim dtNull As DateTime = CDate("01/01/0001")
    Dim dtReminder As DateTime = DateAdd(DateInterval.Day, -90, Today).Date
#End Region

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
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
    Friend WithEvents tabCntrlLicenseeMgmt As System.Windows.Forms.TabControl
    Friend WithEvents tbPageRenewals As System.Windows.Forms.TabPage
    Friend WithEvents tbPageReminders As System.Windows.Forms.TabPage
    Friend WithEvents tbPageExpirations As System.Windows.Forms.TabPage
    Friend WithEvents pnlRenewalsTop As System.Windows.Forms.Panel
    Friend WithEvents pnlRenewalsBottom As System.Windows.Forms.Panel
    Friend WithEvents btnSelectAllforRenewal As System.Windows.Forms.Button
    Friend WithEvents pnlRemindersBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlRemindersTop As System.Windows.Forms.Panel
    Friend WithEvents pnlRemindersContainer As System.Windows.Forms.Panel
    Friend WithEvents ugReminders As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnGenerateInfoNeededLetters As System.Windows.Forms.Button
    Friend WithEvents btnSelectAllforReminding As System.Windows.Forms.Button
    Friend WithEvents pnlExpirationsBottom As System.Windows.Forms.Panel
    Friend WithEvents btnGenerateExpirationLetters As System.Windows.Forms.Button
    Friend WithEvents pnlExpirationsTop As System.Windows.Forms.Panel
    Friend WithEvents btnSelectAllforExpiring As System.Windows.Forms.Button
    Friend WithEvents pnlExpirationsContainer As System.Windows.Forms.Panel
    Friend WithEvents ugExpirations As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlRenewalsContainer As System.Windows.Forms.Panel
    Friend WithEvents ugRenewals As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnInfoLetter As System.Windows.Forms.Button
    Friend WithEvents btnClearAllforReminding As System.Windows.Forms.Button
    Friend WithEvents btnClearAllforExpiring As System.Windows.Forms.Button
    Friend WithEvents btnClearAllforRenewal As System.Windows.Forms.Button
    Friend WithEvents lblReminderDate As System.Windows.Forms.Label
    Friend WithEvents dtPickReminder As System.Windows.Forms.DateTimePicker
    Friend WithEvents btnRefresh As System.Windows.Forms.Button
    Friend WithEvents rBtnLetterGenYes As System.Windows.Forms.RadioButton
    Friend WithEvents lblLetterGenerated As System.Windows.Forms.Label
    Friend WithEvents rBtnLetterGenNo As System.Windows.Forms.RadioButton
    Friend WithEvents rBtnLetterGenEither As System.Windows.Forms.RadioButton
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim Appearance1 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance2 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance3 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance4 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance5 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Dim Appearance6 As Infragistics.Win.Appearance = New Infragistics.Win.Appearance
        Me.tabCntrlLicenseeMgmt = New System.Windows.Forms.TabControl
        Me.tbPageReminders = New System.Windows.Forms.TabPage
        Me.pnlRemindersContainer = New System.Windows.Forms.Panel
        Me.ugReminders = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlRemindersBottom = New System.Windows.Forms.Panel
        Me.btnGenerateInfoNeededLetters = New System.Windows.Forms.Button
        Me.pnlRemindersTop = New System.Windows.Forms.Panel
        Me.btnRefresh = New System.Windows.Forms.Button
        Me.lblReminderDate = New System.Windows.Forms.Label
        Me.dtPickReminder = New System.Windows.Forms.DateTimePicker
        Me.btnClearAllforReminding = New System.Windows.Forms.Button
        Me.btnSelectAllforReminding = New System.Windows.Forms.Button
        Me.rBtnLetterGenYes = New System.Windows.Forms.RadioButton
        Me.lblLetterGenerated = New System.Windows.Forms.Label
        Me.rBtnLetterGenNo = New System.Windows.Forms.RadioButton
        Me.rBtnLetterGenEither = New System.Windows.Forms.RadioButton
        Me.tbPageExpirations = New System.Windows.Forms.TabPage
        Me.pnlExpirationsContainer = New System.Windows.Forms.Panel
        Me.ugExpirations = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlExpirationsBottom = New System.Windows.Forms.Panel
        Me.btnGenerateExpirationLetters = New System.Windows.Forms.Button
        Me.pnlExpirationsTop = New System.Windows.Forms.Panel
        Me.btnClearAllforExpiring = New System.Windows.Forms.Button
        Me.btnSelectAllforExpiring = New System.Windows.Forms.Button
        Me.tbPageRenewals = New System.Windows.Forms.TabPage
        Me.pnlRenewalsContainer = New System.Windows.Forms.Panel
        Me.ugRenewals = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlRenewalsBottom = New System.Windows.Forms.Panel
        Me.btnInfoLetter = New System.Windows.Forms.Button
        Me.pnlRenewalsTop = New System.Windows.Forms.Panel
        Me.btnClearAllforRenewal = New System.Windows.Forms.Button
        Me.btnSelectAllforRenewal = New System.Windows.Forms.Button
        Me.tabCntrlLicenseeMgmt.SuspendLayout()
        Me.tbPageReminders.SuspendLayout()
        Me.pnlRemindersContainer.SuspendLayout()
        CType(Me.ugReminders, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlRemindersBottom.SuspendLayout()
        Me.pnlRemindersTop.SuspendLayout()
        Me.tbPageExpirations.SuspendLayout()
        Me.pnlExpirationsContainer.SuspendLayout()
        CType(Me.ugExpirations, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlExpirationsBottom.SuspendLayout()
        Me.pnlExpirationsTop.SuspendLayout()
        Me.tbPageRenewals.SuspendLayout()
        Me.pnlRenewalsContainer.SuspendLayout()
        CType(Me.ugRenewals, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlRenewalsBottom.SuspendLayout()
        Me.pnlRenewalsTop.SuspendLayout()
        Me.SuspendLayout()
        '
        'tabCntrlLicenseeMgmt
        '
        Me.tabCntrlLicenseeMgmt.Controls.Add(Me.tbPageReminders)
        Me.tabCntrlLicenseeMgmt.Controls.Add(Me.tbPageExpirations)
        Me.tabCntrlLicenseeMgmt.Controls.Add(Me.tbPageRenewals)
        Me.tabCntrlLicenseeMgmt.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tabCntrlLicenseeMgmt.Location = New System.Drawing.Point(0, 0)
        Me.tabCntrlLicenseeMgmt.Name = "tabCntrlLicenseeMgmt"
        Me.tabCntrlLicenseeMgmt.SelectedIndex = 0
        Me.tabCntrlLicenseeMgmt.Size = New System.Drawing.Size(824, 534)
        Me.tabCntrlLicenseeMgmt.TabIndex = 0
        '
        'tbPageReminders
        '
        Me.tbPageReminders.Controls.Add(Me.pnlRemindersContainer)
        Me.tbPageReminders.Controls.Add(Me.pnlRemindersBottom)
        Me.tbPageReminders.Controls.Add(Me.pnlRemindersTop)
        Me.tbPageReminders.Location = New System.Drawing.Point(4, 22)
        Me.tbPageReminders.Name = "tbPageReminders"
        Me.tbPageReminders.Size = New System.Drawing.Size(816, 508)
        Me.tbPageReminders.TabIndex = 1
        Me.tbPageReminders.Text = "Reminders"
        '
        'pnlRemindersContainer
        '
        Me.pnlRemindersContainer.Controls.Add(Me.ugReminders)
        Me.pnlRemindersContainer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlRemindersContainer.Location = New System.Drawing.Point(0, 40)
        Me.pnlRemindersContainer.Name = "pnlRemindersContainer"
        Me.pnlRemindersContainer.Size = New System.Drawing.Size(816, 420)
        Me.pnlRemindersContainer.TabIndex = 1
        '
        'ugReminders
        '
        Me.ugReminders.Cursor = System.Windows.Forms.Cursors.Default
        Appearance1.BorderColor = System.Drawing.SystemColors.ControlLight
        Me.ugReminders.DisplayLayout.Override.CellAppearance = Appearance1
        Me.ugReminders.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Appearance2.BorderColor = System.Drawing.SystemColors.ControlLight
        Me.ugReminders.DisplayLayout.Override.RowAppearance = Appearance2
        Me.ugReminders.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugReminders.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugReminders.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugReminders.Location = New System.Drawing.Point(0, 0)
        Me.ugReminders.Name = "ugReminders"
        Me.ugReminders.Size = New System.Drawing.Size(816, 420)
        Me.ugReminders.TabIndex = 0
        '
        'pnlRemindersBottom
        '
        Me.pnlRemindersBottom.Controls.Add(Me.btnGenerateInfoNeededLetters)
        Me.pnlRemindersBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlRemindersBottom.Location = New System.Drawing.Point(0, 460)
        Me.pnlRemindersBottom.Name = "pnlRemindersBottom"
        Me.pnlRemindersBottom.Size = New System.Drawing.Size(816, 48)
        Me.pnlRemindersBottom.TabIndex = 2
        '
        'btnGenerateInfoNeededLetters
        '
        Me.btnGenerateInfoNeededLetters.Location = New System.Drawing.Point(8, 8)
        Me.btnGenerateInfoNeededLetters.Name = "btnGenerateInfoNeededLetters"
        Me.btnGenerateInfoNeededLetters.Size = New System.Drawing.Size(104, 32)
        Me.btnGenerateInfoNeededLetters.TabIndex = 0
        Me.btnGenerateInfoNeededLetters.Text = "Generate Letters"
        '
        'pnlRemindersTop
        '
        Me.pnlRemindersTop.Controls.Add(Me.btnRefresh)
        Me.pnlRemindersTop.Controls.Add(Me.lblReminderDate)
        Me.pnlRemindersTop.Controls.Add(Me.dtPickReminder)
        Me.pnlRemindersTop.Controls.Add(Me.btnClearAllforReminding)
        Me.pnlRemindersTop.Controls.Add(Me.btnSelectAllforReminding)
        Me.pnlRemindersTop.Controls.Add(Me.rBtnLetterGenYes)
        Me.pnlRemindersTop.Controls.Add(Me.lblLetterGenerated)
        Me.pnlRemindersTop.Controls.Add(Me.rBtnLetterGenNo)
        Me.pnlRemindersTop.Controls.Add(Me.rBtnLetterGenEither)
        Me.pnlRemindersTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlRemindersTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlRemindersTop.Name = "pnlRemindersTop"
        Me.pnlRemindersTop.Size = New System.Drawing.Size(816, 40)
        Me.pnlRemindersTop.TabIndex = 0
        '
        'btnRefresh
        '
        Me.btnRefresh.Location = New System.Drawing.Point(640, 8)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.TabIndex = 5
        Me.btnRefresh.Text = "Refresh"
        '
        'lblReminderDate
        '
        Me.lblReminderDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblReminderDate.Location = New System.Drawing.Point(176, 8)
        Me.lblReminderDate.Name = "lblReminderDate"
        Me.lblReminderDate.TabIndex = 4
        Me.lblReminderDate.Text = "Expiration Date:"
        Me.lblReminderDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dtPickReminder
        '
        Me.dtPickReminder.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickReminder.Location = New System.Drawing.Point(280, 9)
        Me.dtPickReminder.Name = "dtPickReminder"
        Me.dtPickReminder.ShowCheckBox = True
        Me.dtPickReminder.Size = New System.Drawing.Size(96, 20)
        Me.dtPickReminder.TabIndex = 3
        '
        'btnClearAllforReminding
        '
        Me.btnClearAllforReminding.Location = New System.Drawing.Point(88, 8)
        Me.btnClearAllforReminding.Name = "btnClearAllforReminding"
        Me.btnClearAllforReminding.Size = New System.Drawing.Size(72, 24)
        Me.btnClearAllforReminding.TabIndex = 2
        Me.btnClearAllforReminding.Text = "Clear All"
        '
        'btnSelectAllforReminding
        '
        Me.btnSelectAllforReminding.Location = New System.Drawing.Point(8, 8)
        Me.btnSelectAllforReminding.Name = "btnSelectAllforReminding"
        Me.btnSelectAllforReminding.Size = New System.Drawing.Size(72, 24)
        Me.btnSelectAllforReminding.TabIndex = 0
        Me.btnSelectAllforReminding.Text = "Select All"
        '
        'rBtnLetterGenYes
        '
        Me.rBtnLetterGenYes.Location = New System.Drawing.Point(488, 8)
        Me.rBtnLetterGenYes.Name = "rBtnLetterGenYes"
        Me.rBtnLetterGenYes.Size = New System.Drawing.Size(48, 24)
        Me.rBtnLetterGenYes.TabIndex = 6
        Me.rBtnLetterGenYes.Text = "Yes"
        '
        'lblLetterGenerated
        '
        Me.lblLetterGenerated.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLetterGenerated.Location = New System.Drawing.Point(384, 8)
        Me.lblLetterGenerated.Name = "lblLetterGenerated"
        Me.lblLetterGenerated.TabIndex = 4
        Me.lblLetterGenerated.Text = "Letter Generated:"
        Me.lblLetterGenerated.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'rBtnLetterGenNo
        '
        Me.rBtnLetterGenNo.Checked = True
        Me.rBtnLetterGenNo.Location = New System.Drawing.Point(536, 8)
        Me.rBtnLetterGenNo.Name = "rBtnLetterGenNo"
        Me.rBtnLetterGenNo.Size = New System.Drawing.Size(48, 24)
        Me.rBtnLetterGenNo.TabIndex = 6
        Me.rBtnLetterGenNo.TabStop = True
        Me.rBtnLetterGenNo.Text = "No"
        '
        'rBtnLetterGenEither
        '
        Me.rBtnLetterGenEither.Location = New System.Drawing.Point(584, 8)
        Me.rBtnLetterGenEither.Name = "rBtnLetterGenEither"
        Me.rBtnLetterGenEither.Size = New System.Drawing.Size(56, 24)
        Me.rBtnLetterGenEither.TabIndex = 6
        Me.rBtnLetterGenEither.Text = "Either"
        '
        'tbPageExpirations
        '
        Me.tbPageExpirations.Controls.Add(Me.pnlExpirationsContainer)
        Me.tbPageExpirations.Controls.Add(Me.pnlExpirationsBottom)
        Me.tbPageExpirations.Controls.Add(Me.pnlExpirationsTop)
        Me.tbPageExpirations.Location = New System.Drawing.Point(4, 22)
        Me.tbPageExpirations.Name = "tbPageExpirations"
        Me.tbPageExpirations.Size = New System.Drawing.Size(816, 508)
        Me.tbPageExpirations.TabIndex = 2
        Me.tbPageExpirations.Text = "Expirations"
        '
        'pnlExpirationsContainer
        '
        Me.pnlExpirationsContainer.Controls.Add(Me.ugExpirations)
        Me.pnlExpirationsContainer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlExpirationsContainer.Location = New System.Drawing.Point(0, 40)
        Me.pnlExpirationsContainer.Name = "pnlExpirationsContainer"
        Me.pnlExpirationsContainer.Size = New System.Drawing.Size(816, 420)
        Me.pnlExpirationsContainer.TabIndex = 1
        '
        'ugExpirations
        '
        Me.ugExpirations.Cursor = System.Windows.Forms.Cursors.Default
        Appearance3.BorderColor = System.Drawing.SystemColors.ControlLight
        Me.ugExpirations.DisplayLayout.Override.CellAppearance = Appearance3
        Me.ugExpirations.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Appearance4.BorderColor = System.Drawing.SystemColors.ControlLight
        Me.ugExpirations.DisplayLayout.Override.RowAppearance = Appearance4
        Me.ugExpirations.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugExpirations.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugExpirations.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugExpirations.Location = New System.Drawing.Point(0, 0)
        Me.ugExpirations.Name = "ugExpirations"
        Me.ugExpirations.Size = New System.Drawing.Size(816, 420)
        Me.ugExpirations.TabIndex = 0
        '
        'pnlExpirationsBottom
        '
        Me.pnlExpirationsBottom.Controls.Add(Me.btnGenerateExpirationLetters)
        Me.pnlExpirationsBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlExpirationsBottom.Location = New System.Drawing.Point(0, 460)
        Me.pnlExpirationsBottom.Name = "pnlExpirationsBottom"
        Me.pnlExpirationsBottom.Size = New System.Drawing.Size(816, 48)
        Me.pnlExpirationsBottom.TabIndex = 2
        '
        'btnGenerateExpirationLetters
        '
        Me.btnGenerateExpirationLetters.Location = New System.Drawing.Point(8, 8)
        Me.btnGenerateExpirationLetters.Name = "btnGenerateExpirationLetters"
        Me.btnGenerateExpirationLetters.Size = New System.Drawing.Size(160, 32)
        Me.btnGenerateExpirationLetters.TabIndex = 0
        Me.btnGenerateExpirationLetters.Text = "Generate Expiration Letters"
        '
        'pnlExpirationsTop
        '
        Me.pnlExpirationsTop.Controls.Add(Me.btnClearAllforExpiring)
        Me.pnlExpirationsTop.Controls.Add(Me.btnSelectAllforExpiring)
        Me.pnlExpirationsTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlExpirationsTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlExpirationsTop.Name = "pnlExpirationsTop"
        Me.pnlExpirationsTop.Size = New System.Drawing.Size(816, 40)
        Me.pnlExpirationsTop.TabIndex = 0
        '
        'btnClearAllforExpiring
        '
        Me.btnClearAllforExpiring.Location = New System.Drawing.Point(88, 8)
        Me.btnClearAllforExpiring.Name = "btnClearAllforExpiring"
        Me.btnClearAllforExpiring.Size = New System.Drawing.Size(72, 24)
        Me.btnClearAllforExpiring.TabIndex = 3
        Me.btnClearAllforExpiring.Text = "Clear All"
        '
        'btnSelectAllforExpiring
        '
        Me.btnSelectAllforExpiring.Location = New System.Drawing.Point(8, 8)
        Me.btnSelectAllforExpiring.Name = "btnSelectAllforExpiring"
        Me.btnSelectAllforExpiring.Size = New System.Drawing.Size(72, 24)
        Me.btnSelectAllforExpiring.TabIndex = 0
        Me.btnSelectAllforExpiring.Text = "Select All"
        '
        'tbPageRenewals
        '
        Me.tbPageRenewals.Controls.Add(Me.pnlRenewalsContainer)
        Me.tbPageRenewals.Controls.Add(Me.pnlRenewalsBottom)
        Me.tbPageRenewals.Controls.Add(Me.pnlRenewalsTop)
        Me.tbPageRenewals.Location = New System.Drawing.Point(4, 22)
        Me.tbPageRenewals.Name = "tbPageRenewals"
        Me.tbPageRenewals.Size = New System.Drawing.Size(816, 508)
        Me.tbPageRenewals.TabIndex = 0
        Me.tbPageRenewals.Text = "Renewals"
        '
        'pnlRenewalsContainer
        '
        Me.pnlRenewalsContainer.Controls.Add(Me.ugRenewals)
        Me.pnlRenewalsContainer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlRenewalsContainer.Location = New System.Drawing.Point(0, 40)
        Me.pnlRenewalsContainer.Name = "pnlRenewalsContainer"
        Me.pnlRenewalsContainer.Size = New System.Drawing.Size(816, 420)
        Me.pnlRenewalsContainer.TabIndex = 1
        '
        'ugRenewals
        '
        Me.ugRenewals.Cursor = System.Windows.Forms.Cursors.Default
        Appearance5.BorderColor = System.Drawing.SystemColors.ControlLight
        Me.ugRenewals.DisplayLayout.Override.CellAppearance = Appearance5
        Me.ugRenewals.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Appearance6.BorderColor = System.Drawing.SystemColors.ControlLight
        Me.ugRenewals.DisplayLayout.Override.RowAppearance = Appearance6
        Me.ugRenewals.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugRenewals.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugRenewals.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugRenewals.Location = New System.Drawing.Point(0, 0)
        Me.ugRenewals.Name = "ugRenewals"
        Me.ugRenewals.Size = New System.Drawing.Size(816, 420)
        Me.ugRenewals.TabIndex = 0
        '
        'pnlRenewalsBottom
        '
        Me.pnlRenewalsBottom.Controls.Add(Me.btnInfoLetter)
        Me.pnlRenewalsBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlRenewalsBottom.Location = New System.Drawing.Point(0, 460)
        Me.pnlRenewalsBottom.Name = "pnlRenewalsBottom"
        Me.pnlRenewalsBottom.Size = New System.Drawing.Size(816, 48)
        Me.pnlRenewalsBottom.TabIndex = 2
        '
        'btnInfoLetter
        '
        Me.btnInfoLetter.Location = New System.Drawing.Point(8, 8)
        Me.btnInfoLetter.Name = "btnInfoLetter"
        Me.btnInfoLetter.Size = New System.Drawing.Size(168, 32)
        Me.btnInfoLetter.TabIndex = 1
        Me.btnInfoLetter.Text = "Generate Info Needed Letter"
        '
        'pnlRenewalsTop
        '
        Me.pnlRenewalsTop.Controls.Add(Me.btnClearAllforRenewal)
        Me.pnlRenewalsTop.Controls.Add(Me.btnSelectAllforRenewal)
        Me.pnlRenewalsTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlRenewalsTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlRenewalsTop.Name = "pnlRenewalsTop"
        Me.pnlRenewalsTop.Size = New System.Drawing.Size(816, 40)
        Me.pnlRenewalsTop.TabIndex = 0
        '
        'btnClearAllforRenewal
        '
        Me.btnClearAllforRenewal.Location = New System.Drawing.Point(88, 8)
        Me.btnClearAllforRenewal.Name = "btnClearAllforRenewal"
        Me.btnClearAllforRenewal.Size = New System.Drawing.Size(72, 24)
        Me.btnClearAllforRenewal.TabIndex = 1
        Me.btnClearAllforRenewal.Text = "Clear All"
        '
        'btnSelectAllforRenewal
        '
        Me.btnSelectAllforRenewal.Location = New System.Drawing.Point(8, 8)
        Me.btnSelectAllforRenewal.Name = "btnSelectAllforRenewal"
        Me.btnSelectAllforRenewal.Size = New System.Drawing.Size(72, 24)
        Me.btnSelectAllforRenewal.TabIndex = 0
        Me.btnSelectAllforRenewal.Text = "Select All"
        '
        'LicenseeManagement
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(824, 534)
        Me.Controls.Add(Me.tabCntrlLicenseeMgmt)
        Me.Name = "LicenseeManagement"
        Me.Text = "LicenseeManagement"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.tabCntrlLicenseeMgmt.ResumeLayout(False)
        Me.tbPageReminders.ResumeLayout(False)
        Me.pnlRemindersContainer.ResumeLayout(False)
        CType(Me.ugReminders, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlRemindersBottom.ResumeLayout(False)
        Me.pnlRemindersTop.ResumeLayout(False)
        Me.tbPageExpirations.ResumeLayout(False)
        Me.pnlExpirationsContainer.ResumeLayout(False)
        CType(Me.ugExpirations, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlExpirationsBottom.ResumeLayout(False)
        Me.pnlExpirationsTop.ResumeLayout(False)
        Me.tbPageRenewals.ResumeLayout(False)
        Me.pnlRenewalsContainer.ResumeLayout(False)
        CType(Me.ugRenewals, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlRenewalsBottom.ResumeLayout(False)
        Me.pnlRenewalsTop.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "Form Events"
    Private Sub LicenseeManagement_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            ' Process Renewals,Reminders and Expiration Logic
            bolLoading = True
            pLicen = New MUSTER.BusinessLogic.pLicensee
            'If pLicen.ProcessRRE() Then
            'tabCntrlLicenseeMgmt.SelectedTab = tbPageRenewals
            'tabCntrlLicenseeMgmt_Click(sender, e)
            'End If
            If tabCntrlLicenseeMgmt.TabPages.Contains(tbPageRenewals) Then tabCntrlLicenseeMgmt.TabPages.Remove(tbPageRenewals)
            tabCntrlLicenseeMgmt.SelectedTab = tbPageReminders
            UIUtilsGen.SetDatePickerValue(dtPickReminder, dtReminder)
            bolLoading = False
            tabCntrlLicenseeMgmt_SelectedIndexChanged(sender, e)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            bolLoading = False
        End Try
    End Sub
    Private Sub LicenseeManagement_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
    End Sub
    Private Sub LicenseeManagement_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed

    End Sub
#End Region

#Region "UI Support Routines"
    Private Sub SelectorClearAll(ByRef uGrid As Infragistics.Win.UltraWinGrid.UltraGrid, Optional ByVal SelectAll As Boolean = True)
        Try
            bolLoading = True
            For Each ugrow In uGrid.Rows
                If SelectAll Then
                    ugrow.Cells("Selected").Value = True
                Else
                    ugrow.Cells("Selected").Value = False
                End If
            Next
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            bolLoading = False
        End Try
    End Sub
#End Region

#Region "Renewals"
    'Private Sub PopulateRenewals()
    '    Try
    '        Dim dsRenewal As DataSet
    '        dsRenewal = pLicen.GetLicenseesByType("RENEWAL", False)
    '        If dsRenewal.Tables.Count > 0 Then
    '            dsRenewal.Tables(0).DefaultView.Sort = "LAST_NAME"
    '            ugRenewals.DataSource = dsRenewal.Tables(0).DefaultView
    '            ugRenewals.DisplayLayout.Bands(0).Columns("Date_Generated").Hidden = True
    '            ugRenewals.DisplayLayout.Bands(0).Columns("LicenseeID").Hidden = True
    '            ugRenewals.DisplayLayout.Bands(0).Columns("LAST_NAME").Hidden = True
    '            ugRenewals.DisplayLayout.Bands(0).Columns("Licensee").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
    '            ugRenewals.DisplayLayout.Bands(0).Columns("LicenseeNo").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
    '            ugRenewals.DisplayLayout.Bands(0).Columns("CompanyName").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
    '            ugRenewals.DisplayLayout.Bands(0).Columns("Cert_Type").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
    '            ugRenewals.DisplayLayout.Bands(0).Columns("Status").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
    '            ugRenewals.DisplayLayout.Bands(0).Columns("issued_Date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

    '            ugRenewals.DisplayLayout.Bands(0).Columns("App_Recvd_Date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
    '            ugRenewals.DisplayLayout.Bands(0).Columns("Expire_Date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
    '            ugRenewals.DisplayLayout.Bands(0).Columns("Address1").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
    '            ugRenewals.DisplayLayout.Bands(0).Columns("Address2").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
    '            ugRenewals.DisplayLayout.Bands(0).Columns("City").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
    '            ugRenewals.DisplayLayout.Bands(0).Columns("State").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

    '            ugRenewals.DisplayLayout.Bands(0).Columns("Zip").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
    '            ugRenewals.DisplayLayout.Bands(0).Columns("Status1").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
    '            ugRenewals.DisplayLayout.Bands(0).Columns("Phone_Number_one").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
    '            ugRenewals.DisplayLayout.Bands(0).Columns("hire_status").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
    '            ugRenewals.DisplayLayout.Bands(0).Columns("employee_letter").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
    '            ugRenewals.DisplayLayout.Bands(0).Columns("fin_resp_end_date").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
    '            ugRenewals.DisplayLayout.Bands(0).Columns("LAST_NAME").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

    '        End If
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub
    'Private Sub btnInfoLetter_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnInfoLetter.Click
    '    Dim strLicenseeIDs As String = String.Empty
    '    Dim bolStatus As Boolean = False

    '    Try
    '        For Each ugrow In ugRenewals.Rows

    '            If ugrow.Cells("Selected").Value Then
    '                bolStatus = True
    '                If strLicenseeIDs <> String.Empty Then
    '                    strLicenseeIDs += "," + ugrow.Cells("LicenseeID").Value.ToString
    '                Else
    '                    strLicenseeIDs += ugrow.Cells("LicenseeID").Value.ToString
    '                End If
    '                UIUtilsGen.Delay(, 1)
    '                'Delay()
    '                oLetter.GenerateLicenseeLetter(ugrow.Cells("LicenseeID").Value, "Info Needed Letter", "InfoNeeded_Letter", "Info Needed Letter for Company", "InfoNeededLetter.doc", , , ugrow, MusterContainer.AppUser.Name)
    '                'Delay()
    '                UIUtilsGen.Delay(, 1)
    '            End If
    '        Next
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub btnSelectAllforRenewal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelectAllforRenewal.Click
    '    Try
    '        SelectorClearAll(ugRenewals, True)
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub btnClearAllforRenewal_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearAllforRenewal.Click
    '    Try
    '        SelectorClearAll(ugRenewals, False)
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
#End Region

#Region "Reminders"
    Private Sub PopulateReminders()
        Try
            Dim dsReminder As DataSet
            If rBtnLetterGenYes.Checked Then
                dsReminder = pLicen.GetLicenseesByType(dtReminder, "REMINDER", False, 1)
            ElseIf rBtnLetterGenNo.Checked Then
                dsReminder = pLicen.GetLicenseesByType(dtReminder, "REMINDER", False, 0)
            Else
                dsReminder = pLicen.GetLicenseesByType(dtReminder, "REMINDER", False)
            End If
            If dsReminder.Tables.Count > 0 Then
                dsReminder.Tables(0).DefaultView.Sort = "LAST_NAME"
                ugReminders.DataSource = dsReminder.Tables(0).DefaultView
                ugReminders.DisplayLayout.Bands(0).Columns("LicenseeID").Hidden = True
                ugReminders.DisplayLayout.Bands(0).Columns("employee_letter").Hidden = True
                ugReminders.DisplayLayout.Bands(0).Columns("fin_resp_end_date").Hidden = True
                ugReminders.DisplayLayout.Bands(0).Columns("LAST_NAME").Hidden = True
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub dtPickReminder_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dtPickReminder.ValueChanged
        If bolLoading Then Exit Sub
        UIUtilsGen.ToggleDateFormat(dtPickReminder)
        UIUtilsGen.FillDateobjectValues(dtReminder, dtPickReminder.Text)
    End Sub

    Private Sub btnSelectAllforReminding_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelectAllforReminding.Click
        Try
            SelectorClearAll(ugReminders, True)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnClearAllforReminding_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearAllforReminding.Click
        Try
            SelectorClearAll(ugReminders, False)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnGenerateInfoNeededLetters_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGenerateInfoNeededLetters.Click
        Try
            Dim strLicenseeIDs As String = String.Empty
            Dim bolStatus As Boolean = False

            For Each ugrow In ugReminders.Rows
                If ugrow.Cells("Selected").Value Then
                    bolStatus = True
                    strLicenseeIDs += ugrow.Cells("LicenseeID").Value.ToString + ","
                    ' generate info needed letters
                    UIUtilsGen.Delay(, 1)
                    oLetter.GenerateLicenseeLetter(ugrow.Cells("LicenseeID").Value, "Info Needed Letter", "InfoNeeded_Letter", "Licensee Info Needed Letter", "InfoNeededLetter.doc", , , ugrow, MusterContainer.AppUser.Name)
                    UIUtilsGen.Delay(, 1)
                End If
            Next

            If bolStatus Then
                strLicenseeIDs = strLicenseeIDs.TrimEnd(",")
                pLicen.UpdateRenewals(strLicenseeIDs, "REMINDER")
                MsgBox("Info Needed Letter(s) generated Successfully")
                PopulateReminders()
            Else
                MsgBox("No Records Found OR Select at least one Licensee")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "Expirations"
    Private Sub PopulateExpirations()
        Dim dsSet As DataSet
        Dim drRow As DataRow
        Try
            dsSet = pLicen.GetLicenseesByType(dtNull, "EXPIRATION", False)
            If dsSet.Tables.Count > 0 Then
                If dsSet.Tables(0).Rows.Count > 0 Then
                    dsSet.Tables(0).DefaultView.RowFilter = "date_generated is null or (date_generated < expire_date) "
                End If
                dsSet.Tables(0).DefaultView.Sort = "LAST_NAME"
            End If

            ugExpirations.DataSource = dsSet.Tables(0).DefaultView
            ugExpirations.DisplayLayout.Bands(0).Columns("Date_Generated").Hidden = True
            ugExpirations.DisplayLayout.Bands(0).Columns("LicenseeID").Hidden = True
            ugExpirations.DisplayLayout.Bands(0).Columns("employee_letter").Hidden = True
            ugExpirations.DisplayLayout.Bands(0).Columns("fin_resp_end_date").Hidden = True
            ugExpirations.DisplayLayout.Bands(0).Columns("LAST_NAME").Hidden = True
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub btnSelectAllforExpiring_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelectAllforExpiring.Click
        Try
            SelectorClearAll(ugExpirations, True)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnClearAllforExpiring_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClearAllforExpiring.Click
        Try
            SelectorClearAll(ugExpirations, False)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    'Private Sub btnProcessCertifiedExpirations_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim dsSet As DataSet
    '    Dim pLicen As New MUSTER.BusinessLogic.pLicensee
    '    Dim pCompany As New MUSTER.BusinessLogic.pCompany
    '    'Dim xLicenseeinfo As New MUSTER.Info.LicenseeInfo
    '    Dim bolsuccess As Boolean = False
    '    Dim drRow As DataRow
    '    Dim dtFinRespEndDate As Date
    '    Dim bolCompanySuccess As Boolean = False
    '    Dim count As Integer = 0
    '    Dim bolstatus As Boolean = False
    '    Dim success As Boolean = False
    '    Try
    '        dsSet = pLicen.GetLicenseesToBeProcessed()
    '        If dsSet.Tables.Count > 0 Then
    '            If dsSet.Tables(0).Rows.Count > 0 Then
    '                For Each drRow In dsSet.Tables(0).Rows
    '                    bolstatus = True
    '                    pLicen.Retrieve(Integer.Parse(drRow.Item("Licensee_ID")))
    '                    'xLicenseeinfo = pLicen.LicenseeInfo
    '                    If drRow("FIN_RESP_END_DATE") Is System.DBNull.Value Then
    '                        dtFinRespEndDate = CDate("01/01/0001")
    '                    Else
    '                        dtFinRespEndDate = CDate(drRow("FIN_RESP_END_DATE"))
    '                    End If
    '                    'If drRow.Item("COMPANY_NAME") <> String.Empty Then
    '                    '    pLicen.LicenseeLogic(xLicenseeinfo, dtFinRespEndDate, True, Integer.Parse(drRow("Licensee_ID")), drRow("Hire_status"))
    '                    'Else
    '                    '    pLicen.LicenseeLogic(xLicenseeinfo, dtFinRespEndDate, , Integer.Parse(drRow("Licensee_ID")), drRow("Hire_status"))
    '                    'End If

    '                    'have to allow for the - intergers from migration and the - intergers from collections
    '                    If pLicen.ID >= -100 And pLicen.ID <= 0 Then
    '                        pLicen.CreatedBy = MusterContainer.AppUser.ID
    '                    Else
    '                        pLicen.ModifiedBy = MusterContainer.AppUser.ID
    '                    End If
    '                    success = False
    '                    success = pLicen.Save(CType(UIUtilsGen.ModuleID.Company, Integer), MusterContainer.AppUser.UserKey, returnVal)
    '                    If Not UIUtilsGen.HasRights(returnVal) Then
    '                        Exit Sub
    '                    End If
    '                    If Not success Then
    '                        bolsuccess = False
    '                        Exit Sub
    '                    End If
    '                    pCompany.Retrieve(Integer.Parse(drRow("COMPANY_ID")))
    '                    pCompany.CTIAC = False
    '                    pCompany.CTC = False
    '                    bolCompanySuccess = pCompany.Save(UIUtilsGen.ModuleID.Company, MusterContainer.AppUser.UserKey, returnVal)
    '                    If Not UIUtilsGen.HasRights(returnVal) Then
    '                        Exit Sub
    '                    End If
    '                    If success And bolCompanySuccess Then
    '                        count = count + 1
    '                    End If
    '                Next
    '            End If
    '        End If
    '        If bolstatus Then
    '            MsgBox("Number of Licensees examined are " + dsSet.Tables(0).Rows.Count.ToString + " . " & vbCrLf & "[" + count.ToString + "] changed to NOT CURRENTLY CERTIFIED")
    '            PopulateExpirations()
    '        Else
    '            MsgBox("Number of Licensses examined are 0")
    '        End If

    '        'For Each ugrow In ugRenewals.Rows
    '        '    If ugrow.Cells("Selected").Value Then

    '        '        bolStatus = True
    '        '        If strLicenseeIDs <> String.Empty Then
    '        '            strLicenseeIDs += "," + ugrow.Cells("LicenseeID").Value.ToString
    '        '        Else
    '        '            strLicenseeIDs += ugrow.Cells("LicenseeID").Value.ToString
    '        '        End If
    '        '        pLicen.Retrieve(Integer.Parse(ugrow.Cells("LicenseeID").Text))
    '        '        xLicenseeinfo = pLicen.LicenseeInfo
    '        '        If ugrow.Cells("CompanyName").Text <> String.Empty Then
    '        '            pLicen.LicenseeLogic(xLicenseeinfo, CDate(ugrow.Cells("FIN_RESP_END_DATE").Text), True, Integer.Parse(ugrow.Cells("LicenseeID").Text), ugrow.Cells("Hire_status").Text)
    '        '        Else
    '        '            pLicen.LicenseeLogic(xLicenseeinfo, CDate(ugrow.Cells("FIN_RESP_END_DATE").Text), , Integer.Parse(ugrow.Cells("LicenseeID").Text), ugrow.Cells("Hire_status").Text)
    '        '        End If
    '        '        Dim success As Boolean = False
    '        '        'have to allow for the - intergers fro migration and the - intergers from collections
    '        '        If pLicen.ID >= -100 And pLicen.ID <= 0 Then
    '        '            pLicen.CreatedBy = MusterContainer.AppUser.ID
    '        '        Else
    '        '            pLicen.ModifiedBy = MusterContainer.AppUser.ID
    '        '        End If
    '        '        success = pLicen.Save(CType(UIUtilsGen.ModuleID.Company, Integer), MusterContainer.AppUser.UserKey, returnVal)
    '        '        If Not UIUtilsGen.HasRights(returnVal) Then
    '        '            Exit Sub
    '        '        End If
    '        '        If Not success Then
    '        '            bolsuccess = False
    '        '            Exit Sub
    '        '        End If
    '        '        UIUtilsGen.Delay(, 1)
    '        '        'Delay()
    '        '        oLetter.GenerateLicenseeLetter(ugrow.Cells("LicenseeID").Value, "Renewal Certificate Letter", "Renewal_Certificate", "Licensee Renewal Certificate for Company", "LicenseeRenewalLetter.doc", , , ugrow, MusterContainer.AppUser.ID)
    '        '        'Delay()
    '        '        UIUtilsGen.Delay(, 1)
    '        '    End If
    '        'Next

    '        'If bolStatus Then
    '        '    pLicen.UpdateRenewals(strLicenseeIDs, "RENEWAL")
    '        '    MsgBox("Licensee Renewal Certificate Generated Successfull.")
    '        '    PopulateRenewals()
    '        'Else
    '        '    MsgBox("Select at least one Licensee")
    '        'End If
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try

    'End Sub
    Private Sub btnGenerateExpirationLetters_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGenerateExpirationLetters.Click
        Try
            Dim strLicenseeIDs As String = String.Empty
            Dim bolStatus As Boolean = False

            For Each ugrow In ugExpirations.Rows
                If ugrow.Cells("Selected").Value Then
                    bolStatus = True
                    strLicenseeIDs += ugrow.Cells("LicenseeID").Value.ToString + ","
                    ' generate expiration letters
                    UIUtilsGen.Delay(, 1)
                    oLetter.GenerateLicenseeLetter(ugrow.Cells("LicenseeID").Value, "Licensee Expired", "Expired", "Licensee Expired Letter", "ExpirationLetter.doc", , , ugrow, MusterContainer.AppUser.Name)
                    UIUtilsGen.Delay(, 1)
                End If
            Next

            If bolStatus Then
                strLicenseeIDs = strLicenseeIDs.TrimEnd(",")
                pLicen.UpdateRenewals(strLicenseeIDs, "EXPIRATION")
                MsgBox("No Longer Certified Letter Generated Successfully.")
                PopulateExpirations()
            Else
                MsgBox("No Records Found OR Select at least one Licensee")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "UI Events"
    'Private Sub tabCntrlLicenseeMgmt_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabCntrlLicenseeMgmt.Click
    '    Try
    '        Select Case tabCntrlLicenseeMgmt.SelectedTab.Name.ToUpper
    '            Case "TBPAGERENEWALS"
    '                PopulateRenewals()
    '                ''ugRenewals.DataSource = pLicen.GetLicenseesByType("RENEWAL", False)
    '                ''ugRenewals.DisplayLayout.Bands(0).Columns("Date_Generated").Hidden = True
    '            Case "TBPAGEREMINDERS"
    '                PopulateReminders()
    '                'ugReminders.DataSource = pLicen.GetLicenseesByType("REMINDER", False)
    '            Case "TBPAGEEXPIRATIONS"
    '                PopulateExpirations()
    '                'ugExpirations.DataSource = pLicen.GetLicenseesByType("EXPIRATION", False)
    '                'ugRenewals.DisplayLayout.Bands(0).Columns("Date_Generated").Hidden = True
    '        End Select
    '        'ugRenewals.DisplayLayout.Bands(0).Columns("LicenseeID").Hidden = True
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    Private Sub tabCntrlLicenseeMgmt_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles tabCntrlLicenseeMgmt.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            Select Case tabCntrlLicenseeMgmt.SelectedTab.Name
                'Case tbPageRenewals.Name
                '    PopulateRenewals()
                '    'ugRenewals.DataSource = pLicen.GetLicenseesByType("RENEWAL", False)
                '    'ugRenewals.DisplayLayout.Bands(0).Columns("LicenseeID").Hidden = True
                '    'ugRenewals.DisplayLayout.Bands(0).Columns("Date_Generated").Hidden = True
            Case tbPageReminders.Name
                    PopulateReminders()
                    'ugReminders.DataSource = pLicen.GetLicenseesByType("REMINDER", False)
                    'ugRenewals.DisplayLayout.Bands(0).Columns("LicenseeID").Hidden = True
                Case tbPageExpirations.Name
                    'ugExpirations.DataSource = pLicen.GetLicenseesByType("EXPIRATION", False)
                    'ugRenewals.DisplayLayout.Bands(0).Columns("LicenseeID").Hidden = True
                    'ugRenewals.DisplayLayout.Bands(0).Columns("Date_Generated").Hidden = True
                    PopulateExpirations()
            End Select
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.Click
        PopulateReminders()
    End Sub
#End Region

    Private Sub btnInfoLetter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInfoLetter.Click

    End Sub
End Class
