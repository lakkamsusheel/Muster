Public Class UserAdmin
    Inherits System.Windows.Forms.Form
    '-------------------------------------------------------------------------------
    ' MUSTER.MUSTER.UserAdmin
    '   User Administration Screen
    '
    ' Copyright (C) 2004 CIBER, Inc.
    ' All rights reserved.
    '
    ' Release   Initials    Date      Description
    '  1.0        ??      ??/??/??    Original class definition.
    '  1.1        PN      12/28/04    Modified btnCancel event.
    '                                 User_Load,SaveForms,btnNew_Click,user_closing,
    '                                       btnSave,ComboUSERNAME_ValueChanged.
    '                                 New sub UserIsDirty
    '  1.2        JC      12/31/04    Changed local variable pUser to oUser and
    '                                       pGroups to oGroups
    '  1.3        JC      01/04/05    Added code to handle MDI Child integration.
    '  1.4        JC      01/12/05    Added code to handle load behavior.   Also added
    '                                   code to ComboUserName.FindExact on leave. Also
    '                                   eliminated LoadUserData and incorporated it
    '                                   into ComboUserName_Leave.  Finally, modified
    '                                   ComboUserName_Leave to accomodate returning
    '                                   to previously selected user if new condition
    '                                   is aborted by user.
    '  1.5        PN      01/12/05    Added code to handle proper clearing of lists 
    '                                   and other identified issues.
    '  1.6        PN      01/20/05    Modified Cancel event(bug 652)  
    '  1.7        PN      01/24/05    Set Controls tab order(bug 660) 
    '  1.8        PN      01/31/05    Added method  IsEmptyUserName(),events txtName_Enter
    '                                 txtEmail_Enter,txtPhone_Enter,chkInactive_Enter
    '  1.9        AN      02/10/05    Integrated AppFlags new object model
    '  2.0        PN      02/15/05    Set sorted property of listBoxs to true(bug 700)
    '  2.1        PN      02/15/05    Changed clearUserData(bug 698) 
    '  2.2        MR      02/24/05    Modified DisplayUserData() to Validate for ComboUserName.Text instead of ComboUserName.SelectedValue.
    '  2.3        MR      03/03/05    Modified ComboUserName_Leave() to Compare OldUserName with SelectedUserName.
    '-------------------------------------------------------------------------------
    '
    'TODO - Remove comment from VSS version 2/9/05 - JVC 2
    '
#Region "Private Member Variables"
    Private WithEvents oUser As MUSTER.BusinessLogic.pUser
    Dim strPreviousUserName As String = String.Empty

    Dim strUserName As String = String.Empty
    Private bolIsNewUser As Boolean = False
    Private bolLoading As Boolean = False
    Dim bolErrorOccurred As Boolean = False

    Dim userGroupRelInfo As MUSTER.Info.UserGroupRelationInfo
    Dim userInfo As MUSTER.Info.UserInfo
    Dim dtAvailGroups, dtAvailUsers, dtAssignedGroups, dtManagedUsers As DataTable
    Dim dr As DataRow
    Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow

    'Private oGroups As Muster.BusinessLogic.pUserGroupMemberships
    Private WithEvents oChangePassword As New ChangePassword
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
        InitDatatable()
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
    Friend WithEvents lblPhone As System.Windows.Forms.Label
    Friend WithEvents lblEmail As System.Windows.Forms.Label
    Friend WithEvents lblUserName As System.Windows.Forms.Label
    Friend WithEvents lblName As System.Windows.Forms.Label
    Friend WithEvents txtPhone As System.Windows.Forms.TextBox
    Friend WithEvents txtEmail As System.Windows.Forms.TextBox
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents cboPrimaryModule As System.Windows.Forms.ComboBox
    Friend WithEvents lblPrimaryModule As System.Windows.Forms.Label
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents btnNew As System.Windows.Forms.Button
    Friend WithEvents btnResetPassword As System.Windows.Forms.Button
    Friend WithEvents lblAvailableGroups As System.Windows.Forms.Label
    Friend WithEvents lblAssignedGroups As System.Windows.Forms.Label
    Friend WithEvents btnGroupShiftRight As System.Windows.Forms.Button
    Friend WithEvents btnGroupShiftRightAll As System.Windows.Forms.Button
    Friend WithEvents btnGroupShiftLeftAll As System.Windows.Forms.Button
    Friend WithEvents btnGroupShiftLeft As System.Windows.Forms.Button
    Friend WithEvents btnUserShiftLeftAll As System.Windows.Forms.Button
    Friend WithEvents btnUserShiftLeft As System.Windows.Forms.Button
    Friend WithEvents btnUserShiftRightAll As System.Windows.Forms.Button
    Friend WithEvents btnUserShiftRight As System.Windows.Forms.Button
    Friend WithEvents lblAvailableUsers As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents lblSupervisedUsers As System.Windows.Forms.Label
    Friend WithEvents chkInactive As System.Windows.Forms.CheckBox
    Friend WithEvents ComboUserName As System.Windows.Forms.ComboBox
    Friend WithEvents lblpsw As System.Windows.Forms.Label
    'Public WithEvents lblpsw As System.Windows.Forms.Label
    'Public Shared WithEvents lblpsw As System.Windows.Forms.Label
    Friend WithEvents txtUserName As System.Windows.Forms.TextBox
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ChkHEAD_CLOSURE As System.Windows.Forms.CheckBox
    Friend WithEvents ChkHEAD_FINANCIAL As System.Windows.Forms.CheckBox
    Friend WithEvents ChkHEAD_FEES As System.Windows.Forms.CheckBox
    Friend WithEvents ChkHEAD_CANDE As System.Windows.Forms.CheckBox
    Friend WithEvents ChkHEAD_INSPECTION As System.Windows.Forms.CheckBox
    Friend WithEvents ChkHEAD_REGISTRATION As System.Windows.Forms.CheckBox
    Friend WithEvents ChkHEAD_PM As System.Windows.Forms.CheckBox
    Friend WithEvents ChkHEAD_ADMIN As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ugAssignedGroups As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents ugAvailGroups As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents ugAvailUsers As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents ugManagedUsers As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents ChkEXECUTIVE_DIRECTOR As System.Windows.Forms.CheckBox

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblPhone = New System.Windows.Forms.Label
        Me.lblEmail = New System.Windows.Forms.Label
        Me.lblUserName = New System.Windows.Forms.Label
        Me.lblName = New System.Windows.Forms.Label
        Me.txtPhone = New System.Windows.Forms.TextBox
        Me.txtEmail = New System.Windows.Forms.TextBox
        Me.txtName = New System.Windows.Forms.TextBox
        Me.cboPrimaryModule = New System.Windows.Forms.ComboBox
        Me.lblPrimaryModule = New System.Windows.Forms.Label
        Me.chkInactive = New System.Windows.Forms.CheckBox
        Me.btnSearch = New System.Windows.Forms.Button
        Me.btnResetPassword = New System.Windows.Forms.Button
        Me.lblAvailableGroups = New System.Windows.Forms.Label
        Me.lblAssignedGroups = New System.Windows.Forms.Label
        Me.btnGroupShiftRight = New System.Windows.Forms.Button
        Me.btnGroupShiftRightAll = New System.Windows.Forms.Button
        Me.btnGroupShiftLeftAll = New System.Windows.Forms.Button
        Me.btnGroupShiftLeft = New System.Windows.Forms.Button
        Me.btnUserShiftLeftAll = New System.Windows.Forms.Button
        Me.btnUserShiftLeft = New System.Windows.Forms.Button
        Me.btnUserShiftRightAll = New System.Windows.Forms.Button
        Me.btnUserShiftRight = New System.Windows.Forms.Button
        Me.lblSupervisedUsers = New System.Windows.Forms.Label
        Me.lblAvailableUsers = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnNew = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.ComboUserName = New System.Windows.Forms.ComboBox
        Me.lblpsw = New System.Windows.Forms.Label
        Me.txtUserName = New System.Windows.Forms.TextBox
        Me.btnDelete = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.ChkHEAD_CLOSURE = New System.Windows.Forms.CheckBox
        Me.ChkHEAD_FINANCIAL = New System.Windows.Forms.CheckBox
        Me.ChkHEAD_FEES = New System.Windows.Forms.CheckBox
        Me.ChkHEAD_CANDE = New System.Windows.Forms.CheckBox
        Me.ChkHEAD_INSPECTION = New System.Windows.Forms.CheckBox
        Me.ChkHEAD_REGISTRATION = New System.Windows.Forms.CheckBox
        Me.ChkHEAD_PM = New System.Windows.Forms.CheckBox
        Me.ChkHEAD_ADMIN = New System.Windows.Forms.CheckBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.ugAssignedGroups = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.ugAvailGroups = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.ugAvailUsers = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.ugManagedUsers = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.ChkEXECUTIVE_DIRECTOR = New System.Windows.Forms.CheckBox
        Me.GroupBox1.SuspendLayout()
        CType(Me.ugAssignedGroups, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ugAvailGroups, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ugAvailUsers, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ugManagedUsers, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblPhone
        '
        Me.lblPhone.Location = New System.Drawing.Point(24, 112)
        Me.lblPhone.Name = "lblPhone"
        Me.lblPhone.Size = New System.Drawing.Size(56, 16)
        Me.lblPhone.TabIndex = 16
        Me.lblPhone.Text = "Phone"
        Me.lblPhone.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblEmail
        '
        Me.lblEmail.Location = New System.Drawing.Point(24, 88)
        Me.lblEmail.Name = "lblEmail"
        Me.lblEmail.Size = New System.Drawing.Size(56, 16)
        Me.lblEmail.TabIndex = 15
        Me.lblEmail.Text = "Email"
        Me.lblEmail.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblUserName
        '
        Me.lblUserName.Location = New System.Drawing.Point(0, 37)
        Me.lblUserName.Name = "lblUserName"
        Me.lblUserName.Size = New System.Drawing.Size(88, 16)
        Me.lblUserName.TabIndex = 14
        Me.lblUserName.Text = "New User Name"
        Me.lblUserName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblName
        '
        Me.lblName.Location = New System.Drawing.Point(24, 64)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(56, 16)
        Me.lblName.TabIndex = 13
        Me.lblName.Text = "Name"
        Me.lblName.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtPhone
        '
        Me.txtPhone.Location = New System.Drawing.Point(96, 104)
        Me.txtPhone.Name = "txtPhone"
        Me.txtPhone.Size = New System.Drawing.Size(144, 20)
        Me.txtPhone.TabIndex = 50
        Me.txtPhone.Text = ""
        '
        'txtEmail
        '
        Me.txtEmail.Location = New System.Drawing.Point(96, 80)
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.Size = New System.Drawing.Size(144, 20)
        Me.txtEmail.TabIndex = 40
        Me.txtEmail.Text = ""
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(96, 56)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(144, 20)
        Me.txtName.TabIndex = 30
        Me.txtName.Text = ""
        '
        'cboPrimaryModule
        '
        Me.cboPrimaryModule.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPrimaryModule.DropDownWidth = 180
        Me.cboPrimaryModule.ItemHeight = 13
        Me.cboPrimaryModule.Location = New System.Drawing.Point(136, 144)
        Me.cboPrimaryModule.Name = "cboPrimaryModule"
        Me.cboPrimaryModule.Size = New System.Drawing.Size(144, 21)
        Me.cboPrimaryModule.TabIndex = 70
        '
        'lblPrimaryModule
        '
        Me.lblPrimaryModule.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPrimaryModule.Location = New System.Drawing.Point(24, 144)
        Me.lblPrimaryModule.Name = "lblPrimaryModule"
        Me.lblPrimaryModule.Size = New System.Drawing.Size(96, 23)
        Me.lblPrimaryModule.TabIndex = 108
        Me.lblPrimaryModule.Text = "Primary Module"
        '
        'chkInactive
        '
        Me.chkInactive.Location = New System.Drawing.Point(288, 8)
        Me.chkInactive.Name = "chkInactive"
        Me.chkInactive.Size = New System.Drawing.Size(72, 24)
        Me.chkInactive.TabIndex = 12
        Me.chkInactive.Text = "Inactive"
        '
        'btnSearch
        '
        Me.btnSearch.BackColor = System.Drawing.SystemColors.Control
        Me.btnSearch.Enabled = False
        Me.btnSearch.Location = New System.Drawing.Point(248, 8)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(24, 24)
        Me.btnSearch.TabIndex = 11
        Me.btnSearch.Text = "?"
        '
        'btnResetPassword
        '
        Me.btnResetPassword.BackColor = System.Drawing.SystemColors.Control
        Me.btnResetPassword.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnResetPassword.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnResetPassword.Location = New System.Drawing.Point(128, 176)
        Me.btnResetPassword.Name = "btnResetPassword"
        Me.btnResetPassword.Size = New System.Drawing.Size(112, 23)
        Me.btnResetPassword.TabIndex = 80
        Me.btnResetPassword.Text = "Reset Password"
        '
        'lblAvailableGroups
        '
        Me.lblAvailableGroups.Location = New System.Drawing.Point(32, 208)
        Me.lblAvailableGroups.Name = "lblAvailableGroups"
        Me.lblAvailableGroups.Size = New System.Drawing.Size(104, 16)
        Me.lblAvailableGroups.TabIndex = 113
        Me.lblAvailableGroups.Text = "Available Groups"
        '
        'lblAssignedGroups
        '
        Me.lblAssignedGroups.Location = New System.Drawing.Point(288, 208)
        Me.lblAssignedGroups.Name = "lblAssignedGroups"
        Me.lblAssignedGroups.Size = New System.Drawing.Size(104, 16)
        Me.lblAssignedGroups.TabIndex = 115
        Me.lblAssignedGroups.Text = "Assigned Groups"
        '
        'btnGroupShiftRight
        '
        Me.btnGroupShiftRight.BackColor = System.Drawing.SystemColors.Control
        Me.btnGroupShiftRight.Enabled = False
        Me.btnGroupShiftRight.Location = New System.Drawing.Point(236, 232)
        Me.btnGroupShiftRight.Name = "btnGroupShiftRight"
        Me.btnGroupShiftRight.Size = New System.Drawing.Size(32, 24)
        Me.btnGroupShiftRight.TabIndex = 91
        Me.btnGroupShiftRight.Text = ">"
        '
        'btnGroupShiftRightAll
        '
        Me.btnGroupShiftRightAll.BackColor = System.Drawing.SystemColors.Control
        Me.btnGroupShiftRightAll.Enabled = False
        Me.btnGroupShiftRightAll.Location = New System.Drawing.Point(236, 256)
        Me.btnGroupShiftRightAll.Name = "btnGroupShiftRightAll"
        Me.btnGroupShiftRightAll.Size = New System.Drawing.Size(32, 24)
        Me.btnGroupShiftRightAll.TabIndex = 92
        Me.btnGroupShiftRightAll.Text = ">>"
        '
        'btnGroupShiftLeftAll
        '
        Me.btnGroupShiftLeftAll.BackColor = System.Drawing.SystemColors.Control
        Me.btnGroupShiftLeftAll.Enabled = False
        Me.btnGroupShiftLeftAll.Location = New System.Drawing.Point(236, 312)
        Me.btnGroupShiftLeftAll.Name = "btnGroupShiftLeftAll"
        Me.btnGroupShiftLeftAll.Size = New System.Drawing.Size(32, 24)
        Me.btnGroupShiftLeftAll.TabIndex = 94
        Me.btnGroupShiftLeftAll.Text = "<<"
        '
        'btnGroupShiftLeft
        '
        Me.btnGroupShiftLeft.BackColor = System.Drawing.SystemColors.Control
        Me.btnGroupShiftLeft.Enabled = False
        Me.btnGroupShiftLeft.Location = New System.Drawing.Point(236, 288)
        Me.btnGroupShiftLeft.Name = "btnGroupShiftLeft"
        Me.btnGroupShiftLeft.Size = New System.Drawing.Size(32, 24)
        Me.btnGroupShiftLeft.TabIndex = 93
        Me.btnGroupShiftLeft.Text = "<"
        '
        'btnUserShiftLeftAll
        '
        Me.btnUserShiftLeftAll.BackColor = System.Drawing.SystemColors.Control
        Me.btnUserShiftLeftAll.Enabled = False
        Me.btnUserShiftLeftAll.Location = New System.Drawing.Point(236, 464)
        Me.btnUserShiftLeftAll.Name = "btnUserShiftLeftAll"
        Me.btnUserShiftLeftAll.Size = New System.Drawing.Size(32, 24)
        Me.btnUserShiftLeftAll.TabIndex = 104
        Me.btnUserShiftLeftAll.Text = "<<"
        '
        'btnUserShiftLeft
        '
        Me.btnUserShiftLeft.BackColor = System.Drawing.SystemColors.Control
        Me.btnUserShiftLeft.Enabled = False
        Me.btnUserShiftLeft.Location = New System.Drawing.Point(236, 440)
        Me.btnUserShiftLeft.Name = "btnUserShiftLeft"
        Me.btnUserShiftLeft.Size = New System.Drawing.Size(32, 24)
        Me.btnUserShiftLeft.TabIndex = 103
        Me.btnUserShiftLeft.Text = "<"
        '
        'btnUserShiftRightAll
        '
        Me.btnUserShiftRightAll.BackColor = System.Drawing.SystemColors.Control
        Me.btnUserShiftRightAll.Enabled = False
        Me.btnUserShiftRightAll.Location = New System.Drawing.Point(236, 408)
        Me.btnUserShiftRightAll.Name = "btnUserShiftRightAll"
        Me.btnUserShiftRightAll.Size = New System.Drawing.Size(32, 24)
        Me.btnUserShiftRightAll.TabIndex = 102
        Me.btnUserShiftRightAll.Text = ">>"
        '
        'btnUserShiftRight
        '
        Me.btnUserShiftRight.BackColor = System.Drawing.SystemColors.Control
        Me.btnUserShiftRight.Enabled = False
        Me.btnUserShiftRight.Location = New System.Drawing.Point(236, 384)
        Me.btnUserShiftRight.Name = "btnUserShiftRight"
        Me.btnUserShiftRight.Size = New System.Drawing.Size(32, 24)
        Me.btnUserShiftRight.TabIndex = 101
        Me.btnUserShiftRight.Text = ">"
        '
        'lblSupervisedUsers
        '
        Me.lblSupervisedUsers.Location = New System.Drawing.Point(288, 360)
        Me.lblSupervisedUsers.Name = "lblSupervisedUsers"
        Me.lblSupervisedUsers.Size = New System.Drawing.Size(104, 16)
        Me.lblSupervisedUsers.TabIndex = 125
        Me.lblSupervisedUsers.Text = "Supervised Users"
        '
        'lblAvailableUsers
        '
        Me.lblAvailableUsers.Location = New System.Drawing.Point(32, 360)
        Me.lblAvailableUsers.Name = "lblAvailableUsers"
        Me.lblAvailableUsers.Size = New System.Drawing.Size(104, 16)
        Me.lblAvailableUsers.TabIndex = 123
        Me.lblAvailableUsers.Text = "Available Users"
        '
        'btnCancel
        '
        Me.btnCancel.Enabled = False
        Me.btnCancel.Location = New System.Drawing.Point(306, 512)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(80, 26)
        Me.btnCancel.TabIndex = 203
        Me.btnCancel.Text = "Cancel"
        '
        'btnSave
        '
        Me.btnSave.Enabled = False
        Me.btnSave.Location = New System.Drawing.Point(118, 512)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(80, 26)
        Me.btnSave.TabIndex = 201
        Me.btnSave.Text = "Save"
        '
        'btnNew
        '
        Me.btnNew.Location = New System.Drawing.Point(24, 512)
        Me.btnNew.Name = "btnNew"
        Me.btnNew.Size = New System.Drawing.Size(80, 26)
        Me.btnNew.TabIndex = 200
        Me.btnNew.Text = "New"
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(400, 512)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(80, 26)
        Me.btnClose.TabIndex = 204
        Me.btnClose.Text = "Close"
        '
        'ComboUserName
        '
        Me.ComboUserName.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboUserName.Location = New System.Drawing.Point(96, 8)
        Me.ComboUserName.Name = "ComboUserName"
        Me.ComboUserName.Size = New System.Drawing.Size(144, 21)
        Me.ComboUserName.TabIndex = 10
        '
        'lblpsw
        '
        Me.lblpsw.Location = New System.Drawing.Point(0, 0)
        Me.lblpsw.Name = "lblpsw"
        Me.lblpsw.TabIndex = 0
        '
        'txtUserName
        '
        Me.txtUserName.Location = New System.Drawing.Point(96, 32)
        Me.txtUserName.Name = "txtUserName"
        Me.txtUserName.Size = New System.Drawing.Size(144, 20)
        Me.txtUserName.TabIndex = 20
        Me.txtUserName.Text = ""
        '
        'btnDelete
        '
        Me.btnDelete.Enabled = False
        Me.btnDelete.Location = New System.Drawing.Point(212, 512)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 26)
        Me.btnDelete.TabIndex = 202
        Me.btnDelete.Text = "Delete"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.ChkHEAD_CLOSURE)
        Me.GroupBox1.Controls.Add(Me.ChkHEAD_FINANCIAL)
        Me.GroupBox1.Controls.Add(Me.ChkHEAD_FEES)
        Me.GroupBox1.Controls.Add(Me.ChkHEAD_CANDE)
        Me.GroupBox1.Controls.Add(Me.ChkHEAD_INSPECTION)
        Me.GroupBox1.Controls.Add(Me.ChkHEAD_REGISTRATION)
        Me.GroupBox1.Controls.Add(Me.ChkHEAD_PM)
        Me.GroupBox1.Controls.Add(Me.ChkHEAD_ADMIN)
        Me.GroupBox1.Controls.Add(Me.ChkEXECUTIVE_DIRECTOR)
        Me.GroupBox1.Location = New System.Drawing.Point(288, 32)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(152, 168)
        Me.GroupBox1.TabIndex = 60
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Module Heads"
        '
        'ChkHEAD_CLOSURE
        '
        Me.ChkHEAD_CLOSURE.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChkHEAD_CLOSURE.Location = New System.Drawing.Point(8, 32)
        Me.ChkHEAD_CLOSURE.Name = "ChkHEAD_CLOSURE"
        Me.ChkHEAD_CLOSURE.Size = New System.Drawing.Size(125, 15)
        Me.ChkHEAD_CLOSURE.TabIndex = 62
        Me.ChkHEAD_CLOSURE.Text = "CLOSURE"
        '
        'ChkHEAD_FINANCIAL
        '
        Me.ChkHEAD_FINANCIAL.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChkHEAD_FINANCIAL.Location = New System.Drawing.Point(8, 112)
        Me.ChkHEAD_FINANCIAL.Name = "ChkHEAD_FINANCIAL"
        Me.ChkHEAD_FINANCIAL.Size = New System.Drawing.Size(125, 15)
        Me.ChkHEAD_FINANCIAL.TabIndex = 67
        Me.ChkHEAD_FINANCIAL.Text = "FINANCIAL"
        '
        'ChkHEAD_FEES
        '
        Me.ChkHEAD_FEES.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChkHEAD_FEES.Location = New System.Drawing.Point(8, 96)
        Me.ChkHEAD_FEES.Name = "ChkHEAD_FEES"
        Me.ChkHEAD_FEES.Size = New System.Drawing.Size(125, 15)
        Me.ChkHEAD_FEES.TabIndex = 66
        Me.ChkHEAD_FEES.Text = "FEES"
        '
        'ChkHEAD_CANDE
        '
        Me.ChkHEAD_CANDE.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChkHEAD_CANDE.Location = New System.Drawing.Point(8, 80)
        Me.ChkHEAD_CANDE.Name = "ChkHEAD_CANDE"
        Me.ChkHEAD_CANDE.Size = New System.Drawing.Size(125, 15)
        Me.ChkHEAD_CANDE.TabIndex = 65
        Me.ChkHEAD_CANDE.Text = "C && E"
        '
        'ChkHEAD_INSPECTION
        '
        Me.ChkHEAD_INSPECTION.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChkHEAD_INSPECTION.Location = New System.Drawing.Point(8, 64)
        Me.ChkHEAD_INSPECTION.Name = "ChkHEAD_INSPECTION"
        Me.ChkHEAD_INSPECTION.Size = New System.Drawing.Size(125, 15)
        Me.ChkHEAD_INSPECTION.TabIndex = 64
        Me.ChkHEAD_INSPECTION.Text = "INSPECTION"
        '
        'ChkHEAD_REGISTRATION
        '
        Me.ChkHEAD_REGISTRATION.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChkHEAD_REGISTRATION.Location = New System.Drawing.Point(8, 48)
        Me.ChkHEAD_REGISTRATION.Name = "ChkHEAD_REGISTRATION"
        Me.ChkHEAD_REGISTRATION.Size = New System.Drawing.Size(125, 15)
        Me.ChkHEAD_REGISTRATION.TabIndex = 63
        Me.ChkHEAD_REGISTRATION.Text = "REGISTRATION"
        '
        'ChkHEAD_PM
        '
        Me.ChkHEAD_PM.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChkHEAD_PM.Location = New System.Drawing.Point(8, 16)
        Me.ChkHEAD_PM.Name = "ChkHEAD_PM"
        Me.ChkHEAD_PM.Size = New System.Drawing.Size(125, 15)
        Me.ChkHEAD_PM.TabIndex = 61
        Me.ChkHEAD_PM.Text = "TECHNICAL"
        '
        'ChkHEAD_ADMIN
        '
        Me.ChkHEAD_ADMIN.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChkHEAD_ADMIN.Location = New System.Drawing.Point(8, 128)
        Me.ChkHEAD_ADMIN.Name = "ChkHEAD_ADMIN"
        Me.ChkHEAD_ADMIN.Size = New System.Drawing.Size(125, 15)
        Me.ChkHEAD_ADMIN.TabIndex = 68
        Me.ChkHEAD_ADMIN.Text = "ADMIN"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(25, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 16)
        Me.Label1.TabIndex = 205
        Me.Label1.Text = "Users"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ugAssignedGroups
        '
        Me.ugAssignedGroups.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugAssignedGroups.Location = New System.Drawing.Point(280, 224)
        Me.ugAssignedGroups.Name = "ugAssignedGroups"
        Me.ugAssignedGroups.Size = New System.Drawing.Size(200, 128)
        Me.ugAssignedGroups.TabIndex = 206
        '
        'ugAvailGroups
        '
        Me.ugAvailGroups.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugAvailGroups.Location = New System.Drawing.Point(24, 224)
        Me.ugAvailGroups.Name = "ugAvailGroups"
        Me.ugAvailGroups.Size = New System.Drawing.Size(200, 128)
        Me.ugAvailGroups.TabIndex = 206
        '
        'ugAvailUsers
        '
        Me.ugAvailUsers.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugAvailUsers.Location = New System.Drawing.Point(24, 376)
        Me.ugAvailUsers.Name = "ugAvailUsers"
        Me.ugAvailUsers.Size = New System.Drawing.Size(200, 128)
        Me.ugAvailUsers.TabIndex = 206
        '
        'ugManagedUsers
        '
        Me.ugManagedUsers.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugManagedUsers.Location = New System.Drawing.Point(280, 376)
        Me.ugManagedUsers.Name = "ugManagedUsers"
        Me.ugManagedUsers.Size = New System.Drawing.Size(200, 128)
        Me.ugManagedUsers.TabIndex = 206
        '
        'ChkEXECUTIVE_DIRECTOR
        '
        Me.ChkEXECUTIVE_DIRECTOR.Font = New System.Drawing.Font("Microsoft Sans Serif", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ChkEXECUTIVE_DIRECTOR.Location = New System.Drawing.Point(8, 144)
        Me.ChkEXECUTIVE_DIRECTOR.Name = "ChkEXECUTIVE_DIRECTOR"
        Me.ChkEXECUTIVE_DIRECTOR.Size = New System.Drawing.Size(136, 15)
        Me.ChkEXECUTIVE_DIRECTOR.TabIndex = 68
        Me.ChkEXECUTIVE_DIRECTOR.Text = "EXECUTIVE DIRECTOR"
        '
        'UserAdmin
        '
        Me.AcceptButton = Me.btnClose
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(504, 590)
        Me.Controls.Add(Me.ugAssignedGroups)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.txtUserName)
        Me.Controls.Add(Me.txtPhone)
        Me.Controls.Add(Me.txtEmail)
        Me.Controls.Add(Me.txtName)
        Me.Controls.Add(Me.ComboUserName)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnNew)
        Me.Controls.Add(Me.btnUserShiftLeftAll)
        Me.Controls.Add(Me.btnUserShiftLeft)
        Me.Controls.Add(Me.btnUserShiftRightAll)
        Me.Controls.Add(Me.btnUserShiftRight)
        Me.Controls.Add(Me.lblSupervisedUsers)
        Me.Controls.Add(Me.lblAvailableUsers)
        Me.Controls.Add(Me.btnGroupShiftLeftAll)
        Me.Controls.Add(Me.btnGroupShiftLeft)
        Me.Controls.Add(Me.btnGroupShiftRightAll)
        Me.Controls.Add(Me.btnGroupShiftRight)
        Me.Controls.Add(Me.lblAssignedGroups)
        Me.Controls.Add(Me.lblAvailableGroups)
        Me.Controls.Add(Me.btnResetPassword)
        Me.Controls.Add(Me.btnSearch)
        Me.Controls.Add(Me.chkInactive)
        Me.Controls.Add(Me.cboPrimaryModule)
        Me.Controls.Add(Me.lblPrimaryModule)
        Me.Controls.Add(Me.lblPhone)
        Me.Controls.Add(Me.lblEmail)
        Me.Controls.Add(Me.lblUserName)
        Me.Controls.Add(Me.lblName)
        Me.Controls.Add(Me.ugAvailGroups)
        Me.Controls.Add(Me.ugAvailUsers)
        Me.Controls.Add(Me.ugManagedUsers)
        Me.Name = "UserAdmin"
        Me.Text = "Manage Users"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.ugAssignedGroups, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ugAvailGroups, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ugAvailUsers, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ugManagedUsers, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region "Form Level Events"
    Protected Overrides Sub OnLoad(ByVal e As System.EventArgs)
        Try
            bolLoading = True
            oUser = New MUSTER.BusinessLogic.pUser
            ClearUserData()
            ' setting bolLoading = True cause the sub sets bolLoading to false
            bolLoading = True
            InitUsers()
            ComboUserName.SelectedIndex = IIf(ComboUserName.Items.Count > 0, 0, -1)
            bolLoading = False
            If ComboUserName.SelectedIndex >= 0 Then
                btnDelete.Enabled = True
                GetUserData()
                DisplayUserData()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)

        If Not oUser Is Nothing Then
            If oUser.IsDirty Then
                Dim Results As Long = MsgBox("There are unsaved changes. Do you want to save changes before closing?", MsgBoxStyle.YesNoCancel)
                If Results = MsgBoxResult.Yes Then
                    oUser.Flush(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
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

        '
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
    Private Sub InitDatatable()
        dtAvailGroups = New DataTable
        dtAssignedGroups = New DataTable

        dtAvailUsers = New DataTable
        dtManagedUsers = New DataTable

        dtAvailGroups.Columns.Add("GROUP_ID", GetType(Integer))
        dtAvailGroups.Columns.Add("GROUP", GetType(String))
        dtAvailGroups.Columns.Add("INACTIVE", GetType(Boolean))

        dtAvailUsers.Columns.Add("STAFF_ID", GetType(Integer))
        dtAvailUsers.Columns.Add("USER", GetType(String))

        dtAssignedGroups = dtAvailGroups.Clone
        dtManagedUsers = dtAvailUsers.Clone
    End Sub
    Private Sub SetupGroupGrid(ByRef ug As Infragistics.Win.UltraWinGrid.UltraGrid)
        ug.DisplayLayout.Bands(0).Columns("GROUP_ID").Hidden = True
        ug.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        ug.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        ug.DisplayLayout.Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Free
        ug.DisplayLayout.Bands(0).Columns("INACTIVE").Width = 60
        ug.DisplayLayout.Bands(0).Columns("INACTIVE").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
        'ug.DisplayLayout.Bands(0).Columns("GROUP").Width = 130
        ug.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.False
        ug.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
    End Sub
    Private Sub SetupUserGrid(ByRef ug As Infragistics.Win.UltraWinGrid.UltraGrid)
        ug.DisplayLayout.Bands(0).Columns("STAFF_ID").Hidden = True
        ug.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        ug.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        'ug.DisplayLayout.Override.AllowColSizing = Infragistics.Win.UltraWinGrid.AllowColSizing.Free
        ug.DisplayLayout.Bands(0).Columns("USER").Width = 180
        ug.DisplayLayout.Override.RowSelectors = Infragistics.Win.DefaultableBoolean.False
        ug.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
    End Sub
    Private Sub InitUsers()
        ComboUserName.DataSource = oUser.ListUserNames
        ComboUserName.DisplayMember = "USER_ID"
        ComboUserName.ValueMember = "STAFF_ID"
        LoadPrimaryModules()
    End Sub
    Private Sub LoadPrimaryModules()
        cboPrimaryModule.DataSource = oUser.ListPrimaryModules
        cboPrimaryModule.DisplayMember = "PROPERTY_NAME"
        cboPrimaryModule.ValueMember = "PROPERTY_ID"
        'Dim oModProp As New MUSTER.BusinessLogic.pPropertyType("Modules")
        'Dim dtModules As DataTable
        'dtModules = oModProp.PropertiesTable
        'dtModules.DefaultView.RowFilter = "PropType_ID=" + oModProp.ID.ToString
        'Me.cboPrimaryModule.DataSource = dtModules.DefaultView 'oModProp.PropertiesTable
        'Me.cboPrimaryModule.DisplayMember = "Property Name"
        'Me.cboPrimaryModule.ValueMember = "Property ID"
    End Sub
    Private Sub GetUserData()
        Dim strOldUserName As String = oUser.ID
        If bolIsNewUser Then
            Exit Sub
        End If

        strPreviousUserName = ComboUserName.Text
        If ComboUserName.SelectedIndex > -1 Then
            oUser.Retrieve(UIUtilsGen.GetComboBoxValueInt(ComboUserName))
        End If
    End Sub

    Private Sub DisplayUserData()
        If Not bolIsNewUser Then Me.txtUserName.ReadOnly = True

        If ComboUserName.SelectedIndex < 0 Then Exit Sub
        ComboUserName.Enabled = True

        If ComboUserName.Text = String.Empty And Not bolIsNewUser Then Exit Sub

        Me.ChkHEAD_ADMIN.Enabled = Not oUser.ModuleHeadCheck("HEAD_ADMIN", oUser.ID)
        Me.ChkHEAD_CANDE.Enabled = Not oUser.ModuleHeadCheck("HEAD_CANDE", oUser.ID)
        Me.ChkHEAD_CLOSURE.Enabled = Not oUser.ModuleHeadCheck("HEAD_CLOSURE", oUser.ID)
        Me.ChkHEAD_FEES.Enabled = Not oUser.ModuleHeadCheck("HEAD_FEES", oUser.ID)
        Me.ChkHEAD_FINANCIAL.Enabled = Not oUser.ModuleHeadCheck("HEAD_FINANCIAL", oUser.ID)
        Me.ChkHEAD_INSPECTION.Enabled = Not oUser.ModuleHeadCheck("HEAD_INSPECTION", oUser.ID)
        Me.ChkHEAD_PM.Enabled = Not oUser.ModuleHeadCheck("HEAD_PM", oUser.ID)
        Me.ChkHEAD_REGISTRATION.Enabled = Not oUser.ModuleHeadCheck("HEAD_REGISTRATION", oUser.ID)
        Me.ChkEXECUTIVE_DIRECTOR.Enabled = Not oUser.ModuleHeadCheck("EXECUTIVE_DIRECTOR", oUser.ID)

        'Used to determine the previous user
        If Not bolIsNewUser Then strPreviousUserName = Me.ComboUserName.Text
        If Not bolIsNewUser Then Me.btnDelete.Enabled = True
        If Not bolIsNewUser Then Me.txtUserName.Text = Me.ComboUserName.Text
        If Not bolIsNewUser Then ComboUserName.Enabled = True
        If bolIsNewUser Then Me.txtUserName.Text = oUser.ID
        Me.btnResetPassword.Enabled = True
        Me.txtName.Text = oUser.Name
        Me.txtEmail.Text = oUser.EmailAddress
        Me.txtPhone.Text = oUser.PhoneNumber
        Me.chkInactive.Checked = oUser.Active

        Me.ChkHEAD_ADMIN.Checked = oUser.HEAD_ADMIN
        Me.ChkHEAD_CANDE.Checked = oUser.HEAD_CANDE
        Me.ChkHEAD_CLOSURE.Checked = oUser.HEAD_CLOSURE
        Me.ChkHEAD_FEES.Checked = oUser.HEAD_FEES
        Me.ChkHEAD_FINANCIAL.Checked = oUser.HEAD_FINANCIAL
        Me.ChkHEAD_INSPECTION.Checked = oUser.HEAD_INSPECTION
        Me.ChkHEAD_PM.Checked = oUser.HEAD_PM
        Me.ChkHEAD_REGISTRATION.Checked = oUser.HEAD_REGISTRATION
        Me.ChkEXECUTIVE_DIRECTOR.Checked = oUser.EXECUTIVE_DIRECTOR

        UIUtilsGen.SetComboboxItemByValue(cboPrimaryModule, oUser.DefaultModule)
        'Me.cboPrimaryModule.SelectedValue = IIf(oUser.DefaultModule Is Nothing, ComboUserName.Text, oUser.DefaultModule)

        bolLoading = False
        DisplayScreenDisposition()
    End Sub

    Private Sub DisplayScreenDisposition()
        If ComboUserName.SelectedValue.ToString <> "System.Data.DataRowView" Then
            If ComboUserName.SelectedValue.ToString <> String.Empty Then
                LoadScreenListviews()
            End If
        End If
    End Sub

    Private Sub ClearScreenData()
        ugAvailGroups.DataSource = Nothing
        ugAssignedGroups.DataSource = Nothing
        ugAvailUsers.DataSource = Nothing
        ugManagedUsers.DataSource = Nothing

        dtAvailGroups.Rows.Clear()
        dtAssignedGroups.Rows.Clear()
        dtAvailUsers.Rows.Clear()
        dtManagedUsers.Rows.Clear()
    End Sub

    Private Sub LoadScreenListviews()
        ClearScreenData()
        For Each userGroupRelInfo In oUser.UserGroupRelationCollection.Values
            If userGroupRelInfo.StaffID = 0 Or userGroupRelInfo.Deleted Then
                dr = dtAvailGroups.NewRow
                dr("GROUP_ID") = userGroupRelInfo.GroupID
                dr("GROUP") = userGroupRelInfo.GroupName
                dr("INACTIVE") = userGroupRelInfo.Inactive
                dtAvailGroups.Rows.Add(dr)
            Else
                dr = dtAssignedGroups.NewRow
                dr("GROUP_ID") = userGroupRelInfo.GroupID
                dr("GROUP") = userGroupRelInfo.GroupName
                dr("INACTIVE") = userGroupRelInfo.Inactive
                dtAssignedGroups.Rows.Add(dr)
            End If
        Next

        For Each userInfo In oUser.ManagedUsersCollection.Values
            If userInfo.ManagerID = oUser.UserKey Then
                dr = dtManagedUsers.NewRow
                dr("STAFF_ID") = userInfo.UserKey
                dr("USER") = userInfo.ID
                dtManagedUsers.Rows.Add(dr)
            Else
                dr = dtAvailUsers.NewRow
                dr("STAFF_ID") = userInfo.UserKey
                dr("USER") = userInfo.ID
                dtAvailUsers.Rows.Add(dr)
            End If
        Next

        dtAvailGroups.DefaultView.Sort = "GROUP"
        dtAssignedGroups.DefaultView.Sort = "GROUP"
        dtAvailUsers.DefaultView.Sort = "USER"
        dtManagedUsers.DefaultView.Sort = "USER"

        ugAvailGroups.DataSource = dtAvailGroups.DefaultView
        ugAssignedGroups.DataSource = dtAssignedGroups.DefaultView
        ugAvailUsers.DataSource = dtAvailUsers.DefaultView
        ugManagedUsers.DataSource = dtManagedUsers.DefaultView

        If ugAvailGroups.Rows.Count > 0 Then
            ugAvailGroups.ActiveRow = ugAvailGroups.Rows(0)
        End If
        If ugAssignedGroups.Rows.Count > 0 Then
            ugAssignedGroups.ActiveRow = ugAssignedGroups.Rows(0)
        End If
        If ugAvailUsers.Rows.Count > 0 Then
            ugAvailUsers.ActiveRow = ugAvailUsers.Rows(0)
        End If
        If ugManagedUsers.Rows.Count > 0 Then
            ugManagedUsers.ActiveRow = ugManagedUsers.Rows(0)
        End If

        LeftRightEnable()
    End Sub

    Private Sub SetSaveCancel(ByVal bolValue As Boolean)
        btnSave.Enabled = bolValue
        btnCancel.Enabled = bolValue
    End Sub

    Private Function LeftRightEnable()
        Try
            If Me.ugAvailGroups.Rows.Count > 0 Then
                Me.btnGroupShiftRight.Enabled = True
                If Me.ugAvailGroups.Rows.Count > 1 Then
                    Me.btnGroupShiftRightAll.Enabled = True
                Else
                    Me.btnGroupShiftRightAll.Enabled = False
                End If
            Else
                Me.btnGroupShiftRight.Enabled = False

            End If

            If Me.ugAssignedGroups.Rows.Count > 0 Then
                Me.btnGroupShiftLeft.Enabled = True
                If Me.ugAssignedGroups.Rows.Count > 1 Then
                    Me.btnGroupShiftLeftAll.Enabled = True
                Else
                    Me.btnGroupShiftLeftAll.Enabled = False
                End If
            Else
                Me.btnGroupShiftLeft.Enabled = False

            End If

            If Me.ugAvailUsers.Rows.Count > 0 Then
                Me.btnUserShiftRight.Enabled = True
                If Me.ugAvailUsers.Rows.Count > 1 Then
                    Me.btnUserShiftRightAll.Enabled = True
                Else
                    Me.btnUserShiftRightAll.Enabled = False
                End If
            Else
                Me.btnUserShiftRight.Enabled = False

            End If

            If Me.ugManagedUsers.Rows.Count > 0 Then
                Me.btnUserShiftLeft.Enabled = True
                If Me.ugManagedUsers.Rows.Count > 1 Then
                    Me.btnUserShiftLeftAll.Enabled = True
                Else
                    Me.btnUserShiftLeftAll.Enabled = False
                End If
            Else
                Me.btnUserShiftLeft.Enabled = False

            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Private Sub CheckIfDataDirty()
        If oUser.colIsDirty Then
            Dim msgRet As MsgBoxResult = MsgBox("The user data for " & oUser.Name & " has changed.  Do you wish to save the changes?", MsgBoxStyle.Question & MsgBoxStyle.YesNo, "Data Changed")
            If msgRet = MsgBoxResult.Yes Then
                oUser.Flush(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If
            Else
                oUser.Reset()
            End If
        End If
        'Dim sender As System.Object
        'Dim e As System.EventArgs
        'If Not oUser Is Nothing Then
        '    If oUser.IsDirty Then
        '        Dim Results As Long = MsgBox("There are unsaved changes. Do you want to save changes before closing?", MsgBoxStyle.YesNoCancel)
        '        If Results = MsgBoxResult.Yes Then
        '            SaveForms()
        '        Else
        '            oUser.Reset()
        '        End If
        '    End If
        'End If
    End Sub

    'Private Sub ResetForm()
    '    Try
    '        oUser = New MUSTER.BusinessLogic.pUser
    '        'oGroups = New MUSTER.BusinessLogic.pUserGroupMemberships
    '        'LoadUsers()
    '        LoadPrimaryModules()
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub

    'Private Sub LoadUsers()
    '    ComboUserName.DataSource = oUser.ListUserNames
    '    ComboUserName.DisplayMember = "USER_ID"
    '    ComboUserName.ValueMember = "STAFF_ID"
    '    'ComboUserName.ValueMember = "USER_ID"
    'End Sub

    Private Sub ClearUserData()

        bolLoading = True

        Me.txtUserName.Text = String.Empty
        Me.txtName.Text = String.Empty
        Me.txtEmail.Text = String.Empty
        Me.txtPhone.Text = String.Empty
        Me.cboPrimaryModule.SelectedIndex = -1

        'ComboUserName.SelectedIndex = -1
        Me.chkInactive.Checked = False
        Me.lblpsw.Visible = False
        Me.btnResetPassword.Enabled = False
        Me.btnDelete.Enabled = False

        ClearScreenData()
        'Me.txtName.Text = String.Empty
        'Me.txtEmail.Text = String.Empty
        'Me.txtPhone.Text = String.Empty
        'Me.cboPrimaryModule.SelectedIndex = -1
        ''ComboUserName.SelectedIndex = -1
        'Me.chkInactive.Checked = False
        'Me.lblpsw.Visible = False

        Me.ChkHEAD_ADMIN.Checked = False
        Me.ChkHEAD_CANDE.Checked = False
        Me.ChkHEAD_CLOSURE.Checked = False
        Me.ChkHEAD_FEES.Checked = False
        Me.ChkHEAD_FINANCIAL.Checked = False
        Me.ChkHEAD_INSPECTION.Checked = False
        Me.ChkHEAD_PM.Checked = False
        Me.ChkHEAD_REGISTRATION.Checked = False
        Me.ChkEXECUTIVE_DIRECTOR.Checked = False

        Me.cboPrimaryModule.SelectedIndex = -1
        bolLoading = False

    End Sub

    Private Sub IsEmptyUserName()
        Try
            If Not bolIsNewUser Then
                If ComboUserName.Text = String.Empty Then
                    MsgBox("User Names may not be blank.", MsgBoxStyle.Exclamation & MsgBoxStyle.OKOnly, "User Error")
                    ComboUserName.Focus()
                    bolIsNewUser = False
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    'Private Sub ListAvailableGroups()
    '    Dim datarow As DataRow
    '    Dim dtAvailableUserGroups As DataTable
    '    lstAvailableGroups.Items.Clear()
    '    dtAvailableUserGroups = oGroups.ListGroups
    '    If Not dtAvailableUserGroups Is Nothing Then
    '        For Each datarow In dtAvailableUserGroups.Rows
    '            Me.lstAvailableGroups.Items.Add(datarow.Item("GROUP_NAME"))
    '        Next
    '    End If
    'End Sub

    'Private Sub ListAvailableSupervisers()
    '    Dim datarow As DataRow
    '    Dim dtUnSuoUsers As DataTable
    '    dtUnSuoUsers = oUser.ListUnSupervisedUsers
    '    lstAvailableUsers.Items.Clear()
    '    If Not dtUnSuoUsers Is Nothing Then
    '        For Each datarow In dtUnSuoUsers.Rows
    '            If oUser.ManagerID <> datarow.Item("STAFF_ID") And oUser.UserKey <> datarow.Item("staff_ID") Then
    '                Me.lstAvailableUsers.Items.Add(datarow.Item("USER_ID"))
    '            End If
    '        Next
    '    End If
    'End Sub

    Private Function SaveForms() As Boolean
        Dim bolSuccess As Boolean
        Dim strErr As String = String.Empty
        If Me.ComboUserName.Text <> "" Then 'And UIUtills.IsEmailValid(Me.txtEmail.Text) Then
            Try

                Me.lblpsw.Visible = False

                oUser.Save(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Function
                End If
                oUser.Reset()
                bolSuccess = True
            Catch ex As Exception

                Throw ex
                '  End If

            End Try
        Else
            If Me.ComboUserName.Text = String.Empty Then
                strErr += vbTab + "You must specify a username" + vbCrLf
                'Me.ComboUserName.Focus()
            End If
            If Not Me.txtEmail.Text = String.Empty Then
                If Not UIUtilsGen.IsEmailValid(Me.txtEmail.Text) Then
                    strErr += vbTab + "Invalid Email address" + vbCrLf
                End If
            End If
            If strErr.Length > 0 Then
                MsgBox("Invalid/Incomplete User:" + vbCrLf + strErr)
                If Me.ComboUserName.Text = String.Empty Then
                    Me.ComboUserName.Focus()
                End If
            End If
            bolSuccess = False
        End If
        Return bolSuccess
    End Function

    Public Function VerifyPassword(ByVal OldPsw As String, ByVal NewPsw As String) As Boolean
        Try
            oUser.Password = OldPsw
            If oUser.VerifyPassword = True Then
                oUser.Password = NewPsw
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    'Private Sub CheckCurrentUserState()

    '    If oUser.colIsDirty Then
    '        Dim msgRet As MsgBoxResult = MsgBox("The user data has changed.  Do you wish to save the changes?", MsgBoxStyle.Question & MsgBoxStyle.YesNo, "Data Changed")
    '        If msgRet = MsgBoxResult.Yes Then

    '            oUser.Flush(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
    '            If Not UIUtilsGen.HasRights(returnVal) Then
    '                Exit Sub
    '            End If

    '        Else
    '            oUser.ResetCollection()
    '        End If
    '    End If

    'End Sub
#End Region
#Region "UI Control Events"

    Private Sub btnNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNew.Click
        Try
            CheckIfDataDirty()
            bolLoading = True
            bolIsNewUser = True
            ClearUserData()
            ClearScreenData()
            bolLoading = True
            ComboUserName.SelectedIndex = -1
            ComboUserName.Text = String.Empty
            ComboUserName.Enabled = False
            txtUserName.ReadOnly = False
            txtUserName.Text = ""
            txtUserName.Focus()
            Me.cboPrimaryModule.SelectedIndex = -1
            bolLoading = False
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            If Not oUser Is Nothing Then
                oUser.Reset()
                If bolIsNewUser Then bolIsNewUser = False
                If bolIsNewUser Then oUser.Remove(oUser.UserKey)
                If strPreviousUserName <> String.Empty Then
                    ComboUserName.SelectedIndex = ComboUserName.FindStringExact(strPreviousUserName)
                Else
                    ComboUserName.SelectedIndex = 0
                End If
                ClearUserData()
                GetUserData()
                DisplayUserData()
            End If

            'Checking if user object has data,if so then reset and display the data.
            'If Not oUser Is Nothing Then

            '    ComboUserName.Enabled = True
            '    txtUserName.ReadOnly = True
            '    txtUserName.Text = ""

            '    If bolIsNewUser Then
            '        oUser.Remove(oUser.ID)
            '        bolIsNewUser = False
            '    Else
            '        oUser.ResetCollection()
            '    End If
            '    ClearUserData()
            '    bolLoading = True
            '    If strPreviousUserName <> String.Empty Then
            '        ComboUserName.SelectedIndex = ComboUserName.FindStringExact(strPreviousUserName)
            '    Else
            '        Me.ComboUserName.SelectedIndex = 1
            '    End If
            '    bolLoading = False
            '    If IsNothing(ComboUserName) Then
            '        ResetForm()
            '        Exit Sub
            '    End If
            '    If IsNothing(ComboUserName.SelectedValue) Then
            '        ComboUserName.SelectedIndex = 0
            '    End If
            '    oUser.Retrieve(ComboUserName.SelectedValue)
            '    DisplayUserData()
            '    'If oUser.UserKey <> 0 Then
            '    '    DisplayUserData()
            '    'End If
            'End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            ' validation
            Dim strErr As String = String.Empty
            If oUser.ID = String.Empty Then
                strErr = "User Name is required" + vbCrLf
            End If
            If oUser.Name = String.Empty Then
                strErr += "Name is required" + vbCrLf
            End If
            'If oUser.DefaultModule = String.Empty Then
            '    strErr += "Primary Module is required"
            'End If
            If oUser.DefaultModule <= 0 Then
                strErr += "Primary Module is required"
            End If
            If strErr <> String.Empty Then
                MsgBox(strErr, , "User Validation")
                Exit Sub
            End If

            If oUser.UserKey <= 0 Then
                oUser.CreatedBy = MusterContainer.AppUser.ID
            Else
                oUser.ModifiedBy = MusterContainer.AppUser.ID
            End If
            oUser.Save(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If
            If bolIsNewUser Then
                bolIsNewUser = False
                ComboUserName.Enabled = True
                bolLoading = True
                Me.InitUsers()
                bolLoading = False
                ComboUserName.SelectedIndex = ComboUserName.FindStringExact(oUser.ID)
            End If
            MsgBox("User Save Successful", 0, "MUSTER Data Access")

            SetSaveCancel(oUser.IsDirty)

            'Dim bolSuccess As Boolean
            'ComboUserName.Enabled = True
            'txtUserName.ReadOnly = True

            'oUser.Save(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
            'If Not UIUtilsGen.HasRights(returnVal) Then
            '    Exit Sub
            'End If

            'MsgBox("User Save Successful", 0, "MUSTER Data Access")

            'If bolIsNewUser Then
            '    bolIsNewUser = False
            '    strPreviousUserName = Me.txtUserName.Text
            '    txtUserName.Text = ""
            '    ClearUserData()
            '    Me.LoadUsers()
            '    bolLoading = True
            '    ComboUserName.SelectedIndex = ComboUserName.FindStringExact(strPreviousUserName)
            '    bolLoading = False
            '    oUser.Retrieve(strPreviousUserName)
            '    DisplayUserData()
            'End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim msgResult As MsgBoxResult = MsgBox("Are you sure you wish to DELETE the user: " & ComboUserName.Text & " ?", MsgBoxStyle.Question & MsgBoxStyle.YesNo, "DELETE USER")
        If msgResult = MsgBoxResult.Yes Then
            Dim bolLoadingLocal As Boolean = bolLoading
            Try
                bolLoading = True
                oUser.Deleted = True
                For Each userGroupRelInfo In oUser.UserGroupRelationCollection.Values
                    userGroupRelInfo.Deleted = True
                Next
                For Each userInfo In oUser.ManagedUsersCollection.Values
                    userInfo.ManagerID = 0
                Next
                oUser.Save(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If
                MsgBox("User Delete Successful", 0, "MUSTER Data Access")

                ' to prevent any change events on ui
                oUser.Reset()
                oUser.Remove(oUser.UserKey)

                Me.InitUsers()
                ClearScreenData()
                Me.ClearUserData()
                'bolLoading = True
                'ComboUserName.SelectedIndex = -1
                'ComboUserName.Text = String.Empty
                'ClearUserData()
                'Me.InitUsers()
                ''ComboUserName.Focus()
                'Me.txtUserName.ReadOnly = True
                'Me.ComboUserName.Focus()
                'bolLoading = False
            Finally
                bolLoading = False
                ComboUserName.SelectedIndex = 0
                bolLoading = bolLoadingLocal
            End Try
        End If
    End Sub

    Private Sub txtUserName_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUserName.Leave
        Try
            If txtUserName.ReadOnly = False Then
                Dim bolcurrentUser As Boolean = Not (ComboUserName.FindStringExact(Me.txtUserName.Text) > -1)
                If Me.txtUserName.Text <> String.Empty Then
                    If bolcurrentUser Then
                        'Dim msgResult As MsgBoxResult = MsgBox("The user " & ComboUserName.Text & " does not exist!  Do you wish to create the user and a base set of permissions for the user?", MsgBoxStyle.Question & MsgBoxStyle.YesNo, "New User Permissions")
                        'If msgResult = MsgBoxResult.Yes Then
                        Dim strUserName As String = Me.txtUserName.Text
                        Dim oLocalUser As New MUSTER.Info.UserInfo(0, Me.txtUserName.Text, String.Empty, String.Empty, String.Empty, 612, 0, String.Empty, False, False, False, False, False, False, False, False, False, False, MusterContainer.AppUser.ID, Now, String.Empty, CDate("01/01/0001"), False)
                        oLocalUser.Password = "password"
                        oUser.Add(oLocalUser)
                        oUser.Retrieve(strUserName)
                        LoadScreenListviews()
                        'DisplayScreenDisposition()
                        'oUser.Add(strUserName) - commented out 5/31/05 JVC II - overwriting profile
                        'oUser.Save()
                        'LoadUsers()
                        'ComboUserName.SelectedIndex = ComboUserName.FindStringExact(strUserName)
                        'oUser.Retrieve(strUserName)
                        'Else
                        '        bolLoading = False
                        '        bolisNewUser = False
                        '        ClearUserData()
                        '        Exit Sub
                        'End If
                    Else
                        '
                        ' Need this to perform UI reset if not a new user
                        '

                        If Me.txtUserName.Text <> String.Empty Then
                            MsgBox("A user with that username already exists.")
                            strPreviousUserName = Me.txtUserName.Text
                            bolIsNewUser = False
                            btnCancel_Click(sender, e)
                            'ComboUserName.Text = Me.txtUserName.Text
                            '    ComboUserName.SelectedIndex = ComboUserName.FindStringExact(strOldUserName)
                            'oUser.Retrieve(ComboUserName.Text)
                        End If

                    End If


                Else
                    Dim msgResult As MsgBoxResult = MsgBox("Username cannot be blank.Would you like to continue creating a new user?", MsgBoxStyle.Question & MsgBoxStyle.YesNo, "NEW USER")
                    If msgResult = MsgBoxResult.No Then
                        bolIsNewUser = False
                        If strPreviousUserName <> String.Empty Then
                            ComboUserName.SelectedIndex = ComboUserName.FindStringExact(strPreviousUserName)
                        Else
                            Me.ComboUserName.SelectedIndex = 1
                        End If
                    Else
                        Me.txtUserName.Focus()
                        Exit Sub
                    End If
                End If

                'ClearUserData()
                'bolLoading = False

                'If Me.ComboUserName.SelectedValue = String.Empty Then
                '    oUser.ID = Me.ComboUserName.Text
                'End If
                'oUser.Retrieve(ComboUserName.Text)
                'DisplayUserData()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    'Private Sub ComboUserName_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboUserName.Leave

    '    Dim drUserName As DataRow
    '    Dim itemExists As Boolean = False
    '    Dim strUname As String
    '    Dim strOldUserName As String = oUser.ID

    '    Try
    '        If strOldUserName = ComboUserName.Text Then Exit Sub

    '        bolIsNewUser = Not (ComboUserName.FindStringExact(ComboUserName.Text) > -1)
    '        bolLoading = True
    '        If bolIsNewUser And ComboUserName.Text <> String.Empty Then
    '            Dim msgResult As MsgBoxResult = MsgBox("The user " & ComboUserName.Text & " does not exist!  Do you wish to create the user and a base set of permissions for the user?", MsgBoxStyle.Question & MsgBoxStyle.YesNo, "New User Permissions")
    '            If msgResult = MsgBoxResult.Yes Then
    '                Dim strUserName As String = ComboUserName.Text
    '                Dim oLocalUser As New MUSTER.Info.UserInfo(0, ComboUserName.Text, String.Empty, String.Empty, String.Empty, "Registration", 0, String.Empty, False, False, False, False, False, False, False, False, False, True, MusterContainer.AppUser.ID, Now, String.Empty, CDate("01/01/0001"))
    '                oLocalUser.Password = "password"
    '                oUser.Add(oLocalUser)
    '                'oUser.Add(strUserName)
    '                oUser.Save(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
    '                If Not UIUtilsGen.HasRights(returnVal) Then
    '                    Exit Sub
    '                End If

    '                InitUsers()
    '                ComboUserName.SelectedIndex = ComboUserName.FindStringExact(strUserName)
    '                oUser.Retrieve(ComboUserName.Text)
    '            Else
    '                If strOldUserName <> String.Empty Then
    '                    ComboUserName.SelectedIndex = ComboUserName.FindStringExact(strOldUserName)
    '                Else
    '                    bolLoading = False
    '                    bolIsNewUser = False
    '                    ClearUserData()
    '                    Exit Sub
    '                End If
    '            End If
    '        Else
    '            '
    '            ' Need this to perform UI reset if not a new user
    '            '
    '            If ComboUserName.Text <> String.Empty Then
    '                '    ComboUserName.SelectedIndex = ComboUserName.FindStringExact(strOldUserName)
    '                oUser.Retrieve(ComboUserName.Text)
    '            End If

    '        End If
    '        ClearUserData()
    '        bolLoading = False

    '        'If Me.ComboUserName.SelectedValue = String.Empty Then
    '        '    oUser.ID = Me.ComboUserName.Text
    '        'End If
    '        'oUser.Retrieve(ComboUserName.Text)
    '        DisplayUserData()

    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub

    Private Sub ComboUserName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboUserName.SelectedIndexChanged
        If bolLoading Then Exit Sub
        CheckIfDataDirty()
        GetUserData()
        DisplayUserData()
        'If ComboUserName.SelectedValue <> String.Empty And Me.ComboUserName.SelectedValue <> Me.txtUserName.Text Then
        '    Try
        '        CheckCurrentUserState()
        '        ClearUserData()
        '        ComboUserName_Leave(sender, e)
        '    Catch ex As Exception
        '        Throw ex
        '    End Try
        'End If
    End Sub

    Private Sub btnResetPassword_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnResetPassword.Click
        Try
            'oChangePassword.SetUser(oUser)
            'oChangePassword.ShowDialog()
            oUser.Password = "password"
            oUser.Flush(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If
            MsgBox("User password reset Successful")
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub txtName_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtName.TextChanged
        If bolLoading Then Exit Sub
        oUser.Name = Me.txtName.Text
    End Sub

    Private Sub txtEmail_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEmail.TextChanged
        If bolLoading Then Exit Sub
        oUser.EmailAddress = Me.txtEmail.Text
    End Sub

    Private Sub txtPhone_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPhone.TextChanged
        If bolLoading Then Exit Sub
        oUser.PhoneNumber = Me.txtPhone.Text
    End Sub

    Private Sub chkInactive_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkInactive.Click
        If bolLoading Then Exit Sub
        oUser.Active = Me.chkInactive.Checked
    End Sub

    Private Sub cboPrimaryModule_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPrimaryModule.SelectedIndexChanged
        If bolLoading Then Exit Sub
        oUser.DefaultModule = UIUtilsGen.GetComboBoxValueInt(cboPrimaryModule)
        'If Me.cboPrimaryModule.SelectedValue <> "" Then
        '    oUser.DefaultModule = Me.cboPrimaryModule.SelectedValue
        'End If
    End Sub

    Private Sub txtName_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtName.Enter
        Try
            IsEmptyUserName()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub txtEmail_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEmail.Enter
        Try
            IsEmptyUserName()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub txtPhone_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPhone.Enter
        Try
            IsEmptyUserName()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub chkInactive_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkInactive.Enter
        Try
            IsEmptyUserName()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ChkHEAD_PM_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkHEAD_PM.CheckedChanged
        If bolLoading Then Exit Sub
        oUser.HEAD_PM = Me.ChkHEAD_PM.Checked
    End Sub

    Private Sub ChkHEAD_CLOSURE_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkHEAD_CLOSURE.CheckedChanged
        If bolLoading Then Exit Sub
        oUser.HEAD_CLOSURE = Me.ChkHEAD_CLOSURE.Checked
    End Sub

    Private Sub ChkHEAD_REGISTRATION_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkHEAD_REGISTRATION.CheckedChanged
        If bolLoading Then Exit Sub
        oUser.HEAD_REGISTRATION = Me.ChkHEAD_REGISTRATION.Checked
    End Sub

    Private Sub ChkHEAD_INSPECTION_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkHEAD_INSPECTION.CheckedChanged
        If bolLoading Then Exit Sub
        oUser.HEAD_INSPECTION = Me.ChkHEAD_INSPECTION.Checked
    End Sub

    Private Sub ChkHEAD_CANDE_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkHEAD_CANDE.CheckedChanged
        If bolLoading Then Exit Sub
        oUser.HEAD_CANDE = Me.ChkHEAD_CANDE.Checked
    End Sub

    Private Sub ChkHEAD_FEES_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkHEAD_FEES.CheckedChanged
        If bolLoading Then Exit Sub
        oUser.HEAD_FEES = Me.ChkHEAD_FEES.Checked
    End Sub

    Private Sub ChkHEAD_FINANCIAL_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkHEAD_FINANCIAL.CheckedChanged
        If bolLoading Then Exit Sub
        oUser.HEAD_FINANCIAL = Me.ChkHEAD_FINANCIAL.Checked
    End Sub

    Private Sub ChkHEAD_ADMIN_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkHEAD_ADMIN.CheckedChanged
        If bolLoading Then Exit Sub
        oUser.HEAD_ADMIN = Me.ChkHEAD_ADMIN.Checked
    End Sub

    Private Sub ChkEXECUTIVE_DIRECTOR_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChkEXECUTIVE_DIRECTOR.CheckedChanged
        If bolLoading Then Exit Sub
        oUser.EXECUTIVE_DIRECTOR = Me.ChkEXECUTIVE_DIRECTOR.Checked
    End Sub


    Private Sub btnGroupShiftRight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGroupShiftRight.Click
        Try
            If Not ugAvailGroups.Selected.Rows Is Nothing Then
                If ugAvailGroups.Selected.Rows.Count > 0 Then
                    For Each ugRow In ugAvailGroups.Selected.Rows
                        userGroupRelInfo = oUser.UserGroupRelationCollection.Item(oUser.UserKey.ToString + "|" + ugRow.Cells("GROUP_ID").Text)
                        If userGroupRelInfo Is Nothing Then
                            userGroupRelInfo = oUser.UserGroupRelationCollection.Item("0|" + ugRow.Cells("GROUP_ID").Text)
                            userGroupRelInfo.StaffID = oUser.UserKey
                            oUser.UserGroupRelationCollection.ChangeKey("0|" + ugRow.Cells("GROUP_ID").Text, oUser.UserKey.ToString + "|" + ugRow.Cells("GROUP_ID").Text)
                        End If
                        userGroupRelInfo.Deleted = False
                    Next
                    SetSaveCancel(oUser.IsDirty)
                    LoadScreenListviews()
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnGroupShiftLeft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGroupShiftLeft.Click
        Try
            If Not ugAssignedGroups.Selected.Rows Is Nothing Then
                If ugAssignedGroups.Selected.Rows.Count > 0 Then
                    For Each ugRow In ugAssignedGroups.Selected.Rows
                        userGroupRelInfo = oUser.UserGroupRelationCollection.Item(oUser.UserKey.ToString + "|" + ugRow.Cells("GROUP_ID").Text)
                        If userGroupRelInfo.isNew Then
                            oUser.UserGroupRelationCollection.ChangeKey(userGroupRelInfo.ID, "0|" + userGroupRelInfo.GroupID.ToString)
                            userGroupRelInfo.StaffID = 0
                        Else
                            userGroupRelInfo.Deleted = True
                        End If
                    Next
                    SetSaveCancel(oUser.IsDirty)
                    LoadScreenListviews()
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnGroupShiftRightAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGroupShiftRightAll.Click
        Try
            For Each ugRow In ugAvailGroups.Rows
                userGroupRelInfo = oUser.UserGroupRelationCollection.Item(oUser.UserKey.ToString + "|" + ugRow.Cells("GROUP_ID").Text)
                If userGroupRelInfo Is Nothing Then
                    userGroupRelInfo = oUser.UserGroupRelationCollection.Item("0|" + ugRow.Cells("GROUP_ID").Text)
                    userGroupRelInfo.StaffID = oUser.UserKey
                    oUser.UserGroupRelationCollection.ChangeKey("0|" + ugRow.Cells("GROUP_ID").Text, oUser.UserKey.ToString + "|" + ugRow.Cells("GROUP_ID").Text)
                End If
                userGroupRelInfo.Deleted = False
            Next
            SetSaveCancel(oUser.IsDirty)
            LoadScreenListviews()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnGroupShiftLeftAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGroupShiftLeftAll.Click
        Try
            For Each ugRow In ugAssignedGroups.Rows
                userGroupRelInfo = oUser.UserGroupRelationCollection.Item(oUser.UserKey.ToString + "|" + ugRow.Cells("GROUP_ID").Text)
                If userGroupRelInfo.isNew Then
                    oUser.UserGroupRelationCollection.ChangeKey(userGroupRelInfo.ID, "0|" + userGroupRelInfo.GroupID.ToString)
                    userGroupRelInfo.StaffID = 0
                Else
                    userGroupRelInfo.Deleted = True
                End If
            Next
            SetSaveCancel(oUser.IsDirty)
            LoadScreenListviews()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnUserShiftRight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUserShiftRight.Click
        Try
            If Not ugAvailUsers.Selected.Rows Is Nothing Then
                If ugAvailUsers.Selected.Rows.Count > 0 Then
                    For Each ugRow In ugAvailUsers.Selected.Rows
                        userInfo = oUser.ManagedUsersCollection.Item(ugRow.Cells("STAFF_ID").Text)
                        userInfo.ManagerID = oUser.UserKey
                    Next
                    SetSaveCancel(oUser.IsDirty)
                    LoadScreenListviews()
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnUserShiftRightAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUserShiftRightAll.Click
        Try
            For Each ugRow In ugAvailUsers.Rows
                userInfo = oUser.ManagedUsersCollection.Item(ugRow.Cells("STAFF_ID").Text)
                userInfo.ManagerID = oUser.UserKey
            Next
            SetSaveCancel(oUser.IsDirty)
            LoadScreenListviews()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnUserShiftLeft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUserShiftLeft.Click
        Try
            If Not ugManagedUsers.Selected.Rows Is Nothing Then
                If ugManagedUsers.Selected.Rows.Count > 0 Then
                    For Each ugRow In ugManagedUsers.Selected.Rows
                        userInfo = oUser.ManagedUsersCollection.Item(ugRow.Cells("STAFF_ID").Text)
                        userInfo.ManagerID = 0
                    Next
                    SetSaveCancel(oUser.IsDirty)
                    LoadScreenListviews()
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnUserShiftLeftAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUserShiftLeftAll.Click
        Try
            For Each ugRow In ugManagedUsers.Rows
                userInfo = oUser.ManagedUsersCollection.Item(ugRow.Cells("STAFF_ID").Text)
                userInfo.ManagerID = 0
            Next
            SetSaveCancel(oUser.IsDirty)
            LoadScreenListviews()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub


    Private Sub ugAssignedGroups_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugAssignedGroups.InitializeLayout
        SetupGroupGrid(sender)
    End Sub

    Private Sub ugAvailGroups_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugAvailGroups.InitializeLayout
        SetupGroupGrid(sender)
    End Sub

    Private Sub ugAvailUsers_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugAvailUsers.InitializeLayout
        SetupUserGrid(sender)
    End Sub

    Private Sub ugManagedUsers_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugManagedUsers.InitializeLayout
        SetupUserGrid(sender)
    End Sub

    Private Sub ugAssignedGroups_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugAssignedGroups.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            btnGroupShiftLeft_Click(sender, e)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ugAvailGroups_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugAvailGroups.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            btnGroupShiftRight_Click(sender, e)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ugAvailUsers_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugAvailUsers.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            btnUserShiftRight_Click(sender, e)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ugManagedUsers_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugManagedUsers.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            btnUserShiftRight_Click(sender, e)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#Region "External Event Handlers"

    Private Sub oUser_UserExists(ByVal MsgStr As String) Handles oUser.UserExists
        'Don't know what this is for!!
    End Sub

    Private Sub oChangePassword_PasswordChanged() Handles oChangePassword.PasswordChanged
        Me.lblpsw.Visible = True
    End Sub

    'Private Sub oUser_MembershipsChanged(ByVal bolValue As Boolean) Handles oUser.MembershipsChanged
    '    If bolLoading Then Exit Sub
    '    SetSaveCancel(bolValue)
    'End Sub

    Private Sub oUser_UserChanged(ByVal bolValue As Boolean) Handles oUser.UserChanged
        If bolLoading Then Exit Sub
        SetSaveCancel(bolValue)
    End Sub

    Private Sub oUser_UsersChanged(ByVal bolValue As Boolean) Handles oUser.UsersChanged
        If bolLoading Then Exit Sub
        SetSaveCancel(bolValue)
    End Sub
#End Region

End Class
