Public Class ManageProfileData
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
    '  1.1        JC      01/04/05    Added code to handle MDI Child integration.
    '  1.2        JVC2    02/08/2005  Changed ProfileDatum References
    '  1.3        AN      02/10/05    Integrated AppFlags new object model
    '-------------------------------------------------------------------------------
    '
    'TODO - Remove comment from VSS version 2/9/05 - JVC 2
    '
    Private ProfileData As New Muster.BusinessLogic.pProfile
    Friend MyGUID As New System.Guid
    Dim returnVal As String = String.Empty

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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents DATE_LAST_EDITED As System.Windows.Forms.Label
    Friend WithEvents LAST_EDITED_BY As System.Windows.Forms.Label
    Friend WithEvents DATE_CREATED As System.Windows.Forms.Label
    Friend WithEvents CREATED_BY As System.Windows.Forms.Label
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox2 As System.Windows.Forms.ComboBox
    Friend WithEvents profileinfo As Infragistics.Win.UltraWinGrid.UltraGrid
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.DATE_LAST_EDITED = New System.Windows.Forms.Label
        Me.LAST_EDITED_BY = New System.Windows.Forms.Label
        Me.DATE_CREATED = New System.Windows.Forms.Label
        Me.CREATED_BY = New System.Windows.Forms.Label
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.ComboBox2 = New System.Windows.Forms.ComboBox
        Me.profileinfo = New Infragistics.Win.UltraWinGrid.UltraGrid
        CType(Me.profileinfo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Profile User ID"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(16, 42)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 16)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "Profile Key"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(16, 232)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(100, 16)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Created By"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(16, 256)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(100, 16)
        Me.Label4.TabIndex = 3
        Me.Label4.Text = "Created Date"
        '
        'Label5
        '
        Me.Label5.Location = New System.Drawing.Point(16, 280)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(100, 16)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "Modified By"
        '
        'Label6
        '
        Me.Label6.Location = New System.Drawing.Point(16, 304)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(100, 16)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "Modified Date"
        '
        'DATE_LAST_EDITED
        '
        Me.DATE_LAST_EDITED.Location = New System.Drawing.Point(128, 304)
        Me.DATE_LAST_EDITED.Name = "DATE_LAST_EDITED"
        Me.DATE_LAST_EDITED.Size = New System.Drawing.Size(392, 16)
        Me.DATE_LAST_EDITED.TabIndex = 9
        '
        'LAST_EDITED_BY
        '
        Me.LAST_EDITED_BY.Location = New System.Drawing.Point(128, 280)
        Me.LAST_EDITED_BY.Name = "LAST_EDITED_BY"
        Me.LAST_EDITED_BY.Size = New System.Drawing.Size(392, 16)
        Me.LAST_EDITED_BY.TabIndex = 8
        '
        'DATE_CREATED
        '
        Me.DATE_CREATED.Location = New System.Drawing.Point(128, 256)
        Me.DATE_CREATED.Name = "DATE_CREATED"
        Me.DATE_CREATED.Size = New System.Drawing.Size(392, 16)
        Me.DATE_CREATED.TabIndex = 7
        '
        'CREATED_BY
        '
        Me.CREATED_BY.Location = New System.Drawing.Point(128, 232)
        Me.CREATED_BY.Name = "CREATED_BY"
        Me.CREATED_BY.Size = New System.Drawing.Size(392, 16)
        Me.CREATED_BY.TabIndex = 6
        '
        'ComboBox1
        '
        Me.ComboBox1.Location = New System.Drawing.Point(128, 14)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(400, 21)
        Me.ComboBox1.TabIndex = 10
        '
        'ComboBox2
        '
        Me.ComboBox2.Location = New System.Drawing.Point(128, 40)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(400, 21)
        Me.ComboBox2.TabIndex = 11
        '
        'profileinfo
        '
        Me.profileinfo.Cursor = System.Windows.Forms.Cursors.Default
        Me.profileinfo.DisplayLayout.AddNewBox.Hidden = False
        Me.profileinfo.Location = New System.Drawing.Point(16, 64)
        Me.profileinfo.Name = "profileinfo"
        Me.profileinfo.Size = New System.Drawing.Size(512, 160)
        Me.profileinfo.TabIndex = 12
        '
        'ManageProfileData
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(536, 326)
        Me.Controls.Add(Me.profileinfo)
        Me.Controls.Add(Me.ComboBox2)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.DATE_LAST_EDITED)
        Me.Controls.Add(Me.LAST_EDITED_BY)
        Me.Controls.Add(Me.DATE_CREATED)
        Me.Controls.Add(Me.CREATED_BY)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "ManageProfileData"
        Me.Text = "Profile Data Internal Maint."
        CType(Me.profileinfo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub ProfileDatumMaint_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ProfileData.GetAll()
        ComboBox2.Enabled = False
        ComboBox1.DisplayMember = "USER_ID"
        ComboBox1.ValueMember = "USER_ID"
        ComboBox1.DataSource = ProfileData.ProfileUserTable()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.SelectedIndex > -1 And ComboBox1.SelectedValue <> "" Then
            ComboBox2.Enabled = True
            ComboBox2.DisplayMember = "PROFILE_KEY"
            ComboBox2.ValueMember = "PROFILE_KEY"
            ComboBox2.SelectedIndex = -1
            ComboBox2.DataSource = ProfileData.ProfileKeyTable()
        End If
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.SelectedIndex > -1 And ComboBox1.SelectedIndex > -1 And ComboBox1.SelectedValue <> "" Then
            profileinfo.DataSource = ProfileData.ProfileKeyValuesTable(ComboBox1.SelectedValue & "|" & ComboBox2.SelectedValue)
            profileinfo.DataBind()
            profileinfo.DisplayLayout.Bands(0).Columns("PROFILE_MODIFIER_1").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            profileinfo.DisplayLayout.Bands(0).Columns("PROFILE_MODIFIER_2").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
        End If
    End Sub

    Private Sub profileinfo_AfterRowActivate(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles profileinfo.AfterRowActivate
        If ComboBox2.SelectedIndex > -1 And ComboBox1.SelectedIndex > -1 And ComboBox1.SelectedValue <> "" And ComboBox2.SelectedValue <> "" Then
            'MsgBox(profileinfo.ActiveRow.Cells.Item("PROFILE_MODIFIER_1").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw))
            ProfileData.Retrieve(ComboBox1.SelectedValue & "|" & ComboBox2.SelectedValue & "|" & profileinfo.ActiveRow.Cells.Item("PROFILE_MODIFIER_1").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw) & "|" & profileinfo.ActiveRow.Cells.Item("PROFILE_MODIFIER_2").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw))
            CREATED_BY.Text = ProfileData.CreatedBy
            DATE_CREATED.Text = ProfileData.CreatedOn
            LAST_EDITED_BY.Text = ProfileData.ModifiedBy
            DATE_LAST_EDITED.Text = ProfileData.ModifiedOn
        End If
    End Sub

    Private Sub profileinfo_AfterCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles profileinfo.AfterCellUpdate
        Dim cellType As System.Type
        Dim strvalue As String
        cellType = e.Cell.Value.GetType

        If e.Cell.Column.ToString.IndexOf("PROFILE_VALUE") >= 0 Then
            ProfileData.Retrieve(ComboBox1.SelectedValue & "|" & ComboBox2.SelectedValue & "|" & profileinfo.ActiveRow.Cells.Item("PROFILE_MODIFIER_1").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw) & "|" & profileinfo.ActiveRow.Cells.Item("PROFILE_MODIFIER_2").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw))
            ProfileData.ProfileValue = e.Cell.Value
            If ProfileData.User = String.Empty Then
                ProfileData.CreatedBy = MusterContainer.AppUser.ID
            Else
                ProfileData.ModifiedBy = MusterContainer.AppUser.ID
            End If
            ProfileData.Save(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If
        ElseIf e.Cell.Column.ToString.IndexOf("DELETED") >= 0 Then
            If e.Cell.Value = True Then
                ProfileData.Deleted = e.Cell.Value
                If ProfileData.User = String.Empty Then
                    ProfileData.CreatedBy = MusterContainer.AppUser.ID
                Else
                    ProfileData.ModifiedBy = MusterContainer.AppUser.ID
                End If
                ProfileData.Save(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If
            End If
        End If

    End Sub
    'Private Sub profileinfo_BeforeCellUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeCellUpdateEventArgs) Handles profileinfo.BeforeCellUpdate
    'REMOVED - Code DID control which fields were updated. This has been done within Infragistics grid.
    '    Dim cellType As System.Type
    '    Dim strvalue As String
    '    cellType = e.Cell.Value.GetType

    '    If e.Cell.Column.ToString.IndexOf("PROFILE_VALUE") >= 0 Or e.Cell.Column.ToString.IndexOf("DELETED") >= 0 Then
    '        'do nothing
    '    Else
    '        If Not e.Cell.Row.IsAddRow Then
    '            e.Cancel = True
    '        End If
    '    End If
    'End Sub

    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)

        '
        ' Remove any values from the shared collection for this screen
        '
        MusterContainer.AppSemaphores.Remove(MyGUID.ToString)
        '
        ' Log the disposal of the form (exit from Registration form)
        '
        MusterContainer.AppUser.LogExit(MyGUID.ToString)

    End Sub


End Class
