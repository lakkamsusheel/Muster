Public Class SetUpOwner
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal powner As BusinessLogic.pOwner, ByVal moduleID As Integer, ByVal frm As Form, ByVal isnewOwner As Boolean)
        MyBase.New()


        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        pOwn = powner
        nModuleID = moduleID
        frmForm = frm
        bolIsNewOwner = isnewOwner

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
    Public WithEvents pnlOwnerName As System.Windows.Forms.Panel
    Friend WithEvents btnOwnerNameClose As System.Windows.Forms.Button
    Friend WithEvents btnOwnerNameOK As System.Windows.Forms.Button
    Public WithEvents pnlOwnerOrg As System.Windows.Forms.Panel
    Public WithEvents txtOwnerOrgName As System.Windows.Forms.TextBox
    Friend WithEvents lblOwnerOrgName As System.Windows.Forms.Label
    Public WithEvents cmbOwnerOrgEntityCode As System.Windows.Forms.ComboBox
    Friend WithEvents lblOwnerOrgEntityCode As System.Windows.Forms.Label
    Public WithEvents pnlPersonOrganization As System.Windows.Forms.Panel
    Public WithEvents rdOwnerOrg As System.Windows.Forms.RadioButton
    Public WithEvents rdOwnerPerson As System.Windows.Forms.RadioButton
    Public WithEvents pnlOwnerPerson As System.Windows.Forms.Panel
    Public WithEvents cmbOwnerNameSuffix As System.Windows.Forms.ComboBox
    Public WithEvents cmbOwnerNameTitle As System.Windows.Forms.ComboBox
    Friend WithEvents lblOwnerNameSuffix As System.Windows.Forms.Label
    Friend WithEvents lblOwnerNameTitle As System.Windows.Forms.Label
    Public WithEvents txtOwnerMiddleName As System.Windows.Forms.TextBox
    Friend WithEvents lblOwnerMiddleName As System.Windows.Forms.Label
    Public WithEvents txtOwnerLastName As System.Windows.Forms.TextBox
    Public WithEvents txtOwnerFirstName As System.Windows.Forms.TextBox
    Friend WithEvents lblOwnerLastName As System.Windows.Forms.Label
    Friend WithEvents lblOwnerFirstName As System.Windows.Forms.Label
    Public WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Public WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents cmbOwnerEntities As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlOwnerName = New System.Windows.Forms.Panel
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.cmbOwnerEntities = New System.Windows.Forms.ComboBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.pnlOwnerOrg = New System.Windows.Forms.Panel
        Me.txtOwnerOrgName = New System.Windows.Forms.TextBox
        Me.lblOwnerOrgName = New System.Windows.Forms.Label
        Me.cmbOwnerOrgEntityCode = New System.Windows.Forms.ComboBox
        Me.lblOwnerOrgEntityCode = New System.Windows.Forms.Label
        Me.pnlPersonOrganization = New System.Windows.Forms.Panel
        Me.RadioButton1 = New System.Windows.Forms.RadioButton
        Me.rdOwnerOrg = New System.Windows.Forms.RadioButton
        Me.rdOwnerPerson = New System.Windows.Forms.RadioButton
        Me.pnlOwnerPerson = New System.Windows.Forms.Panel
        Me.cmbOwnerNameSuffix = New System.Windows.Forms.ComboBox
        Me.cmbOwnerNameTitle = New System.Windows.Forms.ComboBox
        Me.lblOwnerNameSuffix = New System.Windows.Forms.Label
        Me.lblOwnerNameTitle = New System.Windows.Forms.Label
        Me.txtOwnerMiddleName = New System.Windows.Forms.TextBox
        Me.lblOwnerMiddleName = New System.Windows.Forms.Label
        Me.txtOwnerLastName = New System.Windows.Forms.TextBox
        Me.txtOwnerFirstName = New System.Windows.Forms.TextBox
        Me.lblOwnerLastName = New System.Windows.Forms.Label
        Me.lblOwnerFirstName = New System.Windows.Forms.Label
        Me.btnOwnerNameClose = New System.Windows.Forms.Button
        Me.btnOwnerNameOK = New System.Windows.Forms.Button
        Me.pnlOwnerName.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.pnlOwnerOrg.SuspendLayout()
        Me.pnlPersonOrganization.SuspendLayout()
        Me.pnlOwnerPerson.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlOwnerName
        '
        Me.pnlOwnerName.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.pnlOwnerName.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pnlOwnerName.Controls.Add(Me.Panel1)
        Me.pnlOwnerName.Controls.Add(Me.pnlOwnerOrg)
        Me.pnlOwnerName.Controls.Add(Me.pnlPersonOrganization)
        Me.pnlOwnerName.Controls.Add(Me.pnlOwnerPerson)
        Me.pnlOwnerName.Location = New System.Drawing.Point(0, 8)
        Me.pnlOwnerName.Name = "pnlOwnerName"
        Me.pnlOwnerName.Size = New System.Drawing.Size(320, 240)
        Me.pnlOwnerName.TabIndex = 3
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.cmbOwnerEntities)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Location = New System.Drawing.Point(328, 40)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(112, 192)
        Me.Panel1.TabIndex = 3
        Me.Panel1.Visible = False
        '
        'cmbOwnerEntities
        '
        Me.cmbOwnerEntities.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbOwnerEntities.DropDownWidth = 200
        Me.cmbOwnerEntities.ItemHeight = 13
        Me.cmbOwnerEntities.Items.AddRange(New Object() {""})
        Me.cmbOwnerEntities.Location = New System.Drawing.Point(48, 8)
        Me.cmbOwnerEntities.Name = "cmbOwnerEntities"
        Me.cmbOwnerEntities.Size = New System.Drawing.Size(40, 21)
        Me.cmbOwnerEntities.TabIndex = 89
        Me.cmbOwnerEntities.Visible = False
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 7)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 23)
        Me.Label1.TabIndex = 88
        Me.Label1.Text = "Name"
        Me.Label1.Visible = False
        '
        'pnlOwnerOrg
        '
        Me.pnlOwnerOrg.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.pnlOwnerOrg.Controls.Add(Me.txtOwnerOrgName)
        Me.pnlOwnerOrg.Controls.Add(Me.lblOwnerOrgName)
        Me.pnlOwnerOrg.Controls.Add(Me.cmbOwnerOrgEntityCode)
        Me.pnlOwnerOrg.Controls.Add(Me.lblOwnerOrgEntityCode)
        Me.pnlOwnerOrg.Location = New System.Drawing.Point(0, 176)
        Me.pnlOwnerOrg.Name = "pnlOwnerOrg"
        Me.pnlOwnerOrg.Size = New System.Drawing.Size(304, 56)
        Me.pnlOwnerOrg.TabIndex = 2
        '
        'txtOwnerOrgName
        '
        Me.txtOwnerOrgName.Location = New System.Drawing.Point(97, 7)
        Me.txtOwnerOrgName.Name = "txtOwnerOrgName"
        Me.txtOwnerOrgName.Size = New System.Drawing.Size(192, 20)
        Me.txtOwnerOrgName.TabIndex = 0
        Me.txtOwnerOrgName.Tag = ""
        Me.txtOwnerOrgName.Text = ""
        '
        'lblOwnerOrgName
        '
        Me.lblOwnerOrgName.Location = New System.Drawing.Point(8, 7)
        Me.lblOwnerOrgName.Name = "lblOwnerOrgName"
        Me.lblOwnerOrgName.TabIndex = 88
        Me.lblOwnerOrgName.Text = "Name"
        '
        'cmbOwnerOrgEntityCode
        '
        Me.cmbOwnerOrgEntityCode.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbOwnerOrgEntityCode.DropDownWidth = 200
        Me.cmbOwnerOrgEntityCode.ItemHeight = 13
        Me.cmbOwnerOrgEntityCode.Items.AddRange(New Object() {"Owner Type 1", "Owner Type 2"})
        Me.cmbOwnerOrgEntityCode.Location = New System.Drawing.Point(96, 32)
        Me.cmbOwnerOrgEntityCode.Name = "cmbOwnerOrgEntityCode"
        Me.cmbOwnerOrgEntityCode.Size = New System.Drawing.Size(152, 21)
        Me.cmbOwnerOrgEntityCode.TabIndex = 1
        '
        'lblOwnerOrgEntityCode
        '
        Me.lblOwnerOrgEntityCode.Location = New System.Drawing.Point(8, 32)
        Me.lblOwnerOrgEntityCode.Name = "lblOwnerOrgEntityCode"
        Me.lblOwnerOrgEntityCode.Size = New System.Drawing.Size(72, 23)
        Me.lblOwnerOrgEntityCode.TabIndex = 91
        Me.lblOwnerOrgEntityCode.Text = "Entity Code"
        '
        'pnlPersonOrganization
        '
        Me.pnlPersonOrganization.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.pnlPersonOrganization.Controls.Add(Me.RadioButton1)
        Me.pnlPersonOrganization.Controls.Add(Me.rdOwnerOrg)
        Me.pnlPersonOrganization.Controls.Add(Me.rdOwnerPerson)
        Me.pnlPersonOrganization.Location = New System.Drawing.Point(0, 0)
        Me.pnlPersonOrganization.Name = "pnlPersonOrganization"
        Me.pnlPersonOrganization.Size = New System.Drawing.Size(792, 32)
        Me.pnlPersonOrganization.TabIndex = 0
        '
        'RadioButton1
        '
        Me.RadioButton1.Location = New System.Drawing.Point(336, 8)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.RadioButton1.Size = New System.Drawing.Size(72, 16)
        Me.RadioButton1.TabIndex = 2
        Me.RadioButton1.Text = "Search Owner Entity"
        Me.RadioButton1.TextAlign = System.Drawing.ContentAlignment.TopRight
        Me.RadioButton1.Visible = False
        '
        'rdOwnerOrg
        '
        Me.rdOwnerOrg.Location = New System.Drawing.Point(176, 8)
        Me.rdOwnerOrg.Name = "rdOwnerOrg"
        Me.rdOwnerOrg.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.rdOwnerOrg.Size = New System.Drawing.Size(93, 16)
        Me.rdOwnerOrg.TabIndex = 1
        Me.rdOwnerOrg.Text = "Organization"
        Me.rdOwnerOrg.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'rdOwnerPerson
        '
        Me.rdOwnerPerson.Checked = True
        Me.rdOwnerPerson.Location = New System.Drawing.Point(16, 8)
        Me.rdOwnerPerson.Name = "rdOwnerPerson"
        Me.rdOwnerPerson.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.rdOwnerPerson.Size = New System.Drawing.Size(64, 16)
        Me.rdOwnerPerson.TabIndex = 0
        Me.rdOwnerPerson.TabStop = True
        Me.rdOwnerPerson.Text = "Person"
        Me.rdOwnerPerson.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'pnlOwnerPerson
        '
        Me.pnlOwnerPerson.BackColor = System.Drawing.SystemColors.ControlLightLight
        Me.pnlOwnerPerson.Controls.Add(Me.cmbOwnerNameSuffix)
        Me.pnlOwnerPerson.Controls.Add(Me.cmbOwnerNameTitle)
        Me.pnlOwnerPerson.Controls.Add(Me.lblOwnerNameSuffix)
        Me.pnlOwnerPerson.Controls.Add(Me.lblOwnerNameTitle)
        Me.pnlOwnerPerson.Controls.Add(Me.txtOwnerMiddleName)
        Me.pnlOwnerPerson.Controls.Add(Me.lblOwnerMiddleName)
        Me.pnlOwnerPerson.Controls.Add(Me.txtOwnerLastName)
        Me.pnlOwnerPerson.Controls.Add(Me.txtOwnerFirstName)
        Me.pnlOwnerPerson.Controls.Add(Me.lblOwnerLastName)
        Me.pnlOwnerPerson.Controls.Add(Me.lblOwnerFirstName)
        Me.pnlOwnerPerson.Location = New System.Drawing.Point(0, 40)
        Me.pnlOwnerPerson.Name = "pnlOwnerPerson"
        Me.pnlOwnerPerson.Size = New System.Drawing.Size(304, 128)
        Me.pnlOwnerPerson.TabIndex = 1
        '
        'cmbOwnerNameSuffix
        '
        Me.cmbOwnerNameSuffix.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbOwnerNameSuffix.Items.AddRange(New Object() {"Jr", "Sr", "I", "II", "III", "IV", "V", "VI"})
        Me.cmbOwnerNameSuffix.Location = New System.Drawing.Point(96, 103)
        Me.cmbOwnerNameSuffix.Name = "cmbOwnerNameSuffix"
        Me.cmbOwnerNameSuffix.Size = New System.Drawing.Size(80, 21)
        Me.cmbOwnerNameSuffix.TabIndex = 4
        '
        'cmbOwnerNameTitle
        '
        Me.cmbOwnerNameTitle.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbOwnerNameTitle.Items.AddRange(New Object() {"Mr", "Mrs", "Ms", "Dr", "Sir"})
        Me.cmbOwnerNameTitle.Location = New System.Drawing.Point(96, 8)
        Me.cmbOwnerNameTitle.Name = "cmbOwnerNameTitle"
        Me.cmbOwnerNameTitle.Size = New System.Drawing.Size(80, 21)
        Me.cmbOwnerNameTitle.TabIndex = 0
        '
        'lblOwnerNameSuffix
        '
        Me.lblOwnerNameSuffix.Location = New System.Drawing.Point(8, 103)
        Me.lblOwnerNameSuffix.Name = "lblOwnerNameSuffix"
        Me.lblOwnerNameSuffix.TabIndex = 92
        Me.lblOwnerNameSuffix.Text = "Suffix"
        '
        'lblOwnerNameTitle
        '
        Me.lblOwnerNameTitle.Location = New System.Drawing.Point(8, 8)
        Me.lblOwnerNameTitle.Name = "lblOwnerNameTitle"
        Me.lblOwnerNameTitle.TabIndex = 91
        Me.lblOwnerNameTitle.Text = "Title"
        '
        'txtOwnerMiddleName
        '
        Me.txtOwnerMiddleName.Location = New System.Drawing.Point(96, 56)
        Me.txtOwnerMiddleName.Name = "txtOwnerMiddleName"
        Me.txtOwnerMiddleName.Size = New System.Drawing.Size(192, 20)
        Me.txtOwnerMiddleName.TabIndex = 2
        Me.txtOwnerMiddleName.Tag = ""
        Me.txtOwnerMiddleName.Text = ""
        '
        'lblOwnerMiddleName
        '
        Me.lblOwnerMiddleName.Location = New System.Drawing.Point(8, 56)
        Me.lblOwnerMiddleName.Name = "lblOwnerMiddleName"
        Me.lblOwnerMiddleName.Size = New System.Drawing.Size(100, 17)
        Me.lblOwnerMiddleName.TabIndex = 89
        Me.lblOwnerMiddleName.Text = "Middle Name"
        '
        'txtOwnerLastName
        '
        Me.txtOwnerLastName.Location = New System.Drawing.Point(96, 80)
        Me.txtOwnerLastName.Name = "txtOwnerLastName"
        Me.txtOwnerLastName.Size = New System.Drawing.Size(192, 20)
        Me.txtOwnerLastName.TabIndex = 3
        Me.txtOwnerLastName.Tag = ""
        Me.txtOwnerLastName.Text = ""
        '
        'txtOwnerFirstName
        '
        Me.txtOwnerFirstName.Location = New System.Drawing.Point(96, 32)
        Me.txtOwnerFirstName.Name = "txtOwnerFirstName"
        Me.txtOwnerFirstName.Size = New System.Drawing.Size(192, 20)
        Me.txtOwnerFirstName.TabIndex = 1
        Me.txtOwnerFirstName.Tag = ""
        Me.txtOwnerFirstName.Text = ""
        '
        'lblOwnerLastName
        '
        Me.lblOwnerLastName.Location = New System.Drawing.Point(8, 80)
        Me.lblOwnerLastName.Name = "lblOwnerLastName"
        Me.lblOwnerLastName.TabIndex = 86
        Me.lblOwnerLastName.Text = "Last Name"
        '
        'lblOwnerFirstName
        '
        Me.lblOwnerFirstName.Location = New System.Drawing.Point(8, 32)
        Me.lblOwnerFirstName.Name = "lblOwnerFirstName"
        Me.lblOwnerFirstName.TabIndex = 85
        Me.lblOwnerFirstName.Text = "First Name"
        '
        'btnOwnerNameClose
        '
        Me.btnOwnerNameClose.BackColor = System.Drawing.SystemColors.Control
        Me.btnOwnerNameClose.Location = New System.Drawing.Point(104, 256)
        Me.btnOwnerNameClose.Name = "btnOwnerNameClose"
        Me.btnOwnerNameClose.Size = New System.Drawing.Size(60, 24)
        Me.btnOwnerNameClose.TabIndex = 2
        Me.btnOwnerNameClose.Text = "Close"
        '
        'btnOwnerNameOK
        '
        Me.btnOwnerNameOK.BackColor = System.Drawing.SystemColors.Control
        Me.btnOwnerNameOK.Enabled = False
        Me.btnOwnerNameOK.Location = New System.Drawing.Point(0, 256)
        Me.btnOwnerNameOK.Name = "btnOwnerNameOK"
        Me.btnOwnerNameOK.Size = New System.Drawing.Size(56, 24)
        Me.btnOwnerNameOK.TabIndex = 0
        Me.btnOwnerNameOK.Text = "Save"
        '
        'SetUpOwner
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(336, 285)
        Me.Controls.Add(Me.pnlOwnerName)
        Me.Controls.Add(Me.btnOwnerNameOK)
        Me.Controls.Add(Me.btnOwnerNameClose)
        Me.Name = "SetUpOwner"
        Me.Text = "Setting Up Owner Name"
        Me.pnlOwnerName.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.pnlOwnerOrg.ResumeLayout(False)
        Me.pnlPersonOrganization.ResumeLayout(False)
        Me.pnlOwnerPerson.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "private members"
    Private nPerson_ID As Integer = 0
    Private nOrg_ID As Integer = 0
    Private WithEvents pOwn As BusinessLogic.pOwner
    Private nModuleID As Integer
    Private frmForm As Form
    Public bolLoading As Boolean = False
    Public bolIsNewOwner As Boolean = False

#End Region

#Region "Public Members"
    Public ReadOnly Property Person_ID() As Integer
        Get
            Return nperson_id
        End Get
    End Property

    Public ReadOnly Property Org_ID() As Integer
        Get
            Return nOrg_ID
        End Get
    End Property
#End Region


#Region "Methods"

    Private Sub SetPersonaSaveCancel(ByVal bolstate As Boolean)
        btnOwnerNameOK.Enabled = bolstate
    End Sub


    Private Sub ShowError(ByVal ex As Exception)
        Dim MyErr As New ErrorReport(ex)
        MyErr.ShowDialog()
    End Sub

    Public Sub LoadComboBoxes()

        If Not bolIsNewOwner Then
            Me.RadioButton1.Enabled = False
            Me.rdOwnerOrg.Enabled = False
            Me.rdOwnerPerson.Enabled = False
            ' UIUtilsGen.ResetOwnerName(Me, pOwn, False)

        Else
            Me.RadioButton1.Enabled = True
            Me.rdOwnerOrg.Enabled = True
            Me.rdOwnerPerson.Enabled = True

            UIUtilsGen.ClearPersona(Me)

        End If

        UIUtilsGen.PopulateOrgEntityType(Me.cmbOwnerOrgEntityCode, pOwn)
        UIUtilsGen.PopulateOwnerEntityList(Me.cmbOwnerEntities, pOwn, Not bolIsNewOwner)

    End Sub

#End Region

#Region "Events"

    Private Sub btnOwnerNameClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOwnerNameClose.Click
        Me.Close()

    End Sub

    Private Sub SetUpOwner_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadComboBoxes()

        UIUtilsGen.SwapOrgPersonDisplay(Me)

        btnOwnerNameOK.Enabled = False
    End Sub

    Private Sub txtOwnerFirstName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOwnerFirstName.TextChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillPersona(txtOwnerFirstName, pOwn)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub txtOwnerMiddleName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOwnerMiddleName.TextChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillPersona(txtOwnerMiddleName, pOwn)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub txtOwnerLastName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOwnerLastName.TextChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillPersona(txtOwnerLastName, pOwn)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub cmbOwnerNameSuffix_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbOwnerNameSuffix.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillPersona(cmbOwnerNameSuffix, pOwn)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub txtOwnerOrgName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtOwnerOrgName.TextChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillPersona(txtOwnerOrgName, pOwn)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub cmbOwnerOrgEntityCode_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbOwnerOrgEntityCode.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillPersona(cmbOwnerOrgEntityCode, pOwn)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub cmbEntityCode_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbOwnerEntities.SelectedIndexChanged
        If bolLoading Then Exit Sub
        If Me.cmbOwnerEntities.SelectedIndex <= 0 Then
            SetPersonaSaveCancel(False)
            Exit Sub
        End If

        Try
            UIUtilsGen.FillPersona(cmbOwnerEntities, pOwn)
            '   UIUtilsGen.ResetOwnerName(Me, pOwn, True)

        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub pnlPersonOrganization_LostFocus(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pnlPersonOrganization.LostFocus
        Try
            cmbOwnerNameTitle.Focus()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub rdOwnerPerson_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdOwnerPerson.Click
        Try
            ' UIUtilsGen.rdOwnerPersonClick(Me, pOwn, frmForm)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub rdOwnerOrg_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdOwnerOrg.Click
        Try
            ' UIUtilsGen.rdOwnerOrgClick(Me, pOwn, frmForm)
            If rdOwnerOrg.Checked And pOwn.BPersona.Org_Entity_Code = 0 Then
                cmbOwnerOrgEntityCode.DisplayMember = "PROPERTY_NAME"
                cmbOwnerOrgEntityCode.ValueMember = "PROPERTY_ID"
                cmbOwnerOrgEntityCode.SelectedValue = 539  ' Default "UnderGround Storage Tank Owner"

            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub rdListOrg_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.Click
        Try
            '  UIUtilsGen.rdOwnerEntityListClick(Me, pOwn)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub
    Private Sub cmbOwnerNameTitle_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbOwnerNameTitle.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.FillPersona(cmbOwnerNameTitle, pOwn)
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub btnOwnerNameOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOwnerNameOK.Click
        Try
            Dim success As Boolean = False
            Dim returnVal As String = String.Empty

            If bolIsNewOwner Then
                If nPerson_ID > 0 Then
                    pOwn.BPersona.PersonId = nPerson_ID
                Else
                    pOwn.BPersona.OrgID = nOrg_ID
                End If
            End If

            If pOwn.BPersona.PersonId > 0 Or pOwn.BPersona.OrgID > 0 Then
                pOwn.BPersona.ModifiedBy = MusterContainer.AppUser.ID
            Else
                pOwn.BPersona.CreatedBy = MusterContainer.AppUser.ID
            End If

            success = pOwn.BPersona.Save(nModuleID, MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            If success Then
                '  UIUtilsGen.SetOwnerName(Me, frmForm, pOwn)
            Else
                Exit Sub
            End If

            CType(frmForm, Object).txtOwnerAddress.Focus()
            Close()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Sub PersonaChanged(ByVal bolState As Boolean) Handles pOwn.evtPersonaChanged
        If Not bolLoading Then
            SetPersonaSaveCancel(bolState)
        End If

    End Sub



#End Region


End Class
