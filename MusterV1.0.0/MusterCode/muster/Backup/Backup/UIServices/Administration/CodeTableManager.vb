Public Class CodeTableManager
    Inherits System.Windows.Forms.Form
    '-------------------------------------------------------------------------------
    ' MUSTER.MUSTER.CodeTableManager
    '   Form used to manage system property tables.
    '
    ' Copyright (C) 2004 CIBER, Inc.
    ' All rights reserved.
    '
    ' Release   Initials    Date      Description
    '  1.0        ??      ??/??/??    Original class definition.
    '  1.1        JC      01/03/05    Modified to incorporate window activation code
    '                                   for MDI child
    '  1.2        AN      02/10/05    Integrated AppFlags new object model
    '  1.3        JVCII   02/14/05    Prevented addition of rows to parent grid.
    '  1.4        AN      02/25/05    Added ADD for parent grid
    '-------------------------------------------------------------------------------
    '
    'TODO - Remove comment from VSS version 2/9/05 - JVC 2
    '
#Region "Private Member Variables"
    Private WithEvents oPropType As New Muster.BusinessLogic.pPropertyType 'InfoRepository.MusterProperties
    'Private oPropertyType As New Muster.BusinessLogic.pPropertyType 'InfoRepository.MusterPropertyType
    Private bolCancel As Boolean = False
    Private bolLoadingCombo As Boolean = False
    Private bolAccFlag As Boolean = True
    Private dtProperties As DataTable
    Private dtAssociatedProperties As DataTable
    Private nPropertyActiveIndex As Integer = -1
    Private nAssociatedPropertyActiveIndex As Integer = -1
    Private ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
    Friend MyGUID As New System.Guid
    Dim returnVal As String = String.Empty
#End Region
#Region " Windows Form Designer generated code "

    Public Sub New(Optional ByRef frm As Windows.Forms.Form = Nothing)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        '
        'Load the entity combo box 
        '
        Me.PopulateEntity()
        '
        'Force load the property types box
        '
        If cmbEntity.Items.Count = 0 Then
            MsgBox("Form did not initialize properly - shutting down form.", MsgBoxStyle.Critical & MsgBoxStyle.OKOnly, "Error Loading Form!")
        Else
            cmbEntity.SelectedIndex = 0
            '
            'Force load the properties grid
            '
            If cboPropertyType.Items.Count > 0 Then
                cboPropertyType.SelectedIndex = 0
            End If

            MyGUID = System.Guid.NewGuid


            MusterContainer.AppUser.LogEntry(Me.Text, MyGUID.ToString)

            '2/10/2005 - AN - Changed AppFlags to New Object
            'MusterContainer.AppSemaphores.PutValuePair(MyGUID.ToString, "WindowName", Me.Text)
            'MusterContainer.AppSemaphores.PutValuePair("0", "ActiveForm", MyGUID)
            MusterContainer.AppSemaphores.Retrieve(MyGUID.ToString, "WindowName", Me.Text)
            MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGUID)

            If Not frm Is Nothing Then
                If frm.IsMdiContainer Then
                    Me.MdiParent = frm
                End If
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
    Friend WithEvents lblEntity As System.Windows.Forms.Label
    Friend WithEvents lblPropertyType As System.Windows.Forms.Label
    Friend WithEvents ugProperties As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnAddProperty As System.Windows.Forms.Button
    Friend WithEvents lblAvailableAssociatedProperties As System.Windows.Forms.Label
    Friend WithEvents cboAAssociatedproperties As System.Windows.Forms.ComboBox
    Friend WithEvents ugAssociatedProperties As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnRemoveProperty As System.Windows.Forms.Button
    Friend WithEvents cboPropertyType As System.Windows.Forms.ComboBox
    Friend WithEvents cmbEntity As System.Windows.Forms.ComboBox
    Friend WithEvents btnSaveProperties As System.Windows.Forms.Button
    Friend WithEvents pnlAdminTop As System.Windows.Forms.Panel
    Friend WithEvents pnlAdminRight As System.Windows.Forms.Panel
    Friend WithEvents pnlPropertiesAdmin As System.Windows.Forms.Panel
    Friend WithEvents pnlAvailableProperties As System.Windows.Forms.Panel
    Friend WithEvents pnlAssociatedProperties As System.Windows.Forms.Panel
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnReset As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblEntity = New System.Windows.Forms.Label
        Me.lblPropertyType = New System.Windows.Forms.Label
        Me.ugProperties = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.cboAAssociatedproperties = New System.Windows.Forms.ComboBox
        Me.lblAvailableAssociatedProperties = New System.Windows.Forms.Label
        Me.btnAddProperty = New System.Windows.Forms.Button
        Me.ugAssociatedProperties = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.btnRemoveProperty = New System.Windows.Forms.Button
        Me.cboPropertyType = New System.Windows.Forms.ComboBox
        Me.cmbEntity = New System.Windows.Forms.ComboBox
        Me.btnSaveProperties = New System.Windows.Forms.Button
        Me.pnlAdminTop = New System.Windows.Forms.Panel
        Me.pnlAdminRight = New System.Windows.Forms.Panel
        Me.btnReset = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.pnlPropertiesAdmin = New System.Windows.Forms.Panel
        Me.pnlAvailableProperties = New System.Windows.Forms.Panel
        Me.pnlAssociatedProperties = New System.Windows.Forms.Panel
        CType(Me.ugProperties, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.ugAssociatedProperties, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlAdminTop.SuspendLayout()
        Me.pnlAdminRight.SuspendLayout()
        Me.pnlPropertiesAdmin.SuspendLayout()
        Me.pnlAvailableProperties.SuspendLayout()
        Me.pnlAssociatedProperties.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblEntity
        '
        Me.lblEntity.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEntity.Location = New System.Drawing.Point(16, 8)
        Me.lblEntity.Name = "lblEntity"
        Me.lblEntity.Size = New System.Drawing.Size(88, 20)
        Me.lblEntity.TabIndex = 183
        Me.lblEntity.Text = "Entity:"
        '
        'lblPropertyType
        '
        Me.lblPropertyType.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPropertyType.Location = New System.Drawing.Point(16, 40)
        Me.lblPropertyType.Name = "lblPropertyType"
        Me.lblPropertyType.Size = New System.Drawing.Size(88, 20)
        Me.lblPropertyType.TabIndex = 185
        Me.lblPropertyType.Text = "Property Type:"
        '
        'ugProperties
        '
        Me.ugProperties.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugProperties.DisplayLayout.AddNewBox.Hidden = False
        Me.ugProperties.DisplayLayout.AddNewBox.Prompt = "Add new property"
        Me.ugProperties.DisplayLayout.AddNewBox.Style = Infragistics.Win.UltraWinGrid.AddNewBoxStyle.Compact
        Me.ugProperties.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        Me.ugProperties.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugProperties.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugProperties.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ugProperties.Location = New System.Drawing.Point(20, 3)
        Me.ugProperties.Name = "ugProperties"
        Me.ugProperties.Size = New System.Drawing.Size(548, 205)
        Me.ugProperties.TabIndex = 0
        Me.ugProperties.Text = "Properties"
        '
        'cboAAssociatedproperties
        '
        Me.cboAAssociatedproperties.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAAssociatedproperties.DropDownWidth = 440
        Me.cboAAssociatedproperties.ItemHeight = 13
        Me.cboAAssociatedproperties.Location = New System.Drawing.Point(120, 16)
        Me.cboAAssociatedproperties.Name = "cboAAssociatedproperties"
        Me.cboAAssociatedproperties.Size = New System.Drawing.Size(336, 21)
        Me.cboAAssociatedproperties.TabIndex = 0
        '
        'lblAvailableAssociatedProperties
        '
        Me.lblAvailableAssociatedProperties.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAvailableAssociatedProperties.Location = New System.Drawing.Point(24, 8)
        Me.lblAvailableAssociatedProperties.Name = "lblAvailableAssociatedProperties"
        Me.lblAvailableAssociatedProperties.Size = New System.Drawing.Size(88, 48)
        Me.lblAvailableAssociatedProperties.TabIndex = 188
        Me.lblAvailableAssociatedProperties.Text = "Available Associated Properties"
        '
        'btnAddProperty
        '
        Me.btnAddProperty.BackColor = System.Drawing.SystemColors.Control
        Me.btnAddProperty.Enabled = False
        Me.btnAddProperty.Location = New System.Drawing.Point(464, 16)
        Me.btnAddProperty.Name = "btnAddProperty"
        Me.btnAddProperty.Size = New System.Drawing.Size(88, 24)
        Me.btnAddProperty.TabIndex = 1
        Me.btnAddProperty.Text = "Add Property"
        '
        'ugAssociatedProperties
        '
        Me.ugAssociatedProperties.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugAssociatedProperties.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugAssociatedProperties.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugAssociatedProperties.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ugAssociatedProperties.Location = New System.Drawing.Point(20, 0)
        Me.ugAssociatedProperties.Name = "ugAssociatedProperties"
        Me.ugAssociatedProperties.Size = New System.Drawing.Size(548, 176)
        Me.ugAssociatedProperties.TabIndex = 0
        Me.ugAssociatedProperties.Text = "Associated Properties"
        '
        'btnRemoveProperty
        '
        Me.btnRemoveProperty.BackColor = System.Drawing.SystemColors.Control
        Me.btnRemoveProperty.Enabled = False
        Me.btnRemoveProperty.Location = New System.Drawing.Point(16, 264)
        Me.btnRemoveProperty.Name = "btnRemoveProperty"
        Me.btnRemoveProperty.Size = New System.Drawing.Size(80, 48)
        Me.btnRemoveProperty.TabIndex = 3
        Me.btnRemoveProperty.Text = "Remove Associated Property"
        '
        'cboPropertyType
        '
        Me.cboPropertyType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboPropertyType.DropDownWidth = 300
        Me.cboPropertyType.ItemHeight = 13
        Me.cboPropertyType.Location = New System.Drawing.Point(112, 40)
        Me.cboPropertyType.MaxDropDownItems = 10
        Me.cboPropertyType.Name = "cboPropertyType"
        Me.cboPropertyType.Size = New System.Drawing.Size(280, 21)
        Me.cboPropertyType.TabIndex = 1
        '
        'cmbEntity
        '
        Me.cmbEntity.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbEntity.DropDownWidth = 250
        Me.cmbEntity.Location = New System.Drawing.Point(112, 8)
        Me.cmbEntity.Name = "cmbEntity"
        Me.cmbEntity.Size = New System.Drawing.Size(160, 21)
        Me.cmbEntity.TabIndex = 0
        '
        'btnSaveProperties
        '
        Me.btnSaveProperties.Enabled = False
        Me.btnSaveProperties.Location = New System.Drawing.Point(16, 16)
        Me.btnSaveProperties.Name = "btnSaveProperties"
        Me.btnSaveProperties.Size = New System.Drawing.Size(75, 48)
        Me.btnSaveProperties.TabIndex = 0
        Me.btnSaveProperties.Text = "Save Properties"
        '
        'pnlAdminTop
        '
        Me.pnlAdminTop.Controls.Add(Me.lblEntity)
        Me.pnlAdminTop.Controls.Add(Me.cmbEntity)
        Me.pnlAdminTop.Controls.Add(Me.lblPropertyType)
        Me.pnlAdminTop.Controls.Add(Me.cboPropertyType)
        Me.pnlAdminTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlAdminTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlAdminTop.Name = "pnlAdminTop"
        Me.pnlAdminTop.Size = New System.Drawing.Size(688, 72)
        Me.pnlAdminTop.TabIndex = 195
        '
        'pnlAdminRight
        '
        Me.pnlAdminRight.Controls.Add(Me.btnReset)
        Me.pnlAdminRight.Controls.Add(Me.btnClose)
        Me.pnlAdminRight.Controls.Add(Me.btnSaveProperties)
        Me.pnlAdminRight.Controls.Add(Me.btnRemoveProperty)
        Me.pnlAdminRight.Dock = System.Windows.Forms.DockStyle.Right
        Me.pnlAdminRight.Location = New System.Drawing.Point(568, 72)
        Me.pnlAdminRight.Name = "pnlAdminRight"
        Me.pnlAdminRight.Size = New System.Drawing.Size(120, 454)
        Me.pnlAdminRight.TabIndex = 196
        '
        'btnReset
        '
        Me.btnReset.Enabled = False
        Me.btnReset.Location = New System.Drawing.Point(16, 80)
        Me.btnReset.Name = "btnReset"
        Me.btnReset.Size = New System.Drawing.Size(75, 48)
        Me.btnReset.TabIndex = 1
        Me.btnReset.Text = "Reset"
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(16, 144)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(75, 48)
        Me.btnClose.TabIndex = 2
        Me.btnClose.Text = "Close"
        '
        'pnlPropertiesAdmin
        '
        Me.pnlPropertiesAdmin.Controls.Add(Me.ugProperties)
        Me.pnlPropertiesAdmin.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlPropertiesAdmin.DockPadding.Left = 20
        Me.pnlPropertiesAdmin.DockPadding.Top = 3
        Me.pnlPropertiesAdmin.Location = New System.Drawing.Point(0, 72)
        Me.pnlPropertiesAdmin.Name = "pnlPropertiesAdmin"
        Me.pnlPropertiesAdmin.Size = New System.Drawing.Size(568, 208)
        Me.pnlPropertiesAdmin.TabIndex = 197
        '
        'pnlAvailableProperties
        '
        Me.pnlAvailableProperties.Controls.Add(Me.lblAvailableAssociatedProperties)
        Me.pnlAvailableProperties.Controls.Add(Me.btnAddProperty)
        Me.pnlAvailableProperties.Controls.Add(Me.cboAAssociatedproperties)
        Me.pnlAvailableProperties.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlAvailableProperties.Location = New System.Drawing.Point(0, 280)
        Me.pnlAvailableProperties.Name = "pnlAvailableProperties"
        Me.pnlAvailableProperties.Size = New System.Drawing.Size(568, 56)
        Me.pnlAvailableProperties.TabIndex = 0
        '
        'pnlAssociatedProperties
        '
        Me.pnlAssociatedProperties.Controls.Add(Me.ugAssociatedProperties)
        Me.pnlAssociatedProperties.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlAssociatedProperties.DockPadding.Left = 20
        Me.pnlAssociatedProperties.Location = New System.Drawing.Point(0, 336)
        Me.pnlAssociatedProperties.Name = "pnlAssociatedProperties"
        Me.pnlAssociatedProperties.Size = New System.Drawing.Size(568, 176)
        Me.pnlAssociatedProperties.TabIndex = 199
        '
        'CodeTableManager
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(688, 526)
        Me.Controls.Add(Me.pnlAssociatedProperties)
        Me.Controls.Add(Me.pnlAvailableProperties)
        Me.Controls.Add(Me.pnlPropertiesAdmin)
        Me.Controls.Add(Me.pnlAdminRight)
        Me.Controls.Add(Me.pnlAdminTop)
        Me.Name = "CodeTableManager"
        Me.Text = "Manage Code Table Properties"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.ugProperties, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.ugAssociatedProperties, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlAdminTop.ResumeLayout(False)
        Me.pnlAdminRight.ResumeLayout(False)
        Me.pnlPropertiesAdmin.ResumeLayout(False)
        Me.pnlAvailableProperties.ResumeLayout(False)
        Me.pnlAssociatedProperties.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub PopulateEntity()
        Try
            Dim xView As DataView
            Dim oEntities As New MUSTER.BusinessLogic.pEntity
            bolLoadingCombo = True
            'oEntities.GetAll()
            oEntities.GetEntityAll()
            xView = oEntities.EntityCombo.DefaultView
            xView.Sort = "Entity Name"
            cmbEntity.DataSource = xView
            cmbEntity.DisplayMember = "Entity Name"
            cmbEntity.ValueMember = "Entity ID"
            cmbEntity.SelectedIndex = -1
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            bolLoadingCombo = False
        End Try
    End Sub



    Private Sub ugProperties_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugProperties.AfterRowActivate
        If ugProperties.ActiveRow.Cells(0).Value = 0 Then
            ugAssociatedProperties.DataSource = Nothing
            cboAAssociatedproperties.DataSource = Nothing
            btnAddProperty.Enabled = False
            Exit Sub
        End If
        'ugProperties.ActiveRow.Appearance.BackColor = system.Drawing.Color.
        Dim nSelectedPropertyID As Int64 = CType(ugProperties.ActiveRow.Cells("Property ID").Value, Int64)
        Dim i As Int32 = 0
        Dim nProperty_ID As Int64 = 0
        Try
            nAssociatedPropertyActiveIndex = -1
            If nPropertyActiveIndex <> -1 Then
                ugProperties.Rows(nPropertyActiveIndex).Appearance.BackColor = System.Drawing.Color.White
                ugProperties.Rows(nPropertyActiveIndex).Appearance.ForeColor = System.Drawing.Color.Black
            End If
            ugProperties.ActiveRow.Appearance.BackColor = System.Drawing.SystemColors.Highlight
            ugProperties.ActiveRow.Appearance.ForeColor = System.Drawing.Color.White

            dtAssociatedProperties.Rows.Clear()
            'dtAssociatedProperties = oPropType.PropertiesTable(nSelectedPropertyID)
            dtAssociatedProperties = oPropType.RetrieveChildProperties(nSelectedPropertyID).PropertiesTable

            ugAssociatedProperties.DataSource = dtAssociatedProperties
            ugAssociatedProperties.DataBind()
            For i = 0 To ugAssociatedProperties.DisplayLayout.Bands(0).Columns.Count - 1
                ugAssociatedProperties.DisplayLayout.Bands(0).Columns(i).CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            Next
            ugAssociatedProperties.DisplayLayout.Bands(0).Columns("Parent Property").Hidden = True

            'ugAssociatedProperties.ActiveRow.Selected = True
            fillAvailablePropertyCombo()

            nPropertyActiveIndex = ugProperties.ActiveRow.Index
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub fillAvailablePropertyCombo()
        '
        '  Modified 6/2/05 - JVC - Added MaintenanceMode parameter to calls to getAvailableProperties
        '
        Dim nProperty_ID As Int64 = 0
        Dim StrAssociatedPropertiesCaption As String
        Dim nStringIndex As Integer

        Dim nSelectedPropertyID As Int64 = CType(ugProperties.ActiveRow.Cells("Property ID").Value, Int64)
        If ugAssociatedProperties.Rows.Count > 0 And nSelectedPropertyID > 0 Then
            btnRemoveProperty.Enabled = True
            ugAssociatedProperties.Rows(0).Activate()
            If Not IsNothing(cboAAssociatedproperties.DataSource) And cboAAssociatedproperties.SelectedIndex <> -1 Then
                nProperty_ID = cboAAssociatedproperties.SelectedValue
            End If
            cboAAssociatedproperties.DataSource = oPropType.getAvailableProperties(nSelectedPropertyID, CType(ugAssociatedProperties.Rows(0).Cells("Property ID").Value, Int64), True)
        Else
            btnRemoveProperty.Enabled = False
            cboAAssociatedproperties.DataSource = oPropType.getAvailableProperties(nSelectedPropertyID, 0, True)
        End If
        If Not IsNothing(cboAAssociatedproperties.DataSource) Then
            cboAAssociatedproperties.DisplayMember = "AVAILABLE_PROPERTY_DISPLAY"
            cboAAssociatedproperties.ValueMember = "PROPERTY_ID"

        End If

        If nProperty_ID > 0 Then


            cboAAssociatedproperties.SelectedIndex = -1
            cboAAssociatedproperties.SelectedValue = nProperty_ID
        Else
            'To display the associated property type in the caption of associated property list
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            StrAssociatedPropertiesCaption = cboAAssociatedproperties.Text
            If StrAssociatedPropertiesCaption <> String.Empty Then
                nStringIndex = cboAAssociatedproperties.Text.IndexOf("-")
                ugAssociatedProperties.Text = "Property List of Associated Property Type " & StrAssociatedPropertiesCaption.Substring(0, nStringIndex - 1)
            Else
                ugAssociatedProperties.Text = "Associated Property List"
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            cboAAssociatedproperties.SelectedIndex = -1

        End If
    End Sub

    Private Sub cmbEntity_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbEntity.SelectedIndexChanged
        If bolLoadingCombo Then Exit Sub

        Try
            bolLoadingCombo = True


            cboPropertyType.DataSource = oPropType.GetByEntity(CLng(cmbEntity.SelectedValue)) 'oPropType.GetByEntity(CLng(cmbEntity.SelectedValue))
            cboPropertyType.DisplayMember = "Property Type Name"
            cboPropertyType.ValueMember = "Property Type ID"
            cboPropertyType.SelectedIndex = -1
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            'Clear the grids when new entity is selected
            Me.ugProperties.DataSource = Nothing
            Me.ugProperties.Text = "Properties"
            Me.ugAssociatedProperties.DataSource = Nothing
            Me.ugAssociatedProperties.Text = "Associated Properties"
            Me.cboAAssociatedproperties.DataSource = Nothing
            'If ugProperties.Rows.Count <= 0 Then
            '    Me.btnSaveProperties.Enabled = False
            'End If
            If ugAssociatedProperties.Rows.Count <= 0 Then
                Me.btnRemoveProperty.Enabled = False
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            bolLoadingCombo = False
        End Try
    End Sub

    Friend Sub CodeTableManager_Load()
        PopulateEntity()
    End Sub

    Private Sub RefreshData()
        Dim nPropActiveRow As Integer
        Dim nPropTypeSelIndex As Integer
        Try
            If ugProperties.Rows.Count > 0 Then
                nPropActiveRow = ugProperties.ActiveRow.Index
            End If
            'oPropType = New Muster.BusinessLogic.pPropertyType 'InfoRepository.MusterProperties
            nPropTypeSelIndex = cboPropertyType.SelectedIndex
            cboPropertyType.SelectedIndex = -1
            cboPropertyType.SelectedIndex = nPropTypeSelIndex

            ugProperties.Rows(nPropActiveRow).Activate()

        Catch ex As Exception

        End Try
    End Sub

    Private Sub cboPropertyType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboPropertyType.SelectedIndexChanged
        Dim strResult As MsgBoxResult
        Dim dtAssociatedPropertiesCaption As DataTable

        nPropertyActiveIndex = -1
        If bolLoadingCombo Or cboPropertyType.SelectedIndex = -1 Then Exit Sub

        Dim dRowProperties As DataRow
        Try
            If oPropType.IsDirty() And bolCancel = False Then
                strResult = MsgBox("Property Type " & oPropType.Name & " has been modified - do you wish to save?", MsgBoxStyle.YesNo Or MsgBoxStyle.Question, "Property Type Modified")
            End If
            bolCancel = False
            'oPropertyType = 
            oPropType.Retrieve(Integer.Parse(cboPropertyType.SelectedValue))  'oPropType.GetPropertyType(Integer.Parse(cboPropertyType.SelectedValue))
            dtProperties = oPropType.Properties.PropertiesTable 'PropertiesTable
            dtProperties.DefaultView.Sort = "PROPERTY POSITION ASC"
            ugProperties.DataSource = dtProperties.DefaultView
            ugProperties.DataBind()
            ugProperties.DisplayLayout.Bands(0).Columns("Property ID").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugProperties.DisplayLayout.Bands(0).Columns("Created On").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugProperties.DisplayLayout.Bands(0).Columns("Created By").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugProperties.DisplayLayout.Bands(0).Columns("Modified On").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugProperties.DisplayLayout.Bands(0).Columns("Modified By").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugProperties.DisplayLayout.Bands(0).Columns("Parent Property").Hidden = True
            ugProperties.Text = "Property List for Property Type " & oPropType.Name.ToString & " (" & oPropType.Name.ToString & ")"

            If cboPropertyType.Items.Count > 0 Then
                If dtProperties.Rows.Count > 0 Then
                    'dtAssociatedProperties = oPropType.PropertiesTable(CType(ugProperties.Rows(0).Cells("Property ID").Value, Int64))
                    dtAssociatedProperties = oPropType.RetrieveChildProperties(CType(ugProperties.Rows(0).Cells("Property ID").Value, Int64)).PropertiesTable
                    dtAssociatedProperties.DefaultView.Sort = "Property ID ASC"
                    ugAssociatedProperties.DataSource = dtAssociatedProperties.DefaultView
                    ugAssociatedProperties.DataBind()

                End If

            Else
                'cmbSecProps.DataSource = oPropertyType.PropertyCombo(0)
                'lblSecProps.Text = "No Related Properties"
            End If
            'If ugProperties.Rows.Count > 0 Then
            '    btnSaveProperties.Enabled = True
            'End If
            If ugAssociatedProperties.Rows.Count > 0 Then
                btnRemoveProperty.Enabled = True
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnSaveProperties_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveProperties.Click
        'Dim oProperties As New InfoRepository.MusterProperties
        Dim strErrMsg As String = String.Empty
        Dim strText As String = String.Empty
        Dim i As Integer = 0
        Try
            If Not ugrow Is Nothing Then
                For i = 0 To ugrow.Cells.Count - 1
                    strText += ugrow.Cells(i).Text
                    i += 1
                Next i
            End If
            If strText = String.Empty Then
                ugrow = ugProperties.Rows.GetRowWithListIndex(ugProperties.ActiveRow.Index)
            End If
            'MsgBox(ugrow.Cells("Property Name").Text + vbTab + ugrow.Cells("Property Position").Text)
            If ugrow.Cells("Property Name").Text = String.Empty Then
                strErrMsg += vbTab + "Property Name is missing" + vbCrLf
            End If
            If ugrow.Cells("Property Position").Text = String.Empty Then
                strErrMsg += vbTab + "Property Position is missing" + vbCrLf
            End If


            If strErrMsg.Trim.Length > 0 Then
                MsgBox("Invalid/Incomplete Property" + vbCrLf + strErrMsg)
                Exit Sub
            Else
                If Not IsNothing(dtProperties) Then
                    'oProperties.PutProperties(dtProperties, cboPropertyType.SelectedValue)
                End If
            End If
            Me.oPropType.Flush(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If
            Me.RefreshData()
            MsgBox("Property Save Successful", 0, "MUSTER Data Access")
            '
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub


    Private Sub cboAAssociatedproperties_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboAAssociatedproperties.SelectedIndexChanged
        Try
            If cboAAssociatedproperties.SelectedIndex = -1 Then
                btnAddProperty.Enabled = False
            Else
                btnAddProperty.Enabled = True
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnAddProperty_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddProperty.Click
        Try
            Dim oProperties As New MUSTER.BusinessLogic.pPropertyType
            Dim drow As DataRow
            Dim nActiveRow As Integer

            drow = dtAssociatedProperties.NewRow
            drow("Property ID") = cboAAssociatedproperties.SelectedValue
            drow("Parent Property") = ugProperties.ActiveRow.Cells("Property ID").Value
            dtAssociatedProperties.Rows.Add(drow)
            'nActiveRow = ugProperties.ActiveRow.Index

            'MsgBox(dtPropertyRel.Rows.Count)
            Dim success As Boolean = False
            success = oProperties.PutPropertyRelation(dtAssociatedProperties, CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If
            If success Then
                'dtAssociatedProperties = oPropType.PropertiesTable(CType(ugProperties.Rows(0).Cells("Property ID").Value, Int64))
                dtAssociatedProperties = oPropType.RetrieveChildProperties(CType(ugProperties.Rows(0).Cells("Property ID").Value, Int64)).PropertiesTable
                ugAssociatedProperties.DataSource = dtAssociatedProperties
                ugAssociatedProperties.DataBind()
                RefreshData()
                MsgBox("Add Successful", 0, "MUSTER Data Access")
            End If

            'ugProperties.Rows(0).Activate()
            'ugProperties.Rows(nActiveRow).Activate()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnRemoveProperty_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemoveProperty.Click
        Dim bolSuccess As Boolean
        Try
            'Dim oProperties As New Muster.BusinessLogic.pPropertyType
            bolSuccess = oPropType.DeletePropertyRelation(ugAssociatedProperties.ActiveRow.Cells("Parent Property").Value, ugAssociatedProperties.ActiveRow.Cells("Property ID").Value) 'Then                       
            dtAssociatedProperties = oPropType.RetrieveChildProperties(CType(ugProperties.Rows(0).Cells("Property ID").Value, Int64)).PropertiesTable
            ugAssociatedProperties.DataSource = dtAssociatedProperties
            ugAssociatedProperties.DataBind()
            RefreshData()
            'End If
            If bolSuccess Then
                MsgBox("Remove Successful", 0, "MUSTER Data Access")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub ugAssociatedProperties_AfterRowActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugAssociatedProperties.AfterRowActivate
        Try
            If nAssociatedPropertyActiveIndex <> -1 Then
                ugAssociatedProperties.Rows(nAssociatedPropertyActiveIndex).Appearance.BackColor = System.Drawing.Color.White
                ugAssociatedProperties.Rows(nAssociatedPropertyActiveIndex).Appearance.ForeColor = System.Drawing.Color.Black
            End If
            ugAssociatedProperties.ActiveRow.Appearance.BackColor = System.Drawing.SystemColors.Highlight
            ugAssociatedProperties.ActiveRow.Appearance.ForeColor = System.Drawing.Color.White
            nAssociatedPropertyActiveIndex = ugAssociatedProperties.ActiveRow.Index
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugProperties_AfterCellActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugProperties.AfterCellActivate
        ' If ugProperties.ActiveCell.Text = String.Empty Then
        If ugProperties.Rows.Count > 0 Then
            ugProperties.ActiveRow = ugProperties.ActiveCell.Row
            ugrow = ugProperties.Rows.GetRowWithListIndex(ugProperties.ActiveRow.Index)
        End If
    End Sub


    Private Sub ugProperties_AfterRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles ugProperties.AfterRowUpdate
        'MsgBox("here" & ugProperties.ActiveRow.Index & "-" & ugProperties.Rows.Count)
        'MsgBox(e.Row.IsAddRow)
        'MsgBox(e.Row.Cells.Item("Parent Property").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw) & "-" & e.Row.Cells.Item("Property ID").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw))
        If e.Row.Cells.Item("Property ID").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw) <> "0" Then

            btnSaveProperties.Enabled = True
            'e.Row.Cells.
            'MsgBox(oPropType.RetrieveProperty(ugProperties.ActiveRow.Cells.Item("Property ID").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw)).Name)
            'MsgBox(ugProperties.ActiveRow.Cells.Item("Property ID").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw))
            Dim oProperty As MUSTER.Info.PropertyInfo
            oProperty = oPropType.RetrieveProperty(ugProperties.ActiveRow.Cells.Item("Property ID").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw))
            With ugProperties.ActiveRow.Cells
                Me.oPropType.PropertyName = .Item("Property Name").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw)
                Me.oPropType.PropertyPropDesc = .Item("Property Description").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw)
                Me.oPropType.PropertyPropPos = .Item("Property Position").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw)
                Me.oPropType.PropertyPropIsActive = .Item("Property Active").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw)
       
                'oProperty. = .Item("Parent Property").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw)
                'oProperty.Name = .Item("Property ID").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw)
                'oProperty.Name = .Item("Property ID").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw)
            End With
            'MsgBox(oProperty.IsDirty)
        Else
            If ActiveControl.Name <> "btnReset" Then
                Dim LocalPropertyInfo As New MUSTER.Info.PropertyInfo
                'LocalPropertyInfo.ID = Nothing
                LocalPropertyInfo.BUSINESSTAG = 0
                LocalPropertyInfo.PropDesc = ""
                LocalPropertyInfo.Name = e.Row.Cells.Item("Property Name").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw)
                LocalPropertyInfo.PropDesc = e.Row.Cells.Item("Property Description").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw)
                LocalPropertyInfo.PropIsActive = e.Row.Cells.Item("Property Active").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw)
                If e.Row.Cells.Item("Property Position").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw) <> "" Then
                    LocalPropertyInfo.PropPos = e.Row.Cells.Item("Property Position").GetText(Infragistics.Win.UltraWinMaskedEdit.MaskMode.Raw)
                End If
                LocalPropertyInfo.PropType_ID = Me.cboPropertyType.SelectedValue.ToString
                oPropType.Properties.Add(LocalPropertyInfo)
                If ActiveControl.Name <> "btnClose" Then
                    oPropType.Flush(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If
                Else
                    btnClose_Click(sender, e)
                End If
                'cmbEntity_SelectedIndexChanged(sender, e)
                'Me.btnSaveProperties_Click(sender, e)
                'oPropType.Properties = Nothing
            End If
            RefreshData()
            'Me.cboPropertyType_SelectedIndexChanged(sender, e)
        End If
    End Sub

    Protected Overrides Sub OnActivated(ByVal e As System.EventArgs)
        If cmbEntity.Items.Count = 0 Then Me.Close()
    End Sub

    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)
        If Not oPropType Is Nothing Then
            If oPropType.IsDirty Then
                Dim Results As Long = MsgBox("There are unsaved changes. Do you want to save changes before closing?", MsgBoxStyle.YesNoCancel)
                If Results = MsgBoxResult.Yes Then
                    oPropType.Flush(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If
                Else
                    If Results = MsgBoxResult.Cancel Then
                        bolCancel = True
                        e.Cancel = True
                        Exit Sub
                    End If
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

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub PropertyTypeChanged(ByVal bolValue As Boolean) Handles oPropType.ColChanged, oPropType.PropertyColChanged, oPropType.PropertyInfoChanged, oPropType.PropertyTypeChanged
        'If bolLoading Then Exit Sub
        SetSaveCancel(oPropType.IsDirty)
    End Sub

    Private Sub SetSaveCancel(ByVal bolValue As Boolean)
        Me.btnSaveProperties.Enabled = bolValue
        Me.btnReset.Enabled = bolValue
    End Sub

    Private Sub btnReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Me.oPropType.Reset()
        cboPropertyType_SelectedIndexChanged(Me.cboPropertyType, e)
    End Sub

    Private Sub ugProperties_AfterRowInsert(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles ugProperties.AfterRowInsert
        'Dim LocalPropertyInfo As New MUSTER.Info.PropertyInfo
        'LocalPropertyInfo.BUSINESSTAG = 0
        'LocalPropertyInfo.PropDesc = ""
        'LocalPropertyInfo.PropType_ID = Me.cboPropertyType.SelectedValue.ToString
        'oPropType.Properties.Add(LocalPropertyInfo)        
        'MsgBox(oPropType.PropertyID)
        'e.Row.Cells.Item("Property ID").Value = LocalPropertyInfo.ID

        Dim propPos As Integer = 1
        For Each drow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugProperties.Rows
            If Not drow.IsAddRow Then
                If drow.Cells("Property Position").Value > propPos Then
                    propPos = drow.Cells("Property Position").Value
                End If
            End If
        Next
        e.Row.Cells("Property Position").Value = propPos + 1

        Me.btnSaveProperties.Enabled = True
        Me.btnReset.Enabled = True
    End Sub

    'Private Sub ugProperties_BeforeRowInsert(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.BeforeRowInsertEventArgs) Handles ugProperties.BeforeRowInsert
    ' check if there is a property with position 0 (zero)
    ' if there is, do not allow adding new row
    'For Each drow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugProperties.Rows
    '    If drow.Cells("Property Position").Value = 0 Then
    '        MsgBox("Unable to add new row:" + vbCrLf + _
    '        "Column 'Property Position' is constrained to be unique. Value '0' is already present.", MsgBoxStyle.Exclamation, "Data Error")
    '        e.Cancel = True
    '        Exit For
    '    End If
    'Next
    'End Sub

    Private Sub ugProperties_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugProperties.CellChange
        If "Property Position".Equals(e.Cell.Column.Key) Then
            If e.Cell.Text = "0" Then
                MsgBox("Unable to update Property Position:" + vbCrLf + _
                "Value must be greater than 0.", MsgBoxStyle.Exclamation, "Data Error")
                e.Cell.CancelUpdate()
            End If
        End If
    End Sub
End Class
