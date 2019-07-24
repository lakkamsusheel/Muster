Public Class ActivityList
    '*******************************************************************
    '
    ' Modified 7/25/05 by JVC to clear grid when NEW button is clicked...
    ' Modified 7/28/05 by JVC to check for activity use before deletion
    '
    '*******************************************************************
    Inherits System.Windows.Forms.Form
    Private WithEvents oTecActivity As New MUSTER.BusinessLogic.pTecAct
    Private WithEvents oTecDocument As New MUSTER.BusinessLogic.pTecDoc
    Private bolLoading As Boolean = True
    Private bolIsFinancial As Boolean = False
    Private returnVal As String = String.Empty

    Public Property IsFinancial() As Boolean
        Get
            Return bolIsFinancial
        End Get
        Set(ByVal Value As Boolean)
            bolIsFinancial = Value
        End Set
    End Property

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
    Friend WithEvents cmbActivity As System.Windows.Forms.ComboBox
    Friend WithEvents chkActive As System.Windows.Forms.CheckBox
    Friend WithEvents lblWarn As System.Windows.Forms.Label
    Friend WithEvents lblActivity As System.Windows.Forms.Label
    Friend WithEvents lblAct As System.Windows.Forms.Label
    Friend WithEvents txtWarn As System.Windows.Forms.TextBox
    Friend WithEvents txtAct As System.Windows.Forms.TextBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents GBDocument As System.Windows.Forms.GroupBox
    Friend WithEvents ugAvailableDocuments As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnRemove As System.Windows.Forms.Button
    Friend WithEvents cmbDocument As System.Windows.Forms.ComboBox
    Friend WithEvents btnAdd As System.Windows.Forms.Button
    Friend WithEvents txtActivity As System.Windows.Forms.TextBox
    Friend WithEvents btnAddNew As System.Windows.Forms.Button
    Friend WithEvents chkShowAll As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtCost As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cboCostMode As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmbActivity = New System.Windows.Forms.ComboBox
        Me.chkActive = New System.Windows.Forms.CheckBox
        Me.lblWarn = New System.Windows.Forms.Label
        Me.lblActivity = New System.Windows.Forms.Label
        Me.lblAct = New System.Windows.Forms.Label
        Me.txtWarn = New System.Windows.Forms.TextBox
        Me.txtAct = New System.Windows.Forms.TextBox
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.GBDocument = New System.Windows.Forms.GroupBox
        Me.ugAvailableDocuments = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.btnRemove = New System.Windows.Forms.Button
        Me.cmbDocument = New System.Windows.Forms.ComboBox
        Me.btnAdd = New System.Windows.Forms.Button
        Me.btnAddNew = New System.Windows.Forms.Button
        Me.txtActivity = New System.Windows.Forms.TextBox
        Me.chkShowAll = New System.Windows.Forms.CheckBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtCost = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.cboCostMode = New System.Windows.Forms.ComboBox
        Me.GBDocument.SuspendLayout()
        CType(Me.ugAvailableDocuments, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmbActivity
        '
        Me.cmbActivity.Location = New System.Drawing.Point(104, 24)
        Me.cmbActivity.Name = "cmbActivity"
        Me.cmbActivity.Size = New System.Drawing.Size(344, 21)
        Me.cmbActivity.TabIndex = 0
        '
        'chkActive
        '
        Me.chkActive.Checked = True
        Me.chkActive.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkActive.Location = New System.Drawing.Point(464, 48)
        Me.chkActive.Name = "chkActive"
        Me.chkActive.Size = New System.Drawing.Size(72, 24)
        Me.chkActive.TabIndex = 4
        Me.chkActive.Text = "Active"
        '
        'lblWarn
        '
        Me.lblWarn.Location = New System.Drawing.Point(184, 80)
        Me.lblWarn.Name = "lblWarn"
        Me.lblWarn.Size = New System.Drawing.Size(56, 16)
        Me.lblWarn.TabIndex = 140
        Me.lblWarn.Text = "Warn:"
        Me.lblWarn.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblActivity
        '
        Me.lblActivity.Location = New System.Drawing.Point(24, 24)
        Me.lblActivity.Name = "lblActivity"
        Me.lblActivity.Size = New System.Drawing.Size(64, 16)
        Me.lblActivity.TabIndex = 139
        Me.lblActivity.Text = "Activity:"
        Me.lblActivity.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'lblAct
        '
        Me.lblAct.Location = New System.Drawing.Point(32, 80)
        Me.lblAct.Name = "lblAct"
        Me.lblAct.Size = New System.Drawing.Size(56, 16)
        Me.lblAct.TabIndex = 138
        Me.lblAct.Text = "Act:"
        Me.lblAct.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtWarn
        '
        Me.txtWarn.Location = New System.Drawing.Point(240, 80)
        Me.txtWarn.Name = "txtWarn"
        Me.txtWarn.Size = New System.Drawing.Size(72, 20)
        Me.txtWarn.TabIndex = 2
        Me.txtWarn.Text = ""
        '
        'txtAct
        '
        Me.txtAct.Location = New System.Drawing.Point(104, 80)
        Me.txtAct.Name = "txtAct"
        Me.txtAct.Size = New System.Drawing.Size(72, 20)
        Me.txtAct.TabIndex = 1
        Me.txtAct.Text = ""
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(240, 408)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 23)
        Me.btnDelete.TabIndex = 11
        Me.btnDelete.Text = "Delete"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(328, 408)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(80, 23)
        Me.btnCancel.TabIndex = 10
        Me.btnCancel.Text = "Cancel"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(152, 408)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(80, 23)
        Me.btnSave.TabIndex = 9
        Me.btnSave.Text = "Save"
        '
        'GBDocument
        '
        Me.GBDocument.Controls.Add(Me.ugAvailableDocuments)
        Me.GBDocument.Controls.Add(Me.btnRemove)
        Me.GBDocument.Controls.Add(Me.cmbDocument)
        Me.GBDocument.Controls.Add(Me.btnAdd)
        Me.GBDocument.Location = New System.Drawing.Point(8, 104)
        Me.GBDocument.Name = "GBDocument"
        Me.GBDocument.Size = New System.Drawing.Size(536, 288)
        Me.GBDocument.TabIndex = 200
        Me.GBDocument.TabStop = False
        Me.GBDocument.Text = "Documents"
        '
        'ugAvailableDocuments
        '
        Me.ugAvailableDocuments.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugAvailableDocuments.DisplayLayout.AutoFitColumns = True
        Me.ugAvailableDocuments.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle
        Me.ugAvailableDocuments.DisplayLayout.Override.SelectTypeRow = Infragistics.Win.UltraWinGrid.SelectType.Single
        Me.ugAvailableDocuments.Location = New System.Drawing.Point(152, 56)
        Me.ugAvailableDocuments.Name = "ugAvailableDocuments"
        Me.ugAvailableDocuments.Size = New System.Drawing.Size(336, 192)
        Me.ugAvailableDocuments.TabIndex = 204
        Me.ugAvailableDocuments.Text = "Available Documents"
        '
        'btnRemove
        '
        Me.btnRemove.Location = New System.Drawing.Point(32, 56)
        Me.btnRemove.Name = "btnRemove"
        Me.btnRemove.Size = New System.Drawing.Size(112, 25)
        Me.btnRemove.TabIndex = 202
        Me.btnRemove.Text = "Remove Document"
        '
        'cmbDocument
        '
        Me.cmbDocument.Location = New System.Drawing.Point(152, 24)
        Me.cmbDocument.Name = "cmbDocument"
        Me.cmbDocument.Size = New System.Drawing.Size(336, 21)
        Me.cmbDocument.TabIndex = 200
        '
        'btnAdd
        '
        Me.btnAdd.Location = New System.Drawing.Point(32, 24)
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.Size = New System.Drawing.Size(112, 24)
        Me.btnAdd.TabIndex = 201
        Me.btnAdd.Text = "Add Document"
        '
        'btnAddNew
        '
        Me.btnAddNew.Location = New System.Drawing.Point(0, 0)
        Me.btnAddNew.Name = "btnAddNew"
        Me.btnAddNew.Size = New System.Drawing.Size(64, 16)
        Me.btnAddNew.TabIndex = 201
        Me.btnAddNew.Text = "Add New"
        '
        'txtActivity
        '
        Me.txtActivity.Location = New System.Drawing.Point(104, 48)
        Me.txtActivity.Name = "txtActivity"
        Me.txtActivity.Size = New System.Drawing.Size(344, 20)
        Me.txtActivity.TabIndex = 202
        Me.txtActivity.Text = ""
        '
        'chkShowAll
        '
        Me.chkShowAll.Location = New System.Drawing.Point(464, 72)
        Me.chkShowAll.Name = "chkShowAll"
        Me.chkShowAll.Size = New System.Drawing.Size(80, 24)
        Me.chkShowAll.TabIndex = 203
        Me.chkShowAll.Text = "Show All"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 16)
        Me.Label1.TabIndex = 204
        Me.Label1.Text = "Activity Name:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'txtCost
        '
        Me.txtCost.Enabled = False
        Me.txtCost.Location = New System.Drawing.Point(376, 112)
        Me.txtCost.Name = "txtCost"
        Me.txtCost.Size = New System.Drawing.Size(72, 20)
        Me.txtCost.TabIndex = 206
        Me.txtCost.Text = ""
        '
        'Label2
        '
        Me.Label2.Enabled = False
        Me.Label2.Location = New System.Drawing.Point(280, 112)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(96, 16)
        Me.Label2.TabIndex = 208
        Me.Label2.Text = "Estimated Cost"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(32, 112)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(64, 16)
        Me.Label3.TabIndex = 207
        Me.Label3.Text = "Cost Mode:"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cboCostMode
        '
        Me.cboCostMode.Location = New System.Drawing.Point(104, 112)
        Me.cboCostMode.Name = "cboCostMode"
        Me.cboCostMode.Size = New System.Drawing.Size(168, 21)
        Me.cboCostMode.TabIndex = 209
        '
        'ActivityList
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(552, 446)
        Me.Controls.Add(Me.GBDocument)
        Me.Controls.Add(Me.cboCostMode)
        Me.Controls.Add(Me.txtCost)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.chkShowAll)
        Me.Controls.Add(Me.btnAddNew)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.chkActive)
        Me.Controls.Add(Me.txtWarn)
        Me.Controls.Add(Me.txtAct)
        Me.Controls.Add(Me.txtActivity)
        Me.Controls.Add(Me.lblWarn)
        Me.Controls.Add(Me.lblActivity)
        Me.Controls.Add(Me.lblAct)
        Me.Controls.Add(Me.cmbActivity)
        Me.Name = "ActivityList"
        Me.Text = "Add / Modify Activity List"
        Me.GBDocument.ResumeLayout(False)
        CType(Me.ugAvailableDocuments, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Page Event Routines "

    Private Sub ActivityList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadForm()

        If IsFinancial Then

            GBDocument.Visible = False
            btnAddNew.Visible = False
            txtAct.Enabled = False
            txtWarn.Enabled = False
            txtActivity.Enabled = False
            chkActive.Enabled = False
            chkShowAll.Enabled = False
            cmbDocument.Enabled = False
            btnDelete.Visible = False

        End If

    End Sub
    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        SaveActivity()
    End Sub
    Private Sub btnAddNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddNew.Click

        Try

            If oTecActivity.IsDirty Then
                Select Case MsgBox("Changes were made.  Do you wish to save before continuing?", MsgBoxStyle.YesNoCancel)
                    Case MsgBoxResult.Yes
                        SaveActivity()
                    Case MsgBoxResult.No
                        Exit Select
                    Case MsgBoxResult.Cancel
                        Exit Sub
                End Select
            End If

            bolLoading = True
            cmbActivity.SelectedIndex = -1
            cmbActivity.SelectedIndex = -1
            cmbActivity.Enabled = False
            txtActivity.Visible = True
            txtActivity.Text = String.Empty
            bolLoading = False

            oTecActivity.Retrieve(0)

            txtWarn.Text = 0
            txtAct.Text = 0
            chkActive.Checked = True
            oTecActivity.Active = True
            oTecActivity.ActDays = 0
            oTecActivity.WarnDays = 0
            ugAvailableDocuments.DataSource = Nothing

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Add New " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()

        End Try
    End Sub
    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        'cmbActivity.Visible = True
        'txtActivity.Visible = False

        Me.Close()

    End Sub
    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        Dim oTecDoc As New MUSTER.BusinessLogic.pTecDoc
        Dim oTecDocInfo As MUSTER.Info.TecDocInfo
        Dim bolDuplicate As Boolean = False

        For Each oTecDocInfo In oTecActivity.DocumentsCollection.Values
            If oTecDocInfo.ID = cmbDocument.SelectedValue Then
                bolDuplicate = True
            End If
        Next
        If Not bolDuplicate Then
            oTecDoc.Retrieve(cmbDocument.SelectedValue)
            oTecActivity.DocumentsCollection.Add(oTecDoc.InfoObject)
            oTecActivity.IsDirty = True
        End If

        LoadDocumentGrid()
    End Sub
    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try
            If MsgBox("Do you wish to delete this Activity?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                If oTecActivity.InUse Then
                    MsgBox("The activity is currently in use and may not be deleted.", MsgBoxStyle.Exclamation & MsgBoxStyle.OKOnly, "Activity In Use")
                Else
                    oTecActivity.Deleted = True
                    oTecActivity.ModifiedBy = MusterContainer.AppUser.ID
                    oTecActivity.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), MusterContainer.AppUser.UserKey, returnVal)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If
                    LoadForm()
                End If
            End If
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Delete " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRemove.Click
        Dim oTecDocInfo As MUSTER.Info.TecDocInfo
        Dim bolFound As Boolean = False
        If ugAvailableDocuments.Rows.Count > 0 Then
            For Each oTecDocInfo In oTecActivity.DocumentsCollection.Values
                If oTecDocInfo.Name = ugAvailableDocuments.ActiveRow.Cells("value").Value Then
                    bolFound = True
                    Exit For
                End If
            Next
            If bolFound Then
                oTecActivity.DocumentsCollection.Remove(oTecDocInfo)
                LoadDocumentGrid()
                oTecActivity.IsDirty = True
            End If
        End If
    End Sub

    Private Sub txtActivity_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtActivity.TextChanged
        If bolLoading Then Exit Sub
        oTecActivity.Name = txtActivity.Text
    End Sub
    Private Sub txtAct_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtAct.TextChanged
        If bolLoading Then Exit Sub
        If IsNumeric(txtAct.Text) Then
            oTecActivity.ActDays = txtAct.Text
        ElseIf txtAct.Text = "" Then
            oTecActivity.ActDays = 0
        Else
            txtAct.Text = 0
            oTecActivity.ActDays = 0
            MsgBox("Numeric Data Only Allowed In This Field")
        End If

    End Sub
    Private Sub txtWarn_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtWarn.TextChanged
        If bolLoading Then Exit Sub

        If IsNumeric(txtWarn.Text) Then
            oTecActivity.WarnDays = txtWarn.Text
        ElseIf txtWarn.Text = "" Then
            oTecActivity.WarnDays = 0
        Else
            txtWarn.Text = 0
            oTecActivity.WarnDays = 0
            MsgBox("Numeric Data Only Allowed In This Field")
        End If
    End Sub
    Private Sub cmbActivity_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbActivity.SelectedIndexChanged
        If bolLoading Then Exit Sub

        Try
            If oTecActivity.IsDirty Then
                Select Case MsgBox("Changes were made.  Do you wish to save before continuing?", MsgBoxStyle.YesNoCancel)
                    Case MsgBoxResult.Yes
                        SaveActivity()
                    Case MsgBoxResult.No
                        Exit Select
                    Case MsgBoxResult.Cancel
                        Exit Sub
                End Select
            End If
            LoadActivityForm()
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Change cmbActivity " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub


    Private Sub cboCostMode_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cboCostMode.SelectedIndexChanged

        If IsFinancial AndAlso TypeOf cboCostMode.SelectedValue Is Integer AndAlso Not cboCostMode.SelectedValue = MUSTER.Info.TecActInfo.ActivityCostModeEnum.NotFundable Then
            txtCost.Enabled = True

        Else
            txtCost.Text = "$0.00"
            txtCost.Enabled = False

        End If

    End Sub

    Private Sub chkActive_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkActive.CheckedChanged
        If bolLoading Then Exit Sub

        oTecActivity.Active = chkActive.Checked

    End Sub
    Private Sub chkShowAll_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowAll.CheckedChanged
        Try
            If chkShowAll.Checked Then
                PopulateActivityList(True)
            Else
                PopulateActivityList(False)
            End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Load System DropDowns " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub

#End Region

#Region " Populate Routines "

    Private Sub LoadForm()
        bolLoading = True
        LoadDropDowns()
        cmbActivity.SelectedIndex = 0
        cmbDocument.SelectedIndex = 0
        bolLoading = False
        LoadActivityForm()
    End Sub
    Private Sub PopulateActivityList(Optional ByVal ShowAll As Boolean = False)
        Dim dtActList As DataTable
        Try
            If ShowAll Then
                dtActList = oTecActivity.PopulateTecActivityList(False)
            Else
                dtActList = oTecActivity.PopulateTecActivityList(False, False)
            End If

            cmbActivity.DataSource = dtActList
            cmbActivity.DisplayMember = "PROPERTY_NAME"
            cmbActivity.ValueMember = "PROPERTY_ID"

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Load Activity " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub LoadDropDowns()

        Dim dtDocList As DataTable
        Dim dtCostTypes As DataTable

        Try

            dtDocList = oTecDocument.PopulateTecDocumentList(False)
            PopulateActivityList(False)
            cmbDocument.DataSource = dtDocList
            cmbDocument.DisplayMember = "PROPERTY_NAME"
            cmbDocument.ValueMember = "PROPERTY_ID"

            dtCostTypes = Me.oTecActivity.PopulateCostModes
            cboCostMode.DataSource = dtCostTypes
            cboCostMode.DisplayMember = "PROPERTY_NAME"
            cboCostMode.ValueMember = "PROPERTY_ID"

            'If cmbActivity.Items.Count > 0 Then
            '    cmbActivity.Visible = True
            '    txtActivity.Visible = False
            'Else
            '    cmbActivity.Visible = False
            '    txtActivity.Visible = True
            'End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Load System DropDowns " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub
    Private Sub LoadActivityForm()

        Try
            bolLoading = True
            If Not cmbActivity.SelectedValue Is Nothing Then

                oTecActivity.Retrieve(cmbActivity.SelectedValue)
                txtCost.Text = String.Format("{0:C}", oTecActivity.Cost)
                cboCostMode.SelectedValue = oTecActivity.CostMode
                txtActivity.Text = oTecActivity.Name
                txtAct.Text = oTecActivity.ActDays
                txtWarn.Text = oTecActivity.WarnDays
                chkActive.Checked = oTecActivity.Active
                LoadDocumentGrid()
            End If


        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Load Activity Form" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        Finally
            bolLoading = False
        End Try

    End Sub
    Private Sub LoadDocumentGrid()
        Dim oTecDoc As New MUSTER.BusinessLogic.pTecDoc
        Dim oTecDocInfo As MUSTER.Info.TecDocInfo
        Dim alDocuments As New ArrayList

        For Each oTecDocInfo In oTecActivity.DocumentsCollection.Values
            alDocuments.Add(oTecDocInfo.Name)
        Next

        ugAvailableDocuments.DataSource = alDocuments
        ugAvailableDocuments.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        If ugAvailableDocuments.Rows.Count > 0 Then
            ugAvailableDocuments.DisplayLayout.Bands(0).Columns("Value").Header.Caption = "Document"
            ugAvailableDocuments.Rows.Band.Columns(0).SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
        End If

    End Sub

#End Region

#Region " General Routines "

    Private Sub SaveActivity()
        Dim bolReload As Boolean = False

        Try
            If Not (txtAct.Text > "0") Then
                MsgBox("Act Days Must Be Greater Than Zero.")
                txtAct.Focus()
                Exit Sub
            End If
            If Not (txtWarn.Text > "0") Then
                MsgBox("Warn Days Must Be Greater Than Zero.")
                txtWarn.Focus()
                Exit Sub
            End If
            If Not (txtActivity.Text > "") And (oTecActivity.Activity_ID <= 0) Then
                MsgBox("Activity Name Required.")
                txtActivity.Focus()
                Exit Sub
            End If

            txtCost.Text = txtCost.Text.Replace("$", String.Empty)

            If IsFinancial Then

                If txtCost.Text.Length > 0 AndAlso Not IsNumeric(txtCost.Text) Then
                    MsgBox("Cost must be in numeric form.")
                    txtCost.Focus()
                    Exit Sub
                ElseIf txtCost.Text.Length = 0 Then
                    txtCost.Text = "0"
                ElseIf cboCostMode.SelectedValue = 0 Then
                    txtCost.Text = "0"
                End If

                oTecActivity.Cost = Convert.ToDouble(txtCost.Text)
                oTecActivity.CostMode = cboCostMode.SelectedValue

            End If


            If oTecActivity.Activity_ID <= 0 Then
                bolReload = True
                oTecActivity.CreatedBy = MusterContainer.AppUser.ID
            Else
                oTecActivity.ModifiedBy = MusterContainer.AppUser.ID
            End If

            oTecActivity.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), MusterContainer.AppUser.UserKey, returnVal)
            If Not returnVal = String.Empty Then
                MessageBox.Show(returnVal.ToString(), "ACCESS DENIED", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            Else
                MsgBox("Activity Saved Successfully.")
            End If
            cmbActivity.Enabled = True
            ' If bolReload Then
            bolLoading = True
            LoadDropDowns()
            cmbActivity.SelectedIndex = -1
            txtActivity.Text = String.Empty
            bolLoading = False
            cmbActivity.SelectedValue = oTecActivity.Activity_ID
            If Not oTecActivity.Active Then
                txtActivity.Text = String.Empty
            Else
                txtActivity.Text = oTecActivity.Name
            End If

            'cmbActivity.Visible = True
            'txtActivity.Visible = False
            '  End If

        Catch ex As Exception
            If ex.Message = "Duplicate Entry" Then
                MessageBox.Show("Duplicate Activity Name. Please Enter different Name.", "Duplicate Entry")
                txtActivity.Focus()
                Exit Sub
            End If
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Add New " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub

#End Region

End Class
