Public Class LicenseeComplianceEvent
    Inherits System.Windows.Forms.Form
#Region "Private User Variables"
    Dim pLCE As MUSTER.BusinessLogic.pLicenseeComplianceEvent
    Friend WithEvents objCitation As CitationList
    Dim dtCitation As New DataTable
    Dim strFrom As String = ""
    Dim bolLoading As Boolean
    Dim returnVal As String = String.Empty
#End Region
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        pLCE = New MUSTER.BusinessLogic.pLicenseeComplianceEvent
    End Sub

    Public Sub New(ByRef LCE As MUSTER.BusinessLogic.pLicenseeComplianceEvent, Optional ByVal From As String = "EDIT")
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        pLCE = LCE
        strFrom = From
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
    Public WithEvents dtComplianceEventDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblExceptionGrantedDate As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnCitationDelete As System.Windows.Forms.Button
    Friend WithEvents btnCitationAdd As System.Windows.Forms.Button
    Friend WithEvents ugCitations As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents lblLicensee As System.Windows.Forms.Label
    Public WithEvents txtLicensee As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmbSearchForOwner As System.Windows.Forms.ComboBox
    Friend WithEvents chkSearchforOwner As System.Windows.Forms.CheckBox
    Friend WithEvents cmbFacility As System.Windows.Forms.ComboBox
    Friend WithEvents lblFacility As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtFacilityID As System.Windows.Forms.TextBox
    Friend WithEvents btnLicensees As System.Windows.Forms.Button
    Public WithEvents dtCitationDueDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtCitationDueDate As System.Windows.Forms.Label
    Friend WithEvents pnlLCEBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlLCEDetails As System.Windows.Forms.Panel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.dtComplianceEventDate = New System.Windows.Forms.DateTimePicker
        Me.lblExceptionGrantedDate = New System.Windows.Forms.Label
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnCitationDelete = New System.Windows.Forms.Button
        Me.btnCitationAdd = New System.Windows.Forms.Button
        Me.ugCitations = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.btnLicensees = New System.Windows.Forms.Button
        Me.lblLicensee = New System.Windows.Forms.Label
        Me.txtLicensee = New System.Windows.Forms.TextBox
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtFacilityID = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.cmbSearchForOwner = New System.Windows.Forms.ComboBox
        Me.chkSearchforOwner = New System.Windows.Forms.CheckBox
        Me.cmbFacility = New System.Windows.Forms.ComboBox
        Me.lblFacility = New System.Windows.Forms.Label
        Me.dtCitationDueDate = New System.Windows.Forms.DateTimePicker
        Me.txtCitationDueDate = New System.Windows.Forms.Label
        Me.pnlLCEBottom = New System.Windows.Forms.Panel
        Me.pnlLCEDetails = New System.Windows.Forms.Panel
        CType(Me.ugCitations, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox1.SuspendLayout()
        Me.pnlLCEBottom.SuspendLayout()
        Me.pnlLCEDetails.SuspendLayout()
        Me.SuspendLayout()
        '
        'dtComplianceEventDate
        '
        Me.dtComplianceEventDate.Checked = False
        Me.dtComplianceEventDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtComplianceEventDate.Location = New System.Drawing.Point(152, 168)
        Me.dtComplianceEventDate.Name = "dtComplianceEventDate"
        Me.dtComplianceEventDate.ShowCheckBox = True
        Me.dtComplianceEventDate.Size = New System.Drawing.Size(104, 20)
        Me.dtComplianceEventDate.TabIndex = 249
        '
        'lblExceptionGrantedDate
        '
        Me.lblExceptionGrantedDate.Location = New System.Drawing.Point(8, 168)
        Me.lblExceptionGrantedDate.Name = "lblExceptionGrantedDate"
        Me.lblExceptionGrantedDate.Size = New System.Drawing.Size(128, 16)
        Me.lblExceptionGrantedDate.TabIndex = 257
        Me.lblExceptionGrantedDate.Text = "Compliance Event Date:"
        Me.lblExceptionGrantedDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(304, 5)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(88, 23)
        Me.btnClose.TabIndex = 254
        Me.btnClose.Text = "Close"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(208, 5)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(88, 23)
        Me.btnCancel.TabIndex = 255
        Me.btnCancel.Text = "Cancel"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(112, 5)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(88, 23)
        Me.btnSave.TabIndex = 253
        Me.btnSave.Text = "Save"
        '
        'btnCitationDelete
        '
        Me.btnCitationDelete.Location = New System.Drawing.Point(104, 376)
        Me.btnCitationDelete.Name = "btnCitationDelete"
        Me.btnCitationDelete.Size = New System.Drawing.Size(88, 23)
        Me.btnCitationDelete.TabIndex = 252
        Me.btnCitationDelete.Text = "Delete Citation"
        '
        'btnCitationAdd
        '
        Me.btnCitationAdd.Location = New System.Drawing.Point(8, 376)
        Me.btnCitationAdd.Name = "btnCitationAdd"
        Me.btnCitationAdd.Size = New System.Drawing.Size(88, 23)
        Me.btnCitationAdd.TabIndex = 251
        Me.btnCitationAdd.Text = "Add Citation"
        '
        'ugCitations
        '
        Me.ugCitations.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugCitations.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugCitations.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugCitations.DisplayLayout.ScrollBounds = Infragistics.Win.UltraWinGrid.ScrollBounds.ScrollToFill
        Me.ugCitations.Location = New System.Drawing.Point(6, 230)
        Me.ugCitations.Name = "ugCitations"
        Me.ugCitations.Size = New System.Drawing.Size(400, 136)
        Me.ugCitations.TabIndex = 250
        Me.ugCitations.Text = "Citations"
        '
        'btnLicensees
        '
        Me.btnLicensees.BackColor = System.Drawing.SystemColors.Control
        Me.btnLicensees.Location = New System.Drawing.Point(384, 8)
        Me.btnLicensees.Name = "btnLicensees"
        Me.btnLicensees.Size = New System.Drawing.Size(24, 24)
        Me.btnLicensees.TabIndex = 260
        Me.btnLicensees.Text = "?"
        '
        'lblLicensee
        '
        Me.lblLicensee.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLicensee.Location = New System.Drawing.Point(88, 8)
        Me.lblLicensee.Name = "lblLicensee"
        Me.lblLicensee.Size = New System.Drawing.Size(56, 16)
        Me.lblLicensee.TabIndex = 262
        Me.lblLicensee.Text = "Licensee:"
        '
        'txtLicensee
        '
        Me.txtLicensee.AcceptsTab = True
        Me.txtLicensee.AutoSize = False
        Me.txtLicensee.Location = New System.Drawing.Point(152, 8)
        Me.txtLicensee.Name = "txtLicensee"
        Me.txtLicensee.ReadOnly = True
        Me.txtLicensee.Size = New System.Drawing.Size(224, 21)
        Me.txtLicensee.TabIndex = 261
        Me.txtLicensee.Text = ""
        Me.txtLicensee.WordWrap = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.txtFacilityID)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.cmbSearchForOwner)
        Me.GroupBox1.Controls.Add(Me.chkSearchforOwner)
        Me.GroupBox1.Controls.Add(Me.cmbFacility)
        Me.GroupBox1.Controls.Add(Me.lblFacility)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 40)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(392, 112)
        Me.GroupBox1.TabIndex = 265
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Facility:"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(40, 80)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(32, 16)
        Me.Label2.TabIndex = 271
        Me.Label2.Text = "(OR)"
        '
        'txtFacilityID
        '
        Me.txtFacilityID.Location = New System.Drawing.Point(144, 80)
        Me.txtFacilityID.Name = "txtFacilityID"
        Me.txtFacilityID.Size = New System.Drawing.Size(104, 20)
        Me.txtFacilityID.TabIndex = 270
        Me.txtFacilityID.Text = ""
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(72, 80)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(64, 16)
        Me.Label1.TabIndex = 269
        Me.Label1.Text = "Facility ID:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbSearchForOwner
        '
        Me.cmbSearchForOwner.Location = New System.Drawing.Point(144, 16)
        Me.cmbSearchForOwner.Name = "cmbSearchForOwner"
        Me.cmbSearchForOwner.Size = New System.Drawing.Size(232, 21)
        Me.cmbSearchForOwner.TabIndex = 268
        '
        'chkSearchforOwner
        '
        Me.chkSearchforOwner.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkSearchforOwner.Location = New System.Drawing.Point(16, 16)
        Me.chkSearchforOwner.Name = "chkSearchforOwner"
        Me.chkSearchforOwner.Size = New System.Drawing.Size(120, 16)
        Me.chkSearchforOwner.TabIndex = 267
        Me.chkSearchforOwner.Tag = "644"
        Me.chkSearchforOwner.Text = "Search for Owner"
        '
        'cmbFacility
        '
        Me.cmbFacility.Location = New System.Drawing.Point(144, 48)
        Me.cmbFacility.Name = "cmbFacility"
        Me.cmbFacility.Size = New System.Drawing.Size(232, 21)
        Me.cmbFacility.TabIndex = 265
        '
        'lblFacility
        '
        Me.lblFacility.Location = New System.Drawing.Point(88, 48)
        Me.lblFacility.Name = "lblFacility"
        Me.lblFacility.Size = New System.Drawing.Size(48, 16)
        Me.lblFacility.TabIndex = 266
        Me.lblFacility.Text = "Facility:"
        Me.lblFacility.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'dtCitationDueDate
        '
        Me.dtCitationDueDate.Checked = False
        Me.dtCitationDueDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtCitationDueDate.Location = New System.Drawing.Point(152, 200)
        Me.dtCitationDueDate.Name = "dtCitationDueDate"
        Me.dtCitationDueDate.ShowCheckBox = True
        Me.dtCitationDueDate.Size = New System.Drawing.Size(104, 20)
        Me.dtCitationDueDate.TabIndex = 266
        '
        'txtCitationDueDate
        '
        Me.txtCitationDueDate.Location = New System.Drawing.Point(24, 200)
        Me.txtCitationDueDate.Name = "txtCitationDueDate"
        Me.txtCitationDueDate.Size = New System.Drawing.Size(112, 16)
        Me.txtCitationDueDate.TabIndex = 267
        Me.txtCitationDueDate.Text = "Citation Due Date:"
        Me.txtCitationDueDate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlLCEBottom
        '
        Me.pnlLCEBottom.Controls.Add(Me.btnSave)
        Me.pnlLCEBottom.Controls.Add(Me.btnClose)
        Me.pnlLCEBottom.Controls.Add(Me.btnCancel)
        Me.pnlLCEBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlLCEBottom.Location = New System.Drawing.Point(0, 406)
        Me.pnlLCEBottom.Name = "pnlLCEBottom"
        Me.pnlLCEBottom.Size = New System.Drawing.Size(432, 32)
        Me.pnlLCEBottom.TabIndex = 268
        '
        'pnlLCEDetails
        '
        Me.pnlLCEDetails.Controls.Add(Me.txtLicensee)
        Me.pnlLCEDetails.Controls.Add(Me.lblLicensee)
        Me.pnlLCEDetails.Controls.Add(Me.btnLicensees)
        Me.pnlLCEDetails.Controls.Add(Me.GroupBox1)
        Me.pnlLCEDetails.Controls.Add(Me.dtCitationDueDate)
        Me.pnlLCEDetails.Controls.Add(Me.txtCitationDueDate)
        Me.pnlLCEDetails.Controls.Add(Me.dtComplianceEventDate)
        Me.pnlLCEDetails.Controls.Add(Me.lblExceptionGrantedDate)
        Me.pnlLCEDetails.Controls.Add(Me.ugCitations)
        Me.pnlLCEDetails.Controls.Add(Me.btnCitationAdd)
        Me.pnlLCEDetails.Controls.Add(Me.btnCitationDelete)
        Me.pnlLCEDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlLCEDetails.Location = New System.Drawing.Point(0, 0)
        Me.pnlLCEDetails.Name = "pnlLCEDetails"
        Me.pnlLCEDetails.Size = New System.Drawing.Size(432, 406)
        Me.pnlLCEDetails.TabIndex = 269
        '
        'LicenseeComplianceEvent
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(432, 438)
        Me.Controls.Add(Me.pnlLCEDetails)
        Me.Controls.Add(Me.pnlLCEBottom)
        Me.Name = "LicenseeComplianceEvent"
        Me.Text = "LicenseeComplianceEvent"
        CType(Me.ugCitations, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox1.ResumeLayout(False)
        Me.pnlLCEBottom.ResumeLayout(False)
        Me.pnlLCEDetails.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub PopulateFacility(ByVal ownerID As Integer, Optional ByVal facID As Integer = 0)
        cmbFacility.DataSource = pLCE.PopulateFacilityName(pLCE.OwnerID)
        cmbFacility.DisplayMember = "FACILITY"
        cmbFacility.ValueMember = "FACILITY_ID"
        If facID = 0 Then
            cmbFacility.SelectedIndex = -1
            If cmbFacility.SelectedIndex <> -1 Then
                cmbFacility.SelectedIndex = -1
            End If
        Else

            UIUtilsGen.SetComboboxItemByValue(cmbFacility, pLCE.FacilityID)
        End If

    End Sub

    Private Sub chkSearchforOwner_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSearchforOwner.CheckedChanged
        Try
            txtFacilityID.Text = ""
            If chkSearchforOwner.Checked Then
                cmbFacility.Enabled = True
                cmbSearchForOwner.Enabled = True
                txtFacilityID.Enabled = False
                If Not cmbFacility.SelectedText Is Nothing Then
                    txtFacilityID.Text = cmbFacility.SelectedValue
                End If
            Else
                cmbFacility.Enabled = False
                cmbSearchForOwner.Enabled = False
                txtFacilityID.Enabled = True
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub LicenseeComplianceEvent_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            bolLoading = True

            chkSearchforOwner.Checked = True

            UIUtilsGen.CreateEmptyFormatDatePicker(dtComplianceEventDate)
            UIUtilsGen.CreateEmptyFormatDatePicker(dtCitationDueDate)

            cmbSearchForOwner.DisplayMember = "o_name"
            cmbSearchForOwner.ValueMember = "o_id"
            'bolCombo = True
            cmbSearchForOwner.DataSource = pLCE.PopulateOwnerName()
            If strFrom = "ADD" Then
                cmbSearchForOwner.SelectedIndex = -1

                'dtComplianceEventDate.Value = Now
                'dtCitationDueDate.Value = Now.AddDays(30)
                'bolCombo = False
                bolLoading = False
            Else
                'dtComplianceEventDate.Value = pLCE.LCEDate
                'dtCitationDueDate.Value = pLCE.CitationDueDate
                'txtLicensee.Text = pLCE.LicenseeName
                'txtFacilityID.Text = pLCE.FacilityID.ToString
                'bolCombo = False
                'cmbSearchForOwner.SelectedValue = pLCE.OwnerID
                'cmbFacility.SelectedValue = pLCE.FacilityID
                'ugCitations.DataSource = pLCE.getLicenseeCitation(pLCE.ID)
                PopulateLCE()
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub PopulateLCE()
        txtLicensee.Text = pLCE.LicenseeName
        txtLicensee.Text = pLCE.LicenseeName
        'bolLoading = False
        cmbSearchForOwner.SelectedValue = pLCE.OwnerID
        ' populate facility
        PopulateFacility(pLCE.OwnerID, pLCE.FacilityID)
        txtFacilityID.Text = pLCE.FacilityID.ToString
        ugCitations.DataSource = pLCE.getLicenseeCitation(pLCE.ID)
        UIUtilsGen.SetDatePickerValue(Me.dtComplianceEventDate, pLCE.LCEDate)
        UIUtilsGen.SetDatePickerValue(Me.dtCitationDueDate, pLCE.CitationDueDate)
        bolLoading = False

    End Sub

    Private Sub btnLicensees_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLicensees.Click
        Try
            Dim Licensees As New CAELicensees(pLCE)
            Licensees.ShowDialog()
            txtLicensee.Text = pLCE.LicenseeName
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub cmbSearchForOwner_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbSearchForOwner.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            'Or Not strFrom = "EDIT" Then
            pLCE.OwnerID = cmbSearchForOwner.SelectedValue
            bolLoading = True
            pLCE.FacilityID = 0
            PopulateFacility(pLCE.OwnerID, pLCE.FacilityID)
            txtFacilityID.Text = String.Empty
            bolLoading = False
            'If cmbFacility.Items.Count = 0 Then
            '    txtFacilityID.Text = String.Empty
            'End If


        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnCitationAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCitationAdd.Click
        Try
            If ugCitations.Rows.Count = 0 Then
                objCitation = New CitationList("LCE", , pLCE)
                objCitation.ShowDialog()
            Else
                MsgBox("Delete existing citation to add new citation")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub



    Private Sub txtFacilityID_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFacilityID.TextChanged
        If bolLoading Then Exit Sub
        Try

        If Not txtFacilityID.Text = "" Then
            pLCE.FacilityID = CInt(txtFacilityID.Text)
            'If cmbFacility.Text <> "System.Data.DataRowView" Then
            '    pLCE.FacilityName = cmbFacility.Text
            'End If
        Else
            pLCE.FacilityID = 0
        End If
        If Not strFrom = "EDIT" Then
            UIUtilsGen.SetDatePickerValue(Me.dtComplianceEventDate, Now.Today)
            UIUtilsGen.SetDatePickerValue(Me.dtCitationDueDate, Now.AddDays(30))
        End If

        Catch ex As Exception
            If UCase(ex.Message).IndexOfAny(UCase("Cast from string ''" + txtFacilityID.Text + "'' to type 'Integer' is not valid.")) >= 0 Then
                MsgBox("Invalid FacilityID")
                Exit Sub
            End If
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub objCitation_evtLCECitationSelected(ByVal drow As Infragistics.Win.UltraWinGrid.UltraGridRow) Handles objCitation.evtLCECitationSelected
        If bolLoading Then Exit Sub
        Try
            citationGrid(drow)
            ugCitations.DataSource = dtCitation
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub citationGrid(ByVal ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow)
        Dim drRow As DataRow
        Try
            If dtCitation.Columns.Count = 0 Then
                dtCitation.Columns.Add("CITATION_ID")
                dtCitation.Columns.Add("CITATION_TEXT")
                dtCitation.Columns.Add("POLICY_PENALTY")
            End If
            drRow = dtCitation.NewRow
            drRow("CITATION_ID") = Integer.Parse(ugRow.Cells("CITATION_ID").Value)
            drRow("CITATION_TEXT") = ugRow.Cells("CITATION_TEXT").Value
            drRow("POLICY_PENALTY") = Integer.Parse(ugRow.Cells("POLICY_PENALTY").Value)
            pLCE.PolicyAmount = Integer.Parse(ugRow.Cells("POLICY_PENALTY").Value)
            dtCitation.Rows.Add(drRow)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnCitationDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCitationDelete.Click
        Try
            If Not ugCitations.ActiveRow Is Nothing Then
                ugCitations.ActiveRow.Delete(False)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub CheckLicenseeRevoke()
        Dim oLCEInfoLocal As MUSTER.Info.LicenseeComplianceEventInfo
        For Each oLCEInfoLocal In pLCE.ColLCEvents.Values
            If oLCEInfoLocal.LicenseeID = pLCE.LicenseeID And (DateDiff(DateInterval.Year, oLCEInfoLocal.LCEDate, Now.Today) < 3) And Not oLCEInfoLocal.Rescinded And Not oLCEInfoLocal.Deleted And oLCEInfoLocal.ID > 0 Then
                If strFrom = "ADD" Then
                    Dim oLicensee As New MUSTER.BusinessLogic.pLicensee
                    oLicensee.Retrieve(oLCEInfoLocal.LicenseeID)
                    oLicensee.STATUS_ID = "REVOKED"
                    If oLicensee.ID <= 0 And oLicensee.ID > -99 Then 'there are 99 licensees with a negative licensee ID
                        oLicensee.CreatedBy = MusterContainer.AppUser.ID
                    Else
                        oLicensee.ModifiedBy = MusterContainer.AppUser.ID
                    End If
                    oLicensee.Save(CType(UIUtilsGen.ModuleID.Company, Integer), MusterContainer.AppUser.UserKey, returnVal, True)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If

                    MsgBox("Licensee " + oLCEInfoLocal.LicenseeName.ToString + " is Revoked")
                    Exit Sub
                End If
            End If
        Next
    End Sub
    Friend Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            If ugCitations.Rows.Count = 0 Then
                MsgBox("Please add a citation and then Save")
                Exit Sub
            End If
            If ValidateData() Then
                If Not strFrom = "EDIT" Then
                    pLCE.LCEStatus = 995
                    pLCE.Status = "NEW".ToUpper
                    pLCE.NextDueDate = Now.Today.AddDays(30)
                    pLCE.PendingLetter = 1173
                    pLCE.PendingLetterTemplateNum = UIUtilsGen.GetLetterTemplateNumPropertyID(UIUtilsGen.LCELetterTemplateNum.NOV)
                End If

                If Date.Compare(CDate(dtCitationDueDate.Text), CDate("01/01/0001")) = 0 Then
                    pLCE.CitationDueDate = Now.AddDays(30)
                End If
                If Date.Compare(CDate(Me.dtComplianceEventDate.Text), CDate("01/01/0001")) = 0 Then
                    pLCE.LCEDate = Now.Today()
                End If
                Me.CheckLicenseeRevoke()

                If pLCE.ID <= 0 Then
                    pLCE.CreatedBy = MusterContainer.AppUser.ID
                Else
                    pLCE.ModifiedBy = MusterContainer.AppUser.ID
                End If

                pLCE.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                MsgBox("LCE Saved Successfully")
                Me.Close()
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub txtLicensee_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLicensee.TextChanged
        If bolLoading Then Exit Sub
        pLCE.LicenseeName = txtLicensee.Text
    End Sub

    Private Sub dtComplianceEventDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtComplianceEventDate.ValueChanged
        If bolLoading Then Exit Sub
        ' pLCE.LCEDate = dtComplianceEventDate.Value
        UIUtilsGen.ToggleDateFormat(Me.dtComplianceEventDate)
        UIUtilsGen.FillDateobjectValues(pLCE.LCEDate, dtComplianceEventDate.Text)

    End Sub

    Private Sub dtCitationDueDate_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtCitationDueDate.ValueChanged
        If bolLoading Then Exit Sub
        'pLCE.CitationDueDate = dtCitationDueDate.Value
        UIUtilsGen.ToggleDateFormat(Me.dtCitationDueDate)
        UIUtilsGen.FillDateobjectValues(pLCE.CitationDueDate, dtCitationDueDate.Text)

    End Sub

    'Private Sub LicenseeComplianceEvent_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
    '    Dim result As DialogResult
    '    Try
    '        If pLCE.colIsDirty Then
    '            result = MsgBox("There are unsaved changes. Do you want to save them now", MsgBoxStyle.YesNoCancel)
    '            If result = DialogResult.Yes Then
    '                If ValidateData() Then
    '                    pLCE.Save()
    '                    e.Cancel = True
    '                End If
    '            ElseIf result = DialogResult.No Then
    '                Exit Sub
    '            ElseIf result = DialogResult.Cancel Then
    '                e.Cancel = True
    '                Exit Sub
    '            End If
    '        End If
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub

    'Private Sub cmbSearchForOwner_SelectedValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbSearchForOwner.SelectedValueChanged
    '    Try
    '        cmbFacility.DisplayMember = "FACILITY_NAME"
    '        cmbFacility.ValueMember = "FACILITY_ID"
    '        cmbFacility.DataSource = pLCE.PopulateFacilityName(cmbSearchForOwner.SelectedValue)
    '        If strFrom = "EDIT" Then
    '            cmbFacility.SelectedValue = pLCE.FacilityID
    '        End If
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub

    Private Sub cmbFacility_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbFacility.SelectedIndexChanged
        If bolLoading Then Exit Sub
        pLCE.FacilityID = cmbFacility.SelectedValue
        txtFacilityID.Text = pLCE.FacilityID.ToString

    End Sub
    Friend Function ValidateData() As Boolean
        Dim errStr As String = ""
        Dim validateSuccess As Boolean = True

        Try
            If txtLicensee.Text <> String.Empty Then
                If txtFacilityID.Text <> String.Empty Then
                    'If Date.Compare(oLCEInfo.CitationDueDate, CDate("01/01/0001")) = 0 Then
                    '    errStr += "Citation Due Date cannot be empty" + vbCrLf
                    '    validateSuccess = False
                    'Else
                    '    validateSuccess = True
                    'End If
                Else
                    errStr += "FacilityID cannot be empty" + vbCrLf
                    validateSuccess = False
                End If
            Else
                errStr += "LicenseeID cannot be empty" + vbCrLf
                validateSuccess = False
            End If


            If errStr.Length > 0 Or Not validateSuccess Then
                MsgBox(errStr)
            End If
            Return validateSuccess
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Sub ClearData()
        bolLoading = True
        UIUtilsGen.ClearFields(pnlLCEDetails)
        If dtCitation.Rows.Count > 0 Then
            dtCitation.Clear()
        End If
        cmbFacility.SelectedIndex = -1
        cmbSearchForOwner.SelectedIndex = -1
        bolLoading = False
    End Sub
    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            If Not pLCE Is Nothing Then
                pLCE.Reset()
                ClearData()
                If pLCE.ID > 0 Then
                    bolLoading = True
                    cmbFacility.Enabled = True
                    PopulateLCE()
                    cmbFacility.Enabled = False
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

End Class
