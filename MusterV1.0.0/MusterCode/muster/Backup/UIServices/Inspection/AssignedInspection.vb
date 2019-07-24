Public Class AssignedInspection
    Inherits System.Windows.Forms.Form
#Region "User Defined Variables"
    Private WithEvents oInspection As MUSTER.BusinessLogic.pInspection
    Private ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
    Private bolLoading As Boolean = False
    Friend CallingForm As Form
    Dim returnVal As String = String.Empty
    Dim bolReadOnly As Boolean
#End Region
#Region "Windows Form Designer generated code "

    Public Sub New(ByRef row As Infragistics.Win.UltraWinGrid.UltraGridRow, ByRef oInspec As MUSTER.BusinessLogic.pInspection, ByVal isReadOnly As Boolean)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        'Add any initialization after the InitializeComponent() call
        Cursor.Current = Cursors.AppStarting
        bolLoading = True
        oInspection = oInspec
        ugRow = row
        bolReadOnly = isReadOnly
        Me.LoadAssignedInspection(ugRow)
        bolLoading = False
        Cursor.Current = Cursors.Default
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
    Friend WithEvents pnlAssignedInspecDetails As System.Windows.Forms.Panel
    Friend WithEvents txtAdminComments As System.Windows.Forms.TextBox
    Friend WithEvents txtInspectorComments As System.Windows.Forms.TextBox
    Friend WithEvents lblInspectorComments As System.Windows.Forms.Label
    Friend WithEvents lblCompleted As System.Windows.Forms.Label
    Friend WithEvents lblAssigned As System.Windows.Forms.Label
    Friend WithEvents lblAssignedValue As System.Windows.Forms.Label
    Friend WithEvents lblAdminComments As System.Windows.Forms.Label
    Friend WithEvents lblInspectionTypeValue As System.Windows.Forms.Label
    Friend WithEvents lblFacilityValue As System.Windows.Forms.Label
    Friend WithEvents lblOwnerValue As System.Windows.Forms.Label
    Friend WithEvents lblOwnerPhoneValue As System.Windows.Forms.Label
    Friend WithEvents lblCityValue As System.Windows.Forms.Label
    Friend WithEvents lblStreetValue As System.Windows.Forms.Label
    Friend WithEvents lblOwner As System.Windows.Forms.Label
    Friend WithEvents lblOwnerPhone As System.Windows.Forms.Label
    Friend WithEvents lblStreet As System.Windows.Forms.Label
    Friend WithEvents lblInspectionType As System.Windows.Forms.Label
    Friend WithEvents lblCity As System.Windows.Forms.Label
    Friend WithEvents lblFacility As System.Windows.Forms.Label
    Friend WithEvents pnlAssignedInspecBottom As System.Windows.Forms.Panel
    Friend WithEvents btnPrint As System.Windows.Forms.Button
    Friend WithEvents btnSubmitToCE As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents lblCountyValue As System.Windows.Forms.Label
    Friend WithEvents dtPickCompleted As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblCounty As System.Windows.Forms.Label
    Friend WithEvents lblInspectedBy As System.Windows.Forms.Label
    Friend WithEvents lblInspectedByValue As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlAssignedInspecDetails = New System.Windows.Forms.Panel
        Me.lblInspectedByValue = New System.Windows.Forms.Label
        Me.lblInspectedBy = New System.Windows.Forms.Label
        Me.dtPickCompleted = New System.Windows.Forms.DateTimePicker
        Me.lblCounty = New System.Windows.Forms.Label
        Me.lblCountyValue = New System.Windows.Forms.Label
        Me.txtAdminComments = New System.Windows.Forms.TextBox
        Me.txtInspectorComments = New System.Windows.Forms.TextBox
        Me.lblInspectorComments = New System.Windows.Forms.Label
        Me.lblCompleted = New System.Windows.Forms.Label
        Me.lblAssigned = New System.Windows.Forms.Label
        Me.lblAssignedValue = New System.Windows.Forms.Label
        Me.lblAdminComments = New System.Windows.Forms.Label
        Me.lblInspectionTypeValue = New System.Windows.Forms.Label
        Me.lblFacilityValue = New System.Windows.Forms.Label
        Me.lblOwnerValue = New System.Windows.Forms.Label
        Me.lblOwnerPhoneValue = New System.Windows.Forms.Label
        Me.lblCityValue = New System.Windows.Forms.Label
        Me.lblStreetValue = New System.Windows.Forms.Label
        Me.lblOwner = New System.Windows.Forms.Label
        Me.lblOwnerPhone = New System.Windows.Forms.Label
        Me.lblStreet = New System.Windows.Forms.Label
        Me.lblInspectionType = New System.Windows.Forms.Label
        Me.lblCity = New System.Windows.Forms.Label
        Me.lblFacility = New System.Windows.Forms.Label
        Me.pnlAssignedInspecBottom = New System.Windows.Forms.Panel
        Me.btnPrint = New System.Windows.Forms.Button
        Me.btnSubmitToCE = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.pnlAssignedInspecDetails.SuspendLayout()
        Me.pnlAssignedInspecBottom.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlAssignedInspecDetails
        '
        Me.pnlAssignedInspecDetails.AutoScroll = True
        Me.pnlAssignedInspecDetails.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlAssignedInspecDetails.Controls.Add(Me.lblInspectedByValue)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.lblInspectedBy)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.dtPickCompleted)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.lblCounty)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.lblCountyValue)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.txtAdminComments)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.txtInspectorComments)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.lblInspectorComments)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.lblCompleted)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.lblAssigned)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.lblAssignedValue)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.lblAdminComments)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.lblInspectionTypeValue)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.lblFacilityValue)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.lblOwnerValue)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.lblOwnerPhoneValue)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.lblCityValue)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.lblStreetValue)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.lblOwner)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.lblOwnerPhone)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.lblStreet)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.lblInspectionType)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.lblCity)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.lblFacility)
        Me.pnlAssignedInspecDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlAssignedInspecDetails.Location = New System.Drawing.Point(0, 0)
        Me.pnlAssignedInspecDetails.Name = "pnlAssignedInspecDetails"
        Me.pnlAssignedInspecDetails.Size = New System.Drawing.Size(400, 438)
        Me.pnlAssignedInspecDetails.TabIndex = 0
        '
        'lblInspectedByValue
        '
        Me.lblInspectedByValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblInspectedByValue.Location = New System.Drawing.Point(120, 216)
        Me.lblInspectedByValue.Name = "lblInspectedByValue"
        Me.lblInspectedByValue.Size = New System.Drawing.Size(134, 28)
        Me.lblInspectedByValue.TabIndex = 23
        '
        'lblInspectedBy
        '
        Me.lblInspectedBy.Location = New System.Drawing.Point(24, 216)
        Me.lblInspectedBy.Name = "lblInspectedBy"
        Me.lblInspectedBy.Size = New System.Drawing.Size(88, 17)
        Me.lblInspectedBy.TabIndex = 22
        Me.lblInspectedBy.Text = "Inspected By:"
        Me.lblInspectedBy.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'dtPickCompleted
        '
        Me.dtPickCompleted.Checked = False
        Me.dtPickCompleted.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtPickCompleted.Location = New System.Drawing.Point(120, 336)
        Me.dtPickCompleted.Name = "dtPickCompleted"
        Me.dtPickCompleted.ShowCheckBox = True
        Me.dtPickCompleted.Size = New System.Drawing.Size(120, 20)
        Me.dtPickCompleted.TabIndex = 0
        '
        'lblCounty
        '
        Me.lblCounty.Location = New System.Drawing.Point(70, 157)
        Me.lblCounty.Name = "lblCounty"
        Me.lblCounty.Size = New System.Drawing.Size(43, 17)
        Me.lblCounty.TabIndex = 16
        Me.lblCounty.Text = "County:"
        Me.lblCounty.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblCountyValue
        '
        Me.lblCountyValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCountyValue.Location = New System.Drawing.Point(121, 151)
        Me.lblCountyValue.Name = "lblCountyValue"
        Me.lblCountyValue.Size = New System.Drawing.Size(132, 30)
        Me.lblCountyValue.TabIndex = 7
        '
        'txtAdminComments
        '
        Me.txtAdminComments.Location = New System.Drawing.Point(120, 272)
        Me.txtAdminComments.Multiline = True
        Me.txtAdminComments.Name = "txtAdminComments"
        Me.txtAdminComments.ReadOnly = True
        Me.txtAdminComments.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtAdminComments.Size = New System.Drawing.Size(264, 61)
        Me.txtAdminComments.TabIndex = 10
        Me.txtAdminComments.Text = ""
        '
        'txtInspectorComments
        '
        Me.txtInspectorComments.Location = New System.Drawing.Point(120, 368)
        Me.txtInspectorComments.Multiline = True
        Me.txtInspectorComments.Name = "txtInspectorComments"
        Me.txtInspectorComments.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtInspectorComments.Size = New System.Drawing.Size(263, 61)
        Me.txtInspectorComments.TabIndex = 1
        Me.txtInspectorComments.Text = ""
        '
        'lblInspectorComments
        '
        Me.lblInspectorComments.Location = New System.Drawing.Point(8, 368)
        Me.lblInspectorComments.Name = "lblInspectorComments"
        Me.lblInspectorComments.Size = New System.Drawing.Size(112, 17)
        Me.lblInspectorComments.TabIndex = 21
        Me.lblInspectorComments.Text = "Inspector Comments:"
        Me.lblInspectorComments.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblCompleted
        '
        Me.lblCompleted.Location = New System.Drawing.Point(48, 336)
        Me.lblCompleted.Name = "lblCompleted"
        Me.lblCompleted.Size = New System.Drawing.Size(64, 17)
        Me.lblCompleted.TabIndex = 20
        Me.lblCompleted.Text = "Completed:"
        Me.lblCompleted.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblAssigned
        '
        Me.lblAssigned.Location = New System.Drawing.Point(56, 248)
        Me.lblAssigned.Name = "lblAssigned"
        Me.lblAssigned.Size = New System.Drawing.Size(56, 17)
        Me.lblAssigned.TabIndex = 18
        Me.lblAssigned.Text = "Assigned:"
        Me.lblAssigned.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblAssignedValue
        '
        Me.lblAssignedValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblAssignedValue.Location = New System.Drawing.Point(120, 248)
        Me.lblAssignedValue.Name = "lblAssignedValue"
        Me.lblAssignedValue.Size = New System.Drawing.Size(132, 23)
        Me.lblAssignedValue.TabIndex = 9
        '
        'lblAdminComments
        '
        Me.lblAdminComments.Location = New System.Drawing.Point(16, 272)
        Me.lblAdminComments.Name = "lblAdminComments"
        Me.lblAdminComments.Size = New System.Drawing.Size(100, 17)
        Me.lblAdminComments.TabIndex = 19
        Me.lblAdminComments.Text = "Admin Comments:"
        Me.lblAdminComments.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblInspectionTypeValue
        '
        Me.lblInspectionTypeValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblInspectionTypeValue.Location = New System.Drawing.Point(121, 182)
        Me.lblInspectionTypeValue.Name = "lblInspectionTypeValue"
        Me.lblInspectionTypeValue.Size = New System.Drawing.Size(132, 30)
        Me.lblInspectionTypeValue.TabIndex = 8
        '
        'lblFacilityValue
        '
        Me.lblFacilityValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblFacilityValue.Location = New System.Drawing.Point(121, 3)
        Me.lblFacilityValue.Name = "lblFacilityValue"
        Me.lblFacilityValue.Size = New System.Drawing.Size(264, 30)
        Me.lblFacilityValue.TabIndex = 2
        '
        'lblOwnerValue
        '
        Me.lblOwnerValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblOwnerValue.Location = New System.Drawing.Point(121, 34)
        Me.lblOwnerValue.Name = "lblOwnerValue"
        Me.lblOwnerValue.Size = New System.Drawing.Size(264, 30)
        Me.lblOwnerValue.TabIndex = 3
        '
        'lblOwnerPhoneValue
        '
        Me.lblOwnerPhoneValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblOwnerPhoneValue.Location = New System.Drawing.Point(121, 65)
        Me.lblOwnerPhoneValue.Name = "lblOwnerPhoneValue"
        Me.lblOwnerPhoneValue.Size = New System.Drawing.Size(132, 23)
        Me.lblOwnerPhoneValue.TabIndex = 4
        '
        'lblCityValue
        '
        Me.lblCityValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCityValue.Location = New System.Drawing.Point(121, 120)
        Me.lblCityValue.Name = "lblCityValue"
        Me.lblCityValue.Size = New System.Drawing.Size(132, 30)
        Me.lblCityValue.TabIndex = 6
        '
        'lblStreetValue
        '
        Me.lblStreetValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblStreetValue.Location = New System.Drawing.Point(121, 89)
        Me.lblStreetValue.Name = "lblStreetValue"
        Me.lblStreetValue.Size = New System.Drawing.Size(264, 30)
        Me.lblStreetValue.TabIndex = 5
        '
        'lblOwner
        '
        Me.lblOwner.Location = New System.Drawing.Point(69, 40)
        Me.lblOwner.Name = "lblOwner"
        Me.lblOwner.Size = New System.Drawing.Size(44, 17)
        Me.lblOwner.TabIndex = 12
        Me.lblOwner.Text = "Owner:"
        Me.lblOwner.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblOwnerPhone
        '
        Me.lblOwnerPhone.Location = New System.Drawing.Point(37, 68)
        Me.lblOwnerPhone.Name = "lblOwnerPhone"
        Me.lblOwnerPhone.Size = New System.Drawing.Size(76, 17)
        Me.lblOwnerPhone.TabIndex = 13
        Me.lblOwnerPhone.Text = "Owner Phone:"
        Me.lblOwnerPhone.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblStreet
        '
        Me.lblStreet.Location = New System.Drawing.Point(75, 95)
        Me.lblStreet.Name = "lblStreet"
        Me.lblStreet.Size = New System.Drawing.Size(38, 17)
        Me.lblStreet.TabIndex = 14
        Me.lblStreet.Text = "Street:"
        Me.lblStreet.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblInspectionType
        '
        Me.lblInspectionType.Location = New System.Drawing.Point(25, 188)
        Me.lblInspectionType.Name = "lblInspectionType"
        Me.lblInspectionType.Size = New System.Drawing.Size(88, 17)
        Me.lblInspectionType.TabIndex = 17
        Me.lblInspectionType.Text = "Inspection Type:"
        Me.lblInspectionType.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblCity
        '
        Me.lblCity.Location = New System.Drawing.Point(83, 126)
        Me.lblCity.Name = "lblCity"
        Me.lblCity.Size = New System.Drawing.Size(30, 17)
        Me.lblCity.TabIndex = 15
        Me.lblCity.Text = "City:"
        Me.lblCity.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblFacility
        '
        Me.lblFacility.Location = New System.Drawing.Point(67, 9)
        Me.lblFacility.Name = "lblFacility"
        Me.lblFacility.Size = New System.Drawing.Size(46, 17)
        Me.lblFacility.TabIndex = 11
        Me.lblFacility.Text = "Facility:"
        Me.lblFacility.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'pnlAssignedInspecBottom
        '
        Me.pnlAssignedInspecBottom.Controls.Add(Me.btnPrint)
        Me.pnlAssignedInspecBottom.Controls.Add(Me.btnSubmitToCE)
        Me.pnlAssignedInspecBottom.Controls.Add(Me.btnCancel)
        Me.pnlAssignedInspecBottom.Controls.Add(Me.btnSave)
        Me.pnlAssignedInspecBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlAssignedInspecBottom.Location = New System.Drawing.Point(0, 438)
        Me.pnlAssignedInspecBottom.Name = "pnlAssignedInspecBottom"
        Me.pnlAssignedInspecBottom.Size = New System.Drawing.Size(400, 40)
        Me.pnlAssignedInspecBottom.TabIndex = 1
        '
        'btnPrint
        '
        Me.btnPrint.Location = New System.Drawing.Point(285, 8)
        Me.btnPrint.Name = "btnPrint"
        Me.btnPrint.TabIndex = 3
        Me.btnPrint.Text = "Print"
        '
        'btnSubmitToCE
        '
        Me.btnSubmitToCE.Location = New System.Drawing.Point(189, 8)
        Me.btnSubmitToCE.Name = "btnSubmitToCE"
        Me.btnSubmitToCE.Size = New System.Drawing.Size(88, 23)
        Me.btnSubmitToCE.TabIndex = 2
        Me.btnSubmitToCE.Text = "Submit to C&&E"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(109, 8)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "Cancel"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(29, 8)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.TabIndex = 0
        Me.btnSave.Text = "Save"
        '
        'AssignedInspection
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(400, 478)
        Me.Controls.Add(Me.pnlAssignedInspecDetails)
        Me.Controls.Add(Me.pnlAssignedInspecBottom)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "AssignedInspection"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Assigned Inspection"
        Me.pnlAssignedInspecDetails.ResumeLayout(False)
        Me.pnlAssignedInspecBottom.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region "UI Support Routines"
    Private Sub LoadAssignedInspection(ByVal ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow)
        Try
            Me.lblFacilityValue.Text = "#" + ugRow.Cells("FACILITY ID").Value.ToString + " - " + ugRow.Cells("FACILITY").Value
            Me.lblOwnerValue.Text = ugRow.Cells("OWNER NAME").Value
            Me.lblOwnerPhoneValue.Text = IIf(ugRow.Cells("OWNER PHONE").Value Is System.DBNull.Value, String.Empty, ugRow.Cells("OWNER PHONE").Value)
            Me.lblStreetValue.Text = IIf(ugRow.Cells("ADDRESS_LINE_ONE").Value Is System.DBNull.Value, String.Empty, ugRow.Cells("ADDRESS_LINE_ONE").Value)
            Me.lblCityValue.Text = IIf(ugRow.Cells("CITY").Value Is System.DBNull.Value, String.Empty, ugRow.Cells("CITY").Value)
            Me.lblCountyValue.Text = IIf(ugRow.Cells("COUNTY").Value Is System.DBNull.Value, String.Empty, ugRow.Cells("COUNTY").Value)
            Me.lblInspectionTypeValue.Text = IIf(ugRow.Cells("INSPECTION TYPE").Value Is System.DBNull.Value, String.Empty, ugRow.Cells("INSPECTION TYPE").Value)
            Me.lblInspectedByValue.Text = IIf(ugRow.Cells("INSPECTED BY").Value Is System.DBNull.Value, String.Empty, ugRow.Cells("INSPECTED BY").Value)
            Me.lblAssignedValue.Text = IIf(ugRow.Cells("ASSIGNED DATE").Value Is System.DBNull.Value, String.Empty, ugRow.Cells("ASSIGNED DATE").Value)
            Me.txtAdminComments.Text = IIf(ugRow.Cells("ADMIN COMMENTS").Value Is System.DBNull.Value, String.Empty, ugRow.Cells("ADMIN COMMENTS").Value)
            Me.txtInspectorComments.Text = IIf(ugRow.Cells("INSPECTOR COMMENTS").Value Is System.DBNull.Value, String.Empty, ugRow.Cells("INSPECTOR COMMENTS").Value)
            If ugRow.Cells("COMPLETED").Value Is System.DBNull.Value Then
                UIUtilsGen.CreateEmptyFormatDatePicker(dtPickCompleted)
            Else
                UIUtilsGen.SetDatePickerValue(dtPickCompleted, ugRow.Cells("COMPLETED").Value)
            End If
            btnSave.Enabled = IIf(oInspection.IsDirty, True, False)
            btnSubmitToCE.Enabled = IIf(ugRow.Cells("SUBMITTED").Value Is DBNull.Value, True, False)
            Me.dtPickCompleted.Focus()
            If bolReadOnly Then
                btnSave.Enabled = False
                btnSubmitToCE.Enabled = False
                btnPrint.Enabled = False
                dtPickCompleted.Enabled = False
                txtInspectorComments.ReadOnly = True
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Friend Sub PrintAssignedInspection()
        Dim oLetter As New Reg_Letters
        Dim strTitle As String
        Dim strData As String
        Dim rows, cols As Int16
        Dim documentInfo As String = String.Empty
        Try
            ' if there is an open oce, attach the letter
            Dim oce As New MUSTER.BusinessLogic.pOwnerComplianceEvent
            documentInfo = oce.GetOpenOCEDocumentInfoforOwner(oInspection.OwnerID)

            strTitle = "Assigned Inspection for " + lblFacility.Text + lblFacilityValue.Text
            rows = 11
            cols = 2
            strData = lblFacility.Text + "|" + _
                        lblFacilityValue.Text + "|" + _
                        lblOwner.Text + "|" + _
                        lblOwnerValue.Text + "|" + _
                        lblOwnerPhone.Text + "|" + _
                        lblOwnerPhoneValue.Text + "|" + _
                        lblStreet.Text + "|" + _
                        lblStreetValue.Text + "|" + _
                        lblCity.Text + "|" + _
                        lblCityValue.Text + "|" + _
                        lblCounty.Text + "|" + _
                        lblCountyValue.Text + "|" + _
                        lblInspectionType.Text + "|" + _
                        lblInspectionTypeValue.Text + "|" + _
                        lblAssigned.Text + "|" + _
                        lblAssignedValue.Text + "|" + _
                        lblAdminComments.Text + "|" + _
                        txtAdminComments.Text + "|" + _
                        lblCompleted.Text + "|" + _
                        UIUtilsGen.GetDatePickerValue(dtPickCompleted).ToString + "|" + _
                        lblInspectorComments.Text + "|" + _
                        txtInspectorComments.Text
            oLetter.GenerateGenericLetter(UIUtilsGen.ModuleID.Inspection, strTitle, strData, cols, False, "INSPECTION", CType(ugRow.Cells("FACILITY ID").Value, Int64), 6, "INS_AssignedInspection_", "Assigned Inspection", , Word.WdTableFormat.wdTableFormatGrid2, True, False, True, False, False, False, True, False, True, documentInfo)
            MsgBox("Assigned Inspection Printed")
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Function ValidateInspection() As Boolean
        Try
            If Date.Compare(oInspection.Completed, CDate("01/01/0001")) = 0 Then
                MsgBox("Required: " + lblCompleted.Text.Trim.TrimEnd(":"))
                Return False
            End If
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region
#Region "UI Control Events"
    Private Sub dtPickCompleted_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtPickCompleted.ValueChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(Me.dtPickCompleted)
            UIUtilsGen.FillDateobjectValues(oInspection.Completed, dtPickCompleted.Text)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtInspectorComments_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtInspectorComments.TextChanged
        If bolLoading Then Exit Sub
        Try
            oInspection.InspectorComments = txtInspectorComments.Text
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnSave_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            If ValidateInspection() Then
                If oInspection.ID <= 0 Then
                    oInspection.CreatedBy = MusterContainer.AppUser.ID
                Else
                    oInspection.ModifiedBy = MusterContainer.AppUser.ID
                End If

                Dim bolSaveSuccess As Boolean = False
                bolSaveSuccess = oInspection.Save(CType(UIUtilsGen.ModuleID.Inspection, Integer), MusterContainer.AppUser.UserKey, returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                If bolSaveSuccess Then
                    ugRow.Cells("COMPLETED").Value = oInspection.Completed
                    ugRow.Cells("INSPECTOR COMMENTS").Value = oInspection.InspectorComments
                    MsgBox("Inspection Saved")
                End If


            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            oInspection.Reset()
            Me.Close()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnSubmitToCE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSubmitToCE.Click
        Try
            If ValidateInspection() Then
                oInspection.SubmittedDate = Now.Date
                If oInspection.ID <= 0 Then
                    oInspection.CreatedBy = MusterContainer.AppUser.ID
                Else
                    oInspection.ModifiedBy = MusterContainer.AppUser.ID
                End If

                Dim bolSuccess As Boolean = False
                bolSuccess = oInspection.Save(CType(UIUtilsGen.ModuleID.Inspection, Integer), MusterContainer.AppUser.UserKey, returnVal)

                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If

                If bolSuccess Then
                    MsgBox("Inspection submitted to C&E")
                End If

                CallingForm.Tag = "1"
                Me.Close()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Try
            Cursor.Current = Cursors.AppStarting
            PrintAssignedInspection()
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)
        Try
            If oInspection.IsDirty And Not bolReadOnly Then
                Dim Results As Long = MsgBox("There are unsaved changes.  Do you wish to save the changes?", MsgBoxStyle.Question & MsgBoxStyle.YesNoCancel, "Data Changed")
                If Results = MsgBoxResult.Yes Then
                    btnSave.PerformClick()
                ElseIf Results = MsgBoxResult.Cancel Then
                    e.Cancel = True
                    Exit Sub
                Else
                    oInspection.Reset()
                End If
            End If
            oInspection.Remove(oInspection.ID)
            oInspection = New MUSTER.BusinessLogic.pInspection
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#Region "External Events"
    Private Sub InspectionChanged(ByVal bolValue As Boolean) Handles oInspection.evtInspectionChanged
        If bolLoading Then Exit Sub
        Try
            btnSave.Enabled = bolValue Or oInspection.IsDirty
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    'Private Sub ColChanged(ByVal bolValue As Boolean) Handles oInspection.evtColChanged
    '    If bolLoading Then Exit Sub
    '    Try
    '        btnSave.Enabled = bolValue Or oInspection.IsDirty
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
#End Region

End Class
