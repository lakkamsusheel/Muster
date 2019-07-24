Public Class CandEAssignedInspection
    Inherits System.Windows.Forms.Form
#Region "Private Member Variables"
    Friend CallingForm As Form
    Private bolLoading As Boolean = False
    Private nInspectionID As Integer = 0

    Private pFCE As New MUSTER.BusinessLogic.pFacilityComplianceEvent
    Private WithEvents oInspection As MUSTER.BusinessLogic.pInspection
    Dim returnVal As String = String.Empty
#End Region

#Region " Windows Form Designer generated code "

    'Public Sub New(Optional ByVal OwnerID As Integer = 0, Optional ByVal inspID As Integer = 0, Optional ByVal facID As Int64 = 0, Optional ByVal FCEID As Int64 = 0, Optional ByVal ownName As String = "")
    Public Sub New(Optional ByVal inspID As Integer = 0)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        nInspectionID = inspID
        oInspection = New MUSTER.BusinessLogic.pInspection
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
    Friend WithEvents lblAdminComments As System.Windows.Forms.Label
    Friend WithEvents lblInspectionType As System.Windows.Forms.Label
    Friend WithEvents pnlAssignedInspecBottom As System.Windows.Forms.Panel
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents lblInspection As System.Windows.Forms.Label
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnCreateFCE As System.Windows.Forms.Button
    Friend WithEvents txtSubmitted As System.Windows.Forms.TextBox
    Friend WithEvents lblSubmitted As System.Windows.Forms.Label
    Friend WithEvents cmbInspectionType As System.Windows.Forms.ComboBox
    Friend WithEvents lblAssigned As System.Windows.Forms.Label
    Friend WithEvents txtCompleted As System.Windows.Forms.TextBox
    Friend WithEvents lblCompleted As System.Windows.Forms.Label
    Friend WithEvents cmbInspector As System.Windows.Forms.ComboBox
    Public WithEvents dtAssigned As System.Windows.Forms.DateTimePicker
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtFacilityID As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cmbSearchForOwner As System.Windows.Forms.ComboBox
    Friend WithEvents chkSearchforOwner As System.Windows.Forms.CheckBox
    Friend WithEvents cmbFacility As System.Windows.Forms.ComboBox
    Friend WithEvents lblFacility As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlAssignedInspecDetails = New System.Windows.Forms.Panel
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtFacilityID = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.cmbSearchForOwner = New System.Windows.Forms.ComboBox
        Me.chkSearchforOwner = New System.Windows.Forms.CheckBox
        Me.cmbFacility = New System.Windows.Forms.ComboBox
        Me.lblFacility = New System.Windows.Forms.Label
        Me.dtAssigned = New System.Windows.Forms.DateTimePicker
        Me.txtCompleted = New System.Windows.Forms.TextBox
        Me.lblCompleted = New System.Windows.Forms.Label
        Me.lblAssigned = New System.Windows.Forms.Label
        Me.cmbInspectionType = New System.Windows.Forms.ComboBox
        Me.cmbInspector = New System.Windows.Forms.ComboBox
        Me.lblInspection = New System.Windows.Forms.Label
        Me.pnlAssignedInspecBottom = New System.Windows.Forms.Panel
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnCreateFCE = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.txtAdminComments = New System.Windows.Forms.TextBox
        Me.txtInspectorComments = New System.Windows.Forms.TextBox
        Me.txtSubmitted = New System.Windows.Forms.TextBox
        Me.lblInspectorComments = New System.Windows.Forms.Label
        Me.lblSubmitted = New System.Windows.Forms.Label
        Me.lblAdminComments = New System.Windows.Forms.Label
        Me.lblInspectionType = New System.Windows.Forms.Label
        Me.pnlAssignedInspecDetails.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.pnlAssignedInspecBottom.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlAssignedInspecDetails
        '
        Me.pnlAssignedInspecDetails.AutoScroll = True
        Me.pnlAssignedInspecDetails.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlAssignedInspecDetails.Controls.Add(Me.GroupBox1)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.dtAssigned)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.txtCompleted)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.lblCompleted)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.lblAssigned)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.cmbInspectionType)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.cmbInspector)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.lblInspection)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.pnlAssignedInspecBottom)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.txtAdminComments)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.txtInspectorComments)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.txtSubmitted)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.lblInspectorComments)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.lblSubmitted)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.lblAdminComments)
        Me.pnlAssignedInspecDetails.Controls.Add(Me.lblInspectionType)
        Me.pnlAssignedInspecDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlAssignedInspecDetails.Location = New System.Drawing.Point(0, 0)
        Me.pnlAssignedInspecDetails.Name = "pnlAssignedInspecDetails"
        Me.pnlAssignedInspecDetails.Size = New System.Drawing.Size(416, 494)
        Me.pnlAssignedInspecDetails.TabIndex = 0
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
        Me.GroupBox1.Location = New System.Drawing.Point(16, 8)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(392, 112)
        Me.GroupBox1.TabIndex = 266
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
        'dtAssigned
        '
        Me.dtAssigned.Checked = False
        Me.dtAssigned.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtAssigned.Location = New System.Drawing.Point(112, 280)
        Me.dtAssigned.Name = "dtAssigned"
        Me.dtAssigned.ShowCheckBox = True
        Me.dtAssigned.Size = New System.Drawing.Size(104, 20)
        Me.dtAssigned.TabIndex = 6
        '
        'txtCompleted
        '
        Me.txtCompleted.Location = New System.Drawing.Point(112, 328)
        Me.txtCompleted.Name = "txtCompleted"
        Me.txtCompleted.ReadOnly = True
        Me.txtCompleted.Size = New System.Drawing.Size(168, 20)
        Me.txtCompleted.TabIndex = 8
        Me.txtCompleted.Text = ""
        '
        'lblCompleted
        '
        Me.lblCompleted.Location = New System.Drawing.Point(40, 328)
        Me.lblCompleted.Name = "lblCompleted"
        Me.lblCompleted.Size = New System.Drawing.Size(64, 17)
        Me.lblCompleted.TabIndex = 45
        Me.lblCompleted.Text = "Completed:"
        Me.lblCompleted.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblAssigned
        '
        Me.lblAssigned.Location = New System.Drawing.Point(48, 280)
        Me.lblAssigned.Name = "lblAssigned"
        Me.lblAssigned.Size = New System.Drawing.Size(56, 16)
        Me.lblAssigned.TabIndex = 43
        Me.lblAssigned.Text = "Assigned:"
        Me.lblAssigned.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'cmbInspectionType
        '
        Me.cmbInspectionType.Location = New System.Drawing.Point(112, 160)
        Me.cmbInspectionType.Name = "cmbInspectionType"
        Me.cmbInspectionType.Size = New System.Drawing.Size(160, 21)
        Me.cmbInspectionType.TabIndex = 4
        '
        'cmbInspector
        '
        Me.cmbInspector.Location = New System.Drawing.Point(112, 136)
        Me.cmbInspector.Name = "cmbInspector"
        Me.cmbInspector.Size = New System.Drawing.Size(160, 21)
        Me.cmbInspector.TabIndex = 3
        '
        'lblInspection
        '
        Me.lblInspection.Location = New System.Drawing.Point(40, 136)
        Me.lblInspection.Name = "lblInspection"
        Me.lblInspection.Size = New System.Drawing.Size(64, 16)
        Me.lblInspection.TabIndex = 34
        Me.lblInspection.Text = "Inspector:"
        Me.lblInspection.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'pnlAssignedInspecBottom
        '
        Me.pnlAssignedInspecBottom.Controls.Add(Me.btnClose)
        Me.pnlAssignedInspecBottom.Controls.Add(Me.btnCreateFCE)
        Me.pnlAssignedInspecBottom.Controls.Add(Me.btnCancel)
        Me.pnlAssignedInspecBottom.Controls.Add(Me.btnSave)
        Me.pnlAssignedInspecBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlAssignedInspecBottom.Location = New System.Drawing.Point(0, 450)
        Me.pnlAssignedInspecBottom.Name = "pnlAssignedInspecBottom"
        Me.pnlAssignedInspecBottom.Size = New System.Drawing.Size(412, 40)
        Me.pnlAssignedInspecBottom.TabIndex = 1
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(200, 8)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.TabIndex = 2
        Me.btnClose.Text = "Close"
        '
        'btnCreateFCE
        '
        Me.btnCreateFCE.Location = New System.Drawing.Point(288, 8)
        Me.btnCreateFCE.Name = "btnCreateFCE"
        Me.btnCreateFCE.Size = New System.Drawing.Size(80, 23)
        Me.btnCreateFCE.TabIndex = 3
        Me.btnCreateFCE.Text = "Create FCE"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(120, 8)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "Cancel"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(40, 8)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.TabIndex = 0
        Me.btnSave.Text = "Save"
        '
        'txtAdminComments
        '
        Me.txtAdminComments.Location = New System.Drawing.Point(112, 192)
        Me.txtAdminComments.Multiline = True
        Me.txtAdminComments.Name = "txtAdminComments"
        Me.txtAdminComments.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtAdminComments.Size = New System.Drawing.Size(264, 80)
        Me.txtAdminComments.TabIndex = 5
        Me.txtAdminComments.Text = ""
        '
        'txtInspectorComments
        '
        Me.txtInspectorComments.Location = New System.Drawing.Point(112, 360)
        Me.txtInspectorComments.Multiline = True
        Me.txtInspectorComments.Name = "txtInspectorComments"
        Me.txtInspectorComments.ReadOnly = True
        Me.txtInspectorComments.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtInspectorComments.Size = New System.Drawing.Size(263, 80)
        Me.txtInspectorComments.TabIndex = 9
        Me.txtInspectorComments.Text = ""
        '
        'txtSubmitted
        '
        Me.txtSubmitted.Location = New System.Drawing.Point(112, 304)
        Me.txtSubmitted.Name = "txtSubmitted"
        Me.txtSubmitted.ReadOnly = True
        Me.txtSubmitted.Size = New System.Drawing.Size(167, 20)
        Me.txtSubmitted.TabIndex = 7
        Me.txtSubmitted.Text = ""
        '
        'lblInspectorComments
        '
        Me.lblInspectorComments.Location = New System.Drawing.Point(24, 360)
        Me.lblInspectorComments.Name = "lblInspectorComments"
        Me.lblInspectorComments.Size = New System.Drawing.Size(80, 32)
        Me.lblInspectorComments.TabIndex = 28
        Me.lblInspectorComments.Text = "Inspector Comments"
        Me.lblInspectorComments.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblSubmitted
        '
        Me.lblSubmitted.Location = New System.Drawing.Point(40, 304)
        Me.lblSubmitted.Name = "lblSubmitted"
        Me.lblSubmitted.Size = New System.Drawing.Size(64, 17)
        Me.lblSubmitted.TabIndex = 27
        Me.lblSubmitted.Text = "Submitted:"
        Me.lblSubmitted.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblAdminComments
        '
        Me.lblAdminComments.Location = New System.Drawing.Point(40, 192)
        Me.lblAdminComments.Name = "lblAdminComments"
        Me.lblAdminComments.Size = New System.Drawing.Size(64, 32)
        Me.lblAdminComments.TabIndex = 23
        Me.lblAdminComments.Text = "Admin Comments"
        Me.lblAdminComments.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblInspectionType
        '
        Me.lblInspectionType.Location = New System.Drawing.Point(16, 160)
        Me.lblInspectionType.Name = "lblInspectionType"
        Me.lblInspectionType.Size = New System.Drawing.Size(88, 17)
        Me.lblInspectionType.TabIndex = 13
        Me.lblInspectionType.Text = "Inspection Type:"
        Me.lblInspectionType.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'CandEAssignedInspection
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(416, 494)
        Me.ControlBox = False
        Me.Controls.Add(Me.pnlAssignedInspecDetails)
        Me.Name = "CandEAssignedInspection"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "CandEAssignedInspection"
        Me.pnlAssignedInspecDetails.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.pnlAssignedInspecBottom.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region "UI Support Routines"
    Private Sub Populate()
        Dim bolLoadingLocal As Boolean = bolLoading
        Try
            bolLoading = True
            If nInspectionID = 0 Then ' ADD
                chkSearchforOwner.Checked = True
                chkSearchforOwner.Enabled = True
                cmbSearchForOwner.Enabled = True
                cmbFacility.Enabled = True
                cmbFacility.DataSource = Nothing
                txtFacilityID.Enabled = False
                txtFacilityID.Text = ""
                cmbInspector.Enabled = False
                cmbInspector.DataSource = Nothing
                cmbInspectionType.Enabled = False
                cmbInspectionType.DataSource = Nothing
                txtAdminComments.Enabled = False
                txtAdminComments.Text = ""
                dtAssigned.Enabled = False
                dtAssigned.Text = ""
                txtSubmitted.Text = ""
                txtCompleted.Text = ""
                txtInspectorComments.Text = ""
                PopulateOwners()
            Else ' EDIT
                oInspection.Retrieve(nInspectionID)

                chkSearchforOwner.Checked = False
                cmbSearchForOwner.Enabled = True
                cmbFacility.Enabled = True
                PopulateOwners(oInspection.OwnerID)
                PopulateFacility(oInspection.OwnerID, oInspection.FacilityID)
                cmbSearchForOwner.Enabled = False
                cmbFacility.Enabled = False

                txtFacilityID.Enabled = True
                txtFacilityID.Text = oInspection.FacilityID.ToString
                cmbInspector.Enabled = True
                PopulateInspectors(oInspection.StaffID)
                cmbInspectionType.Enabled = True
                PopulateInspectionType(oInspection.InspectionType)
                txtAdminComments.Enabled = True
                txtAdminComments.Text = oInspection.AdminComments
                dtAssigned.Enabled = True
                UIUtilsGen.SetDatePickerValue(dtAssigned, oInspection.AssignedDate)
                If Date.Compare(oInspection.SubmittedDate, CDate("01/01/0001")) = 0 Then
                    txtSubmitted.Text = ""
                Else
                    txtSubmitted.Text = oInspection.SubmittedDate.ToShortDateString
                End If
                If Date.Compare(oInspection.Completed, CDate("01/01/0001")) = 0 Then
                    txtCompleted.Text = ""
                Else
                    txtCompleted.Text = oInspection.Completed.ToShortDateString
                End If
                txtInspectorComments.Text = oInspection.InspectorComments
            End If
            EnableSave(oInspection.IsDirty)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            bolLoading = bolLoadingLocal
        End Try
    End Sub
    Private Sub PopulateOwners(Optional ByVal ownID As Integer = 0)
        Try
            cmbSearchForOwner.DataSource = pFCE.GetOwners.Tables(0)
            cmbSearchForOwner.DisplayMember = "o_name"
            cmbSearchForOwner.ValueMember = "o_id"

            oInspection.OwnerID = ownID

            If ownID > 0 Then
                If cmbSearchForOwner.Enabled Then
                    UIUtilsGen.SetComboboxItemByValue(cmbSearchForOwner, ownID)
                Else
                    Dim bolLoadingLocal As Boolean = bolLoading
                    Try
                        bolLoading = True
                        cmbSearchForOwner.Enabled = True
                        UIUtilsGen.SetComboboxItemByValue(cmbSearchForOwner, ownID)
                    Finally
                        cmbSearchForOwner.Enabled = False
                        bolLoading = bolLoadingLocal
                    End Try
                End If
            Else
                cmbSearchForOwner.SelectedIndex = -1
                If cmbSearchForOwner.SelectedIndex <> -1 Then
                    cmbSearchForOwner.SelectedIndex = -1
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub PopulateFacility(ByVal OwnerID As Integer, Optional ByVal facID As Integer = 0)
        Try
            cmbFacility.DataSource = pFCE.GetFacilities(OwnerID).Tables(0)
            cmbFacility.DisplayMember = "FACILITY"
            cmbFacility.ValueMember = "FACILITY_ID"

            oInspection.FacilityID = facID

            If facID > 0 Then
                If cmbFacility.Enabled Then
                    UIUtilsGen.SetComboboxItemByValue(cmbFacility, facID)
                Else
                    Dim bolLoadingLocal As Boolean = bolLoading
                    Try
                        bolLoading = True
                        cmbFacility.Enabled = True
                        UIUtilsGen.SetComboboxItemByValue(cmbFacility, facID)
                    Finally
                        cmbFacility.Enabled = False
                        bolLoading = bolLoadingLocal
                    End Try
                End If
            Else
                cmbFacility.SelectedIndex = -1
                If cmbFacility.SelectedIndex <> -1 Then
                    cmbFacility.SelectedIndex = -1
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub PopulateInspectors(Optional ByVal staffID As Integer = 0)
        Try
            cmbInspector.DataSource = oInspection.GetInspectors.Tables(0)
            cmbInspector.DisplayMember = "USER_NAME"
            cmbInspector.ValueMember = "STAFF_ID"

            oInspection.StaffID = staffID

            If staffID > 0 Then
                UIUtilsGen.SetComboboxItemByValue(cmbInspector, staffID)
            Else
                cmbInspector.SelectedIndex = -1
                If cmbInspector.SelectedIndex <> -1 Then
                    cmbInspector.SelectedIndex = -1
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub PopulateInspectionType(Optional ByVal inspType As Integer = 0)
        Try
            cmbInspectionType.DataSource = oInspection.GetInspectionTypes.Tables(0)
            cmbInspectionType.DisplayMember = "PROPERTY_NAME"
            cmbInspectionType.ValueMember = "PROPERTY_ID"

            oInspection.InspectionType = inspType

            If inspType > 0 Then
                UIUtilsGen.SetComboboxItemByValue(cmbInspectionType, inspType)
            Else
                cmbInspectionType.SelectedIndex = -1
                If cmbInspectionType.SelectedIndex <> -1 Then
                    cmbInspectionType.SelectedIndex = -1
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub EnableSave(ByVal bolValue As Boolean)
        btnSave.Enabled = bolValue
        If Not bolValue Then
            If Date.Compare(oInspection.SubmittedDate, CDate("01/01/0001")) <> 0 Then
                btnCreateFCE.Enabled = True
            Else
                btnCreateFCE.Enabled = False
            End If
        Else
            btnCreateFCE.Enabled = False
        End If
    End Sub
#End Region

#Region "UI Control Events"
    Private Sub chkSearchforOwner_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSearchforOwner.CheckedChanged
        If bolLoading Then Exit Sub
        Dim bolLoadingLocal As Boolean = bolLoading
        Try
            bolLoading = True
            If chkSearchforOwner.Checked Then
                cmbFacility.Enabled = True
                cmbSearchForOwner.Enabled = True
                txtFacilityID.Enabled = False
                txtFacilityID.Text = ""
                PopulateOwners(oInspection.OwnerID)
                PopulateFacility(oInspection.OwnerID, oInspection.FacilityID)
            Else
                cmbFacility.Enabled = False
                cmbSearchForOwner.Enabled = False
                cmbFacility.DataSource = Nothing
                cmbSearchForOwner.DataSource = Nothing

                txtFacilityID.Enabled = True

                ' to get owner_id for a facility_id entered
                If oInspection.FacilityID > 0 Then
                    txtFacilityID.Text = oInspection.FacilityID
                    Dim dt As DataTable = oInspection.GetFacOwner(oInspection.FacilityID).Tables(0)
                    cmbSearchForOwner.DataSource = dt
                    cmbSearchForOwner.DisplayMember = "o_name"
                    cmbSearchForOwner.ValueMember = "o_id"
                    If dt.Rows.Count > 0 Then
                        cmbSearchForOwner.SelectedIndex = 0
                        oInspection.OwnerID = cmbSearchForOwner.SelectedValue
                        PopulateFacility(oInspection.OwnerID, oInspection.FacilityID)
                    Else
                        oInspection.OwnerID = 0
                    End If
                Else
                    txtFacilityID.Text = ""
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            bolLoading = bolLoadingLocal
        End Try
    End Sub
    Private Sub cmbSearchForOwner_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbSearchForOwner.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Dim bolLoadingLocal As Boolean = bolLoading
        Try
            oInspection.OwnerID = UIUtilsGen.GetComboBoxValue(cmbSearchForOwner)
            bolLoading = True
            PopulateFacility(oInspection.OwnerID, oInspection.FacilityID)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            bolLoading = bolLoadingLocal
        End Try
    End Sub
    Private Sub cmbFacility_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbFacility.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Dim bolLoadingLocal As Boolean = bolLoading
        Try
            oInspection.FacilityID = UIUtilsGen.GetComboBoxValue(cmbFacility)
            oInspection.StaffID = oInspection.GetCAEFacilityAssignedTo(oInspection.FacilityID)

            bolLoading = True

            cmbInspector.Enabled = True
            cmbInspectionType.Enabled = True
            PopulateInspectors(oInspection.StaffID)
            PopulateInspectionType(oInspection.InspectionType)

            txtAdminComments.Enabled = True
            dtAssigned.Enabled = True
            If Date.Compare(oInspection.AssignedDate, CDate("01/01/0001")) = 0 Then
                UIUtilsGen.SetDatePickerValue(dtAssigned, Today.Date)
                oInspection.AssignedDate = Today.Date
            Else
                UIUtilsGen.SetDatePickerValue(dtAssigned, oInspection.AssignedDate)
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            bolLoading = bolLoadingLocal
        End Try
    End Sub
    Private Sub txtFacilityID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFacilityID.LostFocus
        If bolLoading Then Exit Sub
        If txtFacilityID.Text = "0" Then Exit Sub
        If txtFacilityID.Text = oInspection.FacilityID.ToString Then Exit Sub

        If Not IsNumeric(txtFacilityID.Text) Then
            MsgBox("Facility ID must be numeric")
            txtFacilityID.Text = oInspection.FacilityID.ToString
            Exit Sub
        End If

        Dim bolLoadingLocal As Boolean = bolLoading
        Try
            bolLoading = True
            ' to get owner_id for a facility_id entered
            Dim dt As DataTable = oInspection.GetFacOwner(CType(txtFacilityID.Text, Integer)).Tables(0)
            cmbSearchForOwner.DataSource = dt
            cmbSearchForOwner.DisplayMember = "o_name"
            cmbSearchForOwner.ValueMember = "o_id"
            If dt.Rows.Count > 0 Then
                oInspection.FacilityID = CType(txtFacilityID.Text, Integer)
                cmbSearchForOwner.SelectedIndex = 0
                oInspection.OwnerID = cmbSearchForOwner.SelectedValue
                PopulateFacility(oInspection.OwnerID, oInspection.FacilityID)
            Else
                MsgBox("Owner does not exists for selected Facility")
                txtFacilityID.Text = oInspection.FacilityID.ToString
                oInspection.OwnerID = 0
                cmbFacility.DataSource = Nothing
            End If

            oInspection.StaffID = oInspection.GetCAEFacilityAssignedTo(oInspection.FacilityID)

            cmbInspector.Enabled = True
            cmbInspectionType.Enabled = True
            PopulateInspectors(oInspection.StaffID)
            PopulateInspectionType(oInspection.InspectionType)

            txtAdminComments.Enabled = True
            dtAssigned.Enabled = True
            UIUtilsGen.SetDatePickerValue(dtAssigned, Today.Date)
            oInspection.AssignedDate = Today.Date
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            bolLoading = bolLoadingLocal
        End Try
    End Sub
    Private Sub cmbInspector_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbInspector.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            oInspection.StaffID = UIUtilsGen.GetComboBoxValue(cmbInspector)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub cmbInspectionType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbInspectionType.SelectedIndexChanged
        If bolLoading Then Exit Sub
        Try
            oInspection.InspectionType = UIUtilsGen.GetComboBoxValue(cmbInspectionType)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub txtAdminComments_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAdminComments.TextChanged
        If bolLoading Then Exit Sub
        Try
            oInspection.AdminComments = txtAdminComments.Text
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub dtAssigned_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtAssigned.ValueChanged
        If bolLoading Then Exit Sub
        Try
            UIUtilsGen.ToggleDateFormat(dtAssigned)
            oInspection.AssignedDate = UIUtilsGen.GetDatePickerValue(dtAssigned)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            Dim strErr As String = String.Empty
            If oInspection.OwnerID = 0 Then
                strErr += "Owner" + vbCrLf
            End If
            If oInspection.FacilityID = 0 Then
                strErr += "Facility" + vbCrLf
            End If
            If oInspection.StaffID = 0 Then
                strErr += "Inspector" + vbCrLf
            End If
            If oInspection.InspectionType = 0 Then
                strErr += "Inspection Type" + vbCrLf
            End If
            If Date.Compare(oInspection.AssignedDate, CDate("01/01/0001")) = 0 Then
                strErr += "Assigned Date" + vbCrLf
            End If
            If strErr.Length > 0 Then
                MsgBox("The following are required: " + vbCrLf + strErr)
                Exit Sub
            End If

            If oInspection.ID <= 0 Then
                oInspection.CreatedBy = MusterContainer.AppUser.ID
            Else
                oInspection.ModifiedBy = MusterContainer.AppUser.ID
            End If

            oInspection.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal, True)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            MsgBox("Assigned Inspection saved successfully")
            CallingForm.Tag = "1"
            Me.Close()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            oInspection.Reset()
            Populate()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        If CallingForm.Tag <> "1" Then
            CallingForm.Tag = "1"
        End If
        Me.Close()
    End Sub
    Private Sub btnCreateFCE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCreateFCE.Click
        Try
            If oInspection.IsDirty Then
                Dim Results As Long = MsgBox("There are unsaved changes.  Do you wish to save the changes?", MsgBoxStyle.Question & MsgBoxStyle.YesNoCancel, "Data Changed")
                If Results = MsgBoxResult.Yes Then
                    btnSave.PerformClick()
                Else
                    Exit Sub
                End If
            End If
            'pFCE.Retrieve(0)
            'pFCE.InspectionID = oInspection.ID
            'pFCE.OwnerID = oInspection.OwnerID
            'pFCE.FacilityID = oInspection.FacilityID
            'pFCE.Source = "ADMIN"
            'pFCE.FCEDate = Today.Date
            Dim frmFCE As New FacilityComplianceEvent(oInspection.OwnerID, oInspection.ID, oInspection.FacilityID, 0, )
            frmFCE.CallingForm = Me
            Me.Tag = "0"
            frmFCE.ShowDialog()
            If Me.Tag = "1" Then
                ' FCE was Created
                oInspection.InspectionAccepted = True
                oInspection.Save(CType(UIUtilsGen.ModuleID.CAE, Integer), MusterContainer.AppUser.UserKey, returnVal, True)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If
                CallingForm.Tag = "1"
                Me.Close()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "Form Events"
    Private Sub CandEAssignedInspection_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            bolLoading = True
            Populate()
            bolLoading = False
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub CandEAssignedInspection_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Try
            If oInspection.IsDirty Then
                Dim Results As Long = MsgBox("There are unsaved changes.  Do you wish to save the changes?", MsgBoxStyle.Question & MsgBoxStyle.YesNoCancel, "Data Changed")
                If Results = MsgBoxResult.Yes Then
                    btnSave.PerformClick()
                ElseIf Results = MsgBoxResult.Cancel Then
                    e.Cancel = True
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

#Region "External Events"
    Private Sub InspectionErr(ByVal MsgStr As String) Handles oInspection.evtInspectionErr
        MsgBox(MsgStr)
    End Sub
    Private Sub InspectionChanged(ByVal bolValue As Boolean) Handles oInspection.evtInspectionChanged
        EnableSave(bolValue)
    End Sub
#End Region

End Class
