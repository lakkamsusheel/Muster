'Imports InfoRepository
Public Class TransferEvent
    Inherits System.Windows.Forms.Form

    Friend CallingForm As Form
    Friend FacilityID As Int64
    Friend FacilityName As String

    'Friend nOwnerID As Int64
    Private dsFacilityEvents As DataSet
    Private oFacility As New MUSTER.BusinessLogic.pFacility
    Private oOwner As New MUSTER.BusinessLogic.pOwner
    Private bolLoading As Boolean = False
    Private returnVal As String = String.Empty

#Region " Windows Form Designer generated code "

    Public Sub New(Optional ByRef frmMusterContainer As MusterContainer = Nothing)
        'Public Sub New(Optional ByRef frmMusterContainer As MusterContainer = Nothing, Optional ByRef oRegGrp As InfoRepository.Registrations = Nothing)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        Me.MdiParent = frmMusterContainer
        'Add any initialization after the InitializeComponent() call
        ' AddHandler TransferOwnership.Load, AddressOf mstContainer.loadTransferOwner
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
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents Splitter2 As System.Windows.Forms.Splitter
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents pnlBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlMainLeft As System.Windows.Forms.Panel
    Friend WithEvents pnlRight As System.Windows.Forms.Panel
    Friend WithEvents pnlLeftSub As System.Windows.Forms.Panel
    Friend WithEvents pnlMiddle As System.Windows.Forms.Panel
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents btnShiftRight As System.Windows.Forms.Button
    Friend WithEvents btnShiftLeft As System.Windows.Forms.Button
    Friend WithEvents btnSaveTransfer As System.Windows.Forms.Button
    Friend WithEvents pnlEventsCount As System.Windows.Forms.Panel
    Friend WithEvents lblNoOfEventsOwnerPotentialValue As System.Windows.Forms.Label
    Friend WithEvents lblNoOfEventsOwnerPotential As System.Windows.Forms.Label
    Friend WithEvents lblNoOfEventsOwnerValue As System.Windows.Forms.Label
    Friend WithEvents lblNoofEventsForOwner As System.Windows.Forms.Label
    Friend WithEvents lstViewEvent1 As System.Windows.Forms.ListView
    Friend WithEvents colHeaderSelectEvent1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents colHeaderLUSTEvent1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents colHeaderOpenDate1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents pnlFacility1Top As System.Windows.Forms.Panel
    Friend WithEvents lstViewEvent2 As System.Windows.Forms.ListView
    Friend WithEvents colHeaderSelectEvent2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents colHeaderLUSTEvent2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents colHeaderOpenDate2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents pnlFacility2Top As System.Windows.Forms.Panel
    Friend WithEvents cmbFacility As System.Windows.Forms.ComboBox
    Friend WithEvents lblFacility As System.Windows.Forms.Label
    Friend WithEvents lblCurrFacID As System.Windows.Forms.Label
    Friend WithEvents lblCurrFacIDValue As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlBottom = New System.Windows.Forms.Panel
        Me.btnSaveTransfer = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.pnlEventsCount = New System.Windows.Forms.Panel
        Me.lblNoOfEventsOwnerPotentialValue = New System.Windows.Forms.Label
        Me.lblNoOfEventsOwnerPotential = New System.Windows.Forms.Label
        Me.lblNoOfEventsOwnerValue = New System.Windows.Forms.Label
        Me.lblNoofEventsForOwner = New System.Windows.Forms.Label
        Me.pnlMainLeft = New System.Windows.Forms.Panel
        Me.Splitter2 = New System.Windows.Forms.Splitter
        Me.pnlMiddle = New System.Windows.Forms.Panel
        Me.btnShiftLeft = New System.Windows.Forms.Button
        Me.btnShiftRight = New System.Windows.Forms.Button
        Me.pnlLeftSub = New System.Windows.Forms.Panel
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.lstViewEvent1 = New System.Windows.Forms.ListView
        Me.colHeaderSelectEvent1 = New System.Windows.Forms.ColumnHeader
        Me.colHeaderLUSTEvent1 = New System.Windows.Forms.ColumnHeader
        Me.colHeaderOpenDate1 = New System.Windows.Forms.ColumnHeader
        Me.pnlFacility1Top = New System.Windows.Forms.Panel
        Me.lblCurrFacID = New System.Windows.Forms.Label
        Me.lblCurrFacIDValue = New System.Windows.Forms.Label
        Me.pnlRight = New System.Windows.Forms.Panel
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.lstViewEvent2 = New System.Windows.Forms.ListView
        Me.colHeaderSelectEvent2 = New System.Windows.Forms.ColumnHeader
        Me.colHeaderLUSTEvent2 = New System.Windows.Forms.ColumnHeader
        Me.colHeaderOpenDate2 = New System.Windows.Forms.ColumnHeader
        Me.pnlFacility2Top = New System.Windows.Forms.Panel
        Me.cmbFacility = New System.Windows.Forms.ComboBox
        Me.lblFacility = New System.Windows.Forms.Label
        Me.Splitter1 = New System.Windows.Forms.Splitter
        Me.pnlBottom.SuspendLayout()
        Me.pnlEventsCount.SuspendLayout()
        Me.pnlMainLeft.SuspendLayout()
        Me.pnlMiddle.SuspendLayout()
        Me.pnlLeftSub.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.pnlFacility1Top.SuspendLayout()
        Me.pnlRight.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.pnlFacility2Top.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlBottom
        '
        Me.pnlBottom.Controls.Add(Me.btnSaveTransfer)
        Me.pnlBottom.Controls.Add(Me.btnCancel)
        Me.pnlBottom.Controls.Add(Me.pnlEventsCount)
        Me.pnlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlBottom.DockPadding.Left = 10
        Me.pnlBottom.Location = New System.Drawing.Point(0, 438)
        Me.pnlBottom.Name = "pnlBottom"
        Me.pnlBottom.Size = New System.Drawing.Size(800, 160)
        Me.pnlBottom.TabIndex = 0
        '
        'btnSaveTransfer
        '
        Me.btnSaveTransfer.Location = New System.Drawing.Point(288, 128)
        Me.btnSaveTransfer.Name = "btnSaveTransfer"
        Me.btnSaveTransfer.Size = New System.Drawing.Size(104, 23)
        Me.btnSaveTransfer.TabIndex = 14
        Me.btnSaveTransfer.Text = "Save Transfer"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(400, 128)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(104, 23)
        Me.btnCancel.TabIndex = 13
        Me.btnCancel.Text = "Cancel"
        '
        'pnlEventsCount
        '
        Me.pnlEventsCount.Controls.Add(Me.lblNoOfEventsOwnerPotentialValue)
        Me.pnlEventsCount.Controls.Add(Me.lblNoOfEventsOwnerPotential)
        Me.pnlEventsCount.Controls.Add(Me.lblNoOfEventsOwnerValue)
        Me.pnlEventsCount.Controls.Add(Me.lblNoofEventsForOwner)
        Me.pnlEventsCount.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlEventsCount.Location = New System.Drawing.Point(10, 0)
        Me.pnlEventsCount.Name = "pnlEventsCount"
        Me.pnlEventsCount.Size = New System.Drawing.Size(790, 16)
        Me.pnlEventsCount.TabIndex = 16
        '
        'lblNoOfEventsOwnerPotentialValue
        '
        Me.lblNoOfEventsOwnerPotentialValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNoOfEventsOwnerPotentialValue.Dock = System.Windows.Forms.DockStyle.Right
        Me.lblNoOfEventsOwnerPotentialValue.Location = New System.Drawing.Point(750, 0)
        Me.lblNoOfEventsOwnerPotentialValue.Name = "lblNoOfEventsOwnerPotentialValue"
        Me.lblNoOfEventsOwnerPotentialValue.Size = New System.Drawing.Size(40, 16)
        Me.lblNoOfEventsOwnerPotentialValue.TabIndex = 3
        '
        'lblNoOfEventsOwnerPotential
        '
        Me.lblNoOfEventsOwnerPotential.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNoOfEventsOwnerPotential.Location = New System.Drawing.Point(624, 0)
        Me.lblNoOfEventsOwnerPotential.Name = "lblNoOfEventsOwnerPotential"
        Me.lblNoOfEventsOwnerPotential.Size = New System.Drawing.Size(115, 18)
        Me.lblNoOfEventsOwnerPotential.TabIndex = 2
        Me.lblNoOfEventsOwnerPotential.Text = "No. of LUST Events:"
        '
        'lblNoOfEventsOwnerValue
        '
        Me.lblNoOfEventsOwnerValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNoOfEventsOwnerValue.Location = New System.Drawing.Point(120, 0)
        Me.lblNoOfEventsOwnerValue.Name = "lblNoOfEventsOwnerValue"
        Me.lblNoOfEventsOwnerValue.Size = New System.Drawing.Size(40, 18)
        Me.lblNoOfEventsOwnerValue.TabIndex = 1
        '
        'lblNoofEventsForOwner
        '
        Me.lblNoofEventsForOwner.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNoofEventsForOwner.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblNoofEventsForOwner.Location = New System.Drawing.Point(0, 0)
        Me.lblNoofEventsForOwner.Name = "lblNoofEventsForOwner"
        Me.lblNoofEventsForOwner.Size = New System.Drawing.Size(112, 16)
        Me.lblNoofEventsForOwner.TabIndex = 0
        Me.lblNoofEventsForOwner.Text = "No. of LUST Events:"
        '
        'pnlMainLeft
        '
        Me.pnlMainLeft.Controls.Add(Me.Splitter2)
        Me.pnlMainLeft.Controls.Add(Me.pnlMiddle)
        Me.pnlMainLeft.Controls.Add(Me.pnlLeftSub)
        Me.pnlMainLeft.Dock = System.Windows.Forms.DockStyle.Left
        Me.pnlMainLeft.DockPadding.Left = 10
        Me.pnlMainLeft.Location = New System.Drawing.Point(0, 0)
        Me.pnlMainLeft.Name = "pnlMainLeft"
        Me.pnlMainLeft.Size = New System.Drawing.Size(440, 438)
        Me.pnlMainLeft.TabIndex = 1
        '
        'Splitter2
        '
        Me.Splitter2.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Splitter2.Location = New System.Drawing.Point(386, 0)
        Me.Splitter2.Name = "Splitter2"
        Me.Splitter2.Size = New System.Drawing.Size(3, 438)
        Me.Splitter2.TabIndex = 2
        Me.Splitter2.TabStop = False
        '
        'pnlMiddle
        '
        Me.pnlMiddle.Controls.Add(Me.btnShiftLeft)
        Me.pnlMiddle.Controls.Add(Me.btnShiftRight)
        Me.pnlMiddle.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlMiddle.Location = New System.Drawing.Point(386, 0)
        Me.pnlMiddle.Name = "pnlMiddle"
        Me.pnlMiddle.Size = New System.Drawing.Size(54, 438)
        Me.pnlMiddle.TabIndex = 1
        '
        'btnShiftLeft
        '
        Me.btnShiftLeft.Font = New System.Drawing.Font("Arial Black", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnShiftLeft.Location = New System.Drawing.Point(11, 167)
        Me.btnShiftLeft.Name = "btnShiftLeft"
        Me.btnShiftLeft.Size = New System.Drawing.Size(32, 32)
        Me.btnShiftLeft.TabIndex = 1
        Me.btnShiftLeft.Text = "<<"
        '
        'btnShiftRight
        '
        Me.btnShiftRight.Font = New System.Drawing.Font("Arial Black", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnShiftRight.Location = New System.Drawing.Point(10, 128)
        Me.btnShiftRight.Name = "btnShiftRight"
        Me.btnShiftRight.Size = New System.Drawing.Size(32, 32)
        Me.btnShiftRight.TabIndex = 0
        Me.btnShiftRight.Text = ">>"
        '
        'pnlLeftSub
        '
        Me.pnlLeftSub.Controls.Add(Me.GroupBox1)
        Me.pnlLeftSub.Dock = System.Windows.Forms.DockStyle.Left
        Me.pnlLeftSub.Location = New System.Drawing.Point(10, 0)
        Me.pnlLeftSub.Name = "pnlLeftSub"
        Me.pnlLeftSub.Size = New System.Drawing.Size(376, 438)
        Me.pnlLeftSub.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lstViewEvent1)
        Me.GroupBox1.Controls.Add(Me.pnlFacility1Top)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(376, 438)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        '
        'lstViewEvent1
        '
        Me.lstViewEvent1.CheckBoxes = True
        Me.lstViewEvent1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.colHeaderSelectEvent1, Me.colHeaderLUSTEvent1, Me.colHeaderOpenDate1})
        Me.lstViewEvent1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstViewEvent1.FullRowSelect = True
        Me.lstViewEvent1.GridLines = True
        Me.lstViewEvent1.Location = New System.Drawing.Point(3, 88)
        Me.lstViewEvent1.Name = "lstViewEvent1"
        Me.lstViewEvent1.Size = New System.Drawing.Size(370, 347)
        Me.lstViewEvent1.TabIndex = 13
        Me.lstViewEvent1.View = System.Windows.Forms.View.Details
        '
        'colHeaderSelectEvent1
        '
        Me.colHeaderSelectEvent1.Text = ""
        Me.colHeaderSelectEvent1.Width = 22
        '
        'colHeaderLUSTEvent1
        '
        Me.colHeaderLUSTEvent1.Text = "LUST Event"
        Me.colHeaderLUSTEvent1.Width = 100
        '
        'colHeaderOpenDate1
        '
        Me.colHeaderOpenDate1.Text = "Open Date"
        Me.colHeaderOpenDate1.Width = 94
        '
        'pnlFacility1Top
        '
        Me.pnlFacility1Top.Controls.Add(Me.lblCurrFacIDValue)
        Me.pnlFacility1Top.Controls.Add(Me.lblCurrFacID)
        Me.pnlFacility1Top.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlFacility1Top.Location = New System.Drawing.Point(3, 16)
        Me.pnlFacility1Top.Name = "pnlFacility1Top"
        Me.pnlFacility1Top.Size = New System.Drawing.Size(370, 72)
        Me.pnlFacility1Top.TabIndex = 12
        '
        'lblCurrFacID
        '
        Me.lblCurrFacID.Location = New System.Drawing.Point(8, 24)
        Me.lblCurrFacID.Name = "lblCurrFacID"
        Me.lblCurrFacID.Size = New System.Drawing.Size(139, 23)
        Me.lblCurrFacID.TabIndex = 0
        Me.lblCurrFacID.Text = "Current Facility ID (Name):"
        '
        'lblCurrFacIDValue
        '
        Me.lblCurrFacIDValue.Location = New System.Drawing.Point(135, 24)
        Me.lblCurrFacIDValue.Name = "lblCurrFacIDValue"
        Me.lblCurrFacIDValue.Size = New System.Drawing.Size(225, 23)
        Me.lblCurrFacIDValue.TabIndex = 0
        Me.lblCurrFacIDValue.Text = "203"
        '
        'pnlRight
        '
        Me.pnlRight.Controls.Add(Me.GroupBox2)
        Me.pnlRight.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlRight.DockPadding.Left = 4
        Me.pnlRight.Location = New System.Drawing.Point(440, 0)
        Me.pnlRight.Name = "pnlRight"
        Me.pnlRight.Size = New System.Drawing.Size(360, 438)
        Me.pnlRight.TabIndex = 2
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.lstViewEvent2)
        Me.GroupBox2.Controls.Add(Me.pnlFacility2Top)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.Location = New System.Drawing.Point(4, 0)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(356, 438)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        '
        'lstViewEvent2
        '
        Me.lstViewEvent2.CheckBoxes = True
        Me.lstViewEvent2.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.colHeaderSelectEvent2, Me.colHeaderLUSTEvent2, Me.colHeaderOpenDate2})
        Me.lstViewEvent2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstViewEvent2.FullRowSelect = True
        Me.lstViewEvent2.GridLines = True
        Me.lstViewEvent2.Location = New System.Drawing.Point(3, 88)
        Me.lstViewEvent2.Name = "lstViewEvent2"
        Me.lstViewEvent2.Size = New System.Drawing.Size(350, 347)
        Me.lstViewEvent2.TabIndex = 10
        Me.lstViewEvent2.View = System.Windows.Forms.View.Details
        '
        'colHeaderSelectEvent2
        '
        Me.colHeaderSelectEvent2.Text = ""
        Me.colHeaderSelectEvent2.Width = 22
        '
        'colHeaderLUSTEvent2
        '
        Me.colHeaderLUSTEvent2.Text = "LUST Event"
        Me.colHeaderLUSTEvent2.Width = 71
        '
        'colHeaderOpenDate2
        '
        Me.colHeaderOpenDate2.Text = "Open Date"
        Me.colHeaderOpenDate2.Width = 100
        '
        'pnlFacility2Top
        '
        Me.pnlFacility2Top.Controls.Add(Me.cmbFacility)
        Me.pnlFacility2Top.Controls.Add(Me.lblFacility)
        Me.pnlFacility2Top.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlFacility2Top.Location = New System.Drawing.Point(3, 16)
        Me.pnlFacility2Top.Name = "pnlFacility2Top"
        Me.pnlFacility2Top.Size = New System.Drawing.Size(350, 72)
        Me.pnlFacility2Top.TabIndex = 9
        '
        'cmbFacility
        '
        Me.cmbFacility.Location = New System.Drawing.Point(8, 24)
        Me.cmbFacility.Name = "cmbFacility"
        Me.cmbFacility.Size = New System.Drawing.Size(296, 21)
        Me.cmbFacility.TabIndex = 8
        '
        'lblFacility
        '
        Me.lblFacility.Location = New System.Drawing.Point(8, 7)
        Me.lblFacility.Name = "lblFacility"
        Me.lblFacility.Size = New System.Drawing.Size(128, 23)
        Me.lblFacility.TabIndex = 5
        Me.lblFacility.Text = "New Facility (ID - Name)"
        '
        'Splitter1
        '
        Me.Splitter1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Splitter1.Location = New System.Drawing.Point(440, 0)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(3, 438)
        Me.Splitter1.TabIndex = 3
        Me.Splitter1.TabStop = False
        '
        'TransferEvent
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(800, 598)
        Me.Controls.Add(Me.Splitter1)
        Me.Controls.Add(Me.pnlRight)
        Me.Controls.Add(Me.pnlMainLeft)
        Me.Controls.Add(Me.pnlBottom)
        Me.Name = "TransferEvent"
        Me.Text = "Technical - Transfer Event"
        Me.pnlBottom.ResumeLayout(False)
        Me.pnlEventsCount.ResumeLayout(False)
        Me.pnlMainLeft.ResumeLayout(False)
        Me.pnlMiddle.ResumeLayout(False)
        Me.pnlLeftSub.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.pnlFacility1Top.ResumeLayout(False)
        Me.pnlRight.ResumeLayout(False)
        Me.GroupBox2.ResumeLayout(False)
        Me.pnlFacility2Top.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub btnShiftRight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShiftRight.Click
        Dim tmpListView As ListViewItem

        Try

            For Each tmpListView In lstViewEvent1.Items
                If tmpListView.Checked Then
                    tmpListView.BackColor = System.Drawing.Color.PaleGreen
                    lstViewEvent1.Items.Remove(tmpListView)
                    lstViewEvent2.Items.Add(tmpListView)
                    tmpListView.Checked = False
                End If
            Next
            lblNoOfEventsOwnerValue.Text = lstViewEvent1.Items.Count
            lblNoOfEventsOwnerPotentialValue.Text = lstViewEvent2.Items.Count
        Catch ex As Exception

        End Try

    End Sub

    Private Sub btnShiftLeft_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShiftLeft.Click
        Dim tmpListView As ListViewItem

        Try

            For Each tmpListView In lstViewEvent2.Items
                If tmpListView.Checked Then
                    If tmpListView.Tag <> 0 Then
                        tmpListView.BackColor = System.Drawing.Color.White
                        lstViewEvent2.Items.Remove(tmpListView)
                        lstViewEvent1.Items.Add(tmpListView)
                    End If
                    tmpListView.Checked = False
                End If
            Next
            lblNoOfEventsOwnerValue.Text = lstViewEvent1.Items.Count
            lblNoOfEventsOwnerPotentialValue.Text = lstViewEvent2.Items.Count
        Catch ex As Exception

        End Try

    End Sub

    Private Sub TransferEvent_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        lblCurrFacIDValue.Text = FacilityID.ToString + " (" + FacilityName + ")"
        oFacility.Retrieve(oOwner.OwnerInfo, FacilityID, , "FACILITY")
        dsFacilityEvents = oFacility.LustEventDataset(True)

        bolLoading = True
        PopulatecmbFacility()
        bolLoading = False

        Load_lstEvent1()
        Load_lstEvent2()

    End Sub

    Private Sub PopulatecmbFacility()

        Try
            Dim dtLustEventStatus As DataTable = oOwner.PopulateOwnerFacilities(0, True)
            dtLustEventStatus.Columns.Add("FACILITY_ID_NAME")
            For Each dr As DataRow In dtLustEventStatus.Rows
                dr("FACILITY_ID_NAME") = dr("FACILITY_ID").ToString + " - " + dr("FACILITY_NAME")
            Next
            If Not IsNothing(dtLustEventStatus) Then

                cmbFacility.DataSource = dtLustEventStatus
                cmbFacility.DisplayMember = "FACILITY_ID_NAME"
                cmbFacility.ValueMember = "FACILITY_ID"
            Else
                cmbFacility.DataSource = Nothing
            End If
            cmbFacility.SelectedIndex = -1

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Populate Owner Facilities" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub Load_lstEvent1()
        Dim tmpListView As ListViewItem
        'Dim dtOwn2Facilities As DataTable
        Dim drEvent As DataRow
        Try
            lstViewEvent1.Items.Clear()

            For Each drEvent In dsFacilityEvents.Tables(0).Rows
                If drEvent(7) = "Open" Then
                    tmpListView = New ListViewItem(New String() {"", "Lust Event # " & drEvent(2), IIf(drEvent(3) Is DBNull.Value, "", drEvent(3))})
                    tmpListView.Tag = drEvent(0)
                    lstViewEvent1.Items.Add(tmpListView)
                End If
            Next

            lblNoOfEventsOwnerValue.Text = lstViewEvent1.Items.Count
            If lstViewEvent1.Items.Count = 0 Then
                btnSaveTransfer.Enabled = False
            Else
                btnSaveTransfer.Enabled = True
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub cmbFacility_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbFacility.SelectedIndexChanged
        If bolLoading Then
            Exit Sub
        End If

        Load_lstEvent1()
        Load_lstEvent2()

    End Sub

    Private Sub Load_lstEvent2()
        Dim tmpListView As ListViewItem
        Dim dsFac2Events As DataSet
        Dim drEvent As DataRow
        Dim xFacility As New MUSTER.BusinessLogic.pFacility

        Try

            lstViewEvent2.Items.Clear()

            If cmbFacility.SelectedValue > 0 Then

                xFacility.Retrieve(oOwner.OwnerInfo, cmbFacility.SelectedValue, , "FACILITY")
                dsFac2Events = xFacility.LustEventDataset

                For Each drEvent In dsFac2Events.Tables(0).Rows
                    tmpListView = New ListViewItem(New String() {"", "Lust Event # " & drEvent(2), IIf(drEvent(3) Is DBNull.Value, "", drEvent(3))})
                    tmpListView.Tag = 0
                    lstViewEvent2.Items.Add(tmpListView)
                Next

                Me.lblNoOfEventsOwnerPotentialValue.Text = lstViewEvent2.Items.Count
            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub btnSaveTransfer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveTransfer.Click
        Dim tmpListView As ListViewItem
        Dim oLustEvent As New MUSTER.BusinessLogic.pLustEvent
        Dim bolRecordSaved As Boolean = False
        Dim dsSet As DataSet
        Try
            For Each tmpListView In lstViewEvent2.Items
                If tmpListView.Tag <> 0 Then

                    oLustEvent.Retrieve(tmpListView.Tag)
                    dsSet = oLustEvent.GetTecOpenFinPO(oLustEvent.ID)
                    If dsSet.Tables.Count > 0 Then
                        If dsSet.Tables(0).Rows.Count > 0 Then
                            If Not dsSet.Tables(0).Rows(0).Item("Balance") Is System.DBNull.Value Then
                                MsgBox("Cannot transfer Lust Event " + oLustEvent.ID.ToString + " .The Lust Event has open financial POs")
                                Exit For
                            End If
                        End If
                    End If
                    oLustEvent.FacilityID = cmbFacility.SelectedValue
                    oLustEvent.ModifiedBy = MusterContainer.AppUser.ID
                    oLustEvent.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), MusterContainer.AppUser.UserKey, returnVal)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If
                    bolRecordSaved = True
                End If
            Next

            If bolRecordSaved Then
                MsgBox("Event(s) transferred")
                Me.Close()
            End If

        Catch ex As Exception

        End Try
    End Sub
End Class
