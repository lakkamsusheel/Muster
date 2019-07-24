Public Class RemediationSystemList
    Inherits System.Windows.Forms.Form

#Region " Local Variables "
    Friend EventActivityID As Int64
    Friend CallingForm As Form
    Friend Mode As Int16
    Private bolLoading As Boolean
    Private oLustRemediation As New MUSTER.BusinessLogic.pLustRemediation
#End Region


#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()
        bolLoading = True
        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        bolLoading = False
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
    Public WithEvents ugSystemList As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents rdoShowAll As System.Windows.Forms.RadioButton
    Friend WithEvents rdoShowAvailable As System.Windows.Forms.RadioButton
    Friend WithEvents btnExpandCollapseAll As System.Windows.Forms.Button
    Friend WithEvents btnAddNew As System.Windows.Forms.Button
    Friend WithEvents pnlMiddle As System.Windows.Forms.Panel
    Friend WithEvents pnlBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlTop As System.Windows.Forms.Panel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.ugSystemList = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlMiddle = New System.Windows.Forms.Panel
        Me.pnlBottom = New System.Windows.Forms.Panel
        Me.btnAddNew = New System.Windows.Forms.Button
        Me.pnlTop = New System.Windows.Forms.Panel
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.rdoShowAll = New System.Windows.Forms.RadioButton
        Me.rdoShowAvailable = New System.Windows.Forms.RadioButton
        Me.btnExpandCollapseAll = New System.Windows.Forms.Button
        CType(Me.ugSystemList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlMiddle.SuspendLayout()
        Me.pnlBottom.SuspendLayout()
        Me.pnlTop.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'ugSystemList
        '
        Me.ugSystemList.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugSystemList.DisplayLayout.AutoFitColumns = True
        Me.ugSystemList.DisplayLayout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
        Me.ugSystemList.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        Me.ugSystemList.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        Me.ugSystemList.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortSingle
        Me.ugSystemList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugSystemList.Location = New System.Drawing.Point(0, 0)
        Me.ugSystemList.Name = "ugSystemList"
        Me.ugSystemList.Size = New System.Drawing.Size(1000, 230)
        Me.ugSystemList.TabIndex = 22
        '
        'pnlMiddle
        '
        Me.pnlMiddle.Controls.Add(Me.ugSystemList)
        Me.pnlMiddle.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlMiddle.Location = New System.Drawing.Point(0, 64)
        Me.pnlMiddle.Name = "pnlMiddle"
        Me.pnlMiddle.Size = New System.Drawing.Size(1000, 230)
        Me.pnlMiddle.TabIndex = 23
        '
        'pnlBottom
        '
        Me.pnlBottom.Controls.Add(Me.btnAddNew)
        Me.pnlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlBottom.Location = New System.Drawing.Point(0, 302)
        Me.pnlBottom.Name = "pnlBottom"
        Me.pnlBottom.Size = New System.Drawing.Size(1000, 40)
        Me.pnlBottom.TabIndex = 23
        '
        'btnAddNew
        '
        Me.btnAddNew.Location = New System.Drawing.Point(16, 9)
        Me.btnAddNew.Name = "btnAddNew"
        Me.btnAddNew.Size = New System.Drawing.Size(96, 23)
        Me.btnAddNew.TabIndex = 5
        Me.btnAddNew.Text = "Add &New"
        '
        'pnlTop
        '
        Me.pnlTop.Controls.Add(Me.GroupBox1)
        Me.pnlTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlTop.Name = "pnlTop"
        Me.pnlTop.Size = New System.Drawing.Size(1000, 64)
        Me.pnlTop.TabIndex = 23
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.rdoShowAll)
        Me.GroupBox1.Controls.Add(Me.rdoShowAvailable)
        Me.GroupBox1.Controls.Add(Me.btnExpandCollapseAll)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(1000, 64)
        Me.GroupBox1.TabIndex = 6
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Remediation System List"
        '
        'rdoShowAll
        '
        Me.rdoShowAll.Location = New System.Drawing.Point(176, 24)
        Me.rdoShowAll.Name = "rdoShowAll"
        Me.rdoShowAll.Size = New System.Drawing.Size(120, 24)
        Me.rdoShowAll.TabIndex = 2
        Me.rdoShowAll.Text = "Show All Systems"
        '
        'rdoShowAvailable
        '
        Me.rdoShowAvailable.Checked = True
        Me.rdoShowAvailable.Location = New System.Drawing.Point(8, 24)
        Me.rdoShowAvailable.Name = "rdoShowAvailable"
        Me.rdoShowAvailable.Size = New System.Drawing.Size(152, 24)
        Me.rdoShowAvailable.TabIndex = 1
        Me.rdoShowAvailable.TabStop = True
        Me.rdoShowAvailable.Text = "Show Available Systems"
        '
        'btnExpandCollapseAll
        '
        Me.btnExpandCollapseAll.Location = New System.Drawing.Point(312, 24)
        Me.btnExpandCollapseAll.Name = "btnExpandCollapseAll"
        Me.btnExpandCollapseAll.Size = New System.Drawing.Size(88, 23)
        Me.btnExpandCollapseAll.TabIndex = 23
        Me.btnExpandCollapseAll.Text = "Expand All"
        Me.btnExpandCollapseAll.Visible = False
        '
        'RemediationSystemList
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(1000, 342)
        Me.Controls.Add(Me.pnlMiddle)
        Me.Controls.Add(Me.pnlBottom)
        Me.Controls.Add(Me.pnlTop)
        Me.Name = "RemediationSystemList"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Remediation Systems"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.ugSystemList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlMiddle.ResumeLayout(False)
        Me.pnlBottom.ResumeLayout(False)
        Me.pnlTop.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


    Private Sub RemediationSystemList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            bolLoading = True
            If Mode = 0 Then
                rdoShowAvailable.Checked = True
                GroupBox1.Enabled = False
            Else
                rdoShowAll.Checked = True
            End If
            LoadSystemList()
        Finally
            bolLoading = False
        End Try
    End Sub

    Private Sub LoadSystemList()
        Cursor = Cursors.WaitCursor

        If rdoShowAvailable.Checked = True Then
            btnExpandCollapseAll.Visible = False
            ugSystemList.DataSource = oLustRemediation.AvailableSystemsDataset()
        Else
            btnExpandCollapseAll.Visible = True
            ugSystemList.DataSource = oLustRemediation.HistoricalSystemsDataset()
        End If

        'ugSystemList.Rows.ExpandAll(True)
        ugSystemList.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
        ugSystemList.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
        If ugSystemList.Rows.Count > 0 Then
            ugSystemList.DisplayLayout.Bands(0).Columns("REM_SYSTEM_ID").Hidden = True
            ugSystemList.DisplayLayout.Bands(0).Columns("SYSTEM_DEC").Hidden = True
            ugSystemList.DisplayLayout.Bands(0).Columns("START_DATE").Hidden = True
            ugSystemList.DisplayLayout.Bands(0).Columns("Facility_ID").Header.Caption = "Facility ID"
            ugSystemList.DisplayLayout.Bands(0).Columns("Facility_Name").Header.Caption = "Facility Name"
            ugSystemList.DisplayLayout.Bands(0).Columns("SYSTEM_SEQ").Header.Caption = "Sequence"
            ugSystemList.DisplayLayout.Bands(0).Columns("REM_SYSTEM_TYPE").Header.Caption = "System Type"
            ugSystemList.DisplayLayout.Bands(0).Columns("DESCRIPTION").Header.Caption = "Description"
            ugSystemList.DisplayLayout.Bands(0).Columns("SYSTEM_OWNER").Header.Caption = "System Owner"
            ugSystemList.DisplayLayout.Bands(0).Columns("VACPUMP1_Size").Header.Caption = "Vac Pump1 Size" 'BUILDING_SIZE, VACPUMP1_Size, VACPUMP2_Size,
            ugSystemList.DisplayLayout.Bands(0).Columns("VACPUMP2_Size").Header.Caption = "Vac Pump2 Size"
            ugSystemList.DisplayLayout.Bands(0).Columns("BUILDING_SIZE").Header.Caption = "Building Size"
            'ugSystemList.DisplayLayout.Bands(0).Columns("REM_SYSTEM_TYPE").Hidden = True
            'ugSystemList.DisplayLayout.Bands(0).Columns("MOTOR_SIZE").Hidden = True

            'ugSystemList.DisplayLayout.Bands(0).Columns("Facility_ID").Width = 35
            'ugSystemList.DisplayLayout.Bands(0).Columns("Facility_Name").Width = 70
            'ugSystemList.DisplayLayout.Bands(0).Columns("SYSTEM_SEQ").Width = 35
            'ugSystemList.DisplayLayout.Bands(0).Columns("REM_SYSTEM_TYPE").Width = 45
            'ugSystemList.DisplayLayout.Bands(0).Columns("DESCRIPTION").Width = 70
            'ugSystemList.DisplayLayout.Bands(0).Columns("SYSTEM_OWNER").Width = 75
            'ugSystemList.DisplayLayout.Bands(0).Columns("VACPUMP1_Size").Width = 90
            'ugSystemList.DisplayLayout.Bands(0).Columns("VACPUMP2_Size").Width = 90
            'ugSystemList.DisplayLayout.Bands(0).Columns("BUILDING_SIZE").Width = 45
            If Mode = 1 And rdoShowAll.Checked Then
                ugSystemList.DisplayLayout.Bands(1).Columns("REM_SYSTEM_ID").Hidden = True
                ugSystemList.DisplayLayout.Bands(1).Columns("SYSTEM_DEC").Hidden = True
                ugSystemList.DisplayLayout.Bands(1).Columns("START_DATE").Hidden = True
                ugSystemList.DisplayLayout.Bands(1).Columns("Facility_ID").Header.Caption = "Facility ID"
                ugSystemList.DisplayLayout.Bands(1).Columns("Facility_Name").Header.Caption = "Facility Name"
                ugSystemList.DisplayLayout.Bands(1).Columns("SYSTEM_SEQ").Header.Caption = "Sequence"
                ugSystemList.DisplayLayout.Bands(1).Columns("REM_SYSTEM_TYPE").Header.Caption = "System Type"
                ugSystemList.DisplayLayout.Bands(1).Columns("DESCRIPTION").Header.Caption = "Description"
                ugSystemList.DisplayLayout.Bands(1).Columns("SYSTEM_OWNER").Header.Caption = "System Owner"
                ugSystemList.DisplayLayout.Bands(1).Columns("VACPUMP1_Size").Header.Caption = "Vac Pump1 Size"
                ugSystemList.DisplayLayout.Bands(1).Columns("VACPUMP2_Size").Header.Caption = "Vac Pump2 Size"
                ugSystemList.DisplayLayout.Bands(1).Columns("BUILDING_SIZE").Header.Caption = "Building Size"

                'ugSystemList.DisplayLayout.Bands(1).Columns("Facility_ID").Width = 35
                'ugSystemList.DisplayLayout.Bands(1).Columns("Facility_Name").Width = 75
                'ugSystemList.DisplayLayout.Bands(1).Columns("SYSTEM_SEQ").Width = 35
                'ugSystemList.DisplayLayout.Bands(1).Columns("REM_SYSTEM_TYPE").Width = 45
                'ugSystemList.DisplayLayout.Bands(1).Columns("DESCRIPTION").Width = 70
                'ugSystemList.DisplayLayout.Bands(1).Columns("SYSTEM_OWNER").Width = 70
                'ugSystemList.DisplayLayout.Bands(1).Columns("VACPUMP1_Size").Width = 90
                'ugSystemList.DisplayLayout.Bands(1).Columns("VACPUMP2_Size").Width = 90
                'ugSystemList.DisplayLayout.Bands(1).Columns("BUILDING_SIZE").Width = 45
            End If
            'ugSystemList.DisplayLayout.Bands(1).Columns("EVENT_ACTIVITY_DOCUMENT_ID").Hidden = True
            'ugSystemList.DisplayLayout.Bands(0).Columns("DESCRIPTION").Header.Caption = "Manufacturer"
            'ugTankandPipes.DisplayLayout.Bands(0).Columns("FuelType").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            'ugTankandPipes.DisplayLayout.Bands(1).Columns("Pipe Site ID").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit


        End If
        Cursor = Cursors.Default
    End Sub

    Private Sub btnAddNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddNew.Click

        ShowRemediationSystemWindow(0, 0)
        If EventActivityID > 0 Then
            Me.Close()
        End If

    End Sub

    Private Sub ShowRemediationSystemWindow(ByVal intMode As Int16, ByVal intSystemID As Integer)
        Dim frmRemSys As New RemediationSystem
        Try
            frmRemSys.CallingForm = Me
            frmRemSys.Mode = intMode
            frmRemSys.nSystemID = intSystemID
            frmRemSys.nActivityID = EventActivityID
            frmRemSys.ShowDialog()

            If Not Me.Tag Is Nothing Then
                If Me.Tag = "1" Then
                    Me.Tag = "0"
                    LoadSystemList()
                End If
            End If
        Finally
            frmRemSys = Nothing
        End Try
    End Sub

    Private Sub ugSystemList_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ugSystemList.DoubleClick
        If bolLoading Then Exit Sub
        If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
            Me.Cursor = Cursors.Default
            Exit Sub
        End If
        Try
            ''Clone the remediation system for use with this activity....
            'oLustRemediation.Retrieve(ugSystemList.ActiveRow.Cells("REM_SYSTEM_ID").Value)
            'oLustRemediation.SystemDeclaration = 0
            'oLustRemediation.ID = 0
            'oLustRemediation.Save()
            If Me.oLustRemediation.CheckRemediationSystemPermissions(Integer.Parse(ugSystemList.ActiveRow.Cells("REM_SYSTEM_ID").Value)) Then
                If ugSystemList.ActiveRow.Band.Index = 0 Then
                    ShowRemediationSystemWindow(0, ugSystemList.ActiveRow.Cells("REM_SYSTEM_ID").Value)
                Else
                    ShowRemediationSystemWindow(3, ugSystemList.ActiveRow.Cells("REM_SYSTEM_ID").Value)
                End If

                If EventActivityID > 0 Then
                    Me.Close()
                End If
            Else
                MsgBox("Remediation System cannot be Modified. It is associated to an Event which is NOT technically completed!")
                Exit Sub
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub rdoShowAvailable_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoShowAvailable.CheckedChanged
        If bolLoading Then Exit Sub
        LoadSystemList()
    End Sub

    Private Sub rdoShowAll_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdoShowAll.CheckedChanged
        If bolLoading Then Exit Sub
        LoadSystemList()
    End Sub

    Private Sub btnExpandCollapseAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExpandCollapseAll.Click
        If bolLoading Then Exit Sub
        If btnExpandCollapseAll.Text = "Expand All" Then
            ugSystemList.Rows.ExpandAll(True)
            btnExpandCollapseAll.Text = "Collapse All"
        Else
            ugSystemList.Rows.CollapseAll(True)
            btnExpandCollapseAll.Text = "Expand All"
        End If
    End Sub
End Class
