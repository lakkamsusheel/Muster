Public Class CAPPreMonthly
    Inherits System.Windows.Forms.Form

#Region "User Defined Variables"
    Public MyGuid As New System.Guid
    Private oReport As MUSTER.BusinessLogic.pReport
    Private rp As New Remove_Pencil
    Private nSelected As Integer = 0
    Private bolLoading As Boolean = False
#End Region

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        MyGuid = System.Guid.NewGuid
        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        MusterContainer.AppUser.LogEntry("CAP", MyGuid.ToString)
        MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "CAP")
    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        MusterContainer.AppSemaphores.Remove(MyGuid.ToString)
        MusterContainer.AppUser.LogExit(MyGuid.ToString)
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
    Friend WithEvents pnlTop As System.Windows.Forms.Panel
    Friend WithEvents pnlBottom As System.Windows.Forms.Panel
    Friend WithEvents ugCAPPreMonthly As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents btnSelectAll As System.Windows.Forms.Button
    Friend WithEvents btnSelectNone As System.Windows.Forms.Button
    Friend WithEvents btnShowProcessed As System.Windows.Forms.Button
    Friend WithEvents lblSelected As System.Windows.Forms.Label
    Friend WithEvents lblSelectedValue As System.Windows.Forms.Label
    Friend WithEvents btnMarkProcessed As System.Windows.Forms.Button
    Friend WithEvents btnExpandAll As System.Windows.Forms.Button
    Friend WithEvents btnRefresh As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlTop = New System.Windows.Forms.Panel
        Me.ugCAPPreMonthly = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlBottom = New System.Windows.Forms.Panel
        Me.lblSelected = New System.Windows.Forms.Label
        Me.btnSelectAll = New System.Windows.Forms.Button
        Me.btnSelectNone = New System.Windows.Forms.Button
        Me.btnShowProcessed = New System.Windows.Forms.Button
        Me.btnMarkProcessed = New System.Windows.Forms.Button
        Me.lblSelectedValue = New System.Windows.Forms.Label
        Me.btnExpandAll = New System.Windows.Forms.Button
        Me.btnRefresh = New System.Windows.Forms.Button
        Me.pnlTop.SuspendLayout()
        CType(Me.ugCAPPreMonthly, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlBottom.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlTop
        '
        Me.pnlTop.Controls.Add(Me.ugCAPPreMonthly)
        Me.pnlTop.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlTop.Name = "pnlTop"
        Me.pnlTop.Size = New System.Drawing.Size(592, 293)
        Me.pnlTop.TabIndex = 0
        '
        'ugCAPPreMonthly
        '
        Me.ugCAPPreMonthly.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugCAPPreMonthly.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugCAPPreMonthly.Location = New System.Drawing.Point(0, 0)
        Me.ugCAPPreMonthly.Name = "ugCAPPreMonthly"
        Me.ugCAPPreMonthly.Size = New System.Drawing.Size(592, 293)
        Me.ugCAPPreMonthly.TabIndex = 0
        '
        'pnlBottom
        '
        Me.pnlBottom.BackColor = System.Drawing.SystemColors.Control
        Me.pnlBottom.Controls.Add(Me.lblSelected)
        Me.pnlBottom.Controls.Add(Me.btnSelectAll)
        Me.pnlBottom.Controls.Add(Me.btnSelectNone)
        Me.pnlBottom.Controls.Add(Me.btnShowProcessed)
        Me.pnlBottom.Controls.Add(Me.btnMarkProcessed)
        Me.pnlBottom.Controls.Add(Me.lblSelectedValue)
        Me.pnlBottom.Controls.Add(Me.btnExpandAll)
        Me.pnlBottom.Controls.Add(Me.btnRefresh)
        Me.pnlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlBottom.Location = New System.Drawing.Point(0, 293)
        Me.pnlBottom.Name = "pnlBottom"
        Me.pnlBottom.Size = New System.Drawing.Size(592, 80)
        Me.pnlBottom.TabIndex = 0
        '
        'lblSelected
        '
        Me.lblSelected.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSelected.Location = New System.Drawing.Point(8, 8)
        Me.lblSelected.Name = "lblSelected"
        Me.lblSelected.Size = New System.Drawing.Size(57, 23)
        Me.lblSelected.TabIndex = 1
        Me.lblSelected.Text = "Selected :"
        Me.lblSelected.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'btnSelectAll
        '
        Me.btnSelectAll.Location = New System.Drawing.Point(8, 40)
        Me.btnSelectAll.Name = "btnSelectAll"
        Me.btnSelectAll.TabIndex = 0
        Me.btnSelectAll.Text = "Select All"
        '
        'btnSelectNone
        '
        Me.btnSelectNone.Location = New System.Drawing.Point(91, 40)
        Me.btnSelectNone.Name = "btnSelectNone"
        Me.btnSelectNone.TabIndex = 0
        Me.btnSelectNone.Text = "Select None"
        '
        'btnShowProcessed
        '
        Me.btnShowProcessed.Location = New System.Drawing.Point(257, 40)
        Me.btnShowProcessed.Name = "btnShowProcessed"
        Me.btnShowProcessed.Size = New System.Drawing.Size(115, 23)
        Me.btnShowProcessed.TabIndex = 0
        Me.btnShowProcessed.Text = "Show Processed"
        '
        'btnMarkProcessed
        '
        Me.btnMarkProcessed.Location = New System.Drawing.Point(472, 40)
        Me.btnMarkProcessed.Name = "btnMarkProcessed"
        Me.btnMarkProcessed.Size = New System.Drawing.Size(112, 23)
        Me.btnMarkProcessed.TabIndex = 0
        Me.btnMarkProcessed.Text = "Mark Processed"
        '
        'lblSelectedValue
        '
        Me.lblSelectedValue.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSelectedValue.Location = New System.Drawing.Point(64, 8)
        Me.lblSelectedValue.Name = "lblSelectedValue"
        Me.lblSelectedValue.Size = New System.Drawing.Size(144, 23)
        Me.lblSelectedValue.TabIndex = 1
        '
        'btnExpandAll
        '
        Me.btnExpandAll.Location = New System.Drawing.Point(174, 40)
        Me.btnExpandAll.Name = "btnExpandAll"
        Me.btnExpandAll.TabIndex = 0
        Me.btnExpandAll.Text = "Expand All"
        '
        'btnRefresh
        '
        Me.btnRefresh.Location = New System.Drawing.Point(380, 40)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(60, 23)
        Me.btnRefresh.TabIndex = 0
        Me.btnRefresh.Text = "Refresh"
        '
        'CAPPreMonthly
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(592, 373)
        Me.Controls.Add(Me.pnlTop)
        Me.Controls.Add(Me.pnlBottom)
        Me.Name = "CAPPreMonthly"
        Me.Text = "CAPPreMonthly"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.pnlTop.ResumeLayout(False)
        CType(Me.ugCAPPreMonthly, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlBottom.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub CAPPreMonthly_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGuid, "CAP")
    End Sub

    Private Sub CAPPreMonthly_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.TextChanged
        MusterContainer.AppSemaphores.Retrieve(MyGuid.ToString, "WindowName", Me.Text, "CAP")
    End Sub

    Private Sub CAPPreMonthly_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            oReport = New MUSTER.BusinessLogic.pReport
            LoadGrid()
            UpdateSelectedTotals()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub ShowError(ByVal ex As Exception)
        Dim MyErr As New ErrorReport(ex)
        MyErr.ShowDialog()
    End Sub
    Private Sub UpdateSelectedTotals()
        lblSelectedValue.Text = nSelected.ToString + " / " + ugCAPPreMonthly.Rows.Count.ToString
        If nSelected = ugCAPPreMonthly.Rows.Count Then
            btnSelectAll.Enabled = False
        Else
            btnSelectAll.Enabled = True
        End If
        If nSelected = 0 Then
            btnSelectNone.Enabled = False
            btnMarkProcessed.Enabled = False
        Else
            btnSelectNone.Enabled = True
            btnMarkProcessed.Enabled = True
        End If
    End Sub
    Private Function LoadGrid(Optional ByVal showPrev As Boolean = False)
        ugCAPPreMonthly.DataSource = oReport.GetCapPreMonthly(showPrev)
        ugCAPPreMonthly.DrawFilter = rp
        nSelected = 0
        btnExpandAll.Text = "Expand All"
    End Function

    Private Sub ugCAPPreMonthly_InitializeLayout(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.InitializeLayoutEventArgs) Handles ugCAPPreMonthly.InitializeLayout
        Try
            e.Layout.Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
            e.Layout.Override.AllowColMoving = Infragistics.Win.UltraWinGrid.AllowColMoving.NotAllowed
            e.Layout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False

            e.Layout.Bands(0).Columns("SELECTED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.AllowEdit
            e.Layout.Bands(0).Columns("OWNER_ID").CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly
            e.Layout.Bands(0).Columns("OWNER").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(0).Columns("FACILITY_ID").CellActivation = Infragistics.Win.UltraWinGrid.Activation.ActivateOnly
            e.Layout.Bands(0).Columns("FACILITY").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(0).Columns("CAP_CANDIDATE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(0).Columns("FAC_CAP_CANDIDATE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(0).Columns("PROCESSED_DATE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(0).Columns("PROCESSED_DATE").Header.Caption = "PROCESSED ON"
            e.Layout.Bands(0).Columns("FACILITY_ID").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending

            e.Layout.Bands(0).Columns("CAP_STATUS_ID").Hidden = True
            e.Layout.Bands(0).Columns("INSPECTION_ID").Hidden = True
            'e.Layout.Bands(0).Columns("CAP_CANDIDATE").Hidden = True
            e.Layout.Bands(0).Columns("CAP_STATUS_BEGIN").Hidden = True
            e.Layout.Bands(0).Columns("CAP_STATUS_END").Hidden = True
            e.Layout.Bands(0).Columns("SUBMITTED_DATE").Hidden = True
            e.Layout.Bands(0).Columns("DELETED").Hidden = True
            e.Layout.Bands(0).Columns("CREATED_BY").Hidden = True
            e.Layout.Bands(0).Columns("DATE_CREATED").Hidden = True
            e.Layout.Bands(0).Columns("LAST_EDITED_BY").Hidden = True
            e.Layout.Bands(0).Columns("DATE_LAST_EDITED").Hidden = True
            e.Layout.Bands(0).Columns("FAC_CURRENT_CAP_STATUS").Hidden = True

            e.Layout.Bands(0).Columns("CAP_CANDIDATE").Header.Caption = "CAP CANDIDATE BEFORE"
            e.Layout.Bands(0).Columns("FAC_CAP_CANDIDATE").Header.Caption = "CAP CANDIDATE NOW"

            ' Tank
            e.Layout.Bands(1).Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect

            e.Layout.Bands(1).Columns("TCP_DATE").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("LI_INSTALL").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("LI_INSPECT").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("TTT").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("LAST_EDITED_BY").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("DATE_LAST_EDITED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            e.Layout.Bands(1).Columns("TCP_DATE_END").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("LI_INSTALL_END").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("LI_INSPECT_END").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("TTT_END").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("LAST_EDITED_BY_END").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("DATE_LAST_EDITED_END").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            'e.Layout.Bands(1).Columns("CAP_DATA_ID").Hidden = True
            e.Layout.Bands(1).Columns("CAP_STATUS_ID").Hidden = True
            'e.Layout.Bands(1).Columns("ENTITY_ID").Hidden = True
            'e.Layout.Bands(1).Columns("IS_TANK").Hidden = True
            'e.Layout.Bands(1).Columns("PIPE_CP_TEST").Hidden = True
            'e.Layout.Bands(1).Columns("TERM_CP_LAST_TESTED").Hidden = True
            'e.Layout.Bands(1).Columns("LTT_DATE").Hidden = True
            'e.Layout.Bands(1).Columns("ALLD_TEST_DATE").Hidden = True
            'e.Layout.Bands(1).Columns("DELETED").Hidden = True

            ' Pipe
            e.Layout.Bands(2).Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect

            e.Layout.Bands(2).Columns("PIPE_CP_TEST").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(2).Columns("TERM_CP_TEST").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(2).Columns("LTT").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(2).Columns("ALLD_TEST").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("LAST_EDITED_BY").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("DATE_LAST_EDITED").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            e.Layout.Bands(2).Columns("PIPE_CP_TEST_END").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(2).Columns("TERM_CP_TEST_END").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(2).Columns("LTT_END").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(2).Columns("ALLD_TEST_END").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("LAST_EDITED_BY_END").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            e.Layout.Bands(1).Columns("DATE_LAST_EDITED_END").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

            'e.Layout.Bands(2).Columns("CAP_DATA_ID").Hidden = True
            e.Layout.Bands(2).Columns("CAP_STATUS_ID").Hidden = True
            'e.Layout.Bands(2).Columns("ENTITY_ID").Hidden = True
            'e.Layout.Bands(2).Columns("IS_TANK").Hidden = True
            'e.Layout.Bands(2).Columns("LASTTCPDATE").Hidden = True
            'e.Layout.Bands(2).Columns("LINEDINTERIORINSTALLDATE").Hidden = True
            'e.Layout.Bands(2).Columns("LINEDINTERIORINSPECTDATE").Hidden = True
            'e.Layout.Bands(2).Columns("TTTDATE").Hidden = True
            'e.Layout.Bands(2).Columns("DELETED").Hidden = True

            Dim dtNull As Date = CDate("01/01/0001")

            For Each ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow In e.Layout.Grid.Rows
                If ugRow.Cells("CAP_CANDIDATE").Value <> ugRow.Cells("FAC_CAP_CANDIDATE").Value Then
                    ugRow.Cells("FAC_CAP_CANDIDATE").Appearance.BackColor = Color.Yellow
                End If
                If Not ugRow.ChildBands Is Nothing Then
                    For Each childband As Infragistics.Win.UltraWinGrid.UltraGridChildBand In ugRow.ChildBands
                        If childband.Index = 0 Then
                            ' tank
                            For Each ugChildRow As Infragistics.Win.UltraWinGrid.UltraGridRow In childband.Rows
                                If ugchildrow.Cells("TCP_DATE").Text <> ugchildrow.Cells("TCP_DATE_END").Text Then
                                    ugchildrow.Cells("TCP_DATE_END").Appearance.BackColor = Color.Yellow
                                End If
                                If ugchildrow.Cells("LI_INSTALL").Text <> ugchildrow.Cells("LI_INSTALL_END").Text Then
                                    ugchildrow.Cells("LI_INSTALL_END").Appearance.BackColor = Color.Yellow
                                End If
                                If ugchildrow.Cells("LI_INSPECT").Text <> ugchildrow.Cells("LI_INSPECT_END").Text Then
                                    ugchildrow.Cells("LI_INSPECT_END").Appearance.BackColor = Color.Yellow
                                End If
                                If ugchildrow.Cells("TTT").Text <> ugchildrow.Cells("TTT_END").Text Then
                                    ugchildrow.Cells("TTT_END").Appearance.BackColor = Color.Yellow
                                End If
                            Next
                        ElseIf childband.Index = 1 Then
                            ' pipe
                            For Each ugChildRow As Infragistics.Win.UltraWinGrid.UltraGridRow In childband.Rows
                                If ugchildrow.Cells("PIPE_CP_TEST").Text <> ugchildrow.Cells("PIPE_CP_TEST_END").Text Then
                                    ugchildrow.Cells("PIPE_CP_TEST_END").Appearance.BackColor = Color.Yellow
                                End If
                                If ugchildrow.Cells("TERM_CP_TEST").Text <> ugchildrow.Cells("TERM_CP_TEST_END").Text Then
                                    ugchildrow.Cells("TERM_CP_TEST_END").Appearance.BackColor = Color.Yellow
                                End If
                                If ugchildrow.Cells("LTT").Text <> ugchildrow.Cells("LTT_END").Text Then
                                    ugchildrow.Cells("LTT_END").Appearance.BackColor = Color.Yellow
                                End If
                                If ugchildrow.Cells("ALLD_TEST").Text <> ugchildrow.Cells("ALLD_TEST_END").Text Then
                                    ugchildrow.Cells("ALLD_TEST_END").Appearance.BackColor = Color.Yellow
                                End If
                            Next
                        End If
                    Next
                End If
            Next
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub ugCAPPreMonthly_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugCAPPreMonthly.CellChange
        If bolLoading Then Exit Sub
        Try
            If "SELECTED".Equals(e.Cell.Column.Key.ToUpper) Then
                If e.Cell.Text.ToUpper = "TRUE" Then
                    nSelected += 1
                Else
                    nSelected -= 1
                End If
                UpdateSelectedTotals()
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub ugCAPPreMonthly_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugCAPPreMonthly.DoubleClick
        If bolLoading Then Exit Sub
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If

            Dim nOwnerID As Integer = 0
            Dim nFacilityID As Integer = 0
            Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow

            If ugCAPPreMonthly.ActiveRow.Band.Index = 0 Then
                ugRow = ugCAPPreMonthly.ActiveRow
            ElseIf ugCAPPreMonthly.ActiveRow.Band.Index = 1 Or ugCAPPreMonthly.ActiveRow.Band.Index = 2 Then
                ugRow = ugCAPPreMonthly.ActiveRow.ParentRow
            Else
                MsgBox("Invalid Row Selected")
                Exit Sub
            End If

            nOwnerID = ugRow.Cells("OWNER_ID").Value
            nFacilityID = ugRow.Cells("FACILITY_ID").Value

            Dim frmCAP As CAPSignUp
            Dim pOwn As New MUSTER.BusinessLogic.pOwner

            pOwn.Retrieve(nOwnerID, , , True)

            frmCAP = New CAPSignUp(pOwn)
            frmCAP.nOwnerID = nOwnerID
            frmCAP.nFacilityID = nFacilityID

            If pOwn.OrganizationID = 0 Then
                frmCAP.txtOwner.Text = IIf(pOwn.BPersona.Title.Trim.Length > 0, pOwn.BPersona.Title.ToString() + " ", "") + pOwn.BPersona.FirstName.ToString() + " " + IIf(pOwn.BPersona.MiddleName.Trim.Length > 0, pOwn.BPersona.MiddleName.ToString() + " ", "") + pOwn.BPersona.LastName.ToString() + IIf(pOwn.BPersona.Suffix.Trim.Length > 0, " " + pOwn.BPersona.Suffix.ToString(), "") + " (" + nOwnerID.ToString + ")"
            ElseIf pOwn.PersonID = 0 Then
                frmCAP.txtOwner.Text = pOwn.Organization.Company + " (" + nOwnerID.ToString + ")"
            End If
            Me.Tag = "0"
            frmCAP.CallingForm = Me
            frmCAP.ShowDialog()
            If Me.Tag = "1" Then
                ' clearing the collection
                pOwn.OwnerInfo.facilityCollection = New MUSTER.Info.FacilityCollection
                pOwn.Facilities = New MUSTER.BusinessLogic.pFacility
                ugRow.Appearance.BackColor = Color.DarkGray
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub btnSelectAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelectAll.Click
        Try
            For Each ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugCAPPreMonthly.Rows
                If ugRow.Cells("SELECTED").Value <> True Then
                    ugRow.Cells("SELECTED").Value = True
                End If
            Next
            nSelected = ugCAPPreMonthly.Rows.Count
            UpdateSelectedTotals()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub btnSelectNone_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSelectNone.Click
        Try
            For Each ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugCAPPreMonthly.Rows
                If ugRow.Cells("SELECTED").Value <> False Then
                    ugRow.Cells("SELECTED").Value = False
                End If
            Next
            nSelected = 0
            UpdateSelectedTotals()
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub btnShowProcessed_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShowProcessed.Click
        If btnShowProcessed.Text = "Show Processed" Then
            LoadGrid(True)
            btnShowProcessed.Text = "Show UnProcessed"
        Else
            LoadGrid()
            btnShowProcessed.Text = "Show Processed"
        End If
        UpdateSelectedTotals()
    End Sub

    Private Sub btnMarkProcessed_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMarkProcessed.Click
        Dim bolMarked As Boolean = False
        Try
            'check rights for Tank & Pipe
            If Not MusterContainer.AppUser.HasAccess(CType(UIUtilsGen.ModuleID.CAPProcess, Integer), MusterContainer.AppUser.UserKey, UIUtilsGen.EntityTypes.Tank) Then
                MessageBox.Show("You do not have Rights to CAP Processing")
                Exit Sub
            ElseIf Not MusterContainer.AppUser.HasAccess(CType(UIUtilsGen.ModuleID.CAPProcess, Integer), MusterContainer.AppUser.UserKey, UIUtilsGen.EntityTypes.Pipe) Then
                MessageBox.Show("You do not have Rights to CAP Processing")
                Exit Sub
            End If
            If nSelected < 1 Then
                MsgBox("No rows Selected", MsgBoxStyle.OKOnly, "Select Rows")
                Exit Sub
            Else
                Dim strMarkProcessed As String = String.Empty
                For Each ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow In ugCAPPreMonthly.Rows
                    If ugRow.Cells("SELECTED").Value = True Then
                        strMarkProcessed += ugRow.Cells("CAP_STATUS_ID").Value.ToString + ", "
                    End If
                    If strMarkProcessed.Length > 450 Then
                        oReport.SaveCapPreMonthly(strMarkProcessed.Trim.TrimEnd(","))
                        bolMarked = True
                        strMarkProcessed = String.Empty
                    End If
                Next
                If strMarkProcessed.Length > 0 Then
                    oReport.SaveCapPreMonthly(strMarkProcessed.Trim.TrimEnd(","))
                    LoadGrid()
                    UpdateSelectedTotals()
                Else
                    If bolMarked Then
                        LoadGrid()
                        UpdateSelectedTotals()
                    Else
                        MsgBox("No rows Selected", MsgBoxStyle.OKOnly, "Select Rows")
                        Exit Sub
                    End If
                End If
            End If
        Catch ex As Exception
            ShowError(ex)
        End Try
    End Sub

    Private Sub btnExpandAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExpandAll.Click
        If btnExpandAll.Text = "Expand All" Then
            ugCAPPreMonthly.Rows.ExpandAll(True)
            btnExpandAll.Text = "Collapse All"
        Else
            ugCAPPreMonthly.Rows.CollapseAll(True)
            btnExpandAll.Text = "Expand All"
        End If
    End Sub

    Private Sub btnRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRefresh.Click
        If btnShowProcessed.Text = "Show Processed" Then
            LoadGrid(True)
        Else
            LoadGrid(False)
        End If
        UpdateSelectedTotals()
    End Sub
End Class
