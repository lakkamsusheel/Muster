Public Class InspectionHistory
    Inherits System.Windows.Forms.Form
#Region "User Defined Variables"
    Private WithEvents oInspection As MUSTER.BusinessLogic.pInspection
    Private frmChecklist As CheckList
    Private frmAssignedInspection As AssignedInspection
    Private rp As New Remove_Pencil
    Private bolLoading As Boolean = False
#End Region
#Region "Windows Form Designer generated code "

    Public Sub New(ByRef oInspec As MUSTER.BusinessLogic.pInspection, ByVal dsViewHistory As DataSet, ByVal strFacilityName As String)
        MyBase.New()
        bolLoading = True
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        'Add any initialization after the InitializeComponent() call
        Cursor.Current = Cursors.AppStarting
        oInspection = oInspec
        LoadInspectionHistory(dsViewHistory, strFacilityName)
        bolLoading = False
        If dsViewHistory.Tables(0).Rows.Count > 0 Then
            ugHistory.Rows(0).Activate()
        End If
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
    Friend WithEvents pnlInspecHistoryDetails As System.Windows.Forms.Panel
    Friend WithEvents pnlInspecHistoryBottom As System.Windows.Forms.Panel
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents pnlInspecHistoryHeader As System.Windows.Forms.Panel
    Friend WithEvents lblFacIDValue As System.Windows.Forms.Label
    Friend WithEvents lblHistory As System.Windows.Forms.Label
    Friend WithEvents lblFacID As System.Windows.Forms.Label
    Friend WithEvents pnlInspecHistoryGrid As System.Windows.Forms.Panel
    Friend WithEvents ugHistory As Infragistics.Win.UltraWinGrid.UltraGrid
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlInspecHistoryDetails = New System.Windows.Forms.Panel
        Me.pnlInspecHistoryGrid = New System.Windows.Forms.Panel
        Me.ugHistory = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlInspecHistoryHeader = New System.Windows.Forms.Panel
        Me.lblFacIDValue = New System.Windows.Forms.Label
        Me.lblHistory = New System.Windows.Forms.Label
        Me.lblFacID = New System.Windows.Forms.Label
        Me.pnlInspecHistoryBottom = New System.Windows.Forms.Panel
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnView = New System.Windows.Forms.Button
        Me.pnlInspecHistoryDetails.SuspendLayout()
        Me.pnlInspecHistoryGrid.SuspendLayout()
        CType(Me.ugHistory, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlInspecHistoryHeader.SuspendLayout()
        Me.pnlInspecHistoryBottom.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlInspecHistoryDetails
        '
        Me.pnlInspecHistoryDetails.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlInspecHistoryDetails.Controls.Add(Me.pnlInspecHistoryGrid)
        Me.pnlInspecHistoryDetails.Controls.Add(Me.pnlInspecHistoryHeader)
        Me.pnlInspecHistoryDetails.Controls.Add(Me.pnlInspecHistoryBottom)
        Me.pnlInspecHistoryDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlInspecHistoryDetails.Location = New System.Drawing.Point(0, 0)
        Me.pnlInspecHistoryDetails.Name = "pnlInspecHistoryDetails"
        Me.pnlInspecHistoryDetails.Size = New System.Drawing.Size(504, 357)
        Me.pnlInspecHistoryDetails.TabIndex = 0
        '
        'pnlInspecHistoryGrid
        '
        Me.pnlInspecHistoryGrid.Controls.Add(Me.ugHistory)
        Me.pnlInspecHistoryGrid.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlInspecHistoryGrid.Location = New System.Drawing.Point(0, 51)
        Me.pnlInspecHistoryGrid.Name = "pnlInspecHistoryGrid"
        Me.pnlInspecHistoryGrid.Size = New System.Drawing.Size(500, 262)
        Me.pnlInspecHistoryGrid.TabIndex = 7
        '
        'ugHistory
        '
        Me.ugHistory.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugHistory.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugHistory.Location = New System.Drawing.Point(0, 0)
        Me.ugHistory.Name = "ugHistory"
        Me.ugHistory.Size = New System.Drawing.Size(500, 262)
        Me.ugHistory.TabIndex = 0
        '
        'pnlInspecHistoryHeader
        '
        Me.pnlInspecHistoryHeader.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlInspecHistoryHeader.Controls.Add(Me.lblFacIDValue)
        Me.pnlInspecHistoryHeader.Controls.Add(Me.lblHistory)
        Me.pnlInspecHistoryHeader.Controls.Add(Me.lblFacID)
        Me.pnlInspecHistoryHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlInspecHistoryHeader.Location = New System.Drawing.Point(0, 0)
        Me.pnlInspecHistoryHeader.Name = "pnlInspecHistoryHeader"
        Me.pnlInspecHistoryHeader.Size = New System.Drawing.Size(500, 51)
        Me.pnlInspecHistoryHeader.TabIndex = 0
        '
        'lblFacIDValue
        '
        Me.lblFacIDValue.Location = New System.Drawing.Point(56, 8)
        Me.lblFacIDValue.Name = "lblFacIDValue"
        Me.lblFacIDValue.Size = New System.Drawing.Size(432, 30)
        Me.lblFacIDValue.TabIndex = 0
        '
        'lblHistory
        '
        Me.lblHistory.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHistory.Location = New System.Drawing.Point(4, 32)
        Me.lblHistory.Name = "lblHistory"
        Me.lblHistory.Size = New System.Drawing.Size(44, 17)
        Me.lblHistory.TabIndex = 0
        Me.lblHistory.Text = "History"
        '
        'lblFacID
        '
        Me.lblFacID.Location = New System.Drawing.Point(8, 8)
        Me.lblFacID.Name = "lblFacID"
        Me.lblFacID.Size = New System.Drawing.Size(48, 17)
        Me.lblFacID.TabIndex = 0
        Me.lblFacID.Text = "Facility:"
        Me.lblFacID.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'pnlInspecHistoryBottom
        '
        Me.pnlInspecHistoryBottom.Controls.Add(Me.btnClose)
        Me.pnlInspecHistoryBottom.Controls.Add(Me.btnView)
        Me.pnlInspecHistoryBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlInspecHistoryBottom.Location = New System.Drawing.Point(0, 313)
        Me.pnlInspecHistoryBottom.Name = "pnlInspecHistoryBottom"
        Me.pnlInspecHistoryBottom.Size = New System.Drawing.Size(500, 40)
        Me.pnlInspecHistoryBottom.TabIndex = 1
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(254, 8)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(72, 23)
        Me.btnClose.TabIndex = 1
        Me.btnClose.Text = "Close"
        '
        'btnView
        '
        Me.btnView.Location = New System.Drawing.Point(174, 8)
        Me.btnView.Name = "btnView"
        Me.btnView.Size = New System.Drawing.Size(72, 23)
        Me.btnView.TabIndex = 0
        Me.btnView.Text = "View"
        '
        'InspectionHistory
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(504, 357)
        Me.Controls.Add(Me.pnlInspecHistoryDetails)
        Me.MinimizeBox = False
        Me.Name = "InspectionHistory"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Inspection History"
        Me.pnlInspecHistoryDetails.ResumeLayout(False)
        Me.pnlInspecHistoryGrid.ResumeLayout(False)
        CType(Me.ugHistory, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlInspecHistoryHeader.ResumeLayout(False)
        Me.pnlInspecHistoryBottom.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
#End Region
#Region "UI Support Routines"
    Private Sub LoadInspectionHistory(ByVal dsInspecHistory As DataSet, ByVal strFacilityName As String)
        Try
            'For Each drColumn As DataColumn In dsInspecHistory.Tables(0).Columns
            '    drColumn.ReadOnly = True
            'Next

            ugHistory.DataSource = dsInspecHistory
            ugHistory.DrawFilter = rp

            ugHistory.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
            ugHistory.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
            ugHistory.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
            ugHistory.DisplayLayout.Bands(0).Columns("LAST INSPECTED ON").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            ugHistory.DisplayLayout.Bands(0).Columns("COMPLETED").Hidden = True
            ugHistory.DisplayLayout.Bands(0).Columns("INSPECTION_ID").Hidden = True
            ugHistory.DisplayLayout.Bands(0).Columns("OWNER_ID").Hidden = True
            ugHistory.DisplayLayout.Bands(0).Columns("FACILITY ID").Hidden = True
            ugHistory.DisplayLayout.Bands(0).Columns("SCHEDULED_BY").Hidden = True

            ugHistory.DisplayLayout.Bands(0).Columns("OWNER NAME").Hidden = True
            ugHistory.DisplayLayout.Bands(0).Columns("FACILITY").Hidden = True
            ugHistory.DisplayLayout.Bands(0).Columns("ASSIGNED DATE").Hidden = True
            ugHistory.DisplayLayout.Bands(0).Columns("OWNER PHONE").Hidden = True
            ugHistory.DisplayLayout.Bands(0).Columns("COUNTY").Hidden = True
            ugHistory.DisplayLayout.Bands(0).Columns("ADDRESS_LINE_ONE").Hidden = True
            ugHistory.DisplayLayout.Bands(0).Columns("CITY").Hidden = True
            ugHistory.DisplayLayout.Bands(0).Columns("ADMIN COMMENTS").Hidden = True
            ugHistory.DisplayLayout.Bands(0).Columns("COMPLETED").Hidden = True
            ugHistory.DisplayLayout.Bands(0).Columns("SUBMITTED").Hidden = True
            ugHistory.DisplayLayout.Bands(0).Columns("INSPECTOR COMMENTS").Hidden = True

            lblFacIDValue.Text = "#" + (ugHistory.Rows(0).Cells("FACILITY ID").Value).ToString + " - " + strFacilityName
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#Region "UI Control Events"
    Private Sub btnView_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnView.Click
        Dim ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            If ugHistory.Rows.Count > 0 Then
                If Not ugHistory.ActiveRow Is Nothing Then
                    ugrow = ugHistory.ActiveRow
                End If
            End If
            If ugrow Is Nothing Then
                MsgBox("Please select an inspection to view")
                Exit Sub
            End If
            oInspection.Retrieve(CType(ugrow.Cells("INSPECTION_ID").Value, Int64), , CType(ugrow.Cells("FACILITY ID").Value, Int64), CType(ugrow.Cells("OWNER_ID").Value, Int64))
            If ugrow.Cells("SCHEDULED_BY").Value Is DBNull.Value Then
                frmAssignedInspection = New AssignedInspection(ugrow, oInspection, True)
                frmAssignedInspection.CallingForm = Me
                frmAssignedInspection.ShowDialog()
            Else
                frmChecklist = New CheckList(oInspection, True)
                frmChecklist.WindowState = FormWindowState.Maximized
                frmChecklist.ShowDialog()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    'Private Sub ugHistory_BeforeRowActivate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.RowEventArgs) Handles ugHistory.BeforeRowActivate
    '    If bolLoading Then Exit Sub
    '    Try
    '        If e.Row.Cells("SCHEDULED_BY").Value Is DBNull.Value Then
    '            btnView.Enabled = False
    '        Else
    '            If e.Row.Cells("SCHEDULED_BY").Text = String.Empty Then
    '                btnView.Enabled = False
    '            Else
    '                btnView.Enabled = True
    '            End If
    '        End If
    '        'If e.Row.Cells("INSPECTION TYPE").Text = "Compliance Audit" Then
    '        '    btnView.Enabled = True
    '        'Else
    '        '    btnView.Enabled = False
    '        'End If
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    Private Sub ugHistory_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugHistory.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            btnView_Click(sender, e)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

End Class
