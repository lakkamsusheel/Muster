Public Class CMList
    Inherits System.Windows.Forms.Form
#Region "User Defined Variables"
    Private WithEvents oManager As MUSTER.BusinessLogic.pLicensee
    Private frmCM As Managers
    ' Private frmAssignedManager As AssignedManager
    Private rp As New Remove_Pencil
    Private bolLoading As Boolean = False
#End Region
#Region "Windows Form Designer generated code "

    Public Sub New(ByRef oMgr As MUSTER.BusinessLogic.pLicensee, ByVal dsCMList As DataSet, ByVal facility_ID As Integer)
        MyBase.New()
        bolLoading = True
        'This call is required by the Windows Form Designer.
        InitializeComponent()
        'Add any initialization after the InitializeComponent() call
        Cursor.Current = Cursors.AppStarting
        oManager = oMgr
        LoadCMList(dsCMList, facility_ID)
        bolLoading = False
        If dsCMList.Tables(0).Rows.Count > 0 Then
            ugCMList.Rows(0).Activate()
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
    Friend WithEvents pnlMgrCMListDetails As System.Windows.Forms.Panel
    Friend WithEvents pnlMgrCMListBottom As System.Windows.Forms.Panel
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnView As System.Windows.Forms.Button
    Friend WithEvents pnlMgrCMListHeader As System.Windows.Forms.Panel
    Friend WithEvents lblFacIDValue As System.Windows.Forms.Label
    Friend WithEvents lblCMList As System.Windows.Forms.Label
    Friend WithEvents lblFacID As System.Windows.Forms.Label
    Friend WithEvents pnlMgrCMListGrid As System.Windows.Forms.Panel
    Friend WithEvents ugCMList As Infragistics.Win.UltraWinGrid.UltraGrid
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlMgrCMListDetails = New System.Windows.Forms.Panel
        Me.pnlMgrCMListGrid = New System.Windows.Forms.Panel
        Me.ugCMList = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlMgrCMListHeader = New System.Windows.Forms.Panel
        Me.lblFacIDValue = New System.Windows.Forms.Label
        Me.lblCMList = New System.Windows.Forms.Label
        Me.lblFacID = New System.Windows.Forms.Label
        Me.pnlMgrCMListBottom = New System.Windows.Forms.Panel
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnView = New System.Windows.Forms.Button
        Me.pnlMgrCMListDetails.SuspendLayout()
        Me.pnlMgrCMListGrid.SuspendLayout()
        CType(Me.ugCMList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlMgrCMListHeader.SuspendLayout()
        Me.pnlMgrCMListBottom.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlMgrCMListDetails
        '
        Me.pnlMgrCMListDetails.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlMgrCMListDetails.Controls.Add(Me.pnlMgrCMListGrid)
        Me.pnlMgrCMListDetails.Controls.Add(Me.pnlMgrCMListHeader)
        Me.pnlMgrCMListDetails.Controls.Add(Me.pnlMgrCMListBottom)
        Me.pnlMgrCMListDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlMgrCMListDetails.Location = New System.Drawing.Point(0, 0)
        Me.pnlMgrCMListDetails.Name = "pnlMgrCMListDetails"
        Me.pnlMgrCMListDetails.Size = New System.Drawing.Size(504, 357)
        Me.pnlMgrCMListDetails.TabIndex = 0
        '
        'pnlMgrCMListGrid
        '
        Me.pnlMgrCMListGrid.Controls.Add(Me.ugCMList)
        Me.pnlMgrCMListGrid.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlMgrCMListGrid.Location = New System.Drawing.Point(0, 51)
        Me.pnlMgrCMListGrid.Name = "pnlMgrCMListGrid"
        Me.pnlMgrCMListGrid.Size = New System.Drawing.Size(500, 262)
        Me.pnlMgrCMListGrid.TabIndex = 7
        '
        'ugCMList
        '
        Me.ugCMList.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugCMList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugCMList.Location = New System.Drawing.Point(0, 0)
        Me.ugCMList.Name = "ugCMList"
        Me.ugCMList.Size = New System.Drawing.Size(500, 262)
        Me.ugCMList.TabIndex = 0
        '
        'pnlMgrCMListHeader
        '
        Me.pnlMgrCMListHeader.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlMgrCMListHeader.Controls.Add(Me.lblFacIDValue)
        Me.pnlMgrCMListHeader.Controls.Add(Me.lblCMList)
        Me.pnlMgrCMListHeader.Controls.Add(Me.lblFacID)
        Me.pnlMgrCMListHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlMgrCMListHeader.Location = New System.Drawing.Point(0, 0)
        Me.pnlMgrCMListHeader.Name = "pnlMgrCMListHeader"
        Me.pnlMgrCMListHeader.Size = New System.Drawing.Size(500, 51)
        Me.pnlMgrCMListHeader.TabIndex = 0
        '
        'lblFacIDValue
        '
        Me.lblFacIDValue.Location = New System.Drawing.Point(152, 8)
        Me.lblFacIDValue.Name = "lblFacIDValue"
        Me.lblFacIDValue.Size = New System.Drawing.Size(152, 16)
        Me.lblFacIDValue.TabIndex = 0
        '
        'lblCMList
        '
        Me.lblCMList.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCMList.Location = New System.Drawing.Point(4, 24)
        Me.lblCMList.Name = "lblCMList"
        Me.lblCMList.Size = New System.Drawing.Size(212, 24)
        Me.lblCMList.TabIndex = 0
        Me.lblCMList.Text = "Compliance Manager List"
        '
        'lblFacID
        '
        Me.lblFacID.Location = New System.Drawing.Point(8, 8)
        Me.lblFacID.Name = "lblFacID"
        Me.lblFacID.Size = New System.Drawing.Size(128, 17)
        Me.lblFacID.TabIndex = 0
        Me.lblFacID.Text = "Facility:"
        Me.lblFacID.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'pnlMgrCMListBottom
        '
        Me.pnlMgrCMListBottom.Controls.Add(Me.btnClose)
        Me.pnlMgrCMListBottom.Controls.Add(Me.btnView)
        Me.pnlMgrCMListBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlMgrCMListBottom.Location = New System.Drawing.Point(0, 313)
        Me.pnlMgrCMListBottom.Name = "pnlMgrCMListBottom"
        Me.pnlMgrCMListBottom.Size = New System.Drawing.Size(500, 40)
        Me.pnlMgrCMListBottom.TabIndex = 1
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
        'CMList
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(504, 357)
        Me.Controls.Add(Me.pnlMgrCMListDetails)
        Me.MinimizeBox = False
        Me.Name = "CMList"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Manager CMList"
        Me.pnlMgrCMListDetails.ResumeLayout(False)
        Me.pnlMgrCMListGrid.ResumeLayout(False)
        CType(Me.ugCMList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlMgrCMListHeader.ResumeLayout(False)
        Me.pnlMgrCMListBottom.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
#End Region
#Region "UI Support Routines"
    Private Sub LoadCMList(ByVal dsCMList As DataSet, ByVal facility_ID As Integer)
        Try

            ugCMList.DataSource = dsCMList
            ugCMList.DrawFilter = rp

            ugCMList.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
            ugCMList.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect
            ugCMList.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
            ugCMList.DisplayLayout.Bands(0).Columns("MGRFACRELATION_ID").Hidden = True
            ugCMList.DisplayLayout.Bands(0).Columns("Manager_ID").Hidden = False
            ugCMList.DisplayLayout.Bands(0).Columns("FACILITY_ID").Hidden = True
            ugCMList.DisplayLayout.Bands(0).Columns("RELATION_ID").Hidden = True
            ugCMList.DisplayLayout.Bands(0).Columns("RELATION_DESC").Hidden = False

            ugCMList.DisplayLayout.Bands(0).Columns("First_Name").Hidden = False
            ugCMList.DisplayLayout.Bands(0).Columns("Last_Name").Hidden = False
            ugCMList.DisplayLayout.Bands(0).Columns("DELETED").Hidden = True
            lblFacIDValue.Text = "#" + (ugCMList.Rows(0).Cells("FACILITY_ID").Value).ToString
            '  lblFacIDValue.Text = "#" + (ugCMList.Rows(0).Cells("FACILITY ID").Value).ToString + " - " + strFacilityName
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
            If ugCMList.Rows.Count > 0 Then
                If Not ugCMList.ActiveRow Is Nothing Then
                    ugrow = ugCMList.ActiveRow
                End If
            End If
            If ugrow Is Nothing Then
                MsgBox("Please select a Manager to view")
                Exit Sub
            End If
            '  oManager.Retrieve(CType(ugrow.Cells("Manager_ID").Value, Int64)

            frmCM = New Managers(ugrow.Cells("Manager_ID").Value)
            frmCM.WindowState = FormWindowState.Maximized
            frmCM.ShowDialog()

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

    Private Sub ugCMList_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugCMList.DoubleClick
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

    Private Sub lblCMList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lblCMList.Click

    End Sub
End Class