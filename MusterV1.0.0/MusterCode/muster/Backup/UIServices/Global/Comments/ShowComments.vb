'MR - 9/4/2004
Public Class ShowComments
    '-------------------------------------------------------------------------------
    ' MUSTER.MUSTER.ShowComment.vb
    '   Provides the mechanism for displaying comments for the app.
    '
    ' Copyright (C) 2004 CIBER, Inc.
    ' All rights reserved.
    '
    ' Release   Initials    Date        Description
    '  1.0        ??      8/??/04    Original class definition.
    '  1.1        JC      1/02/04    Changed AppUser.UserName to AppUser.ID to
    '                                  accomodate new use of pUser by application.
    '  1.2        JC      8/2/05     Changed rowfilter to append to existing rowfilter
    '
    '-------------------------------------------------------------------------------
    Inherits System.Windows.Forms.Form
    Dim ACfrm As AddComments
#Region "User Defined Variables"
    Friend dTableGlobal As New System.Data.DataTable("tGlobalComments")
    Friend dsComments As New System.Data.DataSet
    Dim result As DialogResult
    Friend ts As DataGridTableStyle
    Private WithEvents pCommentLocal As MUSTER.BusinessLogic.pComments
    Friend tblComments As DataTable
    Dim strModuleName As String = String.Empty
    Dim nEntityTypeID As Integer = 0
    Dim nEntityID As Int64 = 0
    Dim strEntityAddnInfo As String = String.Empty
    Dim returnVal As String = String.Empty
    Friend nCommentsCount As Integer = 0
#End Region
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub
    Public Sub New(ByVal EntityID As Int64, ByVal EntityType As Integer, ByVal strModule As String, ByVal strEntityName As String, Optional ByRef oComments As MUSTER.BusinessLogic.pComments = Nothing, Optional ByVal Parenttext As String = "", Optional ByVal entityAddnInfo As String = "", Optional ByVal enableShowAllModules As Boolean = True)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        '
        ' Getting the Current Entity Details.
        '
        Me.Text = Parenttext + " Comments"
        lblModuleName.Text = strModule
        nEntityTypeID = EntityType
        nEntityID = EntityID
        lblUserName.Text = MusterContainer.AppUser.ID
        strEntityAddnInfo = entityAddnInfo
        lblEntityName.Text = strEntityName
        pCommentLocal = oComments
        ' not disabling show all modules at event level. just left it here incase need to disable in future
        'chkBoxShowAllModules.Enabled = enableShowAllModules

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
    Friend WithEvents lblEntityName As System.Windows.Forms.Label
    Friend WithEvents lblComments As System.Windows.Forms.Label
    Friend WithEvents lblModuleName As System.Windows.Forms.Label
    Friend WithEvents lblEntityType As System.Windows.Forms.Label
    Friend WithEvents lblEntityID As System.Windows.Forms.Label
    Friend WithEvents dgComments As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents lblUserName As System.Windows.Forms.Label
    Friend WithEvents pnlCommentTop As System.Windows.Forms.Panel
    Friend WithEvents pnlCommentGrid As System.Windows.Forms.Panel
    Friend WithEvents pnlCommentBottom As System.Windows.Forms.Panel
    Friend WithEvents btnModifyComment As System.Windows.Forms.Button
    Friend WithEvents btnCommentCancel As System.Windows.Forms.Button
    Friend WithEvents btnCommentAdd As System.Windows.Forms.Button
    Friend WithEvents btnCommentDelete As System.Windows.Forms.Button
    Friend WithEvents chkBoxShowAllModules As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlCommentTop = New System.Windows.Forms.Panel
        Me.lblEntityName = New System.Windows.Forms.Label
        Me.lblComments = New System.Windows.Forms.Label
        Me.lblEntityID = New System.Windows.Forms.Label
        Me.lblEntityType = New System.Windows.Forms.Label
        Me.lblModuleName = New System.Windows.Forms.Label
        Me.lblUserName = New System.Windows.Forms.Label
        Me.dgComments = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlCommentGrid = New System.Windows.Forms.Panel
        Me.pnlCommentBottom = New System.Windows.Forms.Panel
        Me.chkBoxShowAllModules = New System.Windows.Forms.CheckBox
        Me.btnModifyComment = New System.Windows.Forms.Button
        Me.btnCommentCancel = New System.Windows.Forms.Button
        Me.btnCommentAdd = New System.Windows.Forms.Button
        Me.btnCommentDelete = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.pnlCommentTop.SuspendLayout()
        CType(Me.dgComments, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlCommentGrid.SuspendLayout()
        Me.pnlCommentBottom.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlCommentTop
        '
        Me.pnlCommentTop.Controls.Add(Me.lblEntityName)
        Me.pnlCommentTop.Controls.Add(Me.lblComments)
        Me.pnlCommentTop.Controls.Add(Me.lblEntityID)
        Me.pnlCommentTop.Controls.Add(Me.lblEntityType)
        Me.pnlCommentTop.Controls.Add(Me.lblModuleName)
        Me.pnlCommentTop.Controls.Add(Me.lblUserName)
        Me.pnlCommentTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlCommentTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlCommentTop.Name = "pnlCommentTop"
        Me.pnlCommentTop.Size = New System.Drawing.Size(816, 32)
        Me.pnlCommentTop.TabIndex = 0
        '
        'lblEntityName
        '
        Me.lblEntityName.Font = New System.Drawing.Font("Trebuchet MS", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEntityName.Location = New System.Drawing.Point(104, 8)
        Me.lblEntityName.Name = "lblEntityName"
        Me.lblEntityName.Size = New System.Drawing.Size(456, 16)
        Me.lblEntityName.TabIndex = 114
        '
        'lblComments
        '
        Me.lblComments.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblComments.Location = New System.Drawing.Point(8, 8)
        Me.lblComments.Name = "lblComments"
        Me.lblComments.Size = New System.Drawing.Size(88, 16)
        Me.lblComments.TabIndex = 113
        Me.lblComments.Text = "Comments for: "
        '
        'lblEntityID
        '
        Me.lblEntityID.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEntityID.Location = New System.Drawing.Point(568, 8)
        Me.lblEntityID.Name = "lblEntityID"
        Me.lblEntityID.Size = New System.Drawing.Size(16, 16)
        Me.lblEntityID.TabIndex = 117
        Me.lblEntityID.Visible = False
        '
        'lblEntityType
        '
        Me.lblEntityType.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEntityType.Location = New System.Drawing.Point(592, 8)
        Me.lblEntityType.Name = "lblEntityType"
        Me.lblEntityType.Size = New System.Drawing.Size(16, 16)
        Me.lblEntityType.TabIndex = 116
        Me.lblEntityType.Visible = False
        '
        'lblModuleName
        '
        Me.lblModuleName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblModuleName.Location = New System.Drawing.Point(624, 8)
        Me.lblModuleName.Name = "lblModuleName"
        Me.lblModuleName.Size = New System.Drawing.Size(16, 16)
        Me.lblModuleName.TabIndex = 115
        Me.lblModuleName.Visible = False
        '
        'lblUserName
        '
        Me.lblUserName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUserName.Location = New System.Drawing.Point(656, 8)
        Me.lblUserName.Name = "lblUserName"
        Me.lblUserName.Size = New System.Drawing.Size(16, 16)
        Me.lblUserName.TabIndex = 120
        Me.lblUserName.Visible = False
        '
        'dgComments
        '
        Me.dgComments.Cursor = System.Windows.Forms.Cursors.Default
        Me.dgComments.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgComments.Location = New System.Drawing.Point(0, 0)
        Me.dgComments.Name = "dgComments"
        Me.dgComments.Size = New System.Drawing.Size(816, 397)
        Me.dgComments.TabIndex = 119
        Me.dgComments.Text = "Comments"
        '
        'pnlCommentGrid
        '
        Me.pnlCommentGrid.Controls.Add(Me.dgComments)
        Me.pnlCommentGrid.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlCommentGrid.Location = New System.Drawing.Point(0, 32)
        Me.pnlCommentGrid.Name = "pnlCommentGrid"
        Me.pnlCommentGrid.Size = New System.Drawing.Size(816, 397)
        Me.pnlCommentGrid.TabIndex = 123
        '
        'pnlCommentBottom
        '
        Me.pnlCommentBottom.Controls.Add(Me.Label1)
        Me.pnlCommentBottom.Controls.Add(Me.chkBoxShowAllModules)
        Me.pnlCommentBottom.Controls.Add(Me.btnModifyComment)
        Me.pnlCommentBottom.Controls.Add(Me.btnCommentCancel)
        Me.pnlCommentBottom.Controls.Add(Me.btnCommentAdd)
        Me.pnlCommentBottom.Controls.Add(Me.btnCommentDelete)
        Me.pnlCommentBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlCommentBottom.Location = New System.Drawing.Point(0, 429)
        Me.pnlCommentBottom.Name = "pnlCommentBottom"
        Me.pnlCommentBottom.Size = New System.Drawing.Size(816, 40)
        Me.pnlCommentBottom.TabIndex = 124
        '
        'chkBoxShowAllModules
        '
        Me.chkBoxShowAllModules.Location = New System.Drawing.Point(656, 8)
        Me.chkBoxShowAllModules.Name = "chkBoxShowAllModules"
        Me.chkBoxShowAllModules.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkBoxShowAllModules.Size = New System.Drawing.Size(113, 24)
        Me.chkBoxShowAllModules.TabIndex = 119
        Me.chkBoxShowAllModules.Text = "Show All Modules"
        '
        'btnModifyComment
        '
        Me.btnModifyComment.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnModifyComment.Location = New System.Drawing.Point(304, 8)
        Me.btnModifyComment.Name = "btnModifyComment"
        Me.btnModifyComment.Size = New System.Drawing.Size(112, 26)
        Me.btnModifyComment.TabIndex = 118
        Me.btnModifyComment.Text = "Modify Comment"
        '
        'btnCommentCancel
        '
        Me.btnCommentCancel.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCommentCancel.Location = New System.Drawing.Point(544, 8)
        Me.btnCommentCancel.Name = "btnCommentCancel"
        Me.btnCommentCancel.Size = New System.Drawing.Size(75, 26)
        Me.btnCommentCancel.TabIndex = 107
        Me.btnCommentCancel.Text = "Close"
        '
        'btnCommentAdd
        '
        Me.btnCommentAdd.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCommentAdd.Location = New System.Drawing.Point(200, 8)
        Me.btnCommentAdd.Name = "btnCommentAdd"
        Me.btnCommentAdd.Size = New System.Drawing.Size(96, 26)
        Me.btnCommentAdd.TabIndex = 106
        Me.btnCommentAdd.Text = "Add Comment"
        '
        'btnCommentDelete
        '
        Me.btnCommentDelete.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCommentDelete.Location = New System.Drawing.Point(424, 8)
        Me.btnCommentDelete.Name = "btnCommentDelete"
        Me.btnCommentDelete.Size = New System.Drawing.Size(112, 26)
        Me.btnCommentDelete.TabIndex = 108
        Me.btnCommentDelete.Text = "Delete Comment"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Yellow
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(16, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(88, 16)
        Me.Label1.TabIndex = 120
        Me.Label1.Text = "Migrated Row"
        '
        'ShowComments
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(816, 469)
        Me.ControlBox = False
        Me.Controls.Add(Me.pnlCommentGrid)
        Me.Controls.Add(Me.pnlCommentTop)
        Me.Controls.Add(Me.pnlCommentBottom)
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.IsMdiContainer = True
        Me.Name = "ShowComments"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ShowComments"
        Me.pnlCommentTop.ResumeLayout(False)
        CType(Me.dgComments, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlCommentGrid.ResumeLayout(False)
        Me.pnlCommentBottom.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private Sub ShowComments_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            dgComments.DataSource = pCommentLocal.GetComments(lblModuleName.Text, nEntityTypeID, CInt(nEntityID), strEntityAddnInfo)
            SetupGrid()
            nCommentsCount = dgComments.Rows.Count
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCommentAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCommentAdd.Click
        Try
            If IsNothing(ACfrm) Then
                ACfrm = New AddComments(nEntityID, nEntityTypeID, lblModuleName.Text, lblEntityName.Text, pCommentLocal, , , strEntityAddnInfo)
                AddHandler ACfrm.Closing, AddressOf frmAddCommentClosing
                AddHandler ACfrm.Closed, AddressOf frmAddCommentClosed
            End If
            ACfrm.ShowDialog()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub frmAddCommentClosing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        Try
            dgComments.DataSource = pCommentLocal.GetComments(IIf(chkBoxShowAllModules.Checked, String.Empty, lblModuleName.Text), nEntityTypeID, CInt(nEntityID), strEntityAddnInfo)
            SetupGrid()
            nCommentsCount = dgComments.Rows.Count
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub frmAddCommentClosed(ByVal sender As Object, ByVal e As System.EventArgs)
        ACfrm = Nothing
    End Sub
    Private Sub btnCommentCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCommentCancel.Click
        Me.Close()
    End Sub
    Private Sub btnCommentDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCommentDelete.Click
        Try
            If dgComments.Rows.Count <= 0 Then Exit Sub

            If dgComments.ActiveRow Is Nothing Then
                MsgBox("Select row to Delete")
                Exit Sub
            Else
                result = MessageBox.Show("Are you Sure you want to Delete this Record?", "Comments", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If result = DialogResult.No Then
                    Exit Sub
                Else
                    pCommentLocal.Retrieve(Integer.Parse(dgComments.ActiveRow.Cells("COMMENT_ID").Text), dgComments.ActiveRow.Cells("USER ID").Text)
                    pCommentLocal.Deleted = True
                    If pCommentLocal.ID <= 0 Then
                        pCommentLocal.CreatedBy = MusterContainer.AppUser.ID
                    Else
                        pCommentLocal.ModifiedBy = MusterContainer.AppUser.ID
                    End If
                    pCommentLocal.Save(CType(UIUtilsGen.ModuleID.[Global], Integer), MusterContainer.AppUser.UserKey, returnVal)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If
                    pCommentLocal.Remove(Integer.Parse(dgComments.ActiveRow.Cells("COMMENT_ID").Text))
                    dgComments.ActiveRow.Delete(False)
                    nCommentsCount = dgComments.Rows.Count
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub chkboxShowAllModules_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkBoxShowAllModules.CheckedChanged
        Try
            If chkBoxShowAllModules.Checked = True Then
                dgComments.DataSource = pCommentLocal.GetComments(, nEntityTypeID, nEntityID, strEntityAddnInfo)
            Else
                dgComments.DataSource = pCommentLocal.GetComments(lblModuleName.Text, nEntityTypeID, nEntityID, strEntityAddnInfo)
            End If
            SetupGrid()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnModifyComment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnModifyComment.Click
        Try
            If Not dgComments.Rows.Count > 0 Then Exit Sub
            If dgComments.ActiveRow Is Nothing Then
                MsgBox("Select row to Modify.")
                Exit Sub
            End If
            If IsNothing(ACfrm) Then
                ACfrm = New AddComments(dgComments.ActiveRow.Cells("ENTITY ID").Value, dgComments.ActiveRow.Cells("ENTITY_TYPE").Value, dgComments.ActiveRow.Cells("MODULE").Value, lblEntityName.Text, pCommentLocal, dgComments.ActiveRow.Cells("COMMENT_ID").Value, dgComments.ActiveRow, IIf(dgComments.ActiveRow.Cells("ENTITY_ADDITIONAL_INFO").Value Is DBNull.Value, String.Empty, dgComments.ActiveRow.Cells("ENTITY_ADDITIONAL_INFO").Value))
                AddHandler ACfrm.Closing, AddressOf frmAddCommentClosing
                AddHandler ACfrm.Closed, AddressOf frmAddCommentClosed
            End If
            ACfrm.ShowDialog()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub dgComments_AfterSelectChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.AfterSelectChangeEventArgs) Handles dgComments.AfterSelectChange
        ProcessRowTest()
    End Sub
    Private Sub dgComments_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles dgComments.DoubleClick
        Try
            If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
                Me.Cursor = Cursors.Default
                Exit Sub
            End If
            ProcessRowTest()
            If btnModifyComment.Enabled Then
                btnModifyComment_Click(sender, e)
            Else
                MsgBox("You Cannot Modify the selected row")
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ProcessRowTest()
        Dim dtSuperUsers As DataTable
        Dim drow As DataRow
        Dim bolEnableFlag As Boolean = False
        Try
            With dgComments.Selected
                If .Rows.Count > 0 Then
                    If MusterContainer.AppUser.ID.Trim.ToUpper = dgComments.ActiveRow.Cells("USER ID").Text.Trim.ToUpper Then
                        btnCommentDelete.Enabled = True
                        btnModifyComment.Enabled = True
                        Exit Sub
                    Else
                        dtSuperUsers = MusterContainer.AppUser.ListSupervisedUsers()
                        If dtSuperUsers.Rows.Count > 0 Then
                            For Each drow In dtSuperUsers.Rows
                                If drow("USER_ID").ToString.Trim.ToUpper = dgComments.ActiveRow.Cells("USER ID").Text.Trim.ToUpper Then
                                    bolEnableFlag = True
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                Else
                    If dgComments.Rows.Count > 0 Then
                        If MusterContainer.AppUser.ID.Trim.ToUpper = dgComments.Rows(0).Cells("USER ID").Text.Trim.ToUpper Then
                            btnCommentDelete.Enabled = True
                            btnModifyComment.Enabled = True
                            Exit Sub
                        Else
                            dtSuperUsers = MusterContainer.AppUser.ListSupervisedUsers()
                            If dtSuperUsers.Rows.Count > 0 Then
                                For Each drow In dtSuperUsers.Rows
                                    If drow("USER_ID").ToString.Trim.ToUpper = dgComments.Rows(0).Cells("USER ID").Text.Trim.ToUpper Then
                                        bolEnableFlag = True
                                        Exit For
                                    End If
                                Next
                            End If
                        End If
                    Else
                        Exit Sub
                    End If
                End If
            End With
            If bolEnableFlag Then
                btnCommentDelete.Enabled = True
                btnModifyComment.Enabled = True
            Else
                btnCommentDelete.Enabled = False
                btnModifyComment.Enabled = False
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub SetupGrid()
        Try
            dgComments.DisplayLayout.Bands(0).Columns("ENTITY_TYPE").Hidden = True
            dgComments.DisplayLayout.Bands(0).Columns("CREATEDBY").Hidden = True
            dgComments.DisplayLayout.Bands(0).Columns("COMMENT_DATE").Hidden = True
            dgComments.DisplayLayout.Bands(0).Columns("ENTITY_ADDITIONAL_INFO").Hidden = True
            dgComments.DisplayLayout.Bands(0).Columns("COMMENT_ID").Hidden = True '9
            dgComments.DisplayLayout.Bands(0).Columns("DELETED").Hidden = True '8
            dgComments.DisplayLayout.Bands(0).Columns("LAST_EDITED_BY").Hidden = True
            dgComments.DisplayLayout.Bands(0).Columns("DATE_LAST_EDITED").Hidden = True

            dgComments.DisplayLayout.Bands(0).Columns("CREATED ON").Width = 90
            dgComments.DisplayLayout.Bands(0).Columns("COMMENT").Width = 275
            dgComments.DisplayLayout.Bands(0).Columns("MODULE").Width = 80
            dgComments.DisplayLayout.Bands(0).Columns("USER ID").Width = 75
            dgComments.DisplayLayout.Bands(0).Columns("VIEWABLE BY").Width = 100
            dgComments.DisplayLayout.Bands(0).Columns("COMMENT").CellMultiLine = Infragistics.Win.DefaultableBoolean.True
            dgComments.DisplayLayout.Bands(0).Columns("COMMENT").VertScrollBar = True
            dgComments.DisplayLayout.Bands(0).Columns("COMMENT").AutoSizeEdit = Infragistics.Win.DefaultableBoolean.True

            dgComments.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.Free
            dgComments.DisplayLayout.Override.RowSizingArea = Infragistics.Win.UltraWinGrid.RowSizingArea.EntireRow
            dgComments.DisplayLayout.Bands(0).Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.AutoFree
            dgComments.DisplayLayout.Bands(0).Override.RowSizingAutoMaxLines = 8

            dgComments.DisplayLayout.Bands(0).Columns("CREATED ON").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
            dgComments.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False
            dgComments.DisplayLayout.Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect

            dgComments.DisplayLayout.Bands(0).Columns("CREATED ON").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Descending
            'dgComments.DisplayLayout.Bands(0).Columns("ENTITY TYPE").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            'dgComments.DisplayLayout.Bands(0).Columns("MODULE").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            'dgComments.DisplayLayout.Bands(0).Columns("USER ID").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            'dgComments.DisplayLayout.Bands(0).Columns("COMMENT").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            'dgComments.DisplayLayout.Bands(0).Columns("VIEWABLE BY").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
            'dgComments.DisplayLayout.Bands(0).Columns("ENTITY ID").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending

            If dgComments.Rows.Count > 0 Then
                ' 2675 if comment was created before golive (oct 13 2007) change created on background color to indicate migrated row
                For Each ugrow As Infragistics.Win.UltraWinGrid.UltraGridRow In dgComments.Rows
                    If Not ugrow.Cells("CREATED ON").Value Is DBNull.Value Then
                        If Date.Compare(CDate("10/13/2006"), ugrow.Cells("CREATED ON").Value) >= 0 Then
                            ugrow.Appearance.BackColor = Color.Yellow
                        End If
                    End If
                Next
                ProcessRowTest()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Friend Function GetCounts() As Integer
        Dim bolGridHasRows As Boolean = False
        If dgComments Is Nothing Then
            bolGridHasRows = False
        ElseIf dgComments.DataSource Is Nothing Then
            bolGridHasRows = False
        End If
        If bolGridHasRows Then
            nCommentsCount = dgComments.Rows.Count
        Else
            nCommentsCount = pCommentLocal.GetComments(lblModuleName.Text, nEntityTypeID, CInt(nEntityID), strEntityAddnInfo).Tables(0).Rows.Count
        End If
        Return nCommentsCount
    End Function

    Private Sub dgComments_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles dgComments.MouseDown
        If e.Button = MouseButtons.Right Then
            Dim row As Infragistics.Win.UltraWinGrid.UltraGridRow = Nothing
            Dim element As Infragistics.Win.UIElement = Nothing
            Dim point As New System.Drawing.Point(e.X, e.Y)
            element = dgComments.DisplayLayout.UIElement.ElementFromPoint(point)
            Try
                row = element.SelectableItem
                If Not row Is Nothing Then
                    dgComments.ActiveRow = row
                    MessageBoxCustom.Show(row.Cells("COMMENT").Text, "View Comment", MessageBoxButtons.OK, MessageBoxIcon.Information, , KnownColor.White)
                End If
            Catch ex As Exception
                Exit Try
            End Try
        End If
    End Sub
End Class
