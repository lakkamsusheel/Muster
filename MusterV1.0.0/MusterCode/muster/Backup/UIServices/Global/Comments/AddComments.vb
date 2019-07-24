'MR - 8/16/2004
Public Class AddComments
    Inherits System.Windows.Forms.Form
    Friend WithEvents frmRegServices As MusterContainer
    Friend drow As Infragistics.Win.UltraWinGrid.UltraGridRow
    'Newly Added Variables
    Private WithEvents pCommentLocal As MUSTER.BusinessLogic.pComments
    Private Selectedrow As Infragistics.Win.UltraWinGrid.UltraGridRow
    Dim nCommentID As Integer = 0
    Dim nEntityTypeID As Integer = 0
    Dim nEntityID As Int64 = 0
    Dim strEntityAddnInfo As String = String.Empty
    Dim returnVal As String = String.Empty
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub
    Public Sub New(ByVal EntityID As Int64, ByVal EntityType As Integer, ByVal strModule As String, ByVal strEntityName As String, Optional ByRef pComments As MUSTER.BusinessLogic.pComments = Nothing, Optional ByVal CommentID As Integer = 0, Optional ByVal Selrow As Infragistics.Win.UltraWinGrid.UltraGridRow = Nothing, Optional ByVal entityAddnInfo As String = "")
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        '
        ' Getting the Current Entity Details.
        '
        Me.Text = "Add Comments"
        lblModuleName.Text = strModule
        nEntityTypeID = EntityType
        nEntityID = EntityID
        strEntityAddnInfo = entityAddnInfo
        lblEntityName.Text = strEntityName
        lblUserName.Text = MusterContainer.AppUser.ID
        nCommentID = CommentID
        pCommentLocal = pComments
        Selectedrow = Selrow

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
    Friend WithEvents lblComment As System.Windows.Forms.Label
    Friend WithEvents grpScope As System.Windows.Forms.GroupBox
    Friend WithEvents rdInternal As System.Windows.Forms.RadioButton
    Friend WithEvents rdExternal As System.Windows.Forms.RadioButton
    Friend WithEvents lblCommentFor As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSaveComment As System.Windows.Forms.Button
    Friend WithEvents txtComment As System.Windows.Forms.TextBox
    Friend WithEvents lblEntityType As System.Windows.Forms.Label
    Friend WithEvents lblModuleName As System.Windows.Forms.Label
    Friend WithEvents lblEntityName As System.Windows.Forms.Label
    Friend WithEvents lblEntityID As System.Windows.Forms.Label
    Friend WithEvents lblModule As System.Windows.Forms.Label
    Friend WithEvents lblUserName As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.lblComment = New System.Windows.Forms.Label
        Me.grpScope = New System.Windows.Forms.GroupBox
        Me.rdInternal = New System.Windows.Forms.RadioButton
        Me.rdExternal = New System.Windows.Forms.RadioButton
        Me.lblCommentFor = New System.Windows.Forms.Label
        Me.lblEntityName = New System.Windows.Forms.Label
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSaveComment = New System.Windows.Forms.Button
        Me.txtComment = New System.Windows.Forms.TextBox
        Me.lblEntityType = New System.Windows.Forms.Label
        Me.lblModuleName = New System.Windows.Forms.Label
        Me.lblEntityID = New System.Windows.Forms.Label
        Me.lblModule = New System.Windows.Forms.Label
        Me.lblUserName = New System.Windows.Forms.Label
        Me.grpScope.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblComment
        '
        Me.lblComment.Location = New System.Drawing.Point(24, 72)
        Me.lblComment.Name = "lblComment"
        Me.lblComment.Size = New System.Drawing.Size(64, 16)
        Me.lblComment.TabIndex = 35
        Me.lblComment.Text = "Comment"
        '
        'grpScope
        '
        Me.grpScope.Controls.Add(Me.rdInternal)
        Me.grpScope.Controls.Add(Me.rdExternal)
        Me.grpScope.Location = New System.Drawing.Point(24, 264)
        Me.grpScope.Name = "grpScope"
        Me.grpScope.Size = New System.Drawing.Size(128, 72)
        Me.grpScope.TabIndex = 1
        Me.grpScope.TabStop = False
        Me.grpScope.Text = "Viewable By"
        '
        'rdInternal
        '
        Me.rdInternal.Checked = True
        Me.rdInternal.Location = New System.Drawing.Point(24, 40)
        Me.rdInternal.Name = "rdInternal"
        Me.rdInternal.Size = New System.Drawing.Size(80, 24)
        Me.rdInternal.TabIndex = 1
        Me.rdInternal.TabStop = True
        Me.rdInternal.Text = "Internal"
        '
        'rdExternal
        '
        Me.rdExternal.Location = New System.Drawing.Point(24, 16)
        Me.rdExternal.Name = "rdExternal"
        Me.rdExternal.Size = New System.Drawing.Size(80, 24)
        Me.rdExternal.TabIndex = 0
        Me.rdExternal.Text = "External"
        '
        'lblCommentFor
        '
        Me.lblCommentFor.Location = New System.Drawing.Point(24, 16)
        Me.lblCommentFor.Name = "lblCommentFor"
        Me.lblCommentFor.Size = New System.Drawing.Size(80, 16)
        Me.lblCommentFor.TabIndex = 45
        Me.lblCommentFor.Text = "Comment for "
        '
        'lblEntityName
        '
        Me.lblEntityName.Font = New System.Drawing.Font("Trebuchet MS", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEntityName.Location = New System.Drawing.Point(96, 16)
        Me.lblEntityName.Name = "lblEntityName"
        Me.lblEntityName.Size = New System.Drawing.Size(248, 24)
        Me.lblEntityName.TabIndex = 46
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(200, 352)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(80, 23)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Text = "Cancel"
        '
        'btnSaveComment
        '
        Me.btnSaveComment.Location = New System.Drawing.Point(88, 352)
        Me.btnSaveComment.Name = "btnSaveComment"
        Me.btnSaveComment.Size = New System.Drawing.Size(96, 23)
        Me.btnSaveComment.TabIndex = 2
        Me.btnSaveComment.Text = "Save Comment"
        '
        'txtComment
        '
        Me.txtComment.Location = New System.Drawing.Point(24, 96)
        Me.txtComment.Multiline = True
        Me.txtComment.Name = "txtComment"
        Me.txtComment.Size = New System.Drawing.Size(312, 152)
        Me.txtComment.TabIndex = 0
        Me.txtComment.Text = ""
        '
        'lblEntityType
        '
        Me.lblEntityType.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEntityType.Location = New System.Drawing.Point(296, 48)
        Me.lblEntityType.Name = "lblEntityType"
        Me.lblEntityType.Size = New System.Drawing.Size(16, 16)
        Me.lblEntityType.TabIndex = 118
        Me.lblEntityType.Visible = False
        '
        'lblModuleName
        '
        Me.lblModuleName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblModuleName.Location = New System.Drawing.Point(80, 48)
        Me.lblModuleName.Name = "lblModuleName"
        Me.lblModuleName.Size = New System.Drawing.Size(128, 16)
        Me.lblModuleName.TabIndex = 117
        Me.lblModuleName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lblEntityID
        '
        Me.lblEntityID.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEntityID.Location = New System.Drawing.Point(328, 0)
        Me.lblEntityID.Name = "lblEntityID"
        Me.lblEntityID.Size = New System.Drawing.Size(16, 16)
        Me.lblEntityID.TabIndex = 119
        Me.lblEntityID.Visible = False
        '
        'lblModule
        '
        Me.lblModule.Location = New System.Drawing.Point(24, 48)
        Me.lblModule.Name = "lblModule"
        Me.lblModule.Size = New System.Drawing.Size(56, 16)
        Me.lblModule.TabIndex = 120
        Me.lblModule.Text = "Module :"
        '
        'lblUserName
        '
        Me.lblUserName.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblUserName.Location = New System.Drawing.Point(336, 264)
        Me.lblUserName.Name = "lblUserName"
        Me.lblUserName.Size = New System.Drawing.Size(16, 16)
        Me.lblUserName.TabIndex = 121
        Me.lblUserName.Visible = False
        '
        'AddComments
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(352, 390)
        Me.ControlBox = False
        Me.Controls.Add(Me.lblUserName)
        Me.Controls.Add(Me.lblModule)
        Me.Controls.Add(Me.lblEntityID)
        Me.Controls.Add(Me.lblEntityType)
        Me.Controls.Add(Me.lblModuleName)
        Me.Controls.Add(Me.txtComment)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSaveComment)
        Me.Controls.Add(Me.lblEntityName)
        Me.Controls.Add(Me.lblCommentFor)
        Me.Controls.Add(Me.grpScope)
        Me.Controls.Add(Me.lblComment)
        Me.Name = "AddComments"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "AddComments"
        Me.grpScope.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
#End Region
    Private Sub btnSaveComment_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveComment.Click
        Dim oCommentInfo As MUSTER.Info.CommentsInfo
        Try
            If nCommentID > 0 Then
                oCommentInfo = pCommentLocal.Retrieve(nCommentID)
                pCommentLocal.EntityID = nEntityID
                pCommentLocal.EntityAdditionalInfo = strEntityAddnInfo
                pCommentLocal.EntityType = nEntityTypeID
                pCommentLocal.Comments = txtComment.Text
                pCommentLocal.CommentsScope = IIf(rdExternal.Checked, "External", "Internal")
                pCommentLocal.UserID = lblUserName.Text
                pCommentLocal.CommentDate = Date.Now()
                pCommentLocal.ModuleName = lblModuleName.Text
                pCommentLocal.ModifiedBy = MusterContainer.AppUser.ID
            Else
                oCommentInfo = New MUSTER.Info.CommentsInfo
                oCommentInfo.EntityID = nEntityID
                oCommentInfo.EntityAdditionalInfo = strEntityAddnInfo
                oCommentInfo.EntityType = nEntityTypeID
                oCommentInfo.Comments = txtComment.Text
                oCommentInfo.CommentsScope = IIf(rdExternal.Checked, "External", "Internal")
                oCommentInfo.UserID = lblUserName.Text
                oCommentInfo.CommentDate = Date.Now()
                oCommentInfo.ModuleName = lblModuleName.Text
                oCommentInfo.CreatedBy = MusterContainer.AppUser.ID
                pCommentLocal.Add(oCommentInfo)
            End If
            pCommentLocal.Save(CType(UIUtilsGen.ModuleID.[Global], Integer), MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If
            nCommentID = pCommentLocal.ID
            Me.Close()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        pCommentLocal.Reset()
        Me.Close()
    End Sub
    Private Sub AddComments_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            If nCommentID > 0 Then
                ' Modify Comment
                Me.Text = "Modify Comment"
                lblModuleName.Text = Selectedrow.Cells("MODULE").Value
                lblEntityID.Text = Selectedrow.Cells("ENTITY ID").Value
                lblEntityType.Text = Selectedrow.Cells("ENTITY_TYPE").Value
                txtComment.Text = Selectedrow.Cells("COMMENT").Value
                lblUserName.Text = Selectedrow.Cells("USER ID").Value
                If Selectedrow.Cells("VIEWABLE BY").Value = "External" Then
                    rdInternal.Checked = False
                    rdExternal.Checked = True
                Else
                    rdInternal.Checked = True
                    rdExternal.Checked = False
                End If
            End If
            txtComment.Focus()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
End Class
