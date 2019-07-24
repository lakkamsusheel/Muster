Public Class FinTextAdmin
    Inherits System.Windows.Forms.Form

    Friend ReasonType As Int64
    Private oFinText As New MUSTER.BusinessLogic.pFinancialText
    Private bolLoading As Boolean = False
    Dim returnVal As String = String.Empty

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

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
    Friend WithEvents cmbText As System.Windows.Forms.ComboBox
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents lblTitle As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents lblText As System.Windows.Forms.Label
    Friend WithEvents lblName As System.Windows.Forms.Label
    Friend WithEvents btnAddNew As System.Windows.Forms.Button
    Friend WithEvents txtText As System.Windows.Forms.TextBox
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents chkActive As System.Windows.Forms.CheckBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmbText = New System.Windows.Forms.ComboBox
        Me.btnSave = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.lblTitle = New System.Windows.Forms.Label
        Me.btnAddNew = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.chkActive = New System.Windows.Forms.CheckBox
        Me.lblText = New System.Windows.Forms.Label
        Me.lblName = New System.Windows.Forms.Label
        Me.txtText = New System.Windows.Forms.TextBox
        Me.txtName = New System.Windows.Forms.TextBox
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmbText
        '
        Me.cmbText.Location = New System.Drawing.Point(152, 40)
        Me.cmbText.Name = "cmbText"
        Me.cmbText.Size = New System.Drawing.Size(456, 21)
        Me.cmbText.TabIndex = 1
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(328, 304)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(80, 23)
        Me.btnSave.TabIndex = 5
        Me.btnSave.Text = "Save"
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(416, 304)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(80, 23)
        Me.btnDelete.TabIndex = 6
        Me.btnDelete.Text = "Delete"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(504, 304)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(80, 23)
        Me.btnCancel.TabIndex = 7
        Me.btnCancel.Text = "Cancel"
        '
        'lblTitle
        '
        Me.lblTitle.Location = New System.Drawing.Point(16, 40)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(128, 23)
        Me.lblTitle.TabIndex = 6
        '
        'btnAddNew
        '
        Me.btnAddNew.Location = New System.Drawing.Point(0, 0)
        Me.btnAddNew.Name = "btnAddNew"
        Me.btnAddNew.Size = New System.Drawing.Size(64, 16)
        Me.btnAddNew.TabIndex = 9
        Me.btnAddNew.Text = "Add New"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.chkActive)
        Me.GroupBox1.Controls.Add(Me.lblText)
        Me.GroupBox1.Controls.Add(Me.lblName)
        Me.GroupBox1.Controls.Add(Me.txtText)
        Me.GroupBox1.Controls.Add(Me.txtName)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 72)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(608, 224)
        Me.GroupBox1.TabIndex = 10
        Me.GroupBox1.TabStop = False
        '
        'chkActive
        '
        Me.chkActive.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkActive.Checked = True
        Me.chkActive.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkActive.Location = New System.Drawing.Point(8, 80)
        Me.chkActive.Name = "chkActive"
        Me.chkActive.Size = New System.Drawing.Size(80, 24)
        Me.chkActive.TabIndex = 4
        Me.chkActive.Text = "Active"
        '
        'lblText
        '
        Me.lblText.Location = New System.Drawing.Point(8, 48)
        Me.lblText.Name = "lblText"
        Me.lblText.Size = New System.Drawing.Size(120, 23)
        Me.lblText.TabIndex = 12
        Me.lblText.Text = "Text"
        '
        'lblName
        '
        Me.lblName.Location = New System.Drawing.Point(8, 16)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(120, 23)
        Me.lblName.TabIndex = 11
        Me.lblName.Text = "Name"
        '
        'txtText
        '
        Me.txtText.Location = New System.Drawing.Point(144, 48)
        Me.txtText.Multiline = True
        Me.txtText.Name = "txtText"
        Me.txtText.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtText.Size = New System.Drawing.Size(452, 152)
        Me.txtText.TabIndex = 3
        Me.txtText.Text = ""
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(144, 16)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(452, 20)
        Me.txtName.TabIndex = 2
        Me.txtName.Text = ""
        '
        'FinTextAdmin
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(624, 342)
        Me.Controls.Add(Me.btnAddNew)
        Me.Controls.Add(Me.lblTitle)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.cmbText)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "FinTextAdmin"
        Me.Text = "FinTextAdmin"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


#Region " Page Events "
    Private Sub btnAddNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddNew.Click
        SetForm_AddNew()

    End Sub



    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try


            If oFinText.ID <= 0 Then
                oFinText.CreatedBy = MusterContainer.AppUser.ID
            Else
                oFinText.ModifiedBy = MusterContainer.AppUser.ID
            End If
            oFinText.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            MsgBox("System Text Saved")
            lblTitle.Visible = True
            cmbText.Visible = True
            bolLoading = True
            LoadcmbText()
            bolLoading = False
            SetForm_Update()

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Populate Tec Activity Documents" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub


    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            lblTitle.Visible = True
            cmbText.Visible = True
            SetForm_Update()
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Populate Tec Activity Documents" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try
            oFinText.Deleted = True
            oFinText.ModifiedBy = MusterContainer.AppUser.ID
            oFinText.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            MsgBox("System Text Saved")
            bolLoading = True
            LoadcmbText()
            bolLoading = False
            SetForm_Update()

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Populate Tec Activity Documents" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub FinTextAdmin_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            Select Case ReasonType
                Case 983
                    Me.Text = "Manage - Primary Condition of Reimbursement"
                Case 984
                    Me.Text = "Manage - Additional Condition of Reimbursement"
                Case 986
                    Me.Text = "Manage - Incomplete Application Reason"
                Case 987
                    Me.Text = "Manage - Deduction Reason"
            End Select
            bolLoading = True
            LoadcmbText()
            SetForm_Update()

            bolLoading = False
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Populate Tec Activity Documents" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try


    End Sub

#End Region

#Region " Processing code "
    Private Sub SetForm_AddNew()
        Try '
            lblTitle.Visible = False
            cmbText.Visible = False

            oFinText.Retrieve(0)
            oFinText.Text_Type = ReasonType
            oFinText.Active = True
            bolLoading = True
            txtName.Text = ""
            txtText.Text = ""
            chkActive.Checked = True

            bolLoading = False

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Populate Tec Activity Documents" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub SetForm_Update()

        Try
            If Not IsNothing(cmbText.SelectedValue) Then
                oFinText.Retrieve(cmbText.SelectedValue)
            Else
                oFinText.Retrieve(0)
            End If
            bolLoading = True

            txtName.Text = oFinText.Name
            txtText.Text = oFinText.Text
            chkActive.Checked = oFinText.Active

            bolLoading = False


        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Populate Tec Activity Documents" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub


#End Region
    

    

    

    Private Sub LoadcmbText()
        Dim dtText As DataTable

        dtText = oFinText.GetFinancialTextTable(ReasonType)

        If Not IsNothing(dtText) Then
            cmbText.DataSource = dtText
            cmbText.DisplayMember = "Text_Name"
            cmbText.ValueMember = "Text_ID"

            cmbText.SelectedIndex = 0
        Else
            cmbText.DataSource = Nothing
            SetForm_AddNew()
        End If


    End Sub

    Private Sub cmbText_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbText.SelectedIndexChanged
        If bolLoading Then Exit Sub
        If oFinText.IsDirty Then
            If MsgBox("Save Changes?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                If oFinText.ID <= 0 Then
                    oFinText.CreatedBy = MusterContainer.AppUser.ID
                Else
                    oFinText.ModifiedBy = MusterContainer.AppUser.ID
                End If
                oFinText.Save(CType(UIUtilsGen.ModuleID.Financial, Integer), MusterContainer.AppUser.UserKey, returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If
            End If
        End If
        SetForm_Update()

    End Sub

    Private Sub txtName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtName.TextChanged
        If bolLoading Then Exit Sub

        oFinText.Name = txtName.Text

    End Sub

    Private Sub txtText_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtText.TextChanged
        If bolLoading Then Exit Sub

        oFinText.Text = txtText.Text

    End Sub


    Private Sub chkActive_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkActive.CheckedChanged
        If bolLoading Then Exit Sub

        oFinText.Active = chkActive.Checked

    End Sub
End Class
