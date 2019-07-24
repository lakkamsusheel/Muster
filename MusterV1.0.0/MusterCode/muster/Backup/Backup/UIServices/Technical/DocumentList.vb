Public Class DocumentList
    Inherits System.Windows.Forms.Form


#Region " Local Variables "
    Private WithEvents oTecDocument As New MUSTER.BusinessLogic.pTecDoc
    Private bolLoading As Boolean = True
    Private returnVal As String = String.Empty
#End Region

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
    Friend WithEvents chkActive As System.Windows.Forms.CheckBox
    Friend WithEvents lblActivity As System.Windows.Forms.Label
    Friend WithEvents cmbType As System.Windows.Forms.ComboBox
    Friend WithEvents lblDocument As System.Windows.Forms.Label
    Friend WithEvents lblFileName As System.Windows.Forms.Label
    Friend WithEvents txtFileName As System.Windows.Forms.TextBox
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents cmbDocument As System.Windows.Forms.ComboBox
    Friend WithEvents gbMGPTFStatus As System.Windows.Forms.GroupBox
    Friend WithEvents chkSTFS As System.Windows.Forms.CheckBox
    Friend WithEvents chkNTFE As System.Windows.Forms.CheckBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtTriggerField As System.Windows.Forms.TextBox
    Friend WithEvents GPAutoDocs As System.Windows.Forms.GroupBox
    Friend WithEvents cmbAutoDoc4 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbAutoDoc3 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbAutoDoc2 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbAutoDoc1 As System.Windows.Forms.ComboBox
    Friend WithEvents txtDocument As System.Windows.Forms.TextBox
    Friend WithEvents btnAddNew As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents cmbTriggerField As System.Windows.Forms.ComboBox
    Friend WithEvents chkShowAll As System.Windows.Forms.CheckBox
    Friend WithEvents lblDocumentName As System.Windows.Forms.Label
    Friend WithEvents cmbAutoDoc5 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbAutoDoc7 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbAutoDoc9 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbAutoDoc10 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbAutoDoc8 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbAutoDoc6 As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.cmbType = New System.Windows.Forms.ComboBox
        Me.chkActive = New System.Windows.Forms.CheckBox
        Me.lblActivity = New System.Windows.Forms.Label
        Me.lblDocument = New System.Windows.Forms.Label
        Me.lblFileName = New System.Windows.Forms.Label
        Me.txtFileName = New System.Windows.Forms.TextBox
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.cmbDocument = New System.Windows.Forms.ComboBox
        Me.gbMGPTFStatus = New System.Windows.Forms.GroupBox
        Me.chkNTFE = New System.Windows.Forms.CheckBox
        Me.chkSTFS = New System.Windows.Forms.CheckBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtTriggerField = New System.Windows.Forms.TextBox
        Me.GPAutoDocs = New System.Windows.Forms.GroupBox
        Me.cmbAutoDoc6 = New System.Windows.Forms.ComboBox
        Me.cmbAutoDoc8 = New System.Windows.Forms.ComboBox
        Me.cmbAutoDoc10 = New System.Windows.Forms.ComboBox
        Me.cmbAutoDoc9 = New System.Windows.Forms.ComboBox
        Me.cmbAutoDoc7 = New System.Windows.Forms.ComboBox
        Me.cmbAutoDoc5 = New System.Windows.Forms.ComboBox
        Me.cmbAutoDoc4 = New System.Windows.Forms.ComboBox
        Me.cmbAutoDoc3 = New System.Windows.Forms.ComboBox
        Me.cmbAutoDoc2 = New System.Windows.Forms.ComboBox
        Me.cmbAutoDoc1 = New System.Windows.Forms.ComboBox
        Me.txtDocument = New System.Windows.Forms.TextBox
        Me.btnAddNew = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.cmbTriggerField = New System.Windows.Forms.ComboBox
        Me.chkShowAll = New System.Windows.Forms.CheckBox
        Me.lblDocumentName = New System.Windows.Forms.Label
        Me.gbMGPTFStatus.SuspendLayout()
        Me.GPAutoDocs.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmbType
        '
        Me.cmbType.Location = New System.Drawing.Point(96, 80)
        Me.cmbType.Name = "cmbType"
        Me.cmbType.Size = New System.Drawing.Size(128, 21)
        Me.cmbType.TabIndex = 2
        '
        'chkActive
        '
        Me.chkActive.Checked = True
        Me.chkActive.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkActive.Location = New System.Drawing.Point(256, 80)
        Me.chkActive.Name = "chkActive"
        Me.chkActive.Size = New System.Drawing.Size(72, 24)
        Me.chkActive.TabIndex = 3
        Me.chkActive.Text = "Active"
        '
        'lblActivity
        '
        Me.lblActivity.Location = New System.Drawing.Point(24, 88)
        Me.lblActivity.Name = "lblActivity"
        Me.lblActivity.Size = New System.Drawing.Size(64, 16)
        Me.lblActivity.TabIndex = 148
        Me.lblActivity.Text = "Type:"
        Me.lblActivity.TextAlign = System.Drawing.ContentAlignment.BottomRight
        '
        'lblDocument
        '
        Me.lblDocument.Location = New System.Drawing.Point(16, 29)
        Me.lblDocument.Name = "lblDocument"
        Me.lblDocument.Size = New System.Drawing.Size(72, 16)
        Me.lblDocument.TabIndex = 154
        Me.lblDocument.Text = "Document:"
        Me.lblDocument.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblFileName
        '
        Me.lblFileName.Location = New System.Drawing.Point(16, 112)
        Me.lblFileName.Name = "lblFileName"
        Me.lblFileName.Size = New System.Drawing.Size(72, 16)
        Me.lblFileName.TabIndex = 158
        Me.lblFileName.Text = "FileName:"
        Me.lblFileName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtFileName
        '
        Me.txtFileName.Location = New System.Drawing.Point(96, 112)
        Me.txtFileName.Name = "txtFileName"
        Me.txtFileName.Size = New System.Drawing.Size(232, 20)
        Me.txtFileName.TabIndex = 4
        Me.txtFileName.Text = ""
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(232, 304)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(96, 23)
        Me.btnCancel.TabIndex = 13
        Me.btnCancel.Text = "Cancel"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(16, 304)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(96, 23)
        Me.btnSave.TabIndex = 12
        Me.btnSave.Text = "Save"
        '
        'cmbDocument
        '
        Me.cmbDocument.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbDocument.Location = New System.Drawing.Point(96, 24)
        Me.cmbDocument.Name = "cmbDocument"
        Me.cmbDocument.Size = New System.Drawing.Size(232, 21)
        Me.cmbDocument.TabIndex = 1
        '
        'gbMGPTFStatus
        '
        Me.gbMGPTFStatus.Controls.Add(Me.chkNTFE)
        Me.gbMGPTFStatus.Controls.Add(Me.chkSTFS)
        Me.gbMGPTFStatus.Location = New System.Drawing.Point(96, 168)
        Me.gbMGPTFStatus.Name = "gbMGPTFStatus"
        Me.gbMGPTFStatus.Size = New System.Drawing.Size(208, 72)
        Me.gbMGPTFStatus.TabIndex = 6
        Me.gbMGPTFStatus.TabStop = False
        Me.gbMGPTFStatus.Text = "Available for MGPTF Status:"
        '
        'chkNTFE
        '
        Me.chkNTFE.Location = New System.Drawing.Point(8, 40)
        Me.chkNTFE.Name = "chkNTFE"
        Me.chkNTFE.Size = New System.Drawing.Size(176, 24)
        Me.chkNTFE.TabIndex = 7
        Me.chkNTFE.Text = "NTFE, EUD"
        '
        'chkSTFS
        '
        Me.chkSTFS.Location = New System.Drawing.Point(8, 16)
        Me.chkSTFS.Name = "chkSTFS"
        Me.chkSTFS.Size = New System.Drawing.Size(176, 24)
        Me.chkSTFS.TabIndex = 6
        Me.chkSTFS.Text = "STFS, STFS-Direct, Federal"
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 144)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(72, 16)
        Me.Label1.TabIndex = 170
        Me.Label1.Text = "Trigger Field:"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtTriggerField
        '
        Me.txtTriggerField.Location = New System.Drawing.Point(288, 248)
        Me.txtTriggerField.Name = "txtTriggerField"
        Me.txtTriggerField.Size = New System.Drawing.Size(48, 20)
        Me.txtTriggerField.TabIndex = 5
        Me.txtTriggerField.Text = ""
        Me.txtTriggerField.Visible = False
        '
        'GPAutoDocs
        '
        Me.GPAutoDocs.Controls.Add(Me.cmbAutoDoc6)
        Me.GPAutoDocs.Controls.Add(Me.cmbAutoDoc8)
        Me.GPAutoDocs.Controls.Add(Me.cmbAutoDoc10)
        Me.GPAutoDocs.Controls.Add(Me.cmbAutoDoc9)
        Me.GPAutoDocs.Controls.Add(Me.cmbAutoDoc7)
        Me.GPAutoDocs.Controls.Add(Me.cmbAutoDoc5)
        Me.GPAutoDocs.Controls.Add(Me.cmbAutoDoc4)
        Me.GPAutoDocs.Controls.Add(Me.cmbAutoDoc3)
        Me.GPAutoDocs.Controls.Add(Me.cmbAutoDoc2)
        Me.GPAutoDocs.Controls.Add(Me.cmbAutoDoc1)
        Me.GPAutoDocs.Location = New System.Drawing.Point(344, 16)
        Me.GPAutoDocs.Name = "GPAutoDocs"
        Me.GPAutoDocs.Size = New System.Drawing.Size(248, 320)
        Me.GPAutoDocs.TabIndex = 8
        Me.GPAutoDocs.TabStop = False
        Me.GPAutoDocs.Text = "Automatically Created Documents"
        '
        'cmbAutoDoc6
        '
        Me.cmbAutoDoc6.Location = New System.Drawing.Point(8, 160)
        Me.cmbAutoDoc6.Name = "cmbAutoDoc6"
        Me.cmbAutoDoc6.Size = New System.Drawing.Size(232, 21)
        Me.cmbAutoDoc6.TabIndex = 17
        '
        'cmbAutoDoc8
        '
        Me.cmbAutoDoc8.Location = New System.Drawing.Point(8, 224)
        Me.cmbAutoDoc8.Name = "cmbAutoDoc8"
        Me.cmbAutoDoc8.Size = New System.Drawing.Size(232, 21)
        Me.cmbAutoDoc8.TabIndex = 16
        '
        'cmbAutoDoc10
        '
        Me.cmbAutoDoc10.Location = New System.Drawing.Point(8, 288)
        Me.cmbAutoDoc10.Name = "cmbAutoDoc10"
        Me.cmbAutoDoc10.Size = New System.Drawing.Size(232, 21)
        Me.cmbAutoDoc10.TabIndex = 15
        '
        'cmbAutoDoc9
        '
        Me.cmbAutoDoc9.Location = New System.Drawing.Point(8, 256)
        Me.cmbAutoDoc9.Name = "cmbAutoDoc9"
        Me.cmbAutoDoc9.Size = New System.Drawing.Size(232, 21)
        Me.cmbAutoDoc9.TabIndex = 14
        '
        'cmbAutoDoc7
        '
        Me.cmbAutoDoc7.Location = New System.Drawing.Point(8, 192)
        Me.cmbAutoDoc7.Name = "cmbAutoDoc7"
        Me.cmbAutoDoc7.Size = New System.Drawing.Size(232, 21)
        Me.cmbAutoDoc7.TabIndex = 13
        '
        'cmbAutoDoc5
        '
        Me.cmbAutoDoc5.Location = New System.Drawing.Point(8, 128)
        Me.cmbAutoDoc5.Name = "cmbAutoDoc5"
        Me.cmbAutoDoc5.Size = New System.Drawing.Size(232, 21)
        Me.cmbAutoDoc5.TabIndex = 12
        '
        'cmbAutoDoc4
        '
        Me.cmbAutoDoc4.Location = New System.Drawing.Point(8, 96)
        Me.cmbAutoDoc4.Name = "cmbAutoDoc4"
        Me.cmbAutoDoc4.Size = New System.Drawing.Size(232, 21)
        Me.cmbAutoDoc4.TabIndex = 11
        '
        'cmbAutoDoc3
        '
        Me.cmbAutoDoc3.Location = New System.Drawing.Point(8, 67)
        Me.cmbAutoDoc3.Name = "cmbAutoDoc3"
        Me.cmbAutoDoc3.Size = New System.Drawing.Size(232, 21)
        Me.cmbAutoDoc3.TabIndex = 10
        '
        'cmbAutoDoc2
        '
        Me.cmbAutoDoc2.Location = New System.Drawing.Point(8, 41)
        Me.cmbAutoDoc2.Name = "cmbAutoDoc2"
        Me.cmbAutoDoc2.Size = New System.Drawing.Size(232, 21)
        Me.cmbAutoDoc2.TabIndex = 9
        '
        'cmbAutoDoc1
        '
        Me.cmbAutoDoc1.Location = New System.Drawing.Point(8, 15)
        Me.cmbAutoDoc1.Name = "cmbAutoDoc1"
        Me.cmbAutoDoc1.Size = New System.Drawing.Size(232, 21)
        Me.cmbAutoDoc1.TabIndex = 8
        '
        'txtDocument
        '
        Me.txtDocument.Location = New System.Drawing.Point(96, 48)
        Me.txtDocument.Name = "txtDocument"
        Me.txtDocument.Size = New System.Drawing.Size(232, 20)
        Me.txtDocument.TabIndex = 1
        Me.txtDocument.Text = ""
        '
        'btnAddNew
        '
        Me.btnAddNew.Location = New System.Drawing.Point(0, 0)
        Me.btnAddNew.Name = "btnAddNew"
        Me.btnAddNew.Size = New System.Drawing.Size(64, 16)
        Me.btnAddNew.TabIndex = 176
        Me.btnAddNew.TabStop = False
        Me.btnAddNew.Text = "Add New"
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(128, 304)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(96, 23)
        Me.btnDelete.TabIndex = 177
        Me.btnDelete.Text = "Delete"
        '
        'cmbTriggerField
        '
        Me.cmbTriggerField.Location = New System.Drawing.Point(96, 136)
        Me.cmbTriggerField.Name = "cmbTriggerField"
        Me.cmbTriggerField.Size = New System.Drawing.Size(232, 21)
        Me.cmbTriggerField.TabIndex = 178
        '
        'chkShowAll
        '
        Me.chkShowAll.Location = New System.Drawing.Point(104, 248)
        Me.chkShowAll.Name = "chkShowAll"
        Me.chkShowAll.Size = New System.Drawing.Size(80, 24)
        Me.chkShowAll.TabIndex = 179
        Me.chkShowAll.Text = "Show All"
        '
        'lblDocumentName
        '
        Me.lblDocumentName.Location = New System.Drawing.Point(16, 48)
        Me.lblDocumentName.Name = "lblDocumentName"
        Me.lblDocumentName.Size = New System.Drawing.Size(72, 16)
        Me.lblDocumentName.TabIndex = 180
        Me.lblDocumentName.Text = "Name:"
        Me.lblDocumentName.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'DocumentList
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(600, 341)
        Me.Controls.Add(Me.lblDocumentName)
        Me.Controls.Add(Me.chkShowAll)
        Me.Controls.Add(Me.cmbTriggerField)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.btnAddNew)
        Me.Controls.Add(Me.GPAutoDocs)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtTriggerField)
        Me.Controls.Add(Me.txtFileName)
        Me.Controls.Add(Me.chkActive)
        Me.Controls.Add(Me.txtDocument)
        Me.Controls.Add(Me.gbMGPTFStatus)
        Me.Controls.Add(Me.cmbDocument)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.lblFileName)
        Me.Controls.Add(Me.lblDocument)
        Me.Controls.Add(Me.cmbType)
        Me.Controls.Add(Me.lblActivity)
        Me.Name = "DocumentList"
        Me.Text = "Add / Modify Document List"
        Me.gbMGPTFStatus.ResumeLayout(False)
        Me.GPAutoDocs.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

#Region " Events "
    Private Sub DocumentList_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        bolLoading = True
        LoadDropDowns()
        cmbDocument.SelectedIndex = 0
        bolLoading = False
        LoadDocumentForm()

    End Sub

    Private Sub cmbDocument_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbDocument.SelectedIndexChanged
        If bolLoading = True Then
            Exit Sub
        End If
        Try
            If oTecDocument.IsDirty Then
                Select Case MsgBox("Changes were made.  Do you wish to save before continuing?", MsgBoxStyle.YesNoCancel)
                    Case MsgBoxResult.Yes
                        SaveDocument()
                    Case MsgBoxResult.No
                        Exit Select
                    Case MsgBoxResult.Cancel
                        Exit Sub
                End Select
            End If
            LoadDocumentForm()
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Change cmbDocument " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Sub cmbType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbType.SelectedIndexChanged
        If bolLoading = True Then
            Exit Sub
        End If
        oTecDocument.DocType = cmbType.SelectedValue
    End Sub

    Private Sub chkActive_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkActive.CheckedChanged
        If bolLoading = True Then
            Exit Sub
        End If

        oTecDocument.Active = chkActive.Checked
    End Sub

    Private Sub txtFileName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFileName.TextChanged
        If bolLoading = True Then
            Exit Sub
        End If

        oTecDocument.FileName = txtFileName.Text
    End Sub

    Private Sub chkSTFS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSTFS.CheckedChanged
        If bolLoading = True Then
            Exit Sub
        End If

        oTecDocument.STFS_Flag = chkSTFS.Checked
    End Sub

    Private Sub chkNTFE_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkNTFE.CheckedChanged
        If bolLoading = True Then
            Exit Sub
        End If

        oTecDocument.NTFE_Flag = chkNTFE.Checked
    End Sub

    Private Sub cmbAutoDoc1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbAutoDoc1.SelectedIndexChanged
        If bolLoading = True Then
            Exit Sub
        End If

        oTecDocument.Auto_Doc_1 = cmbAutoDoc1.SelectedValue
    End Sub
    Private Sub cmbAutoDoc2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbAutoDoc2.SelectedIndexChanged
        If bolLoading = True Then
            Exit Sub
        End If

        oTecDocument.Auto_Doc_2 = cmbAutoDoc2.SelectedValue
    End Sub
    Private Sub cmbAutoDoc3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbAutoDoc3.SelectedIndexChanged
        If bolLoading = True Then
            Exit Sub
        End If

        oTecDocument.Auto_Doc_3 = cmbAutoDoc3.SelectedValue
    End Sub
    Private Sub cmbAutoDoc4_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbAutoDoc4.SelectedIndexChanged
        If bolLoading = True Then
            Exit Sub
        End If

        oTecDocument.Auto_Doc_4 = cmbAutoDoc4.SelectedValue
    End Sub
    Private Sub cmbAutoDoc5_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbAutoDoc5.SelectedIndexChanged
        If bolLoading = True Then
            Exit Sub
        End If

        oTecDocument.Auto_Doc_5 = cmbAutoDoc5.SelectedValue
    End Sub
    Private Sub cmbAutoDoc6_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbAutoDoc6.SelectedIndexChanged
        If bolLoading = True Then
            Exit Sub
        End If

        oTecDocument.Auto_Doc_6 = cmbAutoDoc6.SelectedValue
    End Sub
    Private Sub cmbAutoDoc7_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbAutoDoc7.SelectedIndexChanged
        If bolLoading = True Then
            Exit Sub
        End If

        oTecDocument.Auto_Doc_7 = cmbAutoDoc7.SelectedValue
    End Sub
    Private Sub cmbAutoDoc8_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbAutoDoc8.SelectedIndexChanged
        If bolLoading = True Then
            Exit Sub
        End If

        oTecDocument.Auto_Doc_8 = cmbAutoDoc8.SelectedValue
    End Sub
    Private Sub cmbAutoDoc9_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbAutoDoc9.SelectedIndexChanged
        If bolLoading = True Then
            Exit Sub
        End If

        oTecDocument.Auto_Doc_9 = cmbAutoDoc9.SelectedValue
    End Sub
    Private Sub cmbAutoDoc10_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbAutoDoc10.SelectedIndexChanged
        If bolLoading = True Then
            Exit Sub
        End If

        oTecDocument.Auto_Doc_10 = cmbAutoDoc10.SelectedValue
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        oTecDocument.Reset()
        ' txtDocument.Visible = False
        ' cmbDocument.Visible = True

        LoadDocumentForm()
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            SaveDocument()
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Save Document " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Sub btnAddNew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddNew.Click

        Try

            If oTecDocument.IsDirty Then
                Select Case MsgBox("Changes were made.  Do you wish to save before continuing?", MsgBoxStyle.YesNoCancel)
                    Case MsgBoxResult.Yes
                        SaveDocument()
                    Case MsgBoxResult.No
                        Exit Select
                    Case MsgBoxResult.Cancel
                        Exit Sub
                End Select
            End If

            'cmbDocument.Visible = False
            'txtDocument.Visible = True
            bolLoading = True
            cmbDocument.SelectedIndex = -1
            cmbDocument.SelectedIndex = -1
            cmbDocument.Enabled = False
            txtDocument.Visible = True
            txtDocument.Text = String.Empty
            bolLoading = False


            oTecDocument.Retrieve(0)

            txtFileName.Text = oTecDocument.FileName
            chkActive.Checked = True
            oTecDocument.Active = True
            chkNTFE.Checked = oTecDocument.NTFE_Flag
            chkSTFS.Checked = oTecDocument.STFS_Flag
            'cmbType.SelectedValue = oTecDocument.DocType
            cmbType.SelectedIndex = -1
            cmbType.SelectedIndex = -1
            txtDocument.Text = oTecDocument.Name
            cmbAutoDoc1.SelectedValue = oTecDocument.Auto_Doc_1
            cmbAutoDoc2.SelectedValue = oTecDocument.Auto_Doc_2
            cmbAutoDoc3.SelectedValue = oTecDocument.Auto_Doc_3
            cmbAutoDoc4.SelectedValue = oTecDocument.Auto_Doc_4
            cmbAutoDoc5.SelectedValue = oTecDocument.Auto_Doc_5
            cmbAutoDoc6.SelectedValue = oTecDocument.Auto_Doc_6
            cmbAutoDoc7.SelectedValue = oTecDocument.Auto_Doc_7
            cmbAutoDoc8.SelectedValue = oTecDocument.Auto_Doc_8
            cmbAutoDoc9.SelectedValue = oTecDocument.Auto_Doc_9
            cmbAutoDoc10.SelectedValue = oTecDocument.Auto_Doc_10
            cmbTriggerField.SelectedValue = oTecDocument.Trigger_Field

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Add New " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()

        End Try

    End Sub

    Private Sub txtDocument_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDocument.TextChanged
        If bolLoading = True Then
            Exit Sub
        End If
        oTecDocument.Name = txtDocument.Text
    End Sub

    Private Sub cmbTriggerField_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbTriggerField.SelectedIndexChanged
        If bolLoading = True Then
            Exit Sub
        End If
        oTecDocument.Trigger_Field = cmbTriggerField.SelectedValue
    End Sub

    Private Sub DocumentList_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If oTecDocument.IsDirty Then
            If MsgBox("Changes were made.  Do you wish to save before continuing?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                SaveDocument()
            End If
        End If
    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click

        Try
            If oTecDocument.GetLustEventUsage > 0 Then
                MsgBox("You Cannot Delete This Document.  It's Currently Associated With A Lust Event?")
                Exit Sub
            End If
            If oTecDocument.GetActivityUsage > 0 Then
                MsgBox("You Cannot Delete This Document.  It's Currently Associated With An Activity?")
                Exit Sub
            End If


            If MsgBox("Do you wish to delete this Document?", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                oTecDocument.Deleted = True
                oTecDocument.ModifiedBy = MusterContainer.AppUser.ID
                oTecDocument.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), MusterContainer.AppUser.UserKey, returnVal)
                If Not returnVal = String.Empty Then
                    MessageBox.Show(returnVal.ToString(), "ACCESS DENIED", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Sub
                End If

                bolLoading = True
                LoadDropDowns()
                cmbDocument.SelectedIndex = 0
                bolLoading = False
                LoadDocumentForm()
            End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Delete " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()

        End Try
    End Sub

#End Region

#Region " Populate Routines "

    Private Sub PopulateDocumentList(Optional ByVal ShowAll As Boolean = False)
        Dim dtDocList As DataTable
        Try
            If ShowAll Then
                dtDocList = oTecDocument.PopulateTecDocumentList(False)
            Else
                dtDocList = oTecDocument.PopulateTecDocumentList(False, False)
            End If

            cmbDocument.DataSource = dtDocList
            cmbDocument.DisplayMember = "PROPERTY_NAME"
            cmbDocument.ValueMember = "PROPERTY_ID"
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Load System DropDowns " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub
    Private Sub LoadDropDowns()

        PopulateDocumentList(False)
        Dim dtTriggerList As DataTable = oTecDocument.PopulateTecDocumentTriggerList(True)
        Dim dtAddDocList1 As DataTable = oTecDocument.PopulateTecDocumentList(True)
        Dim dtAddDocList2 As DataTable = oTecDocument.PopulateTecDocumentList(True)
        Dim dtAddDocList3 As DataTable = oTecDocument.PopulateTecDocumentList(True)
        Dim dtAddDocList4 As DataTable = oTecDocument.PopulateTecDocumentList(True)
        Dim dtAddDocList5 As DataTable = oTecDocument.PopulateTecDocumentList(True)
        Dim dtAddDocList6 As DataTable = oTecDocument.PopulateTecDocumentList(True)
        Dim dtAddDocList7 As DataTable = oTecDocument.PopulateTecDocumentList(True)
        Dim dtAddDocList8 As DataTable = oTecDocument.PopulateTecDocumentList(True)
        Dim dtAddDocList9 As DataTable = oTecDocument.PopulateTecDocumentList(True)
        Dim dtAddDocList10 As DataTable = oTecDocument.PopulateTecDocumentList(True)

        Dim dtDocTypes As DataTable = oTecDocument.PopulateTecDocumentTypes
        Try


            cmbAutoDoc1.DataSource = dtAddDocList1
            cmbAutoDoc1.DisplayMember = "PROPERTY_NAME"
            cmbAutoDoc1.ValueMember = "PROPERTY_ID"

            cmbAutoDoc2.DataSource = dtAddDocList2
            cmbAutoDoc2.DisplayMember = "PROPERTY_NAME"
            cmbAutoDoc2.ValueMember = "PROPERTY_ID"

            cmbAutoDoc3.DataSource = dtAddDocList3
            cmbAutoDoc3.DisplayMember = "PROPERTY_NAME"
            cmbAutoDoc3.ValueMember = "PROPERTY_ID"

            cmbAutoDoc4.DataSource = dtAddDocList4
            cmbAutoDoc4.DisplayMember = "PROPERTY_NAME"
            cmbAutoDoc4.ValueMember = "PROPERTY_ID"

            cmbAutoDoc5.DataSource = dtAddDocList5
            cmbAutoDoc5.DisplayMember = "PROPERTY_NAME"
            cmbAutoDoc5.ValueMember = "PROPERTY_ID"

            cmbAutoDoc6.DataSource = dtAddDocList6
            cmbAutoDoc6.DisplayMember = "PROPERTY_NAME"
            cmbAutoDoc6.ValueMember = "PROPERTY_ID"

            cmbAutoDoc7.DataSource = dtAddDocList7
            cmbAutoDoc7.DisplayMember = "PROPERTY_NAME"
            cmbAutoDoc7.ValueMember = "PROPERTY_ID"

            cmbAutoDoc8.DataSource = dtAddDocList8
            cmbAutoDoc8.DisplayMember = "PROPERTY_NAME"
            cmbAutoDoc8.ValueMember = "PROPERTY_ID"

            cmbAutoDoc9.DataSource = dtAddDocList9
            cmbAutoDoc9.DisplayMember = "PROPERTY_NAME"
            cmbAutoDoc9.ValueMember = "PROPERTY_ID"

            cmbAutoDoc10.DataSource = dtAddDocList10
            cmbAutoDoc10.DisplayMember = "PROPERTY_NAME"
            cmbAutoDoc10.ValueMember = "PROPERTY_ID"

            cmbType.DataSource = dtDocTypes
            cmbType.DisplayMember = "PROPERTY_NAME"
            cmbType.ValueMember = "PROPERTY_ID"

            cmbTriggerField.DataSource = dtTriggerList
            cmbTriggerField.DisplayMember = "PROPERTY_NAME"
            cmbTriggerField.ValueMember = "PROPERTY_ID"

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Load System DropDowns " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Sub LoadDocumentForm()

        Try
            bolLoading = True
            If Not cmbDocument.SelectedValue Is Nothing Then
                oTecDocument.Retrieve(cmbDocument.SelectedValue)
                txtDocument.Text = oTecDocument.Name
                txtFileName.Text = oTecDocument.FileName
                chkActive.Checked = oTecDocument.Active
                chkNTFE.Checked = oTecDocument.NTFE_Flag
                chkSTFS.Checked = oTecDocument.STFS_Flag
                cmbType.SelectedValue = oTecDocument.DocType
                cmbAutoDoc1.SelectedValue = oTecDocument.Auto_Doc_1
                cmbAutoDoc2.SelectedValue = oTecDocument.Auto_Doc_2
                cmbAutoDoc3.SelectedValue = oTecDocument.Auto_Doc_3
                cmbAutoDoc4.SelectedValue = oTecDocument.Auto_Doc_4
                cmbAutoDoc5.SelectedValue = oTecDocument.Auto_Doc_5
                cmbAutoDoc6.SelectedValue = oTecDocument.Auto_Doc_6
                cmbAutoDoc7.SelectedValue = oTecDocument.Auto_Doc_7
                cmbAutoDoc8.SelectedValue = oTecDocument.Auto_Doc_8
                cmbAutoDoc9.SelectedValue = oTecDocument.Auto_Doc_9
                cmbAutoDoc10.SelectedValue = oTecDocument.Auto_Doc_10
                cmbTriggerField.SelectedValue = oTecDocument.Trigger_Field
            End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Load Document Form" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        Finally
            bolLoading = False
        End Try

    End Sub
#End Region

    Private Sub SaveDocument()
        Dim bolReload As Boolean = False

        Try
            If chkNTFE.Checked = False And chkSTFS.Checked = False Then
                MsgBox("At least one MGPTF Status must be selected")
                Exit Sub
            End If
            If Me.cmbType.Text = "" Then
                MsgBox("Document Type must be selected")
                Exit Sub
            End If
            If oTecDocument.ID = 0 Then
                bolReload = True
            End If
            If oTecDocument.ID <= 0 Then
                oTecDocument.CreatedBy = MusterContainer.AppUser.ID
            Else
                oTecDocument.ModifiedBy = MusterContainer.AppUser.ID
            End If

            oTecDocument.Save(CType(UIUtilsGen.ModuleID.Technical, Integer), MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If
            MsgBox("Document Saved Successfully.")
            cmbDocument.Enabled = True
            'If bolReload Then
            bolLoading = True
            LoadDropDowns()
            cmbDocument.SelectedIndex = -1
            txtDocument.Text = String.Empty
            bolLoading = False
            cmbDocument.SelectedValue = oTecDocument.ID
            If Not oTecDocument.Active Then
                txtDocument.Text = String.Empty
            Else
                txtDocument.Text = oTecDocument.Name
            End If

            'cmbDocument.Visible = True
            'txtDocument.Visible = False
            ' End If

        Catch ex As Exception
            If ex.Message = "Duplicate Entry" Then
                MessageBox.Show("Duplicate Document Name. Please Enter different Name.", "Duplicate Entry")
                txtDocument.Focus()
                Exit Sub
            End If
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Add New " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub chkShowAll_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowAll.CheckedChanged
        Try
            If chkShowAll.Checked Then
                PopulateDocumentList(True)
            Else
                PopulateDocumentList(False)
            End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Load System DropDowns " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub

End Class
