Public Class FilePathAdmin
    '-------------------------------------------------------------------------------
    ' MUSTER.MUSTER.FilePaths
    '   Provides the UI for managing system file paths
    '
    ' Copyright (C) 2004 CIBER, Inc.
    ' All rights reserved.
    '
    ' Release   Initials    Date        Description
    '  1.0        PN      12/??/04    Original class definition.
    '  1.1        JVC2    12/27/04    Rewrote interface to reuse common code and 
    '                                    leverage the new FilePathIsDirty event
    '                                    which activates/deactivates save and cancel
    '                                    controls.
    '  1.2        JVC2     01/03/05    Added code to handle MDI child integration with app
    '  1.3        JVC2     01/12/05    Added code to update filepath attributes on
    '                                     textchanged event of controls.
    '  1.4        AN      02/10/05    Integrated AppFlags new object model
    '-------------------------------------------------------------------------------
    '
    'TODO - Remove comment from VSS version 2/9/05 - JVC 2
    '
    Inherits System.Windows.Forms.Form

#Region "Private Member Variables"
    Private WithEvents oFilePath As Muster.BusinessLogic.pFilePaths
    Friend MyGUID As New System.Guid
    Private bolLoading As Boolean = False
    Dim returnVal As String = String.Empty
#End Region
#Region " Windows Form Designer generated code "

    Public Sub New(Optional ByRef frm As Windows.Forms.Form = Nothing)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        MyGUID = System.Guid.NewGuid
        MusterContainer.AppUser.LogEntry(Me.Text, MyGUID.ToString)
        MusterContainer.AppSemaphores.Retrieve(MyGUID.ToString, "WindowName", Me.Text)
        MusterContainer.AppSemaphores.Retrieve("0", "ActiveForm", MyGUID)

        If Not frm Is Nothing Then
            If frm.IsMdiContainer Then
                Me.MdiParent = frm
            End If
        End If

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
    Friend WithEvents lblFacilities As System.Windows.Forms.Label
    Friend WithEvents lblImages As System.Windows.Forms.Label
    Friend WithEvents btnSearchFacilities As System.Windows.Forms.Button
    Friend WithEvents txtFacilities As System.Windows.Forms.TextBox
    Friend WithEvents btnSearchManuallyCreated As System.Windows.Forms.Button
    Friend WithEvents lblManuallyCreated As System.Windows.Forms.Label
    Friend WithEvents txtManuallyCreated As System.Windows.Forms.TextBox
    Friend WithEvents lblDocuments As System.Windows.Forms.Label
    Friend WithEvents btnSearchSystemGenerated As System.Windows.Forms.Button
    Friend WithEvents lblSystemGenerated As System.Windows.Forms.Label
    Friend WithEvents txtSystemGenerated As System.Windows.Forms.TextBox
    Friend WithEvents txtSystemArchive As System.Windows.Forms.TextBox
    Friend WithEvents lblSystemArchive As System.Windows.Forms.Label
    Friend WithEvents btnSearchSystemArchive As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents FolderBrowerDialog As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents txtTemplatesPath As System.Windows.Forms.TextBox
    Friend WithEvents btnTemplatesPath As System.Windows.Forms.Button
    Friend WithEvents lblTemplatesPath As System.Windows.Forms.Label
    Friend WithEvents txtReportsPath As System.Windows.Forms.TextBox
    Friend WithEvents btnReportsPath As System.Windows.Forms.Button
    Friend WithEvents txtDBSyncPath As System.Windows.Forms.TextBox
    Friend WithEvents btnDBSyncPath As System.Windows.Forms.Button
    Friend WithEvents lblReportsPath As System.Windows.Forms.Label
    Friend WithEvents lblDBSync As System.Windows.Forms.Label
    Friend WithEvents lblLicensees As System.Windows.Forms.Label
    Friend WithEvents txtLicensees As System.Windows.Forms.TextBox
    Friend WithEvents btnSearchLicensees As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtSketches As System.Windows.Forms.TextBox
    Friend WithEvents btnSearchSketehes As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.btnSearchFacilities = New System.Windows.Forms.Button
        Me.lblFacilities = New System.Windows.Forms.Label
        Me.txtFacilities = New System.Windows.Forms.TextBox
        Me.lblImages = New System.Windows.Forms.Label
        Me.btnSearchManuallyCreated = New System.Windows.Forms.Button
        Me.lblManuallyCreated = New System.Windows.Forms.Label
        Me.txtManuallyCreated = New System.Windows.Forms.TextBox
        Me.lblDocuments = New System.Windows.Forms.Label
        Me.btnSearchSystemGenerated = New System.Windows.Forms.Button
        Me.lblSystemGenerated = New System.Windows.Forms.Label
        Me.txtSystemGenerated = New System.Windows.Forms.TextBox
        Me.txtSystemArchive = New System.Windows.Forms.TextBox
        Me.lblSystemArchive = New System.Windows.Forms.Label
        Me.btnSearchSystemArchive = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.FolderBrowerDialog = New System.Windows.Forms.FolderBrowserDialog
        Me.txtTemplatesPath = New System.Windows.Forms.TextBox
        Me.btnTemplatesPath = New System.Windows.Forms.Button
        Me.lblTemplatesPath = New System.Windows.Forms.Label
        Me.txtReportsPath = New System.Windows.Forms.TextBox
        Me.lblReportsPath = New System.Windows.Forms.Label
        Me.btnReportsPath = New System.Windows.Forms.Button
        Me.lblDBSync = New System.Windows.Forms.Label
        Me.txtDBSyncPath = New System.Windows.Forms.TextBox
        Me.btnDBSyncPath = New System.Windows.Forms.Button
        Me.lblLicensees = New System.Windows.Forms.Label
        Me.txtLicensees = New System.Windows.Forms.TextBox
        Me.btnSearchLicensees = New System.Windows.Forms.Button
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtSketches = New System.Windows.Forms.TextBox
        Me.btnSearchSketehes = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'btnSearchFacilities
        '
        Me.btnSearchFacilities.BackColor = System.Drawing.SystemColors.Control
        Me.btnSearchFacilities.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSearchFacilities.Location = New System.Drawing.Point(408, 40)
        Me.btnSearchFacilities.Name = "btnSearchFacilities"
        Me.btnSearchFacilities.Size = New System.Drawing.Size(32, 24)
        Me.btnSearchFacilities.TabIndex = 187
        Me.btnSearchFacilities.Text = " ..."
        Me.btnSearchFacilities.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'lblFacilities
        '
        Me.lblFacilities.Location = New System.Drawing.Point(16, 40)
        Me.lblFacilities.Name = "lblFacilities"
        Me.lblFacilities.Size = New System.Drawing.Size(96, 16)
        Me.lblFacilities.TabIndex = 186
        Me.lblFacilities.Text = "Facilities"
        Me.lblFacilities.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtFacilities
        '
        Me.txtFacilities.Location = New System.Drawing.Point(120, 40)
        Me.txtFacilities.Name = "txtFacilities"
        Me.txtFacilities.Size = New System.Drawing.Size(280, 20)
        Me.txtFacilities.TabIndex = 185
        Me.txtFacilities.Text = ""
        '
        'lblImages
        '
        Me.lblImages.Location = New System.Drawing.Point(16, 16)
        Me.lblImages.Name = "lblImages"
        Me.lblImages.Size = New System.Drawing.Size(48, 16)
        Me.lblImages.TabIndex = 188
        Me.lblImages.Text = "Images"
        Me.lblImages.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnSearchManuallyCreated
        '
        Me.btnSearchManuallyCreated.BackColor = System.Drawing.SystemColors.Control
        Me.btnSearchManuallyCreated.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSearchManuallyCreated.Location = New System.Drawing.Point(408, 256)
        Me.btnSearchManuallyCreated.Name = "btnSearchManuallyCreated"
        Me.btnSearchManuallyCreated.Size = New System.Drawing.Size(32, 24)
        Me.btnSearchManuallyCreated.TabIndex = 198
        Me.btnSearchManuallyCreated.Text = " ..."
        Me.btnSearchManuallyCreated.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'lblManuallyCreated
        '
        Me.lblManuallyCreated.Location = New System.Drawing.Point(16, 256)
        Me.lblManuallyCreated.Name = "lblManuallyCreated"
        Me.lblManuallyCreated.Size = New System.Drawing.Size(96, 24)
        Me.lblManuallyCreated.TabIndex = 197
        Me.lblManuallyCreated.Text = "Manually Created"
        Me.lblManuallyCreated.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtManuallyCreated
        '
        Me.txtManuallyCreated.Location = New System.Drawing.Point(120, 256)
        Me.txtManuallyCreated.Name = "txtManuallyCreated"
        Me.txtManuallyCreated.Size = New System.Drawing.Size(280, 20)
        Me.txtManuallyCreated.TabIndex = 196
        Me.txtManuallyCreated.Text = ""
        '
        'lblDocuments
        '
        Me.lblDocuments.Location = New System.Drawing.Point(16, 144)
        Me.lblDocuments.Name = "lblDocuments"
        Me.lblDocuments.Size = New System.Drawing.Size(72, 16)
        Me.lblDocuments.TabIndex = 195
        Me.lblDocuments.Text = "Documents"
        Me.lblDocuments.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnSearchSystemGenerated
        '
        Me.btnSearchSystemGenerated.BackColor = System.Drawing.SystemColors.Control
        Me.btnSearchSystemGenerated.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSearchSystemGenerated.Location = New System.Drawing.Point(408, 216)
        Me.btnSearchSystemGenerated.Name = "btnSearchSystemGenerated"
        Me.btnSearchSystemGenerated.Size = New System.Drawing.Size(32, 24)
        Me.btnSearchSystemGenerated.TabIndex = 194
        Me.btnSearchSystemGenerated.Text = " ..."
        Me.btnSearchSystemGenerated.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'lblSystemGenerated
        '
        Me.lblSystemGenerated.Location = New System.Drawing.Point(16, 216)
        Me.lblSystemGenerated.Name = "lblSystemGenerated"
        Me.lblSystemGenerated.Size = New System.Drawing.Size(96, 24)
        Me.lblSystemGenerated.TabIndex = 193
        Me.lblSystemGenerated.Text = "System Generated"
        Me.lblSystemGenerated.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtSystemGenerated
        '
        Me.txtSystemGenerated.Location = New System.Drawing.Point(120, 216)
        Me.txtSystemGenerated.Name = "txtSystemGenerated"
        Me.txtSystemGenerated.Size = New System.Drawing.Size(280, 20)
        Me.txtSystemGenerated.TabIndex = 192
        Me.txtSystemGenerated.Text = ""
        '
        'txtSystemArchive
        '
        Me.txtSystemArchive.Location = New System.Drawing.Point(120, 296)
        Me.txtSystemArchive.Name = "txtSystemArchive"
        Me.txtSystemArchive.Size = New System.Drawing.Size(280, 20)
        Me.txtSystemArchive.TabIndex = 203
        Me.txtSystemArchive.Text = ""
        '
        'lblSystemArchive
        '
        Me.lblSystemArchive.Location = New System.Drawing.Point(16, 296)
        Me.lblSystemArchive.Name = "lblSystemArchive"
        Me.lblSystemArchive.Size = New System.Drawing.Size(96, 27)
        Me.lblSystemArchive.TabIndex = 204
        Me.lblSystemArchive.Text = "System Archive"
        Me.lblSystemArchive.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'btnSearchSystemArchive
        '
        Me.btnSearchSystemArchive.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSearchSystemArchive.Location = New System.Drawing.Point(408, 296)
        Me.btnSearchSystemArchive.Name = "btnSearchSystemArchive"
        Me.btnSearchSystemArchive.Size = New System.Drawing.Size(32, 24)
        Me.btnSearchSystemArchive.TabIndex = 205
        Me.btnSearchSystemArchive.Text = "..."
        '
        'btnClose
        '
        Me.btnClose.Location = New System.Drawing.Point(312, 416)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(80, 23)
        Me.btnClose.TabIndex = 208
        Me.btnClose.Text = "Close"
        '
        'btnCancel
        '
        Me.btnCancel.Enabled = False
        Me.btnCancel.Location = New System.Drawing.Point(224, 416)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(80, 23)
        Me.btnCancel.TabIndex = 207
        Me.btnCancel.Text = "Cancel"
        '
        'btnSave
        '
        Me.btnSave.Enabled = False
        Me.btnSave.Location = New System.Drawing.Point(136, 416)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(80, 23)
        Me.btnSave.TabIndex = 206
        Me.btnSave.Text = "Save"
        '
        'txtTemplatesPath
        '
        Me.txtTemplatesPath.Location = New System.Drawing.Point(120, 176)
        Me.txtTemplatesPath.Name = "txtTemplatesPath"
        Me.txtTemplatesPath.Size = New System.Drawing.Size(280, 20)
        Me.txtTemplatesPath.TabIndex = 209
        Me.txtTemplatesPath.Text = ""
        '
        'btnTemplatesPath
        '
        Me.btnTemplatesPath.BackColor = System.Drawing.SystemColors.Control
        Me.btnTemplatesPath.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnTemplatesPath.Location = New System.Drawing.Point(408, 176)
        Me.btnTemplatesPath.Name = "btnTemplatesPath"
        Me.btnTemplatesPath.Size = New System.Drawing.Size(32, 24)
        Me.btnTemplatesPath.TabIndex = 211
        Me.btnTemplatesPath.Text = " ..."
        Me.btnTemplatesPath.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'lblTemplatesPath
        '
        Me.lblTemplatesPath.Location = New System.Drawing.Point(16, 176)
        Me.lblTemplatesPath.Name = "lblTemplatesPath"
        Me.lblTemplatesPath.Size = New System.Drawing.Size(96, 24)
        Me.lblTemplatesPath.TabIndex = 210
        Me.lblTemplatesPath.Text = "Templates"
        Me.lblTemplatesPath.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtReportsPath
        '
        Me.txtReportsPath.Location = New System.Drawing.Point(120, 336)
        Me.txtReportsPath.Name = "txtReportsPath"
        Me.txtReportsPath.Size = New System.Drawing.Size(280, 20)
        Me.txtReportsPath.TabIndex = 203
        Me.txtReportsPath.Text = ""
        '
        'lblReportsPath
        '
        Me.lblReportsPath.Location = New System.Drawing.Point(16, 336)
        Me.lblReportsPath.Name = "lblReportsPath"
        Me.lblReportsPath.Size = New System.Drawing.Size(96, 27)
        Me.lblReportsPath.TabIndex = 204
        Me.lblReportsPath.Text = "Reports"
        Me.lblReportsPath.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'btnReportsPath
        '
        Me.btnReportsPath.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnReportsPath.Location = New System.Drawing.Point(408, 336)
        Me.btnReportsPath.Name = "btnReportsPath"
        Me.btnReportsPath.Size = New System.Drawing.Size(32, 24)
        Me.btnReportsPath.TabIndex = 205
        Me.btnReportsPath.Text = "..."
        '
        'lblDBSync
        '
        Me.lblDBSync.Location = New System.Drawing.Point(16, 376)
        Me.lblDBSync.Name = "lblDBSync"
        Me.lblDBSync.Size = New System.Drawing.Size(96, 27)
        Me.lblDBSync.TabIndex = 204
        Me.lblDBSync.Text = "DB Sync to Local"
        Me.lblDBSync.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtDBSyncPath
        '
        Me.txtDBSyncPath.Location = New System.Drawing.Point(120, 376)
        Me.txtDBSyncPath.Name = "txtDBSyncPath"
        Me.txtDBSyncPath.Size = New System.Drawing.Size(280, 20)
        Me.txtDBSyncPath.TabIndex = 203
        Me.txtDBSyncPath.Text = ""
        '
        'btnDBSyncPath
        '
        Me.btnDBSyncPath.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDBSyncPath.Location = New System.Drawing.Point(408, 376)
        Me.btnDBSyncPath.Name = "btnDBSyncPath"
        Me.btnDBSyncPath.Size = New System.Drawing.Size(32, 24)
        Me.btnDBSyncPath.TabIndex = 205
        Me.btnDBSyncPath.Text = "..."
        '
        'lblLicensees
        '
        Me.lblLicensees.Location = New System.Drawing.Point(16, 80)
        Me.lblLicensees.Name = "lblLicensees"
        Me.lblLicensees.Size = New System.Drawing.Size(96, 16)
        Me.lblLicensees.TabIndex = 190
        Me.lblLicensees.Text = "Licensees"
        Me.lblLicensees.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtLicensees
        '
        Me.txtLicensees.Location = New System.Drawing.Point(120, 80)
        Me.txtLicensees.Name = "txtLicensees"
        Me.txtLicensees.Size = New System.Drawing.Size(280, 20)
        Me.txtLicensees.TabIndex = 189
        Me.txtLicensees.Text = ""
        '
        'btnSearchLicensees
        '
        Me.btnSearchLicensees.BackColor = System.Drawing.SystemColors.Control
        Me.btnSearchLicensees.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSearchLicensees.Location = New System.Drawing.Point(408, 80)
        Me.btnSearchLicensees.Name = "btnSearchLicensees"
        Me.btnSearchLicensees.Size = New System.Drawing.Size(32, 24)
        Me.btnSearchLicensees.TabIndex = 191
        Me.btnSearchLicensees.Text = " ..."
        Me.btnSearchLicensees.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(16, 120)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(96, 16)
        Me.Label1.TabIndex = 213
        Me.Label1.Text = "Sketches"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtSketches
        '
        Me.txtSketches.Location = New System.Drawing.Point(120, 120)
        Me.txtSketches.Name = "txtSketches"
        Me.txtSketches.Size = New System.Drawing.Size(280, 20)
        Me.txtSketches.TabIndex = 212
        Me.txtSketches.Text = ""
        '
        'btnSearchSketehes
        '
        Me.btnSearchSketehes.BackColor = System.Drawing.SystemColors.Control
        Me.btnSearchSketehes.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSearchSketehes.Location = New System.Drawing.Point(408, 120)
        Me.btnSearchSketehes.Name = "btnSearchSketehes"
        Me.btnSearchSketehes.Size = New System.Drawing.Size(32, 24)
        Me.btnSearchSketehes.TabIndex = 214
        Me.btnSearchSketehes.Text = " ..."
        Me.btnSearchSketehes.TextAlign = System.Drawing.ContentAlignment.TopLeft
        '
        'FilePathAdmin
        '
        Me.AcceptButton = Me.btnClose
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(528, 453)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtSketches)
        Me.Controls.Add(Me.btnSearchSketehes)
        Me.Controls.Add(Me.txtTemplatesPath)
        Me.Controls.Add(Me.btnTemplatesPath)
        Me.Controls.Add(Me.lblTemplatesPath)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.btnSearchSystemArchive)
        Me.Controls.Add(Me.lblSystemArchive)
        Me.Controls.Add(Me.txtSystemArchive)
        Me.Controls.Add(Me.txtManuallyCreated)
        Me.Controls.Add(Me.txtSystemGenerated)
        Me.Controls.Add(Me.txtFacilities)
        Me.Controls.Add(Me.btnSearchManuallyCreated)
        Me.Controls.Add(Me.lblManuallyCreated)
        Me.Controls.Add(Me.lblDocuments)
        Me.Controls.Add(Me.btnSearchSystemGenerated)
        Me.Controls.Add(Me.lblSystemGenerated)
        Me.Controls.Add(Me.lblImages)
        Me.Controls.Add(Me.btnSearchFacilities)
        Me.Controls.Add(Me.lblFacilities)
        Me.Controls.Add(Me.txtReportsPath)
        Me.Controls.Add(Me.lblReportsPath)
        Me.Controls.Add(Me.btnReportsPath)
        Me.Controls.Add(Me.lblDBSync)
        Me.Controls.Add(Me.txtDBSyncPath)
        Me.Controls.Add(Me.btnDBSyncPath)
        Me.Controls.Add(Me.lblLicensees)
        Me.Controls.Add(Me.txtLicensees)
        Me.Controls.Add(Me.btnSearchLicensees)
        Me.Name = "FilePathAdmin"
        Me.Text = "Manage File Paths"
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region "Form Level Events"
    Private Sub FilePaths_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        bolLoading = True
        oFilePath = New Muster.BusinessLogic.pFilePaths
        Try
            Me.BuildPaths()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        Finally
            bolLoading = False
        End Try

    End Sub
    Private Sub FilePaths_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing

        Try
            If oFilePath.colIsDirty Then

                Dim Results As MsgBoxResult = MsgBox("There are unsaved changes. Do you want to save changes before closing?", MsgBoxStyle.Question & MsgBoxStyle.YesNoCancel)
                Select Case Results
                    Case MsgBoxResult.No
                        oFilePath.Reset()
                    Case MsgBoxResult.Yes
                        If ValidPaths() Then
                            oFilePath.Flush(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
                            If Not UIUtilsGen.HasRights(returnVal) Then
                                Exit Sub
                            End If
                        End If
                    Case MsgBoxResult.Cancel
                        e.Cancel = True
                End Select
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub
#End Region
#Region "UI Control Events"

#Region "Images"
    Private Sub btnSearchFacilities_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearchFacilities.Click
        Me.ShowBrowser("Select the path to the Facilities picture directory.", UIUtilsGen.FilePathKey_FacImages, txtFacilities)
    End Sub
    Private Sub btnSearchLicensees_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearchLicensees.Click
        Me.ShowBrowser("Select the path to the Licensees picture directory.", UIUtilsGen.FilePathKey_LicenseesImages, txtLicensees)
    End Sub

    Private Sub btnSearchSketches_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearchSketehes.Click
        Me.ShowBrowser("Select the path to the inspections ketches picture directory.", UIUtilsGen.FilePathKey_Sketches, txtSketches)
    End Sub

    Private Sub txtFacilities_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFacilities.TextChanged
        If bolLoading Then Exit Sub
        SetFilePath(UIUtilsGen.FilePathKey_FacImages, txtFacilities.Text)
    End Sub
    Private Sub txtFacilities_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFacilities.Leave
        Me.ShowBrowser("", UIUtilsGen.FilePathKey_FacImages, txtFacilities, True)
    End Sub

    Private Sub txtLicensees_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLicensees.TextChanged
        If bolLoading Then Exit Sub
        SetFilePath(UIUtilsGen.FilePathKey_LicenseesImages, txtLicensees.Text)
    End Sub
    Private Sub txtLicensees_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLicensees.Leave
        Me.ShowBrowser("", UIUtilsGen.FilePathKey_LicenseesImages, txtLicensees, True)
    End Sub

    Private Sub txtSketches_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSketches.TextChanged
        If bolLoading Then Exit Sub
        SetFilePath(UIUtilsGen.FilePathKey_Sketches, txtSketches.Text)
    End Sub
    Private Sub txtSketches_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSketches.Leave
        Me.ShowBrowser("", UIUtilsGen.FilePathKey_Sketches, txtSketches, True)
    End Sub


#End Region

#Region "Docs"
    Private Sub btnTemplatesPath_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnTemplatesPath.Click
        Me.ShowBrowser("Select the path to the common Templates directory.", UIUtilsGen.FilePathKey_Templates, txtTemplatesPath)
    End Sub
    Private Sub btnSearchSystemGenerated_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearchSystemGenerated.Click
        Me.ShowBrowser("Select the path to the common System Generated documents directory.", UIUtilsGen.FilePathKey_SystemGenerated, txtSystemGenerated)
    End Sub
    Private Sub btnSearchManuallyCreated_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearchManuallyCreated.Click
        Me.ShowBrowser("Select the path to the common Manually Generated documents directory.", UIUtilsGen.FilePathKey_ManuallyCreated, txtManuallyCreated)
    End Sub
    Private Sub btnSearchSystemArchive_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSearchSystemArchive.Click
        Me.ShowBrowser("Select the path to the common System Archive directory.", UIUtilsGen.FilePathKey_SystemArchive, txtSystemArchive)
    End Sub
    Private Sub btnReportsPath_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReportsPath.Click
        Me.ShowBrowser("Select the path to the common Reports directory.", UIUtilsGen.FilePathKey_Reports, txtReportsPath)
    End Sub
    Private Sub btnDBSyncPath_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDBSyncPath.Click
        Me.ShowBrowser("Select the path to the common DB Sync directory.", UIUtilsGen.FilePathKey_DBSync, txtDBSyncPath)
    End Sub

    Private Sub txtTemplatesPath_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTemplatesPath.TextChanged
        If bolLoading Then Exit Sub
        SetFilePath(UIUtilsGen.FilePathKey_Templates, txtTemplatesPath.Text)
    End Sub
    Private Sub txtTemplatesPath_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtTemplatesPath.Leave
        Me.ShowBrowser("", UIUtilsGen.FilePathKey_Templates, txtTemplatesPath, True)
    End Sub

    Private Sub txtSystemGenerated_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSystemGenerated.TextChanged
        If bolLoading Then Exit Sub
        SetFilePath(UIUtilsGen.FilePathKey_SystemGenerated, txtSystemGenerated.Text)
    End Sub
    Private Sub txtSystemGenerated_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSystemGenerated.Leave
        Me.ShowBrowser("", UIUtilsGen.FilePathKey_SystemGenerated, txtSystemGenerated, True)
    End Sub

    Private Sub txtManuallyCreated_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtManuallyCreated.TextChanged
        If bolLoading Then Exit Sub
        SetFilePath(UIUtilsGen.FilePathKey_ManuallyCreated, txtManuallyCreated.Text)
    End Sub
    Private Sub txtManuallyCreated_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtManuallyCreated.Leave
        Me.ShowBrowser("", UIUtilsGen.FilePathKey_ManuallyCreated, txtManuallyCreated, True)
    End Sub

    Private Sub txtSystemArchive_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSystemArchive.TextChanged
        If bolLoading Then Exit Sub
        SetFilePath(UIUtilsGen.FilePathKey_SystemArchive, txtSystemArchive.Text)
    End Sub
    Private Sub txtSystemArchive_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSystemArchive.Leave
        Me.ShowBrowser("", UIUtilsGen.FilePathKey_SystemArchive, txtSystemArchive, True)
    End Sub

    Private Sub txtReportsPath_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtReportsPath.TextChanged
        If bolLoading Then Exit Sub
        SetFilePath(UIUtilsGen.FilePathKey_Reports, txtReportsPath.Text)
    End Sub
    Private Sub txtReportsPath_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtReportsPath.Leave
        Me.ShowBrowser("", UIUtilsGen.FilePathKey_Reports, txtReportsPath, True)
    End Sub

    Private Sub txtDBSyncPath_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDBSyncPath.TextChanged
        If bolLoading Then Exit Sub
        SetFilePath(UIUtilsGen.FilePathKey_DBSync, txtDBSyncPath.Text)
    End Sub
    Private Sub txtDBSyncPath_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDBSyncPath.Leave
        Me.ShowBrowser("", UIUtilsGen.FilePathKey_DBSync, txtDBSyncPath, True)
    End Sub
#End Region

    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        MyBase.Close()
    End Sub
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try
            If ValidPaths() Then
                Try
                    Dim ResetPathing As Boolean = oFilePath.colIsDirty
                    oFilePath.Flush(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If
                    oFilePath.Reset()
                    If ResetPathing Then
                        MusterContainer.ProfileData.GetAll()
                    End If
                    MsgBox("Save Successful")
                Catch ex As Exception
                    Throw ex
                End Try
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Try
            If Not oFilePath Is Nothing Then
                oFilePath.Reset()
                Me.BuildPaths()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

#End Region
#Region "UI Control Support"
    Private Sub SetFilePath(ByVal ID As String, ByVal Value As String)
        Try
            oFilePath.Retrieve(ID)
            oFilePath.FilePath = Value
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub ShowBrowser(ByVal Title As String, ByVal ID As String, ByRef ControlID As System.Windows.Forms.TextBox, Optional ByVal SkipBrowser As Boolean = False)

        Dim DlgResult As DialogResult
        Dim strValue As String

        If Not SkipBrowser Then
            Me.FolderBrowerDialog.Description = Title
            Me.FolderBrowerDialog.SelectedPath = ControlID.Text
            DlgResult = Me.FolderBrowerDialog.ShowDialog(Me)
        Else
            DlgResult = DialogResult.OK
        End If
        If DlgResult <> DialogResult.Cancel Then
            If Not SkipBrowser Then
                strValue = Me.FolderBrowerDialog.SelectedPath.ToString
            Else
                strValue = ControlID.Text
            End If
            ControlID.Text = strValue
            SetFilePath(ID, strValue)
        End If

    End Sub
    Private Sub BuildPaths()

        Dim dtFilePaths As DataTable
        Dim drRow As DataRow
        Try
            dtFilePaths = oFilePath.FilePathTable
            For Each drRow In dtFilePaths.Rows
                Select Case drRow.Item("FileName").toUpper
                    Case "FACILITIES"
                        txtFacilities.Text = drRow.Item("FilePath")
                    Case "MANUALLY_CREATED"
                        Me.txtManuallyCreated.Text = drRow.Item("FilePath")
                    Case "SYSTEM_ARCHIVE"
                        Me.txtSystemArchive.Text = drRow.Item("FilePath")
                    Case "LICENSEES"
                        Me.txtLicensees.Text = drRow.Item("FilePath")
                    Case "SYSTEM_GENERATED"
                        Me.txtSystemGenerated.Text = drRow.Item("FilePath")
                    Case "TEMPLATES"
                        Me.txtTemplatesPath.Text = drRow.Item("FilePath")
                    Case "REPORTS"
                        Me.txtReportsPath.Text = drRow.Item("FilePath")
                    Case "DBSYNC"
                        Me.txtDBSyncPath.Text = drRow.Item("FilePath")
                    Case "SKETCHES"
                        Me.txtSketches.Text = drRow.Item("FilePath")

                End Select
            Next

            btnSave.Enabled = False
            btnCancel.Enabled = False

        Catch ex As Exception
            Throw ex
        End Try

    End Sub
    Private Function CheckPath() As Boolean

        Dim Cntl As Control
        Dim tmpTextBox As System.Windows.Forms.TextBox
        Dim strValidPath As String
        Dim strErrMsg As String = String.Empty

        Try
            For Each Cntl In Me.Controls
                If Cntl.GetType.ToString.ToLower = "system.Windows.Forms.TextBox".ToLower Then
                    tmpTextBox = CType(Cntl, System.Windows.Forms.TextBox)
                    If tmpTextBox.Text <> String.Empty Then
                        strValidPath = UIUtilsGen.IsPathValid(tmpTextBox.Text)
                        If strValidPath <> String.Empty Then
                            strErrMsg += strValidPath + vbCrLf
                        End If
                    End If
                End If
            Next
            If strErrMsg.Length > 0 Then
                MsgBox("Illegal/Invalid Path:" + vbCrLf + strErrMsg)
                Return False
                Exit Function
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Function
    Function ValidPaths() As Boolean

        Dim bolLocalCheck As Boolean = True
        Dim strMsg As String = String.Empty
        Dim bolCheckPaths As Boolean
        Dim sender As Object
        Dim e As System.ComponentModel.CancelEventArgs
        Try
            bolCheckPaths = Me.CheckPath()
            If bolCheckPaths Then
                Return bolCheckPaths
            End If

            If txtFacilities.Text <> String.Empty Then
                If Not System.IO.Directory.Exists(txtFacilities.Text) Then
                    bolLocalCheck = False
                    strMsg = strMsg & vbCrLf & vbTab & "Facilities Image Path"
                End If
            End If
            If txtLicensees.Text <> String.Empty Then
                If Not System.IO.Directory.Exists(txtLicensees.Text) Then
                    bolLocalCheck = False
                    strMsg = strMsg & vbCrLf & vbTab & "Licensees Image Path"
                End If
            End If

            If txtTemplatesPath.Text <> String.Empty Then
                If Not System.IO.Directory.Exists(txtTemplatesPath.Text) Then
                    bolLocalCheck = False
                    strMsg = strMsg & vbCrLf & vbTab & "Templates Path"
                End If
            End If

            If txtSystemGenerated.Text <> String.Empty Then
                If Not System.IO.Directory.Exists(txtSystemGenerated.Text) Then
                    bolLocalCheck = False
                    strMsg = strMsg & vbCrLf & vbTab & "System Generated Documents Path"
                End If
            End If
            If txtManuallyCreated.Text <> String.Empty Then
                If Not System.IO.Directory.Exists(txtManuallyCreated.Text) Then
                    bolLocalCheck = False
                    strMsg = strMsg & vbCrLf & vbTab & "Manually Created Documents Path"
                End If
            End If
            If Me.txtSystemArchive.Text <> String.Empty Then
                If Not System.IO.Directory.Exists(Me.txtSystemArchive.Text) Then
                    bolLocalCheck = False
                    strMsg = strMsg & vbCrLf & vbTab & "System Archive Documents Path"
                End If
            End If
            If Me.txtReportsPath.Text <> String.Empty Then
                If Not System.IO.Directory.Exists(Me.txtReportsPath.Text) Then
                    bolLocalCheck = False
                    strMsg = strMsg & vbCrLf & vbTab & "Reports Path"
                End If
            End If
            If Me.txtDBSyncPath.Text <> String.Empty Then
                If Not System.IO.Directory.Exists(Me.txtDBSyncPath.Text) Then
                    bolLocalCheck = False
                    strMsg = strMsg & vbCrLf & vbTab & "DBSync Path"
                End If
            End If

            If Me.txtSketches.Text <> String.Empty Then
                If Not System.IO.Directory.Exists(Me.txtSketches.Text) Then
                    bolLocalCheck = False
                    strMsg = strMsg & vbCrLf & vbTab & "Sketches Path"
                End If
            End If

            If String.Compare(txtTemplatesPath.Text, txtFacilities.Text, True) = 0 Or _
                String.Compare(txtTemplatesPath.Text, txtLicensees.Text, True) = 0 Or _
                String.Compare(txtTemplatesPath.Text, txtSystemGenerated.Text, True) = 0 Or _
                String.Compare(txtTemplatesPath.Text, txtManuallyCreated.Text, True) = 0 Or _
                String.Compare(txtTemplatesPath.Text, txtSystemArchive.Text, True) = 0 Or _
                String.Compare(txtTemplatesPath.Text, txtSketches.Text, True) = 0 Or _
                String.Compare(txtTemplatesPath.Text, txtReportsPath.Text, True) = 0 Or _
                String.Compare(txtTemplatesPath.Text, txtDBSyncPath.Text, True) = 0 Then
                MessageBox.Show("Templates Path should be Unique.")
                Return False
            End If

            If bolLocalCheck = False Then
                Dim Results As Long = MsgBox("The following paths were NOT found on this system -" & strMsg & vbCrLf & "Do you want to save them anyway?", MsgBoxStyle.Question & MsgBoxStyle.YesNo)
                If Results = MsgBoxResult.No Then
                    Me.btnCancel_Click(sender, e)
                    Return False
                Else
                    Return True
                End If
            Else
                Return True
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Function
#End Region
#Region "External Event Handling"
    Private Sub FileInfoIsDirty(ByVal DirtyState As Boolean) Handles oFilePath.FilePathIsDirty
        Me.btnSave.Enabled = DirtyState
        Me.btnCancel.Enabled = DirtyState
    End Sub
#End Region

    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)
        If Not oFilePath Is Nothing Then
            If oFilePath.IsDirty Then
                Dim Results As Long = MsgBox("There are unsaved changes. Do you want to save changes before closing?", MsgBoxStyle.YesNoCancel)
                If Results = MsgBoxResult.Yes Then
                    oFilePath.Flush(CType(UIUtilsGen.ModuleID.Admin, Integer), MusterContainer.AppUser.UserKey, returnVal, MusterContainer.AppUser.ID)
                    If Not UIUtilsGen.HasRights(returnVal) Then
                        Exit Sub
                    End If
                Else
                    If Results = MsgBoxResult.Cancel Then
                        e.Cancel = True
                        Exit Sub
                    End If
                End If
            End If
        End If

        ' Remove any values from the shared collection for this screen
        '
        MusterContainer.AppSemaphores.Remove(MyGUID.ToString)
        '
        ' Log the disposal of the form (exit from Registration form)
        '
        MusterContainer.AppUser.LogExit(MyGUID.ToString)

    End Sub

    
End Class
