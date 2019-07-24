Public Class LateFeeWaiverRequest
    Inherits System.Windows.Forms.Form
#Region "User Defined Variables"
    Friend LateCertID As Int64
    Friend ugRowMisc As Infragistics.Win.UltraWinGrid.UltraGridRow
    Private oLateFee As New MUSTER.BusinessLogic.pFeeLateFee
    Private bolLoading As Boolean
    Dim returnVal As String = String.Empty
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
    Friend WithEvents pnlLateFeeBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlLateFeeDetails As System.Windows.Forms.Panel
    Friend WithEvents chkApprovalRecommended As System.Windows.Forms.CheckBox
    Friend WithEvents cmbExcuse As System.Windows.Forms.ComboBox
    Friend WithEvents lblExcuse As System.Windows.Forms.Label
    Friend WithEvents lblAmount As System.Windows.Forms.Label
    Friend WithEvents lblLateFeeWaiverRequest As System.Windows.Forms.Label
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Public WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents lblAmountValue As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlLateFeeBottom = New System.Windows.Forms.Panel
        Me.btnDelete = New System.Windows.Forms.Button
        Me.btnCancel = New System.Windows.Forms.Button
        Me.btnSave = New System.Windows.Forms.Button
        Me.pnlLateFeeDetails = New System.Windows.Forms.Panel
        Me.lblAmountValue = New System.Windows.Forms.Label
        Me.chkApprovalRecommended = New System.Windows.Forms.CheckBox
        Me.cmbExcuse = New System.Windows.Forms.ComboBox
        Me.lblExcuse = New System.Windows.Forms.Label
        Me.lblAmount = New System.Windows.Forms.Label
        Me.lblLateFeeWaiverRequest = New System.Windows.Forms.Label
        Me.pnlLateFeeBottom.SuspendLayout()
        Me.pnlLateFeeDetails.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlLateFeeBottom
        '
        Me.pnlLateFeeBottom.Controls.Add(Me.btnDelete)
        Me.pnlLateFeeBottom.Controls.Add(Me.btnCancel)
        Me.pnlLateFeeBottom.Controls.Add(Me.btnSave)
        Me.pnlLateFeeBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlLateFeeBottom.Location = New System.Drawing.Point(0, 134)
        Me.pnlLateFeeBottom.Name = "pnlLateFeeBottom"
        Me.pnlLateFeeBottom.Size = New System.Drawing.Size(480, 40)
        Me.pnlLateFeeBottom.TabIndex = 3
        '
        'btnDelete
        '
        Me.btnDelete.Location = New System.Drawing.Point(280, 8)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.TabIndex = 6
        Me.btnDelete.Text = "Delete"
        '
        'btnCancel
        '
        Me.btnCancel.Location = New System.Drawing.Point(200, 8)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.TabIndex = 5
        Me.btnCancel.Text = "Cancel"
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(120, 8)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.TabIndex = 4
        Me.btnSave.Text = "Save"
        '
        'pnlLateFeeDetails
        '
        Me.pnlLateFeeDetails.Controls.Add(Me.lblAmountValue)
        Me.pnlLateFeeDetails.Controls.Add(Me.chkApprovalRecommended)
        Me.pnlLateFeeDetails.Controls.Add(Me.cmbExcuse)
        Me.pnlLateFeeDetails.Controls.Add(Me.lblExcuse)
        Me.pnlLateFeeDetails.Controls.Add(Me.lblAmount)
        Me.pnlLateFeeDetails.Controls.Add(Me.lblLateFeeWaiverRequest)
        Me.pnlLateFeeDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlLateFeeDetails.Location = New System.Drawing.Point(0, 0)
        Me.pnlLateFeeDetails.Name = "pnlLateFeeDetails"
        Me.pnlLateFeeDetails.Size = New System.Drawing.Size(480, 134)
        Me.pnlLateFeeDetails.TabIndex = 0
        '
        'lblAmountValue
        '
        Me.lblAmountValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblAmountValue.Location = New System.Drawing.Point(176, 39)
        Me.lblAmountValue.Name = "lblAmountValue"
        Me.lblAmountValue.TabIndex = 10
        '
        'chkApprovalRecommended
        '
        Me.chkApprovalRecommended.Location = New System.Drawing.Point(87, 90)
        Me.chkApprovalRecommended.Name = "chkApprovalRecommended"
        Me.chkApprovalRecommended.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkApprovalRecommended.Size = New System.Drawing.Size(104, 32)
        Me.chkApprovalRecommended.TabIndex = 2
        Me.chkApprovalRecommended.Text = "Approval Recommended"
        '
        'cmbExcuse
        '
        Me.cmbExcuse.Location = New System.Drawing.Point(176, 64)
        Me.cmbExcuse.Name = "cmbExcuse"
        Me.cmbExcuse.Size = New System.Drawing.Size(296, 21)
        Me.cmbExcuse.TabIndex = 1
        '
        'lblExcuse
        '
        Me.lblExcuse.Location = New System.Drawing.Point(120, 64)
        Me.lblExcuse.Name = "lblExcuse"
        Me.lblExcuse.Size = New System.Drawing.Size(48, 17)
        Me.lblExcuse.TabIndex = 9
        Me.lblExcuse.Text = "Excuse:"
        '
        'lblAmount
        '
        Me.lblAmount.Location = New System.Drawing.Point(120, 40)
        Me.lblAmount.Name = "lblAmount"
        Me.lblAmount.Size = New System.Drawing.Size(48, 17)
        Me.lblAmount.TabIndex = 7
        Me.lblAmount.Text = "Amount:"
        '
        'lblLateFeeWaiverRequest
        '
        Me.lblLateFeeWaiverRequest.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLateFeeWaiverRequest.Location = New System.Drawing.Point(144, 8)
        Me.lblLateFeeWaiverRequest.Name = "lblLateFeeWaiverRequest"
        Me.lblLateFeeWaiverRequest.Size = New System.Drawing.Size(168, 23)
        Me.lblLateFeeWaiverRequest.TabIndex = 6
        Me.lblLateFeeWaiverRequest.Text = "Late Fee Waiver Request"
        '
        'LateFeeWaiverRequest
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(480, 174)
        Me.ControlBox = False
        Me.Controls.Add(Me.pnlLateFeeDetails)
        Me.Controls.Add(Me.pnlLateFeeBottom)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "LateFeeWaiverRequest"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Late Fee Waiver Request"
        Me.pnlLateFeeBottom.ResumeLayout(False)
        Me.pnlLateFeeDetails.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region "UI Control Events"
    Private Sub btnCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub
#End Region

    Private Sub LateFeeWaiverRequest_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load

        oLateFee.Retrieve(LateCertID)

        If LateCertID <= 0 Then
            If ugRowMisc Is Nothing Then
                MsgBox("Invalid Data. Misc. Invoice Row not selected")
                Me.Close()
                Exit Sub
            Else
                oLateFee.FiscalYear = ugRowMisc.Cells("SFY").Value
                oLateFee.InvoiceNumber = ugRowMisc.Cells("InvoiceID").Value
                oLateFee.LateCharges = ugRowMisc.Cells("Charges").Value
            End If
        End If

        If oLateFee.ProcessWaiver Then
            If oLateFee.WaiveApprovalStatus Then
                MsgBox("Late Fee Processed - Approved")
            Else
                MsgBox("Late Fee Processed - Denied")
            End If
            Me.Close()
            Exit Sub
        End If

        LoadForm()


    End Sub

    Private Sub LoadForm()
        bolLoading = True
        LoadDropDowns()
        lblAmountValue.Text = FormatNumber(oLateFee.LateCharges, 2, TriState.True, TriState.False, TriState.True)
        chkApprovalRecommended.Checked = oLateFee.WaiveApprovalRecommendation
        If LateCertID > 0 Then
            btnDelete.Enabled = True
        Else
            btnDelete.Enabled = False
        End If
        bolLoading = False
    End Sub

    Private Sub LoadDropDowns()
        Try
            Dim dtLateFeeExcuses As DataTable = oLateFee.PopulateLateFeeWaiverExcuses
            If Not IsNothing(dtLateFeeExcuses) Then
                cmbExcuse.DataSource = dtLateFeeExcuses
                cmbExcuse.DisplayMember = "PROPERTY_NAME"
                cmbExcuse.ValueMember = "PROPERTY_ID"
            Else
                cmbExcuse.DataSource = Nothing
            End If

            If IsDBNull(oLateFee.WaiveReason) Then
                cmbExcuse.SelectedIndex = -1
            Else
                cmbExcuse.SelectedValue = oLateFee.WaiveReason
            End If


        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Populate dtLateFeeExcuses " + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Try

            If oLateFee.ID <= 0 Then
                oLateFee.CreatedBy = MusterContainer.AppUser.ID
            Else
                oLateFee.ModifiedBy = MusterContainer.AppUser.ID
            End If

            Dim newExcuse As Boolean = False

            If Me.cmbExcuse.SelectedIndex > -1 AndAlso Me.cmbExcuse.Text.ToUpper <> DirectCast(Me.cmbExcuse.Items(cmbExcuse.SelectedIndex), DataRowView).Item(0).ToUpper AndAlso Me.cmbExcuse.Text <> String.Empty Then
                newExcuse = True
            ElseIf Me.cmbExcuse.SelectedIndex = -1 AndAlso Me.cmbExcuse.Text.Length > 0 Then
                newExcuse = True
            End If

            If newExcuse Then

                If MsgBox(String.Format("Do you want to save the excuse : {0} "" {1} ""  ? ", vbCrLf, Me.cmbExcuse.Text), MsgBoxStyle.YesNo, "New Excuse Wizard") = MsgBoxResult.Yes Then
                    oLateFee.WaiveReason = oLateFee.SaveExcuse(Me.cmbExcuse.Text, MusterContainer.AppUser.ID)
                Else
                    Exit Sub
                End If

            End If

            oLateFee.Save(CType(UIUtilsGen.ModuleID.Fees, Integer), MusterContainer.AppUser.UserKey, returnVal)
            If Not UIUtilsGen.HasRights(returnVal) Then
                Exit Sub
            End If

            MsgBox("Waiver Saved", MsgBoxStyle.Information, "Confirmation")
            Me.Close()
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Save Waiver" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try


    End Sub

    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try

            If MsgBox("Delete Waiver?", MsgBoxStyle.YesNo, "Delete Waiver") = MsgBoxResult.Yes Then
                oLateFee.ProcessWaiver = False
                oLateFee.WaiveApprovalStatus = False
                oLateFee.WaiverFinalizedOn = CDate("01/01/0001")
                oLateFee.WaiveApprovalRecommendation = 0
                oLateFee.WaiveApprovalStatus = 0
                oLateFee.WaiveReason = 0
                oLateFee.ModifiedBy = MusterContainer.AppUser.ID
                oLateFee.Save(CType(UIUtilsGen.ModuleID.Fees, Integer), MusterContainer.AppUser.UserKey, returnVal)
                If Not UIUtilsGen.HasRights(returnVal) Then
                    Exit Sub
                End If
                MsgBox("Waiver Deleted", MsgBoxStyle.Information, "Confirmation")
            End If
            Me.Close()
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot Delete Waiver" + vbCrLf + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Sub cmbExcuse_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbExcuse.SelectedIndexChanged
        If bolLoading Then Exit Sub

        oLateFee.WaiveReason = cmbExcuse.SelectedValue
    End Sub

    Private Sub chkApprovalRecommended_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkApprovalRecommended.CheckedChanged
        If bolLoading Then Exit Sub

        oLateFee.WaiveApprovalRecommendation = chkApprovalRecommended.Checked

    End Sub
End Class
